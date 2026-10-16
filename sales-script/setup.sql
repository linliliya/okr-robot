-- ============================================================
-- 销售话术复盘 · Supabase 初始化脚本
-- 在 skill-hub 那个 Supabase 项目里执行（依赖 skill-hub 已建好的 public.profiles）
-- 用法：Supabase 后台 → SQL Editor → New query → 整段粘贴 → Run
-- 可以重复执行，不会删数据
-- 如需课程资料知识库，完成本脚本后，再执行同目录 knowledge-setup.sql。
-- ============================================================

-- ---------- 成员白名单 ----------
-- skill-hub 开放注册，所以只有名单里的账号能看/改话术和案例。
-- 名单绑定的是账号 id：只能添加已经注册的邮箱，防止有人抢先用同事的邮箱注册来冒用权限
create table if not exists public.sales_members (
  email text primary key check (email = lower(email)),
  user_id uuid not null unique references auth.users(id) on delete cascade,
  added_by uuid references public.profiles(id) on delete set null default auth.uid(),
  added_at timestamptz not null default now()
);

create or replace function public.is_sales_member() returns boolean
  language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.sales_members where user_id = auth.uid());
$$;

-- 页面上「添加成员」走这个函数：按邮箱找到已注册的账号再加进名单
create or replace function public.add_sales_member(p_email text) returns void
  language plpgsql security definer set search_path = public as $$
declare
  e text := lower(trim(coalesce(p_email, '')));
  uid uuid;
begin
  if not public.is_sales_member() then
    raise exception '你还不是成员，不能添加别人';
  end if;
  select id into uid from auth.users where lower(email) = e limit 1;
  if uid is null then
    raise exception '这个邮箱还没注册账号，请对方先在 Skill 市集或本系统注册';
  end if;
  if exists (select 1 from public.sales_members where user_id = uid or email = e) then
    raise exception '这个邮箱已经在名单里了';
  end if;
  insert into public.sales_members (email, user_id, added_by) values (e, uid, auth.uid());
end;
$$;
revoke all on function public.add_sales_member(text) from public, anon;
grant execute on function public.add_sales_member(text) to authenticated;

alter table public.sales_members enable row level security;
drop policy if exists "成员可查看名单" on public.sales_members;
create policy "成员可查看名单" on public.sales_members
  for select to authenticated using (public.is_sales_member());
-- 不开放直接 insert，只能通过 add_sales_member
drop policy if exists "成员可添加成员" on public.sales_members;
drop policy if exists "成员可移除别人" on public.sales_members;
create policy "成员可移除别人" on public.sales_members
  for delete to authenticated
  using (public.is_sales_member() and user_id <> auth.uid());

-- ---------- 通用：创建人/修改人/版本号 ----------
create or replace function public.sales_stamp() returns trigger
  language plpgsql as $$
begin
  if tg_op = 'INSERT' then
    new.created_by := auth.uid();
    new.updated_by := auth.uid();
    new.created_at := now();
    new.updated_at := now();
    new.version := 1;
  else
    new.created_by := old.created_by;
    new.created_at := old.created_at;
    new.updated_by := auth.uid();
    new.updated_at := now();
    new.version := old.version + 1;
  end if;
  return new;
end;
$$;

-- ---------- 话术库 ----------
create table if not exists public.sales_scripts (
  id bigint generated always as identity primary key,
  stage text not null default '',            -- 阶段：邀约入群 / 社群预热 / 直播中 / 直播后追单 ...
  scene text not null,                       -- 场景或异议，如「嫌贵」「再考虑一下」
  customer_says text not null default '',    -- 客户的典型说法（多行）
  reply text not null default '',            -- 推荐回复
  tips text not null default '',             -- 要点 / 禁忌
  channel text not null default 'all' check (channel in ('all','xiaoe','bilibili')),
  status text not null default 'active' check (status in ('active','archived')),
  sort_order integer not null default 0,
  change_note text not null default '',      -- 这一版为什么改
  change_review_id bigint,                   -- 在哪次复盘会上改的
  change_case_ids bigint[] not null default '{}',  -- 依据哪些案例
  version integer not null default 1,
  created_by uuid references public.profiles(id) on delete set null,
  updated_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists sales_scripts_stage_idx on public.sales_scripts (stage, sort_order);

drop trigger if exists sales_scripts_stamp on public.sales_scripts;
create trigger sales_scripts_stamp before insert or update on public.sales_scripts
  for each row execute function public.sales_stamp();

-- 每一版都存一份快照（只由触发器写入）
create table if not exists public.sales_script_versions (
  id bigint generated always as identity primary key,
  script_id bigint not null,                 -- 不加外键：话术删了历史也留着
  version integer not null,
  action text not null,                      -- insert | update | delete
  snapshot jsonb not null,
  change_note text not null default '',
  review_id bigint,
  case_ids bigint[] not null default '{}',
  changed_by uuid references public.profiles(id) on delete set null,
  changed_at timestamptz not null default now()
);
create index if not exists sales_script_versions_script_idx on public.sales_script_versions (script_id, version desc);
create index if not exists sales_script_versions_review_idx on public.sales_script_versions (review_id);

create or replace function public.sales_script_history() returns trigger
  language plpgsql security definer set search_path = public as $$
declare r public.sales_scripts;
begin
  if tg_op = 'DELETE' then r := old; else r := new; end if;
  insert into public.sales_script_versions
    (script_id, version, action, snapshot, change_note, review_id, case_ids, changed_by)
  values
    (r.id, r.version, lower(tg_op), to_jsonb(r),
     case when tg_op = 'DELETE' then '删除' else r.change_note end,
     case when tg_op = 'DELETE' then null else r.change_review_id end,
     case when tg_op = 'DELETE' then '{}'::bigint[] else r.change_case_ids end,
     auth.uid());
  return null;
end;
$$;

drop trigger if exists sales_scripts_history on public.sales_scripts;
create trigger sales_scripts_history after insert or update or delete on public.sales_scripts
  for each row execute function public.sales_script_history();

-- ---------- 案例 ----------
create table if not exists public.sales_cases (
  id bigint generated always as identity primary key,
  case_date date not null,                   -- 成交 / 流失的日期
  review_week date not null,                 -- 归到哪一周的复盘会（周一）
  seller text not null default '',           -- 销售
  channel text not null default 'xiaoe' check (channel in ('xiaoe','bilibili','other')),
  cohort text not null default '',           -- 期次
  result text not null check (result in ('won','lost')),
  profile text not null default '',          -- 客户画像
  dialogue jsonb not null default '[]'::jsonb,  -- [{role:'customer'|'sales'|'note', text}]
  stuck_point text not null default '',      -- 成交 / 流失的关键在哪
  self_review text not null default '',      -- 销售自己的复盘
  scenes text[] not null default '{}',       -- 涉及的场景
  script_refs jsonb not null default '[]'::jsonb, -- 用到的话术 [{id, version}]
  ai_analysis jsonb,                         -- AI 分析结果
  discussion text not null default '',       -- 复盘会讨论结论
  review_status text not null default 'pending' check (review_status in ('pending','discussed','skipped')),
  version integer not null default 1,
  created_by uuid references public.profiles(id) on delete set null,
  updated_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists sales_cases_week_idx on public.sales_cases (review_week desc);

drop trigger if exists sales_cases_stamp on public.sales_cases;
create trigger sales_cases_stamp before insert or update on public.sales_cases
  for each row execute function public.sales_stamp();

-- ---------- 复盘会（一周一条） ----------
create table if not exists public.sales_reviews (
  id bigint generated always as identity primary key,
  week date not null unique,                 -- 周一
  summary text not null default '',          -- 复盘总结
  todos jsonb not null default '[]'::jsonb,  -- [{text, owner, done}]
  ai_brief jsonb,                            -- 会前 AI 分析（含修改建议及处理状态）
  version integer not null default 1,
  created_by uuid references public.profiles(id) on delete set null,
  updated_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

drop trigger if exists sales_reviews_stamp on public.sales_reviews;
create trigger sales_reviews_stamp before insert or update on public.sales_reviews
  for each row execute function public.sales_stamp();

-- ---------- 话术助手的提问记录 + 反馈 ----------
create table if not exists public.sales_ai_logs (
  id bigint generated always as identity primary key,
  message text not null,                     -- 客户说的话
  context text not null default '',          -- 前面的聊天
  channel text not null default 'all',
  result jsonb,                              -- AI 返回的推荐
  feedback text not null default '' check (feedback in ('','good','bad')),
  feedback_note text not null default '',
  created_by uuid references public.profiles(id) on delete set null default auth.uid(),
  created_at timestamptz not null default now()
);
create index if not exists sales_ai_logs_time_idx on public.sales_ai_logs (created_at desc);

-- ---------- 设置（如：给 AI 的业务背景资料） ----------
create table if not exists public.sales_settings (
  key text primary key,
  value text not null default '',
  updated_by uuid references public.profiles(id) on delete set null default auth.uid(),
  updated_at timestamptz not null default now()
);

-- ---------- 行级权限 ----------
alter table public.sales_settings enable row level security;
drop policy if exists "成员可查看" on public.sales_settings;
create policy "成员可查看" on public.sales_settings
  for select to authenticated using (public.is_sales_member());
drop policy if exists "成员可新增" on public.sales_settings;
create policy "成员可新增" on public.sales_settings
  for insert to authenticated with check (public.is_sales_member());
drop policy if exists "成员可修改" on public.sales_settings;
create policy "成员可修改" on public.sales_settings
  for update to authenticated using (public.is_sales_member());

alter table public.sales_scripts enable row level security;
alter table public.sales_script_versions enable row level security;
alter table public.sales_cases enable row level security;
alter table public.sales_reviews enable row level security;
alter table public.sales_ai_logs enable row level security;

do $$
declare t text;
begin
  -- 话术、案例、复盘会：成员都能查看和修改（所有改动有版本号，话术有完整历史）
  foreach t in array array['sales_scripts','sales_cases','sales_reviews'] loop
    execute format('drop policy if exists "成员可查看" on public.%I', t);
    execute format('create policy "成员可查看" on public.%I for select to authenticated using (public.is_sales_member())', t);
    execute format('drop policy if exists "成员可新增" on public.%I', t);
    execute format('create policy "成员可新增" on public.%I for insert to authenticated with check (public.is_sales_member())', t);
    execute format('drop policy if exists "成员可修改" on public.%I', t);
    execute format('create policy "成员可修改" on public.%I for update to authenticated using (public.is_sales_member()) with check (public.is_sales_member())', t);
  end loop;
end $$;

drop policy if exists "成员可删除案例" on public.sales_cases;
create policy "成员可删除案例" on public.sales_cases
  for delete to authenticated using (public.is_sales_member());

-- 话术不在页面上删除（用「停用」），历史表只能看
drop policy if exists "成员可查看" on public.sales_script_versions;
create policy "成员可查看" on public.sales_script_versions
  for select to authenticated using (public.is_sales_member());

drop policy if exists "成员可查看" on public.sales_ai_logs;
create policy "成员可查看" on public.sales_ai_logs
  for select to authenticated using (public.is_sales_member());
drop policy if exists "成员可新增" on public.sales_ai_logs;
create policy "成员可新增" on public.sales_ai_logs
  for insert to authenticated with check (public.is_sales_member() and created_by = auth.uid());
drop policy if exists "本人可反馈" on public.sales_ai_logs;
create policy "本人可反馈" on public.sales_ai_logs
  for update to authenticated using (public.is_sales_member() and created_by = auth.uid());

-- BEGIN AI_LOG_DELETE
-- 已有库可单独执行 ai-log-delete-setup.sql，权限范围与「本人可反馈」一致。
alter table public.sales_ai_logs enable row level security;
drop policy if exists "本人可删除提问" on public.sales_ai_logs;
create policy "本人可删除提问" on public.sales_ai_logs
  for delete to authenticated using (public.is_sales_member() and created_by = auth.uid());
grant select, delete on public.sales_ai_logs to authenticated;

-- 使用调用者权限执行，DELETE 仍必须通过 RLS；返回实际删除的编号，零行不能冒充成功。
create or replace function public.delete_sales_ai_log(p_id bigint) returns bigint
  language plpgsql security invoker set search_path = pg_catalog, public as $$
declare
  owner_id uuid;
  deleted_id bigint;
begin
  if auth.uid() is null then
    raise exception using errcode = '42501', message = 'ai_log_auth_required';
  end if;
  if not coalesce(public.is_sales_member(), false) then
    raise exception using errcode = '42501', message = 'sales_members_only';
  end if;
  if p_id is null or p_id <= 0 then
    raise exception using errcode = '22023', message = 'ai_log_invalid_id';
  end if;
  select created_by into owner_id from public.sales_ai_logs where id = p_id;
  if not found then
    raise exception using errcode = 'P0002', message = 'ai_log_not_found';
  end if;
  if owner_id is distinct from auth.uid() then
    raise exception using errcode = '42501', message = 'ai_log_not_owner';
  end if;
  delete from public.sales_ai_logs where id = p_id and created_by = auth.uid() returning id into deleted_id;
  if deleted_id is null then
    if not exists (select 1 from public.sales_ai_logs where id = p_id) then
      raise exception using errcode = 'P0002', message = 'ai_log_not_found';
    end if;
    raise exception using errcode = '42501', message = 'ai_log_delete_denied';
  end if;
  return deleted_id;
end;
$$;
revoke all on function public.delete_sales_ai_log(bigint) from public, anon;
grant execute on function public.delete_sales_ai_log(bigint) to authenticated;
-- END AI_LOG_DELETE

-- ---------- 第一个成员 ----------
-- 先在页面上注册账号（或用 Skill 市集的账号），再把下面这行的邮箱换成你的，单独执行一次：
-- insert into public.sales_members (email, user_id) select lower(email), id from auth.users where lower(email) = lower('你的邮箱');
