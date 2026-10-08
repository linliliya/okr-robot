-- ============================================================
-- 训练营周复盘 · Supabase 初始化脚本
-- 在 Skill 市集用的同一个 Supabase 项目里执行：
--   SQL Editor → New query → 整段粘贴 → Run
-- 依赖 skill-hub/setup.sql 建好的 public.profiles 表。可以重复执行。
--
-- 安全设计：
--   · Skill 市集开放注册，所以复盘数据只对「成员白名单」里的邮箱开放
--   · 非成员登录后看不到任何复盘数据，也不能把自己加进白名单
--   · 成员不能把自己从白名单删掉（名单里至少会留一个人）
--   · 创建人 / 修改人 / 版本号 / 时间全部由数据库触发器填写，网页端改不了
--   · 每次新建、修改、删除都会在历史表留一份整行快照，删掉的期次可以找回
-- ============================================================

-- ---------- 0. 前置检查 ----------
do $$
begin
  if to_regclass('public.profiles') is null then
    raise exception '没找到 public.profiles 表：请先在这个 Supabase 项目里执行 skill-hub/setup.sql';
  end if;
end $$;


-- ============================================================
-- 1. 成员白名单
-- ============================================================
create table if not exists public.camp_review_members (
  email text primary key
    constraint camp_review_members_email_check
    check (email = lower(btrim(email)) and email ~ '^[^@[:space:]]+@[^@[:space:]]+\.[^@[:space:]]+$'),
  added_by uuid default auth.uid()
    constraint camp_review_members_added_by_fkey references public.profiles(id) on delete set null,
  added_at timestamptz not null default now()
);

-- 当前登录用户是否在白名单里（security definer：绕过白名单表自身的行级权限，避免策略递归）
create or replace function public.is_camp_review_member()
returns boolean
language sql stable security definer set search_path = ''
as $$
  select exists (
    select 1 from public.camp_review_members m
    where m.email = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;

revoke execute on function public.is_camp_review_member() from public, anon;
grant execute on function public.is_camp_review_member() to authenticated;

-- 添加成员时：统一小写、记录添加人；只允许添加已经注册过的邮箱
-- （Skill 市集关了邮箱验证，提前放进白名单的邮箱有被别人抢先注册的风险）
create or replace function public.camp_review_members_before_insert()
returns trigger
language plpgsql security definer set search_path = ''
as $$
declare
  registered boolean;
begin
  -- 非成员先拦下（行级权限也会拦；这里提前拦，是为了不让非成员借此打听某个邮箱有没有注册过）
  -- 后台 SQL Editor 里执行时没有登录身份（auth.uid() 为空），放行
  if auth.uid() is not null and not public.is_camp_review_member() then
    raise exception '只有复盘系统成员才能添加成员' using errcode = '42501';
  end if;
  new.email := lower(btrim(new.email));
  new.added_by := (select p.id from public.profiles p where p.id = auth.uid());
  new.added_at := now();
  begin
    registered := exists (select 1 from auth.users u where lower(u.email) = new.email);
  exception when insufficient_privilege then
    registered := true;   -- 读不到账号表时不拦，避免整个添加功能不可用
  end;
  if not registered then
    raise exception '邮箱 % 还没有注册账号：请对方先在 Skill 市集或复盘系统注册，再添加', new.email
      using errcode = 'P0001';
  end if;
  return new;
end;
$$;

drop trigger if exists camp_review_members_before_insert on public.camp_review_members;
create trigger camp_review_members_before_insert
  before insert on public.camp_review_members
  for each row execute function public.camp_review_members_before_insert();

-- 账号改邮箱 / 被删除时同步白名单，避免旧邮箱空出来后被别人注册顶替
-- 出错只记警告，绝不影响登录注册
create or replace function public.camp_review_sync_member_email()
returns trigger
language plpgsql security definer set search_path = ''
as $$
begin
  begin
    if tg_op = 'DELETE' then
      delete from public.camp_review_members where email = lower(old.email);
    elsif old.email is not null and lower(old.email) is distinct from lower(new.email) then
      if new.email is null
         or exists (select 1 from public.camp_review_members where email = lower(new.email)) then
        delete from public.camp_review_members where email = lower(old.email);
      else
        update public.camp_review_members set email = lower(new.email) where email = lower(old.email);
      end if;
    end if;
  exception when others then
    raise warning 'camp_review_sync_member_email: %', sqlerrm;
  end;
  return null;
end;
$$;

do $$
begin
  drop trigger if exists camp_review_member_email_sync on auth.users;
  create trigger camp_review_member_email_sync
    after update of email or delete on auth.users
    for each row execute function public.camp_review_sync_member_email();
exception when insufficient_privilege then
  raise notice '没有权限在 auth.users 上建触发器，已跳过邮箱同步（不影响主要功能）';
end $$;

alter table public.camp_review_members enable row level security;

drop policy if exists "成员可查看成员名单" on public.camp_review_members;
create policy "成员可查看成员名单" on public.camp_review_members
  for select to authenticated
  using ((select public.is_camp_review_member()));

drop policy if exists "成员可添加成员" on public.camp_review_members;
create policy "成员可添加成员" on public.camp_review_members
  for insert to authenticated
  with check ((select public.is_camp_review_member()));

drop policy if exists "成员可移除其他成员" on public.camp_review_members;
create policy "成员可移除其他成员" on public.camp_review_members
  for delete to authenticated
  using (
    (select public.is_camp_review_member())
    and email <> lower(coalesce((select auth.jwt() ->> 'email'), ''))
  );

-- 不开放修改；添加时只能填邮箱，其余字段由触发器填写
revoke all on table public.camp_review_members from anon, authenticated;
grant select, delete on table public.camp_review_members to authenticated;
grant insert (email) on table public.camp_review_members to authenticated;


-- ============================================================
-- 2. 期次表（一期 = 一场直播）
-- ============================================================
create table if not exists public.camp_cohorts (
  id bigint generated always as identity primary key,
  channel text not null
    constraint camp_cohorts_channel_check check (channel in ('xiaoe', 'bilibili')),
  name text not null
    constraint camp_cohorts_name_check check (btrim(name) <> ''),
  live_date date not null,
  closed boolean not null default false,
  summary text not null default '',
  data jsonb not null default '{}'::jsonb
    constraint camp_cohorts_data_check check (jsonb_typeof(data) = 'object'),
  version integer not null default 1,
  created_by uuid
    constraint camp_cohorts_created_by_fkey references public.profiles(id) on delete set null,
  updated_by uuid
    constraint camp_cohorts_updated_by_fkey references public.profiles(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists camp_cohorts_channel_live_date_idx
  on public.camp_cohorts (channel, live_date desc);

-- 历史表：触发器写入；不加外键，期次删掉后历史仍保留
create table if not exists public.camp_cohort_history (
  id bigint generated always as identity primary key,
  cohort_id bigint not null,
  action text not null
    constraint camp_cohort_history_action_check check (action in ('insert', 'update', 'delete')),
  snapshot jsonb not null,
  changed_by uuid
    constraint camp_cohort_history_changed_by_fkey references public.profiles(id) on delete set null,
  changed_at timestamptz not null default now()
);

create index if not exists camp_cohort_history_cohort_idx
  on public.camp_cohort_history (cohort_id, id desc);

-- 写入前：创建人、修改人、时间、版本号一律由数据库决定，忽略网页端传来的值
create or replace function public.camp_cohorts_before_write()
returns trigger
language plpgsql security definer set search_path = ''
as $$
declare
  editor uuid := (select p.id from public.profiles p where p.id = auth.uid());
begin
  if tg_op = 'INSERT' then
    new.created_by := editor;
    new.updated_by := editor;
    new.created_at := now();
    new.updated_at := new.created_at;
    new.version := 1;
    return new;
  end if;

  -- 外键级联（删除账号后把创建人/修改人置空）不算一次编辑，原样放行
  -- 网页端的修改都在第 1 层触发，走不到这里
  if pg_trigger_depth() > 1 then
    return new;
  end if;

  new.created_by := old.created_by;
  new.created_at := old.created_at;
  new.updated_by := editor;
  new.updated_at := now();
  new.version := old.version + 1;   -- 乐观锁：网页端用 version 做条件更新
  return new;
end;
$$;

drop trigger if exists camp_cohorts_before_write on public.camp_cohorts;
create trigger camp_cohorts_before_write
  before insert or update on public.camp_cohorts
  for each row execute function public.camp_cohorts_before_write();

-- 写入后：整行快照进历史表（delete 时存删除前的整行）
create or replace function public.camp_cohorts_log_history()
returns trigger
language plpgsql security definer set search_path = ''
as $$
declare
  editor uuid := (select p.id from public.profiles p where p.id = auth.uid());
begin
  if tg_op = 'DELETE' then
    insert into public.camp_cohort_history (cohort_id, action, snapshot, changed_by)
    values (old.id, 'delete', to_jsonb(old), editor);
  else
    insert into public.camp_cohort_history (cohort_id, action, snapshot, changed_by)
    values (new.id, lower(tg_op), to_jsonb(new), editor);
  end if;
  return null;
end;
$$;

drop trigger if exists camp_cohorts_log_history on public.camp_cohorts;
create trigger camp_cohorts_log_history
  after insert or update or delete on public.camp_cohorts
  for each row execute function public.camp_cohorts_log_history();

-- 行级权限：只有成员能读写
alter table public.camp_cohorts enable row level security;

drop policy if exists "成员可查看期次" on public.camp_cohorts;
create policy "成员可查看期次" on public.camp_cohorts
  for select to authenticated
  using ((select public.is_camp_review_member()));

drop policy if exists "成员可新建期次" on public.camp_cohorts;
create policy "成员可新建期次" on public.camp_cohorts
  for insert to authenticated
  with check ((select public.is_camp_review_member()));

drop policy if exists "成员可修改期次" on public.camp_cohorts;
create policy "成员可修改期次" on public.camp_cohorts
  for update to authenticated
  using ((select public.is_camp_review_member()))
  with check ((select public.is_camp_review_member()));

drop policy if exists "成员可删除期次" on public.camp_cohorts;
create policy "成员可删除期次" on public.camp_cohorts
  for delete to authenticated
  using ((select public.is_camp_review_member()));

-- 列级权限：网页端只能写业务字段；id、版本号、创建人、修改人、时间都碰不到
revoke all on table public.camp_cohorts from anon, authenticated;
grant select, delete on table public.camp_cohorts to authenticated;
grant insert (channel, name, live_date, closed, summary, data) on table public.camp_cohorts to authenticated;
grant update (channel, name, live_date, closed, summary, data) on table public.camp_cohorts to authenticated;

-- 历史表：成员只读，没有任何写入权限（只能由上面的触发器写）
alter table public.camp_cohort_history enable row level security;

drop policy if exists "成员可查看修改记录" on public.camp_cohort_history;
create policy "成员可查看修改记录" on public.camp_cohort_history
  for select to authenticated
  using ((select public.is_camp_review_member()));

revoke all on table public.camp_cohort_history from anon, authenticated;
grant select on table public.camp_cohort_history to authenticated;


-- 让接口层立即识别新表
notify pgrst, 'reload schema';

-- ============================================================
-- 3. 添加第一批成员（对方要先注册账号；把邮箱换成真实的再执行）
-- ============================================================
-- insert into public.camp_review_members (email) values ('你的邮箱');
-- insert into public.camp_review_members (email) values ('同事的邮箱');
