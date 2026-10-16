-- ============================================================
-- Skill 共享平台 · Supabase 初始化脚本
-- 用法：Supabase 后台 → SQL Editor → 新建查询 → 整段粘贴 → Run
-- ============================================================

begin;

-- ---------- 用户资料表（注册时自动创建） ----------
create table public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  display_name text not null default '',
  department text not null default '其他',
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;

create policy "登录用户可查看所有资料" on public.profiles
  for select to authenticated using (true);
create policy "只能创建自己的资料" on public.profiles
  for insert to authenticated with check (auth.uid() = id);
create policy "只能修改自己的资料" on public.profiles
  for update to authenticated using (auth.uid() = id);

-- 注册成功后自动写入 profiles（姓名、部门来自注册表单）
create or replace function public.handle_new_user()
returns trigger
language plpgsql security definer set search_path = public
as $$
begin
  insert into public.profiles (id, display_name, department)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'display_name', '未命名'),
    coalesce(new.raw_user_meta_data->>'department', '其他')
  );
  return new;
end;
$$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- ---------- skill 主表 ----------
create table public.skills (
  id bigint generated always as identity primary key,
  slug text unique not null,
  title text not null,
  description text not null default '',
  category text not null default '其他',
  tags text not null default '',
  author_id uuid not null references public.profiles(id),
  downloads integer not null default 0,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.skills enable row level security;

create policy "登录用户可浏览" on public.skills
  for select to authenticated using (true);
create policy "登录用户可发布" on public.skills
  for insert to authenticated with check (auth.uid() = author_id);
-- 最后一个版本被删除时，skill 本身也要能删掉（仅限 skill 作者）
create policy "作者可删除自己的 skill" on public.skills
  for delete to authenticated using (auth.uid() = author_id);

-- ---------- 版本表 ----------
create table public.versions (
  id bigint generated always as identity primary key,
  skill_id bigint not null references public.skills(id) on delete cascade,
  version integer not null,
  changelog text not null default '',
  file_path text not null,
  skill_md text not null default '',
  uploader_id uuid not null references public.profiles(id),
  uploaded_at timestamptz not null default now(),
  unique (skill_id, version)
);

alter table public.versions enable row level security;

create policy "登录用户可查看版本" on public.versions
  for select to authenticated using (true);
create policy "可删除自己上传的版本" on public.versions
  for delete to authenticated using (auth.uid() = uploader_id);

-- ---------- 外部推荐字段与权限（与增量迁移保持一致） ----------
alter table public.skills
  add column if not exists source_type text not null default 'upload',
  add column if not exists source_url text not null default '',
  add column if not exists source_key text not null default '',
  add column if not exists source_repo text not null default '',
  add column if not exists source_path text not null default '',
  add column if not exists original_author text not null default '',
  add column if not exists recommendation text not null default '',
  add column if not exists tested boolean not null default false,
  add column if not exists tested_with text not null default '',
  add column if not exists tested_version text not null default '',
  add column if not exists install_notes text not null default '';

alter table public.skills
  drop constraint if exists skills_source_type_check,
  drop constraint if exists skills_external_source_check,
  drop constraint if exists skills_external_text_lengths_check;

alter table public.skills
  add constraint skills_source_type_check
    check (source_type in ('upload', 'external')),
  add constraint skills_external_source_check check (
    source_type <> 'external' or (
      -- 只允许 HTTPS，域名部分不能含账号、密码或空白字符。
      source_url ~ '^https://[A-Za-z0-9]([A-Za-z0-9.-]*[A-Za-z0-9])?(:[0-9]{1,5})?([/?#][^[:space:][:cntrl:]]*)?$'
      and source_url !~ '[[:space:][:cntrl:]]'
      and position(chr(92) in source_url) = 0
      and length(source_key) > 0
      and source_key !~ '[[:space:][:cntrl:]]'
      and title ~ '[^[:space:]]'
    )
  ),
  add constraint skills_external_text_lengths_check check (
    -- 不给历史上传内容补加长度限制，保留旧客户端的上传方式。
    source_type <> 'external' or (
      length(title) <= 200
      and length(description) <= 4000
      and length(tags) <= 500
      and length(source_url) <= 2048
      and octet_length(source_key) <= 2048
      and length(source_repo) <= 200
      and length(source_path) <= 1024
      and length(original_author) <= 200
      and length(recommendation) <= 4000
      and length(tested_with) <= 200
      and length(tested_version) <= 200
      and length(install_notes) <= 8000
    )
  );

-- 不同标题可以指向同一来源；唯一性由前端规范化后的来源标识决定。
create unique index if not exists skills_external_source_key_unique
  on public.skills (source_key)
  where source_type = 'external' and source_key <> '';

-- 内部条目继续由全体同事协作，外部推荐只允许分享者修改。
drop policy if exists "登录用户可更新" on public.skills;
create policy "登录用户可更新" on public.skills
  for update to authenticated
  using (source_type = 'upload' or auth.uid() = author_id)
  with check (source_type = 'upload' or auth.uid() = author_id);

-- 避免利用内部条目的协作权限，改作者、改来源类型或绕过来源去重。
-- 换来源需要新建条目；简介、试用记录等内容仍可修改。
create or replace function public.protect_skill_source_identity()
returns trigger
language plpgsql set search_path = public
as $$
begin
  if new.author_id is distinct from old.author_id
    or new.source_type is distinct from old.source_type
    or new.source_key is distinct from old.source_key
    or new.source_url is distinct from old.source_url
    or new.source_repo is distinct from old.source_repo
    or new.source_path is distinct from old.source_path then
    raise exception 'Skill 的作者与来源信息在创建后不能修改，请新建条目。'
      using errcode = '23514';
  end if;
  return new;
end;
$$;

drop trigger if exists protect_skill_source_identity on public.skills;
create trigger protect_skill_source_identity
  before update on public.skills
  for each row execute function public.protect_skill_source_identity();

-- 外部推荐不存 ZIP 版本。RLS 和触发器同时限制插入与关联变更。
drop policy if exists "登录用户可上传版本" on public.versions;
create policy "登录用户可上传版本" on public.versions
  for insert to authenticated with check (
    auth.uid() = uploader_id
    and exists (
      select 1 from public.skills
      where skills.id = versions.skill_id and skills.source_type = 'upload'
    )
  );

create or replace function public.require_upload_skill_for_version()
returns trigger
language plpgsql set search_path = public
as $$
begin
  if not exists (
    select 1 from public.skills
    where id = new.skill_id and source_type = 'upload'
  ) then
    raise exception '只有内部上传的 Skill 可以添加文件版本。'
      using errcode = '23514';
  end if;
  return new;
end;
$$;

drop trigger if exists require_upload_skill_for_version on public.versions;
create trigger require_upload_skill_for_version
  before insert or update of skill_id on public.versions
  for each row execute function public.require_upload_skill_for_version();

-- 若先前手动写入过不符合来源规则的数据，整次迁移回滚，交由维护者核对。
do $$
begin
  if exists (
    select 1 from public.versions v
    join public.skills s on s.id = v.skill_id
    where s.source_type = 'external'
  ) then
    raise exception '已有外部推荐关联了文件版本，请核对数据后重新执行迁移。';
  end if;
end;
$$;

-- 下载计数仅记录内部文件下载。复制外部安装指引不计为下载或安装。
create or replace function public.increment_downloads(p_skill_id bigint)
returns void
language sql security definer set search_path = public
as $$
  update public.skills set downloads = downloads + 1
  where id = p_skill_id and source_type = 'upload' and auth.uid() is not null;
$$;

revoke all on function public.increment_downloads(bigint) from public, anon;
grant execute on function public.increment_downloads(bigint) to authenticated;

-- ---------- 适用岗位（与 update-skill-roles.sql 保持一致） ----------
alter table public.skills
  add column if not exists roles text[] not null default '{}';

alter table public.skills
  alter column roles set default '{}',
  alter column roles set not null,
  drop constraint if exists skills_roles_check;

-- 只转换已知旧选项。未知值或畸形数组让整次迁移回滚，留给维护者核对。
do $$
begin
  if exists (
    select 1 from public.skills
    where not case
      when cardinality(roles) = 0 then true
      when array_ndims(roles) = 1 then
        array_lower(roles, 1) = 1
        and cardinality(roles) <= 12
        and array_position(roles, null) is null
        and roles <@ array[
          '通用', '内容创作', '课程运营', '市场运营', '销售', '教学教研',
          '人力行政', '产品经理', '设计师', '剪辑',
          '客户服务', '研发工程师', '财务', '管理者'
        ]::text[]
        and (not ('通用' = any(roles)) or cardinality(roles) = 1)
      else false
    end
  ) then
    raise exception '已有 Skill 的岗位含未知值或格式异常，请核对数据后重新执行迁移。'
      using errcode = '23514';
  end if;
end;
$$;

with normalized as (
  select s.id, coalesce((
    select array_agg(mapped.role order by mapped.first_position)
    from (
      select case when item.role = '市场运营' then '课程运营' else item.role end as role,
        min(item.position) as first_position
      from unnest(s.roles) with ordinality as item(role, position)
      where item.role not in ('客户服务', '研发工程师', '财务', '管理者')
      group by case when item.role = '市场运营' then '课程运营' else item.role end
    ) mapped
  ), '{}'::text[]) as roles
  from public.skills s
)
update public.skills s set roles = normalized.roles
from normalized
where s.id = normalized.id and s.roles is distinct from normalized.roles;

alter table public.skills
  add constraint skills_roles_check check (
    case
      when cardinality(roles) = 0 then true
      when array_ndims(roles) = 1 then
        array_lower(roles, 1) = 1
        and cardinality(roles) <= 9
        and array_position(roles, null) is null
        and roles <@ array[
          '通用', '内容创作', '课程运营', '销售', '教学教研',
          '人力行政', '产品经理', '设计师', '剪辑'
        ]::text[]
        and (not ('通用' = any(roles)) or cardinality(roles) = 1)
      else false
    end
  );

-- ---------- 功能标题（与 update-skill-display-titles.sql 保持一致） ----------
alter table public.skills
  add column if not exists display_title text not null default '';

alter table public.skills
  alter column display_title set default '',
  alter column display_title set not null,
  drop constraint if exists skills_display_title_check;

update public.skills
set display_title = case title
  when 'de-ai-copy' then '去AI味写作'
  when 'ai-moments' then '生成 AI 资讯朋友圈'
  when 'book-insights' then '生成读书感悟朋友圈'
  when 'python-crawler-success' then '生成 Python 爬虫训练营成功案例朋友圈'
end
where display_title = ''
  and title in ('de-ai-copy', 'ai-moments', 'book-insights', 'python-crawler-success');

alter table public.skills
  add constraint skills_display_title_check check (
    display_title = '' or (
      char_length(display_title) <= 80
      -- 与 JavaScript trim 使用相同的空白字符集合。
      and btrim(display_title, U&'\0009\000A\000B\000C\000D\0020\00A0\1680\2000\2001\2002\2003\2004\2005\2006\2007\2008\2009\200A\2028\2029\202F\205F\3000\FEFF') <> ''
      and display_title !~ U&'[\0001-\001F\007F-\009F\2028\2029]'
    )
  );

-- ---------- 文件存储桶 ----------
insert into storage.buckets (id, name, public)
values ('skills', 'skills', false)
on conflict (id) do nothing;

create policy "登录用户可下载文件" on storage.objects
  for select to authenticated using (bucket_id = 'skills');
create policy "登录用户可上传文件" on storage.objects
  for insert to authenticated with check (bucket_id = 'skills');
create policy "可删除自己上传的文件" on storage.objects
  for delete to authenticated
  using (bucket_id = 'skills' and (owner = auth.uid() or owner_id = auth.uid()::text));

commit;
