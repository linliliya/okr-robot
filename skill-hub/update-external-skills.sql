-- Skill 市集 · 外部 Skill 推荐
-- 已有站点在 Supabase SQL Editor 中完整执行本文件。可重复运行。
-- 全新站点执行 setup.sql 即可，其中已包含本次更新。
-- 所有修改在同一事务中完成。任一检查失败会回滚本次更新。

begin;

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
      and recommendation ~ '[^[:space:]]'
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
-- 换来源需要新建条目；推荐理由、试用记录等内容仍可修改。
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

commit;
