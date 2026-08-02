-- ============================================================
-- Skill 共享平台 · Supabase 初始化脚本
-- 用法：Supabase 后台 → SQL Editor → 新建查询 → 整段粘贴 → Run
-- ============================================================

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
-- 团队内协作：任何登录用户都可以更新 skill 信息（比如上传新版本时刷新描述）
create policy "登录用户可更新" on public.skills
  for update to authenticated using (true);

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
create policy "登录用户可上传版本" on public.versions
  for insert to authenticated with check (auth.uid() = uploader_id);

-- ---------- 下载计数（绕过行级权限，安全地 +1） ----------
create or replace function public.increment_downloads(p_skill_id bigint)
returns void
language sql security definer set search_path = public
as $$
  update public.skills set downloads = downloads + 1 where id = p_skill_id;
$$;

-- ---------- 文件存储桶 ----------
insert into storage.buckets (id, name, public)
values ('skills', 'skills', false)
on conflict (id) do nothing;

create policy "登录用户可下载文件" on storage.objects
  for select to authenticated using (bucket_id = 'skills');
create policy "登录用户可上传文件" on storage.objects
  for insert to authenticated with check (bucket_id = 'skills');
