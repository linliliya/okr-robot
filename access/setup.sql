-- Central access policy. Seed the approved roster separately; do not publish it.
create schema if not exists app_private;
revoke all on schema app_private from public, anon, authenticated;
create table if not exists app_private.staff_access (
 email text primary key check (email = lower(btrim(email)) and email ~ '^[^@[:space:]]+@[^@[:space:]]+\.[^@[:space:]]+$'),
 role text not null check (role in ('admin','analyst','staff')),
 active boolean not null default true,
 updated_at timestamptz not null default now()
);
alter table app_private.staff_access enable row level security;
revoke all on app_private.staff_access from public, anon, authenticated;

create or replace function public.my_app_access()
returns jsonb language sql stable security definer set search_path = '' as $$
 select coalesce((select jsonb_build_object('allowed',true,'role',s.role,'email',s.email,
   'apps',case when s.role in ('admin','analyst') then
     '["skill-hub","okr","crawler","ai","camp-review","sales-script"]'::jsonb
     else '["skill-hub","okr","crawler","ai"]'::jsonb end)
 from app_private.staff_access s join auth.users u on lower(u.email)=s.email
 where u.id=auth.uid() and s.active and u.email_confirmed_at is not null
   and (u.banned_until is null or u.banned_until <= now()) and not coalesce(u.is_anonymous,false)),
 '{"allowed":false,"role":null,"apps":[]}'::jsonb);
$$;
create or replace function public.has_app_access(p_app text)
returns boolean language sql stable security definer set search_path = '' as $$
 select coalesce((public.my_app_access()->'apps') ? p_app,false);
$$;
create or replace function public.is_access_admin()
returns boolean language sql stable security definer set search_path = '' as $$
 select coalesce(public.my_app_access()->>'role'='admin',false);
$$;
create or replace function public.is_camp_review_member()
returns boolean language sql stable security definer set search_path = '' as $$
 select public.has_app_access('camp-review');
$$;
create or replace function public.is_sales_member()
returns boolean language sql stable security definer set search_path = '' as $$
 select public.has_app_access('sales-script');
$$;

-- Auth-side restriction also applies to callers who bypass the website.
create or replace function app_private.check_signup_email()
returns trigger language plpgsql security definer set search_path = '' as $$
begin
 if tg_op='UPDATE' then
   if lower(new.email) is distinct from lower(old.email) then
     raise exception '邮箱属于管理员配置的身份，不能自行更改' using errcode='42501';
   end if;
   return new;
 end if;
 if new.email is null or coalesce(new.is_anonymous,false) or not exists
   (select 1 from app_private.staff_access s where s.email=lower(btrim(new.email)) and s.active) then
   raise exception '仅限管理员名单中的邮箱注册' using errcode='42501';
 end if;
 return new;
end;
$$;
drop trigger if exists staff_signup_allowlist on auth.users;
create trigger staff_signup_allowlist before insert or update of email on auth.users
 for each row execute function app_private.check_signup_email();

create or replace function public.list_staff_access()
returns jsonb language plpgsql stable security definer set search_path='' as $$
begin
 if not public.is_access_admin() then raise exception '仅管理员可管理账号权限' using errcode='42501'; end if;
 return coalesce((select jsonb_agg(jsonb_build_object('email',s.email,'role',s.role,'active',s.active,
   'registered',u.id is not null,'confirmed',u.email_confirmed_at is not null) order by s.email)
   from app_private.staff_access s left join auth.users u on lower(u.email)=s.email),'[]'::jsonb);
end;
$$;
create or replace function public.set_staff_access(p_email text,p_role text,p_active boolean)
returns void language plpgsql security definer set search_path='' as $$
declare e text:=lower(btrim(p_email));
begin
 if not public.is_access_admin() then raise exception '仅管理员可管理账号权限' using errcode='42501'; end if;
 if p_role not in ('staff','analyst') or p_role is null or p_active is null or e is null then
   raise exception '权限格式不正确' using errcode='22023'; end if;
 if exists(select 1 from app_private.staff_access where email=e and role='admin') then
   raise exception '不能修改管理员自身权限' using errcode='42501'; end if;
 insert into app_private.staff_access(email,role,active) values(e,p_role,p_active)
 on conflict(email) do update set role=excluded.role,active=excluded.active,updated_at=now();
 update auth.users set banned_until=case when p_active then null else timestamptz '2099-12-31' end
 where lower(email)=e;
 if not p_active then
   update auth.refresh_tokens set revoked=true where user_id in (select id::text from auth.users where lower(email)=e);
   update auth.sessions set not_after=now() where user_id in (select id from auth.users where lower(email)=e);
 end if;
end;
$$;

-- Preserve existing business policies and add a mandatory gate to every operation.
do $$
declare t text;
begin
 foreach t in array array['profiles','skills','versions'] loop
   execute format('drop policy if exists staff_access_required on public.%I',t);
   execute format('create policy staff_access_required on public.%I as restrictive for all to authenticated using ((select public.has_app_access(''skill-hub''))) with check ((select public.has_app_access(''skill-hub'')))',t);
 end loop;
end $$;
drop policy if exists staff_files_required on storage.objects;
create policy staff_files_required on storage.objects as restrictive for all to authenticated
 using (bucket_id <> 'skills' or (select public.has_app_access('skill-hub')))
 with check (bucket_id <> 'skills' or (select public.has_app_access('skill-hub')));
create or replace function public.increment_downloads(p_skill_id bigint)
returns void language plpgsql security definer set search_path='' as $$
begin
 if not public.has_app_access('skill-hub') then raise exception 'access_denied' using errcode='42501'; end if;
 update public.skills set downloads=downloads+1 where id=p_skill_id and source_type='upload';
end;
$$;

create or replace function public.schedule_require_member(p_program text)
returns void language plpgsql security definer set search_path='' as $$
begin
 if p_program is null or p_program not in ('ai','crawler') then raise exception 'schedule_invalid_program' using errcode='22023'; end if;
 if not public.has_app_access(p_program) then raise exception 'schedule_members_only' using errcode='42501'; end if;
end;
$$;
drop policy if exists schedule_camps_member_read on public.schedule_camps;
create policy schedule_camps_member_read on public.schedule_camps for select to authenticated
 using (public.has_app_access(program));
drop policy if exists schedule_tasks_member_read on public.schedule_tasks;
create policy schedule_tasks_member_read on public.schedule_tasks for select to authenticated
 using (public.has_app_access(program) and exists(select 1 from public.schedule_camps c where c.id=camp_id and c.deleted_at is null));

-- Retire old membership editors so an analyst cannot grant privileges to others.
revoke insert,update,delete on public.camp_review_members,public.sales_members from authenticated,anon;
create or replace function public.add_sales_member(p_email text)
returns void language plpgsql security definer set search_path='' as $$
begin
 raise exception '请在统一账号权限页面管理名单' using errcode='42501';
end;
$$;
revoke all on function app_private.check_signup_email() from public,anon,authenticated;
revoke all on function public.my_app_access(),public.has_app_access(text),public.is_access_admin(),public.list_staff_access(),public.set_staff_access(text,text,boolean),public.is_camp_review_member(),public.is_sales_member(),public.increment_downloads(bigint),public.add_sales_member(text) from public,anon;
grant execute on function public.my_app_access(),public.has_app_access(text),public.is_access_admin(),public.list_staff_access(),public.set_staff_access(text,text,boolean),public.is_camp_review_member(),public.is_sales_member(),public.increment_downloads(bigint) to authenticated;

-- A single-email check gives the form a useful error without exposing the roster.
create or replace function public.can_register_email(p_email text)
returns boolean language sql stable security definer set search_path='' as $$
 select exists(select 1 from app_private.staff_access where email=lower(btrim(p_email)) and active);
$$;
revoke all on function public.can_register_email(text) from public;
grant execute on function public.can_register_email(text) to anon,authenticated;
