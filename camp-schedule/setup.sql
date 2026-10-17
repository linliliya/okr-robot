-- Shared schedules. Reuses the internal operations team's existing membership.
-- Apply to the existing Supabase project before publishing the frontend.
create table if not exists public.schedule_camps (
  id uuid primary key default gen_random_uuid(),
  program text not null check (program in ('ai', 'crawler')),
  term integer not null check (term between 1 and 10000),
  open_date date not null check (open_date between date '2000-01-01' and date '2100-12-31'),
  skips integer[] not null default '{}' check (
    cardinality(skips) <= 12 and array_position(skips, null) is null
    and skips <@ array[1,2,3,4,5,6,7,8,9,10,11,12]),
  version integer not null default 1,
  updated_by uuid references auth.users(id) on delete set null,
  updated_at timestamptz not null default now(),
  deleted_at timestamptz,
  unique (id, program)
);
create unique index if not exists schedule_camps_active_term
  on public.schedule_camps(program, term) where deleted_at is null;

create table if not exists public.schedule_tasks (
  camp_id uuid not null,
  program text not null,
  task_key text not null check (
    length(task_key) <= 500
    and task_key ~ '^(open|(week|rest)-(1[0-2]|[1-9]))\|(all|[0-6])\|.+$'),
  completed boolean not null,
  version integer not null default 1,
  updated_by uuid references auth.users(id) on delete set null,
  updated_at timestamptz not null default now(),
  primary key (camp_id, task_key),
  foreign key (camp_id, program) references public.schedule_camps(id, program)
);
create index if not exists schedule_tasks_program on public.schedule_tasks(program);

alter table public.schedule_camps enable row level security;
alter table public.schedule_tasks enable row level security;
revoke all on public.schedule_camps, public.schedule_tasks from public, anon, authenticated;
grant select on public.schedule_camps, public.schedule_tasks to authenticated;

drop policy if exists schedule_camps_member_read on public.schedule_camps;
create policy schedule_camps_member_read on public.schedule_camps for select to authenticated
  using ((select public.is_camp_review_member()));
drop policy if exists schedule_tasks_member_read on public.schedule_tasks;
create policy schedule_tasks_member_read on public.schedule_tasks for select to authenticated
  using ((select public.is_camp_review_member()) and exists (
    select 1 from public.schedule_camps c where c.id = camp_id and c.deleted_at is null));

create or replace function public.schedule_require_member(p_program text)
returns void language plpgsql security definer set search_path = '' as $$
begin
  if auth.uid() is null or not public.is_camp_review_member() then
    raise exception 'schedule_members_only' using errcode = '42501';
  end if;
  if p_program is null or p_program not in ('ai', 'crawler') then
    raise exception 'schedule_invalid_program' using errcode = '22023';
  end if;
end;
$$;

create or replace function public.schedule_snapshot(p_program text)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare result jsonb;
begin
  perform public.schedule_require_member(p_program);
  select jsonb_build_object(
    'camps', coalesce((select jsonb_agg(to_jsonb(c) order by c.open_date, c.term)
      from public.schedule_camps c where c.program = p_program and c.deleted_at is null), '[]'::jsonb),
    'tasks', coalesce((select jsonb_agg(to_jsonb(t)) from public.schedule_tasks t
      join public.schedule_camps c on c.id = t.camp_id
      where c.program = p_program and c.deleted_at is null), '[]'::jsonb)
  ) into result;
  return result;
end;
$$;

create or replace function public.save_schedule_camp(
  p_program text, p_id uuid, p_version integer, p_term integer, p_open_date date, p_skips integer[])
returns jsonb language plpgsql security definer set search_path = '' as $$
declare result public.schedule_camps; normalized integer[];
begin
  perform public.schedule_require_member(p_program);
  if p_skips is null or array_position(p_skips, null) is not null then
    raise exception 'schedule_invalid_skips' using errcode = '22023';
  end if;
  select coalesce(array_agg(distinct n order by n), '{}'::integer[]) into normalized from unnest(p_skips) n;
  if p_id is null then
    insert into public.schedule_camps(program, term, open_date, skips, updated_by)
      values(p_program, p_term, p_open_date, normalized, auth.uid()) returning * into result;
  else
    update public.schedule_camps set term = p_term, open_date = p_open_date, skips = normalized,
      version = version + 1, updated_by = auth.uid(), updated_at = clock_timestamp()
      where id = p_id and program = p_program and version = p_version and deleted_at is null
      returning * into result;
    if not found then raise exception 'schedule_conflict' using errcode = 'P0001'; end if;
  end if;
  return to_jsonb(result);
end;
$$;

create or replace function public.delete_schedule_camp(p_program text, p_id uuid, p_version integer)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare result public.schedule_camps;
begin
  perform public.schedule_require_member(p_program);
  update public.schedule_camps set deleted_at = clock_timestamp(), updated_at = clock_timestamp(),
    updated_by = auth.uid(), version = version + 1
    where id = p_id and program = p_program and version = p_version and deleted_at is null
    returning * into result;
  if not found then raise exception 'schedule_conflict' using errcode = 'P0001'; end if;
  return to_jsonb(result);
end;
$$;

create or replace function public.set_schedule_task(
  p_program text, p_camp_id uuid, p_task_key text, p_version integer, p_completed boolean)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare result public.schedule_tasks;
begin
  perform public.schedule_require_member(p_program);
  perform 1 from public.schedule_camps where id = p_camp_id and program = p_program
    and deleted_at is null for share;
  if not found then raise exception 'schedule_conflict' using errcode = 'P0001'; end if;
  if p_version = 0 then
    insert into public.schedule_tasks(camp_id, program, task_key, completed, updated_by)
      values(p_camp_id, p_program, p_task_key, p_completed, auth.uid())
      on conflict (camp_id, task_key) do nothing returning * into result;
  else
    update public.schedule_tasks set completed = p_completed, version = version + 1,
      updated_at = clock_timestamp(), updated_by = auth.uid()
      where camp_id = p_camp_id and program = p_program and task_key = p_task_key and version = p_version
      returning * into result;
  end if;
  if result.camp_id is null then raise exception 'schedule_conflict' using errcode = 'P0001'; end if;
  return to_jsonb(result);
end;
$$;

-- Import creates missing terms only. It never overwrites an existing team term
-- or its task progress. One RPC is one transaction, including all task rows.
create or replace function public.import_schedule_camps(p_program text, p_camps jsonb)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare entry jsonb; saved jsonb; task text; normalized integer[];
  inserted_terms integer[] := '{}'; skipped_terms integer[] := '{}';
begin
  perform public.schedule_require_member(p_program);
  if jsonb_typeof(p_camps) is distinct from 'array' or jsonb_array_length(p_camps) > 200 then
    raise exception 'schedule_invalid_import' using errcode = '22023';
  end if;
  for entry in select value from jsonb_array_elements(p_camps) loop
    if jsonb_typeof(entry->'skips') is distinct from 'array'
       or jsonb_typeof(entry->'completed_tasks') is distinct from 'array'
       or jsonb_array_length(entry->'completed_tasks') > 10000 then
      raise exception 'schedule_invalid_import' using errcode = '22023';
    end if;
    select coalesce(array_agg(value::integer), '{}'::integer[]) into normalized
      from jsonb_array_elements_text(entry->'skips');
    begin
      saved := public.save_schedule_camp(p_program, null, null, (entry->>'term')::integer,
        (entry->>'open_date')::date, normalized);
    exception when unique_violation then
      skipped_terms := array_append(skipped_terms, (entry->>'term')::integer);
      continue;
    end;
    for task in select distinct value from jsonb_array_elements_text(entry->'completed_tasks') loop
      perform public.set_schedule_task(p_program, (saved->>'id')::uuid, task, 0, true);
    end loop;
    inserted_terms := array_append(inserted_terms, (entry->>'term')::integer);
  end loop;
  return jsonb_build_object('inserted', inserted_terms, 'skipped', skipped_terms);
end;
$$;

revoke all on function public.schedule_require_member(text), public.schedule_snapshot(text),
  public.save_schedule_camp(text,uuid,integer,integer,date,integer[]),
  public.delete_schedule_camp(text,uuid,integer),
  public.set_schedule_task(text,uuid,text,integer,boolean),
  public.import_schedule_camps(text,jsonb) from public, anon, authenticated;
grant execute on function public.schedule_snapshot(text),
  public.save_schedule_camp(text,uuid,integer,integer,date,integer[]),
  public.delete_schedule_camp(text,uuid,integer),
  public.set_schedule_task(text,uuid,text,integer,boolean),
  public.import_schedule_camps(text,jsonb) to authenticated;

do $$
begin
  if exists (select 1 from pg_publication where pubname = 'supabase_realtime') then
    if not exists (select 1 from pg_publication_tables where pubname = 'supabase_realtime'
      and schemaname = 'public' and tablename = 'schedule_camps') then
      alter publication supabase_realtime add table public.schedule_camps;
    end if;
    if not exists (select 1 from pg_publication_tables where pubname = 'supabase_realtime'
      and schemaname = 'public' and tablename = 'schedule_tasks') then
      alter publication supabase_realtime add table public.schedule_tasks;
    end if;
  end if;
end;
$$;
