-- 销售话术复盘 · 课程资料知识库
-- 先执行 setup.sql（依赖 public.profiles、public.is_sales_member）。
-- 整段在 Supabase SQL Editor 执行；可重复运行，不删除已有资料。
-- 资料只通过 save_sales_knowledge 原子保存；搜索为关键词/中文双字匹配，无 embedding。
begin;

create table if not exists public.sales_knowledge_documents (
  id bigint generated always as identity primary key,
  title text not null check (char_length(title) between 1 and 240),
  project text not null default '' check (char_length(project) <= 120),
  kind text not null default 'other' check (kind in ('detail','live','faq','other')),
  channel text not null default 'all' check (channel in ('all','xiaoe','bilibili')),
  content text not null check (char_length(content) between 1 and 200000),
  status text not null default 'active' check (status in ('active','archived')),
  source_note text not null default '' check (char_length(source_note) <= 1000),
  version integer not null default 1,
  char_count integer not null default 0,
  chunk_count integer not null default 0,
  created_by uuid references public.profiles(id) on delete set null,
  updated_by uuid references public.profiles(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists sales_knowledge_scope_idx on public.sales_knowledge_documents (status, project, channel);

create table if not exists public.sales_knowledge_chunks (
  id bigint generated always as identity primary key,
  document_id bigint not null references public.sales_knowledge_documents(id) on delete cascade,
  ordinal integer not null,
  content text not null check (char_length(content) between 1 and 900),
  version integer not null,
  unique(document_id, ordinal)
);

alter table public.sales_knowledge_documents enable row level security;
alter table public.sales_knowledge_chunks enable row level security;
drop policy if exists "成员可查看资料" on public.sales_knowledge_documents;
create policy "成员可查看资料" on public.sales_knowledge_documents
  for select to authenticated using (public.is_sales_member());
drop policy if exists "成员可查看资料块" on public.sales_knowledge_chunks;
create policy "成员可查看资料块" on public.sales_knowledge_chunks
  for select to authenticated using (public.is_sales_member());
-- 禁止绕过 RPC 修改正文或分块，避免两者不一致。
revoke all on public.sales_knowledge_documents, public.sales_knowledge_chunks from public, anon, authenticated;
grant select on public.sales_knowledge_documents, public.sales_knowledge_chunks to authenticated;
revoke all on sequence public.sales_knowledge_documents_id_seq, public.sales_knowledge_chunks_id_seq from public, anon, authenticated;

create or replace function public.save_sales_knowledge(p_document jsonb)
returns public.sales_knowledge_documents
language plpgsql security definer set search_path = pg_catalog, public as $$
declare
  d public.sales_knowledge_documents;
  doc_id bigint := nullif(p_document->>'id', '')::bigint;
  expected_version integer := nullif(p_document->>'version', '')::integer;
  doc_title text := btrim(coalesce(p_document->>'title', ''));
  doc_project text := btrim(coalesce(p_document->>'project', ''));
  doc_kind text := coalesce(p_document->>'kind', 'other');
  doc_channel text := coalesce(p_document->>'channel', 'all');
  doc_content text := coalesce(p_document->>'content', '');
  doc_status text := coalesce(p_document->>'status', 'active');
  doc_source text := btrim(coalesce(p_document->>'source_note', ''));
  content_length integer := char_length(doc_content);
  start_at integer := 1;
  end_at integer;
  newline_at integer;
  part text;
  part_count integer := 0;
begin
  if auth.uid() is null or not coalesce(public.is_sales_member(), false) then
    raise exception using errcode = '42501', message = 'sales_members_only';
  end if;
  if p_document is null or jsonb_typeof(p_document) <> 'object' or doc_title = ''
    or char_length(doc_title) > 240 or char_length(doc_project) > 120 or char_length(doc_source) > 1000
    or content_length > 200000 or doc_content ~ '^\s*$'
    or doc_kind not in ('detail','live','faq','other') or doc_channel not in ('all','xiaoe','bilibili')
    or doc_status not in ('active','archived') then
    raise exception using errcode = '22023', message = 'knowledge_invalid_document';
  end if;

  if doc_id is null then
    insert into public.sales_knowledge_documents
      (title, project, kind, channel, content, status, source_note, char_count, created_by, updated_by)
    values (doc_title, doc_project, doc_kind, doc_channel, doc_content, doc_status, doc_source, content_length, auth.uid(), auth.uid())
    returning * into d;
  else
    if expected_version is null or expected_version < 1 then
      raise exception using errcode = '22023', message = 'knowledge_version_required';
    end if;
    update public.sales_knowledge_documents set title = doc_title, project = doc_project, kind = doc_kind,
      channel = doc_channel, content = doc_content, status = doc_status, source_note = doc_source,
      char_count = content_length, version = version + 1, updated_by = auth.uid(), updated_at = now()
    where id = doc_id and version = expected_version returning * into d;
    if not found then
      if not exists (select 1 from public.sales_knowledge_documents where id = doc_id) then
        raise exception using errcode = 'P0002', message = 'knowledge_not_found';
      end if;
      raise exception using errcode = 'P0001', message = 'knowledge_conflict';
    end if;
  end if;

  delete from public.sales_knowledge_chunks where document_id = d.id;
  -- 每块最多 900 字，相邻块重叠 100 字。优先在后 300 字的换行处结束。
  while start_at <= content_length loop
    end_at := least(start_at + 899, content_length);
    if end_at < content_length then
      newline_at := strpos(reverse(substring(doc_content from start_at + 600 for 300)), E'\n');
      if newline_at > 0 then end_at := start_at + 900 - newline_at; end if;
    end if;
    part := substring(doc_content from start_at for end_at - start_at + 1);
    if part !~ '^\s*$' then
      part_count := part_count + 1;
      insert into public.sales_knowledge_chunks(document_id, ordinal, content, version)
        values(d.id, part_count, part, d.version);
    end if;
    exit when end_at = content_length;
    start_at := end_at - 99;
  end loop;
  update public.sales_knowledge_documents set chunk_count = part_count where id = d.id returning * into d;
  return d;
end;
$$;
revoke all on function public.save_sales_knowledge(jsonb) from public, anon;
grant execute on function public.save_sales_knowledge(jsonb) to authenticated;

create or replace function public.search_sales_knowledge(
  p_terms text[], p_project text default '', p_channel text default 'all', p_limit integer default 6
)
returns table(chunk_id bigint, document_id bigint, title text, project text, kind text,
  content text, ordinal integer, version integer, source_note text)
language plpgsql stable security definer set search_path = pg_catalog, public as $$
begin
  if auth.uid() is null or not coalesce(public.is_sales_member(), false) then
    raise exception using errcode = '42501', message = 'sales_members_only';
  end if;
  if coalesce(p_channel, 'all') not in ('all','xiaoe','bilibili') then
    raise exception using errcode = '22023', message = 'knowledge_invalid_channel';
  end if;
  return query
  with terms as (
    select distinct lower(btrim(t)) as term from unnest(p_terms[1:64]) t
    where char_length(btrim(t)) between 1 and 64
  ), scored as (
    select c.id as cid, d.id as did, d.title as heading, d.project as proj, d.kind as doc_kind,
      c.content as body, c.ordinal as ord, d.version as ver, d.source_note as source,
      sum(case when strpos(lower(c.content), t.term) > 0 then least(char_length(t.term), 8) else 0 end) as body_score,
      sum(case when strpos(lower(d.title), t.term) > 0 then least(char_length(t.term), 8) * 2 else 0 end) as title_score
    from public.sales_knowledge_chunks c
    join public.sales_knowledge_documents d on d.id = c.document_id and c.version = d.version
    cross join terms t
    where d.status = 'active' and (d.project = '' or d.project = btrim(coalesce(p_project, '')))
      and (d.channel = 'all' or d.channel = coalesce(p_channel, 'all'))
    group by c.id, d.id
  ), scoped as (
    select *, bool_or(body_score > 0) over(partition by did) as has_body_match,
      row_number() over(partition by did order by ord) as first_chunk
    from scored
  ), relevant as (
    -- 正文有命中时排除同文档的纯标题命中；只有标题匹配的资料最多取首块。
    select * from scoped where body_score > 0 or (title_score > 0 and not has_body_match and first_chunk = 1)
  ), budgeted as (
    select *, sum(char_length(body) + char_length(heading) + char_length(proj) + char_length(source) + 100)
      over(order by body_score desc, title_score desc, did desc, ord asc) as used_chars
    from relevant
  )
  select cid, did, heading, proj, doc_kind, body, ord, ver, source from budgeted
  where used_chars <= 6000 order by body_score desc, title_score desc, did desc, ord asc
  limit greatest(1, least(coalesce(p_limit, 6), 6));
end;
$$;
revoke all on function public.search_sales_knowledge(text[], text, text, integer) from public, anon;
grant execute on function public.search_sales_knowledge(text[], text, text, integer) to authenticated;

commit;
