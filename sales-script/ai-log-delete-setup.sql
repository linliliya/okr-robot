-- 销售话术复盘 · 只删除本人提问记录
-- 已完成 setup.sql 的数据库执行此增量；可重复运行，不删除已有记录。
begin;

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

commit;
