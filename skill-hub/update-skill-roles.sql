-- 为已有 Skill 增加适用岗位。已有站点先执行 update-external-skills.sql。
-- 可以重复执行；首次增加岗位时，已有条目保留为空数组。
-- 已有岗位按当前选项升级，不改变条目的其他内容。

begin;

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

commit;
