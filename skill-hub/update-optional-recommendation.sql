-- 已有外部分享功能的站点执行本文件，取消推荐原因必填。
-- 可重复执行。保留 recommendation 列和已有内容，只更新来源约束。

begin;

alter table public.skills
  drop constraint if exists skills_external_source_check;

alter table public.skills
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
  );

commit;
