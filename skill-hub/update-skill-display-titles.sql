-- 为 Skill 增加独立的功能标题，原有 title 继续保存 Skill 名称。
-- 可以重复执行。只回填下面四个已确认的名称，不覆盖已有功能标题。

begin;

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

commit;
