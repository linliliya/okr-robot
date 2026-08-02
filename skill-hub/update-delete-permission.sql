-- ============================================================
-- 增量更新：允许上传人删除自己上传的版本（2026-08-02）
-- 用法：Supabase 后台 → SQL Editor → New query → 整段粘贴 → Run
-- （已经运行过 setup.sql 的项目跑这个即可；全新安装直接跑 setup.sql，
--   里面已包含这些规则，不需要再跑本文件）
-- ============================================================

create policy "可删除自己上传的版本" on public.versions
  for delete to authenticated using (auth.uid() = uploader_id);

-- 最后一个版本被删除时，skill 本身也要能删掉（仅限 skill 作者）
create policy "作者可删除自己的 skill" on public.skills
  for delete to authenticated using (auth.uid() = author_id);

-- 删除版本时同步删除存储的 zip 文件（只能删自己传的）
create policy "可删除自己上传的文件" on storage.objects
  for delete to authenticated
  using (bucket_id = 'skills' and (owner = auth.uid() or owner_id = auth.uid()::text));
