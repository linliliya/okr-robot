// 普通适配器测试无外部依赖；可用 PGLITE_MODULE 指定本地 @electric-sql/pglite 运行真实 SQL 测试。
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const K = require('../knowledge.js');
const source = fs.readFileSync(require.resolve('../db.js'), 'utf8');
const copy = x => JSON.parse(JSON.stringify(x));
const config = { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_ANON_KEY: 'public-test-key' };
const document = { title: '示例课程说明', project: '', kind: 'detail', channel: 'all', content: '退款需要先核对订单', status: 'active', source_note: '测试构造资料' };

function loadDB({ demo = false, state = {}, storageFailure = () => false, execute = () => ({ data: [], error: null }) } = {}) {
  let storage = JSON.stringify(state);
  const requests = [];
  const client = {
    rpc(name, params) { const request = { type: 'rpc', name, params: copy(params) }; requests.push(request); return Promise.resolve().then(() => execute(request)); },
    from(table) {
      const request = { type: 'table', table, filters: [], orders: [] };
      const q = {
        select(fields) { request.fields = fields; return q; },
        order(field, opts) { request.orders.push([field, opts]); return q; },
        range(from, to) { request.range = [from, to]; return q; },
        eq(field, value) { request.filters.push([field, value]); return q; },
        maybeSingle() { request.single = true; return q; },
        then(resolve, reject) { requests.push(copy(request)); return Promise.resolve().then(() => execute(request)).then(resolve, reject); },
      };
      return q;
    },
  };
  const window = { location: { search: demo ? '?demo=1' : '' }, SalesKnowledge: K, supabase: { createClient: () => client } };
  vm.runInNewContext(source, { window, localStorage: { getItem: () => storage, setItem: (_k, value) => {
    if (storageFailure()) throw new Error('QuotaExceededError'); storage = value;
  } } });
  window.SalesDB.init(demo ? undefined : config);
  return { db: window.SalesDB, requests, stored: () => JSON.parse(storage) };
}

test('知识列表分页只选元数据，单篇编辑才取全文', async () => {
  const rows = Array.from({ length: 1001 }, (_, i) => ({ ...document, id: i + 1, version: 1, updated_at: '2026-01-01', char_count: 12, chunk_count: 1 }));
  const h = loadDB({ execute: r => ({ data: r.single ? rows[0] : rows.slice(r.range[0], r.range[1] + 1).map(d =>
    Object.fromEntries(r.fields.split(',').map(k => [k, d[k]]))) }) });
  const metadata = await h.db.listKnowledgeDocuments();
  assert.equal(metadata.length, 1001);
  assert(metadata.every(d => !('content' in d)));
  assert.deepEqual(h.requests.map(r => r.range), [[0, 999], [1000, 1999]]);
  assert(h.requests.every(r => r.fields !== '*' && !r.fields.split(',').includes('content')));
  assert.equal((await h.db.getKnowledgeDocument(1)).content, document.content);
  assert.equal(h.requests[2].fields, '*');
  assert.deepEqual(h.requests[2].filters, [['id', 1]]);
});

test('搜索只调用服务端 RPC，传递关键词和有限范围，不读取全库正文', async () => {
  const hit = { document_id: 1, chunk_id: 2, content: '退款条款', ordinal: 1, version: 2 };
  const h = loadDB({ execute: () => ({ data: [hit] }) });
  assert.deepEqual(copy(await h.db.searchKnowledge('Python退款', { project: ' Python ', channel: 'xiaoe', limit: 99 })), [hit]);
  assert.deepEqual(h.requests, [{ type: 'rpc', name: 'search_sales_knowledge', params: {
    p_terms: K.queryTerms('Python退款'), p_project: 'Python', p_channel: 'xiaoe', p_limit: 6,
  } }]);
});

test('保存只发一个原子 RPC，客户端校验长度和版本，不能注入自建分块', async () => {
  const h = loadDB({ execute: r => ({ data: { ...r.params.p_document, id: 8, version: 3 } }) });
  const saved = await h.db.saveKnowledgeDocument({ ...document, id: 8, version: 2, chunk_count: 999, chunks: ['假资料'] });
  assert.equal(saved.version, 3);
  assert.deepEqual(h.requests[0], { type: 'rpc', name: 'save_sales_knowledge', params: { p_document: { ...document, id: 8, version: 2 } } });
  await assert.rejects(h.db.saveKnowledgeDocument({ ...document, content: '甲'.repeat(200001) }), e => e.code === 'invalid');
  await assert.rejects(h.db.saveKnowledgeDocument({ ...document, id: 8 }), e => e.code === 'invalid');
  assert.equal(h.requests.length, 1);
});

test('缺知识表/RPC单独报 knowledge_missing，权限、网络、冲突不混为未初始化', async () => {
  for (const failure of [{ code: '42P01', message: 'missing table' }, { code: 'PGRST202', message: 'missing RPC' }]) {
    const h = loadDB({ execute: () => ({ error: failure }) });
    assert.equal(h.db.mode, 'supabase');
    await assert.rejects(h.db.listKnowledgeDocuments(), e => e.code === 'knowledge_missing' && /knowledge-setup\.sql/.test(e.message));
    await assert.rejects(h.db.searchKnowledge('退款'), e => e.code === 'knowledge_missing');
  }
  for (const [failure, expected] of [[{ code: '42501', message: 'sales_members_only' }, 'not_member'],
    [{ code: 'P0001', message: 'knowledge_conflict' }, 'conflict'], [{ code: 'P0002', message: 'knowledge_not_found' }, 'not_found'],
    [{ code: '22023', message: 'knowledge_invalid_document' }, 'invalid'], [new Error('Failed to fetch'), 'network']]) {
    const h = loadDB({ execute: () => { if (failure instanceof Error) throw failure; return { error: failure }; } });
    await assert.rejects(h.db.saveKnowledgeDocument(document), e => e.code === expected);
  }
});

test('旧演示数据可直接使用知识库，修改原子替换分块，冲突不改内容，停用即退出检索', async () => {
  const h = loadDB({ demo: true, state: { scripts: [{ id: 100, scene: '原有话术' }], seq: 101 } });
  assert.equal((await h.db.listScripts())[0].id, 100);
  assert.deepEqual(copy(await h.db.listKnowledgeDocuments()), []);
  const first = await h.db.saveKnowledgeDocument({ ...document, content: '退款旧规则'.repeat(220) });
  assert(first.chunk_count > 1);
  assert.equal((await h.db.searchKnowledge('退款')).length, first.chunk_count);
  const revised = await h.db.saveKnowledgeDocument({ ...first, content: '学习时间新规则' });
  assert.equal(revised.version, 2);
  assert.equal(revised.chunk_count, 1);
  await assert.rejects(h.db.saveKnowledgeDocument({ ...first, content: '不能覆盖的新内容' }), e => e.code === 'conflict');
  assert.equal((await h.db.getKnowledgeDocument(first.id)).content, '学习时间新规则');
  assert.equal((await h.db.searchKnowledge('退款')).length, 0);
  assert.equal((await h.db.searchKnowledge('学习时间'))[0].version, 2);
  await h.db.saveKnowledgeDocument({ ...revised, status: 'archived' });
  assert.equal((await h.db.searchKnowledge('学习时间')).length, 0);
  assert.equal(h.stored().knowledgeChunks.length, 1);
  assert(!('content' in (await h.db.listKnowledgeDocuments())[0]));
});

test('演示检索与约定一致，通用/项目和渠道不会互串', async () => {
  const { db } = loadDB({ demo: true });
  await db.saveKnowledgeDocument(document);
  await db.saveKnowledgeDocument({ ...document, project: 'Python', channel: 'xiaoe' });
  await db.saveKnowledgeDocument({ ...document, project: 'AI', channel: 'bilibili' });
  assert.equal((await db.searchKnowledge('退款')).length, 1);
  const hits = copy(await db.searchKnowledge('退款', { project: 'Python', channel: 'xiaoe' }));
  assert.equal(hits.length, 2);
  assert(hits.every(c => c.project === '' || c.project === 'Python'));
  assert(hits.every(c => c.ordinal === 1 && c.version === 1 && c.source_note === document.source_note));
});

test('演示存储容量不足时明确报错，正文、版本和检索结果都保持原样', async () => {
  let fail = false;
  const { db } = loadDB({ demo: true, storageFailure: () => fail });
  const saved = await db.saveKnowledgeDocument(document);
  fail = true;
  await assert.rejects(db.saveKnowledgeDocument({ ...saved, content: '保存失败的新内容' }), /浏览器没有保存成功/);
  assert.equal((await db.getKnowledgeDocument(saved.id)).content, document.content);
  assert.equal((await db.getKnowledgeDocument(saved.id)).version, 1);
  assert.equal((await db.searchKnowledge('退款')).length, 1);
});

test('真实 SQL 验证幂等、原子重建、RLS和范围过滤', { skip: !process.env.PGLITE_MODULE }, async () => {
  const { PGlite } = require(process.env.PGLITE_MODULE);
  const pg = new PGlite();
  const member = '00000000-0000-0000-0000-000000000001', stranger = '00000000-0000-0000-0000-000000000002';
  try {
    await pg.exec(`create role anon; create role authenticated; create schema auth;
      create table auth.users(id uuid primary key, email text);
      create table public.profiles(id uuid primary key, display_name text);
      create function auth.uid() returns uuid language sql stable as $$select nullif(current_setting('request.jwt.claim.sub', true), '')::uuid$$;
      grant usage on schema auth, public to anon, authenticated;`);
    await pg.exec(fs.readFileSync(require.resolve('../setup.sql'), 'utf8'));
    const sql = fs.readFileSync(require.resolve('../knowledge-setup.sql'), 'utf8');
    await pg.exec(sql); await pg.exec(sql);
    await pg.query('insert into auth.users values($1, $2), ($3, $4)', [member, 'member@example.test', stranger, 'stranger@example.test']);
    await pg.query('insert into public.profiles values($1, $2), ($3, $4)', [member, '测试成员', stranger, '非成员']);
    await pg.query('insert into public.sales_members(email,user_id) values($1,$2)', ['member@example.test', member]);
    const identity = async id => { await pg.query("select set_config('request.jwt.claim.sub',$1,false)", [id]); };
    await identity(member); await pg.exec('set role authenticated');
    const save = async d => (await pg.query('select * from public.save_sales_knowledge($1::jsonb)', [JSON.stringify(d)])).rows[0];
    const search = async (query, project = '', channel = 'all') => (await pg.query(
      'select * from public.search_sales_knowledge($1::text[],$2,$3,6)', [K.queryTerms(query), project, channel])).rows;
    const original = { ...document, content: '退款流程'.repeat(170) + '\n' + '课程安排😀'.repeat(220) };
    const first = await save(original);
    const parts = (await pg.query('select content from public.sales_knowledge_chunks where document_id=$1 order by ordinal', [first.id])).rows.map(r => r.content);
    assert.deepEqual(parts, K.chunkText(original.content), 'SQL与浏览器分块一致');
    assert.equal(first.chunk_count, parts.length);
    await assert.rejects(pg.query('update public.sales_knowledge_documents set content=$1 where id=$2', ['绕过RPC', first.id]), e => e.code === '42501');
    await assert.rejects(pg.query('delete from public.sales_knowledge_chunks'), e => e.code === '42501');
    const revision = await save({ ...first, content: '学习时间新规则' });
    assert.equal(revision.version, 2);
    await assert.rejects(save({ ...first, content: '陈旧版本' }), /knowledge_conflict/);
    assert.equal((await pg.query('select count(*)::int as n from public.sales_knowledge_chunks where document_id=$1', [first.id])).rows[0].n, 1);
    assert.equal((await search('退款')).length, 0);
    await save({ ...document, project: 'Python', channel: 'xiaoe' });
    await save({ ...document, project: 'AI', channel: 'bilibili' });
    assert.equal((await search('退款')).length, 0);
    assert.equal((await search('退款', 'Python', 'xiaoe')).length, 1);
    assert.equal((await search('退款', 'Python', 'bilibili')).length, 0);
    await save({ ...revision, status: 'archived' });
    assert.equal((await search('学习时间')).length, 0);
    await assert.rejects(save({ ...document, content: '甲'.repeat(200001) }), e => e.code === '22023');
    const maxDoc = await save({ ...document, title: '上限验证', content: '上限匹配词'.repeat(40000), source_note: '来源'.repeat(500) });
    assert.equal(maxDoc.char_count, 200000);
    const limitedHits = await search('上限匹配词');
    assert(limitedHits.length > 0 && limitedHits.length <= 6);
    assert(limitedHits.every(h => Array.from(h.content).length <= 900));
    assert(limitedHits.reduce((n, h) => n + Array.from(h.content + h.title + h.project + h.source_note).length + 100, 0) <= 6000);
    const longDoc = await save({ ...document, title: '甲训练营销转直播', content: '填充正文'.repeat(8000) + '\n回放保留60天', project: '回放测试' });
    const otherDoc = await save({ ...document, title: '回放说明', content: '回放可在课程页面找到', project: '回放测试' });
    const replayHits = await search('错过直播了，回放保留多久', '回放测试');
    assert(replayHits.some(c => c.document_id === longDoc.id && c.content.includes('回放保留60天')));
    assert(replayHits.filter(c => c.document_id === longDoc.id).every(c => c.content.includes('回放')));
    assert(replayHits.some(c => c.document_id === otherDoc.id));
    const sqlChunks = (await pg.query(`select c.id as chunk_id,c.document_id,c.content,c.ordinal,c.version,
      d.title,d.project,d.channel,d.status,d.kind,d.source_note from public.sales_knowledge_chunks c
      join public.sales_knowledge_documents d on d.id=c.document_id`)).rows;
    const jsHits = K.rankChunks(sqlChunks, '错过直播了，回放保留多久', { project: '回放测试' });
    assert.deepEqual(replayHits.map(c => c.chunk_id), jsHits.map(c => c.chunk_id), 'SQL和演示排序及标题降级策略一致');
    assert.deepEqual((await search('直播', '回放测试')).map(c => [c.document_id, c.ordinal]), [[longDoc.id, 1]]);

    // 强制分块写入失败，检查原文、版本与旧块一并回滚。
    const before = await save({ ...document, content: '回滚前原文' });
    await pg.exec(`reset role;
      create function public.test_reject_chunk() returns trigger language plpgsql as $$begin
        if new.content like '%强制失败%' then raise exception '模拟分块写入失败'; end if; return new; end;$$;
      create trigger test_reject_chunk before insert on public.sales_knowledge_chunks for each row execute function public.test_reject_chunk();
      set role authenticated;`);
    await assert.rejects(save({ ...before, content: '强制失败' }), /模拟分块写入失败/);
    const after = (await pg.query('select content,version from public.sales_knowledge_documents where id=$1', [before.id])).rows[0];
    assert.deepEqual(after, { content: before.content, version: before.version });
    assert.equal((await search('回滚前原文'))[0].content, before.content);

    await identity(stranger);
    assert.equal((await pg.query('select * from public.sales_knowledge_documents')).rows.length, 0);
    assert.equal((await pg.query('select * from public.sales_knowledge_chunks')).rows.length, 0);
    await assert.rejects(save(document), e => e.code === '42501');
    await assert.rejects(search('退款'), e => e.code === '42501');
    await pg.exec('reset role; set role anon');
    await assert.rejects(search('退款'), e => e.code === '42501');
  } finally { await pg.close(); }
});
