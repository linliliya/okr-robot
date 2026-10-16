// 全部使用内存数据。PGLITE_MODULE 可指向已安装的 PGlite，验证真实 PostgreSQL 权限。
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require.resolve('../db.js'), 'utf8');
const setupSQL = fs.readFileSync(require.resolve('../setup.sql'), 'utf8');
const migrationSQL = fs.readFileSync(require.resolve('../ai-log-delete-setup.sql'), 'utf8');
const migrationBlock = /-- BEGIN AI_LOG_DELETE[\s\S]*?-- END AI_LOG_DELETE/;
const copy = x => JSON.parse(JSON.stringify(x));

function fixture({ demo = false, state = {}, execute = () => ({ data: 7 }), storageFailure = () => false } = {}) {
  const requests = [];
  let stored = JSON.stringify(state), writes = 0;
  const browserWindow = { location: { search: demo ? '?demo=1' : '' }, supabase: { createClient: () => ({
    rpc(name, args) { requests.push({ name, args: copy(args) }); return Promise.resolve().then(() => execute(name, args)); },
    from() { throw new Error('删除接口不应通过客户端普通表写入'); },
  }) } };
  vm.runInNewContext(source, { window: browserWindow, localStorage: {
    getItem: () => stored,
    setItem(_key, value) { if (storageFailure()) throw new Error('QuotaExceededError'); stored = value; writes++; },
  } });
  browserWindow.SalesDB.init(demo ? undefined : { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_ANON_KEY: 'public-test-key' });
  return { db: browserWindow.SalesDB, requests, saved: () => JSON.parse(stored), writes: () => writes };
}

test('删除 API 只调用受 RLS 约束的 RPC，校验服务器确认的实际编号', async () => {
  for (const deleted of [7, '7']) {
    const h = fixture({ execute: () => ({ data: deleted }) });
    assert.deepEqual(copy(await h.db.deleteAiLog('7')), { id: 7 });
    assert.deepEqual(h.requests, [{ name: 'delete_sales_ai_log', args: { p_id: 7 } }]);
  }
  for (const deleted of [null, undefined, [], {}, { id: 7 }, [7], false, 0, 8, '7junk']) {
    const h = fixture({ execute: () => ({ data: deleted }) });
    await assert.rejects(h.db.deleteAiLog(7), e => e.code === 'unknown' && /未确认删除成功/.test(e.message));
  }
});

test('无效编号在本地拒绝，不向服务器发送删除请求', async () => {
  const h = fixture();
  for (const id of [undefined, null, '', ' ', '7 or 1=1', 0, -1, 1.5, true, [], {}, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1]) {
    await assert.rejects(h.db.deleteAiLog(id), e => e.code === 'invalid');
  }
  assert.equal(h.requests.length, 0);
});

test('缺迁移、未登录、非成员、他人记录、不存在与网络错误均不能伪装删除成功', async () => {
  for (const [error, code, message] of [
    [{ code: 'PGRST202', message: 'Could not find the function' }, 'ai_log_delete_missing', /ai-log-delete-setup\.sql/],
    [{ code: '42883', message: 'function does not exist' }, 'ai_log_delete_missing', /ai-log-delete-setup\.sql/],
    [{ code: '42501', message: 'ai_log_auth_required' }, 'auth', /重新登录/],
    [{ code: '42501', message: 'sales_members_only' }, 'not_member', /还没开通/],
    [{ code: '42501', message: 'ai_log_not_owner' }, 'forbidden', /自己创建/],
    [{ code: '42501', message: 'ai_log_delete_denied' }, 'ai_log_delete_missing', /记录未删除/],
    [{ code: '42501', message: 'permission denied for table sales_ai_logs' }, 'forbidden', /没有权限/],
    [{ code: 'P0002', message: 'ai_log_not_found' }, 'not_found', /已不存在/],
    [{ code: '22023', message: 'ai_log_invalid_id' }, 'invalid', /编号无效/],
    [new Error('Failed to fetch'), 'network', /网络连接失败/],
  ]) {
    const h = fixture({ execute: () => { if (error instanceof Error) throw error; return { error }; } });
    await assert.rejects(h.db.deleteAiLog(7), e => e.code === code && message.test(e.message));
  }
});

const demoState = () => ({
  logs: [
    { id: 7, created_by: 'demo-user', message: '本人合成提问', created_at: '2026-01-01', feedback: 'bad', feedback_note: '合成反馈' },
    { id: 8, created_by: 'other-user', message: '他人合成提问', created_at: '2026-01-01' },
    { id: 9, created_by: 'demo-user', message: '本人另一条提问', created_at: '2026-01-01' },
  ], scripts: [{ id: 1, scene: '合成场景' }], cases: [{ id: 2 }], reviews: [{ id: 3 }], settings: {}, seq: 10,
});

test('演示模式只删除本人的指定记录，并持久化，不改变其他业务数据', async () => {
  const h = fixture({ demo: true, state: demoState() });
  await assert.rejects(h.db.deleteAiLog(8), e => e.code === 'forbidden');
  await assert.rejects(h.db.deleteAiLog(999), e => e.code === 'not_found');
  assert.equal(h.writes(), 0);
  assert.deepEqual(copy(await h.db.deleteAiLog(7)), { id: 7 });
  assert.deepEqual(h.saved().logs.map(l => l.id), [8, 9]);
  assert.deepEqual(copy(await h.db.listAiLogs({ createdBy: 'demo-user' })).map(l => l.id), [9]);
  for (const field of ['scripts', 'cases', 'reviews', 'seq']) assert.deepEqual(h.saved()[field], demoState()[field]);
  await assert.rejects(h.db.deleteAiLog(7), e => e.code === 'not_found');
  assert.equal(h.writes(), 1);
});

test('演示持久化失败时记录和反馈仍留在内存及存储，可安全重试', async () => {
  let fail = true;
  const h = fixture({ demo: true, state: demoState(), storageFailure: () => fail });
  await assert.rejects(h.db.deleteAiLog(7), /记录仍然保留/);
  assert.deepEqual(h.saved(), demoState());
  const kept = (await h.db.listAiLogs({ all: true })).find(l => l.id === 7);
  assert.equal(kept.feedback, 'bad');
  assert.equal(kept.feedback_note, '合成反馈');
  assert.equal(h.writes(), 0);
  fail = false;
  await h.db.deleteAiLog(7);
  assert.equal((await h.db.listAiLogs({ all: true })).some(l => l.id === 7), false);
});

test('新装脚本与增量脚本的删除权限、函数定义保持一致', () => {
  assert.equal(setupSQL.match(migrationBlock)[0], migrationSQL.match(migrationBlock)[0]);
});

test('实际 PostgreSQL 验证增量幂等、本人权限、RLS不可绕过和零行报错', { skip: !process.env.PGLITE_MODULE }, async () => {
  const { PGlite } = require(process.env.PGLITE_MODULE), pg = new PGlite();
  const a = '00000000-0000-0000-0000-000000000001', b = '00000000-0000-0000-0000-000000000002';
  const outsider = '00000000-0000-0000-0000-000000000003';
  try {
    await pg.exec(`create role anon; create role authenticated; create schema auth;
      create table auth.users(id uuid primary key,email text);
      create table public.profiles(id uuid primary key,display_name text);
      create function auth.uid() returns uuid language sql stable as $$select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid$$;
      grant usage on schema auth, public to anon, authenticated;`);
    await pg.exec(setupSQL.replace(migrationBlock, ''));
    await pg.query('insert into auth.users values($1,$2),($3,$4),($5,$6)', [a, 'a@example.test', b, 'b@example.test', outsider, 'outside@example.test']);
    await pg.query('insert into public.profiles values($1,$2),($3,$4),($5,$6)', [a, '成员甲', b, '成员乙', outsider, '非成员']);
    await pg.query('insert into public.sales_members(email,user_id) values($1,$2),($3,$4)', ['a@example.test', a, 'b@example.test', b]);
    const inserted = (await pg.query(`insert into public.sales_ai_logs(message,created_by) values
      ('甲第一条',$1),('乙第一条',$2),('甲第二条',$1),('非成员原记录',$3),('已注销账号原记录',null),('乙第二条',$2) returning id`, [a, b, outsider])).rows.map(r => r.id);
    const [aFirst, bFirst, aSecond, outsiderLog, orphanLog, bSecond] = inserted;
    const identity = async id => { await pg.query("select set_config('request.jwt.claim.sub',$1,false)", [id]); };
    const callDelete = id => pg.query('select public.delete_sales_ai_log($1) as id', [id]);

    // 模拟旧 Supabase 已给表权限，但尚没有 DELETE policy/RPC 的状态。
    await pg.exec('grant select, delete on public.sales_ai_logs to authenticated; set role authenticated');
    await identity(a);
    assert.equal((await pg.query('delete from public.sales_ai_logs where id=$1 returning id', [aFirst])).rows.length, 0);
    await assert.rejects(callDelete(aFirst), e => e.code === '42883');
    await pg.exec('reset role');
    await pg.exec(migrationSQL); await pg.exec(migrationSQL); await pg.exec(setupSQL);
    assert.equal((await pg.query('select count(*)::int as n from public.sales_ai_logs')).rows[0].n, inserted.length, '重复安装不删除旧数据');
    assert.equal((await pg.query("select prosecdef from pg_proc where oid='public.delete_sales_ai_log(bigint)'::regprocedure")).rows[0].prosecdef, false, 'RPC不能绕过RLS');
    assert.equal((await pg.query("select has_function_privilege('anon','public.delete_sales_ai_log(bigint)','execute') as allowed")).rows[0].allowed, false);

    await pg.exec('set role authenticated'); await identity(a);
    await assert.rejects(callDelete(bFirst), e => e.code === '42501' && e.message === 'ai_log_not_owner');
    await assert.rejects(callDelete(orphanLog), e => e.code === '42501' && e.message === 'ai_log_not_owner');
    assert.equal((await pg.query('delete from public.sales_ai_logs where id=$1 returning id', [bFirst])).rows.length, 0, '直接DELETE他人记录仍被RLS拦截');
    assert.equal((await callDelete(aFirst)).rows[0].id, aFirst);
    assert.equal((await pg.query('select id from public.sales_ai_logs where id=$1', [aFirst])).rows.length, 0);
    await assert.rejects(callDelete(aFirst), e => e.code === 'P0002' && e.message === 'ai_log_not_found');
    await assert.rejects(callDelete(null), e => e.code === '22023');
    await identity(outsider);
    await assert.rejects(callDelete(outsiderLog), e => e.code === '42501' && e.message === 'sales_members_only');
    assert.equal((await pg.query('delete from public.sales_ai_logs where id=$1 returning id', [outsiderLog])).rows.length, 0);

    // RPC存在但DELETE policy缺失，必须明确报错，不能返回成功。
    await pg.exec('reset role; drop policy "本人可删除提问" on public.sales_ai_logs; set role authenticated');
    await identity(b);
    await assert.rejects(callDelete(bFirst), e => e.code === '42501' && e.message === 'ai_log_delete_denied');
    assert.equal((await pg.query('select id from public.sales_ai_logs where id=$1', [bFirst])).rows.length, 1);
    await pg.exec('reset role'); await pg.exec(migrationSQL); await pg.exec('set role authenticated');
    assert.equal((await callDelete(bFirst)).rows[0].id, bFirst);

    // 删除白名单后，即使知道自己的历史编号也不能再删。
    await pg.exec('reset role');
    await pg.query('delete from public.sales_members where user_id=$1', [a]);
    await pg.exec('set role authenticated'); await identity(a);
    await assert.rejects(callDelete(aSecond), e => e.code === '42501' && e.message === 'sales_members_only');
    assert.equal((await pg.query('delete from public.sales_ai_logs where id=$1 returning id', [aSecond])).rows.length, 0);
    await identity('');
    await assert.rejects(callDelete(bSecond), e => e.code === '42501' && e.message === 'ai_log_auth_required');
    await pg.exec('reset role; set role anon');
    await assert.rejects(callDelete(bSecond), e => e.code === '42501');
    await pg.exec('reset role');
    assert.deepEqual((await pg.query('select id from public.sales_ai_logs order by id')).rows.map(r => r.id), [aSecond, outsiderLog, orphanLog, bSecond]);
  } finally { await pg.close(); }
});
