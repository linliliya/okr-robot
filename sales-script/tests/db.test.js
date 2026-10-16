// 运行：node sales-script/tests/db.test.js（只使用内存模拟，不连接 Supabase）
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../db.js'), 'utf8');
const config = { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_ANON_KEY: 'public-test-key' };

function loadDB({ search = '', sdk, logs = [] } = {}) {
  const window = { location: { search } };
  if (sdk !== undefined) window.supabase = sdk;
  let saved = JSON.stringify({ logs });
  const localStorage = {
    getItem: () => saved,
    setItem: (_key, value) => { saved = value; },
  };
  vm.runInNewContext(source, { window, localStorage });
  return window.SalesDB;
}

function mockSupabase(logs, failure) {
  const requests = [];
  const clientArgs = [];
  const sdk = {
    createClient(...args) {
      clientArgs.push(args);
      return {
        from(table) {
          const state = { table, filters: [], orders: [], limit: null, range: null };
          const q = {
            select() { return q; },
            order(field, { ascending }) { state.orders.push({ field, ascending }); return q; },
            gte(field, value) { state.filters.push({ op: 'gte', field, value }); return q; },
            lt(field, value) { state.filters.push({ op: 'lt', field, value }); return q; },
            eq(field, value) { state.filters.push({ op: 'eq', field, value }); return q; },
            limit(value) { state.limit = value; return q; },
            range(from, to) { state.range = [from, to]; return q; },
            then(resolve, reject) {
              return Promise.resolve().then(() => {
                if (table === 'profiles') return { data: [{ id: 'seller-a', display_name: '销售 A' }], error: null };
                assert.equal(table, 'sales_ai_logs');
                requests.push(state);
                if (failure && state.range && state.range[0] >= 1000) {
                  if (failure instanceof Error) throw failure;
                  return { data: null, error: failure };
                }
                let rows = logs.filter(row => state.filters.every(({ op, field, value }) =>
                  op === 'eq' ? row[field] === value : op === 'gte' ? row[field] >= value : row[field] < value));
                rows.sort((a, b) => {
                  for (const { field, ascending } of state.orders) {
                    if (a[field] !== b[field]) return (a[field] < b[field] ? -1 : 1) * (ascending ? 1 : -1);
                  }
                  return 0;
                });
                if (state.range) rows = rows.slice(state.range[0], state.range[1] + 1);
                if (state.limit !== null) rows = rows.slice(0, state.limit);
                return { data: rows, error: null };
              }).then(resolve, reject);
            },
          };
          return q;
        },
      };
    },
  };
  return { sdk, requests, clientArgs };
}

const since = '2026-10-04T16:00:00.000Z';
const until = '2026-10-11T16:00:00.000Z';
const matching = Array.from({ length: 2305 }, (_, i) => ({
  id: i + 1, created_at: since, created_by: 'seller-a', feedback: 'bad', message: `问题 ${i + 1}`,
}));
const logs = matching.concat([
  { id: 3000, created_at: '2026-10-04T15:59:59.999Z', created_by: 'seller-a', feedback: 'bad' },
  { id: 3001, created_at: until, created_by: 'seller-a', feedback: 'bad' },
  { id: 3002, created_at: since, created_by: 'seller-b', feedback: 'bad' },
  { id: 3003, created_at: since, created_by: 'seller-a', feedback: 'good' },
]);
const filters = { since, until, feedback: 'bad', createdBy: 'seller-a' };
const ids = rows => Array.from(rows, row => row.id);

test('正式模式缺少配置时明确报错，不进入演示模式', () => {
  for (const cfg of [undefined, {}, { SUPABASE_URL: config.SUPABASE_URL }, { SUPABASE_ANON_KEY: config.SUPABASE_ANON_KEY }]) {
    const db = loadDB();
    assert.throws(() => db.init(cfg), e => e.code === 'config' && /config\.js/.test(e.message));
    assert.equal(db.mode, null);
    assert.equal(db.saveCase, undefined);
  }
});

test('正式模式 SDK 未加载或损坏时明确报错', () => {
  for (const sdk of [undefined, {}]) {
    const db = loadDB({ sdk });
    assert.throws(() => db.init(config), e => e.code === 'network' && /数据库组件加载失败/.test(e.message));
    assert.equal(db.mode, null);
  }
});

test('只有完整 demo=1 参数启用演示，演示无需配置或 SDK', async () => {
  for (const search of ['?demo=1', '?other=ok&demo=1', '?demo=1&other=ok']) {
    const db = loadDB({ search });
    assert.equal(db.init().mode, 'demo');
    assert.equal((await db.getSession()).user.id, 'demo-user');
  }
  for (const search of ['', '?demo=0', '?demo=10', '?demo=1extra', '?demo=1.2', '?otherdemo=1']) {
    const db = loadDB({ search });
    assert.throws(() => db.init(config), e => e.code === 'network');
  }
});

test('完整配置和 SDK 正常初始化真实适配器', () => {
  const mock = mockSupabase([]);
  const db = loadDB({ sdk: mock.sdk });
  assert.equal(db.init(config).mode, 'supabase');
  assert.deepEqual(mock.clientArgs, [[config.SUPABASE_URL, config.SUPABASE_ANON_KEY]]);
});

test('all:true 跨三页读取全部反馈，每页保留筛选并稳定排序', async () => {
  const mock = mockSupabase(logs);
  const db = loadDB({ sdk: mock.sdk });
  db.init(config);
  const rows = await db.listAiLogs({ ...filters, all: true });
  assert.equal(rows.length, 2305);
  assert.deepEqual(ids(rows), matching.map(row => row.id).reverse());
  assert.equal(rows[0].created_by_name, '销售 A');
  assert.deepEqual(mock.requests.map(r => r.range), [[0, 999], [1000, 1999], [2000, 2999]]);
  for (const request of mock.requests) {
    assert.equal(request.limit, null);
    assert.deepEqual(request.filters, [
      { op: 'gte', field: 'created_at', value: since },
      { op: 'lt', field: 'created_at', value: until },
      { op: 'eq', field: 'feedback', value: 'bad' },
      { op: 'eq', field: 'created_by', value: 'seller-a' },
    ]);
    assert.deepEqual(request.orders, [{ field: 'created_at', ascending: false }, { field: 'id', ascending: false }]);
  }
});

test('普通列表仍默认 200 条，显式 limit:15 保持不变', async () => {
  const mock = mockSupabase(logs);
  const db = loadDB({ sdk: mock.sdk });
  db.init(config);
  assert.equal((await db.listAiLogs(filters)).length, 200);
  assert.equal((await db.listAiLogs({ ...filters, limit: 15 })).length, 15);
  assert.deepEqual(mock.requests.map(r => r.limit), [200, 15]);
  assert(mock.requests.every(r => r.range === null));
});

test('后续分页失败时向上抛错，不能返回部分数据或空列表', async () => {
  for (const [failure, code] of [[new Error('Failed to fetch'), 'network'], [{ code: '42P01', message: 'table does not exist' }, 'tables_missing']]) {
    const mock = mockSupabase(logs, failure);
    const db = loadDB({ sdk: mock.sdk });
    db.init(config);
    await assert.rejects(db.listAiLogs({ ...filters, all: true }), e => e.code === code && e.cause === failure);
    assert.equal(mock.requests.length, 2);
  }
});

test('演示适配器与真实适配器的筛选、排序、all 和 limit 行为一致', async () => {
  const mock = mockSupabase(logs);
  const real = loadDB({ sdk: mock.sdk });
  real.init(config);
  const demo = loadDB({ search: '?demo=1', logs });
  demo.init();
  for (const options of [filters, { ...filters, limit: 15 }, { ...filters, all: true }]) {
    assert.deepEqual(ids(await demo.listAiLogs(options)), ids(await real.listAiLogs(options)));
  }
});

test('AI 错误保留后端服务状态，客户端连接错误不冒充某个模型故障', async () => {
  const states = { openai: { state: 'unavailable', reason: 'quota' }, deepseek: { state: 'not_configured' } };
  const sdk = { createClient: () => ({ functions: { invoke: async () => ({ error: {
    message: 'Edge Function returned a non-2xx status code',
    context: { status: 502, json: async () => ({ error: '主服务额度不足，备用服务未配置', ai_status: states }) },
  } }) } }) };
  const db = loadDB({ sdk });
  db.init(config);
  await assert.rejects(db.invokeAI({}), e => e.code === 'ai' && e.ai_status === states && /额度不足/.test(e.message));

  const broken = loadDB({ sdk: { createClient: () => ({ functions: { invoke: async () => { throw new Error('Failed to fetch'); } } }) } });
  broken.init(config);
  await assert.rejects(broken.invokeAI({}), e => e.code === 'network' && !e.ai_status && /暂时无法判断/.test(e.message));
});
