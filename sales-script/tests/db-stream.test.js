const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const stream = require('../stream.js');
const source = fs.readFileSync(require.resolve('../db.js'), 'utf8');
function setup({ session = { access_token: 'test-session-token' }, getSession, fetch: fetchImpl } = {}) {
  const calls = [], window = { location: { search: '' }, SalesStream: stream,
    supabase: { createClient: () => ({ auth: { getSession: getSession || (async () => ({ data: { session } })) } }) } };
  vm.runInNewContext(source, { window, fetch: async (...args) => {
    calls.push(args); return fetchImpl ? fetchImpl(...args) : Response.json({ content: '{}', provider: 'openai' });
  } });
  const db = window.SalesDB;
  db.init({ SUPABASE_URL: 'https://example.test/', SUPABASE_ANON_KEY: 'test-public-key', AI_FUNCTION: 'sales-ai', AI_REGION: 'ap-southeast-1' });
  return { db, calls };
}
test('流请求使用当前会话与公开key，传递区域、signal与stream开关', async () => {
  const { db, calls } = setup(), controller = new AbortController(), events = [];
  const result = await db.invokeAIStream({ json: true, messages: [] }, { signal: controller.signal, onEvent: e => events.push(e) });
  assert.equal(calls.length, 1);
  const [url, options] = calls[0];
  assert.equal(url, 'https://example.test/functions/v1/sales-ai');
  assert.equal(options.method, 'POST');
  assert.equal(options.headers.Authorization, 'Bearer test-session-token');
  assert.equal(options.headers.apikey, 'test-public-key');
  assert.equal(options.headers['x-region'], 'ap-southeast-1');
  assert.equal(options.headers.Accept, 'text/event-stream');
  assert.equal(options.signal, controller.signal);
  assert.equal(JSON.parse(options.body).stream, true);
  assert.equal(result.streaming, false);
  assert.equal(events[0].type, 'legacy');
});
test('没有会话或已取消时不发送模型请求', async () => {
  const a = setup({ session: null });
  await assert.rejects(a.db.invokeAIStream({}), e => e.code === 'auth');
  assert.equal(a.calls.length, 0);
  const b = setup(), controller = new AbortController(); controller.abort();
  await assert.rejects(b.db.invokeAIStream({}, { signal: controller.signal }), e => e.name === 'AbortError');
  assert.equal(b.calls.length, 0);
});
test('中转网络失败不伪造服务商状态', async () => {
  const { db } = setup({ fetch: async () => { throw new TypeError('Failed to fetch'); } });
  await assert.rejects(db.invokeAIStream({}), e => e.code === 'network' && !e.ai_status && /暂时无法判断/.test(e.message));
});


test('读取会话尚未完成时取消，立即结束且不再fetch', async () => {
  let resolveSession;
  const { db, calls } = setup({ getSession: () => new Promise(resolve => { resolveSession = resolve; }) });
  const controller = new AbortController();
  const pending = db.invokeAIStream({}, { signal: controller.signal });
  controller.abort();
  await assert.rejects(pending, e => e.name === 'AbortError');
  resolveSession({ data: { session: { access_token: 'later-token' } } });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(calls.length, 0);
});
