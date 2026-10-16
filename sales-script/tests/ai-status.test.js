const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { create } = require('../ai-status.js');

const status = (state, message = '', reason = null) => ({ state, message, reason, checked_at: '2026-10-16T08:00:00.000Z' });
const available = () => ({ openai: status('available'), deepseek: status('unknown') });
const result = (provider = 'openai') => ({
  content: '{}', provider, model: provider === 'openai' ? 'gpt-6-luna' : 'deepseek-flash',
  fallback: provider === 'deepseek',
  ai_status: provider === 'openai' ? available() : {
    openai: status('unavailable', '连接超时', 'timeout'), deepseek: status('available'),
  },
});
function deferred() {
  let resolve, reject;
  const promise = new Promise((a, b) => { resolve = a; reject = b; });
  return { promise, resolve, reject };
}

test('实际服务商随响应更新，DeepSeek 兜底后新请求可恢复 OpenAI', async () => {
  const replies = [result('deepseek'), result()];
  const store = create(() => ({ mode: 'supabase', invokeAI: async () => replies.shift() }));
  await store.invoke({ messages: [] });
  assert.equal(store.getState().lastCall.provider, 'deepseek');
  assert.equal(store.getState().lastCall.fallback, true);
  assert.equal(store.getState().providers.openai.reason, 'timeout');
  await store.invoke({ messages: [] });
  assert.equal(store.getState().lastCall.provider, 'openai');
  assert.equal(store.getState().lastCall.fallback, false);
  assert.equal(store.getState().providers.openai.state, 'available');
  assert.equal(store.getState().providers.deepseek.state, 'available', '未尝试的服务不抹掉已有记录');
});

test('中转服务网络失败不伪造 OpenAI 不可用，也不声称使用了备用服务', async () => {
  const store = create(() => ({ mode: 'supabase', invokeAI: async () => { throw new Error('中转服务连接失败'); } }));
  await assert.rejects(store.invoke({ messages: [] }));
  assert.equal(store.getState().providers.openai.state, 'unknown');
  assert.equal(store.getState().providers.deepseek.state, 'unknown');
  assert.equal(store.getState().lastCall.provider, null);
  assert.equal(store.getState().lastCall.ok, false);
  assert.equal(store.getState().pending, 0);
});

test('两家均失败时保留具体状态', async () => {
  const error = Object.assign(new Error('两家服务暂不可用'), { ai_status: {
    openai: status('unavailable', '额度不足', 'quota'), deepseek: status('not_configured'),
  } });
  const store = create(() => ({ mode: 'supabase', invokeAI: async () => { throw error; } }));
  await assert.rejects(store.invoke({}), error);
  assert.equal(store.getState().providers.openai.reason, 'quota');
  assert.equal(store.getState().providers.deepseek.state, 'not_configured');
});

test('连接探测合并重复点击，不会把可连接显示成已使用', async () => {
  const call = deferred();
  const requests = [];
  const store = create(() => ({ mode: 'supabase', invokeAI: body => { requests.push(body); return call.promise; } }));
  const a = store.check(), b = store.check();
  assert.equal(a, b);
  assert.equal(store.getState().checking, true);
  assert.deepEqual(requests, [{ action: 'status' }]);
  call.resolve({ ai_status: available() });
  await a;
  assert.equal(store.getState().lastCall, null);
  assert.equal(store.getState().providers.openai.source, 'probe');
  assert.equal(store.getState().checking, false);
});

test('较早发起的检查结果或检查失败不覆盖更新的实际调用', async () => {
  for (const fail of [false, true]) {
    const check = deferred();
    const store = create(() => ({ mode: 'supabase', invokeAI: body => body.action ? check.promise : Promise.resolve(result()) }));
    const a = store.check().catch(() => {});
    await store.invoke({ messages: [] });
    if (fail) check.reject(new Error('旧检查失败'));
    else check.resolve({ ai_status: { openai: status('unavailable', '旧检查超时', 'timeout') } });
    await a;
    assert.equal(store.getState().providers.openai.state, 'available');
    assert.equal(store.getState().providers.openai.source, 'request');
    assert.equal(store.getState().issue, '');
  }
});

test('分段并发请求分别计数，旧版函数不猜测服务商', async () => {
  const a = deferred(), b = deferred();
  const queue = [a, b];
  const store = create(() => ({ mode: 'supabase', invokeAI: () => queue.shift().promise }));
  const p1 = store.invoke({}), p2 = store.invoke({});
  assert.equal(store.getState().pending, 2);
  a.resolve({ content: '{}' });
  await p1;
  assert.equal(store.getState().pending, 1);
  assert.equal(store.getState().lastCall.provider, null);
  assert.match(store.getState().issue, /尚未升级/);
  b.resolve(result());
  await p2;
  assert.equal(store.getState().pending, 0);
  assert.equal(store.getState().issue, '');
});

test('演示模式不发起连接探测', async () => {
  const store = create(() => ({ mode: 'demo', invokeAI() { throw new Error('不应该调用'); } }));
  assert.equal(await store.check(), null);
  assert.equal(store.getState().demo, true);
});

test('分段请求反序完成时，较早的成功或失败不能覆盖较新的调用摘要', async () => {
  for (const oldFails of [true, false]) {
    const older = deferred(), newer = deferred();
    const queue = [older, newer];
    const store = create(() => ({ mode: 'supabase', invokeAI: () => queue.shift().promise }));
    const a = store.invoke({}).catch(() => {}), b = store.invoke({}).catch(() => {});
    if (oldFails) newer.resolve(result());
    else newer.reject(Object.assign(new Error('最新请求失败'), { ai_status: { openai: status('unavailable', '额度不足', 'quota') } }));
    await b;
    if (oldFails) older.reject(new Error('旧请求失败'));
    else older.resolve(result());
    await a;
    assert.equal(store.getState().lastCall.ok, oldFails);
    assert.equal(store.getState().issue, oldFails ? '' : '最新请求失败');
    assert.equal(store.getState().providers.openai.state, oldFails ? 'available' : 'unavailable');
    assert.equal(store.getState().pending, 0);
  }
});

test('页面区分探测与实际使用，并转义服务返回的文案', () => {
  const box = {}, button = {};
  const window = {};
  const document = { querySelector: id => id === '#ai-status-content' ? box : button, addEventListener() {} };
  const source = fs.readFileSync(require.resolve('../app.js'), 'utf8')
    .replace('  boot().catch(renderError);', '  window.renderAIStatus = renderAIStatus;');
  vm.runInNewContext(source, { window, document });
  const store = create(() => ({ mode: 'supabase' }));
  const state = store.getState();
  state.providers.openai = { ...status('available'), source: 'probe' };
  window.renderAIStatus(state);
  assert.match(box.innerHTML, /连接检查通过/);
  assert.doesNotMatch(box.innerHTML, /最近使用/);
  state.lastCall = { ok: true, provider: 'deepseek', model: 'deepseek-flash', fallback: true, reason: '<script>bad</script>' };
  state.issue = '<img onerror=bad>';
  window.renderAIStatus(state);
  assert.match(box.innerHTML, /最近使用 DeepSeek/);
  assert.match(box.innerHTML, /下次仍优先尝试 ChatGPT/);
  assert.doesNotMatch(box.innerHTML, /<script>|<img/);
  assert.match(box.innerHTML, /&lt;script&gt;/);
  state.lastCall = { ok: false };
  state.providers.openai = { ...status('available', '无法处理此次请求', 'request'), source: 'request' };
  window.renderAIStatus(state);
  assert.match(box.innerHTML, /连接正常，请求未完成/);
  assert.match(box.innerHTML, /无法处理此次请求/);
  assert.doesNotMatch(box.innerHTML, /最近调用成功/);
});

test('流式状态在完成前更新，取消不把服务记成失败', async () => {
  const d = deferred(), seen = [];
  const store = create(() => ({ mode: 'supabase', invokeAIStream: async (_body, opts) => {
    opts.onEvent({ type: 'status', provider: 'openai', ai_status: available() });
    return d.promise;
  } }));
  const pending = store.invokeStream({}, { onEvent: e => seen.push(e) });
  assert.equal(store.getState().pending, 1);
  assert.equal(store.getState().providers.openai.state, 'available');
  assert.equal(store.getState().lastCall, null);
  d.reject(Object.assign(new Error('已停止'), { name: 'AbortError', code: 'cancelled' }));
  await assert.rejects(pending, /已停止/);
  assert.equal(store.getState().pending, 0);
  assert.equal(store.getState().lastCall, null);
  assert.equal(store.getState().issue, '');
  assert.equal(seen[0].type, 'status');
});

test('旧流事件不能覆盖较新的调用状态', async () => {
  const old = deferred(); let notify;
  const store = create(() => ({ mode: 'supabase',
    invokeAI: async () => result(),
    invokeAIStream: (_body, opts) => { notify = opts.onEvent; return old.promise; },
  }));
  const a = store.invokeStream({});
  await store.invoke({});
  notify({ type: 'reset', ai_status: { openai: status('unavailable', '旧流失败', 'timeout') } });
  assert.equal(store.getState().providers.openai.state, 'available');
  old.resolve(result('deepseek')); await a;
  assert.equal(store.getState().lastCall.provider, 'openai');
  assert.equal(store.getState().pending, 0);
});


test('旧适配器没有流方法时标明完整JSON结果，只调用一次', async () => {
  let calls = 0; const events = [];
  const store = create(() => ({ mode: 'supabase', invokeAI: async () => { calls++; return result(); } }));
  const response = await store.invokeStream({}, { onEvent: e => events.push(e) });
  assert.equal(response.streaming, false);
  assert.equal(calls, 1);
  assert.equal(events[0].type, 'legacy');
  assert.equal(store.getState().lastCall.provider, 'openai');
});
