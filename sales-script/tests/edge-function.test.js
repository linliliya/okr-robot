// 运行：node sales-script/tests/edge-function.test.js
// 使用 Node 的 TypeScript 转换与 VM，全部鉴权和上游请求均在内存中模拟。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const { stripTypeScriptTypes } = require('node:module');

const source = fs.readFileSync(path.join(__dirname, '../edge-function/sales-ai.ts'), 'utf8');
const compiled = stripTypeScriptTypes(source.replace(/^import \{ createClient \} from .*;\s*$/m, ''), { mode: 'transform' });
const KEYS = { OPENAI_API_KEY: 'openai-test-secret', DEEPSEEK_API_KEY: 'deepseek-test-secret' };
const messages = [{ role: 'user', content: '请输出 JSON' }];
const jsonResponse = (data, status = 200) => new Response(JSON.stringify(data), { status, headers: { 'Content-Type': 'application/json' } });
const completion = (content = '{"items":[]}', extra = {}) => jsonResponse({ choices: [{ message: { content }, finish_reason: 'stop' }], usage: { total_tokens: 5 }, ...extra });
const upstreamError = (status, code = '', message = '') => jsonResponse({ error: { code, message } }, status);

function fixture({ env = KEYS, fetch: responder = () => completion(), member = true, userError = false, memberError = null, shortTimeouts = false } = {}) {
  let handler;
  const calls = [], authCalls = [], budgets = [], timers = new Set();
  const variables = { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_ANON_KEY: 'test-public-key', ...env };
  const context = {
    URL, Request, Response, AbortController, ReadableStream, TextEncoder, TextDecoder,
    Deno: { env: { get: key => variables[key] }, serve: fn => { handler = fn; } },
    createClient: (_url, _key, options) => ({
      auth: { getUser: async token => {
        authCalls.push({ method: 'getUser', token, authorization: options.global.headers.Authorization });
        return userError ? { data: null, error: { message: 'private-auth-detail' } } : { data: { user: { id: 'member-id' } }, error: null };
      } },
      rpc: async name => { authCalls.push({ method: name }); return { data: member, error: memberError }; },
    }),
    fetch: async (url, options) => {
      const call = { url, method: options.method, headers: options.headers, body: options.body ? JSON.parse(options.body) : null, signal: options.signal };
      calls.push(call);
      return responder(call, calls.length);
    },
    setTimeout: (fn, ms) => {
      budgets.push(ms);
      const id = setTimeout(() => { timers.delete(id); fn(); }, shortTimeouts ? 5 : ms);
      timers.add(id); return id;
    },
    clearTimeout: id => { timers.delete(id); clearTimeout(id); },
  };
  vm.runInNewContext(compiled, context, { filename: 'sales-ai-under-test.js' });
  const rawRequest = async (body = { messages, json: true }, options = {}) => {
    const headers = { 'Content-Type': 'application/json', apikey: 'browser-public-key' };
    if (!options.noAuth) headers.Authorization = 'Bearer logged-in-token';
    return handler(new Request('https://example.supabase.co/functions/v1/sales-ai', {
      method: options.method || 'POST', headers, signal: options.signal,
      ...((options.method === 'GET' || options.method === 'OPTIONS') ? {} : { body: typeof body === 'string' ? body : JSON.stringify(body) }),
    }));
  };
  return {
    calls, authCalls, budgets, rawRequest, activeTimers: () => timers.size,
    async request(body = { messages, json: true }, options = {}) {
      const response = await rawRequest(body, options);
      const text = await response.text();
      return { status: response.status, headers: response.headers, body: text === 'ok' ? text : JSON.parse(text) };
    },
  };
}

function assertPrivate(result) {
  const text = JSON.stringify(result.body);
  for (const secret of [...Object.values(KEYS), 'private-auth-detail', 'private-upstream-detail']) assert(!text.includes(secret));
}
function assertChecked(status) {
  assert.equal(typeof status.checked_at, 'string');
  assert(!Number.isNaN(Date.parse(status.checked_at)));
}

test('优先 OpenAI，成功返回实际提供方且不请求 DeepSeek', async () => {
  const f = fixture();
  const result = await f.request();
  assert.equal(result.status, 200);
  assert.equal(result.body.provider, 'openai');
  assert.equal(result.body.fallback, false);
  assert.equal(result.body.model, 'gpt-6-luna');
  assert.equal(result.body.ai_status.openai.state, 'available');
  assert.equal(result.body.ai_status.deepseek.state, 'unknown');
  assert.equal(result.body.ai_status.deepseek.checked_at, null);
  assert.equal(f.calls.length, 1);
  assert.equal(f.calls[0].url, 'https://api.openai.com/v1/chat/completions');
  assert.equal(f.calls[0].headers.Authorization, 'Bearer ' + KEYS.OPENAI_API_KEY);
  assert.equal(f.calls[0].body.response_format.type, 'json_object');
  assert.equal(f.calls[0].body.reasoning_effort, 'none');
  assert.deepEqual(f.budgets, [45000]);
  assertChecked(result.body.ai_status.openai);
  assertPrivate(result);
});

test('OpenAI 网络异常后使用 DeepSeek，下次请求重新尝试 OpenAI', async () => {
  let openaiAttempts = 0;
  const f = fixture({ fetch: call => {
    if (call.url.includes('openai') && ++openaiAttempts === 1) throw new Error('private-upstream-detail');
    return completion();
  } });
  const fallback = await f.request();
  assert.equal(fallback.body.provider, 'deepseek');
  assert.equal(fallback.body.fallback, true);
  assert.equal(fallback.body.ai_status.openai.reason, 'network');
  assert.equal(fallback.body.ai_status.deepseek.state, 'available');
  assert.equal(f.calls[1].headers.Authorization, 'Bearer ' + KEYS.DEEPSEEK_API_KEY);
  assert.equal(f.calls[1].body.model, 'deepseek-flash');
  assert.equal(f.calls[1].body.reasoning_effort, undefined);
  assert.deepEqual(f.budgets, [45000, 90000]);
  const recovered = await f.request();
  assert.equal(recovered.body.provider, 'openai');
  assert.equal(recovered.body.fallback, false);
  assert.equal(f.calls.length, 3);
  assertPrivate(fallback);
});

for (const phase of ['headers', 'body']) {
  test(`OpenAI ${phase === 'body' ? '读取响应体' : '等待响应头'}超时后兜底，并取消主请求`, async () => {
    const never = () => new Promise(() => {});
    const f = fixture({ shortTimeouts: true, fetch: call => {
      if (!call.url.includes('openai')) return completion();
      return phase === 'headers' ? never() : { ok: true, status: 200, text: never };
    } });
    const result = await f.request();
    assert.equal(result.status, 200);
    assert.equal(result.body.provider, 'deepseek');
    assert.equal(result.body.ai_status.openai.reason, 'timeout');
    assert.equal(f.calls[0].signal.aborted, true);
    assert.deepEqual(f.budgets, [45000, 90000]);
  });
}

for (const [status, code, reason] of [
  [401, 'invalid_api_key', 'auth'], [403, 'forbidden', 'auth'],
  [403, 'unsupported_country_region_territory', 'region'], [402, '', 'quota'],
  [429, 'insufficient_quota', 'quota'], [429, 'rate_limit_exceeded', 'rate_limit'],
  [500, '', 'server'], [503, '', 'server'],
]) {
  test(`上游 ${status} ${code || '无错误码'} 分类为 ${reason} 并兜底`, async () => {
    const f = fixture({ fetch: (_call, n) => n === 1 ? upstreamError(status, code, 'private-upstream-detail') : completion() });
    const result = await f.request();
    assert.equal(result.status, 200);
    assert.equal(result.body.provider, 'deepseek');
    assert.equal(result.body.ai_status.openai.reason, reason);
    assertPrivate(result);
  });
}

for (const [name, response] of [
  ['空内容', () => completion('')],
  ['坏上游 JSON', () => new Response('<html>private-upstream-detail</html>')],
  ['JSON 模式内容不合法', () => completion('{broken')],
]) {
  test(`${name}可以兜底`, async () => {
    const f = fixture({ fetch: (_call, n) => n === 1 ? response() : completion() });
    const result = await f.request();
    assert.equal(result.body.provider, 'deepseek');
    assert.equal(result.body.ai_status.openai.reason, 'invalid_response');
    assertPrivate(result);
  });
}

test('普通文本模式不要求内容为 JSON', async () => {
  const f = fixture({ fetch: () => completion('普通文本回复') });
  const result = await f.request({ messages });
  assert.equal(result.status, 200);
  assert.equal(result.body.content, '普通文本回复');
  assert.equal(f.calls.length, 1);
  assert.equal(f.calls[0].body.response_format, undefined);
});

for (const [name, response] of [
  ['非法请求 400', () => upstreamError(400, 'invalid_request_error', 'private-upstream-detail')],
  ['显式拒绝', () => completion('', { choices: [{ message: { refusal: '内容无法处理' } }] })],
  ['内容过滤结束', () => completion('', { choices: [{ message: { content: null }, finish_reason: 'content_filter' }] })],
  ['HTTP 内容策略拒绝', () => upstreamError(403, 'content_policy_violation', 'private-upstream-detail')],
]) {
  test(`${name}不会切换到另一家`, async () => {
    const f = fixture({ fetch: response });
    const result = await f.request();
    assert.equal(result.status, 400);
    assert.equal(result.body.provider, null);
    assert.equal(result.body.ai_status.openai.reason, 'request');
    assert.equal(f.calls.length, 1);
    assertPrivate(result);
  });
}

test('两家都失败时返回双方原因，不能误报已使用 DeepSeek', async () => {
  const f = fixture({ fetch: (_call, n) => n === 1 ? upstreamError(429, 'insufficient_quota') : upstreamError(503, '', 'private-upstream-detail') });
  const result = await f.request();
  assert.equal(result.status, 502);
  assert.equal(result.body.provider, null);
  assert.equal(result.body.ai_status.openai.reason, 'quota');
  assert.equal(result.body.ai_status.deepseek.reason, 'server');
  assert.equal(result.body.ai_status.deepseek.state, 'unavailable');
  assertPrivate(result);
});

test('status 并行 GET 两家的 models，不触发生成，完成后记录时间', async () => {
  const releases = [];
  const f = fixture({ fetch: call => new Promise(resolve => {
    releases.push(() => resolve(call.url.includes('openai') ? jsonResponse({ data: [] }) : upstreamError(401, 'invalid_api_key')));
  }) });
  const started = Date.now();
  const promise = f.request({ action: 'status' });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(f.calls.length, 2, '探测应并行发起');
  for (const call of f.calls) { assert.equal(call.method, 'GET'); assert(call.url.endsWith('/models')); assert.equal(call.body, null); }
  assert.deepEqual(f.budgets, [10000, 10000]);
  releases.forEach(release => release());
  const result = await promise;
  assert.equal(result.status, 200);
  assert.equal(result.body.ai_status.openai.state, 'available');
  assert.equal(result.body.ai_status.deepseek.reason, 'auth');
  for (const status of Object.values(result.body.ai_status)) { assertChecked(status); assert(Date.parse(status.checked_at) >= started); }
});

test('status 响应体超时给出 timeout，未配置提供方不发请求', async () => {
  const f = fixture({ env: { OPENAI_API_KEY: KEYS.OPENAI_API_KEY }, shortTimeouts: true, fetch: () => ({ ok: true, status: 200, text: () => new Promise(() => {}) }) });
  const result = await f.request({ action: 'status' });
  assert.equal(result.body.ai_status.openai.reason, 'timeout');
  assert.equal(result.body.ai_status.deepseek.state, 'not_configured');
  assert.equal(result.body.ai_status.deepseek.checked_at, null);
  assert.equal(f.calls.length, 1);
});

for (const [name, settings, requestOptions, expected] of [
  ['未登录', {}, { noAuth: true }, 401],
  ['登录无效', { userError: true }, {}, 401],
  ['未开通成员', { member: false }, {}, 403],
  ['成员查询失败', { memberError: { message: 'private-auth-detail' } }, {}, 500],
]) {
  test(`${name}时生成与状态探测都不发上游请求`, async () => {
    const f = fixture(settings);
    for (const body of [{ messages }, { action: 'status' }]) {
      const result = await f.request(body, requestOptions);
      assert.equal(result.status, expected);
      assert.equal(result.body.provider, null);
      assertPrivate(result);
    }
    assert.equal(f.calls.length, 0);
  });
}

test('未配置任何 key 时返回可识别状态且不发上游请求', async () => {
  const f = fixture({ env: {} });
  const result = await f.request();
  assert.equal(result.status, 503);
  assert.equal(result.body.provider, null);
  for (const status of Object.values(result.body.ai_status)) {
    assert.equal(status.state, 'not_configured'); assert.equal(status.reason, 'not_configured'); assert.equal(status.checked_at, null);
  }
  const health = await f.request({ action: 'status' });
  assert.equal(health.status, 200);
  assert.equal(f.calls.length, 0);
});

for (const [base, expected] of [
  [undefined, 'deepseek'], ['', 'deepseek'], ['https://api.deepseek.com', 'deepseek'],
  ['https://api.deepseek.com/v1/', 'deepseek'], ['https://api.openai.com/v1', 'openai'],
]) {
  test(`旧 key 按官方地址 ${base || '默认'} 只交给 ${expected}`, async () => {
    const f = fixture({ env: { LLM_API_KEY: 'legacy-secret', LLM_BASE_URL: base, LLM_MODEL: 'legacy-model' } });
    const result = await f.request();
    assert.equal(result.body.provider, expected);
    assert.equal(result.body.model, 'legacy-model');
    assert.equal(result.body.fallback, expected === 'deepseek');
    assert.equal(f.calls.length, 1);
    assert(f.calls[0].url.startsWith(expected === 'openai' ? 'https://api.openai.com/' : 'https://api.deepseek.com/'));
    assert.equal(f.calls[0].headers.Authorization, 'Bearer legacy-secret');
  });
}

for (const base of ['https://example.com/v1', 'https://api.openai.com.evil.example/v1', 'https://api.openai.com@evil.example/v1', 'http://api.openai.com/v1', 'not-a-url']) {
  test(`旧自定义地址不迁移密钥 ${base}`, async () => {
    const f = fixture({ env: { LLM_API_KEY: 'legacy-secret', LLM_BASE_URL: base } });
    const result = await f.request();
    assert.equal(result.status, 503);
    assert.equal(f.calls.length, 0);
    assert(!JSON.stringify(result.body).includes('legacy-secret'));
  });
}

test('默认截图请求使用 GPT-6 Luna，故障时交给支持图片的 DeepSeek Flash', async () => {
  const screenshotMessages = [{ role: 'user', content: [
    { type: 'text', text: '请将截图中的对话转成 JSON' },
    { type: 'image_url', image_url: { url: 'data:image/png;base64,dGVzdA==' } },
  ] }];
  const f = fixture({ fetch: (_call, n) => n === 1 ? upstreamError(503) : completion() });
  const result = await f.request({ messages: screenshotMessages, vision: true, json: true });
  assert.equal(result.status, 200);
  assert.equal(f.calls[0].body.model, 'gpt-6-luna');
  assert.equal(f.calls[0].body.reasoning_effort, 'none');
  assert.equal(f.calls[1].body.model, 'deepseek-flash');
  assert.equal(f.calls[1].body.reasoning_effort, undefined);
  for (const call of f.calls) {
    assert.deepEqual(call.body.messages, screenshotMessages);
    assert.equal(call.body.response_format.type, 'json_object');
  }
  assert.equal(result.body.provider, 'deepseek');
});

test('显式文字模型继续覆盖默认值，不向其他型号发送 Luna 的推理参数', async () => {
  const f = fixture({ env: { ...KEYS, OPENAI_MODEL: 'custom-primary', DEEPSEEK_MODEL: 'custom-backup' },
    fetch: (_call, n) => n === 1 ? upstreamError(503) : completion() });
  const result = await f.request();
  assert.equal(f.calls[0].body.model, 'custom-primary');
  assert.equal(f.calls[1].body.model, 'custom-backup');
  for (const call of f.calls) assert.equal(call.body.reasoning_effort, undefined);
  assert.equal(result.body.model, 'custom-backup');
});

test('专用配置优先于旧配置，视觉请求使用对应提供方的视觉模型', async () => {
  const f = fixture({ env: { ...KEYS, LLM_API_KEY: 'legacy-secret', LLM_BASE_URL: 'https://api.openai.com/v1', LLM_MODEL: 'old-model',
    OPENAI_VISION_MODEL: 'primary-vision', DEEPSEEK_VISION_MODEL: 'backup-vision' },
  fetch: (_call, n) => n === 1 ? upstreamError(500) : completion() });
  const result = await f.request({ messages, vision: true, json: true });
  assert.equal(f.calls[0].body.model, 'primary-vision');
  assert.equal(f.calls[1].body.model, 'backup-vision');
  for (const call of f.calls) assert.equal(call.body.reasoning_effort, undefined);
  assert.equal(f.calls[0].headers.Authorization, 'Bearer ' + KEYS.OPENAI_API_KEY);
  assert.equal(f.calls[1].headers.Authorization, 'Bearer ' + KEYS.DEEPSEEK_API_KEY);
  assert.equal(result.body.model, 'backup-vision');
});

test('坏请求、缺消息和预检都不调用上游', async () => {
  const f = fixture();
  for (const body of ['{bad', null, [], {}, { messages: [] }]) assert.equal((await f.request(body)).status, 400);
  assert.equal((await f.request(undefined, { method: 'GET' })).status, 405);
  const preflight = await f.request(undefined, { method: 'OPTIONS', noAuth: true });
  assert.equal(preflight.status, 200);
  assert.equal(preflight.headers.get('Access-Control-Allow-Origin'), '*');
  assert.equal(f.calls.length, 0);
});

const encoder = new TextEncoder();
const streamBody = { messages, json: true, stream: true };
const deltaFrame = (text, extra = {}) => `data: ${JSON.stringify({ choices: [{ index: 0, delta: { content: text }, finish_reason: null, ...extra }] })}\n\n`;
const stopFrame = 'data: {"choices":[{"index":0,"delta":{},"finish_reason":"stop"}]}\n\n';
const doneFrame = 'data: [DONE]\n\n';
function sseResponse(parts) {
  return new Response(new ReadableStream({ start(controller) {
    for (const part of parts) controller.enqueue(typeof part === 'string' ? encoder.encode(part) : part);
    controller.close();
  } }), { headers: { 'Content-Type': 'text/event-stream' } });
}
const streamedCompletion = (content = '{"items":[]}', usage = { total_tokens: 9 }) => sseResponse([
  deltaFrame(content), stopFrame, `data: ${JSON.stringify({ choices: [], usage })}\n\n`, doneFrame,
]);
function eventReader(response) {
  const reader = response.body.getReader(), decoder = new TextDecoder();
  let buffer = '';
  return {
    reader,
    async next() {
      for (;;) {
        const end = buffer.indexOf('\n\n');
        if (end >= 0) {
          const frame = buffer.slice(0, end); buffer = buffer.slice(end + 2);
          const lines = frame.split('\n');
          return { event: lines.find(l => l.startsWith('event: ')).slice(7),
            data: JSON.parse(lines.filter(l => l.startsWith('data: ')).map(l => l.slice(6)).join('\n')) };
        }
        const { done, value } = await reader.read();
        if (done) { assert.equal(buffer, ''); return null; }
        buffer += decoder.decode(value, { stream: true });
      }
    },
    async all() { const events = []; for (;;) { const event = await this.next(); if (!event) return events; events.push(event); } },
  };
}
const streamEvents = async f => eventReader(await f.rawRequest(streamBody)).all();
const nextTurn = () => new Promise(resolve => setImmediate(resolve));
async function promptly(promise) {
  let timer;
  try { return await Promise.race([promise, new Promise((_resolve, reject) => { timer = setTimeout(() => reject(new Error('未增量返回')), 1000); })]); }
  finally { clearTimeout(timer); }
}

test('首个 delta 在上游结束前立即可读，最终 done 才确认 JSON 与实际提供方', async () => {
  let upstream, cancelled = false;
  const f = fixture({ fetch: () => new Response(new ReadableStream({
    start(controller) { upstream = controller; }, cancel() { cancelled = true; },
  }), { headers: { 'Content-Type': 'text/event-stream' } }) });
  const response = await f.rawRequest(streamBody), events = eventReader(response);
  assert.match(response.headers.get('Content-Type'), /text\/event-stream/);
  assert.match(response.headers.get('Cache-Control'), /no-transform/);
  assert.equal((await promptly(events.next())).event, 'status');
  upstream.enqueue(encoder.encode(deltaFrame('{"items":')));
  assert.deepEqual(await promptly(events.next()), { event: 'delta', data: { text: '{"items":' } });
  assert.equal(cancelled, false, '首段输出时上游仍保持连接');
  upstream.enqueue(encoder.encode(deltaFrame('[]}') + stopFrame + 'data: {"choices":[],"usage":{"total_tokens":7}}\n\n' + doneFrame));
  const rest = await events.all(), done = rest.at(-1);
  assert.equal(done.event, 'done');
  assert.equal(done.data.content, '{"items":[]}');
  assert.equal(done.data.provider, 'openai');
  assert.equal(done.data.fallback, false);
  assert.equal(done.data.model, 'gpt-6-luna');
  assert.equal(done.data.usage.total_tokens, 7);
  assert.equal(done.data.ai_status.openai.state, 'available');
  assert.equal(f.calls[0].body.stream, true);
  assert.equal(f.calls[0].body.stream_options.include_usage, true);
  assert.equal(f.calls[0].body.reasoning_effort, 'none');
  assert.equal(f.calls.length, 1);
  assert.equal(cancelled, true);
  assert.equal(f.activeTimers(), 0);
});

test('SSE解析兼容逐字节UTF-8、拆开的CRLF、多行data、角色块、心跳与usage块', async () => {
  const content = '{"文本":"你好😀"}';
  const wire = ': keepalive\r\nid: 1\r\nevent: message\r\ndata: {"choices":\r\ndata: [{"index":0,"delta":{"role":"assistant"},"finish_reason":null}]}\r\n\r\n'
    + deltaFrame(content).replace(/\n/g, '\r\n')
    + 'data: {"choices":[{"index":0,"delta":{"reasoning_content":"不应显示"},"finish_reason":"stop"}],"usage":{"total_tokens":8}}\r\n\r\n'
    + doneFrame.replace(/\n/g, '\r\n');
  const f = fixture({ fetch: () => sseResponse(Array.from(encoder.encode(wire), b => new Uint8Array([b]))) });
  const events = await streamEvents(f), done = events.at(-1);
  assert.deepEqual(events.filter(e => e.event === 'delta').map(e => e.data.text), [content]);
  assert.equal(done.event, 'done');
  assert.equal(done.data.content, content);
  assert.equal(done.data.usage.total_tokens, 8);
  assert(!JSON.stringify(events).includes('不应显示'));
  assert.equal(f.activeTimers(), 0);
});

test('主服务输出部分后断流，先reset再发备用内容，最终结果不混入旧内容', async () => {
  const f = fixture({ fetch: (_call, n) => n === 1 ? sseResponse([deltaFrame('{"旧结果":')]) : streamedCompletion('{"新结果":true}') });
  const events = await streamEvents(f);
  assert.deepEqual(events.map(e => e.event), ['status', 'delta', 'reset', 'status', 'delta', 'done']);
  assert.equal(events[2].data.ai_status.openai.reason, 'invalid_response');
  const done = events.at(-1).data;
  assert.equal(done.content, '{"新结果":true}');
  assert.equal(done.provider, 'deepseek');
  assert.equal(done.fallback, true);
  assert.equal(f.calls[1].body.reasoning_effort, undefined);
  assert.deepEqual(f.budgets, [45000, 90000]);
  assert.equal(f.activeTimers(), 0);
});

test('上游reader网络错误与完整但非法JSON均触发备用并清除已输出草稿', async () => {
  for (const kind of ['reader', 'json']) {
    const f = fixture({ fetch: (_call, n) => {
      if (n > 1) return streamedCompletion();
      if (kind === 'json') return streamedCompletion('{不合法');
      let pulled = false;
      return new Response(new ReadableStream({ pull(controller) {
        if (!pulled) { pulled = true; controller.enqueue(encoder.encode(deltaFrame('{"draft":'))); }
        else controller.error(new Error('private-upstream-detail'));
      } }));
    } });
    const events = await streamEvents(f);
    assert(events.some(e => e.event === 'reset'));
    assert.equal(events.at(-1).data.provider, 'deepseek');
    assert.equal(events.at(-1).data.ai_status.openai.reason, kind === 'reader' ? 'network' : 'invalid_response');
    assertPrivate({ body: events });
  }
});

for (const [name, parts] of [
  ['空流', []],
  ['缺少DONE', [deltaFrame('{}'), stopFrame]],
  ['缺少正常结束状态', [deltaFrame('{}'), doneFrame]],
  ['仅DONE', [doneFrame]],
  ['length截断', [deltaFrame('{}', { finish_reason: 'length' }), doneFrame]],
  ['工具调用结束', [deltaFrame('{}', { finish_reason: 'tool_calls' }), doneFrame]],
  ['坏事件JSON', ['data: {broken}\n\n']],
  ['错误choices格式', ['data: {"choices":{}}\n\n']],
  ['结束后仍输出内容', [deltaFrame('{}'), stopFrame, deltaFrame('额外内容'), doneFrame]],
  ['无效UTF8', [new Uint8Array([0xff, 0xfe])]],
  ['超大响应', [': ' + 'x'.repeat(2 * 1024 * 1024)]],
]) {
  test(`流式${name}不能当成成功`, async () => {
    const f = fixture({ fetch: (_call, n) => n === 1 ? sseResponse(parts) : streamedCompletion() });
    const events = await streamEvents(f), done = events.at(-1);
    assert.equal(done.event, 'done');
    assert.equal(done.data.provider, 'deepseek');
    assert.equal(done.data.ai_status.openai.reason, 'invalid_response');
    assert.equal(events.filter(e => e.event === 'done').length, 1);
    assert.equal(f.activeTimers(), 0);
  });
}

for (const [name, response] of [
  ['delta.refusal', () => sseResponse([deltaFrame('{"draft":'), 'data: {"choices":[{"index":0,"delta":{"refusal":"不能提供"}}]}\n\n'])],
  ['content_filter', () => sseResponse([deltaFrame('{"draft":'), deltaFrame('', { finish_reason: 'content_filter' })])],
  ['HTTP内容拒绝', () => upstreamError(403, 'content_policy_violation', 'private-upstream-detail')],
  ['HTTP非法请求', () => upstreamError(400, 'invalid_request_error', 'private-upstream-detail')],
  ['流内内容拒绝', () => sseResponse(['data: {"error":{"code":"content_policy_violation","message":"private-upstream-detail"}}\n\n'])],
  ['流内400错误', () => sseResponse(['data: {"error":{"status":400,"message":"private-upstream-detail"}}\n\n'])],
]) {
  test(`流式${name}不尝试另一家，也不发done`, async () => {
    const f = fixture({ fetch: response }), events = await streamEvents(f);
    assert.equal(events.at(-1).event, 'error');
    assert.equal(events.at(-1).data.provider, null);
    assert.equal(events.at(-1).data.ai_status.openai.reason, 'request');
    assert(!events.some(e => e.event === 'done' || e.event === 'reset'));
    assert.equal(f.calls.length, 1);
    assertPrivate({ body: events });
  });
}

for (const phase of ['headers', 'body']) {
  test(`流式等待${phase}超时中止主请求后兜底，清理计时器和reader`, async () => {
    let cancelled = false;
    const f = fixture({ shortTimeouts: true, fetch: (_call, n) => {
      if (n > 1) return streamedCompletion();
      if (phase === 'headers') return new Promise(() => {});
      return new Response(new ReadableStream({
        start(controller) { controller.enqueue(encoder.encode(deltaFrame('{"draft":'))); },
        cancel() { cancelled = true; },
      }));
    } });
    const events = await streamEvents(f), done = events.at(-1);
    assert.equal(done.event, 'done');
    assert.equal(done.data.provider, 'deepseek');
    assert.equal(done.data.ai_status.openai.reason, 'timeout');
    assert.equal(f.calls[0].signal.aborted, true);
    assert.equal(events.some(e => e.event === 'reset'), phase === 'body');
    if (phase === 'body') assert.equal(cancelled, true);
    assert.equal(f.activeTimers(), 0);
  });
}

for (const method of ['reader', 'request']) {
  test(`客户端通过${method}取消时中止上游，不发备用请求`, async () => {
    let cancelled = false;
    const client = new AbortController();
    const f = fixture({ fetch: () => new Response(new ReadableStream({
      start(controller) { controller.enqueue(encoder.encode(deltaFrame('{"draft":'))); },
      cancel() { cancelled = true; },
    })) });
    const events = eventReader(await f.rawRequest(streamBody, { signal: client.signal }));
    assert.equal((await events.next()).event, 'status');
    assert.equal((await events.next()).event, 'delta');
    if (method === 'reader') await events.reader.cancel(); else client.abort();
    await nextTurn();
    assert.equal(f.calls[0].signal.aborted, true);
    assert.equal(cancelled, true);
    assert.equal(f.calls.length, 1);
    assert.equal(f.activeTimers(), 0);
    if (method === 'request') assert.equal(await events.next(), null);
  });
}

test('等待响应头时取消，即使失控mock稍后返回也不读取其正文或启动备用', async () => {
  let release, cancelled = false;
  const client = new AbortController(), f = fixture({ fetch: () => new Promise(resolve => { release = resolve; }) });
  const events = eventReader(await f.rawRequest(streamBody, { signal: client.signal }));
  assert.equal((await events.next()).event, 'status');
  client.abort();
  assert.equal(await events.next(), null);
  release(new Response(new ReadableStream({ cancel() { cancelled = true; } })));
  await nextTurn();
  assert.equal(cancelled, true);
  assert.equal(f.calls.length, 1);
  assert.equal(f.activeTimers(), 0);
});

test('流式两家失败和未配置都只返回受控error，权限失败仍是HTTP错误且无上游请求', async () => {
  const failed = fixture({ fetch: (_call, n) => { if (n === 1) throw new Error('private-upstream-detail'); return upstreamError(503, '', 'private-upstream-detail'); } });
  const events = await streamEvents(failed);
  assert.equal(events.at(-1).event, 'error');
  assert.equal(events.at(-1).data.ai_status.openai.reason, 'network');
  assert.equal(events.at(-1).data.ai_status.deepseek.reason, 'server');
  assert.equal(events.at(-1).data.provider, null);
  assertPrivate({ body: events });
  const empty = fixture({ env: {} }), noKeys = await streamEvents(empty);
  assert.deepEqual(noKeys.map(e => e.event), ['error']);
  assert.equal(noKeys[0].data.ai_status.openai.state, 'not_configured');
  assert.equal(empty.calls.length, 0);
  for (const [settings, options, status] of [[{}, { noAuth: true }, 401], [{ member: false }, {}, 403], [{ userError: true }, {}, 401]]) {
    const f = fixture(settings), response = await f.request(streamBody, options);
    assert.equal(response.status, status);
    assert.equal(f.calls.length, 0);
    assertPrivate(response);
  }
});

test('流内额度/限流错误保留原因并兜底，每次新请求仍优先主服务', async () => {
  for (const [code, reason] of [['insufficient_quota', 'quota'], ['rate_limit_exceeded', 'rate_limit']]) {
    let attempts = 0;
    const f = fixture({ fetch: call => {
      if (call.url.includes('openai') && ++attempts === 1) return sseResponse([`data: ${JSON.stringify({ error: { code, message: 'private-upstream-detail' } })}\n\n`]);
      return streamedCompletion();
    } });
    const failed = await streamEvents(f);
    assert.equal(failed.at(-1).data.provider, 'deepseek');
    assert.equal(failed.at(-1).data.ai_status.openai.reason, reason);
    assert.equal((await streamEvents(f)).at(-1).data.provider, 'openai');
    assertPrivate({ body: failed });
  }
});

test('流式普通文本与视觉自定义模型保持原设置，Luna参数不会传给其他型号', async () => {
  const f = fixture({ env: { ...KEYS, OPENAI_VISION_MODEL: 'custom-vision', DEEPSEEK_VISION_MODEL: 'custom-backup' },
    fetch: (_call, n) => n === 1 ? upstreamError(500) : streamedCompletion('普通文本回复') });
  const events = await eventReader(await f.rawRequest({ messages, vision: true, stream: true })).all();
  assert.equal(events.at(-1).data.content, '普通文本回复');
  assert.equal(events.at(-1).data.model, 'custom-backup');
  assert.equal(f.calls[0].body.model, 'custom-vision');
  for (const call of f.calls) {
    assert.equal(call.body.reasoning_effort, undefined);
    assert.equal(call.body.response_format, undefined);
    assert.equal(call.body.stream, true);
  }
});
