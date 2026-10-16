const test = require('node:test');
const assert = require('node:assert/strict');
const { readResponse } = require('../stream.js');
const encode = s => new TextEncoder().encode(s);
const event = (type, data) => `event: ${type}\r\ndata: ${JSON.stringify(data)}\r\n\r\n`;
function fixture() {
  let controller, cancelled = false;
  const body = new ReadableStream({ start(c) { controller = c; }, cancel() { cancelled = true; } });
  return { response: new Response(body, { headers: { 'content-type': 'text/event-stream' } }),
    send(s) { controller.enqueue(typeof s === 'string' ? encode(s) : s); }, close() { controller.close(); }, get cancelled() { return cancelled; } };
}
const tick = () => new Promise(resolve => setImmediate(resolve));
test('SSE 首段在完成前交付，UTF8字节拆分、CRLF及重置保持顺序', async () => {
  const f = fixture(), events = [];
  let completed = false;
  const promise = readResponse(f.response, { onEvent: e => events.push(e) }).then(x => { completed = true; return x; });
  const bytes = encode(': keepalive\r\n\r\n' + event('status', { provider: 'openai' }) + event('delta', { text: '中文草稿' }));
  for (const byte of bytes) f.send(Uint8Array.of(byte));
  await tick();
  assert.equal(completed, false);
  assert.equal(events[1].text, '中文草稿');
  f.send(event('reset', { message: '切换' }));
  f.send(event('delta', { text: '备用' }));
  f.send(event('done', { content: '{"text":"备用"}', provider: 'deepseek' }));
  const result = await promise;
  assert.equal(result.streaming, true);
  assert.equal(result.provider, 'deepseek');
  assert.deepEqual(events.map(e => e.type), ['status', 'delta', 'reset', 'delta', 'done']);
  assert.equal(f.cancelled, true);
});
test('多data行按SSE规范组合', async () => {
  const f = fixture();
  f.send('event: done\ndata: {\ndata: "content":"ok"}\n\n');
  assert.equal((await readResponse(f.response)).content, 'ok');
});
test('没有done的断流、坏JSON和空事件对象不能作为完成结果', async () => {
  for (const text of [event('delta', { text: '半截' }), 'event: done\ndata: bad\n\n', 'event: error\ndata: null\n\n', event('done', { content: '' })]) {
    const f = fixture(); f.send(text); f.close();
    await assert.rejects(readResponse(f.response), /未完成|格式异常|完整内容/);
  }
});
test('服务端错误保留状态，取消立即终止reader', async () => {
  const f = fixture(), status = { openai: { state: 'unavailable' } };
  f.send(event('error', { error: '生成失败', ai_status: status }));
  await assert.rejects(readResponse(f.response), e => e.message === '生成失败' && e.ai_status.openai.state === 'unavailable');
  const g = fixture(), controller = new AbortController();
  const pending = readResponse(g.response, { signal: controller.signal });
  controller.abort();
  await assert.rejects(pending, e => e.name === 'AbortError' && e.code === 'cancelled');
  assert.equal(g.cancelled, true);
});
test('旧JSON响应复用一次且提示部署，错误不作成功', async () => {
  const seen = [];
  const result = await readResponse(Response.json({ content: '{}', provider: 'openai' }), { onEvent: e => seen.push(e) });
  assert.equal(result.streaming, false);
  assert.deepEqual(seen.map(e => e.type), ['legacy']);
  await assert.rejects(readResponse(Response.json({ error: '请登录' }, { status: 401 })), e => e.code === 'auth');
  await assert.rejects(readResponse(Response.json({ error: 'missing' }, { status: 404 })), /部署/);
  await assert.rejects(readResponse(Response.json(null)), /完整内容/);
});
