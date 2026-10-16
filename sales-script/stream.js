// 浏览器和测试共用的 SSE 读取器。只有 done 事件表示完整结果。
(function (root) {
  'use strict';
  const MAX_CHARS = 4 * 1024 * 1024;
  const abortError = () => Object.assign(new Error('已停止生成'), { name: 'AbortError', code: 'cancelled' });
  const error = (message, payload) => Object.assign(new Error(message), {
    code: 'ai', ai_status: payload && payload.ai_status,
  });

  async function readResponse(response, opts) {
    const { signal, onEvent = () => {} } = opts || {};
    if (signal && signal.aborted) throw abortError();
    const streaming = /text\/event-stream/i.test(response.headers.get('content-type') || '');
    if (!streaming) {
      let payload;
      try { payload = await response.json(); }
      catch (e) {
        if (signal && signal.aborted) throw abortError();
        throw error('AI 中转服务返回了无法识别的内容，请检查函数部署');
      }
      if (signal && signal.aborted) throw abortError();
      if (!payload || typeof payload !== 'object') throw error('AI 没有返回完整内容');
      if (!response.ok || payload.error) {
        const err = error(payload.error || 'AI 调用失败', payload);
        if (response.status === 401) err.code = 'auth';
        if (response.status === 404) err.message = 'AI 中转函数还没部署，请按 SETUP.md 更新 sales-ai';
        throw err;
      }
      if (typeof payload.content !== 'string') throw error('AI 没有返回完整内容');
      // 旧函数忽略 stream 字段时，复用本次结果，不再付费生成一遍。
      onEvent({ type: 'legacy', message: '当前云端函数尚未启用流式输出，结果已完整返回；更新 sales-ai 后可边生成边显示。' });
      return { ...payload, streaming: false };
    }
    if (!response.ok || !response.body) throw error('流式连接未建立，请重试');
    const reader = response.body.getReader(), decoder = new TextDecoder();
    let buffer = '', total = 0, eventType = 'message', data = [], result = null;
    let rejectAbort;
    const aborted = new Promise((_resolve, reject) => { rejectAbort = reject; });
    const onAbort = () => { reader.cancel().catch(() => {}); rejectAbort(abortError()); };
    if (signal) signal.addEventListener('abort', onAbort, { once: true });
    const dispatch = () => {
      if (!data.length) { eventType = 'message'; return; }
      let payload;
      try { payload = JSON.parse(data.join('\n')); }
      catch { throw error('流式内容格式异常，本次生成未完成，请重试'); }
      if (!payload || typeof payload !== 'object') throw error('流式内容格式异常，请重试');
      const type = eventType;
      data = []; eventType = 'message';
      if (type === 'error') throw error(payload.error || '生成中断，请重试', payload);
      if (type === 'done') {
        if (!payload || typeof payload.content !== 'string' || !payload.content.trim()) throw error('AI 没有返回完整内容', payload);
        result = { ...payload, streaming: true };
      }
      if (['status', 'delta', 'reset', 'done'].includes(type)) onEvent({ ...payload, type });
    };
    const line = value => {
      if (value.endsWith('\r')) value = value.slice(0, -1);
      if (!value) { dispatch(); return; }
      if (value.startsWith('event:')) eventType = value.slice(6).trim();
      else if (value.startsWith('data:')) data.push(value.slice(5).replace(/^ /, ''));
    };
    try {
      while (!result) {
        if (signal && signal.aborted) throw abortError();
        const next = signal ? await Promise.race([reader.read(), aborted]) : await reader.read();
        if (signal && signal.aborted) throw abortError();
        if (next.done) {
          buffer += decoder.decode();
          if (buffer) line(buffer);
          // EOF 不能把一份未完成的草稿当成最终结果。
          throw error('连接提前结束，本次生成未完成，请重试');
        }
        const text = decoder.decode(next.value, { stream: true });
        total += text.length;
        if (total > MAX_CHARS) throw error('AI 输出过长，本次生成已停止，请缩小问题后重试');
        buffer += text;
        let end;
        while (!result && (end = buffer.indexOf('\n')) >= 0) {
          const value = buffer.slice(0, end); buffer = buffer.slice(end + 1); line(value);
        }
      }
      return result;
    } finally {
      if (signal) signal.removeEventListener('abort', onAbort);
      void reader.cancel().catch(() => {});
      reader.releaseLock();
    }
  }
  const api = { readResponse, abortError };
  root.SalesStream = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
