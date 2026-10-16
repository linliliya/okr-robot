// AI 调用状态。只记录后端实际返回的服务商，不根据浏览器网络猜测。
(function (root) {
  'use strict';

  function create(getDB) {
    const unknown = () => ({ state: 'unknown', reason: null, message: '尚未检查', checked_at: null, source: null });
    const state = {
      providers: { openai: unknown(), deepseek: unknown() },
      pending: 0, checking: false, lastCall: null, issue: '',
    };
    const listeners = new Set();
    const revisions = { openai: 0, deepseek: 0 };
    let sequence = 0, issueRevision = 0, lastCallRevision = 0, checkPromise = null;
    const snapshot = () => JSON.parse(JSON.stringify({ ...state, demo: getDB().mode === 'demo' }));
    const publish = () => listeners.forEach(fn => fn(snapshot()));
    function setIssue(message, revision) {
      if (revision < issueRevision) return;
      issueRevision = revision;
      state.issue = message;
    }

    function merge(status, source, revision) {
      if (!status || typeof status !== 'object') return false;
      let received = false;
      for (const id of ['openai', 'deepseek']) {
        const next = status[id];
        if (!next || !['available', 'unavailable', 'not_configured', 'unknown'].includes(next.state)) continue;
        received = true;
        // 未尝试的备用服务不能抹掉之前的结果，较早发起的探测也不能覆盖新调用。
        if (next.state === 'unknown' || revision < revisions[id]) continue;
        revisions[id] = revision;
        state.providers[id] = {
          state: next.state, reason: next.reason || null, message: String(next.message || ''),
          checked_at: next.checked_at || null, source,
        };
      }
      return received;
    }

    async function request(body, probe, streamOptions) {
      const revision = ++sequence;
      if (probe) state.checking = true;
      else state.pending++;
      state.issue = '';
      publish();
      try {
        let result = streamOptions && typeof getDB().invokeAIStream === 'function'
          ? await getDB().invokeAIStream(body, { ...streamOptions, onEvent(event) {
            if (event.type === 'status' || event.type === 'reset') {
              merge(event.ai_status, 'request', revision);
              publish();
            }
            if (streamOptions.onEvent) streamOptions.onEvent(event);
          } })
          : await getDB().invokeAI(body);
        if (streamOptions && typeof getDB().invokeAIStream !== 'function') {
          result = { ...result, streaming: false };
          if (streamOptions.onEvent) streamOptions.onEvent({ type: 'legacy', message: '当前页面组件尚未启用流式输出，本次已返回完整结果；请刷新页面后重试。' });
        }
        const hasStatus = merge(result && result.ai_status, probe ? 'probe' : 'request', revision);
        setIssue(hasStatus ? '' : 'AI 函数尚未升级，无法确认服务商，请按 SETUP.md 重新部署。', revision);
        if (!probe && revision >= lastCallRevision) {
          lastCallRevision = revision;
          const provider = result && ['openai', 'deepseek'].includes(result.provider) ? result.provider : null;
          state.lastCall = {
            ok: true, provider, model: String(result && result.model || ''),
            fallback: provider === 'deepseek' && result.fallback === true,
            reason: result && result.ai_status && result.ai_status.openai && result.ai_status.openai.message || '',
            at: new Date().toISOString(),
          };
        }
        return result;
      } catch (error) {
        if (error && (error.name === 'AbortError' || error.code === 'cancelled')) throw error;
        merge(error && error.ai_status, probe ? 'probe' : 'request', revision);
        setIssue(error && error.message || 'AI 服务连接失败，请重试', revision);
        if (!probe && revision >= lastCallRevision) {
          lastCallRevision = revision;
          state.lastCall = { ok: false, provider: null, model: '', fallback: false, at: new Date().toISOString() };
        }
        throw error;
      } finally {
        if (probe) state.checking = false;
        else state.pending--;
        publish();
      }
    }

    return {
      getState: snapshot,
      subscribe(fn) { listeners.add(fn); fn(snapshot()); return () => listeners.delete(fn); },
      invoke(body) { return request(body, false); },
      invokeStream(body, options) { return request(body, false, options || {}); },
      check() {
        if (getDB().mode === 'demo') return Promise.resolve(null);
        if (checkPromise) return checkPromise;
        checkPromise = request({ action: 'status' }, true).finally(() => { checkPromise = null; });
        return checkPromise;
      },
    };
  }

  if (typeof module === 'object' && module.exports) module.exports = { create };
  else root.SalesAIStatus = create(() => root.SalesDB);
})(typeof window !== 'undefined' ? window : globalThis);
