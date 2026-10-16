// 跨多次保存的流程：保留已完成的结果，失败后可以继续。
(function (root) {
  'use strict';

  async function importScripts(items, save) {
    const saved = [];
    for (const item of items) {
      if (!item._on || item._saved) continue;
      try {
        const row = await save(item);
        item._saved = row;
        saved.push(row);
      } catch (error) {
        return { saved, error };
      }
    }
    return { saved, error: null };
  }

  function proposalTarget(proposal, scripts) {
    const id = proposal.applied_script_id || proposal.script_id;
    if (!id) return null;
    const script = scripts.find(s => Number(s.id) === Number(id));
    if (!script) throw new Error('关联的话术已不在当前话术库中，请刷新后核对，避免重复新增。');
    return script;
  }

  root.SalesWorkflows = { importScripts, proposalTarget };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.SalesWorkflows;
})(typeof window !== 'undefined' ? window : globalThis);
