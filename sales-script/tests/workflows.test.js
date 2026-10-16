const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../workflows.js');

test('批量导入中断后保留编辑内容，重试跳过已成功项', async () => {
  const items = [
    { scene: '示例A', reply: '已人工修改的回复A', _on: true },
    { scene: '示例B', reply: '已人工修改的回复B', _on: true },
    { scene: '示例C', reply: '已人工修改的回复C', _on: true },
    { scene: '不选', reply: '保留但不导入', _on: false },
  ];
  const calls = [];
  let fail = true;
  const save = async item => {
    calls.push(item.scene);
    if (item.scene === '示例B' && fail) throw new Error('模拟断网');
    return { id: calls.length, scene: item.scene, reply: item.reply };
  };
  const first = await W.importScripts(items, save);
  assert.equal(first.error.message, '模拟断网');
  assert.deepEqual(first.saved.map(s => s.scene), ['示例A']);
  assert.equal(items[1].reply, '已人工修改的回复B');
  assert.equal(items[2].reply, '已人工修改的回复C');
  assert.equal(items[1]._saved, undefined);
  fail = false;
  const retry = await W.importScripts(items, save);
  assert.equal(retry.error, null);
  assert.deepEqual(retry.saved.map(s => s.scene), ['示例B', '示例C']);
  assert.deepEqual(calls, ['示例A', '示例B', '示例B', '示例C']);
  assert.equal(items[3]._saved, undefined);
});

test('第一条即失败时整批草稿保留', async () => {
  const items = [{ scene: '示例', reply: '修改后', _on: true }];
  const before = JSON.stringify(items);
  const result = await W.importScripts(items, async () => { throw new Error('失败'); });
  assert.equal(result.saved.length, 0);
  assert.equal(JSON.stringify(items), before);
});

test('新增建议采纳、重新讨论、再次采纳只创建一条话术', async () => {
  const scripts = [];
  const proposal = { script_id: null, state: 'open' };
  let inserts = 0, updates = 0;
  async function apply() {
    const target = W.proposalTarget(proposal, scripts);
    if (target) updates++;
    else { inserts++; scripts.push({ id: 101, scene: '示例' }); }
    Object.assign(proposal, { state: 'applied', applied_script_id: (target || scripts[0]).id });
  }
  await apply();
  proposal.state = 'open';
  await apply();
  assert.equal(inserts, 1);
  assert.equal(updates, 1);
  assert.equal(scripts.length, 1);
});

test('关联话术缺失不能退回新增，防止再次采纳制造重复项', () => {
  assert.throws(() => W.proposalTarget({ applied_script_id: 101 }, []), /刷新后核对/);
  assert.throws(() => W.proposalTarget({ script_id: 102 }, []), /刷新后核对/);
});
