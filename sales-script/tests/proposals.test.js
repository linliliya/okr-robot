// 运行：node sales-script/tests/proposals.test.js
// 执行真实 app.js 的采纳闭包，仅替换弹窗和数据保存，避免复制被测流程。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const W = require('../workflows.js');

const appSource = fs.readFileSync(path.join(__dirname, '../app.js'), 'utf8');
const start = appSource.indexOf('    const setProposal = async');
const end = appSource.indexOf('    // ② 逐个过案例', start);
assert(start >= 0 && end > start, '找不到 app.js 的建议采纳闭包，请更新测试的提取位置');
const proposalSource = appSource.slice(start, end);
const helperStart = appSource.indexOf('  function prepareCustomerReply(');
const helperEnd = appSource.indexOf('  function salesInternalHtml(', helperStart);
assert(helperStart >= 0 && helperEnd > helperStart, '找不到 app.js 的 AI 草稿检查函数');
const helperSource = appSource.slice(helperStart, helperEnd);
const contentSource = appSource.match(/^  const CONTENT_FIELDS = .*$/m)[0] + '\n' + appSource.match(/^  const pickContent = .*$/m)[0];

function fixture({ failFirstSave = false, replaceReviewOnFailure = false, proposal = {} } = {}) {
  let currentReview = {
    id: 1, week: '2026-10-05', version: 1,
    ai_brief: {
      proposals: [{ key: 'proposal-1', state: 'open', script_id: null, scene: '嫌贵', new_reply: '先问清顾虑', case_ids: [], ...proposal }],
    },
  };
  const clone = value => JSON.parse(JSON.stringify(value));
  const persistedReview = clone(currentReview);
  const scripts = [];
  const editorTargets = [];
  const editorOptions = [];
  const errors = [];
  const drawnStates = [];
  let saveCalls = 0;
  const context = {
    week: currentReview.week,
    S: { scripts },
    window: { SalesReplyContent: require('../reply-content.js') },
    W,
    review: () => currentReview,
    ensureReview: async () => currentReview,
    updateReview: async (_week, mutate) => {
      saveCalls++;
      const next = clone(currentReview);
      mutate(next);
      if (failFirstSave && saveCalls === 1) {
        // 真实 updateReview 在版本冲突后会重新读取并替换 S.reviews。
        if (replaceReviewOnFailure) currentReview = clone(persistedReview);
        throw new Error('模拟复盘状态保存断网');
      }
      next.version++;
      currentReview = next;
      return next;
    },
    openScriptEditor: async (target, options) => {
      editorTargets.push(target ? target.id : null);
      editorOptions.push(options);
      if (target) return target;
      const saved = { id: 91, version: 1, scene: '嫌贵', reply: '先问清顾虑' };
      scripts.push(saved);
      return saved;
    },
    drawBrief: () => drawnStates.push(clone(currentReview.ai_brief.proposals[0])),
    drawChanges: () => {},
    showErr: error => errors.push(error),
    toast: () => {},
  };
  vm.createContext(context);
  vm.runInContext(contentSource + '\n' + helperSource + proposalSource + '\nthis.applyProposalForTest = applyProposal;', context, { filename: 'app-proposals-under-test.js' });
  drawnStates.length = 0;
  return {
    apply: () => context.applyProposalForTest('proposal-1'),
    proposal: () => currentReview.ai_brief.proposals[0],
    scripts, editorTargets, editorOptions, errors, drawnStates,
    saveCalls: () => saveCalls,
  };
}

test('建议正常采纳后保存关联、标为已采纳，重复点击不重复创建', async () => {
  const f = fixture();
  await f.apply();
  await f.apply();
  assert.equal(f.proposal().state, 'applied');
  assert.equal(f.proposal().applied_script_id, 91);
  assert.equal(f.proposal().applied_version, 1);
  assert.deepEqual(f.editorTargets, [null]);
  assert.equal(f.scripts.length, 1);
  assert.equal(f.errors.length, 0);
  assert.equal(f.saveCalls(), 1);
});

test('建议的内部核实事项只预填到使用提示，客户正文保持原话', async () => {
  const f = fixture({ proposal: { new_reply: '先说说你的学习目标\n我帮你确认这期的安排', new_tips: '先了解学习节奏', internal_notes: ['核实结营后的答疑期限'] } });
  await f.apply();
  assert.equal(f.errors.length, 0);
  assert.equal(f.editorOptions[0].prefill.reply, '先说说你的学习目标\n我帮你确认这期的安排');
  assert.equal(f.editorOptions[0].prefill.tips, '先了解学习节奏\n核实结营后的答疑期限');
  assert.equal(f.proposal().state, 'applied');
});

test('旧建议正文混入来源或待核实时不允许采纳入库', async () => {
  for (const new_reply of ['按通用资料，长期有效', '长期有效\n【待核实】这期是否适用', '截止时间是【优惠截止时间】']) {
    const f = fixture({ proposal: { new_reply } });
    await f.apply();
    assert.equal(f.editorOptions.length, 0);
    assert.equal(f.saveCalls(), 0);
    assert.equal(f.proposal().state, 'open');
  }
});

for (const replaceReviewOnFailure of [false, true]) {
  test(`话术已保存但复盘状态失败后可重试并复用原话术${replaceReviewOnFailure ? '（期间重新读取了复盘）' : ''}`, async () => {
    const f = fixture({ failFirstSave: true, replaceReviewOnFailure });
    await f.apply();
    assert.equal(f.proposal().state, 'open', '状态未写入时仍应允许再次点击采纳');
    assert.equal(f.proposal().applied_script_id, 91, '保留已创建的话术关联');
    assert.equal(f.proposal().applied_version, 1);
    assert.equal(f.errors.length, 1);
    assert(f.drawnStates.some(p => p.state === 'open' && p.applied_script_id === 91), '失败后重绘可重试的建议');

    await f.apply();
    assert.deepEqual(f.editorTargets, [null, 91], '第二次弹窗必须编辑已创建的话术');
    assert.equal(f.scripts.length, 1, '重试不能重复新增');
    assert.equal(f.saveCalls(), 2);
    assert.equal(f.proposal().state, 'applied');
    assert.equal(f.proposal().applied_script_id, 91);
    assert.equal(f.proposal().applied_version, 1);
  });
}
