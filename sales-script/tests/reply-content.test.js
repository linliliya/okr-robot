const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const { prepare } = require('../reply-content.js');

test('干净内容保留原空格、空白行和每种换行，且不按句号或显示折行拆分', () => {
  const original = '  同学你好。可以慢慢了解！  \r\n\t\r我帮你确认一下这期安排\n\n' + '这是同一条很长的消息。'.repeat(100) + '\r\n';
  assert.deepEqual(prepare(original), {
    text: original, internal_notes: [], withheld_messages: [], has_internal_content: false,
  });
  for (const blank of ['', ' \t\r\n　', null, undefined]) {
    assert.equal(prepare(blank).text, blank == null ? '' : blank);
    assert.equal(prepare(blank).has_internal_content, false);
  }
});

test('正常客户问句、课程资料和确认动作不视为内部备注', () => {
  const messages = [
    '我帮你确认一下这期安排', '你之前看过课程详情页吗', '课程资料长期可用',
    '我再帮你核实一下，确认后告诉你。', '你可以根据自己的时间安排学习。',
    '我把课程资料发给你，你先看一下。', '你可以在资料里查看练习题。',
    '这份资料里有练习题和参考答案。', '你可以按照老师讲的步骤操作。',
    '课程详情页上有报名入口。', '我们收到资料后会跟你联系。', '这份参考资料可以收藏。',
    '这节课会教你搭建自己的知识库', '我们会练习对检索结果排序',
    '你会学习如何给知识库分块，以及如何整理资料片段。',
    '课上会讲解检索片段的选择方式，你可以先看参考片段。',
  ];
  const text = messages.join('\r\n');
  assert.equal(prepare(text).text, text);
  assert.deepEqual(prepare(text).withheld_messages, []);
});

test('内部来源及占位标志隔离整条消息，不能留下同条承诺', () => {
  const messages = [
    '课程永久有效。【待核实】这期是否适用',
    '可以随时回看。[待核实]有效期',
    '【待确认】本期安排', '［待补充］课时数', '待补充占位：价格',
    '同学，按通用资料的安排，回放长期有效', '知识库写着可以退款',
    '检索结果中未找到答疑期限', '资料片段：每周有直播',
    '来源 ref doc-123-v2-p0', '[doc-1-v10-p9] 退款政策',
    '同学，按目前资料，课程长期有效', '根据目前提供的资料，回放长期有效',
    '根据详情页，课程长期有效', '资料里建议每周预留3～4小时，你平时晚上方便吗？',
    '详情页中注明结营时间', '资料显示每周有直播', '内部备注：服务期限需核对', '  待核实：本期安排  ',
    '根据知识库，回放长期有效', '知识库里显示有效期为一年', '本次检索结果未提供服务期限',
    '检索到的资料提到每周直播', '参考片段：课程长期有效', '参考片段 1：回放安排',
  ];
  for (const text of messages) {
    const result = prepare(text);
    assert.equal(result.text, '', text);
    assert.deepEqual(result.withheld_messages, [text], text);
    assert.deepEqual(result.internal_notes, [text], text);
    assert.equal(result.has_internal_content, true, text);
  }
});

test('按整条隔离后保留干净消息的字面空格和原分隔符', () => {
  const first = '  你好呀  ';
  const last = '\t你想先了解哪个部分？ ';
  const original = first + '\r\n按目前资料，回放长期有效\r\n' + last + '\r';
  const result = prepare(original);
  assert.equal(result.text, first + '\r\n' + last + '\r');
  assert.deepEqual(result.withheld_messages, ['按目前资料，回放长期有效']);
  assert.equal(original, first + '\r\n按目前资料，回放长期有效\r\n' + last + '\r');
  assert.equal(prepare('【待确认】期限\r你好\n再见').text, '你好\n再见');
  assert.equal(prepare('你好\r【待确认】期限').text, '你好\r');
  assert.equal(prepare('\r\n【待确认】期限\r\n\t').text, '');
});

test('明确未填业务字段整条隔离，课程标题和已填标签保留', () => {
  const placeholders = [
    '现在报名有优惠，优惠截止是【优惠截止时间】，可以先帮你锁定。',
    '定金已经帮你锁定名额啦。尾款截止是【尾款截止时间】，开营前会先发预习资料。',
    '报名截止是[报名截止日期]，错过就没有名额了。',
    '本次活动截止到［ 活动截止时间 ］。',
    '我们会在【开营日期】开营。', '回放可以看到【有效期】。',
  ];
  for (const message of placeholders) {
    const result = prepare(message + '\r\n你现在主要卡在哪一步？');
    assert.equal(result.text, '你现在主要卡在哪一步？', message);
    assert.deepEqual(result.withheld_messages, [message], message);
    assert.deepEqual(result.internal_notes, [message], message);
  }
  const clean = [
    '你可以先看看【Python 入门训练营】的课程介绍。',
    '你看到的【课程标题】是页面上的栏目名称。',
    '【直播安排】每周三晚上八点，我们一起答疑。',
    '【优惠截止时间：本周五20点】你可以提前安排。',
    '【尾款截止时间说明】这部分我可以再帮你解释。',
    '这份【课程资料】可以随时查看。',
  ].join('\r\n');
  assert.equal(prepare(clean).text, clean);
  assert.deepEqual(prepare(clean).withheld_messages, []);
});

// 以下摄影工作坊和安排均为虚构回归样例。
test('两条消息泄漏：两条消息均整体隔离，不把未证实承诺留作可复制话术', () => {
  const text = '同学，按目前通用资料，示例摄影工作坊的练习包可使用90天，六周活动结束后还可以继续练习\n六周是资料里的集中练习周期，和练习包有效期不同。【待核实】你问的这一组是否沿用此安排，以及作品点评开放到哪一天';
  const result = prepare(text);
  assert.equal(result.text, '');
  assert.deepEqual(result.withheld_messages, text.split('\n'));
  assert.deepEqual(result.internal_notes, text.split('\n'));
});

test('三条消息泄漏：中间混合消息也整体隔离，不能仅删除资料归因措辞', () => {
  const text = '同学，按通用资料的安排，示例摄影工作坊不用在六周内做完所有练习，练习包可使用90天\n这周没空可以下周继续。资料里建议每周预留1～2小时，你通常哪天方便练习？\n【待核实】这组是否适用以上安排，以及作品点评可以提交到哪一天';
  const result = prepare(text);
  assert.equal(result.text, '');
  assert.deepEqual(result.withheld_messages, text.split('\n'));
});

test('传入内部备注不污染客户文本或伪装成被隔离消息，输入数组不变', () => {
  const notes = Object.freeze(['  先核对本期安排  ', '', '  先核对本期安排  ', null, { note: '对象忽略' }]);
  const result = prepare('我帮你确认一下这期安排', notes);
  assert.equal(result.text, '我帮你确认一下这期安排');
  assert.deepEqual(result.internal_notes, ['  先核对本期安排  ']);
  assert.deepEqual(result.withheld_messages, []);
  assert.equal(result.has_internal_content, true);
  assert.equal(notes.length, 5);
  assert.deepEqual(prepare('你好', '核对价格').internal_notes, ['核对价格']);
});

test('隔离诊断保留重复原消息，内部备注去重，连续调用无正则状态泄漏', () => {
  const note = '  【待核实】服务期限  ';
  for (let i = 0; i < 3; i++) {
    const result = prepare(note + '\r' + note, [note]);
    assert.deepEqual(result.withheld_messages, [note, note]);
    assert.deepEqual(result.internal_notes, [note]);
    assert.equal(result.text, '');
  }
});

test('浏览器 UMD 暴露相同 prepare 接口', () => {
  const context = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../reply-content.js'), 'utf8'), context);
  assert.equal(typeof context.SalesReplyContent.prepare, 'function');
  assert.equal(context.SalesReplyContent.prepare('你好\r\n').text, '你好\r\n');
  assert.equal(context.SalesReplyContent.prepare('【待核实】期限').text, '');
});
