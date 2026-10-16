// 运行：node sales-script/tests/logic.test.js
const assert = require('assert');
const L = require('../logic.js');

// 日期
assert.strictEqual(L.weekStart('2026-10-08'), '2026-10-05');
assert.strictEqual(L.weekStart('2026-10-05'), '2026-10-05');
assert.strictEqual(L.weekStart('2026-10-11'), '2026-10-05');
assert.strictEqual(L.weekStart('2026-10-12'), '2026-10-12');
assert.strictEqual(L.weekLabel('2026-12-28'), '12月28日–1月3日');
assert.strictEqual(L.addDays('2026-02-28', 1), '2026-03-01');

// 脱敏
const m = L.maskSensitive('电话 138 1234 5678，+86-13912345678，微信号：abc_12345，vx abcdef1，邮箱 a.b@c.com，身份证 11010119900307123X，wxid_ab12');
assert(!/1381234|13912345678|abc_12345|abcdef1|a\.b@c|11010119900307123X|wxid_ab12/.test(m), m);
assert.strictEqual(L.maskSensitive('价格 1999 元，第 12345 期'), '价格 1999 元，第 12345 期');

// 聊天解析
assert.deepStrictEqual(L.parseDialogue('客户：太贵了\n销售：理解\n补一句\n我：好的\n备注：[图片]'), [
  { role: 'customer', text: '太贵了' }, { role: 'sales', text: '理解\n补一句' }, { role: 'sales', text: '好的' }, { role: 'note', text: '[图片]' },
]);
assert.strictEqual(L.parseDialogue('小王 10:21\n你好'), null);
assert.strictEqual(L.dialogueToText([{ role: 'customer', text: 'a' }, { role: 'x', text: 'b' }, { role: 'sales', text: ' ' }]), '客户：a\n备注：b');

// 话术消息只按显式换行拆分，空白行不占序号，其余字符保持原样。
assert.deepStrictEqual(L.splitReplyMessages('\r\n  第一条。还有一句！  \r\n\t\r第二条\n　\n第三条\r'), [
  '  第一条。还有一句！  ', '第二条', '第三条',
]);
for (const empty of [undefined, null, '', ' \t\r\n　']) assert.deepStrictEqual(L.splitReplyMessages(empty), []);
const unbrokenReply = '这段很长，只是显示时可能折行。还有问号？'.repeat(40) + '\\n这不是实际换行';
assert.deepStrictEqual(L.splitReplyMessages(unbrokenReply), [unbrokenReply]);

// 话术库分组 / 搜索
const scripts = [
  { id: 1, stage: '直播后追单', scene: '嫌贵', reply: '先问顾虑', customer_says: '太贵了', tips: '', channel: 'all', sort_order: 0 },
  { id: 2, stage: '自定义阶段', scene: 'A', reply: 'x', customer_says: '', tips: '', channel: 'xiaoe', sort_order: 0 },
  { id: 3, stage: '邀约入群', scene: 'B', reply: 'y', customer_says: '', tips: '', channel: 'bilibili', sort_order: 0 },
  { id: 4, stage: '', scene: 'C', reply: 'z', customer_says: '', tips: '', channel: 'all', sort_order: 0 },
];
assert.deepStrictEqual(L.groupScriptsByStage(scripts).map(g => g.stage), ['邀约入群', '直播后追单', '自定义阶段', '其他']);
assert.deepStrictEqual(L.searchScripts(scripts, '太贵 顾虑').map(s => s.id), [1]);
assert.deepStrictEqual(scripts.filter(s => L.scriptChannelMatch(s, 'xiaoe')).map(s => s.id), [1, 2, 4]);
const promptScript = { ...scripts[0], reply: '  先问顾虑  \r\n\t\n再介绍课程\r最后确认下一步' };
const originalReply = promptScript.reply;
assert(L.scriptToPrompt(promptScript).includes('推荐回复（以下消息按顺序发送）：\n消息 1：  先问顾虑  \n消息 2：再介绍课程\n消息 3：最后确认下一步'));
assert.strictEqual(promptScript.reply, originalReply, '生成 AI 上下文不能改写保存的话术');

// 复盘会分组：成交流失都有的场景排前面，未分类放最后
const cases = [
  { id: 1, result: 'lost', scenes: ['没时间'] }, { id: 2, result: 'lost', scenes: ['没时间'] },
  { id: 3, result: 'won', scenes: ['嫌贵'] }, { id: 4, result: 'lost', scenes: ['嫌贵', '没时间'] },
  { id: 5, result: 'won', scenes: [] },
];
assert.deepStrictEqual(L.groupCasesByScene(cases).map(g => [g.scene, g.won.length, g.lost.length]),
  [['嫌贵', 1, 1], ['没时间', 0, 2], ['未分类', 1, 0]]);

// 效果统计：同一案例重复引用只算一次
const u = L.scriptUsage([
  { result: 'won', script_refs: [{ id: 1, version: 2 }, { id: 1, version: 2 }] },
  { result: 'lost', script_refs: [{ id: 1, version: 1 }] },
  { result: 'won', script_refs: null },
]);
assert.deepStrictEqual(u.get(1), { won: 1, lost: 1, byVersion: { 1: { won: 0, lost: 1 }, 2: { won: 1, lost: 0 } } });
assert.strictEqual(L.usageText(u.get(1)), '案例里用到 2 次 · 成交 1');

// 文本对比：末尾加一句不应该把上一句标成删除
assert.deepStrictEqual(L.diffText('能理解。你看看', '能理解。你看看\n新的一句。'),
  [{ type: 'same', text: '能理解。你看看' }, { type: 'add', text: '\n新的一句。' }]);
assert.deepStrictEqual(L.diffText('A。B。C。', 'A。X。C。'),
  [{ type: 'same', text: 'A。' }, { type: 'del', text: 'B。' }, { type: 'add', text: 'X。' }, { type: 'same', text: 'C。' }]);

// JSON 提取
assert.deepStrictEqual(L.extractJSON('```json\n{"a":1}\n```'), { a: 1 });
assert.deepStrictEqual(L.extractJSON('好的，结果如下：{"a":[1,2]} 希望有帮助'), { a: [1, 2] });
assert.strictEqual(L.extractJSON('不是 JSON'), null);

// 文档切块
const doc = ['一'.repeat(1000), '二'.repeat(1000), '三'.repeat(1500), '四'.repeat(5000)].join('\n\n');
const chunks = L.splitDoc(doc, 3000);
assert(chunks.every(c => c.length <= 3000), chunks.map(c => c.length));
assert.strictEqual(chunks.join('').replace(/\n/g, '').length, 8500);

// 导入条目整理
const s = L.normImportedScript({ stage: '直播中', scene: '嫌贵', customer_says: ['a', 'b'], reply: 'r', channel: 'weird' });
assert.strictEqual(s.customer_says, 'a\nb');
assert.strictEqual(s.channel, 'all');
assert.strictEqual(L.normImportedScript({ reply: '  第一条  \r\n第二条\n' }).reply, '  第一条  \r\n第二条\n');
assert.strictEqual(L.normImportedScript({ reply: ['  第一条  ', '第二条'] }).reply, '  第一条  \n第二条');

// 校验
assert.deepStrictEqual(L.validateScript(Object.assign(L.emptyScript(), { scene: '嫌贵', reply: 'x' })), []);
assert.strictEqual(L.validateCase(Object.assign(L.emptyCase('销售A'), { dialogue: [] })).length, 1);

console.log('logic.test.js 全部通过');
