const test = require('node:test');
const assert = require('node:assert/strict');
const K = require('../knowledge.js');

test('资料分块不丢字、不切开 emoji，每块最多900字并重叠100字', () => {
  const text = '课程说明'.repeat(300) + '😀练习'.repeat(800);
  const chunks = K.chunkText(text);
  assert(chunks.length > 2);
  assert(chunks.every(c => Array.from(c).length <= 900));
  assert.equal(chunks[0] + chunks.slice(1).map(c => Array.from(c).slice(100).join('')).join(''), text);
  chunks.slice(1).forEach((c, i) => assert.equal(Array.from(chunks[i]).slice(-100).join(''), Array.from(c).slice(0, 100).join('')));
  assert.equal(K.chunkText('甲'.repeat(650) + '\n' + '乙'.repeat(500))[0], '甲'.repeat(650) + '\n');
  assert.deepEqual(K.chunkText(' \r\n\t'), []);
  assert.throws(() => K.chunkText('甲'.repeat(200001)), /200000/);
});

test('检索词包含中文双字与英文关键词，去重且有上限', () => {
  const terms = K.queryTerms('Python Python 课程退款如何处理？');
  assert.equal(terms.filter(x => x === 'python').length, 1);
  assert(terms.includes('课程') && terms.includes('退款') && terms.includes('如何'));
  assert.deepEqual(K.queryTerms('贵'), ['贵']);
  assert.deepEqual(K.queryTerms('？!!!'), []);
  assert(K.queryTerms(Array.from({ length: 200 }, (_, i) => 'term' + i).join(' ')).length <= 64);
});

const chunk = (document_id, extra = {}) => ({ chunk_id: document_id, document_id, title: '课程说明', project: '', kind: 'detail',
  channel: 'all', status: 'active', content: '退款流程说明', ordinal: 1, version: 1, source_note: '', ...extra });

test('关键词检索隔离项目和渠道，停用和无匹配资料不返回', () => {
  const data = [chunk(1), chunk(2, { project: 'Python' }), chunk(3, { project: 'AI' }),
    chunk(4, { channel: 'xiaoe' }), chunk(5, { channel: 'bilibili' }), chunk(6, { status: 'archived' })];
  const ids = opts => K.rankChunks(data, '退款', opts).map(c => c.document_id).sort();
  assert.deepEqual(ids({}), [1]);
  assert.deepEqual(ids({ project: 'Python', channel: 'xiaoe' }), [1, 2, 4]);
  assert.deepEqual(ids({ project: 'AI', channel: 'bilibili' }), [1, 3, 5]);
  assert.deepEqual(K.rankChunks(data, '火星旅行', {}), []);
  assert.deepEqual(K.rankChunks(data, '', {}), []);
});

test('命中按相关度排序，最多六段且计入来源说明预算，不改原对象', () => {
  const data = Array.from({ length: 10 }, (_, i) => chunk(i + 1, { content: '退款'.repeat(450), source_note: '示例来源'.repeat(100) }));
  data.unshift(chunk(99, { title: '退款流程', content: '退款流程先查看订单' }));
  const before = JSON.stringify(data), hits = K.rankChunks(data, '退款流程', { limit: 100 });
  assert.strictEqual(hits[0], data[0]);
  assert(hits.length <= 6);
  assert(hits.reduce((n, c) => n + Array.from(c.content + c.title + c.project + c.source_note).length + 100, 0) <= 6000);
  assert.equal(JSON.stringify(data), before);
});

test('长文标题不会让无关正文占满检索结果，只有标题匹配时最多取首块', () => {
  const data = K.chunkText('填充正文'.repeat(8000) + '\n回放保留60天').map((content, i) => chunk(1, {
    title: '甲训练营销转直播', content, ordinal: i + 1,
  }));
  data.push(chunk(2, { title: '回放说明', content: '回放可在课程页面找到', ordinal: 1 }));
  const hits = K.rankChunks(data, '错过直播了，回放保留多久');
  assert(hits.some(c => c.document_id === 1 && c.content.includes('回放保留60天')));
  assert(hits.filter(c => c.document_id === 1).every(c => c.content.includes('回放')));
  assert(hits.some(c => c.document_id === 2));
  assert.deepEqual(K.rankChunks(data, '直播').map(c => [c.document_id, c.ordinal]), [[1, 1]]);
});

test('风格示例保留原话、相关优先，没命中也选启用且适用的有限示例', () => {
  const scripts = [
    { id: 1, scene: '问时间', reply: '先说时间\n再问安排', channel: 'all', status: 'active' },
    { id: 2, scene: '退款', reply: '  先问情况  \n再确认订单', channel: 'xiaoe', status: 'active' },
    { id: 3, scene: '退款', reply: '另一渠道回复', channel: 'bilibili', status: 'active' },
    { id: 4, scene: '退款', reply: '旧话术', channel: 'all', status: 'archived' },
  ];
  const selected = K.selectStyleScripts(scripts, '退款', { channel: 'xiaoe', limit: 1 });
  assert.strictEqual(selected[0], scripts[1]);
  assert.equal(selected[0].reply, '  先问情况  \n再确认订单');
  assert.deepEqual(K.selectStyleScripts(scripts, '未匹配的问题', { channel: 'xiaoe' }).map(s => s.id), [1, 2]);
  assert.deepEqual(K.selectStyleScripts(scripts, '', { maxChars: 0 }), []);
  const long = { ...scripts[0], reply: '甲'.repeat(8000) };
  assert.deepEqual(K.selectStyleScripts([long, scripts[1]], '', { maxChars: 7000 }).map(s => s.id), [2]);
  assert.deepEqual(K.selectStyleScripts([{ ...scripts[0], reply: '甲\n'.repeat(1000) }], '', { maxChars: 7000 }), [], '消息序号也占提示预算');
});

const beginner = { id: 20, version: 3, scene: '零基础怕学不会', customer_says: '没有编程基础，担心学不会',
  reply: '  营里有不少同学也是从零开始。  \r\n我们先从安装软件练起。', channel: 'all', status: 'active' };

test('明确顾虑直接匹配既定场景与客户说法，中文同义表达可命中', () => {
  for (const message of ['没有编程基础，担心学不会', '我完全没有编程基础，担心自己跟不上', '我是编程小白，怕听不懂',
    '我一点基础都没有，能不能学会？', '零基础怕学不会']) {
    const result = K.matchScript([beginner], message, { channel: 'xiaoe' });
    assert.equal(result.type, 'exact', message);
    assert.strictEqual(result.script, beginner);
    assert.equal(result.script.reply, beginner.reply);
  }
  assert.equal(K.matchScript([{ ...beginner, scene: '课程常见问题', customer_says: '我是零基础，怕学不会' }], '没有编程基础，担心跟不上').type, 'exact');
  assert.equal(K.matchScript([{ ...beginner, scene: '担心零基础学不会', customer_says: '' }], '没有编程基础，担心跟不上').type, 'exact');
});

test('时长、有效期、上课安排与具体就业目标不能因为同属大类就当作同义问题', () => {
  for (const [scene, message] of [
    ['课程有效期', '课程学多久？'], ['课程学多久', '回放保留多久'], ['课程有效期', '回放有效期'],
    ['开课时间', '上课时间？'], ['上课时间', '什么时候开营'],
    ['就业', '我想接单'], ['接单', '我想找工作'], ['保证效果', '保证学会吗'],
  ]) {
    assert.equal(K.matchScript([{ ...beginner, scene, customer_says: '' }], message).type, 'none', scene + ' ≠ ' + message);
  }
  for (const [scene, message] of [['课程学多久', '我想问课程多久'], ['开营日期', '什么时候开营'], ['上课时段', '几点上课'], ['退款', '怎么退费']]) {
    assert.equal(K.matchScript([{ ...beginner, scene, customer_says: '' }], message).type, 'exact', scene + ' = ' + message);
  }
});

test('具体金额与时间条件保守核实，不能借通用意图直接回复原话', () => {
  for (const [scene, message] of [['嫌贵', '1999 有点贵'], ['上课时间', '周二几点上课'], ['课程有效期', '课程有效期三个月吗'], ['嫌贵', '五百块有点贵']]) {
    assert.equal(K.matchScript([{ ...beginner, scene, customer_says: message }], message).type, 'none', message);
  }
});

test('直接匹配遇到否定、反向已有基础、多重顾虑、未知补充意图时保守退出', () => {
  for (const message of ['我不是零基础，也不担心学不会', '我有编程基础，但担心跟不上', '零基础，不过不怕学不会',
    '没有编程基础，担心学不会，而且太贵了', '没有编程基础，担心学不会，平时没时间', '没有编程基础，担心学不会，想找工作',
    '没有编程基础，担心学不会，可以只学 Java 吗', '只是想问一下课程']) {
    assert.equal(K.matchScript([beginner], message).type, 'none', message);
  }
  assert.equal(K.matchScript([{ ...beginner, scene: '退款', customer_says: '怎么退款', reply: '零基础怕学不会，我们可以退款' }], '没有编程基础，担心学不会').type, 'none');
});

test('同样明确的候选不任意选一条，停用与非当前渠道不能直接匹配', () => {
  const duplicate = { ...beginner, id: 21, reply: '另一种回法' };
  const result = K.matchScript([beginner, duplicate], '没有编程基础，担心学不会');
  assert.equal(result.type, 'ambiguous');
  assert.equal(result.script, null);
  assert.deepEqual(result.candidates.map(s => s.id), [20, 21]);
  assert.equal(K.matchScript([{ ...beginner, status: 'archived' }], '零基础怕学不会').type, 'none');
  const exclusive = { ...beginner, channel: 'bilibili' };
  assert.equal(K.matchScript([exclusive], '零基础怕学不会', { channel: 'all' }).type, 'exact', '不限渠道允许唯一明确匹配');
  assert.equal(K.matchScript([exclusive, { ...beginner, id: 22, channel: 'xiaoe' }], '零基础怕学不会', { channel: 'all' }).type, 'ambiguous');
  assert.equal(K.matchScript([exclusive], '零基础怕学不会', { channel: 'xiaoe' }).type, 'none');
  assert.equal(K.matchScript([exclusive], '零基础怕学不会', { channel: 'bilibili' }).type, 'exact');
});
