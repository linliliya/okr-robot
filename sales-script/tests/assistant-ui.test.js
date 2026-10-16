const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const L = require('../logic.js');
function render(result) {
  const window = { SalesLogic: L, SalesReplyContent: require('../reply-content.js') }, document = { addEventListener() {} };
  const source = fs.readFileSync(require.resolve('../app.js'), 'utf8').replace('  boot().catch(renderError);', '  window.renderSuggestion = renderSuggestion; window.state = S;');
  vm.runInNewContext(source, { window, document });
  window.state.scripts = [{ id: 1, scene: '零基础怕学不会', version: 2 }];
  const el = { innerHTML: '', querySelectorAll: () => [], querySelector: () => null };
  window.renderSuggestion(el, result, null, '客户问题');
  return el.innerHTML;
}
const reply = { text: '原话第一条\n原话第二条', based_on: [1], knowledge_refs: [] };
test('历史原话依据使用保存时版本，不冒充当前版本', () => {
  const html = render({ source: 'library', matched_script: { id: 1, version: 1 }, replies: [reply],
    context_used: { sent_to_ai: false, knowledge_status: 'skipped', style_scripts: [{ id: 1, version: 1, scene: '零基础怕学不会', reply: reply.text }], knowledge: [] } });
  assert.match(html, /依据：<a[^>]+>「零基础怕学不会」v1<\/a>/);
  assert.doesNotMatch(html, /依据：<a[^>]+>「零基础怕学不会」v2<\/a>/);
  assert.match(html, /已匹配话术库原话 · 未调用 AI/);
  assert.match(html, /未检索知识库/);
  assert.doesNotMatch(html, /以上片段已提供给 AI/);
});
test('旧历史没有版本快照时显示未知，不用当前版本补造', () => {
  const html = render({ replies: [reply] });
  assert.match(html, /依据：<a[^>]+>「零基础怕学不会」（版本未知）/);
});
test('模型输出中的HTML只按消息文字显示', () => {
  const html = render({ replies: [{ text: '<img src=x onerror=bad>\n第二条', based_on: [] }] });
  assert.doesNotMatch(html, /<img/);
  assert.match(html, /&lt;img src=x onerror=bad&gt;/);
});


test('内部提醒单独显示，正常客户正文保留原话和消息', () => {
  const html = render({ internal_notes: ['核实结营后的答疑安排'], replies: [{ text: '你更关心观看期限，还是答疑服务？\n我帮你确认一下这期的安排。', internal_notes: ['对照当前期次的服务说明'], based_on: [] }] });
  assert.match(html, /销售内部提示/);
  assert.match(html, /不会复制/);
  assert.match(html, /data-copy="0"/);
  const messages = [...html.matchAll(/<div class="message-text">([^<]*)<\/div>/g)].map(x => x[1]);
  assert.deepEqual(messages, ['你更关心观看期限，还是答疑服务？', '我帮你确认一下这期的安排。']);
  assert(!messages.join('').includes('对照当前期次'));
});
test('历史正文混入来源或待核实时整份方案不提供复制，也不保留去掉限定的承诺', () => {
  for (const text of ['同学，按目前通用资料，课程长期有效。', '课程长期有效\n【待核实】这期是否适用', '放心，按资料的安排可以一直看。']) {
    const html = render({ replies: [{ text, based_on: [] }] });
    assert.doesNotMatch(html, /data-copy[=-]|class="message-text"/);
    assert.match(html, /查看需要整理的原稿/);
    assert.match(html, /重新生成后再发送/);
  }
});
test('一个方案被隔离不阻止其他完整且干净的方案复制', () => {
  const html = render({ replies: [{ text: '按通用资料，长期有效', based_on: [] }, { text: '你希望每周安排多久学习？', based_on: [] }] });
  assert.doesNotMatch(html, /data-copy="0"/);
  assert.match(html, /data-copy="1"/);
  assert.match(html, /class="message-text">你希望每周安排多久学习？/);
});
test('内部备注中的HTML按文字显示', () => {
  const html = render({ replies: [{ text: '正常回复', internal_notes: ['<img src=x onerror=bad>'], based_on: [] }] });
  assert.doesNotMatch(html, /<img/);
  assert.match(html, /&lt;img src=x onerror=bad&gt;/);
});
test('流式预览只显示完整消息，不展示半截标记及内部来源', () => {
  const window = { SalesLogic: L, SalesReplyContent: require('../reply-content.js') }, document = { addEventListener() {} };
  const source = fs.readFileSync(require.resolve('../app.js'), 'utf8').replace('  boot().catch(renderError);', '  window.preview = assistantPreviewHtml;');
  vm.runInNewContext(source, { window, document });
  assert.doesNotMatch(window.preview([{ text:'同学，按目前通', text_complete:false }]), /class="message-text"/);
  assert.doesNotMatch(window.preview([{ text:'同学，按目前通用资料，长期有效', text_complete:true }]), /class="message-text"/);
  assert.match(window.preview([{ text:'先说说你的学习目标\n第二条还没写完', text_complete:false }]), /class="message-text">先说说你的学习目标/);
  assert.doesNotMatch(window.preview([{ text:'先说说你的学习目标\n第二条还没写完', text_complete:false }]), /第二条还没写完/);
  assert.match(window.preview([{ text:'单条消息已经写完', text_complete:true }]), /class="message-text">单条消息已经写完/);
});
