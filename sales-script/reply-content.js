// 只分离明显的内部来源/占位备注；不改写话术，也不判断其承诺是否真实。
(function (root, factory) {
  'use strict';
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SalesReplyContent = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const INTERNAL_MARKERS = [
    /通用资料/,
    /\bdoc-\d+-v\d+-p\d+\b/i,
    /[【\[［]\s*(?:待核实|待确认|待补充|待填写|待填|需核实|需确认|占位(?:符)?|内部(?:备注|说明|提醒)?)\s*(?:[：:][^】\]］]*)?[】\]］]/,
    // 只识别明确的未填业务字段；不把任意书名、课程标题或已填写的标签当作占位符。
    /[【\[［]\s*(?:(?:优惠|尾款|报名|活动)截止(?:时间|日期)|开营(?:时间|日期)|有效期)\s*[】\]］]/,
    /(?:^|[。！？；;])\s*(?:待核实|待确认|待补充|待填写|内部备注|内部说明|销售备注|来源说明)\s*[：:]/,
    /(?:待补充|待填写|待填)\s*占位(?:符)?/,
    /(?:按照?|依据|根据|参照)\s*(?:(?:目前|当前|现有|已有|已提供|所提供|提供|上传|这份|该|本|手头|课程|招生|官方|通用)的?\s*)*(?:资料|详情页|课程介绍|课程详情|文档|知识库|检索结果|参考片段|检索片段)/,
    /(?:资料|详情页|文档)(?:里|中|上)的?\s*(?:建议|提到|写着|写明|写了|说明|说了|显示|标注|注明|介绍|提及)/,
    /(?:资料|详情页|文档)\s*(?:建议|显示|注明|标注|提到|写明)/,
    /(?:知识库|检索结果|资料片段|参考片段|检索片段)(?:里|中|上)?的?\s*(?:建议|提到|写着|写明|写了|说明|显示|标注|注明|给出|未(?:找到|提到|明确|说明)|没有(?:找到|提到|说明))/,
    /本次的?检索结果|检索到的(?:资料|片段|内容)/,
    /(?:^|[，。！？；;])\s*(?:资料片段|参考片段|检索片段|检索结果|知识库)\s*\d*\s*[:：]/,
  ];

  function prepare(text, internalNotes = []) {
    const original = String(text == null ? '' : text);
    const notes = [];
    const appendNote = note => {
      if (typeof note === 'string' && note.trim() && !notes.includes(note)) notes.push(note);
    };
    // 兼容旧记录的单条备注，同时避免把对象隐式转成客户或销售文案。
    (Array.isArray(internalNotes) ? internalNotes : [internalNotes]).forEach(appendNote);

    const withheld = [];
    const kept = [];
    // 分隔符单独保留：没有命中时整串原样返回，命中时不规范化其他消息。
    const pieces = original.split(/(\r\n|\r|\n)/);
    for (let i = 0; i < pieces.length; i += 2) {
      const message = pieces[i];
      if (INTERNAL_MARKERS.some(pattern => pattern.test(message))) {
        withheld.push(message);
        appendNote(message);
      } else {
        kept.push(message + (pieces[i + 1] || ''));
      }
    }
    const filtered = kept.join('');
    return {
      text: withheld.length ? (filtered.trim() ? filtered : '') : original,
      internal_notes: notes,
      withheld_messages: withheld,
      has_internal_content: notes.length > 0 || withheld.length > 0,
    };
  }

  return { prepare };
});
