// ═══════════════════════════════════════════════════════════════════
// AI 编程训练营 · 本文件只存 SOP 内容和常用链接。
//
// 营期（第几期、哪天开营、跳过哪几周）在页面上直接添加和管理：
//   - 数据保存在浏览器 localStorage，并实时同步到网址 # 后面
//   - 把网址复制发给别人，对方打开就能看到同样的营期配置
//
// 页面根据开营日期自动推算：
//   开营周   = 开营直播当周（月底/下月初开营，周一 ~ 周日开营直播）
//   第1~12周 = 开营后的12个自然周（周一起算）
//   第12周   = 结营周（周日结营直播）
//   休息周   = 添加营期时勾选的长假周，课程整体顺延一周
//   ⚠ 考试周暂未确定，考试相关任务暂不排入；确定后在此文件补充
// ═══════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════
// ① 常用链接
// ═══════════════════════════════════════════════════════════════════
const LINKS = {
  duxue:    { text: '督学 SOP',        url: 'https://n6fo0mbcz6.feishu.cn/wiki/ZoqkwjnbEiUX1xkYIYYcrddRnUh' },
  dayi:     { text: '答疑 SOP',        url: 'https://n6fo0mbcz6.feishu.cn/wiki/PbvawohPWi6XE5k06fecCzTvnfd' },
  jiangshi: { text: '辅助讲师 SOP',    url: 'https://n6fo0mbcz6.feishu.cn/wiki/FI2qwF0fbiEasckCGTdcdUPXndd' },
  pyq:      { text: '朋友圈公式',      url: 'https://n6fo0mbcz6.feishu.cn/wiki/WzSuwr9uxi6BpMkZBeGcdO7JnZc' },
  pigai:    { text: '批改文档模板',    url: 'https://n6fo0mbcz6.feishu.cn/wiki/DpoQwGGIAiHpnfkd4BUcAUOlnpe' },
  anli:     { text: '好评案例收集',    url: 'https://n6fo0mbcz6.feishu.cn/docx/ORt5dOYn8oKqEKx5q4icZeHVnof' },
  zhibo:    { text: '直播重建说明',    url: 'https://n6fo0mbcz6.feishu.cn/minutes/obcnloc6sc5xg9789t77331u' },
  yunpan:   { text: '云空间案例',      url: 'https://n6fo0mbcz6.feishu.cn/drive/folder/Inb5fsXzrlNZqjdN7gTcPKedn1g' },
  wenjuan:  { text: '入学档案问卷',    url: 'https://n6fo0mbcz6.feishu.cn/share/base/form/shrcnISwi4LEd5CrlnFrOs2gCgd' },
  tijiao:   { text: '作业提交指南',    url: 'https://n6fo0mbcz6.feishu.cn/docx/O02LdlCDHorC5ExnjBgc39Nynug' },
  fenxiang: { text: '分享稿案例',      url: 'https://n6fo0mbcz6.feishu.cn/wiki/JMF5wm7s3ixxAukpCNgcjqJanrg' },
  top5:     { text: '名单表格示例',    url: 'https://n6fo0mbcz6.feishu.cn/wiki/Wj1fwyPjBiywgLkwalYc0iqtnMc' },
  jiangpin: { text: '奖品信息问卷',    url: 'https://n6fo0mbcz6.feishu.cn/wiki/CRPmwNtQti382gkKIhLcDRi7njd' },
};

// ═══════════════════════════════════════════════════════════════════
// ② SOP 内容 —— 班主任每周工作（改动频率低）
//    day:    1~6 = 周一~周六，0 = 周日，'all' = 全周/日常
//    time:   有硬时间点的任务填这里（如 '15:00前'、'直播前'），页面会标红并按时间排序
//    follow: true = 跨周跟进项（如寄奖品），完成勾掉前会一直出现在今日待办
// ═══════════════════════════════════════════════════════════════════

// ── 休息周（长假顺延，添加营期时勾选）──
const SOP_REST = [
  { day: 'all', t: '长假休息周：本周无课程更新，后续课程与作业顺延一周' },
  { day: 'all', t: '提前在群里通知学员假期安排与课程恢复时间', d: '确认小鹅通每周课程/作业的解锁时间已相应调整' },
  { day: 'all', t: '假期期间关注群消息，答疑可降低频率' },
];

// ── 开营周（月底准备，周日开营直播）──
const SOP_OPEN = [
  { day: 'all', t: '接待已购学员（开营前持续进行）',
    d: '核对订单后，给学员发入学档案问卷；B站学员需额外打标签，并发送鹅圈子兑换码', link: 'wenjuan' },
  { day: 'all', t: '创建本期课程',
    d: '鹅圈子复制并改每周解锁时间；答疑直播全部删除重建（批量创建-视频直播-传统直播间）；B站课程视频复制整理', link: 'zhibo' },
  { day: 'all', t: '更新本期云空间配套资料',
    d: '新建「【AI编程训练营】第x期 配套资料」共享文件夹：上课链接文档、直播答疑文档文件夹、配套代码链接文档、讲义/常见问题/赠品快捷方式', link: 'yunpan' },
  { day: 'all', t: '建立学员小队群',
    d: '按群模板建群 → 改群名「xx队 AI编程训练营x期交流群」→ 发布群公告（资料云空间、提交作业方式、答疑时间）', link: 'tijiao' },
  { day: 'all', t: '更新开营直播 keynote 与逐字稿', d: '更新小队分组、上课链接与小程序码、新的成功案例' },
  { day: 5, t: '新建本期鹅圈子兑换码，发给 B 站渠道学员', d: 'B站学员需通过兑换码进入学员专属圈子' },
  { day: 6, t: '学员随机分组，开营前拉进对应小队群', d: '未分队学员来问时：告知分队并拉群' },
  { day: 6, t: '鹅圈子新建小队并把成员移入对应小队' },
  { day: 0, time: '直播前', t: '把学员拉进小队群 + 发送开营直播提醒', d: '1v1 私聊全员 + 小队群 @所有人' },
  { day: 0, time: '晚上', t: '开营典礼直播', d: '直播中维护讨论区秩序' },
  { day: 0, t: '开营直播后收尾', d: '群发回看提醒' },
];

// ── 常规周（第1~12周每周固定动作）──
const SOP_REGULAR = [
  { day: 1, t: '发布课程及打卡等更新提醒', link: 'duxue' },
  { day: 1, t: '公布上周成长值榜单', link: 'duxue', notW1: true },
  { day: 1, t: '直播答疑文档同步到群公告的答疑文档文件夹' },
  { day: 1, t: '发布朋友圈 ①', d: '成果案例/好评/人设/价值观/日常/资讯，每周3条', link: 'pyq' },
  { day: 1, t: '开始整理上周作业批改', link: 'pigai', notW1: true },
  { day: 2, t: '私聊搜刮一波好评和反馈' },
  { day: 2, t: '查看鹅圈子成果分享打卡内容，跟进学员' },
  { day: 3, t: '批改完上周作业，发送到学员打卡评论区', link: 'duxue', notW1: true },
  { day: 3, t: '发布朋友圈 ②', link: 'pyq' },
  { day: 4, t: '督促前两周未完成作业/随堂练习的同学，询问有无困难', link: 'duxue', notW1: true },
  { day: 4, t: '选出优秀分享给出精选/点赞', d: '部分可留作朋友圈素材' },
  { day: 5, t: '提醒本周作业截止时间', link: 'duxue' },
  { day: 5, t: '开启本周答疑问卷收集', d: '问卷设置截止时间：本周六 14:00（老师直播在周六）' },
  { day: 5, t: '发布朋友圈 ③', d: '每周3条一般在周五前发完', link: 'pyq' },
  { day: 5, t: '收集好评/成果案例 3 条至案例文档', link: 'anli' },
  { day: 6, time: '14:00前', t: '结束答疑问卷收集，整理后发给答疑老师' },
  { day: 0, t: '提醒本周作业提交截止' },
  { day: 'all', t: '日常：新购付费学员接待入群', d: '订单核验 → 发入学档案问卷（B站学员另发圈子兑换码）→ 填完按小队人数均衡分队 → 拉小队群 + 鹅圈子分队', link: 'wenjuan' },
  { day: 'all', t: '日常：督学 / 答疑 / 辅助讲师直播答疑', d: '班主任答疑：工作日 10:00–20:00；VIP讲师答疑：24小时内', link: 'duxue' },
];

// ── 特殊周叠加任务（考试周未定，暂无考试相关任务）──
const SOP_EXTRAS = {
  1: [
    { day: 'all', t: '邀请答疑讲师进新小队群并作介绍', d: '先私聊讲师确认，入群后按话术模板介绍', tag: 'open' },
  ],
};

// ── 结营周（第12周，替代大部分常规动作）──
const SOP_CLOSING = [
  { day: 1, t: '群发私聊学员成果调研问卷，收集成果与反馈', d: '被选中作分享案例的有惊喜奖励', tag: 'close' },
  { day: 3, t: '选出两名优秀学员，邀请做结营文字分享', tag: 'close' },
  { day: 3, t: '整理学员对话版分享逐字稿，请学员确认修改', link: 'fenxiang', tag: 'close' },
  { day: 3, t: '鹅圈子添加结营证书', d: '复制往期证书，选择对应期数的随堂练习/作业/打卡', tag: 'close' },
  { day: 3, t: '统计冠军队伍和 Top 5 成员名单', d: '冠军队伍 = Top 1 次数最多小队；Top 5 = 总成长值最高的五位学员', link: 'top5', tag: 'close' },
  { day: 3, t: '更新结营直播 keynote 与逐字稿', d: '获奖小队/学员头像/奖品图/小队名单；更新优惠券和推荐链接', tag: 'close' },
  { day: 0, time: '上午', t: '发送结营直播预告 + 文字分享预告', d: '晚 19:30 小鹅通直播；下午 17:00 发优秀学员对话分享', tag: 'close' },
  { day: 0, time: '17:00', t: '交流群发布优秀学员对话版分享', tag: 'close' },
  { day: 0, time: '19:30', t: '结营典礼直播', d: '直播中提醒获奖学员找班班填写奖品信息收集问卷', link: 'jiangpin', tag: 'close' },
  { day: 0, time: '直播后', t: '群发优惠券与推广返现说明', tag: 'close' },
  { day: 'all', follow: true, t: '邮寄实物奖励及发放虚拟奖励', link: 'jiangpin', tag: 'close' },
  { day: 'all', t: '常规督学动作照常进行', d: '作业批改、答疑问卷、朋友圈等按每周节奏正常执行' },
];
