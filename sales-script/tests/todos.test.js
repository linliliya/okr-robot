// 运行：node sales-script/tests/todos.test.js
const assert = require('assert');
const L = require('../logic.js');

const a = { id: 'todo-a', text: '试用话术 A', owner: '小王', done: false };
const b = { id: 'todo-b', text: '试用话术 B', owner: '小李', done: false };
const initial = { version: 1, todos: [a, b] };
const target = L.todoTarget(initial, 0);

// 原 bug：第一页勾选 A，保存冲突后 A 已被别人删除，绝不能把当前位置的 B 标完成。
const deleted = { version: 2, todos: [b] };
assert.throws(() => L.setTodoDone(deleted, target, true), { code: 'stale_todo' });
assert.strictEqual(deleted.todos[0].done, false);

// 同名任务、重排与改名不影响身份；只修改最初选中的那一条。
const reordered = { version: 2, todos: [{ ...b, text: a.text }, { ...a, text: '调整后的任务 A' }] };
const updated = L.setTodoDone(reordered, target, true);
assert.deepStrictEqual(updated.map(t => [t.id, t.done]), [['todo-b', false], ['todo-a', true]]);
assert.strictEqual(updated[1].text, '调整后的任务 A');
assert.strictEqual(reordered.todos[1].done, false, '不直接修改数据库快照');

// 已有 ID 异常重复时拒绝操作，避免静默选错任务。
assert.throws(() => L.setTodoDone({ version: 2, todos: [a, { ...a }] }, target, true), { code: 'stale_todo' });

// 兼容旧数据：无 ID 且版本相同可以操作；版本变化后，即使文字相同，也不能按旧 index 重试。
const legacyA = { text: '跟进客户', owner: '小王', done: false };
const legacyB = { text: '跟进客户', owner: '小王', done: false };
const legacy = { version: 3, todos: [legacyA, legacyB] };
const legacyTarget = L.todoTarget(legacy, 0);
assert.deepStrictEqual(L.setTodoDone(legacy, legacyTarget, true).map(t => t.done), [true, false]);
for (const todos of [[legacyB], [legacyB, legacyA], [{ ...legacyA, text: '改名后的任务' }, legacyB]]) {
  const changed = { version: 4, todos };
  assert.throws(() => L.setTodoDone(changed, legacyTarget, true), { code: 'stale_todo' });
  assert(todos.every(t => !t.done));
}
assert.throws(() => L.setTodoDone({ version: 3, todos: [{ ...legacyA, owner: '其他人' }] }, legacyTarget, true), { code: 'stale_todo' });

// 保存总结时保留既有 ID，为旧任务补 ID；清洗文本不会改变身份。
let sequence = 0;
const clean = L.normTodos([
  { ...a, text: '  试用话术 A  ', owner: ' 小王 ' },
  { text: '  跟进客户  ', owner: ' 小李 ', done: 1 },
  { id: 'blank', text: ' ', owner: '', done: false },
], () => `generated-${++sequence}`);
assert.deepStrictEqual(clean, [a, { id: 'generated-1', text: '跟进客户', owner: '小李', done: true }]);
assert.deepStrictEqual(L.normTodos(clean, () => { throw new Error('已有 ID 不应重新分配'); }), clean);
assert.deepStrictEqual(L.normTodos(null, () => 'unused'), []);

console.log('todos.test.js 全部通过');
