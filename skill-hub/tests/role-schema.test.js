const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const KEY = "skillhub.external-sharing.preview.v1";
const previewSource = fs.readFileSync(require.resolve("../preview.js"), "utf8");
const plain = value => JSON.parse(JSON.stringify(value));

function preview(saved) {
  const storage = new Map(saved ? [[KEY, JSON.stringify(saved)]] : []);
  const context = vm.createContext({
    window: {}, URLSearchParams,
    location: { search: "?preview=1", hash: "#/", reload() {} },
    sessionStorage: {
      getItem: key => storage.get(key) ?? null,
      setItem: (key, value) => storage.set(key, value),
      removeItem: key => storage.delete(key),
    },
  });
  vm.runInContext(previewSource, context);
  return {
    client: context.window.SKILLHUB_PREVIEW_CLIENT,
    saved: () => JSON.parse(storage.get(KEY)),
  };
}

test("fresh preview supplies role examples and unlabelled new skills default to an empty list", async () => {
  const app = preview();
  const initial = app.saved();
  assert.equal(initial.schema, 3);
  assert.deepEqual(initial.skills.find(s => s.slug === "course-outline-example").roles, ["教学教研"]);
  assert.deepEqual(initial.skills.find(s => s.slug === "anthropic-pdf").roles, ["通用"]);
  const inserted = await app.client.from("skills").insert({ slug: "untagged", title: "Untitled", author_id: "demo-user" }).select().single();
  assert.equal(inserted.error, null);
  assert.deepEqual(plain(inserted.data.roles), []);
});

test("preview role upgrade preserves user entries, seed edits, files, and explicit role choices", () => {
  const original = preview().saved();
  original.schema = 1;
  original.skills.forEach(skill => { delete skill.roles; });
  original.skills[0].title = "我修改的示例标题";
  original.skills[1].roles = ["剪辑"];
  original.skills.push({ id: 100, slug: "my-skill", title: "我的条目", source_type: "upload" });
  original.files["my-skill.zip"] = { base64: "YWJj", type: "application/zip" };
  original.nextId = 101;
  const migrated = preview(original).saved();
  assert.equal(migrated.schema, 3);
  assert.equal(migrated.skills[0].title, original.skills[0].title);
  assert.deepEqual(migrated.skills[0].roles, ["设计师", "产品经理"]);
  assert.deepEqual(migrated.skills[1].roles, ["剪辑"]);
  assert.deepEqual(migrated.skills.find(s => s.id === 100), { ...original.skills[4], roles: [] });
  assert.deepEqual(migrated.files, original.files);
  assert.deepEqual(migrated.versions, original.versions);
  assert.equal(migrated.nextId, 101);
  migrated.skills[0].roles = [];
  assert.deepEqual(preview(migrated).saved(), migrated);
});

test("preview rejects invalid roles and general roles mixed with a specific role", async () => {
  const app = preview();
  for (const roles of [null, "销售", [null], [["销售"]], ["不存在的岗位"], ["通用", "销售"], ["市场运营"], ["客户服务"], ["研发工程师"], ["财务"], ["管理者"]]) {
    const result = await app.client.from("skills").update({ roles }).eq("id", 1);
    assert.equal(result.error?.code, "23514", JSON.stringify(roles));
  }
  const valid = await app.client.from("skills").update({ roles: ["销售", "课程运营"] }).eq("id", 1).select().single();
  assert.equal(valid.error, null);
  assert.deepEqual(plain(valid.data.roles), ["销售", "课程运营"]);
});

test("existing role selections migrate without losing skills or files", () => {
  const original = preview().saved();
  original.schema = 2;
  original.skills[0].roles = ["市场运营", "设计师", "客户服务", "研发工程师", "财务", "管理者", "课程运营"];
  original.skills[1].roles = ["财务", "客户服务"];
  original.files["kept.zip"] = { base64: "YWJj", type: "application/zip" };
  const migrated = preview(original).saved();
  assert.equal(migrated.schema, 3);
  assert.deepEqual(migrated.skills[0].roles, ["课程运营", "设计师"]);
  assert.deepEqual(migrated.skills[1].roles, []);
  assert.deepEqual(migrated.skills[0], { ...original.skills[0], roles: ["课程运营", "设计师"] });
  assert.deepEqual(migrated.files, original.files);
  assert.deepEqual(migrated.versions, original.versions);
  assert.deepEqual(preview(migrated).saved(), migrated);
});

test("fresh setup and the role migration share the same role constraints", () => {
  const setup = fs.readFileSync(require.resolve("../setup.sql"), "utf8");
  const migration = fs.readFileSync(require.resolve("../update-skill-roles.sql"), "utf8");
  const body = migration.slice(migration.indexOf("alter table"), migration.lastIndexOf("commit;")).trim();
  assert.ok(setup.includes(body));
  const constraint = migration.slice(migration.lastIndexOf("add constraint skills_roles_check"));
  const allowed = constraint.match(/roles <@ array\[([\s\S]*?)\]::text\[\]/)[1];
  assert.deepEqual([...allowed.matchAll(/'([^']+)'/g)].map(match => match[1]), require("../shared.js").ROLE_OPTIONS);
});
