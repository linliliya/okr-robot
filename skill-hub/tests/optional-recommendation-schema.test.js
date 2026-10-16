const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const KEY = "skillhub.external-sharing.preview.v1";
const previewSource = fs.readFileSync(require.resolve("../preview.js"), "utf8");

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

test("external preview entries need no recommendation and new entries default to an empty legacy value", async () => {
  const app = preview();
  app.saved().skills.forEach(skill => assert.equal(skill.recommendation, ""));
  const inserted = await app.client.from("skills").insert({
    slug: "description-only", title: "description-only", display_title: "整理文档",
    description: "根据输入整理文档。", author_id: "demo-user", source_type: "external",
    source_url: "https://example.com/description-only", source_key: "url:https://example.com/description-only",
  }).select().single();
  assert.equal(inserted.error, null);
  assert.equal(inserted.data.recommendation, "");
  assert.equal(inserted.data.description, "根据输入整理文档。");
});

test("preview preserves existing recommendation data and content when reloading or editing other fields", async () => {
  const original = preview().saved();
  original.skills[0].recommendation = "旧版填写的推荐原因，保留在历史数据里。";
  original.files["kept.zip"] = { base64: "YWJj", type: "application/zip" };
  const app = preview(original);
  assert.deepEqual(app.saved(), original);
  const updated = await app.client.from("skills").update({ description: "更新后的简介" }).eq("id", 1).select().single();
  assert.equal(updated.error, null);
  assert.equal(updated.data.recommendation, original.skills[0].recommendation);
  assert.deepEqual(app.saved().versions, original.versions);
  assert.deepEqual(app.saved().files, original.files);
});

test("fresh setup and both migrations share the same source constraint without requiring recommendation text", () => {
  const names = ["setup.sql", "update-external-skills.sql", "update-optional-recommendation.sql"];
  const sources = names.map(name => fs.readFileSync(require.resolve(`../${name}`), "utf8"));
  const constraint = source => source.match(/add constraint skills_external_source_check check \(([\s\S]*?)\n  \)/)[1];
  const expected = constraint(sources[0]);
  sources.forEach(source => {
    assert.equal(constraint(source), expected);
    assert.doesNotMatch(constraint(source), /recommendation/);
  });
  const migration = sources[2];
  assert.match(migration, /\bbegin;[\s\S]*\bcommit;/);
  assert.doesNotMatch(migration, /\b(?:update|delete|insert)\s+(?:into\s+|from\s+)?public\./i);
  assert.doesNotMatch(migration, /\bdrop\s+(?:column|policy|trigger|index)\b/i);
  assert.deepEqual([...migration.matchAll(/drop constraint if exists ([a-z_]+)/g)].map(match => match[1]), ["skills_external_source_check"]);
  sources.slice(0, 2).forEach(source => {
    assert.match(source, /recommendation text not null default ''/);
    assert.match(source, /length\(recommendation\) <= 4000/);
  });
});
