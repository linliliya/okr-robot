const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const KEY = "skillhub.external-sharing.preview.v1";
const previewSource = fs.readFileSync(require.resolve("../preview.js"), "utf8");
const plain = value => JSON.parse(JSON.stringify(value));
const KNOWN_TITLES = {
  "de-ai-copy": "去AI味写作",
  "ai-moments": "生成 AI 资讯朋友圈",
  "book-insights": "生成读书感悟朋友圈",
  "python-crawler-success": "生成 Python 爬虫训练营成功案例朋友圈",
};

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

test("preview examples have separate function titles and internal Skill names match downloaded files", () => {
  const saved = preview().saved();
  assert.equal(saved.schema, 4);
  for (const skill of saved.skills) {
    assert.ok(skill.display_title);
    assert.notEqual(skill.display_title, skill.title);
    if (skill.source_type === "upload") {
      const version = saved.versions.find(item => item.skill_id === skill.id);
      assert.equal(skill.title, version.skill_md.match(/^name: (.+)$/m)[1]);
    }
  }
});

test("display title upgrade preserves user titles, files, versions, metadata, and custom display titles", () => {
  const original = preview().saved();
  original.schema = 3;
  original.skills.forEach(skill => { delete skill.display_title; });
  original.skills[0].title = "我改过的示例";
  original.skills[1].display_title = "我填写的功能标题";
  original.skills[2].title = "课程大纲助手（示例）";
  original.skills.push({ ...original.skills[0], id: 100, slug: "mine", title: "my-own-skill", description: "保留用户填写的说明" });
  original.files["my-skill.zip"] = { base64: "YWJj", type: "application/zip" };
  original.nextId = 101;
  original.signedIn = false;

  const migrated = preview(original).saved();
  assert.equal(migrated.schema, 4);
  assert.equal(migrated.skills[0].display_title, "");
  assert.equal(migrated.skills[1].display_title, "我填写的功能标题");
  assert.equal(migrated.skills[2].display_title, "生成课程大纲（示例）");
  assert.equal(migrated.skills[2].title, "课程大纲助手（示例）");
  assert.equal(migrated.skills.at(-1).display_title, "");
  for (let i = 0; i < original.skills.length; i++) {
    const { display_title: beforeTitle, ...before } = original.skills[i];
    const { display_title: afterTitle, ...after } = migrated.skills[i];
    assert.deepEqual(after, before);
  }
  for (const field of ["files", "versions", "profiles", "nextId", "signedIn"]) {
    assert.deepEqual(migrated[field], original[field]);
  }
  assert.deepEqual(preview(migrated).saved(), migrated);
});

test("display title upgrade backfills only the four exact names and never overwrites a custom title", () => {
  const original = preview().saved();
  original.schema = 3;
  original.skills = Object.keys(KNOWN_TITLES).map((title, i) => ({
    ...original.skills[0], id: 100 + i, slug: title, title, display_title: "",
  }));
  original.skills.push(
    { ...original.skills[0], id: 104, slug: "custom", display_title: "我的自定义标题" },
    { ...original.skills[0], id: 105, slug: "case", title: "DE-AI-COPY" },
    { ...original.skills[0], id: 106, slug: "spaced", title: " de-ai-copy " },
    { ...original.skills[0], id: 107, slug: "unknown", title: "unknown-skill" },
  );
  const saved = preview(original).saved();
  for (const skill of saved.skills.slice(0, 4)) assert.equal(skill.display_title, KNOWN_TITLES[skill.title]);
  assert.equal(saved.skills[4].display_title, "我的自定义标题");
  saved.skills.slice(5).forEach(skill => assert.equal(skill.display_title, ""));
  assert.deepEqual(preview(saved).saved(), saved);
});

test("preview allows old clients, validates Unicode title length, and changes only the requested title", async () => {
  const app = preview();
  const inserted = await app.client.from("skills").insert({ slug: "old-client", title: "old-client", author_id: "demo-user" }).select().single();
  assert.equal(inserted.error, null);
  assert.equal(inserted.data.display_title, "");
  const before = app.saved();
  for (const display_title of [null, 5, [], {}, " ", "\u00a0", "\uFEFF", "\u3000", "a".repeat(81), "😀".repeat(81), "有效\n标题", "有效\r标题", "a\t", "a\0", "a\u0085", "a\u009f", "a\u2028", "a\u2029"]) {
    const result = await app.client.from("skills").update({ display_title }).eq("id", 1);
    assert.equal(result.error?.code, "23514", JSON.stringify(display_title));
    assert.deepEqual(app.saved(), before);
  }
  for (const display_title of ["", "去AI味写作", "😀".repeat(80), "  整理文档  "]) {
    const result = await app.client.from("skills").update({ display_title }).eq("id", 1).select().single();
    assert.equal(result.error, null);
    const { author, ...actual } = plain(result.data);
    assert.deepEqual(actual, { ...before.skills[0], display_title });
  }
});

test("fresh setup includes the exact display title migration and keeps original names untouched", () => {
  const setup = fs.readFileSync(require.resolve("../setup.sql"), "utf8");
  const migration = fs.readFileSync(require.resolve("../update-skill-display-titles.sql"), "utf8");
  const body = migration.slice(migration.indexOf("alter table"), migration.lastIndexOf("commit;")).trim();
  assert.ok(setup.includes(body));
  assert.match(migration, /\bbegin;[\s\S]*\bcommit;/);
  assert.match(migration, /display_title text not null default ''/);
  assert.match(migration, /where display_title = ''/);
  assert.doesNotMatch(migration, /\bset (?:title|slug|source_\w+|author_id|roles)\s*=/i);
  const mappings = [...migration.matchAll(/when '([^']+)' then '([^']+)'/g)];
  assert.deepEqual(Object.fromEntries(mappings.map(match => [match[1], match[2]])), KNOWN_TITLES);
});
