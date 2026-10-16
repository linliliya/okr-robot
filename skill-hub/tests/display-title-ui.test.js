const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const H = require("../shared.js");

const appSource = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8")
  .replace(/\nroute\(\);\s*$/, "\n");

function decode(value) {
  return String(value || "").replace(/&(?:amp|lt|gt|quot|#39);/g,
    entity => ({ "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'" }[entity]));
}

function harness(options = {}) {
  const nodes = new Map(), updates = [], copied = [];
  let dialog, failure = options.failure;
  const skill = {
    id: 7, slug: "pdf-helper", title: "pdf-helper", display_title: "整理 PDF 课程讲义",
    description: "Prepare course handouts", recommendation: "Useful for handouts",
    author_id: "demo-user", source_type: "external",
    source_url: "https://github.com/example/skills/tree/main/skills/pdf",
    source_key: "github:example/skills:skills/pdf", source_path: "skills/pdf",
    category: "效率工具", roles: ["教学教研"],
    ...options.skill,
  };
  function element(id) {
    if (!nodes.has(id)) {
      const node = {
        id, value: "", textContent: "", style: {}, disabled: false,
        setAttribute() {}, focus() {}, scrollIntoView() {},
        querySelectorAll() { return []; },
        showModal() { this.open = true; }, close() { this.open = false; },
        remove() { this.removed = true; },
      };
      let html = "";
      Object.defineProperty(node, "innerHTML", {
        get() { return html; },
        set(value) {
          html = value;
          if (id !== "dialog") return;
          const attrs = html.match(/<input\b([^>]*\bid="edit-display-title"[^>]*)>/)?.[1];
          if (attrs) element("edit-display-title").value = decode(attrs.match(/\bvalue="([^"]*)"/)?.[1]);
        },
      });
      nodes.set(id, node);
    }
    return nodes.get(id);
  }
  const sb = {
    from(table) {
      let payload;
      const filters = [];
      const query = {
        update(value) { payload = value; return this; },
        eq(key, value) { filters.push([key, value]); return this; },
        select() { return this; }, order() { return this; },
        async single() {
          updates.push({ table, payload, filters });
          if (failure) return { data: null, error: failure };
          if (options.noRows) return { data: null, error: null };
          Object.assign(skill, payload);
          return { data: { id: skill.id }, error: null };
        },
        async maybeSingle() { return { data: skill, error: null }; },
        then(resolve, reject) { return Promise.resolve({ data: [], error: null }).then(resolve, reject); },
      };
      return query;
    },
  };
  const context = vm.createContext({
    window: { SKILLHUB_PREVIEW_CLIENT: sb, SkillHub: H, addEventListener() {} },
    document: { getElementById: element, createElement() { dialog = element("dialog"); return dialog; }, body: { appendChild() {} } },
    location: { hash: "#/" }, URL,
    navigator: { clipboard: { async writeText(value) { copied.push(value); } } },
    setTimeout() { return 0; }, clearTimeout() {}, skillForTest: skill,
  });
  const user = Object.hasOwn(options, "user") ? options.user : { id: "demo-user" };
  vm.runInContext(`${appSource}\nme = ${JSON.stringify(user)};`, context);
  return {
    skill, element, updates, copied,
    cardHTML() { return vm.runInContext("cardHTML(skillForTest)", context); },
    headingHTML() { return vm.runInContext("skillHeadingHTML(skillForTest)", context); },
    showDetail() { return vm.runInContext("viewDetail(skillForTest.slug)", context); },
    bindEditor() { vm.runInContext("bindTitleEditor(skillForTest)", context); },
    openEditor() { this.bindEditor(); element("edit-skill-title").onclick(); },
    submit() { return element("title-edit-form").onsubmit({ preventDefault() {} }); },
    isOpen() { return !!dialog?.open && !dialog?.removed; },
    setFailure(error) { failure = error; },
  };
}

for (const source_type of ["upload", "external"]) {
  test(`${source_type} cards and details lead with the function and preserve the original name`, async () => {
    const app = harness({ skill: { source_type } });
    assert.match(app.cardHTML(), /<h3>整理 PDF 课程讲义<\/h3>/);
    assert.match(app.cardHTML(), /class="skill-original-name">pdf-helper<\/div>/);
    await app.showDetail();
    assert.match(app.element("app").innerHTML, /<h1>整理 PDF 课程讲义<\/h1>/);
    assert.match(app.element("app").innerHTML, /Skill 名称<\/span> <code>pdf-helper<\/code>/);
  });
}

test("function titles and original names are escaped in cards and details", async () => {
  const app = harness({ skill: { title: '<img src=x onerror="bad()">', display_title: '<script>bad()</script>' } });
  assert.match(app.cardHTML(), /<h3>&lt;script&gt;bad\(\)&lt;\/script&gt;<\/h3>/);
  assert.doesNotMatch(app.cardHTML(), /<img|<script/);
  await app.showDetail();
  assert.doesNotMatch(app.element("app").innerHTML, /<img|<script/);
});

test("editing the function title updates only display metadata and leaves installation identity intact", async () => {
  const app = harness();
  const before = { ...app.skill }, installBefore = H.buildInstallPrompt(app.skill);
  app.openEditor();
  assert.equal(app.element("edit-display-title").value, before.display_title);
  app.element("edit-display-title").value = "  提取 PDF 中的课程资料  ";
  await app.submit();
  assert.equal(app.updates.length, 1);
  assert.equal(app.updates[0].table, "skills");
  assert.deepEqual(app.updates[0].filters, [["id", before.id]]);
  assert.deepEqual(Object.keys(app.updates[0].payload).sort(), ["display_title", "updated_at"]);
  assert.equal(app.skill.display_title, "提取 PDF 中的课程资料");
  for (const key of ["title", "slug", "source_url", "source_key", "source_path", "description", "recommendation", "roles"]) {
    assert.equal(app.skill[key], before[key]);
  }
  assert.equal(H.buildInstallPrompt(app.skill), installBefore);
  await app.showDetail();
  await app.element("copy-install").onclick();
  assert.equal(app.copied[0], installBefore);
  assert.equal(app.isOpen(), false);
});

test("a failed title save retains the user's input and allows retry", async () => {
  const app = harness({ failure: { code: "42703", message: "display_title is missing" } });
  app.openEditor();
  app.element("edit-display-title").value = "整理课程讲义";
  await app.submit();
  assert.equal(app.skill.display_title, "整理 PDF 课程讲义");
  assert.equal(app.element("edit-display-title").value, "整理课程讲义");
  assert.equal(app.isOpen(), true);
  assert.equal(app.element("save-title-edit").disabled, false);
  assert.equal(app.element("cancel-title-edit").disabled, false);
  assert.match(app.element("title-edit-error").textContent, /还未启用/);
  app.setFailure(null);
  await app.submit();
  assert.equal(app.skill.display_title, "整理课程讲义");
  assert.equal(app.isOpen(), false);
});

for (const [label, value] of [["empty", "   "], ["over 80 characters", "课".repeat(81)], ["multiple lines", "整理\n课程讲义"]]) {
  test(`a ${label} function title cannot reach the database`, async () => {
    const app = harness();
    app.openEditor();
    app.element("edit-display-title").value = value;
    await app.submit();
    assert.equal(app.updates.length, 0);
    assert.equal(app.isOpen(), true);
    assert.match(app.element("title-edit-error").textContent, /功能标题/);
  });
}

test("an 80-character function title is accepted", async () => {
  const app = harness();
  app.openEditor();
  app.element("edit-display-title").value = "课".repeat(80);
  await app.submit();
  assert.equal(app.skill.display_title, "课".repeat(80));
  assert.equal(app.isOpen(), false);
});

test("canceling the title editor writes nothing", () => {
  const app = harness();
  app.openEditor();
  app.element("edit-display-title").value = "不保存的标题";
  app.element("cancel-title-edit").onclick();
  assert.equal(app.updates.length, 0);
  assert.equal(app.isOpen(), false);
  assert.equal(app.skill.display_title, "整理 PDF 课程讲义");
});

test("a save that updates no row reports failure and preserves the dialog", async () => {
  const app = harness({ noRows: true });
  app.openEditor();
  app.element("edit-display-title").value = "整理课程讲义";
  await app.submit();
  assert.equal(app.isOpen(), true);
  assert.match(app.element("title-edit-error").textContent, /权限|没有保存成功/);
});

for (const [source_type, user, editable] of [
  ["upload", { id: "colleague" }, true],
  ["external", { id: "demo-user" }, true],
  ["external", { id: "colleague" }, false],
  ["upload", null, false],
  ["external", null, false],
]) {
  test(`title editing for ${source_type} is ${editable ? "enabled" : "disabled"} for ${user?.id || "a visitor"}`, () => {
    const app = harness({ skill: { source_type }, user });
    assert.equal(app.headingHTML().includes('id="edit-skill-title"'), editable);
    app.bindEditor();
    assert.equal(typeof app.element("edit-skill-title").onclick === "function", editable);
  });
}
