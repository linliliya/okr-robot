const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const H = require("../shared.js");

const appSource = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8")
  .replace(/\nroute\(\);\s*$/, "\n");

function deferred() {
  let resolve, reject;
  const promise = new Promise((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function decode(value) {
  return String(value || "").replace(/&(?:amp|lt|gt|quot|#39);/g,
    entity => ({ "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'" }[entity]));
}

function readme(extra = {}) {
  return {
    markdown: "# Readme\n\nREADME-only contents", path: "skills/pdf/README.md",
    url: "https://github.com/example/skills/blob/main/skills/pdf/README.md",
    kind: "readme", scope: "skill", repo: "example/skills", ref: "main", ...extra,
  };
}

function harness(options = {}) {
  const nodes = new Map(), reads = [], renders = [], writes = [];
  let panels = [];
  const skill = {
    id: 7, slug: "pdf-helper", title: "pdf-helper", display_title: "整理 PDF 课程讲义",
    description: "A concise introduction to the PDF helper.", author_id: "demo-user",
    source_type: "external", source_url: "https://github.com/example/skills/tree/main/skills/pdf",
    source_key: "github:example/skills:skills/pdf", source_path: "skills/pdf",
    category: "效率工具", roles: ["教学教研"], ...options.skill,
  };
  const skills = options.skills || [skill];

  function makePanel(attrs, body) {
    const retry = {};
    const content = {
      innerHTML: body.match(/<div class="readme-content"[^>]*>([\s\S]*)<\/div>/)?.[1] || "",
      querySelector(selector) {
        return selector === "[data-readme-retry]" && this.innerHTML.includes("data-readme-retry") ? retry : null;
      },
    };
    return {
      dataset: { readmeSlug: decode(attrs.match(/data-readme-slug="([^"]*)"/)?.[1]) },
      open: /\bopen(?:\s|$)/.test(attrs), content,
      querySelector(selector) { return selector === ".readme-content" ? content : null; },
    };
  }
  function element(id) {
    if (!nodes.has(id)) {
      let html = "";
      const node = {
        id, style: {}, value: "", textContent: "", disabled: false,
        focus() {}, setAttribute() {}, scrollIntoView() {},
        querySelectorAll(selector) { return id === "app" && selector === "[data-readme-slug]" ? panels : []; },
        contains(node) { return id === "app" && panels.includes(node); },
      };
      Object.defineProperty(node, "innerHTML", {
        get() { return html; },
        set(value) {
          html = value;
          if (id === "app") panels = [...value.matchAll(/<details\b([^>]*data-readme-slug="[^"]*"[^>]*)>([\s\S]*?)<\/details>/g)]
            .map(([, attrs, body]) => makePanel(attrs, body));
        },
      });
      nodes.set(id, node);
    }
    return nodes.get(id);
  }
  const document = { getElementById: element };
  const marked = { parse() { throw new Error("The renderer owns Markdown parsing"); } };
  const DOMPurify = { sanitize() { throw new Error("The renderer owns sanitization"); } };
  const sb = {
    auth: { async getSession() { return { data: { session: { user: { id: "demo-user", email: "colleague@example.test" } } } }; } },
    from(table) {
      return {
        select() { return this; }, eq() { return this; },
        async maybeSingle() { return { data: table === "profiles" ? { display_name: "Colleague" } : skill, error: null }; },
        insert(payload) { writes.push({ table, payload }); throw new Error("README reading must not write to the database"); },
        update(payload) { writes.push({ table, payload }); throw new Error("README reading must not write to the database"); },
        delete() { writes.push({ table }); throw new Error("README reading must not delete data"); },
      };
    },
  };
  const context = vm.createContext({
    window: {
      SKILLHUB_PREVIEW_CLIENT: sb, SkillHub: H, marked, DOMPurify, addEventListener() {},
      SkillHubImporter: {
        async readReadme(target, settings) {
          reads.push({ skill: target, signal: settings.signal });
          return options.readReadme ? options.readReadme(target, settings, reads.length) : readme();
        },
      },
      SkillHubReadme: {
        render(result, settings) {
          renders.push({ result, settings });
          return `<p data-rendered-readme>${result.markdown}</p>`;
        },
      },
    },
    document, location: { hash: "#/" }, URL, AbortController,
    navigator: { clipboard: { async writeText() {} } },
    setTimeout() { return 0; }, clearTimeout() {}, skillForTest: skill, skillsForTest: skills,
  });
  vm.runInContext(`${appSource}\nme = { id: "demo-user" };`, context);
  return {
    skill, skills, reads, renders, writes, element, document, marked, DOMPurify,
    panels() { return panels; },
    showList() { return vm.runInContext("skillCache = skillsForTest; renderList()", context); },
    showDetail() { return vm.runInContext("viewExternalDetail(skillForTest)", context); },
    toggle(panel, open) { panel.open = open; return panel.ontoggle(); },
    retry(panel) { return panel.content.querySelector("[data-readme-retry]").onclick(); },
    filter(query) { element("search").oninput({ target: { value: query } }); },
    navigate() { return vm.runInContext('location.hash = "#/share"; route()', context); },
    replaceWithoutRoute() { element("app").innerHTML = "<p>A replacement view</p>"; },
    async settle() { await new Promise(resolve => setImmediate(resolve)); },
  };
}

test("list README controls are outside card navigation and remain lazy until opened", () => {
  const app = harness();
  app.skills.push({ ...app.skill, id: 8, slug: "internal-helper", source_type: "upload" });
  app.showList();
  assert.match(app.element("app").innerHTML, /<article class="skill-card"><a class="card"/);
  assert.match(app.element("app").innerHTML, /<\/a><details class="readme-section"/);
  assert.equal(app.panels().length, 1, "internal uploads do not receive a remote README control");
  assert.equal(app.panels()[0].open, false);
  assert.equal(app.reads.length, 0);
});

test("opening a README calls the importer and renderer without changing Skill metadata", async () => {
  const app = harness();
  const before = structuredClone(app.skill);
  app.showList();
  const panel = app.panels()[0];
  await app.toggle(panel, true);
  assert.equal(app.reads.length, 1);
  assert.equal(app.reads[0].skill, app.skill);
  assert.equal(app.reads[0].signal.aborted, false);
  assert.equal(app.renders.length, 1);
  assert.equal(app.renders[0].settings.marked, app.marked);
  assert.equal(app.renders[0].settings.DOMPurify, app.DOMPurify);
  assert.equal(app.renders[0].settings.document, app.document);
  assert.match(panel.content.innerHTML, /data-rendered-readme/);
  assert.ok(panel.content.innerHTML.includes(readme().url));
  assert.ok(app.element("app").innerHTML.includes(app.skill.description));
  assert.deepEqual(app.skill, before);
  assert.deepEqual(app.writes, []);
  await app.toggle(panel, false);
  await app.toggle(panel, true);
  assert.equal(app.reads.length, 1, "an already rendered README is reused after reopening");
});

test("external detail opens and reads the README while keeping the short introduction", async () => {
  const app = harness();
  app.showDetail();
  await app.settle();
  assert.equal(app.panels().length, 1);
  assert.equal(app.panels()[0].open, true);
  assert.equal(app.reads.length, 1);
  assert.match(app.panels()[0].content.innerHTML, /README-only contents/);
  assert.match(app.element("app").innerHTML, /<h1>整理 PDF 课程讲义<\/h1>/);
  assert.ok(app.element("app").innerHTML.includes(`<p class="desc">${app.skill.description}</p>`));
  assert.deepEqual(app.writes, []);
});

test("closing an in-flight README aborts it and a late result cannot replace a reopened request", async () => {
  const first = deferred(), second = deferred();
  const app = harness({ readReadme: (_skill, _settings, call) => call === 1 ? first.promise : second.promise });
  app.showList();
  const panel = app.panels()[0];
  const oldRead = app.toggle(panel, true);
  await app.toggle(panel, false);
  assert.equal(app.reads[0].signal.aborted, true);
  const newRead = app.toggle(panel, true);
  assert.equal(app.reads.length, 2);
  first.resolve(readme({ markdown: "STALE CONTENT" }));
  await oldRead;
  assert.equal(app.renders.length, 0);
  assert.doesNotMatch(panel.content.innerHTML, /STALE CONTENT/);
  second.resolve(readme({ markdown: "CURRENT CONTENT" }));
  await newRead;
  assert.equal(app.renders.length, 1);
  assert.match(panel.content.innerHTML, /CURRENT CONTENT/);
});

test("a failed read provides a safe project link and an explicit retry that can succeed", async () => {
  const app = harness({ readReadme: (_skill, _settings, call) => {
    if (call === 1) throw Object.assign(new Error("UNTRUSTED REMOTE ERROR"), { code: "NETWORK_ERROR" });
    return readme();
  } });
  app.showList();
  const panel = app.panels()[0];
  await app.toggle(panel, true);
  assert.match(panel.content.innerHTML, /查看原项目/);
  assert.ok(panel.content.innerHTML.includes(`href="${app.skill.source_url}"`));
  assert.match(panel.content.innerHTML, /target="_blank" rel="noopener noreferrer"/);
  assert.doesNotMatch(panel.content.innerHTML, /UNTRUSTED REMOTE ERROR/);
  await app.retry(panel);
  assert.equal(app.reads.length, 2);
  assert.match(panel.content.innerHTML, /README-only contents/);
});

test("a failed README can also retry by closing and reopening it", async () => {
  const app = harness({ readReadme: (_skill, _settings, call) => {
    if (call === 1) throw Object.assign(new Error("Not found"), { code: "NO_README" });
    return readme();
  } });
  app.showList();
  const panel = app.panels()[0];
  await app.toggle(panel, true);
  await app.toggle(panel, false);
  await app.toggle(panel, true);
  assert.equal(app.reads.length, 2);
  assert.match(panel.content.innerHTML, /README-only contents/);
});

test("non-GitHub sources retain a useful original-project fallback", async () => {
  const app = harness({
    skill: { source_url: "https://example.org/skills/pdf", source_key: "url:https://example.org/skills/pdf", source_path: "" },
    readReadme() { throw Object.assign(new Error("Unsupported"), { code: "UNSUPPORTED_SOURCE" }); },
  });
  app.showDetail();
  await app.settle();
  const content = app.panels()[0].content.innerHTML;
  assert.match(content, /这个来源暂时无法直接读取 README/);
  assert.match(content, /href="https:\/\/example.org\/skills\/pdf"/);
  assert.equal(app.renders.length, 0);
});

for (const replacement of ["filter", "navigation", "detached view"]) {
  test(`${replacement} prevents a previous README request from rendering after the view changes`, async () => {
    const waiting = deferred();
    const app = harness({ readReadme: () => waiting.promise });
    app.showList();
    const oldPanel = app.panels()[0];
    const pending = app.toggle(oldPanel, true);
    if (replacement === "filter") app.filter("no matching skill");
    else if (replacement === "navigation") await app.navigate();
    else app.replaceWithoutRoute();
    if (replacement !== "detached view") assert.equal(app.reads[0].signal.aborted, true);
    waiting.resolve(readme({ markdown: "STALE CONTENT" }));
    await pending;
    assert.equal(app.renders.length, 0);
    assert.doesNotMatch(app.element("app").innerHTML, /STALE CONTENT/);
    assert.doesNotMatch(oldPanel.content.innerHTML, /STALE CONTENT/);
    assert.deepEqual(app.writes, []);
  });
}

for (const [kind, scope, expected] of [
  ["skill", "skill", /以下展示 SKILL\.md/],
  ["readme", "repository", /以下为整个仓库的说明/],
]) {
  test(`README fallback ${kind}/${scope} explains which source document is being displayed`, async () => {
    const app = harness({ readReadme: () => readme({ kind, scope }) });
    app.showList();
    const panel = app.panels()[0];
    await app.toggle(panel, true);
    assert.match(panel.content.innerHTML, expected);
  });
}
