const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { randomUUID } = require("node:crypto");

// Execute the application's real recommendation handlers. Network imports and
// Supabase are mocked at their boundaries so race tests remain deterministic.
const appSource = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8")
  .replace(/\nroute\(\);\s*$/, "\n");
const H = require("../shared.js");

function deferred() {
  let resolve, reject;
  const promise = new Promise((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function decode(value) {
  return String(value || "").replace(/&(?:amp|lt|gt|quot|#39);/g,
    entity => ({ "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'" }[entity]));
}

async function harness(options = {}) {
  const nodes = new Map();
  let roleInputs = [];
  const renderedIds = new Set();
  const timers = new Map();
  let timerId = 0;
  function element(id) {
    if (!nodes.has(id)) {
      const classes = new Set();
      const node = {
        id, name: "", tagName: "DIV", value: "", textContent: "", style: {}, dataset: {},
        disabled: false, hidden: false, open: false, required: false, readOnly: false,
        classList: {
          add(name) { classes.add(name); }, remove(name) { classes.delete(name); },
          contains(name) { return classes.has(name); },
          toggle(name, force) {
            const add = force ?? !classes.has(name);
            if (add) classes.add(name); else classes.delete(name);
            return add;
          },
        },
        setAttribute(name, value) { this[name] = value; },
        removeAttribute(name) { delete this[name]; },
        addEventListener(type, handler) { this[`on${type}`] = handler; },
        focus() {}, scrollIntoView() {},
        querySelectorAll(selector) {
          if (selector === 'input[name="roles"]') return roleInputs;
          if (this.id !== "recommend-form") return [];
          const tags = selector.split(",").map(tag => tag.trim().toUpperCase());
          return [...nodes.values(), ...roleInputs].filter(child => tags.includes(child.tagName));
        },
      };
      let html = "";
      Object.defineProperty(node, "innerHTML", {
        get() { return html; },
        set(value) {
          html = value;
          if (id === "app") {
            // Browser innerHTML replaces child nodes, including same-ID forms.
            for (const renderedId of renderedIds) nodes.delete(renderedId);
            renderedIds.clear();
            parseMarkup(value);
          }
          if (node.tagName === "SELECT") {
            const options = [...value.matchAll(/<option\b([^>]*)>([\s\S]*?)<\/option>/g)];
            const chosen = options.find(option => /\bselected\b/.test(option[1])) || options[0];
            if (chosen) node.value = decode(chosen[1].match(/\bvalue="([^"]*)"/)?.[1] ?? chosen[2]);
          }
        },
      });
      nodes.set(id, node);
    }
    return nodes.get(id);
  }
  function parseMarkup(html) {
    roleInputs = [...html.matchAll(/<input\b([^>]*\bname="roles"[^>]*)>/gi)].map(([, attrs]) => ({
      tagName: "INPUT", name: "roles", value: decode(attrs.match(/\bvalue="([^"]*)"/)?.[1]),
      checked: /\bchecked(?:\s|$)/.test(attrs), disabled: false,
    }));
    for (const match of html.matchAll(/<([a-z][a-z0-9-]*)\b([^>]*\bid="([^"]+)"[^>]*)>/gi)) {
      const [, tag, attrs, id] = match;
      const node = element(id);
      renderedIds.add(id);
      node.tagName = tag.toUpperCase();
      node.name = attrs.match(/\bname="([^"]*)"/)?.[1] || "";
      node.value = decode(attrs.match(/\bvalue="([^"]*)"/)?.[1]);
      node.hidden = /\bhidden(?:\s|$)/.test(attrs);
      node.open = /\bopen(?:\s|$)/.test(attrs);
      node.disabled = /\bdisabled(?:\s|$)/.test(attrs);
      node.required = /\brequired(?:\s|$)/.test(attrs);
      node.readOnly = /\breadonly(?:\s|$)/.test(attrs);
      if (["SELECT", "TEXTAREA"].includes(node.tagName)) {
        const content = html.slice(match.index + match[0].length).split(`</${tag}>`)[0];
        if (node.tagName === "TEXTAREA") node.value = decode(content);
        else node.innerHTML = content;
      }
    }
  }

  const db = { skills: (options.existing || []).map(skill => ({ ...skill })) };
  const calls = [];
  let nextId = 100;
  const sb = {
    from(table) {
      let operation = "select", payload, single = false;
      const filters = [];
      const comparisons = [];
      return {
        select() { return this; },
        eq(key, value) {
          comparisons.push({ key, value });
          filters.push(row => String(row[key]) === String(value));
          return this;
        },
        in(key, values) {
          comparisons.push({ key, values });
          filters.push(row => values.includes(row[key]));
          return this;
        },
        limit() { return this; },
        single() { single = true; return this; },
        maybeSingle() { single = true; return this; },
        insert(value) { operation = "insert"; payload = value; return this; },
        update(value) { operation = "update"; payload = value; return this; },
        then(resolve, reject) {
          return Promise.resolve().then(() => {
            calls.push({ table, operation, payload, comparisons });
            let rows = db[table].filter(row => filters.every(match => match(row)));
            if (operation === "insert") {
              rows = [{ id: nextId++, recommendation: "", ...payload }];
              db[table].push(...rows);
            } else if (operation === "update") rows.forEach(row => Object.assign(row, payload));
            return { data: single ? rows.at(-1) || null : rows, error: null };
          }).then(resolve, reject);
        },
      };
    },
  };
  const importerCalls = [];
  const importer = {
    async inspect(...args) {
      importerCalls.push({ method: "inspect", args });
      if (!options.inspect) throw new Error("No import response configured");
      return options.inspect(...args);
    },
    async readChoice(...args) {
      importerCalls.push({ method: "readChoice", args });
      if (!options.readChoice) throw new Error("No choice response configured");
      return options.readChoice(...args);
    },
  };
  const location = { hash: "#/recommend" };
  const context = vm.createContext({
    window: {
      SKILLHUB_PREVIEW_CLIENT: sb, SkillHub: H, SkillHubImporter: importer,
      addEventListener() {},
    },
    document: { getElementById: element },
    location, navigator: {}, crypto: { randomUUID }, AbortController, URL,
    setTimeout(callback) { timers.set(++timerId, callback); return timerId; },
    clearTimeout(id) { timers.delete(id); },
    FormData: class {
      get(key) { return [...nodes.values()].find(node => node.name === key)?.value || ""; }
      getAll(key) { return key === "roles" ? roleInputs.filter(input => input.checked && !input.disabled).map(input => input.value) : []; }
    },
  });
  vm.runInContext(`${appSource}\nme = { id: "demo-user" };`, context, { filename: "app.js" });
  await vm.runInContext(`viewRecommend(${JSON.stringify(options.slug || "")})`, context);
  return {
    db, calls, importerCalls, element, location,
    selectedRoles() { return roleInputs.filter(input => input.checked).map(input => input.value); },
    chooseRole(role, checked = true) {
      const input = roleInputs.find(input => input.value === role);
      assert.ok(input, `The form includes the ${role} role`);
      input.checked = checked;
      element("skill-role-fields").onchange({ target: input });
    },
    input(id, value) {
      const node = element(id);
      node.value = value;
      return node.oninput?.({ target: node });
    },
    read() { return element("import-button").onclick({ preventDefault() {} }); },
    submit() { return element("recommend-form").onsubmit({ preventDefault() {}, target: element("recommend-form") }); },
    openRecommendation(slug = "") { return vm.runInContext(`viewRecommend(${JSON.stringify(slug)})`, context); },
    async runTimers() {
      const pending = [...timers.values()];
      timers.clear();
      for (const callback of pending) await callback();
    },
  };
}

function imported(name = "pdf", extra = {}) {
  return {
    type: "skill",
    skill: {
      source: H.normalizeSource(`https://github.com/example/skills/tree/main/skills/${name}`),
      title: `${name} assistant`, display_title: "处理工作文档", description: `Work with ${name} files`,
      original_author: "example", category: "效率工具", tags: `${name},documents`,
      ...extra,
    },
  };
}

test("an imported Skill can publish with its function title and introduction without a recommendation reason", async () => {
  const result = imported();
  const app = await harness({ inspect: async () => result });
  app.input("r-url", result.skill.source.url);
  await app.read();
  assert.equal(app.element("r-title").value, result.skill.title);
  assert.equal(app.element("r-description").value, result.skill.description);
  assert.equal(app.element("r-author").value, result.skill.original_author);
  assert.equal(app.element("r-category").value, result.skill.category);
  assert.equal(app.element("r-tags").value, result.skill.tags);
  assert.doesNotMatch(app.element("app").innerHTML, /id="r-recommendation"|name="recommendation"|为什么推荐给同事/);
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].title, result.skill.title);
  assert.equal(app.db.skills[0].display_title, result.skill.display_title);
  assert.equal(app.db.skills[0].source_key, result.skill.source.key);
  assert.equal(app.db.skills[0].recommendation, "");
  assert.equal(Object.hasOwn(app.calls.find(call => call.operation === "insert").payload, "recommendation"), false);
});

test("imported metadata still needs an introduction before publishing", async () => {
  const result = imported("pdf", { description: "" });
  const app = await harness({ inspect: async () => result });
  app.input("r-url", result.skill.source.url);
  await app.read();
  await app.submit();
  assert.equal(app.db.skills.length, 0);
  assert.match(app.element("recommend-error").textContent, /介绍/);
  assert.equal(app.element("recommend-submit").disabled, false);
  app.input("r-display-title", "整理 PDF 课程讲义");
  app.input("r-description", "Prepare course handouts from PDF files.");
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].description, "Prepare course handouts from PDF files.");
});

test("a source read failure exposes manual entry and permits a complete recommendation", async () => {
  const app = await harness({ inspect: async () => { throw new Error("GitHub is temporarily unavailable"); } });
  app.input("r-url", "https://example.org/my-skill");
  await app.read();
  assert.equal(app.element("metadata-fields").open, true);
  assert.equal(app.element("recommend-submit").disabled, false);
  app.input("r-title", "Manually entered skill");
  app.input("r-display-title", "整理课程资料");
  app.input("r-description", "A manually described use case");
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].title, "Manually entered skill");
  assert.equal(app.db.skills[0].source_url, "https://example.org/my-skill");
});

test("changing the source aborts its old read and ignores a late response", async () => {
  const slow = deferred();
  const stale = imported("old");
  const current = imported("new");
  const app = await harness({ inspect: url => url === stale.skill.source.url ? slow.promise : Promise.resolve(current) });
  app.input("r-url", stale.skill.source.url);
  const first = app.read();
  await Promise.resolve();
  const oldSignal = app.importerCalls[0].args[1].signal;
  app.input("r-url", current.skill.source.url);
  assert.equal(oldSignal.aborted, true);
  await app.read();
  slow.resolve(stale);
  await first;
  assert.equal(app.element("r-title").value, current.skill.title);
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].source_key, current.skill.source.key);
});

test("submitting during metadata import cannot publish an incomplete recommendation", async () => {
  const waiting = deferred();
  const result = imported();
  const app = await harness({ inspect: () => waiting.promise });
  app.input("r-url", result.skill.source.url);
  const pending = app.read();
  await app.submit();
  assert.equal(app.db.skills.length, 0);
  waiting.resolve(result);
  await pending;
  await app.submit();
  assert.equal(app.db.skills.length, 1);
});

test("an imported introduction can be edited before publishing without adding a separate reason", async () => {
  const result = imported();
  const app = await harness({ inspect: async () => result });
  app.input("r-url", result.skill.source.url);
  await app.read();
  app.input("r-display-title", "整理 PDF 课程讲义");
  app.input("r-description", "Prepare printable course handouts from PDF files.");
  await app.submit();
  assert.equal(app.db.skills[0].description, "Prepare printable course handouts from PDF files.");
  assert.equal(app.db.skills[0].recommendation, "");
});

for (const duplicateKey of ["skills:example/skills:pdf", "github:example/skills:skills/pdf"]) {
  test(`a skills.sh import opens an existing recommendation matching ${duplicateKey}`, async () => {
    const skillsUrl = "https://skills.sh/example/skills/pdf";
    const result = imported("pdf", { aliasKeys: ["skills:example/skills:pdf"] });
    const app = await harness({
      inspect: async () => result,
      existing: [{ id: 7, slug: "existing-pdf", source_type: "external", source_key: duplicateKey }],
    });
    app.input("r-url", skillsUrl);
    await app.read();
    await app.submit();
    assert.equal(app.db.skills.length, 1);
    assert.equal(app.calls.some(call => call.operation === "insert"), false);
    assert.equal(app.location.hash, "#/skill/existing-pdf");
  });
}

test("blurring a source starts automatic metadata import", async () => {
  const result = imported();
  const app = await harness({ inspect: async () => result });
  app.input("r-url", result.skill.source.url);
  await app.element("r-url").onblur();
  assert.equal(app.importerCalls.length, 1);
  assert.equal(app.element("r-title").value, result.skill.title);
});

test("pasting a link imports after the browser's following input event", async () => {
  const result = imported();
  const app = await harness({ inspect: async () => result });
  app.element("r-url").onpaste();
  app.input("r-url", result.skill.source.url);
  await app.runTimers();
  // The event timer need not return the read promise to the browser.
  await Promise.resolve();
  await Promise.resolve();
  assert.equal(app.importerCalls.length, 1);
  assert.equal(app.element("r-title").value, result.skill.title);
});

test("a repository choice must be selected and read before publishing", async () => {
  const result = imported("slides");
  const choices = [
    { label: "PDF", path: "skills/pdf", url: "https://github.com/example/skills/tree/main/skills/pdf" },
    { label: "Slides", path: "skills/slides", url: result.skill.source.url },
  ];
  const app = await harness({
    inspect: async () => ({ type: "choices", choices }),
    readChoice: async choice => { assert.equal(choice.path, "skills/slides"); return result; },
  });
  app.input("r-url", "https://github.com/example/skills");
  await app.read();
  assert.equal(app.element("skill-choice-field").hidden, false);
  await app.submit();
  assert.equal(app.db.skills.length, 0);
  app.element("skill-choice").value = "1";
  await app.element("skill-choice").onchange();
  assert.equal(app.element("r-title").value, result.skill.title);
  assert.equal(app.element("r-path").value, "skills/slides");
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].source_key, result.skill.source.key);
});

test("replacing an imported link clears stale metadata before it can be published", async () => {
  const result = imported();
  const app = await harness({ inspect: async () => result });
  app.input("r-url", result.skill.source.url);
  await app.read();
  app.input("r-url", "https://example.org/another-skill");
  assert.equal(app.element("r-title").value, "");
  assert.equal(app.element("r-description").value, "");
  assert.equal(app.element("r-path").value, "");
  assert.equal(app.element("imported-preview").hidden, true);
  assert.equal(app.element("r-display-title").value, "");
  await app.submit();
  assert.equal(app.db.skills.length, 0);
});

test("editing an existing recommendation never imports or changes its saved source", async () => {
  const result = imported();
  const source = result.skill.source;
  const existing = {
    id: 7, slug: "existing-pdf", author_id: "demo-user", source_type: "external",
    source_url: source.url, source_key: source.key, source_repo: source.repo, source_path: source.path,
    title: "Original title", display_title: "处理工作文档", description: "Original description", recommendation: "Original reason",
    category: "其他", tags: "", original_author: "example", tested: false,
  };
  const app = await harness({ existing: [existing], slug: existing.slug });
  assert.equal(app.element("r-url").readOnly, true);
  assert.equal(app.element("r-path").readOnly, true);
  assert.equal(app.element("r-url").onblur, undefined);
  assert.equal(app.element("r-url").onpaste, undefined);
  assert.equal(app.element("import-button").onclick, undefined);
  app.input("r-title", "Revised title");
  app.input("r-description", "Revised description");
  // Even a script changing readonly DOM fields cannot rewrite source identity.
  app.input("r-url", "https://github.com/another/skills/tree/main/other");
  app.input("r-path", "other");
  await app.submit();
  assert.equal(app.importerCalls.length, 0);
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].title, "Revised title");
  assert.equal(app.db.skills[0].recommendation, existing.recommendation);
  for (const key of ["source_url", "source_key", "source_repo", "source_path"]) {
    assert.equal(app.db.skills[0][key], existing[key]);
  }
  const update = app.calls.find(call => call.operation === "update");
  assert.ok(update);
  assert.equal(Object.hasOwn(update.payload, "recommendation"), false);
  assert.equal(Object.keys(update.payload).some(key => key.startsWith("source_")), false);
  assert.equal(app.location.hash, "#/skill/existing-pdf");
});

test("a late import after navigation cannot update the replacement recommendation form", async () => {
  const waiting = deferred();
  const result = imported();
  const app = await harness({ inspect: () => waiting.promise });
  app.input("r-url", result.skill.source.url);
  const oldForm = app.element("recommend-form");
  const oldTitle = app.element("r-title");
  const pending = app.read();
  await app.openRecommendation();
  assert.notEqual(app.element("recommend-form"), oldForm);
  app.input("r-url", "https://example.org/new-recommendation");
  app.input("r-title", "A different skill");
  app.input("r-display-title", "整理课程资料");
  app.input("r-description", "Current page description");
  waiting.resolve(result);
  await pending;
  assert.equal(oldTitle.value, "", "the disconnected form is not updated either");
  assert.equal(app.element("r-url").value, "https://example.org/new-recommendation");
  assert.equal(app.element("r-title").value, "A different skill");
  assert.equal(app.element("r-description").value, "Current page description");
  assert.equal(app.element("import-status").textContent, "");
  assert.equal(app.element("recommend-submit").disabled, false);
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].source_key, "url:https://example.org/new-recommendation");
});

test("selected job roles survive automatic import and are saved with the recommendation", async () => {
  const result = imported();
  const app = await harness({ inspect: async () => result });
  const roles = H.ROLE_OPTIONS.filter(role => role !== "通用").slice(0, 2);
  roles.forEach(role => app.chooseRole(role));
  app.input("r-url", result.skill.source.url);
  await app.read();
  assert.deepEqual(app.selectedRoles(), roles);
  await app.submit();
  assert.deepEqual(app.db.skills[0].roles, roles);
});

test("editing an introduction retains previously selected roles and any saved legacy reason", async () => {
  const result = imported();
  const roles = H.ROLE_OPTIONS.filter(role => role !== "通用").slice(0, 2);
  const existing = {
    id: 7, slug: "existing-pdf", author_id: "demo-user", source_type: "external",
    source_url: result.skill.source.url, source_key: result.skill.source.key,
    title: "PDF helper", display_title: "整理 PDF 讲义", description: "Prepare PDF handouts", recommendation: "Useful for handouts",
    category: "其他", tags: "", roles,
  };
  const app = await harness({ existing: [existing], slug: existing.slug });
  assert.deepEqual(app.selectedRoles(), roles);
  app.input("r-description", "Prepare fresh PDF handouts every week.");
  await app.submit();
  assert.deepEqual(app.db.skills[0].roles, roles);
  assert.equal(app.db.skills[0].description, "Prepare fresh PDF handouts every week.");
  assert.equal(app.db.skills[0].recommendation, existing.recommendation);
});

test("general and individual roles are mutually exclusive while roles remain optional", async () => {
  const app = await harness();
  const roles = H.ROLE_OPTIONS.filter(role => role !== "通用").slice(0, 2);
  roles.forEach(role => app.chooseRole(role));
  assert.deepEqual(app.selectedRoles(), roles);
  app.chooseRole("通用");
  assert.deepEqual(app.selectedRoles(), ["通用"]);
  app.chooseRole(roles[0]);
  assert.deepEqual(app.selectedRoles(), [roles[0]]);
  app.chooseRole(roles[0], false);
  assert.deepEqual(app.selectedRoles(), []);
});

test("an imported Skill needs a function title before sharing and keeps its original name", async () => {
  const result = imported("pdf", { display_title: "" });
  const app = await harness({ inspect: async () => result });
  app.input("r-url", result.skill.source.url);
  await app.read();
  assert.equal(app.element("r-display-title").value, "");
  await app.submit();
  assert.equal(app.db.skills.length, 0);
  assert.match(app.element("recommend-error").textContent, /功能标题/);
  app.input("r-display-title", "提取 PDF 中的课程资料");
  await app.submit();
  assert.equal(app.db.skills[0].display_title, "提取 PDF 中的课程资料");
  assert.equal(app.db.skills[0].title, result.skill.title);
  assert.equal(app.db.skills[0].source_url, result.skill.source.url);
  assert.equal(H.buildInstallPrompt(app.db.skills[0]), H.buildInstallPrompt({ ...app.db.skills[0], display_title: "另一个功能标题" }));
});

test("an automatic source read preserves a function title the sharer already entered", async () => {
  const result = imported();
  const app = await harness({ inspect: async () => result });
  app.input("r-url", result.skill.source.url);
  app.input("r-display-title", "整理每周课程讲义");
  await app.read();
  assert.equal(app.element("r-display-title").value, "整理每周课程讲义");
  assert.equal(app.element("r-title").value, result.skill.title);
  await app.submit();
  assert.equal(app.db.skills[0].display_title, "整理每周课程讲义");
});

test("editing a recommendation's function title preserves its original name and source", async () => {
  const source = imported().skill.source;
  const existing = {
    id: 7, slug: "existing-pdf", author_id: "demo-user", source_type: "external",
    source_url: source.url, source_key: source.key, source_repo: source.repo, source_path: source.path,
    title: "pdf-helper", display_title: "处理 PDF", description: "Prepare PDF handouts",
    recommendation: "Useful for handouts", category: "其他", tags: "",
  };
  const app = await harness({ existing: [existing], slug: existing.slug });
  assert.equal(app.element("r-display-title").value, existing.display_title);
  app.input("r-display-title", "整理 PDF 课程讲义");
  await app.submit();
  assert.equal(app.db.skills[0].display_title, "整理 PDF 课程讲义");
  assert.equal(app.db.skills[0].title, existing.title);
  assert.equal(app.db.skills[0].source_url, existing.source_url);
  assert.equal(app.db.skills[0].source_key, existing.source_key);
});
