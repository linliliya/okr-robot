const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const H = require("../shared.js");

const appSource = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8")
  .replace(/\nroute\(\);\s*$/, "\n");

function harness(options = {}) {
  const nodes = new Map();
  let inputs = [], dialog, failure = options.failure;
  const copied = [], updates = [];
  const skill = {
    id: 7, slug: "pdf-helper", title: "PDF helper", description: "Prepare course handouts",
    recommendation: "Useful for handouts", author_id: "demo-user", source_type: "external",
    source_url: "https://github.com/example/skills/tree/main/skills/pdf",
    source_key: "github:example/skills:skills/pdf", source_path: "skills/pdf",
    category: "效率工具", roles: [H.ROLE_OPTIONS.find(role => role !== "通用")],
  };
  function element(id) {
    if (!nodes.has(id)) {
      const node = {
        id, textContent: "", style: {}, value: "", disabled: false,
        setAttribute() {}, focus() {}, scrollIntoView() {},
        showModal() { this.open = true; },
        close() { this.open = false; },
        remove() { this.removed = true; },
        querySelectorAll(selector) { return selector === 'input[name="roles"]' ? inputs : []; },
      };
      let html = "";
      Object.defineProperty(node, "innerHTML", {
        get() { return html; },
        set(value) {
          html = value;
          if (id !== "dialog") return;
          inputs = [...html.matchAll(/<input\b([^>]*\bname="roles"[^>]*)>/gi)].map(([, attrs]) => ({
            name: "roles", value: attrs.match(/\bvalue="([^"]*)"/)?.[1] || "",
            checked: /\bchecked(?:\s|$)/.test(attrs),
          }));
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
      return {
        update(value) { payload = value; return this; },
        eq(key, value) { filters.push([key, value]); return this; },
        select() { return this; },
        async single() {
          updates.push({ table, payload, filters });
          if (failure) return { data: null, error: failure };
          Object.assign(skill, payload);
          return { data: { id: skill.id }, error: null };
        },
      };
    },
  };
  const context = vm.createContext({
    window: { SKILLHUB_PREVIEW_CLIENT: sb, SkillHub: H, addEventListener() {} },
    document: {
      getElementById: element,
      createElement() { dialog = element("dialog"); return dialog; },
      body: { appendChild() {} },
    },
    location: { hash: "#/" }, URL,
    navigator: { clipboard: { async writeText(value) { copied.push(value); } } },
    setTimeout() { return 0; }, clearTimeout() {},
    FormData: class { getAll(key) { return key === "roles" ? inputs.filter(input => input.checked).map(input => input.value) : []; } },
    skillForTest: skill,
  });
  vm.runInContext(`${appSource}\nme = { id: "demo-user" };`, context);
  return {
    skill, updates, copied, element,
    openEditor() {
      vm.runInContext("bindRoleEditor(skillForTest)", context);
      element("edit-skill-roles").onclick();
    },
    chooseRole(role) {
      const input = inputs.find(input => input.value === role);
      assert.ok(input);
      input.checked = true;
      element("edit-role-fields").onchange({ target: input });
    },
    selectedRoles() { return inputs.filter(input => input.checked).map(input => input.value); },
    submit() { return element("role-edit-form").onsubmit({ preventDefault() {}, target: element("role-edit-form") }); },
    isOpen() { return dialog.open && !dialog.removed; },
    setFailure(error) { failure = error; },
    showDetail() { vm.runInContext("viewExternalDetail(skillForTest)", context); },
  };
}

test("editing roles alone saves only role metadata and preserves the source and content", async () => {
  const app = harness();
  const before = { ...app.skill };
  app.openEditor();
  assert.deepEqual(app.selectedRoles(), before.roles);
  app.chooseRole("通用");
  await app.submit();
  assert.equal(app.updates.length, 1);
  assert.equal(app.updates[0].table, "skills");
  assert.deepEqual(app.updates[0].filters, [["id", before.id]]);
  assert.deepEqual(Object.keys(app.updates[0].payload).sort(), ["roles", "updated_at"]);
  assert.deepEqual(app.skill.roles, ["通用"]);
  for (const key of ["source_url", "source_key", "source_path", "title", "description", "recommendation"]) {
    assert.equal(app.skill[key], before[key]);
  }
  assert.equal(app.isOpen(), false);
});

test("a failed role save keeps the previous data and permits retry without losing the selection", async () => {
  const app = harness({ failure: { code: "42703", message: "roles is missing" } });
  const previousRoles = app.skill.roles;
  app.openEditor();
  app.chooseRole("通用");
  await app.submit();
  assert.deepEqual(app.skill.roles, previousRoles);
  assert.deepEqual(app.selectedRoles(), ["通用"]);
  assert.equal(app.isOpen(), true);
  assert.equal(app.element("save-role-edit").disabled, false);
  assert.match(app.element("role-edit-error").textContent, /还未启用/);
  app.setFailure(null);
  await app.submit();
  assert.deepEqual(app.skill.roles, ["通用"]);
  assert.equal(app.isOpen(), false);
});

test("external installation copies an AI request containing the exact Skill URL", async () => {
  const app = harness();
  app.showDetail();
  await app.element("copy-install").onclick();
  assert.equal(app.copied.length, 1);
  assert.match(app.copied[0], /安装/);
  assert.ok(app.copied[0].includes(app.skill.source_url));
  assert.doesNotMatch(app.copied[0], /npx\s|npm\s/);
  assert.doesNotMatch(app.element("app").innerHTML, /Node\.js|复制安装命令/);
});
