const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { randomUUID } = require("node:crypto");
const H = require("../shared.js");

// Run the actual application handlers. Only the automatic initial route is
// suppressed; DOM, ZIP decoding, and Supabase are boundary mocks.
const appSource = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8")
  .replace(/\nroute\(\);\s*$/, "\n");

function deferred() {
  let resolve;
  const promise = new Promise(done => { resolve = done; });
  return { promise, resolve };
}

function file(name, markdown, extra = {}) {
  return { name, size: 128, markdown, ...extra };
}

function decode(value) {
  return String(value || "").replace(/&(?:amp|lt|gt|quot|#39);/g,
    entity => ({ "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'" }[entity]));
}

async function harness(options = {}) {
  const nodes = new Map();
  let roleInputs = [];
  function element(id) {
    if (!nodes.has(id)) {
      const node = {
      value: "", innerHTML: "", textContent: "", style: {}, files: [],
      disabled: false, hidden: false,
      querySelectorAll: selector => selector === 'input[name="roles"]' ? roleInputs : [],
      classList: { add() {}, remove() {} },
      focus() {}, click() {},
      };
      if (id === "app") {
        let html = "";
        Object.defineProperty(node, "innerHTML", {
          get() { return html; },
          set(value) {
            html = value;
            roleInputs = [...html.matchAll(/<input\b([^>]*\bname="roles"[^>]*)>/gi)].map(([, attrs]) => ({
              name: "roles", value: attrs.match(/\bvalue="([^"]*)"/)?.[1] || "",
              checked: /\bchecked(?:\s|$)/.test(attrs), disabled: false,
            }));
            for (const match of html.matchAll(/<(input|textarea)\b([^>]*\bid="([^"]+)"[^>]*)>/gi)) {
              const [, tag, attrs, fieldId] = match;
              const content = tag.toLowerCase() === "textarea"
                ? html.slice(match.index + match[0].length).split("</textarea>")[0]
                : attrs.match(/\bvalue="([^"]*)"/)?.[1];
              element(fieldId).value = decode(content);
            }
          },
        });
      }
      nodes.set(id, node);
    }
    return nodes.get(id);
  }
  const original = options.existing ? {
    id: 7, slug: "existing-skill", title: "Existing title", display_title: "生成课程提纲", description: "Existing description",
    author_id: "original-author", source_type: "upload", tags: "", category: "其他",
    roles: options.existingRoles || [],
  } : null;
  const db = {
    skills: original ? [{ ...original }] : [],
    versions: original ? [{ id: 8, skill_id: original.id, version: 1, file_path: "old.zip" }] : [],
  };
  const files = new Map(original ? [["old.zip", file("old.zip", "OLD")]] : []);
  const calls = [];
  let nextId = 100;
  const sb = {
    from(table) {
      let operation = "select", payload, single = false, returning = true;
      const filters = [];
      const query = {
        select() { returning = true; return this; },
        eq(key, value) { filters.push(row => String(row[key]) === String(value)); return this; },
        order() { return this; }, limit() { return this; },
        single() { single = true; return this; },
        maybeSingle() { single = true; return this; },
        insert(value) { operation = "insert"; payload = value; returning = false; return this; },
        update(value) { operation = "update"; payload = value; returning = false; return this; },
        delete() { operation = "delete"; returning = false; return this; },
        then(resolve, reject) {
          return Promise.resolve().then(() => {
            calls.push({ table, operation, payload });
            if (table === "versions" && operation === "insert" && options.failVersion) {
              return { data: null, error: { code: "23505", message: "Simulated version insert failure" } };
            }
            let rows = db[table].filter(row => filters.every(match => match(row)));
            if (operation === "insert") {
              rows = [{ id: nextId++, ...payload }];
              db[table].push(...rows);
            } else if (operation === "update") rows.forEach(row => Object.assign(row, payload));
            else if (operation === "delete") db[table] = db[table].filter(row => !rows.includes(row));
            return { data: returning ? single ? rows.at(-1) || null : rows : null, error: null };
          }).then(resolve, reject);
        },
      };
      return query;
    },
    storage: {
      from() {
        return {
          async upload(filePath, picked) {
            calls.push({ table: "storage", operation: "upload", path: filePath, file: picked });
            if (options.uploadWait) await options.uploadWait.promise;
            files.set(filePath, picked);
            return { data: { path: filePath }, error: null };
          },
          async remove(paths) {
            calls.push({ table: "storage", operation: "remove", paths });
            paths.forEach(filePath => files.delete(filePath));
            return { data: [], error: null };
          },
        };
      },
    },
  };
  const fields = { title: "New title", display_title: "生成课程提纲", description: "New description", category: "其他", tags: "", changelog: "New release", ...options.fields };
  const context = vm.createContext({
    window: { SKILLHUB_PREVIEW_CLIENT: sb, SkillHub: require("../shared.js"), addEventListener() {} },
    document: { getElementById: element },
    location: { hash: "#/upload" }, navigator: {}, crypto: { randomUUID },
    setTimeout: () => 0, clearTimeout() {},
    FormData: class {
      get(key) {
        const fieldId = { title: "f-title", display_title: "f-display-title", description: "f-desc" }[key];
        return options.useDOMFields && fieldId ? element(fieldId).value : fields[key] || "";
      }
      getAll(key) { return key === "roles" ? roleInputs.filter(input => input.checked && !input.disabled).map(input => input.value) : []; }
    },
    JSZip: {
      async loadAsync(picked) {
        if (picked.invalid) throw new Error("Invalid ZIP");
        return { files: {
          "example/SKILL.md": { dir: false, async: async () => picked.read ? picked.read() : picked.markdown },
        } };
      },
    },
  });
  vm.runInContext(`${appSource}\nme = { id: "demo-user" };`, context, { filename: "app.js" });
  await vm.runInContext(`viewUpload(${JSON.stringify(original?.slug || "")})`, context);
  return {
    db, files, calls, original, element, fields,
    input(id, value) {
      const node = element(id);
      node.value = value;
      return node.oninput?.({ target: node });
    },
    selectedRoles() { return roleInputs.filter(input => input.checked).map(input => input.value); },
    chooseRole(role, checked = true) {
      const input = roleInputs.find(input => input.value === role);
      assert.ok(input, `The form includes the ${role} role`);
      input.checked = checked;
      element("skill-role-fields").onchange({ target: input });
    },
    choose(picked) {
      element("file").files = [picked];
      return element("file").onchange();
    },
    submit() { return element("up-form").onsubmit({ preventDefault() {}, target: element("up-form") }); },
  };
}

test("a broken replacement ZIP cannot upload the previous file's documentation", async () => {
  const app = await harness();
  await app.choose(file("valid.zip", "PREVIOUS FILE DOCUMENT"));
  await app.choose(file("broken.zip", "", { invalid: true }));
  await app.submit();
  assert.equal(app.calls.filter(call => call.operation === "upload").length, 0);
  assert.equal(app.db.skills.length, 0);
  assert.equal(app.db.versions.length, 0);
});

test("submitting while ZIP documentation is being read cannot publish stale content", async () => {
  const read = deferred();
  const app = await harness();
  await app.choose(file("previous.zip", "PREVIOUS DOCUMENT"));
  const pending = app.choose(file("current.zip", "", { read: () => read.promise }));
  await app.submit();
  assert.equal(app.calls.filter(call => call.operation === "upload").length, 0);
  read.resolve("CURRENT DOCUMENT");
  await pending;
  await app.submit();
  assert.equal(app.db.versions.length, 1);
  assert.equal(app.db.versions[0].skill_md, "CURRENT DOCUMENT");
});

test("a slow earlier ZIP read cannot overwrite the currently selected ZIP", async () => {
  const read = deferred();
  const app = await harness();
  const pending = app.choose(file("slow.zip", "", { read: () => read.promise }));
  // Let the first load reach its asynchronous SKILL.md read before replacing it.
  await Promise.resolve();
  await Promise.resolve();
  await app.choose(file("current.zip", "CURRENT DOCUMENT"));
  read.resolve("STALE DOCUMENT");
  await pending;
  await app.submit();
  assert.equal(app.db.versions.length, 1);
  assert.equal(app.db.versions[0].skill_md, "CURRENT DOCUMENT");
  assert.equal(app.calls.find(call => call.operation === "upload").file.name, "current.zip");
});

test("a failed first version removes the new empty Skill and its uploaded ZIP", async () => {
  const app = await harness({ failVersion: true });
  await app.choose(file("new.zip", "NEW DOCUMENT"));
  await app.submit();
  assert.ok(app.calls.some(call => call.table === "versions" && call.operation === "insert"));
  assert.equal(app.db.skills.length, 0, "no empty Skill remains after failure");
  assert.equal(app.db.versions.length, 0);
  assert.equal(app.files.size, 0, "the failed version's ZIP is removed");
  assert.equal(app.element("up-btn").disabled, false, "the form permits retry");
});

test("a failed new version keeps the existing Skill, metadata, and older file", async () => {
  const app = await harness({ existing: true, failVersion: true });
  await app.choose(file("new.zip", "NEW DOCUMENT"));
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.skills[0].id, app.original.id);
  assert.equal(app.db.skills[0].title, app.original.title);
  assert.equal(app.db.skills[0].display_title, app.original.display_title);
  assert.equal(app.db.skills[0].description, app.original.description);
  assert.equal(app.db.versions.length, 1);
  assert.equal(app.files.size, 1);
  assert.ok(app.files.has("old.zip"));
  assert.equal(app.calls.some(call => call.table === "skills" && call.operation === "delete"), false);
});

test("an in-flight publish cannot start a second upload", async () => {
  const uploadWait = deferred();
  const app = await harness({ uploadWait });
  await app.choose(file("new.zip", "NEW DOCUMENT"));
  const first = app.submit();
  const second = app.submit();
  await Promise.resolve();
  const uploadsBeforeRelease = app.calls.filter(call => call.operation === "upload").length;
  uploadWait.resolve();
  await Promise.all([first, second]);
  assert.equal(uploadsBeforeRelease, 1);
  assert.equal(app.db.skills.length, 1);
  assert.equal(app.db.versions.length, 1);
});

test("uploading a new Skill saves the selected job roles", async () => {
  const app = await harness();
  const roles = H.ROLE_OPTIONS.filter(role => role !== "通用").slice(0, 2);
  roles.forEach(role => app.chooseRole(role));
  await app.choose(file("new.zip", "NEW DOCUMENT"));
  await app.submit();
  assert.equal(app.db.skills.length, 1);
  assert.deepEqual(app.db.skills[0].roles, roles);
});

test("uploading a new version retains the existing Skill's roles", async () => {
  const roles = H.ROLE_OPTIONS.filter(role => role !== "通用").slice(0, 2);
  const app = await harness({ existing: true, existingRoles: roles });
  assert.deepEqual(app.selectedRoles(), roles);
  await app.choose(file("new.zip", "NEW DOCUMENT"));
  await app.submit();
  assert.equal(app.db.versions.length, 2);
  assert.deepEqual(app.db.skills[0].roles, roles);
});

test("uploading preserves the original name separately from the function title", async () => {
  const app = await harness({ fields: { title: "de-ai-copy", display_title: "去AI味写作" } });
  await app.choose(file("de-ai-copy.zip", "---\nname: de-ai-copy\n---\nRewrite a draft."));
  await app.submit();
  assert.equal(app.db.skills[0].title, "de-ai-copy");
  assert.equal(app.db.skills[0].display_title, "去AI味写作");
  assert.match(app.db.skills[0].slug, /^de-ai-copy-/);
  assert.match(app.db.versions[0].skill_md, /name: de-ai-copy/);
});

test("an upload with no function title waits for the user to supply one before storing files", async () => {
  const app = await harness({ fields: { display_title: "   " } });
  await app.choose(file("new.zip", "NEW DOCUMENT"));
  await app.submit();
  assert.equal(app.db.skills.length, 0);
  assert.equal(app.calls.some(call => call.operation === "upload"), false);
  assert.match(app.element("flash").textContent, /功能标题/);
  assert.equal(app.element("up-btn").disabled, false);
  app.fields.display_title = "生成课程提纲";
  await app.submit();
  assert.equal(app.db.skills[0].display_title, "生成课程提纲");
});

test("replacing a ZIP refreshes automatically read names, function titles, and descriptions", async () => {
  const app = await harness({ useDOMFields: true });
  await app.choose(file("de-ai-copy.zip", "---\nname: de-ai-copy\ndescription: 将已有文稿改得更自然。\n---\nRewrite a draft."));
  assert.equal(app.element("f-title").value, "de-ai-copy");
  assert.equal(app.element("f-display-title").value, "去AI味写作");
  assert.equal(app.element("f-desc").value, "将已有文稿改得更自然。");
  const nextMarkdown = "---\nname: book-insights\ndescription: 根据书籍和思考生成朋友圈文案。\n---\nWrite a book reflection.";
  await app.choose(file("book-insights.zip", nextMarkdown));
  assert.equal(app.element("f-title").value, "book-insights");
  assert.equal(app.element("f-display-title").value, "生成读书感悟朋友圈");
  assert.equal(app.element("f-desc").value, "根据书籍和思考生成朋友圈文案。");
  await app.submit();
  assert.equal(app.db.skills[0].title, "book-insights");
  assert.equal(app.db.skills[0].display_title, "生成读书感悟朋友圈");
  assert.equal(app.db.versions[0].skill_md, nextMarkdown);
});

test("replacing a ZIP preserves a function title the user has manually edited", async () => {
  const app = await harness({ useDOMFields: true });
  await app.choose(file("de-ai-copy.zip", "---\nname: de-ai-copy\ndescription: 将已有文稿改得更自然。\n---\nRewrite a draft."));
  app.input("f-display-title", "把读书笔记改成自然的朋友圈文案");
  await app.choose(file("book-insights.zip", "---\nname: book-insights\ndescription: 根据书籍和思考生成朋友圈文案。\n---\nWrite a book reflection."));
  assert.equal(app.element("f-title").value, "book-insights");
  assert.equal(app.element("f-display-title").value, "把读书笔记改成自然的朋友圈文案");
  assert.equal(app.element("f-desc").value, "根据书籍和思考生成朋友圈文案。");
  await app.submit();
  assert.equal(app.db.skills[0].display_title, "把读书笔记改成自然的朋友圈文案");
});

test("uploading a replacement version keeps the existing original name and function title", async () => {
  const app = await harness({ existing: true, useDOMFields: true });
  assert.equal(app.element("f-title").value, app.original.title);
  assert.equal(app.element("f-display-title").value, app.original.display_title);
  await app.choose(file("de-ai-copy.zip", "---\nname: de-ai-copy\ndescription: 将已有文稿改得更自然。\n---\nRewrite a draft."));
  await app.choose(file("book-insights.zip", "---\nname: book-insights\ndescription: 根据书籍和思考生成朋友圈文案。\n---\nWrite a book reflection."));
  assert.equal(app.element("f-title").value, app.original.title);
  assert.equal(app.element("f-display-title").value, app.original.display_title);
  assert.equal(app.element("f-desc").value, app.original.description);
  await app.submit();
  assert.equal(app.db.skills[0].title, app.original.title);
  assert.equal(app.db.skills[0].display_title, app.original.display_title);
  assert.equal(app.db.versions.length, 2);
});
