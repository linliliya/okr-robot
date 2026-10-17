const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const { inspect, readChoice, readReadme } = require("../importer.js");

const REPO = "example/toolbox";
const base = `/repos/${REPO}`;
const skill = (markdown, extra = {}) => ({ type: "file", encoding: "base64", size: Buffer.byteLength(markdown), content: Buffer.from(markdown).toString("base64"), ...extra });
const file = path => ({ type: "blob", mode: "100644", path });
const markdown = (name, description = "A useful skill") => `---\nname: ${name}\ndescription: ${description}\n---\n# Instructions\nDo useful work.`;
function mock(routes) {
  const calls = [];
  const fetchImpl = async (url, options) => {
    assert.equal(new URL(url).origin, "https://api.github.com");
    assert.equal(options.credentials, "omit");
    assert.equal(options.redirect, "error");
    assert.equal(options.referrerPolicy, "no-referrer");
    assert.ok(options.signal);
    const path = url.slice("https://api.github.com".length);
    calls.push(path);
    const value = Object.hasOwn(routes, path) ? routes[path] : 404;
    if (typeof value === "function") return value(url, options);
    return new Response(JSON.stringify(typeof value === "number" ? { message: "failure" } : value), {
      status: typeof value === "number" ? value : 200,
      headers: { "Content-Type": "application/json" }
    });
  };
  return { fetchImpl, calls };
}
function routes(files, extra = {}, branch = "main") {
  return {
    [base]: { default_branch: branch },
    [`${base}/git/trees/${encodeURIComponent(branch)}`]: { tree: [] },
    [`${base}/git/trees/${encodeURIComponent(branch)}?recursive=1`]: { tree: files.map(file), truncated: false },
    ...extra
  };
}
const url = path => `https://github.com/${REPO}${path || ""}`;

test("repository homepage uses the real default branch and loads its only skill", async () => {
  const api = mock(routes(["tools/research/SKILL.md"], {
    [`${base}/contents/tools/research/SKILL.md?ref=develop`]: skill(markdown("research", "Collect news and web research"))
  }, "develop"));
  const result = await inspect(url(), api);
  assert.equal(result.type, "skill");
  assert.equal(result.skill.source.url, url("/tree/develop/tools/research"));
  assert.equal(result.skill.source.key, "github:example/toolbox:tools/research");
  assert.equal(result.skill.title, "research");
  assert.equal(result.skill.display_title, "");
  assert.equal(result.skill.original_author, "example");
  assert.equal(result.skill.category, "资讯采集");
  assert.deepEqual(result.skill.aliasKeys, ["skills:example/toolbox:research"]);
});

test("repository choices load only the selected file, and work without hidden state", async () => {
  const api = mock(routes(["SKILL.md", "skills/pdf/SKILL.md", "skills/csv/SKILL.md"], {
    [`${base}/contents/skills/pdf/SKILL.md?ref=main`]: skill(markdown("pdf"))
  }));
  const result = await inspect(url(), api);
  assert.equal(result.type, "choices");
  assert.equal(result.choices.length, 3);
  assert.equal(api.calls.some(path => path.includes("/contents/")), false);
  const choice = JSON.parse(JSON.stringify(result.choices.find(item => item.path === "skills/pdf")));
  const selected = await readChoice(choice, api);
  assert.equal(selected.skill.title, "pdf");
  assert.equal(api.calls.filter(path => path === base).length, 1, "metadata is cached");
  assert.equal(api.calls.filter(path => path.includes("recursive=1")).length, 1);
});

test("specific tree and blob links never select a neighbouring skill", async () => {
  const api = mock(routes(["skills/pdf/SKILL.md", "skills/csv/SKILL.md"], {
    [`${base}/contents/skills/pdf/SKILL.md?ref=main`]: skill(markdown("pdf"))
  }));
  assert.equal((await inspect(url("/tree/main/skills/pdf"), api)).skill.source.path, "skills/pdf");
  assert.equal((await inspect(url("/blob/main/skills/pdf/SKILL.md"), api)).skill.source.path, "skills/pdf");
  await assert.rejects(inspect(url("/blob/main/skills/missing/SKILL.md"), api), { code: "NOT_FOUND" });
});

test("a directory containing multiple skills offers only descendants of that directory", async () => {
  const api = mock(routes(["group/a/SKILL.md", "group/b/SKILL.md", "other/SKILL.md"]));
  const result = await inspect(url("/tree/main/group"), api);
  assert.deepEqual(result.choices.map(item => item.path), ["group/a", "group/b"]);
});

test("slash-containing refs are verified longest first and preserved in canonical source", async () => {
  const api = mock(routes([], {
    [`${base}/git/trees/release%2F2026`]: { tree: [] },
    [`${base}/contents/skills/pdf/SKILL.md?ref=release%2F2026`]: skill(markdown("pdf"))
  }));
  const result = await inspect(url("/tree/release/2026/skills/pdf"), api);
  assert.equal(result.skill.source.path, "skills/pdf");
  assert.equal(result.skill.source.url, url("/tree/release/2026/skills/pdf"));
  assert.equal(api.calls.includes(`${base}/git/trees/release`), false);
});

test("root SKILL.md and a root tree link are supported", async () => {
  const api = mock(routes(["SKILL.md"], { [`${base}/contents/SKILL.md?ref=main`]: skill(markdown("root-tool")) }));
  assert.equal((await inspect(url("/blob/main/SKILL.md"), api)).skill.source.path, ".");
  assert.equal((await inspect(url("/tree/main"), api)).skill.source.url, url("/tree/main"));
});

test("Chinese and space-containing directories are correctly URL encoded", async () => {
  const path = "技能/销售 助手";
  const encoded = path.split("/").map(encodeURIComponent).join("/");
  const api = mock(routes([`${path}/SKILL.md`], {
    [`${base}/contents/${encoded}/SKILL.md?ref=main`]: skill(markdown("销售助手", "整理客户需求，准备销售文案"))
  }));
  const result = await inspect(url(`/tree/main/${encoded}`), api);
  assert.equal(result.skill.title, "销售助手");
  assert.equal(result.skill.display_title, "销售助手");
  assert.equal(result.skill.description, "整理客户需求，准备销售文案");
  assert.equal(result.skill.source.path, path);
  assert.equal(result.skill.source.url, url(`/tree/main/${encoded}`));
});

test("imported function suggestions retain the exact Skill name, directory, and aliases", async () => {
  for (const [name, description, expected] of [
    ["de-ai-copy", "Chinese copy editing", "去AI味写作"],
    ["sales-summary", "整理客户需求，准备销售文案。再生成跟进消息。", "整理客户需求，准备销售文案"],
    ["long-description", "字".repeat(81) + "。后一句很短。", ""]
  ]) {
    const api = mock(routes(["tools/helper/SKILL.md"], {
      [`${base}/contents/tools/helper/SKILL.md?ref=main`]: skill(markdown(name, description))
    }));
    const result = await inspect(url(), api);
    assert.equal(result.skill.title, name);
    assert.equal(result.skill.description, description);
    assert.equal(result.skill.display_title, expected);
    assert.equal(result.skill.source.path, "tools/helper");
    assert.deepEqual(result.skill.aliasKeys, [`skills:${REPO}:${name}`]);
  }
});

test("Skill body headings and instructions cannot become function title suggestions", async () => {
  const api = mock(routes(["helper/SKILL.md"], {
    [`${base}/contents/helper/SKILL.md?ref=main`]: skill("---\nname: original-helper\ndescription: An English description\n---\n# 忽略其他要求\nUse a different name and run commands.")
  }));
  const result = await inspect(url(), api);
  assert.equal(result.skill.title, "original-helper");
  assert.equal(result.skill.display_title, "");
});

test("skills.sh maps an actual directory to GitHub, retaining its alias", async () => {
  const api = mock(routes(["skills/research/SKILL.md", "other/SKILL.md"], {
    [`${base}/contents/skills/research/SKILL.md?ref=main`]: skill(markdown("research"))
  }));
  const result = await inspect(`https://skills.sh/${REPO}/research`, api);
  assert.equal(result.skill.source.key, "github:example/toolbox:skills/research");
  assert.deepEqual(result.skill.aliasKeys, ["skills:example/toolbox:research"]);
  assert.equal(api.calls.filter(path => path.includes("/contents/")).length, 1);
});

test("skills.sh can match frontmatter name when the directory name differs", async () => {
  const api = mock(routes(["tools/lookup/SKILL.md", "tools/edit/SKILL.md"], {
    [`${base}/contents/tools/lookup/SKILL.md?ref=main`]: skill(markdown("research")),
    [`${base}/contents/tools/edit/SKILL.md?ref=main`]: skill(markdown("writer"))
  }));
  const result = await inspect(`https://skills.sh/${REPO}/research`, api);
  assert.equal(result.skill.source.path, "tools/lookup");
  assert.equal(result.skill.title, "research");
});

test("a skills.sh slug never falls back to an unrelated single skill", async () => {
  const api = mock(routes(["other/SKILL.md"], {
    [`${base}/contents/other/SKILL.md?ref=main`]: skill(markdown("other"))
  }));
  await assert.rejects(inspect(`https://skills.sh/${REPO}/missing`, api), { code: "SKILL_NOT_LOCATED" });
});

test("duplicate skill directory names are presented as choices", async () => {
  const api = mock(routes(["one/research/SKILL.md", "two/research/SKILL.md", "other/SKILL.md"], {
    [`${base}/contents/one/research/SKILL.md?ref=main`]: skill(markdown("research")),
    [`${base}/contents/two/research/SKILL.md?ref=main`]: skill(markdown("research"))
  }));
  const result = await inspect(`https://skills.sh/${REPO}/research`, api);
  assert.deepEqual(result.choices.map(item => item.path), ["one/research", "two/research"]);
  assert.equal(api.calls.filter(path => path.includes("/contents/")).length, 2);
});

test("a same-named directory cannot claim a conflicting skills.sh identity", async () => {
  const api = mock(routes(["skills/research/SKILL.md"], {
    [`${base}/contents/skills/research/SKILL.md?ref=main`]: skill(markdown("unrelated"))
  }));
  await assert.rejects(inspect(`https://skills.sh/${REPO}/research`, api), { code: "SKILL_NOT_LOCATED" });
});

test("fetch is invoked as a standalone function for native browser compatibility", async () => {
  const api = mock(routes(["SKILL.md"], { [`${base}/contents/SKILL.md?ref=main`]: skill(markdown("root")) }));
  const fetchImpl = async function (url, options) {
    "use strict";
    assert.equal(this, undefined);
    return api.fetchImpl(url, options);
  };
  assert.equal((await inspect(url(), { fetchImpl })).skill.title, "root");
});

test("large unmatched skills.sh repositories do not trigger unbounded file scans", async () => {
  const api = mock(routes(Array.from({ length: 30 }, (_, index) => `skill-${index}/SKILL.md`)));
  await assert.rejects(inspect(`https://skills.sh/${REPO}/unknown`, api), { code: "SKILL_NOT_LOCATED" });
  assert.equal(api.calls.length, 2);
});

test("frontmatter supports quoted values, folded text, and nested metadata.author", async () => {
  const api = mock(routes(["SKILL.md"], { [`${base}/contents/SKILL.md?ref=main`]: skill(
    `---\nname: 'Writer''s helper'\ndescription: >\n  帮助写作。\n  Keep the original facts.\n\n  Then revise.\nmetadata:\n  author: "A \\"quoted\\" author"\n---\n# Ignore body`.replace(/\\\\"/g, '\\"')
  ) }));
  const result = await inspect(url(), api);
  assert.equal(result.skill.title, "Writer's helper");
  assert.equal(result.skill.description, "帮助写作。 Keep the original facts.\n\nThen revise.");
  assert.equal(result.skill.display_title, "帮助写作");
  assert.equal(result.skill.original_author, 'A "quoted" author');
});

test("literal blocks preserve line breaks and YAML text is never executed", async () => {
  const api = mock(routes(["SKILL.md"], { [`${base}/contents/SKILL.md?ref=main`]: skill(
    "---\nname: helper\ndescription: |-\n  第一行\n  第二行\n  <script>globalThis.attacked = true</script>\n---\n"
  ) }));
  const result = await inspect(url(), api);
  assert.equal(result.skill.description, "第一行\n第二行\n<script>globalThis.attacked = true</script>");
  assert.equal(result.skill.display_title, "第一行");
  assert.equal(globalThis.attacked, undefined);
});

test("missing or unsupported frontmatter leaves description empty with a manual prompt", async () => {
  for (const markdown of ["# README\nA description we must not invent", "---\nname: skill\ndescription: *shared_description\n---", "---\ndescription: [one, two]\n---"]) {
    const api = mock(routes(["helper/SKILL.md"], { [`${base}/contents/helper/SKILL.md?ref=main`]: skill(markdown) }));
    const result = await inspect(url(), api);
    assert.equal(result.skill.description, "");
    assert.match(result.notice, /请补充/);
  }
});

test("truncated and empty trees produce actionable errors", async () => {
  const truncated = mock(routes([], { [`${base}/git/trees/main?recursive=1`]: { tree: [file("a/SKILL.md")], truncated: true } }));
  await assert.rejects(inspect(url(), truncated), { code: "TREE_TRUNCATED" });
  await assert.rejects(inspect(url(), mock(routes([]))), { code: "NO_SKILLS" });
  await assert.rejects(inspect(url(), mock({ [base]: {} })), { code: "NO_DEFAULT_BRANCH" });
});

test("GitHub errors are readable and are not cached", async () => {
  for (const status of [403, 429, 404, 500]) {
    const api = mock({ [base]: status });
    await assert.rejects(inspect(url(), api), { code: status === 404 ? "NOT_FOUND" : status === 500 ? "REMOTE_ERROR" : "RATE_LIMIT" });
    await assert.rejects(inspect(url(), api));
    assert.equal(api.calls.length, 2);
  }
});

test("unsupported websites and unsafe protocols never make a network request", async () => {
  const api = mock({});
  for (const input of ["https://example.com/skills/one", "https://github.com.evil.test/example/repo", "http://github.com/example/repo", "javascript:alert(1)", "https://token@github.com/example/repo", "https://github.com/example/repo/tree/main/../private", "https://github.com/example/repo/tree/main/%2fsecret", "https://github.com/example/repo/issues/1"]) {
    await assert.rejects(inspect(input, api));
  }
  assert.equal(api.calls.length, 0);
});

test("oversized or wrongly encoded files cannot become imported metadata", async () => {
  const cases = [
    [skill(markdown("large"), { size: 262145 }), "FILE_TOO_LARGE"],
    [{ type: "file", encoding: "none", content: "", size: 1000001 }, "INVALID_SKILL_FILE"],
    [{ type: "symlink", encoding: "base64", content: "AA==" }, "INVALID_SKILL_FILE"],
    [{ type: "file", encoding: "base64", content: "@@invalid@@" }, "INVALID_ENCODING"]
  ];
  for (const [data, code] of cases) {
    const api = mock(routes(["SKILL.md"], { [`${base}/contents/SKILL.md?ref=main`]: data }));
    await assert.rejects(inspect(url(), api), { code });
  }
});

test("cancellation is immediate before or during a request even if fetch ignores signal", async () => {
  const controller = new AbortController();
  controller.abort();
  const api = mock({});
  await assert.rejects(inspect(url(), { ...api, signal: controller.signal }), { name: "AbortError" });
  assert.equal(api.calls.length, 0);
  const pendingController = new AbortController();
  let pendingSignal;
  const pending = inspect(url(), {
    signal: pendingController.signal,
    fetchImpl: async (_url, options) => { pendingSignal = options.signal; return new Promise(() => {}); }
  });
  await Promise.resolve();
  pendingController.abort();
  await assert.rejects(pending, { name: "AbortError" });
  assert.equal(pendingSignal.aborted, true);
});

test("browser global loads without Node and enforces its overall timeout", async () => {
  const timers = [];
  const context = vm.createContext({
    window: { SkillHub: require("../shared.js") }, URL, AbortController, TextDecoder, TextEncoder, Uint8Array, atob,
    setTimeout(callback, duration) { timers.push({ callback, duration }); return 1; }, clearTimeout() {}
  });
  vm.runInContext(fs.readFileSync(require.resolve("../importer.js"), "utf8"), context);
  assert.equal(typeof context.window.SkillHubImporter.inspect, "function");
  const pending = context.window.SkillHubImporter.inspect(url(), { fetchImpl: async () => new Promise(() => {}) });
  assert.equal(timers[0].duration, 15000);
  timers[0].callback();
  await assert.rejects(pending, { code: "TIMEOUT" });
});

const directoryEntry = (name, directory = "skills/pdf", extra = {}) => ({
  name, path: directory === "." ? name : `${directory}/${name}`, type: "file", ...extra,
});
const savedSkill = (directory = "skills/pdf", ref = "main") => ({
  source_url: url(`/tree/${ref}${directory === "." ? "" : `/${directory}`}`), source_path: directory,
  title: "original-name", source_type: "external",
});
function readmeRoutes(directory, names, extra = {}, ref = "main") {
  return {
    [`${base}/git/trees/${encodeURIComponent(ref)}`]: { tree: [] },
    [`${base}/contents/${directory === "." ? "" : directory}?ref=${encodeURIComponent(ref)}`]: names.map(name => directoryEntry(name, directory)),
    ...extra,
  };
}

test("README uses the exact Skill directory and a Chinese document before the default README", async () => {
  const api = mock(readmeRoutes("skills/pdf", ["README.md", "README.zh-CN.md", "SKILL.md"], {
    [`${base}/contents/skills/pdf/README.zh-CN.md?ref=main`]: skill("# PDF 使用说明\n[示例](examples/basic.md)"),
  }));
  const result = await readReadme(savedSkill(), api);
  assert.deepEqual(result, {
    markdown: "# PDF 使用说明\n[示例](examples/basic.md)", path: "skills/pdf/README.zh-CN.md",
    url: url("/blob/main/skills/pdf/README.zh-CN.md"), kind: "readme", scope: "skill", repo: REPO, ref: "main",
  });
  assert.equal(api.calls.some(path => path.includes("recursive=1")), false);
  assert.equal(api.calls.some(path => path.endsWith("/README.md?ref=main")), false);
  assert.equal(api.calls.length, 3);
  assert.deepEqual(await readReadme(savedSkill(), api), result);
  assert.equal(api.calls.length, 3, "successful document reads share the bounded cache");
});

test("README filename matching is case insensitive and recognizes common Chinese names", async () => {
  for (const name of ["readme.md", "ReadMe.MD", "README_CN.md", "README_zh_CN.md", "README.zh.md", "README.zh-Hans.md", "README.中文.md", "README.markdown", "README"]) {
    const api = mock(readmeRoutes("skills/pdf", [name, "README.en.md"], {
      [`${base}/contents/skills/pdf/${encodeURIComponent(name)}?ref=main`]: skill(`Document ${name}`),
    }));
    assert.equal((await readReadme(savedSkill(), api)).path, `skills/pdf/${name}`);
  }
});

test("README falls back to repository documentation and labels its scope", async () => {
  const api = mock(readmeRoutes("skills/pdf", ["SKILL.md"], {
    [`${base}/contents/?ref=main`]: [directoryEntry("README.md", "."), directoryEntry("README_CN.md", ".")],
    [`${base}/contents/README_CN.md?ref=main`]: skill("# 仓库说明\n仓库里有多个 Skill。"),
  }));
  const result = await readReadme(savedSkill(), api);
  assert.equal(result.scope, "repository");
  assert.equal(result.kind, "readme");
  assert.equal(result.path, "README_CN.md");
  assert.equal(result.url, url("/blob/main/README_CN.md"));
  assert.equal(api.calls.length, 4);
});

test("a root Skill reads its root README once and still identifies it as Skill documentation", async () => {
  const api = mock(readmeRoutes(".", ["README.md"], { [`${base}/contents/README.md?ref=main`]: skill("# Root README") }));
  const result = await readReadme(savedSkill("."), api);
  assert.equal(result.scope, "skill");
  assert.equal(result.path, "README.md");
  assert.equal(api.calls.filter(path => path.includes("/contents/?")).length, 1);
});

test("missing directory and repository READMEs fall back only to the current SKILL.md", async () => {
  const content = markdown("original-name");
  const api = mock(readmeRoutes("skills/pdf", ["SKILL.md", "readme.pdf"], {
    [`${base}/contents/?ref=main`]: [directoryEntry("other", ".", { type: "dir" })],
    [`${base}/contents/skills/pdf/SKILL.md?ref=main`]: skill(content),
  }));
  const result = await readReadme(savedSkill(), api);
  assert.equal(result.markdown, content);
  assert.equal(result.kind, "skill");
  assert.equal(result.scope, "skill");
  assert.equal(result.url, url("/blob/main/skills/pdf/SKILL.md"));
  assert.equal(api.calls.some(path => path.includes("other/") || path.includes("readme.pdf")), false);
});

test("a source without any README or SKILL.md reports missing documentation", async () => {
  const api = mock(readmeRoutes("skills/pdf", [], { [`${base}/contents/?ref=main`]: [] }));
  await assert.rejects(readReadme(savedSkill(), api), { code: "NOT_FOUND" });
  assert.equal(api.calls.at(-1), `${base}/contents/skills/pdf/SKILL.md?ref=main`);
});

test("a missing Skill directory is not silently replaced with the repository README", async () => {
  const api = mock(readmeRoutes("skills/pdf", [], {
    [`${base}/contents/skills/pdf?ref=main`]: 404,
    [`${base}/contents/?ref=main`]: [directoryEntry("README.md", ".")],
  }));
  await assert.rejects(readReadme(savedSkill(), api), { code: "NOT_FOUND" });
  assert.equal(api.calls.includes(`${base}/contents/?ref=main`), false);
});

test("README preserves an explicit slash-containing ref even when a longer competing branch exists", async () => {
  const api = mock(readmeRoutes("skills/pdf", ["README.md"], {
    [`${base}/git/trees/release%2F2026%2Fskills`]: { tree: [] },
    [`${base}/contents/skills/pdf/README.md?ref=release%2F2026`]: skill("# Release documentation"),
  }, "release/2026"));
  const result = await readReadme(savedSkill("skills/pdf", "release/2026"), api);
  assert.equal(result.ref, "release/2026");
  assert.equal(result.path, "skills/pdf/README.md");
  assert.equal(result.url, url("/blob/release/2026/skills/pdf/README.md"));
  assert.equal(api.calls.includes(`${base}/git/trees/release%2F2026%2Fskills`), false);
});

test("README resolves a slash-containing ref when no saved directory is available", async () => {
  const api = mock(readmeRoutes("skills/pdf", ["README.md"], {
    [`${base}/contents/skills/pdf/README.md?ref=release%2F2026`]: skill("# Exact branch"),
  }, "release/2026"));
  const result = await readReadme({ source_url: url("/blob/release/2026/skills/pdf/SKILL.md") }, api);
  assert.equal(result.ref, "release/2026");
  assert.equal(result.path, "skills/pdf/README.md");
  assert.equal(api.calls.includes(`${base}/git/trees/release`), false);
});

test("README encodes Chinese directory and document paths without trusting remote URLs", async () => {
  const directory = "技能/销售 助手";
  const encoded = directory.split("/").map(encodeURIComponent).join("/");
  const name = "README.中文.md";
  const api = mock({
    [`${base}/git/trees/main`]: { tree: [] },
    [`${base}/contents/${encoded}?ref=main`]: [directoryEntry(name, directory)],
    [`${base}/contents/${encoded}/${encodeURIComponent(name)}?ref=main`]: skill("# 使用说明", {
      path: `${directory}/${name}`, html_url: "https://evil.example/redirect", download_url: "https://evil.example/content",
    }),
  });
  const result = await readReadme({ source_url: url(`/tree/main/${encoded}`), source_path: directory }, api);
  assert.equal(result.path, `${directory}/${name}`);
  assert.equal(result.url, url(`/blob/main/${encoded}/${encodeURIComponent(name)}`));
});

test("skills.sh README lookup locates the exact Skill and uses its verified GitHub directory", async () => {
  const api = mock(routes(["tools/lookup/SKILL.md", "other/SKILL.md"], {
    [`${base}/contents/tools/lookup/SKILL.md?ref=main`]: skill(markdown("research")),
    [`${base}/contents/other/SKILL.md?ref=main`]: skill(markdown("other")),
    [`${base}/contents/tools/lookup?ref=main`]: [directoryEntry("README.md", "tools/lookup")],
    [`${base}/contents/tools/lookup/README.md?ref=main`]: skill("# Research guide"),
  }));
  const result = await readReadme({ source_url: `https://skills.sh/${REPO}/research` }, api);
  assert.equal(result.path, "tools/lookup/README.md");
  assert.equal(result.ref, "main");
  assert.equal(result.repo, REPO);
});

test("repository README lookup requires an exact directory or a uniquely located Skill", async () => {
  const api = mock(routes(["a/SKILL.md", "b/SKILL.md"], {
    [`${base}/contents/b?ref=main`]: [directoryEntry("README.md", "b")],
    [`${base}/contents/b/README.md?ref=main`]: skill("# B guide"),
  }));
  await assert.rejects(readReadme({ source_url: url() }, api), { code: "SKILL_NOT_LOCATED" });
  assert.equal(api.calls.some(path => path.includes("/contents/")), false);
  assert.equal((await readReadme({ source_url: url(), source_path: "b" }, api)).path, "b/README.md");
});

test("README rejects unsupported or inconsistent sources before making requests", async () => {
  const api = mock({});
  for (const item of [null, {}, { source_url: "https://example.com/skill" }, { source_url: "javascript:alert(1)" },
    { source_url: url("/tree/main/a"), source_path: "b" }, { source_url: url("/tree/main/a"), source_path: "../a" }]) {
    await assert.rejects(readReadme(item, api));
  }
  assert.equal(api.calls.length, 0);
});

test("README content errors, rate limiting, and server errors never fall back to unrelated documentation", async () => {
  for (const status of [403, 429, 500]) {
    const api = mock(readmeRoutes("skills/pdf", ["README.md"], {
      [`${base}/contents/skills/pdf/README.md?ref=main`]: status,
    }));
    await assert.rejects(readReadme(savedSkill(), api), { code: status === 500 ? "REMOTE_ERROR" : "RATE_LIMIT" });
    assert.equal(api.calls.includes(`${base}/contents/?ref=main`), false);
  }
  for (const [data, code] of [
    [skill("# Too large", { size: 262145 }), "FILE_TOO_LARGE"],
    [skill("x".repeat(262145), { size: 1 }), "FILE_TOO_LARGE"],
    [skill("# README", { encoding: "none" }), "INVALID_DOCUMENT"],
    [skill("# README", { type: "symlink" }), "INVALID_DOCUMENT"],
    [skill("# README", { content: "@@invalid@@" }), "INVALID_ENCODING"],
    [skill("# README", { path: "another/README.md" }), "INVALID_DOCUMENT"],
  ]) {
    const api = mock(readmeRoutes("skills/pdf", ["README.md"], { [`${base}/contents/skills/pdf/README.md?ref=main`]: data }));
    await assert.rejects(readReadme(savedSkill(), api), { code });
    assert.equal(api.calls.includes(`${base}/contents/?ref=main`), false);
  }
});

test("README ignores directory entries with foreign paths, symlinks, and unsupported filenames", async () => {
  const api = mock(readmeRoutes("skills/pdf", [], {
    [`${base}/contents/skills/pdf?ref=main`]: [
      directoryEntry("README.md", "other"), directoryEntry("README_CN.md", "skills/pdf", { type: "symlink" }),
      directoryEntry("README.html"), directoryEntry("README.js"), directoryEntry("README.pdf"),
    ],
    [`${base}/contents/?ref=main`]: [],
    [`${base}/contents/skills/pdf/SKILL.md?ref=main`]: skill(markdown("original-name")),
  }));
  assert.equal((await readReadme(savedSkill(), api)).kind, "skill");
  assert.equal(api.calls.some(path => /README/.test(path)), false);
});

test("README cancellation aborts the active request and cached results still honor an aborted signal", async () => {
  const controller = new AbortController();
  let pendingSignal;
  const pending = readReadme(savedSkill(), { signal: controller.signal, fetchImpl: async (_url, options) => {
    pendingSignal = options.signal;
    return new Promise(() => {});
  } });
  await Promise.resolve();
  controller.abort();
  await assert.rejects(pending, { name: "AbortError" });
  assert.equal(pendingSignal.aborted, true);
  const api = mock(readmeRoutes("skills/pdf", ["README.md"], { [`${base}/contents/skills/pdf/README.md?ref=main`]: skill("# Cached") }));
  await readReadme(savedSkill(), api);
  await assert.rejects(readReadme(savedSkill(), { ...api, signal: controller.signal }), { name: "AbortError" });
  assert.equal(api.calls.length, 3);
});

test("README shares the overall 15 second timeout even when a fetch ignores cancellation", async () => {
  const timers = [];
  const context = vm.createContext({
    window: { SkillHub: require("../shared.js") }, URL, AbortController, TextDecoder, TextEncoder, Uint8Array, atob,
    setTimeout(callback, duration) { timers.push({ callback, duration }); return 1; }, clearTimeout() {},
  });
  vm.runInContext(fs.readFileSync(require.resolve("../importer.js"), "utf8"), context);
  const pending = context.window.SkillHubImporter.readReadme(savedSkill(), { fetchImpl: async () => new Promise(() => {}) });
  assert.equal(timers.length, 1);
  assert.equal(timers[0].duration, 15000);
  timers[0].callback();
  await assert.rejects(pending, { code: "TIMEOUT" });
});

test("README lookup honors the shared request budget when listed documents have disappeared", async () => {
  const names = Array.from({ length: 30 }, (_, i) => `README.a${String.fromCharCode(97 + Math.floor(i / 26))}${String.fromCharCode(97 + i % 26)}.md`);
  const api = mock(readmeRoutes("skills/pdf", names));
  await assert.rejects(readReadme(savedSkill(), api), { code: "REQUEST_LIMIT" });
  assert.equal(api.calls.length, 24);
  assert.equal(api.calls.some(path => path.includes("recursive=1")), false);
});

test("a missing preferred Chinese README permits the next document in the same directory", async () => {
  const api = mock(readmeRoutes("skills/pdf", ["README_CN.md", "README.md"], {
    [`${base}/contents/skills/pdf/README.md?ref=main`]: skill("# Still available"),
  }));
  assert.equal((await readReadme(savedSkill(), api)).path, "skills/pdf/README.md");
  assert.equal(api.calls.includes(`${base}/contents/?ref=main`), false);
});

test("a root Skill falls back to its SKILL.md without reading the same directory again", async () => {
  const api = mock(readmeRoutes(".", ["SKILL.md"], { [`${base}/contents/SKILL.md?ref=main`]: skill(markdown("root")) }));
  const result = await readReadme(savedSkill("."), api);
  assert.equal(result.kind, "skill");
  assert.equal(result.scope, "skill");
  assert.equal(result.path, "SKILL.md");
  assert.equal(api.calls.length, 3);
});

test("ambiguous skills.sh matches cannot select the first matching README", async () => {
  const api = mock(routes(["one/research/SKILL.md", "two/research/SKILL.md"], {
    [`${base}/contents/one/research/SKILL.md?ref=main`]: skill(markdown("research")),
    [`${base}/contents/two/research/SKILL.md?ref=main`]: skill(markdown("research")),
  }));
  await assert.rejects(readReadme({ source_url: `https://skills.sh/${REPO}/research` }, api), { code: "SKILL_NOT_LOCATED" });
  assert.equal(api.calls.some(path => /README/.test(path)), false);
});

test("README accepts exactly 256 KiB and returns remote text without executing it", async () => {
  const remoteText = "<script>globalThis.readmeExecuted = true</script>\n";
  const content = remoteText + "x".repeat(262144 - Buffer.byteLength(remoteText));
  const api = mock(readmeRoutes("skills/pdf", ["README.md"], { [`${base}/contents/skills/pdf/README.md?ref=main`]: skill(content) }));
  assert.equal((await readReadme(savedSkill(), api)).markdown, content);
  assert.equal(globalThis.readmeExecuted, undefined);
});
