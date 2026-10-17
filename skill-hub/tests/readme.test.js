const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const { render } = require("../readme.js");

// The deployed page already loads marked and DOMPurify from its CDN. For these
// real-DOM regression tests, install marked@12.0.2, dompurify@3.1.5 and
// jsdom@26.1.0 in a temporary directory, then point NODE_PATH at node_modules.
let dependencies;
try {
  dependencies = {
    marked: require("marked"), DOMPurify: require("dompurify"), JSDOM: require("jsdom").JSDOM
  };
} catch (error) {
  if (error.code !== "MODULE_NOT_FOUND") throw error;
}

function fixture(markdown, extra = {}) {
  return {
    markdown, repo: "example/toolbox", ref: "main", path: "skills/writer/README.md",
    url: "https://github.com/example/toolbox/blob/main/skills/writer/README.md",
    kind: "readme", scope: "skill", ...extra
  };
}

function harness() {
  const window = new dependencies.JSDOM("<!doctype html><html><body></body></html>", {
    url: "https://internal.example.test/skill-hub/"
  }).window;
  const DOMPurify = dependencies.DOMPurify(window);
  const options = { marked: dependencies.marked, DOMPurify, document: window.document };
  return {
    window, options,
    render(source) {
      const container = window.document.createElement("div");
      container.innerHTML = render(source, options);
      return container;
    }
  };
}

test("README renderer is available without Node dependencies in the browser", () => {
  const context = vm.createContext({ window: {}, URL });
  vm.runInContext(fs.readFileSync(require.resolve("../readme.js"), "utf8"), context);
  assert.equal(typeof context.window.SkillHubReadme.render, "function");
  assert.throws(() => render({ markdown: "Hello" }), /暂时不可用/);
});

test("README rendering with real marked, DOMPurify, and a real DOM", {
  skip: !dependencies && "Set NODE_PATH to marked, dompurify, and jsdom test dependencies"
}, async t => {
  await t.test("keeps headings, tables, lists, code, and removes YAML metadata from Skill fallback", () => {
    const app = harness();
    const body = app.render(fixture(`---\nname: hidden-name\ndescription: hidden-description\n---\n# 使用方法\n\n- 读取文件\n- 整理结果\n\n|输入|输出|\n|---|---|\n|文稿|摘要|\n\n\`\`\`js\nconsole.log('<script>text only</script>');\n\`\`\`\n`, { kind: "skill" }));
    assert.equal(body.querySelector("h1").textContent, "使用方法");
    assert.equal(body.querySelectorAll("li").length, 2);
    assert.equal(body.querySelector("table tbody td").textContent, "文稿");
    assert.match(body.querySelector("pre code").textContent, /<script>text only<\/script>/);
    assert.doesNotMatch(body.textContent, /hidden-name|hidden-description/);
    assert.equal(body.querySelector("script"), null);
    app.window.close();
  });

  await t.test("removes active HTML, site-conflicting attributes, and unsafe resources", () => {
    const app = harness();
    const body = app.render(fixture(`<script>alert(1)</script><style>body{display:none}</style>
<iframe src="https://evil.test"></iframe><form><input name="x"></form>
<svg><a xlink:href="javascript:alert(1)">bad</a></svg><math><mi>x</mi></math>
<h2 id="app" name="route" style="position:fixed" onclick="alert(1)" data-action="delete">安全内容</h2>
<a href="javascript:alert(1)">脚本</a><a href="data:text/html,bad">数据</a><a href="http://example.test">明文</a>
<a href="mailto:hello@example.test">邮件</a><a href="java&#x0A;script:alert(1)">编码脚本</a>
<img src="data:image/png;base64,AA==" onerror="alert(1)"><img src="http://example.test/a.png">
<img src="https://images.example.test/a.png" onload="alert(1)" srcset="https://evil.test/a.png 2x" style="position:fixed">
<a href="https://docs.example.test/guide" target="_self" rel="opener">正常链接</a>`));
    assert.equal(body.querySelector("script,style,iframe,form,input,svg,math"), null);
    assert.equal(body.querySelector("[id],[name],[style],[onclick],[onerror],[onload],[srcset],[data-action]"), null);
    assert.equal(body.querySelectorAll("a").length, 1);
    assert.equal(body.querySelector("a").getAttribute("href"), "https://docs.example.test/guide");
    assert.equal(body.querySelector("a").getAttribute("target"), "_blank");
    assert.equal(body.querySelector("a").getAttribute("rel"), "noopener noreferrer");
    assert.equal(body.querySelectorAll("img").length, 1);
    assert.equal(body.querySelector("img").getAttribute("loading"), "lazy");
    assert.equal(body.querySelector("img").getAttribute("referrerpolicy"), "no-referrer");
    assert.equal(app.window.attacked, undefined);
    app.window.close();
  });

  await t.test("resolves relative documents and raw images with the verified slash-containing ref", () => {
    const app = harness();
    const body = app.render(fixture(`[说明](../guide/使用.md#安装) [章节](#usage)\n\n![图](<./images/流程 图.png>)\n\n<img src="/assets/root.png">`, {
      ref: "release/2026", url: "https://github.com/example/toolbox/blob/release/2026/skills/writer/README.md"
    }));
    assert.equal(body.querySelectorAll("a")[0].href, "https://github.com/example/toolbox/blob/release/2026/skills/guide/%E4%BD%BF%E7%94%A8.md#%E5%AE%89%E8%A3%85");
    assert.equal(body.querySelectorAll("a")[1].href, "https://github.com/example/toolbox/blob/release/2026/skills/writer/README.md#usage");
    assert.equal(body.querySelectorAll("img")[0].src, "https://raw.githubusercontent.com/example/toolbox/release/2026/skills/writer/images/%E6%B5%81%E7%A8%8B%20%E5%9B%BE.png");
    assert.equal(body.querySelectorAll("img")[1].src, "https://raw.githubusercontent.com/example/toolbox/release/2026/assets/root.png");
    const encoded = app.render(fixture("![图](./images/%E6%B5%81%E7%A8%8B%20%E5%9B%BE.png)", {
      ref: "release/2026", url: "https://github.com/example/toolbox/blob/release/2026/skills/writer/README.md"
    }));
    assert.equal(encoded.querySelector("img").src, "https://raw.githubusercontent.com/example/toolbox/release/2026/skills/writer/images/%E6%B5%81%E7%A8%8B%20%E5%9B%BE.png");
    app.window.close();
  });

  await t.test("cannot escape the repository or guess a branch from incomplete metadata", () => {
    const app = harness();
    const body = app.render(fixture('<a href="../../../other/repo">outside</a><img src="../../../other.png"><a href="//evil.test/a">scheme relative</a><img src="%2f%2fevil.test/a"><a href="bad%00name.md">controls</a>'));
    assert.equal(body.querySelector("a,img"), null);
    for (const extra of [{ ref: undefined }, { ref: "other" }, { url: "https://evil.test/README.md" }, { repo: "../toolbox" }, { repo: "example/.." }]) {
      const unknown = app.render(fixture('[relative](guide.md) ![image](a.png) [absolute](https://docs.example.test/) ![absolute](https://images.example.test/a.png)', extra));
      assert.equal(unknown.querySelectorAll("a").length, 1);
      assert.equal(unknown.querySelector("a").href, "https://docs.example.test/");
      assert.equal(unknown.querySelectorAll("img").length, 1);
      assert.equal(unknown.querySelector("img").src, "https://images.example.test/a.png");
    }
    app.window.close();
  });

  await t.test("sanitizer options and URL rewriting remain local to this render", () => {
    const app = harness();
    const ordinary = '<h2 id="existing-heading">Existing document</h2><a href="#existing-heading">Jump</a>';
    const before = app.options.DOMPurify.sanitize(ordinary);
    app.render(fixture('<h2 id="external-heading">README</h2><a href="guide.md">Guide</a>'));
    assert.equal(app.options.DOMPurify.sanitize(ordinary), before);
    assert.match(before, /id="existing-heading"/);
    assert.doesNotMatch(before, /target=/);
    app.window.close();
  });
});
