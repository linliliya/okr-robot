(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.SkillHubReadme = api;
})(typeof window !== "undefined" ? window : null, function () {
  "use strict";

  const CONTROL = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/;
  const REPOSITORY = /^[a-z0-9][a-z0-9._-]*\/[a-z0-9._-]+$/i;
  const encodePath = value => value.split("/").map(encodeURIComponent).join("/");

  function httpsUrl(value) {
    if (typeof value !== "string" || CONTROL.test(value) || value.includes("\\")) return "";
    const text = value.trim();
    if (!/^https:\/\//i.test(text)) return "";
    try {
      const url = new URL(text);
      return url.protocol === "https:" && !url.username && !url.password ? url.href : "";
    } catch (_) { return ""; }
  }

  function repositoryContext(source) {
    if (!source || typeof source !== "object" || typeof source.repo !== "string" || !REPOSITORY.test(source.repo)
      || source.repo.split("/").some(part => part === "." || part === "..")) return null;
    const { repo, ref, path } = source;
    const validPath = value => typeof value === "string" && value && !CONTROL.test(value)
      && !value.includes("\\") && value.split("/").every(part => part && part !== "." && part !== "..");
    if (!validPath(ref) || !validPath(path)) return null;
    const url = httpsUrl(source.url);
    if (!url) return null;
    const expected = `https://github.com/${repo}/blob/${encodePath(ref)}/${encodePath(path)}`;
    // A separately verified ref is required. A slash inside a Git branch name
    // cannot be inferred safely from the public blob URL alone.
    if (new URL(url).href !== new URL(expected).href) return null;
    return { repo, ref, path, url };
  }

  function resourceUrl(value, context, image) {
    if (typeof value !== "string" || CONTROL.test(value) || value.includes("\\")) return "";
    const text = value.trim();
    if (!text) return "";
    if (/^[a-z][a-z\d+.-]*:/i.test(text)) return httpsUrl(text);
    if (!context || text.startsWith("//")) return "";
    const match = text.match(/^([^?#]*)([?#].*)?$/);
    if (!match) return "";
    const [, relativePath, suffix = ""] = match;
    const parts = relativePath.startsWith("/") ? [] : context.path.split("/").slice(0, -1);
    if (!relativePath) {
      if (image) return "";
      return httpsUrl(context.url + suffix);
    }
    for (const encoded of relativePath.split("/")) {
      let part;
      try { part = decodeURIComponent(encoded); } catch (_) { return ""; }
      if (CONTROL.test(part) || /[\\/]/.test(part)) return "";
      if (!part || part === ".") continue;
      if (part === "..") {
        if (!parts.length) return "";
        parts.pop();
      } else parts.push(part);
    }
    if (!parts.length) return "";
    const prefix = image
      ? `https://raw.githubusercontent.com/${context.repo}/${encodePath(context.ref)}/`
      : `https://github.com/${context.repo}/blob/${encodePath(context.ref)}/`;
    return httpsUrl(prefix + parts.map(encodeURIComponent).join("/") + suffix);
  }

  function stripFrontmatter(markdown) {
    return markdown.replace(/^\uFEFF/, "").replace(/^---[ \t]*\r?\n[\s\S]*?\r?\n(?:---|\.\.\.)[ \t]*(?:\r?\n|$)/, "");
  }

  function render(source, { marked, DOMPurify, document: domDocument } = {}) {
    if (!source || typeof source.markdown !== "string") throw new TypeError("文档内容无法读取");
    if (!marked || typeof marked.parse !== "function" || !DOMPurify || typeof DOMPurify.sanitize !== "function"
      || !domDocument || typeof domDocument.createElement !== "function") {
      throw new TypeError("文档预览暂时不可用");
    }
    const html = marked.parse(stripFrontmatter(source.markdown), { gfm: true, async: false });
    if (typeof html !== "string") throw new TypeError("文档内容无法读取");
    const fragment = DOMPurify.sanitize(html, {
      RETURN_DOM_FRAGMENT: true,
      ALLOWED_TAGS: ["a", "p", "br", "hr", "blockquote", "pre", "code", "em", "strong", "del", "s",
        "ul", "ol", "li", "h1", "h2", "h3", "h4", "h5", "h6", "table", "thead", "tbody", "tfoot",
        "tr", "th", "td", "img", "details", "summary", "kbd", "samp", "sup", "sub"],
      ALLOWED_ATTR: ["href", "src", "alt", "title", "colspan", "rowspan", "start", "reversed", "align"],
      FORBID_TAGS: ["script", "style", "iframe", "form", "svg", "math", "object", "embed", "template"],
      FORBID_CONTENTS: ["script", "style", "iframe", "form", "svg", "math", "object", "embed", "template"],
      FORBID_ATTR: ["id", "name", "style", "srcset"],
      ALLOW_DATA_ATTR: false,
      ALLOW_ARIA_ATTR: false
    });
    const container = domDocument.createElement("div");
    container.appendChild(fragment);
    const context = repositoryContext(source);
    for (const link of container.querySelectorAll("a")) {
      const href = resourceUrl(link.getAttribute("href"), context, false);
      if (!href) {
        link.replaceWith(...link.childNodes);
        continue;
      }
      link.setAttribute("href", href);
      link.setAttribute("target", "_blank");
      link.setAttribute("rel", "noopener noreferrer");
    }
    for (const image of container.querySelectorAll("img")) {
      const src = resourceUrl(image.getAttribute("src"), context, true);
      if (!src) { image.remove(); continue; }
      image.setAttribute("src", src);
      image.setAttribute("loading", "lazy");
      image.setAttribute("referrerpolicy", "no-referrer");
    }
    return container.innerHTML;
  }

  return { render };
});
