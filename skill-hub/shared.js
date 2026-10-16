(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.SkillHub = api;
})(typeof window !== "undefined" ? window : null, function () {
  "use strict";

  const CONTROL = /[\u0000-\u001f\u007f-\u009f]/;
  const REPOSITORY = /^[a-z0-9][a-z0-9._-]*\/[a-z0-9._-]+$/i;
  const SLUG = /^[a-z0-9][a-z0-9._-]*$/i;
  const ROLE_OPTIONS = Object.freeze(["通用", "内容创作", "课程运营", "销售", "教学教研", "人力行政", "产品经理", "设计师", "剪辑"]);
  const KNOWN_DISPLAY_TITLES = Object.freeze({
    "de-ai-copy": "去AI味写作",
    "ai-moments": "生成 AI 资讯朋友圈",
    "book-insights": "生成读书感悟朋友圈",
    "python-crawler-success": "生成 Python 爬虫训练营成功案例朋友圈"
  });

  function displayTitle(value) {
    if (typeof value !== "string" || CONTROL.test(value) || /[\u2028\u2029]/.test(value)) return "";
    const title = value.trim();
    return [...title].length <= 80 ? title : "";
  }

  function knownDisplayTitle(title) {
    return Object.hasOwn(KNOWN_DISPLAY_TITLES, title) ? KNOWN_DISPLAY_TITLES[title] : "";
  }

  function getDisplayTitle(skill) {
    const title = typeof skill?.title === "string" ? skill.title.trim() : "";
    return displayTitle(skill?.display_title) || knownDisplayTitle(title) || title;
  }

  function suggestDisplayTitle(skill) {
    const title = typeof skill?.title === "string" ? skill.title.trim() : "";
    const chineseTitle = displayTitle(title);
    if (/[\u3400-\u9fff]/.test(chineseTitle)) return chineseTitle;
    const known = knownDisplayTitle(title);
    if (known) return known;
    // Suggest only a complete sentence already present in the description.
    // Longer text and English-only descriptions need a title from the sharer.
    const description = typeof skill?.description === "string" ? skill.description.trim() : "";
    const firstSentence = displayTitle(description.split(/[。！？!?；;\r\n]|\.(?=\s|$)/, 1)[0]);
    return /[\u3400-\u9fff]/.test(firstSentence) ? firstSentence : "";
  }

  function normalizeRoles(raw) {
    if (!Array.isArray(raw)) return [];
    const roles = [...new Set(raw.filter(value => typeof value === "string").map(value => value.trim()).map(value => value === "市场运营" ? "课程运营" : value).filter(value => ROLE_OPTIONS.includes(value)))];
    return roles.includes("通用") ? ["通用"] : roles;
  }

  function fail(message) { throw new Error(message); }

  function decode(value) {
    try { return decodeURIComponent(value); }
    catch (_) { fail("链接中有无法识别的编码，请重新复制链接"); }
  }

  // Check the original path before URL() can silently remove dot segments.
  function checkedUrl(value) {
    if (typeof value !== "string" || CONTROL.test(value) || value.includes("\\")) {
      fail("请填写有效的 HTTPS 链接");
    }
    const raw = value.trim();
    if (!/^https:\/\//i.test(raw)) fail("请使用 https 开头的链接");
    const decoded = decode(raw);
    if (CONTROL.test(decoded) || decoded.includes("\\")) fail("链接中包含无效字符");
    const path = raw.match(/^https:\/\/[^/?#]*([^?#]*)/i)?.[1] || "";
    for (const part of path.split("/")) {
      const segment = decode(part);
      if (segment === "." || segment === ".." || segment.includes("/")) {
        fail("链接中不能包含跳转目录，请复制完整的 Skill 地址");
      }
    }
    let parsed;
    try { parsed = new URL(raw); }
    catch (_) { fail("请填写有效的 HTTPS 链接"); }
    if (parsed.protocol !== "https:" || !parsed.hostname || parsed.username || parsed.password) {
      fail("请填写不带账号密码的 HTTPS 链接");
    }
    return parsed;
  }

  function safeUrl(value) {
    try { return checkedUrl(value).href; }
    catch (_) { return ""; }
  }

  function directory(value) {
    if (typeof value !== "string") fail("请填写 Skill 在仓库中的目录");
    if (CONTROL.test(value) || value.includes("\\")) fail("Skill 目录中包含无效字符");
    const path = value.trim().replace(/\/+$/, "");
    if (!path) return "";
    if (path === ".") return ".";
    if (path.startsWith("/") || path.includes("//")) fail("请填写相对于仓库根目录的路径");
    for (const segment of path.split("/")) {
      const decoded = decode(segment);
      if (decoded === "." || decoded === ".." || /[\\/]/.test(decoded) || CONTROL.test(decoded)) {
        fail("Skill 目录不能包含跳转路径");
      }
    }
    return path;
  }

  function repository(parts) {
    if (parts.length < 2) fail("请粘贴具体 Skill 的完整链接");
    const repo = `${parts[0]}/${parts[1].replace(/\.git$/i, "")}`.toLowerCase();
    if (!REPOSITORY.test(repo) || repo.split("/").some(p => p === "." || p === "..")) {
      fail("仓库地址格式不正确，请重新复制链接");
    }
    return repo;
  }

  const encodePath = path => path.split("/").map(encodeURIComponent).join("/");

  function parseSource(rawUrl, explicitPath = "") {
    const parsed = checkedUrl(rawUrl);
    const host = parsed.hostname.toLowerCase().replace(/^www\./, "");
    const parts = parsed.pathname.replace(/^\/+|\/+$/g, "").split("/").map(decode);
    const override = directory(explicitPath);
    if (host === "github.com" || host === "skills.sh") {
      if (parsed.port) fail("请使用 GitHub 或 skills.sh 的标准 HTTPS 链接");
      if (parts.some(p => !p)) fail("链接中有空目录，请重新复制链接");
      const repo = repository(parts);
      if (host === "skills.sh") {
        if (parts.length !== 3 || !SLUG.test(parts[2])) {
          fail("请粘贴 skills.sh 上某一个 Skill 的页面链接");
        }
        // A directory listing slug is not proof of the path in the Git repository.
        const skillName = parts[2];
        return {
          url: `https://skills.sh/${repo}/${encodeURIComponent(skillName)}`,
          key: `skills:${repo}:${skillName}`, repo, path: "", skillName, kind: "skills", ref: ""
        };
      }

      let path = override;
      let ref = "";
      if (parts.length === 2) {
        if (!path) fail("这个链接指向整个仓库，请填写具体 Skill 目录；根目录就是 Skill 时填 .");
      } else {
        const marker = parts[2];
        if (marker !== "tree" && marker !== "blob") {
          fail("请粘贴 GitHub 的 Skill 目录链接，或 SKILL.md 文件链接");
        }
        const tail = parts.slice(3);
        if (marker === "blob") {
          if (tail[tail.length - 1] !== "SKILL.md") fail("文件链接需要指向 SKILL.md");
          tail.pop();
        }
        if (!tail.length) fail("链接中缺少分支或版本，请重新复制链接");
        if (override === ".") {
          ref = tail.join("/");
        } else if (override) {
          const suffix = override.split("/");
          const match = tail.length > suffix.length && suffix.every((part, i) => part === tail[tail.length - suffix.length + i]);
          if (match) ref = tail.slice(0, -suffix.length).join("/");
          else if (tail.length === 1 && marker === "tree") ref = tail[0];
          else fail("填写的目录与链接不一致，请核对仓库中的完整目录");
        } else {
          ref = tail[0];
          path = tail.slice(1).join("/");
          if (!path && marker === "blob") path = ".";
          if (!path) fail("请填写具体 Skill 目录；根目录就是 Skill 时填 .");
        }
      }
      const url = `https://github.com/${repo}` + (ref ? `/tree/${encodePath(ref)}${path === "." ? "" : `/${encodePath(path)}`}` : "");
      return {
        url, key: `github:${repo}:${path}`, repo, path,
        skillName: path === "." ? repo.split("/")[1] : path.split("/").pop(), kind: "github", ref
      };
    }

    if (override) fail("只有 GitHub 链接需要填写仓库目录");
    return { url: parsed.href, key: `url:${parsed.href}`, repo: "", path: "", skillName: "", kind: "website", ref: "" };
  }

  function normalizeSource(rawUrl, explicitPath = "") {
    const { ref, ...source } = parseSource(rawUrl, explicitPath);
    return source;
  }

  const quoteIdentifier = value => JSON.stringify(value).replace(/[\u2028\u2029]/g, character => `\\u${character.charCodeAt(0).toString(16)}`);

  function buildInstallPrompt(source) {
    if (!source || typeof source !== "object") return "";
    const url = safeUrl(source.url || source.source_url);
    if (!url) return "";
    let normalized;
    try {
      // Revalidate database values as well as freshly submitted form values.
      normalized = parseSource(url, source.path ?? source.source_path ?? "");
    } catch (_) { return ""; }
    const canonicalUrl = safeUrl(normalized.url);
    if (!canonicalUrl) return "";
    // Keep identifiers on a single, quoted line. Remote prose and install notes
    // are never copied into the request as instructions.
    const title = typeof source.title === "string" ? source.title.trim() : "";
    const name = normalized.skillName || (title.length <= 120 && !/[\u0000-\u001f\u007f-\u009f\u2028\u2029]/.test(title) ? title : "");
    const lines = ["请帮我安装这个 Skill，供当前 AI 工具使用。", "", "链接", canonicalUrl];
    if (name) lines.push(`Skill 名称 ${quoteIdentifier(name)}`);
    if (normalized.kind === "github") lines.push(normalized.path === "." ? "Skill 位于仓库根目录" : `仓库目录 ${quoteIdentifier(normalized.path)}`);
    lines.push("", "安装前请阅读原项目的说明，确认当前工具是否支持，并检查需要的文件和依赖。可以安装的话，请完成安装并确认可用，再给我一个实际用法示例；如果当前工具不支持，请说明原因和可用的替代方法。");
    return lines.join("\n");
  }

  function sourceType(skill) {
    const value = typeof skill === "string" ? skill : skill?.source_type ?? skill?.sourceType ?? skill?.source;
    return value === "external" ? "external" : "upload";
  }

  function filterSkills(items, { q = "", category = "", role = "", source = "all", sort = "new" } = {}) {
    const query = String(q).trim().toLowerCase();
    const filtered = items.map((skill, index) => ({ skill, index })).filter(({ skill }) => {
      if (category && skill.category !== category) return false;
      if (source && source !== "all" && sourceType(skill) !== source) return false;
      const roles = normalizeRoles(skill.roles);
      if (role && (!ROLE_OPTIONS.includes(role) || (!roles.includes(role) && !roles.includes("通用")))) return false;
      const searchable = [getDisplayTitle(skill), skill.title, skill.description, skill.tags,
        skill.original_author, skill.originalAuthor,
        skill.author?.display_name, skill.source_url, skill.source_path, roles.join(" ")].filter(Boolean).join(" ").toLowerCase();
      return !query || searchable.includes(query);
    });
    const time = skill => Number.isFinite(Date.parse(skill.updated_at)) ? Date.parse(skill.updated_at) : 0;
    const downloads = skill => sourceType(skill) === "external" ? 0 : Math.max(0, Number(skill.downloads) || 0);
    filtered.sort((a, b) => (sort === "hot" ? downloads(b.skill) - downloads(a.skill) : time(b.skill) - time(a.skill)) || a.index - b.index);
    return filtered.map(({ skill }) => skill);
  }

  return { ROLE_OPTIONS, getDisplayTitle, suggestDisplayTitle, normalizeRoles, normalizeSource, buildInstallPrompt, safeUrl, sourceType, filterSkills };
});
