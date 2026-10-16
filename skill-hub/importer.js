(function (root, factory) {
  const shared = typeof module === "object" && module.exports ? require("./shared.js") : root.SkillHub;
  const api = factory(shared);
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.SkillHubImporter = api;
})(typeof window !== "undefined" ? window : null, function (H) {
  "use strict";

  const API = "https://api.github.com";
  const LIMITS = { timeout: 15000, requests: 24, file: 256 * 1024, response: 2 * 1024 * 1024, nameScan: 12, choices: 200 };
  const REPO = /^[a-z0-9][a-z0-9._-]*\/[a-z0-9._-]+$/i;
  const SLUG = /^[a-z0-9][a-z0-9._-]*$/i;
  const cacheByFetch = new WeakMap();
  const encodePath = value => value.split("/").map(encodeURIComponent).join("/");

  function problem(message, code = "IMPORT_FAILED") {
    const error = new Error(message);
    error.name = "SkillImportError";
    error.code = code;
    return error;
  }
  function cancelled() {
    const error = new Error("已取消读取");
    error.name = "AbortError";
    return error;
  }
  function parseUrl(raw) {
    const safe = H.safeUrl(raw);
    if (!safe) throw problem("请粘贴有效的 HTTPS 链接", "INVALID_URL");
    const url = new URL(safe);
    const host = url.hostname.toLowerCase().replace(/^www\./, "");
    if (url.port || !["github.com", "skills.sh"].includes(host)) {
      throw problem("暂时只能自动读取 GitHub 和 skills.sh 链接。这个网站可以继续手动填写。", "UNSUPPORTED_SOURCE");
    }
    const parts = url.pathname.replace(/^\/+|\/+$/g, "").split("/").map(decodeURIComponent);
    const repo = `${parts[0] || ""}/${(parts[1] || "").replace(/\.git$/i, "")}`.toLowerCase();
    if (!REPO.test(repo) || parts.some(part => !part || part === "." || part === "..")) {
      throw problem("仓库地址格式不正确，请重新复制链接", "INVALID_URL");
    }
    if (host === "skills.sh") {
      const source = H.normalizeSource(raw);
      return { repo, mode: "directory", slug: source.skillName, alias: source.key };
    }
    if (parts.length === 2) return { repo, mode: "repository" };
    if (!["tree", "blob"].includes(parts[2]) || parts.length < 4) {
      throw problem("请粘贴仓库首页、Skill 目录或 SKILL.md 文件链接", "INVALID_URL");
    }
    const tail = parts.slice(3);
    if (parts[2] === "blob") {
      if (tail.pop() !== "SKILL.md" || !tail.length) throw problem("文件链接需要指向 SKILL.md", "INVALID_URL");
    }
    return { repo, mode: parts[2], tail };
  }

  async function run(options, callback) {
    const fetchImpl = options?.fetchImpl || globalThis.fetch;
    if (typeof fetchImpl !== "function") throw problem("当前浏览器无法读取链接，请手动填写。", "UNSUPPORTED_BROWSER");
    const external = options?.signal;
    if (external?.aborted) throw cancelled();
    const controller = new AbortController();
    let stopError;
    let rejectStop;
    const stopped = new Promise((_, reject) => { rejectStop = reject; });
    function stop(error) {
      if (stopError) return;
      stopError = error;
      controller.abort();
      rejectStop(error);
    }
    const onAbort = () => stop(cancelled());
    external?.addEventListener("abort", onAbort, { once: true });
    const timer = setTimeout(() => stop(problem("读取超过 15 秒，请重试，或继续手动填写。", "TIMEOUT")), LIMITS.timeout);
    const context = { fetchImpl, signal: controller.signal, count: 0, check() { if (stopError) throw stopError; } };
    try { return await Promise.race([Promise.resolve().then(() => callback(context)), stopped]); }
    finally { clearTimeout(timer); external?.removeEventListener("abort", onAbort); }
  }

  async function boundedText(response, maxBytes) {
    const declared = Number(response.headers?.get("content-length"));
    if (declared > maxBytes) throw problem("仓库返回的资料太大，请改用具体 Skill 的目录链接，或手动填写。", "RESPONSE_TOO_LARGE");
    if (response.body?.getReader) {
      const reader = response.body.getReader();
      const chunks = [];
      let size = 0;
      try {
        while (true) {
          const { value, done } = await reader.read();
          if (done) break;
          size += value.byteLength;
          if (size > maxBytes) throw problem("仓库返回的资料太大，请改用具体 Skill 的目录链接，或手动填写。", "RESPONSE_TOO_LARGE");
          chunks.push(value);
        }
      } catch (error) { try { await reader.cancel(); } catch (_) {} throw error; }
      finally { reader.releaseLock(); }
      const bytes = new Uint8Array(size);
      let offset = 0;
      chunks.forEach(chunk => { bytes.set(chunk, offset); offset += chunk.byteLength; });
      return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    }
    const text = await response.text();
    if (new TextEncoder().encode(text).byteLength > maxBytes) throw problem("仓库返回的资料太大，请改用具体 Skill 的目录链接，或手动填写。", "RESPONSE_TOO_LARGE");
    return text;
  }

  async function request(context, path, maxBytes = LIMITS.response) {
    context.check();
    let cache = cacheByFetch.get(context.fetchImpl);
    if (!cache) { cache = new Map(); cacheByFetch.set(context.fetchImpl, cache); }
    const cached = cache.get(path);
    if (cached && cached.expires > Date.now()) return cached.data;
    if (++context.count > LIMITS.requests) throw problem("这个仓库需要读取的资料较多，请改用具体 Skill 的目录链接，或手动填写。", "REQUEST_LIMIT");
    let response;
    try {
      const fetchImpl = context.fetchImpl;
      response = await fetchImpl(API + path, {
        signal: context.signal, credentials: "omit", redirect: "error", referrerPolicy: "no-referrer",
        headers: { Accept: "application/vnd.github+json" }
      });
      context.check();
    } catch (error) {
      context.check();
      if (error.name === "AbortError") throw error;
      throw problem("暂时连不上 GitHub，请重试，或继续手动填写。", "NETWORK_ERROR");
    }
    if (response.status === 404) throw problem("没有找到公开的仓库、分支或文件。请检查链接，私有仓库可以手动填写。", "NOT_FOUND");
    if (response.status === 403 || response.status === 429) throw problem("GitHub 暂时限制了读取请求，或这个仓库没有公开访问权限。稍后重试，也可以手动填写。", "RATE_LIMIT");
    if (!response.ok) throw problem("GitHub 暂时无法返回资料，请稍后重试，或手动填写。", "REMOTE_ERROR");
    let data;
    try { data = JSON.parse(await boundedText(response, maxBytes)); }
    catch (error) {
      context.check();
      if (error.name === "SkillImportError") throw error;
      throw problem("GitHub 返回的资料无法读取，请重试，或手动填写。", "INVALID_RESPONSE");
    }
    context.check();
    if (cache.size >= 40) cache.delete(cache.keys().next().value);
    cache.set(path, { data, expires: Date.now() + 5 * 60 * 1000 });
    return data;
  }

  async function repository(context, repo) {
    const info = await request(context, `/repos/${repo}`);
    if (!info || typeof info.default_branch !== "string" || !info.default_branch.trim()) {
      throw problem("没有读到仓库的默认分支，请改用具体 Skill 的目录链接，或手动填写。", "NO_DEFAULT_BRANCH");
    }
    return info;
  }
  async function resolveRef(context, repo, tail) {
    // A slash can belong to a branch name. Verify the longest real ref first;
    // never silently treat its first segment as the branch.
    if (tail.length > 18) throw problem("链接层级太多，暂时无法确认分支和目录，请手动填写。", "REQUEST_LIMIT");
    for (let count = tail.length; count >= 1; count--) {
      const ref = tail.slice(0, count).join("/");
      try {
        const tree = await request(context, `/repos/${repo}/git/trees/${encodeURIComponent(ref)}`);
        if (!tree || !Array.isArray(tree.tree)) throw problem("没有读到仓库目录，请重试。", "INVALID_RESPONSE");
        return { ref, path: tail.slice(count).join("/") || ".", tree };
      } catch (error) { if (error.code !== "NOT_FOUND") throw error; }
    }
    throw problem("没有找到链接中的公开分支或版本，请检查链接，或手动填写。", "NOT_FOUND");
  }
  async function skillFiles(context, repo, ref, scope = ".") {
    const tree = await request(context, `/repos/${repo}/git/trees/${encodeURIComponent(ref)}?recursive=1`);
    if (!tree || !Array.isArray(tree.tree)) throw problem("没有读到仓库目录，请重试。", "INVALID_RESPONSE");
    if (tree.truncated) throw problem("GitHub 没有返回完整目录，暂时无法准确定位 Skill。请粘贴具体 Skill 的目录或 SKILL.md 链接。", "TREE_TRUNCATED");
    return tree.tree.filter(item => item.type === "blob" && item.mode !== "120000" && typeof item.path === "string" &&
      (item.path === "SKILL.md" || item.path.endsWith("/SKILL.md")) &&
      (scope === "." || item.path.startsWith(scope + "/")))
      .map(item => ({ path: item.path === "SKILL.md" ? "." : item.path.slice(0, -9), size: item.size }))
      .sort((a, b) => a.path.localeCompare(b.path));
  }
  function sourceFor(repo, ref, path) {
    return H.normalizeSource(`https://github.com/${repo}/tree/${encodePath(ref)}${path === "." ? "" : "/" + encodePath(path)}`, path);
  }
  function choicesFor(repo, ref, files, notice) {
    if (files.length > LIMITS.choices) throw problem("仓库里有超过 200 个 Skill，请粘贴具体 Skill 的目录链接。", "TOO_MANY_CHOICES");
    return { type: "choices", choices: files.map(file => ({
      label: file.path === "." ? `${repo.split("/")[1]}（根目录）` : file.path,
      path: file.path, url: sourceFor(repo, ref, file.path).url
    })), notice: notice || "这个仓库有多个 Skill，选一个想推荐的。" };
  }

  function unquote(value) {
    const text = value.trim();
    if (!text || /^[&*!\[{]/.test(text)) return "";
    if (text[0] === '"') {
      const match = text.match(/^("(?:[^"\\]|\\.)*")(?:\s+#.*)?\s*$/);
      if (!match) return "";
      try { return JSON.parse(match[1]); } catch (_) { return ""; }
    }
    if (text[0] === "'") {
      const match = text.match(/^'((?:[^']|'')*)'(?:\s+#.*)?\s*$/);
      return match ? match[1].replace(/''/g, "'") : "";
    }
    if (/^(?:null|~|true|false)$/i.test(text) || /^[-+]?\d+(?:\.\d+)?$/.test(text)) return "";
    return text.replace(/\s+#.*$/, "").trim();
  }
  function frontmatter(markdown) {
    const lines = markdown.replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n").split("\n");
    if (!/^---\s*$/.test(lines[0])) return {};
    const end = lines.findIndex((line, index) => index > 0 && /^(?:---|\.\.\.)\s*$/.test(line));
    if (end < 0) return {};
    const fields = {};
    for (let i = 1; i < end; i++) {
      const match = lines[i].match(/^(name|description|metadata):\s*(.*)$/);
      if (!match) continue;
      const key = match[1];
      let value = match[2];
      let next = i + 1;
      const continuation = [];
      while (next < end && (/^\s+\S/.test(lines[next]) || !lines[next].trim())) continuation.push(lines[next++]);
      if (key === "metadata") {
        const author = continuation.find(line => /^\s+author:\s*\S/.test(line));
        if (author) fields.author = unquote(author.replace(/^\s+author:\s*/, ""));
      } else if (/^[|>][+-]?(?:\s+#.*)?$/.test(value.trim())) {
        const nonempty = continuation.filter(line => line.trim());
        const indent = nonempty.length ? Math.min(...nonempty.map(line => line.match(/^\s*/)[0].length)) : 0;
        const block = continuation.map(line => line.slice(indent));
        value = value.trim()[0] === "|" ? block.join("\n") : block.join("\n").replace(/([^\n])\n(?=[^\n])/g, "$1 ");
        fields[key] = value.trim();
      } else {
        fields[key] = unquote(value);
        if (fields[key] && !/^["']/.test(value) && continuation.some(line => line.trim())) {
          // Plain YAML scalar continuation is folded; nested maps are not scalars.
          fields[key] = continuation.some(line => /^\s+\S+:\s/.test(line)) ? "" : [fields[key], ...continuation.map(line => line.trim())].join(" ").trim();
        }
      }
      i = next - 1;
    }
    return fields;
  }
  function categoryFor(title, description) {
    const text = `${title} ${description}`.toLowerCase();
    const rules = [
      ["招聘筛选", /招聘|简历|\b(?:resume|recruiting|recruitment|hiring)\b/],
      ["课程制作", /课程|课件|教学|\b(?:curriculum|lesson|teaching)\b/],
      ["数据处理", /数据处理|电子表格|\b(?:spreadsheet|xlsx|csv|data analysis|data processing)\b/],
      ["文案生成", /写作|文案|润色|\b(?:copywriting|writing|proofread)\b/],
      ["资讯采集", /资讯|新闻|爬虫|\b(?:news|scraper|scraping|web research)\b/],
      ["效率工具", /效率|自动化|\b(?:automation|productivity)\b/]
    ];
    return rules.find(([, regex]) => regex.test(text))?.[0] || "其他";
  }
  async function readSkill(context, repo, ref, path) {
    const filename = path === "." ? "SKILL.md" : `${path}/SKILL.md`;
    const data = await request(context, `/repos/${repo}/contents/${encodePath(filename)}?ref=${encodeURIComponent(ref)}`, 512 * 1024);
    if (!data || data.type !== "file" || typeof data.content !== "string" || data.encoding !== "base64") {
      throw problem("没有读到可用的 SKILL.md 文件，请检查链接，或手动填写。", "INVALID_SKILL_FILE");
    }
    const encoded = data.content.replace(/\s/g, "");
    if (Number(data.size) > LIMITS.file || encoded.length > Math.ceil(LIMITS.file / 3) * 4) {
      throw problem("这个 SKILL.md 超过 256 KB，暂时无法自动读取，请手动填写。", "FILE_TOO_LARGE");
    }
    let markdown;
    try {
      const binary = atob(encoded);
      if (binary.length > LIMITS.file) throw problem("这个 SKILL.md 超过 256 KB，暂时无法自动读取，请手动填写。", "FILE_TOO_LARGE");
      markdown = new TextDecoder("utf-8", { fatal: true }).decode(Uint8Array.from(binary, char => char.charCodeAt(0)));
    } catch (error) {
      if (error.code === "FILE_TOO_LARGE") throw error;
      throw problem("SKILL.md 的文字编码无法读取，请手动填写。", "INVALID_ENCODING");
    }
    const metadata = frontmatter(markdown);
    const source = sourceFor(repo, ref, path);
    const notices = [];
    const originalTitle = metadata.name || (path === "." ? repo.split("/")[1] : path.split("/").pop());
    if (!metadata.name) notices.push("没有读到名称，暂时使用目录名，可自行修改。");
    if (!metadata.description) notices.push("没有读到用途介绍，请补充几句。");
    if (originalTitle.length > 200 || (metadata.description || "").length > 4000) notices.push("原介绍较长，已截取表单能保存的部分，请检查。");
    const title = originalTitle.slice(0, 200);
    const description = (metadata.description || "").slice(0, 4000);
    const aliasKeys = [];
    if (metadata.name && SLUG.test(metadata.name)) aliasKeys.push(`skills:${repo}:${metadata.name}`);
    return { type: "skill", skill: {
      source, title, description, original_author: (metadata.author || repo.split("/")[0]).slice(0, 200),
      category: categoryFor(title, description), tags: "", aliasKeys: [...new Set(aliasKeys)]
    }, ...(notices.length ? { notice: notices.join(" ") } : {}) };
  }

  async function inspect(rawUrl, options = {}) {
    const input = parseUrl(rawUrl);
    return run(options, async context => {
      const info = await repository(context, input.repo);
      let ref = info.default_branch;
      let scope = ".";
      if (input.tail) {
        const resolved = await resolveRef(context, input.repo, input.tail);
        ref = resolved.ref;
        scope = resolved.path;
        try { return await readSkill(context, input.repo, ref, scope); }
        catch (error) { if (error.code !== "NOT_FOUND" || input.mode === "blob") throw error; }
      }
      const files = await skillFiles(context, input.repo, ref, scope);
      if (!files.length) throw problem("这个目录下没有找到 SKILL.md，请检查链接，或继续手动填写。", "NO_SKILLS");
      if (input.mode === "directory") {
        const namedPaths = files.filter(file => file.path.split("/").pop() === input.slug);
        // A directory can have a different frontmatter name. Do not assign a
        // skills.sh identity until the file itself confirms that name.
        const namedMatches = [];
        if (namedPaths.length > LIMITS.nameScan) throw problem("仓库里有较多同名目录，请改用具体 Skill 的 GitHub 目录链接。", "SKILL_NOT_LOCATED");
        for (const file of namedPaths) {
          const result = await readSkill(context, input.repo, ref, file.path);
          if (result.skill.aliasKeys.includes(input.alias)) namedMatches.push({ file, result });
        }
        if (namedMatches.length === 1) return namedMatches[0].result;
        if (namedMatches.length > 1) return choicesFor(input.repo, ref, namedMatches.map(match => match.file), "仓库里有多个 Skill 使用了这个名称，请选择具体目录。");
        if (files.length > LIMITS.nameScan) throw problem("没有找到同名目录，这个仓库的 Skill 较多。请从原项目复制具体 Skill 的目录链接，或手动填写。", "SKILL_NOT_LOCATED");
        const matches = [];
        for (const file of files.filter(file => !namedPaths.includes(file))) {
          const result = await readSkill(context, input.repo, ref, file.path);
          if (result.skill.aliasKeys.includes(input.alias)) matches.push({ file, result });
        }
        if (matches.length === 1) return matches[0].result;
        if (matches.length > 1) return choicesFor(input.repo, ref, matches.map(match => match.file), "仓库里有多个 Skill 使用了这个名称，请选择具体目录。");
        throw problem("没有在仓库中找到这个名称的 Skill。请核对 skills.sh 链接，或改用具体 Skill 的 GitHub 目录链接。", "SKILL_NOT_LOCATED");
      }
      if (files.length === 1) return readSkill(context, input.repo, ref, files[0].path);
      return choicesFor(input.repo, ref, files);
    });
  }
  async function readChoice(choice, options = {}) {
    if (!choice || typeof choice.url !== "string") throw problem("请重新选择一个 Skill", "INVALID_CHOICE");
    return inspect(choice.url, options);
  }
  return { inspect, readChoice };
});
