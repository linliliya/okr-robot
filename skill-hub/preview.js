// Preview data stays in this browser tab. This adapter never calls Supabase.
(() => {
  if (new URLSearchParams(location.search).get("preview") !== "1") return;

  const KEY = "skillhub.external-sharing.preview.v1";
  const USER = { id: "demo-user", email: "preview@example.invalid" };
  const ROLE_OPTIONS = ["通用", "内容创作", "课程运营", "销售", "教学教研", "人力行政", "产品经理", "设计师", "剪辑"];
  const REMOVED_ROLES = new Set(["客户服务", "研发工程师", "财务", "管理者"]);
  const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
  const copy = value => JSON.parse(JSON.stringify(value));
  const same = (a, b) => a === b || (a != null && b != null && String(a) === String(b));
  const ok = data => ({ data, error: null });
  const fail = (message, code = "PREVIEW_ERROR") => ({ data: null, error: { message, code } });
  const timestamp = () => new Date().toISOString();

  function initialState() {
    const now = timestamp();
    const common = {
      author_id: USER.id, downloads: 0, created_at: now, updated_at: now,
      source_url: "", source_key: null, source_repo: "", source_path: "",
      original_author: "", recommendation: "", tested: false,
      tested_with: "", tested_version: "", install_notes: "", roles: [],
    };
    const skills = [
      {
        ...common, id: 1, slug: "web-design-guidelines", source_type: "external",
        title: "Web Design Guidelines", category: "效率工具", tags: "网页,设计检查",
        description: "Vercel 提供的网页设计检查 Skill。这里展示外部推荐的页面效果，具体能力以原项目为准。",
        source_url: "https://github.com/vercel-labs/agent-skills/tree/main/skills/web-design-guidelines",
        source_key: "github:vercel-labs/agent-skills:skills/web-design-guidelines",
        source_repo: "vercel-labs/agent-skills", source_path: "skills/web-design-guidelines",
        original_author: "Vercel",
        roles: ["设计师", "产品经理"],
        recommendation: "示例推荐理由。做活动页、报名页时，可以先看看这份网页检查指南。这个示例尚未记录实际试用结果。",
        install_notes: "先阅读原项目的说明，确认它适合你使用的工具。这里仅用于预览推荐和安装指引。",
      },
      {
        ...common, id: 2, slug: "anthropic-pdf", source_type: "external",
        title: "PDF", category: "数据处理", tags: "PDF,文档",
        description: "Anthropic 提供的 PDF 处理 Skill。这里展示外部推荐的页面效果，具体能力以原项目为准。",
        source_url: "https://github.com/anthropics/skills/tree/main/skills/pdf",
        source_key: "github:anthropics/skills:skills/pdf", source_repo: "anthropics/skills",
        source_path: "skills/pdf", original_author: "Anthropic",
        roles: ["通用"],
        recommendation: "示例推荐理由。经常整理 PDF 资料的同事可以从这个项目开始了解。尚未试用，欢迎实际使用后补充适合的场景。",
        install_notes: "安装前查看原项目列出的运行环境和依赖，按自己的工具完成配置。",
      },
      {
        ...common, id: 3, slug: "course-outline-example", source_type: "upload",
        title: "课程大纲助手（示例）", category: "课程制作", tags: "大纲,备课",
        roles: ["教学教研"],
        description: "一个可下载的内部上传示例，根据课程目标整理大纲。用于体验原有上传和版本流程。",
      },
      {
        ...common, id: 4, slug: "weekly-review-example", source_type: "upload",
        title: "周报整理助手（示例）", category: "文案生成", tags: "周报,工作整理",
        roles: ["通用"],
        description: "一个可下载的内部上传示例，将已有工作记录整理成周报，不补写未经提供的进展。",
      },
    ];
    const docs = [
      "---\nname: course-outline-example\ndescription: 根据用户提供的课程目标整理教学大纲。这是 Skill Hub 预览示例。\n---\n\n# 课程大纲助手（示例）\n\n这是用于体验 Skill Hub 的示例文件。\n\n先请用户提供授课对象、课程目标和课时。材料齐全后，按课时整理主题、练习和预期成果。缺少材料时指出需要补充的部分，不编造已有课程内容。\n",
      "---\nname: weekly-review-example\ndescription: 将已有工作记录整理成简明周报。这是 Skill Hub 预览示例。\n---\n\n# 周报整理助手（示例）\n\n这是用于体验 Skill Hub 的示例文件。\n\n请用户提供本周工作记录。按已完成、进行中和需要协助的事项整理，每件事保留具体进展。只使用用户提供的材料，不补写数据或成果。\n",
    ];
    return {
      schema: 3, signedIn: true, nextId: 100,
      profiles: [{ id: USER.id, display_name: "体验账号", department: "产品", created_at: now }],
      skills,
      versions: [3, 4].map((skillId, index) => ({
        id: index + 1, skill_id: skillId, version: 1,
        changelog: "预览示例，可以下载体验", uploader_id: USER.id, uploaded_at: now,
        file_path: `preview-seed/${skills[skillId - 1].slug}/v1.zip`, skill_md: docs[index],
      })),
      files: {},
    };
  }

  let state = initialState();
  let persistent = true;
  try {
    const saved = JSON.parse(sessionStorage.getItem(KEY) || "null");
    if ([1, 2, 3].includes(saved?.schema) && Array.isArray(saved.skills) && Array.isArray(saved.versions)
        && Array.isArray(saved.profiles) && saved.files && Number.isFinite(saved.nextId)) {
      // Keep existing preview edits and files. Only known seed rows receive example roles.
      const seeds = state.skills;
      saved.skills = saved.skills.map(skill => {
        if (saved.schema < 3 && Array.isArray(skill.roles)) {
          const roles = [...new Set(skill.roles.map(role => role === "市场运营" ? "课程运营" : role).filter(role => !REMOVED_ROLES.has(role)))];
          return { ...skill, roles };
        }
        if (Object.prototype.hasOwnProperty.call(skill, "roles")) return skill;
        const seed = saved.schema === 1 && seeds.find(item => same(item.id, skill.id)
          && item.slug === skill.slug && item.source_key === skill.source_key);
        return { ...skill, roles: seed ? copy(seed.roles) : [] };
      });
      saved.schema = 3;
      state = saved;
    }
    sessionStorage.setItem(KEY, JSON.stringify(state));
  } catch (_) {
    persistent = false;
    window.SKILLHUB_PREVIEW_NOTICE = "浏览器暂时不能保存预览，刷新后会恢复示例。";
  }

  function commit(next) {
    if (persistent) {
      try { sessionStorage.setItem(KEY, JSON.stringify(next)); }
      catch (_) {
        return fail("预览存储空间不足，请删除部分预览文件或选择更小的 ZIP。", "PREVIEW_QUOTA");
      }
    }
    state = next;
    return null;
  }

  function relation(row, table) {
    const value = copy(row);
    if (table === "skills") value.author = copy(state.profiles.find(p => same(p.id, row.author_id)) || null);
    if (table === "versions") value.uploader = copy(state.profiles.find(p => same(p.id, row.uploader_id)) || null);
    return value;
  }

  function project(row, table, columns) {
    const value = relation(row, table);
    if (!columns || columns.split(",").some(part => part.trim() === "*")) return value;
    const fields = columns.split(/,(?![^()]*\))/).map(part => part.trim());
    return Object.fromEntries(fields.map(field => {
      const key = field.split(/[:(]/)[0];
      return [key, value[key]];
    }));
  }

  function validation(next, table) {
    if (table === "skills") {
      const slugs = new Set(), sources = new Set();
      for (const skill of next.skills) {
        if (!Array.isArray(skill.roles) || skill.roles.length > ROLE_OPTIONS.length
            || !skill.roles.every(role => ROLE_OPTIONS.includes(role))
            || (skill.roles.includes("通用") && skill.roles.length !== 1)) {
          return fail("请选择有效的适用岗位。「通用」不能和具体岗位一起选择。", "23514");
        }
        if (slugs.has(skill.slug)) return fail("已有同名条目，请换一个名称。", "23505");
        slugs.add(skill.slug);
        if (skill.source_type === "external" && skill.source_key) {
          if (sources.has(skill.source_key)) return fail("这个 Skill 已经有人推荐过了。", "23505");
          sources.add(skill.source_key);
        }
      }
    }
    if (table === "versions") {
      const seen = new Set();
      for (const version of next.versions) {
        if (!next.skills.some(skill => same(skill.id, version.skill_id))) return fail("找不到对应的 Skill。", "23503");
        const key = `${version.skill_id}:${version.version}`;
        if (seen.has(key)) return fail("这个版本已经存在。", "23505");
        seen.add(key);
      }
    }
    return null;
  }

  class Query {
    constructor(table) {
      this.table = table;
      this.operation = "select";
      this.filters = [];
      this.orders = [];
      this.columns = "*";
      this.returning = true;
      this.one = false;
      this.required = false;
      this.maximum = Infinity;
    }
    select(columns = "*") { this.columns = columns; this.returning = true; return this; }
    eq(field, value) { this.filters.push(row => same(row[field], value)); return this; }
    neq(field, value) { this.filters.push(row => !same(row[field], value)); return this; }
    is(field, value) { this.filters.push(row => row[field] === value); return this; }
    in(field, values) { this.filters.push(row => values.some(value => same(row[field], value))); return this; }
    order(field, options = {}) { this.orders.push({ field, ascending: options.ascending !== false }); return this; }
    limit(value) { this.maximum = Math.max(0, Number(value) || 0); return this; }
    maybeSingle() { this.one = true; return this; }
    single() { this.one = true; this.required = true; return this; }
    insert(value) { this.operation = "insert"; this.payload = copy(value); this.returning = false; return this; }
    update(value) { this.operation = "update"; this.payload = copy(value); this.returning = false; return this; }
    delete() { this.operation = "delete"; this.returning = false; return this; }
    then(resolve, reject) {
      if (!this.result) this.result = Promise.resolve().then(() => this.execute());
      return this.result.then(resolve, reject);
    }
    catch(reject) { return this.then(undefined, reject); }
    execute() {
      if (!["profiles", "skills", "versions"].includes(this.table)) return fail("预览中没有这张数据表。");
      if (!state.signedIn) return fail("请先进入预览账号。", "42501");
      const matches = row => this.filters.every(filter => filter(row));
      let rows;
      if (this.operation === "select") rows = state[this.table].filter(matches);
      else {
        const next = copy(state);
        if (this.operation === "insert") {
          const values = Array.isArray(this.payload) ? this.payload : [this.payload];
          rows = values.map(value => {
            const now = timestamp();
            const defaults = this.table === "skills"
              ? { source_type: "upload", downloads: 0, created_at: now, updated_at: now, tested: false, roles: [] }
              : this.table === "versions" ? { uploaded_at: now } : { created_at: now };
            return { ...defaults, id: next.nextId++, ...value };
          });
          next[this.table].push(...rows);
        } else if (this.operation === "update") {
          rows = [];
          next[this.table] = next[this.table].map(row => {
            if (!matches(row)) return row;
            const updated = { ...row, ...this.payload };
            rows.push(updated);
            return updated;
          });
        } else {
          rows = next[this.table].filter(matches);
          next[this.table] = next[this.table].filter(row => !matches(row));
          if (this.table === "skills") {
            next.versions = next.versions.filter(version => !rows.some(row => same(row.id, version.skill_id)));
          }
        }
        const invalid = validation(next, this.table);
        if (invalid) return invalid;
        const error = commit(next);
        if (error) return error;
      }
      rows = rows.slice().sort((a, b) => {
        for (const { field, ascending } of this.orders) {
          if (a[field] === b[field]) continue;
          const order = a[field] < b[field] ? -1 : 1;
          return ascending ? order : -order;
        }
        return 0;
      }).slice(0, this.maximum);
      if (!this.returning) return ok(null);
      const data = rows.map(row => project(row, this.table, this.columns));
      if (this.one) {
        if (data.length > 1 || (this.required && data.length === 0)) return fail("查询结果无法返回单条记录。", "PGRST116");
        return ok(data[0] || null);
      }
      return ok(data);
    }
  }

  function encode(bytes) {
    let binary = "";
    for (let i = 0; i < bytes.length; i += 32768) binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(binary);
  }

  const storage = {
    from(bucket) {
      const allowed = () => bucket === "skills" && state.signedIn;
      return {
        async upload(path, file, options = {}) {
          if (!allowed()) return fail("无法在预览中访问这个文件目录。");
          if (!file || typeof file.arrayBuffer !== "function") return fail("请选择一个有效的 ZIP 文件。");
          if (file.size > MAX_UPLOAD_BYTES) return fail("预览模式支持 5MB 以内的 ZIP。正式上传上限为 50MB。");
          if (state.files[path] && !options.upsert) return fail("这个文件路径已经存在。", "23505");
          try {
            const entry = { base64: encode(new Uint8Array(await file.arrayBuffer())), type: file.type || "application/zip" };
            const next = copy(state);
            next.files[path] = entry;
            return commit(next) || ok({ path, fullPath: `skills/${path}` });
          } catch (_) { return fail("预览文件保存失败，请尝试更小的 ZIP。"); }
        },
        async download(path) {
          if (!allowed()) return fail("无法在预览中访问这个文件目录。");
          const entry = state.files[path];
          if (entry) {
            const bytes = Uint8Array.from(atob(entry.base64), char => char.charCodeAt(0));
            return ok(new Blob([bytes], { type: entry.type }));
          }
          const version = state.versions.find(item => item.file_path === path && path.startsWith("preview-seed/"));
          const skill = version && state.skills.find(item => same(item.id, version.skill_id));
          if (!skill) return fail("这个预览文件已经被删除。");
          if (!window.JSZip) return fail("ZIP 组件未加载，请刷新页面后重试。");
          const zip = new window.JSZip();
          zip.file(`${skill.slug}/SKILL.md`, version.skill_md);
          return ok(await zip.generateAsync({ type: "blob", mimeType: "application/zip" }));
        },
        async remove(paths) {
          if (!allowed()) return fail("无法在预览中访问这个文件目录。");
          const next = copy(state);
          paths.forEach(path => { delete next.files[path]; });
          return commit(next) || ok(paths.map(name => ({ name })));
        },
      };
    },
  };

  const enter = async () => {
    const next = copy(state);
    next.signedIn = true;
    return commit(next) || ok({ user: copy(USER), session: { user: copy(USER) } });
  };

  window.SKILLHUB_PREVIEW_CLIENT = {
    auth: {
      async getSession() { return ok({ session: state.signedIn ? { user: copy(USER) } : null }); },
      async signOut() {
        const next = copy(state);
        next.signedIn = false;
        return commit(next) || { error: null };
      },
      signInWithPassword: enter,
      signUp: enter,
    },
    from: table => new Query(table),
    storage,
    async rpc(name, args = {}) {
      if (name !== "increment_downloads") return fail("预览不支持这个操作。");
      if (!state.signedIn) return fail("请先进入预览账号。", "42501");
      const next = copy(state);
      const skill = next.skills.find(item => same(item.id, args.p_skill_id));
      if (skill) skill.downloads = (skill.downloads || 0) + 1;
      return commit(next) || ok(null);
    },
  };

  window.resetSkillHubPreview = () => {
    try { sessionStorage.removeItem(KEY); } catch (_) { /* Restricted browser storage. */ }
    location.hash = "#/";
    location.reload();
  };
})();
