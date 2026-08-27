# Skill 市集 · 上线步骤（约 5-10 分钟）

页面代码已经写好，push 到 GitHub 就会自动部署到
`https://okr-internal.aurange.cn/skill-hub/`。
但平台需要一个免费的云端数据库来存账号和 skill 文件，按下面步骤配置一次即可。

## 第 1 步：注册 Supabase（免费）

1. 打开 https://supabase.com ，用 GitHub 账号或邮箱注册
2. 点 **New project** 创建项目：
   - Name 随便填，比如 `skill-hub`
   - Database Password 设置一个并**保存好**（后面基本用不到，但丢了麻烦）
   - Region 选 **Northeast Asia (Tokyo)** 或 **Southeast Asia (Singapore)**，国内访问快一些
3. 等 1-2 分钟项目初始化完成

## 第 2 步：初始化数据库

1. 左侧菜单点 **SQL Editor** → **New query**
2. 打开本目录下的 `setup.sql`，全选复制，粘贴进去
3. 点右下角 **Run**，看到 `Success. No rows returned` 就成功了

## 第 3 步：拿到项目地址和密钥

Supabase 2025 年改版后密钥页面搬了位置，按新版界面操作：

1. **项目地址（Project URL）**：项目主页顶部点 **Connect** 按钮，
   弹窗里就能看到（形如 `https://xxxx.supabase.co`）。
   或者：左侧 **Project Settings**（齿轮图标）→ **Data API** 也有。
2. **密钥**：左侧 **Project Settings** → **API Keys**：
   - 看到 **Publishable key**（`sb_publishable_` 开头）就复制它；
   - 如果没有，点 **Create new API keys** 生成一个再复制；
   - 老项目也可以切到 **Legacy API Keys** 标签页复制 **anon public**
     key（很长的一串）——两种密钥都能用，效果一样。
3. 打开本目录的 `config.js`，把这两个值分别填进
   `SUPABASE_URL` 和 `SUPABASE_ANON_KEY`

> Publishable key / anon key 是设计为可以公开放在网页里的，数据安全
> 由数据库的行级权限规则（setup.sql 里已配好）保证，不用担心泄露问题。

## 第 4 步：关闭邮箱验证（重要）

默认注册后要点邮件里的验证链接才能登录，对同事来说太麻烦：

1. 左侧菜单 **Authentication** → **Sign In / Providers** → **Email**
2. 把 **Confirm email** 开关关掉 → Save

## 第 5 步：发布

```bash
git add skill-hub && git commit -m 'feat: skill 共享平台' && git push
```

几分钟后访问 `https://okr-internal.aurange.cn/skill-hub/` 即可。

## 给同事的使用说明（可直接转发）

> 📍 平台地址：https://okr-internal.aurange.cn/skill-hub/
>
> **第一次用**：点「注册一个账号」，填姓名、部门、邮箱、密码即可。
>
> **下载别人的 skill**：找到想要的 skill → 点下载 → 解压 →
> 把整个文件夹放进电脑的 `~/.claude/skills/` 目录（访达按
> Cmd+Shift+G 输入这个路径就能打开）。
>
> **上传自己的 skill**：在电脑上找到自己的 skill 文件夹（在
> `~/.claude/skills/` 里），右键 →「压缩」得到 zip，然后在平台点
> 「+ 上传 Skill」传上去。名称和描述会自动识别，改进后再传一次
> 会自动变成新版本。

## 日常维护

- 免费额度：数据库 500MB + 文件存储 1GB，团队内部用绰绰有余
- **休眠机制**：免费版连续 7 天无访问会自动休眠（页面会登录失败/连不上）。
  仓库里的 GitHub Action（`.github/workflows/skillhub-keepalive.yml`）每天会
  自动访问一次防止休眠；万一还是休眠了，到 supabase.com 后台打开项目，
  点 **Restore project** 恢复即可，数据不会丢
- 想看谁注册了、删除违规内容：Supabase 后台 → Table Editor 直接改表
- 文件都存在 Storage → skills 桶里，可以直接在后台浏览
