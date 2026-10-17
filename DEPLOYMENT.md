# 网站维护与发布

内部工具继续使用 [internal.aurange.cn](https://internal.aurange.cn/)。

## AI 编程训练营排期

源代码已迁移到 [ai-camp-schedule](https://github.com/linliliya/ai-camp-schedule)。本仓库通过 Git submodule 引用它，目录仍为 `ai-camp-schedule/`。请在独立仓库创建开发分支并提交 Pull Request，由所有者合并到受保护的 `main`。

已有本地项目可以运行下面的命令初始化或更新该目录。

```sh
git submodule update --init --recursive
```

如果需要拉取 AI 排期最新的已合并版本，运行下面的命令。

```sh
git submodule update --remote ai-camp-schedule
```

## 统一发布

`Publish internal tools` 工作流在本仓库的 `main` 更新时运行，也会每 10 分钟检查一次 AI 排期的 `main`。GitHub 调度可能延迟。需要立即发布时，在本仓库 Actions 中手动运行该工作流。

发现 AI 排期有更新后，工作流会更新本仓库的 submodule 引用，再组装静态文件并发布到 GitHub Pages。线上版本记录在 `/deployment.json`。版本未变化时跳过部署，部署失败后会在后续检查中重试。

Pages 的发布方式为 GitHub Actions。自定义域名、各网站路径和 AI 排期的浏览器存储键保持不变。

发布脚本只复制 AI 排期的 `index.html` 和 `data.js`，不会执行独立仓库中的脚本。若网站以后增加图片、样式或其他静态文件，需要由所有者更新发布脚本中的文件清单。

## 权限

员工只需加入 AI 排期仓库。本仓库及其发布流程由所有者管理，不需要向员工提供写入权限或跨仓库凭证。

两个仓库当前均为公开仓库，本次隔离控制修改权限。其他网站仍在本仓库，后续可逐个迁移。旧版本仍保留在 Git 历史中。

本次迁移只涉及代码与发布，Supabase 数据、账号权限和 OKR 通知配置不变。
