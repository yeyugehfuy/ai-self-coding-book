# 第四阶段：个人知识库备份工具模块架构

目标是把当前项目从「单一导出脚本」整理成可扩展的个人知识资产同步中心。未来新增 Notion、飞书文档、Markdown、Obsidian、HTML、Word、PDF 等平台或格式时，应优先增加插件模块，而不是重写主流程。

## 终极目标

用户最终只需要双击：

```text
同步我的知识库.bat
```

工具自动执行：

```text
读取配置
↓
登录（如果需要）
↓
同步石墨 / Notion / 飞书等平台
↓
检查修改
↓
只导出新增或已变更内容
↓
生成 Word / Markdown / HTML
↓
更新 PDF（可选）
↓
生成备份报告
↓
完成
```

## 分层结构

```text
personal_kb_backup/
├── core/          # 主流程、插件协议、任务编排、增量同步边界
├── config/        # 配置读取、账号与工作区配置、输出路径配置
├── auth/          # 登录、Cookie、Token、会话刷新
├── browser/       # 浏览器自动化能力，例如 Playwright/Selenium 封装
├── platforms/     # 输入平台连接器：石墨、Notion、飞书、浏览器网页等
├── formats/       # 输出格式：Word、PDF、Markdown、HTML、Obsidian 等
├── exporters/     # 批量导出策略、文件命名、目录组织
└── reporting/     # 备份报告、错误摘要、变更统计
```

## 模块职责

| 模块 | 负责 | 不负责 |
| --- | --- | --- |
| `core` | 定义统一协议、编排同步流程 | 具体平台登录细节 |
| `config` | 读取配置、校验路径和开关 | 执行导出 |
| `auth` | 登录、保存会话、刷新凭据 | 解析文档内容 |
| `browser` | 浏览器启动、页面操作、下载监听 | 平台业务规则 |
| `platforms` | 拉取源文档并规范化为中间格式 | 生成 Word/PDF |
| `formats` | 将中间格式转换成目标文件 | 登录平台 |
| `exporters` | 批量导出、命名规则、跳过策略 | 直接抓取网页 |
| `reporting` | 生成同步结果报告 | 决定平台 API |

## 插件接口

第四阶段先固定三类扩展点：

1. `AuthProvider`：需要登录的平台实现自己的登录方式。
2. `PlatformConnector`：每个平台只负责「列出文档」和「获取文档」。
3. `FormatExporter`：每种输出格式只负责把标准文档内容写成目标文件。

核心数据流统一为：

```text
PlatformConnector -> DocumentContent(markdown + assets + metadata) -> FormatExporter
```

这样可以做到：

- 新增 Notion：只增加 `platforms/notion/connector.py`。
- 新增飞书：只增加 `platforms/feishu/connector.py`。
- 新增 PDF：只增加或替换 `formats/pdf/exporter.py`。
- 新增 Obsidian：只增加 `formats/obsidian/exporter.py`。

## 当前阶段目录落地

已建立以下骨架：

```text
personal_kb_backup/
├── core/contracts.py
├── core/pipeline.py
├── platforms/shimo/connector.py
├── platforms/notion/
├── platforms/feishu/
├── formats/word/
├── formats/pdf/
├── formats/markdown/
├── formats/html/
├── formats/obsidian/
├── auth/
├── browser/
├── config/
├── exporters/
└── reporting/
```

## 后续开发顺序建议

1. 配置文件：已提供 `config.example.toml`，后续接入正式配置读取与校验。
2. 登录模块：先支持石墨 Cookie/浏览器登录状态复用。
3. 石墨连接器：列出文档、读取更新时间、下载/导出内容。
4. 增量同步：记录文档 ID、更新时间、内容哈希，只处理新增或变更。
5. Word 导出器：生成 `.docx`。
6. PDF 导出器：作为可选步骤从 Word/HTML 生成。
7. 报告模块：输出本次新增、更新、跳过、失败的清单。
8. Windows 双击入口：已提供 `同步我的知识库.bat`，内部调用统一 CLI，占位逻辑后续替换为真实同步。

## GitHub 账号切换建议

如果需要换另一个 GitHub 账号继续维护，推荐用 GitHub Desktop 完成登录切换，然后重新克隆或重新绑定当前仓库远程地址。代码结构模块化后，后续升级通常只需要改某个平台或格式模块，再通过 GitHub Desktop 提交和同步即可。

## 石墨正文等待与调试策略

`tools/shimo_export/export_shimo_diary_to_word.py` 不再依赖旧版石墨页面的固定按钮或工具栏。脚本会先检查正文候选节点是否已经可见；如果正文已出现，就立即导出。否则再等待 `networkidle`，并继续轮询新版正文容器候选 selector。

调试时可追加：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py --url "https://shimo.im/docs/K0GYR3gFakvm4ZLG/" --debug
```

如果仍然超时，脚本会打印当前 URL、页面标题，并保存：

```text
debug/timeout.png
debug/timeout.html
```

## 一键同步与自动更新规划

项目根目录规划保留两个面向普通用户的入口：

```text
同步我的知识库.bat
更新工具.bat
```

`同步我的知识库.bat` 最终负责自动登录、检查新增内容、同步石墨、导出 Word、可选导出 PDF、更新 `backup.json`、生成 `backup-report.txt`，并输出本次同步统计。

`更新工具.bat` 负责自动拉取或下载最新程序，同时保留 `.shimo-browser-profile`、`config.toml`、`backup.json`、`backup-report.txt`、`exports/` 等用户数据和登录状态。当前实现优先支持 Git 仓库内 `git fetch` + `git pull --ff-only` 更新；如果用户不是从 Git 克隆的目录运行，会提示手动下载新版并保留用户数据。

后续主流程仍保持统一管线：

```text
读取配置 -> 登录 -> 同步 -> 增量检查 -> 导出 -> 生成报告 -> 完成
```

## 官方 Word 下载策略

石墨导出不再解析网页源码，也不再尝试自己拼 Word。新的最高优先级策略是模拟人工操作：打开文章，等待页面可操作，点击右上角更多菜单，点击下载/导出，选择 Word，等待浏览器下载完成，然后按文章标题重命名并移动到 `exports/YYYY/MM/`。

同步数据库 `backup.json` 升级为版本 2，按 URL 保存标题、本地文件名、本地路径、下载时间、石墨最后修改时间、可选文件 Hash、导出格式。默认同步模式是 `new`：只有新增文章，或检测到石墨最后修改时间变化的文章，才会重新下载；已下载且未修改的文章会跳过。

`同步我的知识库.bat` 现在先进入菜单，支持第一次同步、同步新增、最近 7 天、最近 30 天、指定日期范围、指定文章、强制重新同步全部。指定日期范围会根据文章标题中的 `26.7.01` / `26.7.15` 这类日期筛选。

### 菜单定位调试

石墨顶部菜单不再依赖 `text=更多`。脚本优先使用 `aria-haspopup`、`aria-expanded`、`aria-label`、`data-testid`、`button:has(svg)`、`[role=button]:has(svg)` 等更稳定的按钮结构；如果仍找不到，会回退到页面右上角可见按钮，并在失败时打印所有可见按钮/菜单项的 `tag`、`role`、`aria-label`、`title`、`text`、`class` 和坐标。

可单独运行按钮诊断：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py --url "https://shimo.im/docs/K0GYR3gFakvm4ZLG/" --debug --dump-buttons --headed
```
