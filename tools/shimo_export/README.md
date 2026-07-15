# 石墨日记批量导出 Word 工具

这个目录提供一个 Playwright 自动化脚本，用于把石墨中的日记文章/文档批量导出为 Word 或 PDF 文件并保存到本地文件夹。它既保留 `--url` / `--urls-file` 的导出方式，也支持直接传入一个石墨文件夹链接，自动递归读取子文件夹中的所有文档。

> 安全提醒：不要把石墨账号、密码写进仓库、聊天记录或命令行历史。推荐先用本地浏览器手动登录，脚本会复用 `.shimo-browser-profile/` 中保存的本地登录态。

## 安装依赖

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install playwright
python -m playwright install chromium
```

## 第一次登录或登录失效后重新登录

```bash
python tools/shimo_export/export_shimo_diary_to_word.py --login-only
```

在打开的浏览器里完成石墨登录后，回到终端按 Enter。登录态会保存在 `.shimo-browser-profile/` 中；该目录只应保留在本机，不要提交到 Git。

如果必须使用账号密码登录，请只通过环境变量传入：

```bash
export SHIMO_USERNAME='你的账号'
export SHIMO_PASSWORD='你的密码'
python tools/shimo_export/export_shimo_diary_to_word.py --login-only
```

## 推荐用法：导出整个石墨文件夹

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --folder "https://shimo.im/folder/xxxxxxxx"
```

脚本会：

- 自动打开石墨文件夹。
- 自动滚动加载文件列表。
- 尝试点击分页中的“下一页”。
- 自动递归进入子文件夹。
- 获取完整文档列表后再开始导出。
- 默认把文件导出到 `exported-shimo-diaries/`。
- 默认导出 `docx` 格式。

## 增量同步、断点续传与覆盖导出

默认启用 `--sync` 智能同步。脚本会把每篇文档的文档 ID、标题、石墨最后修改时间、最后导出时间和导出路径记录到 `exported-shimo-diaries/backup.json`。之后再次运行时，不再根据文件是否存在判断，而是对比石墨最后修改时间和本地备份索引：新增或修改过的文档会重新导出，未修改的文档会跳过。

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --folder "https://shimo.im/folder/xxxxxxxx"
```

如果需要忽略 `backup.json` 中的同步记录并重新覆盖导出，增加 `--force`：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --folder "https://shimo.im/folder/xxxxxxxx" \
  --force
```

## 导出格式

默认导出 Word：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --folder "https://shimo.im/folder/xxxxxxxx" \
  --format docx
```

导出 PDF：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --folder "https://shimo.im/folder/xxxxxxxx" \
  --format pdf
```

同时导出 Word 和 PDF：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --folder "https://shimo.im/folder/xxxxxxxx" \
  --format both
```

## 保留用法：导出单篇或 URL 列表

导出单篇：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --url "https://shimo.im/docs/xxxxxx"
```

创建一个文本文件，例如 `shimo-diary-urls.txt`，每行放一个石墨日记文章/文档链接：

```text
https://shimo.im/docs/xxxxxx
https://shimo.im/docs/yyyyyy
```

批量导出：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --urls-file shimo-diary-urls.txt \
  --output-dir exported-shimo-diaries
```

## 备份索引 backup.json

默认备份索引路径：

```text
exported-shimo-diaries/backup.json
```

其中会记录：

- 文档 ID。
- 标题。
- 石墨文件夹页读取到的最后修改时间。
- 最后导出时间。
- 各格式的导出路径。

如果要指定其他位置，可以使用：

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --folder "https://shimo.im/folder/xxxxxxxx" \
  --backup-db ./backup.json
```

## 进度与失败处理

导出时会显示实时进度，例如：

```text
发现 156 篇日记
导出格式：docx
总任务：156

[1/156]
正在导出：26.7.15
状态：new，格式：docx，剩余：155
✔ 完成：exported-shimo-diaries/26.7.15.docx
```

如果某篇导出失败，脚本会自动跳过并继续导出下一篇，最后输出新增、修改、跳过、失败数量和失败列表。

## 常用参数

- `--folder`：石墨文件夹 URL，可重复传入多个文件夹。
- `--url`：单篇石墨文档 URL，可重复传入多个文档。
- `--urls-file`：一行一个文档 URL 的文本文件。
- `--output-dir`：导出目录，默认 `exported-shimo-diaries/`。
- `--profile-dir`：浏览器登录态目录，默认 `.shimo-browser-profile/`。
- `--format`：导出格式，支持 `docx`、`pdf`、`both`，默认 `docx`。
- `--sync`：使用 `backup.json` 做增量同步；当前默认启用。
- `--backup-db`：指定备份索引 JSON 路径，默认 `<output-dir>/backup.json`。
- `--force`：忽略同步判断并重新导出。
- `--max-scrolls`：每个文件夹最多滚动次数，默认 `80`，用于兼容无限滚动加载。
- `--login-only`：只打开浏览器用于登录，不执行导出。

## 注意事项

- 当前账号必须对目标日记和子文件夹有查看和导出权限。
- 增量同步依赖石墨文件夹页能读取到最后修改时间；如果页面没有暴露最后修改时间，脚本会保守使用已有 `backup.json` 记录跳过未变化的文档，必要时可用 `--force` 重新导出。
- 脚本不能绕过团队或公司禁用导出的权限限制。
- 石墨网页的菜单文案或 DOM 结构可能变化；如果脚本提示找不到导出按钮，请手动打开一篇日记确认导出入口文案，再调整脚本中的 `MENU_TEXTS` 或 `EXPORT_PATTERNS`。
- 文件夹遍历依赖页面中的链接结构。如果石墨大幅改版，可能需要调整 `DOC_URL_PATTERN` 或 `FOLDER_URL_PATTERN`。
