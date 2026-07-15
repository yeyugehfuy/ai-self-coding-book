# 石墨日记批量导出 Word 工具

这个目录提供一个 Playwright 自动化脚本，用于把石墨中的日记文章/文档逐篇导出为 Word 文件并保存到本地文件夹。

> 安全提醒：不要把石墨账号、密码写进仓库、聊天记录或命令行历史。推荐先用本地浏览器手动登录，脚本会复用本地登录态。

## 安装依赖

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install playwright
python -m playwright install chromium
```

## 第一次登录

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

## 准备日记链接列表

创建一个文本文件，例如 `shimo-diary-urls.txt`，每行放一个石墨日记文章/文档链接：

```text
https://shimo.im/docs/xxxxxx
https://shimo.im/docs/yyyyyy
```

## 批量导出为 Word

```bash
python tools/shimo_export/export_shimo_diary_to_word.py \
  --urls-file shimo-diary-urls.txt \
  --output-dir exported-shimo-diaries
```

导出的 `.doc`/`.docx` 文件会保存到 `exported-shimo-diaries/`。

## 注意事项

- 当前账号必须对目标日记有查看和导出权限。
- 石墨网页的菜单文案或 DOM 结构可能变化；如果脚本提示找不到 Word 导出按钮，请手动打开一篇日记确认导出入口文案，再调整脚本中的 `MENU_TEXTS` 或 `EXPORT_WORD_PATTERNS`。
- 若公司或团队禁用了导出功能，脚本无法绕过权限限制。
