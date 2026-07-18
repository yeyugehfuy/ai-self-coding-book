@echo off
setlocal
cd /d "%~dp0"
echo Personal Knowledge Hub sync
echo.
echo 如果你已经在 config.toml 中配置了同步入口，请运行统一 CLI。
echo 当前阶段也可以直接运行：
echo python tools\shimo_export\export_shimo_diary_to_word.py --url "你的石墨文章URL" --debug --headed
echo.
python -m personal_kb_backup.cli
pause
