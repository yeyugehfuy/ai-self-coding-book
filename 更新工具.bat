@echo off
setlocal
cd /d "%~dp0"
echo Personal Knowledge Hub updater
echo.
if not exist ".git" (
  echo [ERROR] 当前目录不是 Git 仓库，暂时无法自动更新。
  echo 请从 GitHub 下载最新版，覆盖程序文件时保留 .shimo-browser-profile、config.toml、backup.json。
  pause
  exit /b 1
)
echo [1/4] 保留用户数据：.shimo-browser-profile config.toml backup.json backup-report.txt
echo [2/4] 检查远程更新...
git fetch --all --prune
if errorlevel 1 goto failed
echo [3/4] 拉取最新版本...
git pull --ff-only
if errorlevel 1 goto failed
echo [4/4] 更新完成。最近提交：
git log -1 --oneline
echo.
echo 如果你本地修改过程序文件导致 pull 失败，请先备份用户数据后重新下载最新版。
pause
exit /b 0
:failed
echo.
echo [ERROR] 自动更新失败。用户数据没有被删除。
echo 请保留 .shimo-browser-profile、config.toml、backup.json 后手动更新程序文件。
pause
exit /b 1
