@echo off
setlocal
cd /d "%~dp0"
python -m personal_kb_backup.cli
pause
