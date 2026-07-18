@echo off
setlocal
cd /d "%~dp0"
echo Personal Knowledge Hub sync
echo.
echo 请选择同步方式：
echo 1. 第一次同步（全部下载）
echo 2. 同步新增（默认推荐）
echo 3. 同步最近7天
echo 4. 同步最近30天
echo 5. 指定日期范围
echo 6. 指定文章
echo 7. 强制重新同步全部
echo.
set /p choice=请输入编号（直接回车默认 2）：
if "%choice%"=="" set choice=2
set mode=new
if "%choice%"=="1" set mode=all
if "%choice%"=="2" set mode=new
if "%choice%"=="3" set mode=7days
if "%choice%"=="4" set mode=30days
if "%choice%"=="5" set mode=range
if "%choice%"=="6" set mode=single
if "%choice%"=="7" set mode=force
set extra=
if "%mode%"=="range" (
  set /p start=开始日期（例如 26.7.01）：
  set /p end=结束日期（例如 26.7.15）：
  set extra=--start "%start%" --end "%end%"
)
if "%mode%"=="single" (
  set /p one_url=请输入石墨文章 URL：
  set extra=--url "%one_url%"
)
echo.
echo 当前模式：%mode%
echo 如果需要批量同步，请把 URL 一行一个保存到 shimo_urls.txt。
if exist shimo_urls.txt (
  python tools\shimo_export\export_shimo_diary_to_word.py --url-file shimo_urls.txt --mode %mode% %extra% --debug
) else (
  python tools\shimo_export\export_shimo_diary_to_word.py --mode %mode% %extra% --debug
)
pause
