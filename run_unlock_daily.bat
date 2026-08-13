@echo off
rem Daily generator for hk_unlock_overview.html (run by Windows Task Scheduler)
cd /d "%~dp0"
set "LOG_DIR=%~dp0logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
echo [%time%] start >> "%LOG_DIR%\unlock_daily.log"
"C:\Python314\python.exe" -X utf8 "%~dp0generate_unlock_html.py" --no-ccass >> "%LOG_DIR%\unlock_daily.log" 2>&1
echo [%time%] done >> "%LOG_DIR%\unlock_daily.log"
