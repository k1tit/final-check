@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM Запуск для сервера: Waitress, доступ по сети. Коллегам давайте только ссылку http://ИМЯ-СЕРВЕРА:8765/
set REPORTS_WEB_SERVER=1
set REPORTS_WEB_HOST=0.0.0.0
if not defined REPORTS_WEB_PORT set REPORTS_WEB_PORT=8765
set REPORTS_WEB_NO_BROWSER=1
python run_checks_web.py
pause
