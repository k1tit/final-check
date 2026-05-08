@echo off
setlocal
cd /d "%~dp0"
REM Для коллег без терминала: install_web_deps.cmd (один раз), затем Запуск_отчётов_веб.vbs
python run_checks_web.py
if errorlevel 1 pause
