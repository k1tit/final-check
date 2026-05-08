@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM Тот же python, что в install_web_deps.cmd (не pythonw — у него часто другой набор пакетов).
echo ===== %DATE% %TIME% =====>> "%~dp0run_checks_web_startup.log"
python "%~dp0run_checks_web.py" >> "%~dp0run_checks_web_startup.log" 2>&1
echo ----- exit %ERRORLEVEL% ----- >> "%~dp0run_checks_web_startup.log"
