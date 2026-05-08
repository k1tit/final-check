@echo off
chcp 65001 >nul
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Нужны права администратора.
  pause
  exit /b 1
)
cd /d "%~dp0"

echo Остановка службы...
net stop PFBPYYZYWeb 2>nul

echo Удаление из реестра служб...
python "%~dp0pf_checks_windows_service.py" remove
if errorlevel 1 (
  echo Если remove не сработал: sc delete PFBPYYZYWeb
)

echo Готово.
pause
