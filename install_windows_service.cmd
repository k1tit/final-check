@echo off
chcp 65001 >nul
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Нужны права администратора: ПКМ по файлу — «Запуск от имени администратора».
  pause
  exit /b 1
)
cd /d "%~dp0"

echo [1/4] Зависимости: Flask, Waitress, pandas, openpyxl, pywin32...
python -m pip install -r "%~dp0requirements-web.txt" -r "%~dp0requirements.txt" "pywin32>=306"
if errorlevel 1 (
  echo Ошибка pip. Убедитесь, что Python в PATH ^(python.org, галочка Add to PATH^).
  pause
  exit /b 1
)

echo [2/4] Регистрация службы PFBPYYZYWeb...
python "%~dp0pf_checks_windows_service.py" install
if errorlevel 1 (
  echo Ошибка install. Если служба уже есть — сначала uninstall_windows_service.cmd
  pause
  exit /b 1
)

echo [3/4] Автозапуск при загрузке Windows...
python "%~dp0pf_checks_windows_service.py" --startup auto update

echo [4/4] Перезапуск при сбое ^(если поддерживает ваша редакция Windows^)...
sc failure PFBPYYZYWeb reset= 86400 actions= restart/60000/restart/120000/restart/300000 >nul 2>&1

echo Запуск службы...
net start PFBPYYZYWeb
if errorlevel 1 (
  echo Служба создана, но не стартовала. См. run_checks_web_error.log и Журнал Windows.
  pause
  exit /b 1
)

echo.
echo ======================================================================
echo Служба запущена. Веб:  http://127.0.0.1:8765/
echo Остановка:  net stop PFBPYYZYWeb
echo Удаление:   uninstall_windows_service.cmd ^(от администратора^)
echo Подробности: Сервис_Windows.txt
echo ======================================================================
pause
