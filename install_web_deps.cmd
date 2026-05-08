@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Установка зависимостей веб-интерфейса (Flask, Waitress)...
python -m pip install -r "%~dp0requirements-web.txt"
if errorlevel 1 (
  echo.
  echo Ошибка. Проверьте, что Python в PATH.
  pause
  exit /b 1
)
echo.
echo Готово. Можно закрыть окно и запускать «Запуск_отчётов_веб.vbs».
timeout /t 5
exit /b 0
