@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Сборка портативной папки (PyInstaller). Нужен Python 3.10+ в PATH.
echo.
python -m pip install "pyinstaller>=6.0,<7"
if errorlevel 1 (
  echo Ошибка pip. Установите Python с python.org с галочкой Add to PATH.
  pause
  exit /b 1
)
echo.
echo Сборка dist\PF_BP_PY_ZY_Web (pandas/openpyxl — несколько минут)...
python -m PyInstaller --noconfirm --clean "%~dp0PF_BP_PY_ZY_Web.spec"
if errorlevel 1 (
  echo Ошибка PyInstaller. Если не хватает модулей — сначала: pip install -r requirements.txt -r requirements-web.txt
  pause
  exit /b 1
)
echo.
echo ======================================================================
echo Готово: папка  dist\PF_BP_PY_ZY_Web
echo.
echo Для пользователя: скопировать ВСЮ папку PF_BP_PY_ZY_Web на ПК.
echo Рядом с PF_BP_PY_ZY_Web.exe положить папку с нулевыми выгрузками:
echo   1 Нулевые файлы выгрузки макроса + файл исключений
echo   (внутри неё — 3801, 3802, ... с xlsx)
echo Отчёты появятся в подпапке  result  рядом с exe.
echo Запуск: двойной щелчок по PF_BP_PY_ZY_Web.exe — откройте в браузере http://127.0.0.1:8765/
echo ======================================================================
pause
