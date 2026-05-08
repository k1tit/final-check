@echo off
setlocal
cd /d "%~dp0"
python "build_checks.py"
if errorlevel 1 (
  echo.
  echo Script failed. Press any key to close...
  pause >nul
  exit /b 1
)
echo.
echo Done. Press any key to close...
pause >nul
