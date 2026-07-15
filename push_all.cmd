@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Push via SSH (origin HTTPS ne rabotaet iz-za SSL)...
git push check_py master
if errorlevel 1 exit /b 1
git push git@github.com:k1tit/final-check.git master
if errorlevel 1 exit /b 1
git push git@github.com:k1tit/pr_ch.git master
if errorlevel 1 exit /b 1
echo OK: check_py, final-check, pr_ch
git log -1 --oneline
git ls-files new_access_pf_checks.py
