@echo off
chcp 65001 >nul
cd /d "%~dp0"
python new_access_pf_checks.py %*
