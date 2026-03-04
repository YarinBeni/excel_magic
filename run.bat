@echo off
cd /d "%~dp0"
py -m pip install --target "%~dp0packages" pandas openpyxl >nul 2>&1
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONPATH=%~dp0packages
py find_duplicates.py
echo.
pause
