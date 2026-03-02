@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -m pip install --target "%~dp0packages" pandas openpyxl
set PYTHONPATH=%~dp0packages
py find_duplicates.py
echo.
pause
