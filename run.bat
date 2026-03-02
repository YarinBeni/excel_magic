@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -m pip install pandas openpyxl
py find_duplicates.py
echo.
pause
