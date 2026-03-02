@echo off
chcp 65001 >nul
cd /d "%~dp0"
pip install pandas openpyxl >nul 2>&1
python find_duplicates.py
echo.
pause
