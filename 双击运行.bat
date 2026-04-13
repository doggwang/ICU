@echo off
chcp 65001 >nul
echo ==========================================
echo ICU 报告处理工具
echo ==========================================
echo.
cd /d "%~dp0"
python run.py
pause
