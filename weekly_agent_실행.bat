@echo off
chcp 65001 >nul
cd /d "%~dp0"
python weekly_agent.py
echo.
echo ============================
echo  Done. Press any key to close this window.
echo ============================
pause >nul
