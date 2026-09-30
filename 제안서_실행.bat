@echo off
chcp 65001 >nul
cd /d "%~dp0"
python problem_proposal.py
echo.
echo ============================
echo  Done. Press any key to close this window.
echo ============================
pause >nul
