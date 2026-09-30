@echo off
title Agenda Agent (practice)
rem ------------------------------------------------------------------
rem  Double-click launcher (Windows).
rem  This file is ASCII only on purpose: Korean text in a batch file can
rem  break goto/labels or garble the screen depending on Windows settings.
rem  All Korean messages are printed by the Python program itself.
rem ------------------------------------------------------------------
pushd "%~dp0"

rem ---------- check Python ----------
set "PY="
py -3 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if defined PY goto :py_ok
python -c "import sys" >nul 2>nul
if not errorlevel 1 set "PY=python"
if defined PY goto :py_ok
echo.
echo [SETUP NEEDED] Python 3 is not installed.
echo   Install it from https://www.python.org/downloads/ and double-click this file again.
echo   On the first install screen, check "Add python.exe to PATH".
echo   Korean guide: see the usage text file in this folder.
goto :end
:py_ok

rem ---------- check Node.js ----------
node -v >nul 2>nul
if not errorlevel 1 goto :node_ok
echo.
echo [SETUP NEEDED] Node.js is not installed. It is needed to create the Word file.
echo   Install the LTS version from https://nodejs.org/ and double-click this file again.
goto :end
:node_ok

rem ---------- Word helper module (already bundled in node_modules) ----------
if exist "node_modules\docx\package.json" goto :mod_ok
echo.
echo [FIRST RUN] Installing the Word helper module. An internet connection is needed.
call npm install docx --no-fund --no-audit
if errorlevel 1 goto :fail_mod
:mod_ok

rem ---------- run the agent (Korean screen guide) ----------
rem If the agent is already running, the program shows a notice and exits.
%PY% agenda_agent.py run
goto :end

:fail_mod
echo.
echo [STOPPED] Module install failed. Check the internet connection and try again.

:end
echo.
pause
popd
