@echo off
chcp 65001 >nul
set PYTHONUTF8=1
title 주간보고 통합
cd /d "%~dp0"

echo ==== 주간보고 통합 시작 ====
echo.

if not exist "weekly_report.py" (
  echo [오류] weekly_report.py 파일이 이 파일과 같은 폴더에 없습니다.
  goto end
)

set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY (
  where python >nul 2>&1 && set "PY=python"
)
if not defined PY (
  echo [오류] Python이 설치되어 있지 않습니다. python.org에서 설치한 뒤 다시 실행해 주세요.
  echo 설치할 때 Add Python to PATH 항목을 꼭 체크해 주세요.
  goto end
)

%PY% -c "import anthropic" >nul 2>&1
if errorlevel 1 (
  echo anthropic 패키지를 처음 한 번 설치합니다. 잠시 기다려 주세요...
  %PY% -m pip install anthropic
  if errorlevel 1 (
    echo [오류] 패키지 설치에 실패했습니다. 인터넷 연결을 확인해 주세요.
    goto end
  )
)

if not defined ANTHROPIC_API_KEY (
  echo [오류] API 키가 등록되어 있지 않습니다.
  echo 명령 창에서 아래 형식으로 한 번 등록한 뒤 이 파일을 다시 실행해 주세요.
  echo   setx ANTHROPIC_API_KEY "발급받은키"
  goto end
)

%PY% weekly_report.py

:end
echo.
echo 아무 키나 누르면 창이 닫힙니다.
pause >nul
