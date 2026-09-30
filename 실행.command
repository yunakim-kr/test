#!/bin/bash
# 협의체 안건 도출 에이전트 (연습용) - Mac/Linux용 실행 파일
cd "$(dirname "$0")" || exit 1
export PYTHONUTF8=1 PYTHONIOENCODING=utf-8

finish() { echo; read -r -p "엔터를 누르면 닫힙니다." _; exit "${1:-0}"; }

PY=""
command -v python3 >/dev/null 2>&1 && python3 -c "import sys" >/dev/null 2>&1 && PY="python3"
[ -z "$PY" ] && { echo "[준비 필요] Python 3가 설치되어 있지 않습니다. https://www.python.org/downloads/"; finish 1; }
command -v node >/dev/null 2>&1 || { echo "[준비 필요] Node.js가 설치되어 있지 않습니다. https://nodejs.org/"; finish 1; }
if [ ! -f node_modules/docx/package.json ]; then
  echo "[처음 한 번만] Word 파일 생성 부품을 설치합니다. 인터넷 연결이 필요합니다."
  npm install docx --no-fund --no-audit || { echo "[멈춤] 부품 설치에 실패했습니다."; finish 1; }
fi

# 선별 - 확정 - 안건서 만들기는 프로그램이 화면 안내로 진행한다.
# 이미 실행 중이면 프로그램이 안내만 보이고 끝난다(중복 실행 방지).
"$PY" agenda_agent.py run
finish $?
