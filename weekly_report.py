"""
팀 주간보고 통합 도우미 (첫 버전)

하는 일
1. 바탕화면 > 연습 폴더에서 '보고_'로 시작하는 .txt 파일을 찾아 읽습니다. (코드)
2. 내용을 Claude에게 보내 3개 구분의 표 초안을 만듭니다. (AI)
3. 결과를 '통합_주간보고.md'로 같은 폴더에 저장합니다. (코드)

사람이 할 일: 저장된 결과를 원문과 대조해 확인합니다.
"""

import os
import sys
import time
from pathlib import Path

# ----- 바꿔 쓸 수 있는 설정 -----
FOLDER_NAME = "연습"              # 바탕화면 안의 폴더 이름
INPUT_PREFIX = "보고_"             # 입력 파일 이름의 시작 부분
OUTPUT_NAME = "통합_주간보고.md"   # 결과 파일 이름
MODEL = "claude-sonnet-5-5"        # 사용할 Claude 모델
A4_CHAR_LIMIT = None               # A4 한 장 글자 수 기준. 첫 실습 후 숫자로 채웁니다. 예: 1800
# --------------------------------

SYSTEM_PROMPT = """당신은 팀 주간보고를 통합하는 도우미입니다.
팀원 3명의 주간보고 원문을 받아 아래 규칙대로 하나의 표 초안을 작성하세요.

[규칙]
1. 원문에 없는 내용은 절대 쓰지 않습니다. 일정·숫자·표현을 추가하거나 추측하지 않습니다.
2. 개조식으로 작성하고, 처음 읽는 사람도 이해하기 쉽게 간결하게 씁니다.
3. 원문의 모든 항목을 아래 3개 구분 중 하나에 넣습니다.
   ○ 정보보호 인력양성 정책개발 및 이행관리
   ○ 정보보호 인력양성 기반 조성 및 협력체계 마련
   ○ 정보보호 인력양성 예산·성과 관리 등
4. 각 구분 안에서 "이번주 실적"과 "다음주 계획"을 나눠 씁니다.
5. "-"에는 '무엇을 했는지(할 것인지) + 일정' 순으로 씁니다.
6. "※"에는 구체적인 내용, 향후 일정 또는 부연 설명을 씁니다. 원문에 근거가 없으면 "※"를 쓰지 않습니다.
7. 다음 팀 용어는 원문 표기 그대로 씁니다.
   AI보안 인재 양성방안, 정보보호 프레임워크 개발, 정보보호 직무역량체계 개발, 플랫폼 구축
8. 어느 구분에 넣을지 애매한 항목은 가장 가까운 구분에 넣고, 표 아래에 "확인 필요: (항목)"으로 따로 적습니다.
9. 오타로 보이는 부분은 고치지 말고 표 아래에 "오타 후보: (원문 → 제안)"으로 적습니다.

[출력 형식]
마크다운 표 1개 (열: 구분 | 이번주 실적 | 다음주 계획), 그 아래에 확인 필요 항목과 오타 후보.
표 안에서 줄바꿈은 <br>로 표시합니다."""


def find_folder():
    """바탕화면 > 연습 폴더를 찾습니다. (한글 윈도우, OneDrive 바탕화면도 확인)"""
    home = Path.home()
    candidates = [
        home / "Desktop",
        home / "바탕 화면",
        home / "OneDrive" / "Desktop",
        home / "OneDrive" / "바탕 화면",
    ]
    for base in candidates:
        folder = base / FOLDER_NAME
        if folder.is_dir():
            return folder
    return None


def read_reports(folder):
    """'보고_'로 시작하는 .txt 파일을 이름순으로 읽습니다."""
    files = sorted(folder.glob(f"{INPUT_PREFIX}*.txt"))
    reports = []
    for f in files:
        # utf-8-sig: 메모장에서 저장한 UTF-8 파일도 안전하게 읽습니다.
        text = f.read_text(encoding="utf-8-sig").strip()
        reports.append((f.name, text))
    return reports


def build_user_message(reports):
    parts = []
    for name, text in reports:
        parts.append(f"=== 파일: {name} ===\n{text}")
    return "아래는 팀원별 주간보고 원문입니다.\n\n" + "\n\n".join(parts)


def call_claude(user_message):
    try:
        import anthropic
    except ImportError:
        sys.exit("anthropic 패키지가 없습니다. 명령 창에서 'pip install anthropic'을 먼저 실행해 주세요.")

    if not os.environ.get("ANTHROPIC_API_KEY"):
        sys.exit("API 키가 설정되어 있지 않습니다. 환경 변수 ANTHROPIC_API_KEY를 먼저 등록해 주세요.")

    client = anthropic.Anthropic()
    response = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    )
    return "".join(block.text for block in response.content if block.type == "text")


def main():
    start = time.time()

    folder = find_folder()
    if folder is None:
        sys.exit(f"바탕화면에서 '{FOLDER_NAME}' 폴더를 찾지 못했습니다. 폴더 이름과 위치를 확인해 주세요.")

    reports = read_reports(folder)
    if not reports:
        sys.exit(f"'{folder}' 안에 '{INPUT_PREFIX}'로 시작하는 .txt 파일이 없습니다.")
    if len(reports) != 3:
        print(f"주의: 입력 파일이 {len(reports)}개입니다. (기대: 3개) 그대로 진행합니다.")

    print("읽은 파일:", ", ".join(name for name, _ in reports))
    print("Claude에게 요청 중입니다. 잠시 기다려 주세요...")

    result = call_claude(build_user_message(reports))

    output_path = folder / OUTPUT_NAME
    output_path.write_text(result, encoding="utf-8")

    char_count = len("".join(result.split()))  # 공백·줄바꿈 제외 글자 수
    print(f"\n저장 완료: {output_path}")
    print(f"결과 글자 수(공백 제외): {char_count}자")
    if A4_CHAR_LIMIT is not None and char_count > A4_CHAR_LIMIT:
        print(f"주의: A4 한 장 기준({A4_CHAR_LIMIT}자)을 넘었습니다. 내용을 줄여야 할 수 있습니다.")
    print(f"걸린 시간: {time.time() - start:.1f}초")
    print("\n다음 할 일: 결과를 원문과 대조해 확인하세요.")
    print("  1) 원문에 없는 내용이 있는지  2) 3개 구분 분류가 맞는지  3) 일정이 빠지지 않았는지")


if __name__ == "__main__":
    main()
