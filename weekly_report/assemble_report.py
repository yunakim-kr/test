"""
[5][6][7] 담당: 요약된 내용을 표 틀에 배치, 분량 점검, 저장, 사람 확인 체크리스트 출력.

입력으로 받는 "요약 파일" 형식은 raw_material.md와 비슷하되, 팀원별 구분 없이
구분×섹션당 이미 통합/요약된 "-"와 "※" 항목만 들어있다고 가정한다 (AI 요약 단계
[4]의 산출물). 이번 범위에서 [4]는 스크립트가 아니라 Claude가 직접 작성한다.

    ## <구분명>
    ### 이번주 실적
    - 항목
    ※ 부연
    ### 다음주 계획
    - 항목

    ## <다음 구분명>
    ...
"""
import datetime
import os

import config


def parse_summary(path):
    """반환: {구분: {섹션: [요약된 줄, ...]}}"""
    result = {cat: {sec: [] for sec in config.SECTIONS} for cat in config.CATEGORIES}
    current_cat = None
    current_sec = None

    with open(path, encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.rstrip("\n")
            if not line.strip():
                continue

            cleaned = line.strip().lstrip("#").strip()

            matched_cat = next((c for c in config.CATEGORIES if c in cleaned), None)
            if matched_cat:
                current_cat = matched_cat
                current_sec = None
                continue

            matched_sec = next(
                (
                    sec
                    for sec, aliases in config.SECTION_ALIASES.items()
                    if any(alias in cleaned for alias in aliases)
                ),
                None,
            )
            if matched_sec:
                current_sec = matched_sec
                continue

            if current_cat and current_sec:
                result[current_cat][current_sec].append(line.strip())

    return result


def build_table(summary):
    """기획서 7번 형식: 이번주 실적/다음주 계획을 세로로 나눈 표, 구분 3개를 행으로."""
    header = "| 구분 | 이번주 실적 | 다음주 계획 |"
    divider = "|---|---|---|"
    rows = [header, divider]

    for cat in config.CATEGORIES:
        cell_this_week = "<br>".join(summary[cat]["이번주 실적"]) or "-"
        cell_next_week = "<br>".join(summary[cat]["다음주 계획"]) or "-"
        rows.append(f"| {cat} | {cell_this_week} | {cell_next_week} |")

    return "\n".join(rows)


def check_length(table_text):
    """분량 점검(임시 기준, config.MAX_CHARS_NO_SPACE / MAX_LINES 참고)."""
    char_count = len(table_text.replace(" ", "").replace("\n", ""))
    line_count = table_text.count("\n") + 1

    warnings = []
    if char_count > config.MAX_CHARS_NO_SPACE:
        warnings.append(
            f"글자 수 {char_count}자가 임시 기준 {config.MAX_CHARS_NO_SPACE}자를 초과했습니다."
        )
    if line_count > config.MAX_LINES:
        warnings.append(
            f"줄 수 {line_count}줄이 임시 기준 {config.MAX_LINES}줄을 초과했습니다."
        )
    return char_count, line_count, warnings


def check_team_terms(table_text):
    """팀 용어가 다른 표기로 바뀌지 않았는지 참고용으로만 점검 (최종 판단은 사람 몫)."""
    notes = []
    for term in config.TEAM_TERMS:
        if term not in table_text:
            notes.append(f"'{term}' 용어가 결과물에 보이지 않습니다 (원래 없을 수도 있음).")
    return notes


def save_output(table_text, folder=config.DATA_FOLDER, auto_generated=False):
    os.makedirs(folder, exist_ok=True)
    date_str = datetime.date.today().strftime("%y%m%d")
    out_path = os.path.join(folder, f"주간보고_통합본_{date_str}.md")
    banner = (
        "> ⚠️ **이 파일은 AI가 자동으로 생성한 초안입니다.** "
        "제출 전 기획서 5번 체크리스트(아래)를 사람이 반드시 확인하세요.\n\n"
        if auto_generated
        else ""
    )
    tail = f"\n\n{checklist_markdown()}\n" if auto_generated else ""
    content = "# 주간보고 통합본 (초안)\n\n" + banner + table_text + tail
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(content)
    return out_path


CHECKLIST_ITEMS = [
    "원문에 없는 내용이 없는지",
    "누락된 내용이 없는지",
    "일정 표기가 정확한지",
    "팀 용어가 표기 그대로 쓰였는지",
    "한글(.hwp)로 옮긴 뒤 서식과 A4 한 장 여부",
]


def checklist_markdown():
    lines = ["## 사람 확인 체크리스트 (기획서 5번)", ""]
    lines += [f"- [ ] {item}" for item in CHECKLIST_ITEMS]
    return "\n".join(lines)


def print_checklist():
    print("\n사람 확인 체크리스트 (기획서 5번):")
    for i, item in enumerate(CHECKLIST_ITEMS, start=1):
        print(f"  {i}. [ ] {item}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) not in (2, 3):
        raise SystemExit(
            "사용법: python assemble_report.py <요약파일 경로> [출력 폴더(생략 시 기본 폴더)]"
        )

    summary_path = sys.argv[1]
    out_folder = sys.argv[2] if len(sys.argv) == 3 else config.DATA_FOLDER

    summary = parse_summary(summary_path)
    table_text = build_table(summary)

    char_count, line_count, warnings = check_length(table_text)
    term_notes = check_team_terms(table_text)

    out_path = save_output(table_text, folder=out_folder)
    print(f"통합본 저장 완료 -> {out_path}")
    print(f"분량: 공백 제외 {char_count}자, {line_count}줄")
    for w in warnings:
        print(f"  경고: {w}")
    for n in term_notes:
        print(f"  참고: {n}")

    print_checklist()
