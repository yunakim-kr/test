"""
[1][2][3] 담당: 입력 파일 탐색, 파싱, 구분/섹션별 원자료 취합.

통합 대상 파일 인식 규칙 (기획서 6번): 폴더 안의 파일 중 파일명이
"이름_날짜(YYMMDD)" 형식(예: 김민준_251121.txt)인 파일만 통합 대상으로
인식한다. 그 외 파일(원자료·요약·결과 파일 등)은 이름이 같은 규칙을
우연히 따르더라도 config.OUTPUT_FILENAME_PREFIXES에 등록해 제외한다.

날짜 태그가 없는 새 파일이 폴더에 올라오면, find_input_files 실행 전에
auto_tag_new_files()가 그 파일이 폴더에 처음 생긴 날짜를 자동으로 붙여
이름을 바꾼다. 팀원이 실제로 맨 처음 올리는 파일은 원본 .hwp이므로
(예: 김민준.hwp -> 김민준_260928.hwp) .hwp도 태깅 대상에 포함한다
(config.AUTO_TAG_EXTS). 다만 .hwp 파일 자체를 읽지는 않으며(기획서 8번),
이 파이프라인이 실제로 파싱하는 것은 그 내용을 옮겨 적은 .txt/.md 사본이다.

가정하는 팀원 입력 파일 형식 (.txt 또는 .md, UTF-8):

    ## <구분명 그대로>
    ### 이번주 실적
    - 항목 1
    - 항목 2
    ### 다음주 계획
    - 항목 1

    ## <다음 구분명>
    ...

구분명은 config.CATEGORIES 값이 줄 안에 포함되어 있으면 인식한다(번호나
"##" 등 앞뒤 장식은 무시). 실제 파일 형식이 다르면 이 파일의 파싱 규칙만
고치면 된다.
"""
import datetime
import os
import glob

import config


def _upload_date_str(path):
    """파일이 폴더에 처음 생긴(업로드된) 날짜를 YYMMDD로 반환한다.
    Windows에서는 st_ctime이 '생성 시각'이며, 다른 위치에서 복사해 온
    시각을 반영하므로 '처음 업로드한 날짜'로 사용하기에 적합하다."""
    ts = os.path.getctime(path)
    return datetime.date.fromtimestamp(ts).strftime("%y%m%d")


def auto_tag_new_files(folder=config.DATA_FOLDER):
    """날짜 태그가 없는 새 파일에 처음 업로드된 날짜(YYMMDD)를 자동으로 붙인다.
    예: 김민준.txt -> 김민준_260928.txt (폴더에 생긴 날짜 기준)
    이미 "이름_날짜" 형식이거나, 이 파이프라인의 산출물인 파일은 건드리지 않는다.
    같은 이름의 파일이 이미 있으면 충돌을 피하기 위해 건드리지 않고 넘어간다."""
    renamed = []
    for ext in config.AUTO_TAG_EXTS:
        for path in glob.glob(os.path.join(folder, f"*{ext}")):
            fname = os.path.basename(path)
            stem, ext_ = os.path.splitext(fname)
            if fname in config.EXCLUDE_FROM_AUTO_RENAME:
                continue
            if stem.startswith(config.OUTPUT_FILENAME_PREFIXES):
                continue
            if config.INPUT_FILENAME_PATTERN.match(stem):
                continue  # 이미 날짜 태그가 있음
            date_str = _upload_date_str(path)
            new_name = f"{stem}_{date_str}{ext_}"
            new_path = os.path.join(folder, new_name)
            if os.path.exists(new_path):
                continue  # 이름 충돌 - 사람 확인 필요, 자동으로 덮어쓰지 않음
            os.rename(path, new_path)
            renamed.append((fname, new_name))
    return renamed


def _is_target_file(path):
    """파일명이 "파일명_날짜(YYMMDD)" 형식인지 확인한다 (기획서 6번 파일 인식 규칙)."""
    stem = os.path.splitext(os.path.basename(path))[0]
    if stem.startswith(config.OUTPUT_FILENAME_PREFIXES):
        return False
    return bool(config.INPUT_FILENAME_PATTERN.match(stem))


def find_input_files(folder=config.DATA_FOLDER):
    candidates = []
    for ext in config.INPUT_EXTS:
        candidates.extend(glob.glob(os.path.join(folder, f"*{ext}")))
    files = sorted(f for f in candidates if _is_target_file(f))
    if len(files) != config.EXPECTED_MEMBER_COUNT:
        raise SystemExit(
            f"'파일명_날짜' 형식의 통합 대상 파일이 {config.EXPECTED_MEMBER_COUNT}개가 아닙니다 "
            f"(찾은 개수: {len(files)}, 폴더: {folder}). "
            f"팀원 3명의 보고 파일이 '이름_날짜.txt(.md)' 형식으로 있는지 확인하세요."
        )
    return files


def _match_category(line):
    cleaned = line.strip().lstrip("#").strip()
    for cat in config.CATEGORIES:
        if cat in cleaned:
            return cat
    return None


def _match_section(line):
    cleaned = line.strip().lstrip("#").strip()
    for sec, aliases in config.SECTION_ALIASES.items():
        if any(alias in cleaned for alias in aliases):
            return sec
    return None


def parse_file(path):
    """반환: {구분: {섹션: [원문 줄, ...]}}"""
    if path.lower().endswith(".hwp"):
        import hwp_reader

        return hwp_reader.extract_hwp_structured(path)

    result = {cat: {sec: [] for sec in config.SECTIONS} for cat in config.CATEGORIES}
    current_cat = None
    current_sec = None

    with open(path, encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.rstrip("\n")
            if not line.strip():
                continue

            cat = _match_category(line)
            if cat:
                current_cat = cat
                current_sec = None
                continue

            sec = _match_section(line)
            if sec:
                current_sec = sec
                continue

            if current_cat and current_sec:
                result[current_cat][current_sec].append(line.strip())

    return result


def build_raw_material(files):
    """반환: {구분: {섹션: {팀원파일명: [원문 줄, ...]}}}"""
    raw = {cat: {sec: {} for sec in config.SECTIONS} for cat in config.CATEGORIES}
    for path in files:
        person = os.path.splitext(os.path.basename(path))[0]
        parsed = parse_file(path)
        for cat in config.CATEGORIES:
            for sec in config.SECTIONS:
                raw[cat][sec][person] = parsed[cat][sec]
    return raw


def write_raw_material(raw, out_path):
    """AI 요약 단계에서 참고할 수 있도록, 구분×섹션별로 3명 원문을 모아 파일로 출력."""
    lines = ["# 원자료 취합 (요약 전, 원문 그대로)\n"]
    for cat in config.CATEGORIES:
        lines.append(f"## {cat}\n")
        for sec in config.SECTIONS:
            lines.append(f"### {sec}\n")
            for person, items in raw[cat][sec].items():
                lines.append(f"**{person}**")
                if items:
                    for item in items:
                        prefix = "" if item.startswith(("-", "※")) else "- "
                        lines.append(f"{prefix}{item}")
                else:
                    lines.append("- (내용 없음)")
                lines.append("")
    content = "\n".join(lines)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(content)
    return out_path


if __name__ == "__main__":
    import hwp_reader

    renamed = auto_tag_new_files()
    for old_name, new_name in renamed:
        print(f"날짜 태그 자동 추가: {old_name} -> {new_name}")

    files = find_input_files()
    try:
        raw = build_raw_material(files)
    except hwp_reader.HwpReadError as e:
        raise SystemExit(
            f"[중단] .hwp 파일을 읽는 데 문제가 생겨 자동화를 멈췄습니다: {e}\n"
            f"사람이 해당 파일을 직접 확인해 주세요."
        )
    out_path = os.path.join(config.DATA_FOLDER, config.RAW_MATERIAL_FILENAME)
    write_raw_material(raw, out_path)
    print(f"입력 파일 {len(files)}개 파싱 완료 -> {out_path}")
