"""
.hwp 원본을 한글 프로그램 COM 자동화로 직접 읽어, parse_reports.parse_file()과
같은 형태({구분: {섹션: [원문 줄, ...]}})로 반환한다.

동작 방식:
1. 한글 프로그램을 COM으로 띄워 .hwp 파일을 연다 (읽기만 하고 저장하지 않음).
2. GetTextFile("HWPML2X")로 문서 전체를 XML로 내보낸다. 이 형식은 표의 각 셀이
   ColAddr(열 번호)를 가지고 있어서, "이번주 실적"/"다음주 계획"이 표의 좌우
   컬럼으로 나뉘어 있어도 셀 위치로 구분할 수 있다 (실제 hwp 서식이 이 구조임).
3. 표(TABLE)마다 헤더 행에서 "주요실적"/"주요계획" 등 키워드로 어느 컬럼이
   어느 섹션인지 확인한다. 확인이 안 되면 기본값(0번째 열=이번주 실적,
   1번째 열=다음주 계획)을 쓴다.
4. 각 셀 안에서 "○ 구분명" 문단을 만나면 그 아래 "-"/"※" 문단들을 해당
   구분·섹션에 쌓는다.

구조가 예상과 다르거나 한글 프로그램/COM 연동이 실패하면, 조용히 빈 결과를
만들지 않고 예외를 던져 상위에서 사람에게 알리도록 한다 (기획서 5번).
"""
import xml.etree.ElementTree as ET

import config


class HwpReadError(Exception):
    """.hwp 읽기에 실패했을 때 - 자동화를 멈추고 사람에게 알려야 하는 상황."""


def _get_hwp_xml(path):
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError as e:
        raise HwpReadError(
            "pywin32가 설치되어 있지 않아 .hwp를 읽을 수 없습니다 (pip install pywin32)."
        ) from e

    pythoncom.CoInitialize()
    try:
        hwp = win32.Dispatch("HWPFrame.HwpObject")
    except Exception as e:
        raise HwpReadError(
            "한글 프로그램을 자동화로 열 수 없습니다. 이 PC에 한글이 설치되어 있는지 확인하세요."
        ) from e

    try:
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    except Exception:
        pass  # 등록 실패해도 열기 자체는 되는 경우가 있어 계속 진행

    ok = hwp.Open(path, "HWP", "")
    if not ok:
        hwp.Quit()
        raise HwpReadError(f"'{path}' 파일을 여는 데 실패했습니다.")

    xml_text = hwp.GetTextFile("HWPML2X", "")
    hwp.Quit()
    return xml_text


def _char_text(char_elem):
    """<CHAR> 안의 <FWSPACE/> 등으로 끊긴 텍스트 조각을 이어 붙인다."""
    parts = [char_elem.text or ""]
    for child in char_elem:
        parts.append(child.tail or "")
    return "".join(parts)


def _para_text(p_elem):
    parts = []
    for text_elem in p_elem.findall("TEXT"):
        char_elem = text_elem.find("CHAR")
        if char_elem is not None:
            parts.append(_char_text(char_elem))
    return "".join(parts).strip()


def _match_category(line):
    for cat in config.CATEGORIES:
        if cat in line:
            return cat
    return None


def _detect_column_sections(table_elem):
    """표의 첫 행에서 각 ColAddr가 실적/계획 중 무엇인지 확인한다.
    확인 안 되면 기본값(0=이번주 실적, 1=다음주 계획)을 쓴다."""
    col_section = {}
    rows = table_elem.findall(".//ROW")
    if rows:
        first_row = rows[0]
        for cell in first_row.findall("CELL"):
            col = cell.get("ColAddr")
            text = " ".join(_para_text(p) for p in cell.iter("P"))
            for sec, aliases in config.SECTION_ALIASES.items():
                if any(alias in text for alias in aliases):
                    col_section[col] = sec
    if "0" not in col_section:
        col_section["0"] = config.SECTIONS[0]
    if "1" not in col_section:
        col_section["1"] = config.SECTIONS[1]
    return col_section


def _extract_from_table(table_elem, result):
    col_section = _detect_column_sections(table_elem)
    for row in table_elem.findall(".//ROW"):
        for cell in row.findall("CELL"):
            col = cell.get("ColAddr")
            section = col_section.get(col)
            if section is None:
                continue
            current_cat = None
            for p in cell.iter("P"):
                text = _para_text(p)
                if not text:
                    continue
                cat = _match_category(text)
                if cat:
                    current_cat = cat
                    continue
                if current_cat and text not in ("○", "-", "※"):
                    result[current_cat][section].append(text)


def extract_hwp_structured(path):
    """반환: {구분: {섹션: [원문 줄, ...]}} - parse_reports.parse_file()과 동일한 형태."""
    xml_text = _get_hwp_xml(path)
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        raise HwpReadError(f"'{path}'의 문서 구조를 해석하는 데 실패했습니다: {e}") from e

    result = {cat: {sec: [] for sec in config.SECTIONS} for cat in config.CATEGORIES}
    tables = root.findall(".//TABLE")
    if not tables:
        raise HwpReadError(
            f"'{path}'에서 표를 찾지 못했습니다. 서식이 예상과 다른 것 같습니다 - 사람 확인이 필요합니다."
        )

    for table in tables:
        _extract_from_table(table, result)

    return result
