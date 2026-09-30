import os
import re

# 입출력 폴더: 바탕화면 '주간업무보고' 폴더 (기획서 6번, 임시 처리 - 실제 위치 다르면 이 값만 교체)
# 주의: 빈 서식/템플릿 원본은 이 폴더에 두지 않는다 - 자동 날짜 태깅·파일 인식 대상에
# 함께 걸려서 팀원 파일로 잘못 인식될 수 있다 (기획서 6번 참고).
DATA_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "주간업무보고")

# 통합 파이프라인이 실제로 읽는 형식: 원본 .hwp(COM 자동화로 직접 읽음),
# 또는 실습용 .txt/.md
INPUT_EXTS = (".hwp", ".txt", ".md")
EXPECTED_MEMBER_COUNT = 3

# 자동 날짜 태깅 대상 확장자: 팀원이 처음 올리는 파일은 원본 .hwp이므로 포함한다.
# (.hwp 파일 자체를 자동으로 읽지는 않음 - 기획서 8번, 이번에 안 할 것)
AUTO_TAG_EXTS = (".hwp", ".txt", ".md")

# 통합 대상 파일명 규칙 (기획서 6번): "파일명_날짜(YYMMDD)" 형식만 인식한다.
# 예: 김민준_251121.txt, 팀원A_260928.md
# 이 형식이 아닌 파일(raw_material.md, summary.md 등)은 자동으로 제외된다.
INPUT_FILENAME_PATTERN = re.compile(r"^.+_\d{6}$")

# 이 스크립트 자신이 만드는 산출물(원자료·요약·결과 파일)은 "이름_날짜" 형식이어도
# 통합 대상에서 제외한다 (예: 주간보고_통합본_20260928.md).
OUTPUT_FILENAME_PREFIXES = ("주간보고_통합본",)

# 자동 날짜 태깅에서 제외할 파일명 (산출물/작업 파일). 확장자 포함, 정확히 일치.
EXCLUDE_FROM_AUTO_RENAME = {"raw_material.md", "summary.md"}

# 3개 구분 (기획서 7번 - 표기 그대로 사용, 순서 고정)
CATEGORIES = [
    "정보보호 인력양성 정책개발 및 이행관리",
    "정보보호 인력양성 기반 조성 및 협력체계 마련",
    "정보보호 인력양성 예산·성과 관리 등",
]

SECTIONS = ["이번주 실적", "다음주 계획"]

# 팀원 원본 파일에서 실제로 쓰이는 표현이 다를 수 있어 동의어를 함께 인식한다.
# (예: 실제 hwp 양식은 "주요실적(기간)" / "주요계획(기간)"으로 표기됨)
SECTION_ALIASES = {
    "이번주 실적": ["이번주 실적", "주요실적", "금주실적", "금주 실적"],
    "다음주 계획": ["다음주 계획", "주요계획", "차주계획", "차주 계획"],
}

# 팀 용어 (표기 그대로 사용해야 하는 용어, 기획서 7번)
TEAM_TERMS = [
    "AI보안 인재 양성방안",
    "정보보호 프레임워크 개발",
    "정보보호 직무역량체계 개발",
    "플랫폼 구축",
]

# 분량 기준 (기획서 7번, 임시 처리 - 실측 기준 확정되면 교체)
MAX_CHARS_NO_SPACE = 2000
MAX_LINES = 40

RAW_MATERIAL_FILENAME = "raw_material.md"
SUMMARY_FILENAME = "summary.md"

# AI 자동 요약 (Claude API) 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")
WRITING_PRINCIPLES_PATH = os.path.join(BASE_DIR, "..", "작성_원칙.md")
ANTHROPIC_MODEL = "claude-sonnet-5"
