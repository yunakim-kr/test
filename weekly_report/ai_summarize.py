"""
raw_material.md를 Claude API에 보내 작성_원칙.md 규칙대로 summary.md를 자동
생성한다. (기획서 5번 "AI" 단계의 자동화 버전 - 사람이 대화로 요청하지 않아도 됨)

주의: 이 단계가 자동화되어도 기획서 5번의 "사람 확인" 체크리스트는 그대로
유효하다. API가 만든 통합본은 초안이며, 최종 제출 전 반드시 사람이 확인해야
한다 - assemble_report.py가 저장하는 파일 상단에 경고 문구를 남긴다.

필요 조건: weekly_report/.env 파일에 ANTHROPIC_API_KEY=sk-... 형태로 키가
있어야 한다. 키가 없거나 API 호출이 실패하면 조용히 넘어가지 않고
AiSummarizeError를 던져 자동화를 멈춘다 (기획서 5번 "실패 시 사람에게 알림").
"""
import os

import config


class AiSummarizeError(Exception):
    """AI 요약 자동 생성에 실패했을 때 - 자동화를 멈추고 사람에게 알려야 하는 상황."""


def _load_api_key():
    try:
        from dotenv import load_dotenv
    except ImportError as e:
        raise AiSummarizeError(
            "python-dotenv가 설치되어 있지 않습니다 (pip install python-dotenv)."
        ) from e

    load_dotenv(config.ENV_PATH)
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise AiSummarizeError(
            f"ANTHROPIC_API_KEY가 설정되어 있지 않습니다. "
            f"'{config.ENV_PATH}' 파일에 ANTHROPIC_API_KEY=sk-... 형태로 넣어주세요."
        )
    return api_key


def _load_writing_principles():
    try:
        with open(config.WRITING_PRINCIPLES_PATH, encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError as e:
        raise AiSummarizeError(
            f"작성 원칙 파일을 찾을 수 없습니다: {config.WRITING_PRINCIPLES_PATH}"
        ) from e


def _build_prompt(raw_material_text):
    principles = _load_writing_principles()
    categories = "\n".join(f"- {c}" for c in config.CATEGORIES)
    terms = ", ".join(config.TEAM_TERMS)

    system_prompt = f"""다음은 주간업무보고 작성 원칙이다. 반드시 그대로 따른다.

{principles}

[이 프로젝트의 고정 규칙]
- 최종 구분(○ 상위항목)은 반드시 아래 3개를 표기 그대로 사용한다. 다른 이름으로 바꾸지 않는다:
{categories}
- 다음 팀 용어가 원문에 등장하면 표기를 절대 바꾸지 않는다: {terms}
- 원문(raw_material.md)에 없는 사실은 한 글자도 추가하지 않는다. 표현과 구조만 재구성한다.
- 같은 구분 안에서 3명의 중복/유사 내용은 하나로 통합한다.

[출력 형식 - 반드시 이 형식을 그대로 따른다. 다른 설명이나 인사말 없이 이 내용만 출력한다]

○ <구분명 그대로>
주요실적
- (키워드) 내용(일정)
※ 필요시 보충설명
주요계획
- (키워드) 내용(일정)
※ 필요시 보충설명

○ <다음 구분명>
...

(3개 구분 모두 출력. 실적이나 계획에 해당 내용이 전혀 없으면 그 섹션은 "주요실적"/"주요계획" 줄만 쓰고 항목 없이 비워둔다. 있는 내용만 쓰고 없는 내용을 지어내지 않는다.)
"""

    user_prompt = f"""아래는 팀원 3명의 이번 주 주간업무보고 원문을 구분·섹션별로 취합한 원자료다.
이 내용만 바탕으로 위 형식에 맞는 summary를 작성하라.

{raw_material_text}
"""
    return system_prompt, user_prompt


def generate_summary(raw_material_text):
    """반환: summary.md에 들어갈 텍스트."""
    try:
        import anthropic
    except ImportError as e:
        raise AiSummarizeError(
            "anthropic 패키지가 설치되어 있지 않습니다 (pip install anthropic)."
        ) from e

    api_key = _load_api_key()
    system_prompt, user_prompt = _build_prompt(raw_material_text)

    client = anthropic.Anthropic(api_key=api_key)
    try:
        response = client.messages.create(
            model=config.ANTHROPIC_MODEL,
            max_tokens=4000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
    except Exception as e:
        raise AiSummarizeError(f"Claude API 호출에 실패했습니다: {e}") from e

    text = "".join(
        block.text for block in response.content if getattr(block, "type", None) == "text"
    ).strip()

    if not text:
        raise AiSummarizeError("Claude API 응답이 비어 있습니다.")

    return text


def write_summary(raw_material_path, out_path=None):
    with open(raw_material_path, encoding="utf-8") as f:
        raw_material_text = f.read()

    summary_text = generate_summary(raw_material_text)

    if out_path is None:
        out_path = os.path.join(os.path.dirname(raw_material_path), config.SUMMARY_FILENAME)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(summary_text)

    return out_path
