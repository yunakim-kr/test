import glob
import os
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
from anthropic import Anthropic, APIError, APIStatusError

REPORT_PATTERN = "주간보고_*.txt"
MODEL = "claude-sonnet-5-5"

SYSTEM_PROMPT = """당신은 여러 팀의 주간보고를 취합해 경영 보고서 초안을 작성하는 비서입니다.
반드시 아래 규칙을 지키세요.

- 제공된 원문(각 팀 주간보고)에 없는 사실이나 수치를 새로 만들어내지 않습니다.
- 담당자나 기한이 원문에 분명하게 나와 있지 않으면 "확인 필요"라고 표시합니다.
- 같은 사안이 여러 팀 보고서에 나오면 하나로 합쳐서 정리하고, 관련 팀 이름과 출처 파일명을 함께 적습니다.
- 제공된 모든 팀의 주요 내용이 결과에서 빠지지 않도록 합니다.
- 결과는 반드시 아래 5개 항목을 이 순서대로, 항목 번호와 제목을 그대로 사용해 작성합니다.

1. 핵심 요약 3줄
2. 팀별 주요 실적
3. 여러 부서가 함께 확인할 이슈와 담당·기한
4. 대표가 결정해야 할 사항
5. 다음 주 주요 일정
"""


def load_api_key() -> str:
    load_dotenv()
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(".env 파일에 ANTHROPIC_API_KEY가 설정되어 있지 않습니다.")
    return api_key


def load_reports(folder: str):
    paths = sorted(glob.glob(os.path.join(folder, REPORT_PATTERN)))
    if not paths:
        raise RuntimeError(
            f"읽을 자료가 없습니다. ('{folder}' 폴더에서 '{REPORT_PATTERN}' 형식의 주간보고 파일을 찾지 못했습니다.)"
        )

    reports = []
    for path in paths:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read().strip()
        if not content:
            raise RuntimeError(f"'{os.path.basename(path)}' 파일 내용이 비어 있습니다.")
        reports.append({"filename": os.path.basename(path), "content": content})
    return reports


def build_user_prompt(reports) -> str:
    parts = ["다음은 각 팀의 이번 주 주간보고 원문입니다. 이를 바탕으로 경영 보고서 초안을 작성해주세요.\n"]
    for r in reports:
        parts.append(f"----- 출처 파일: {r['filename']} -----\n{r['content']}\n")
    return "\n".join(parts)


def generate_report(api_key: str, reports) -> str:
    client = Anthropic(api_key=api_key)
    response = client.messages.create(
        model=MODEL,
        max_tokens=4000,
        thinking={"type": "between_tools"},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": build_user_prompt(reports)}],
    )
    text_blocks = [block.text for block in response.content if getattr(block, "type", None) == "text"]
    text = "\n".join(text_blocks).strip()
    if not text:
        raise RuntimeError("Claude로부터 빈 응답을 받았습니다.")
    return text


def make_output_path(folder: str) -> str:
    today = datetime.now().strftime("%Y-%m-%d")
    base = f"주간경영보고_초안_{today}.txt"
    path = os.path.join(folder, base)
    idx = 1
    while os.path.exists(path):
        path = os.path.join(folder, f"주간경영보고_초안_{today}_{idx}.txt")
        idx += 1
    return path


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))

    try:
        api_key = load_api_key()
        reports = load_reports(folder)
        report_text = generate_report(api_key, reports)
    except (RuntimeError, APIError, APIStatusError) as e:
        print(f"오류가 발생하여 결과 파일을 만들지 않았습니다: {e}")
        return
    except Exception as e:
        print(f"예상하지 못한 오류가 발생하여 결과 파일을 만들지 않았습니다: {e}")
        return

    output_path = make_output_path(folder)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(report_text)

    print(f"읽은 주간보고 파일 수: {len(reports)}")
    for r in reports:
        print(f" - {r['filename']}")
    print(f"결과 저장 위치: {output_path}")


if __name__ == "__main__":
    main()
