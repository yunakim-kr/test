"""
Claude API가 이 프로그램에서 정상적으로 동작하는지 확인하는 간단한 테스트 스크립트.

weekly_report/.env 파일의 ANTHROPIC_API_KEY를 읽어서 사용한다.
API 키 자체는 코드/화면 어디에도 출력하지 않는다.
"""
import os

import config
from dotenv import load_dotenv
import anthropic


def load_api_key():
    load_dotenv(config.ENV_PATH)
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            f"ANTHROPIC_API_KEY가 설정되어 있지 않습니다. "
            f"'{config.ENV_PATH}' 파일에 ANTHROPIC_API_KEY=sk-... 형태로 넣어주세요."
        )
    return api_key


def main():
    api_key = load_api_key()
    client = anthropic.Anthropic(api_key=api_key)

    company_info = "한빛정밀은 자동차 부품을 생산하며, 생산관리팀은 생산 계획과 라인 운영을 맡는다."

    response = client.messages.create(
        model=config.ANTHROPIC_MODEL,
        max_tokens=200,
        messages=[
            {
                "role": "user",
                "content": f"{company_info}\n\n이 회사와 팀을 한 문장으로 소개해줘.",
            }
        ],
    )

    text = "".join(
        block.text for block in response.content if getattr(block, "type", None) == "text"
    ).strip()

    print("Claude 응답:")
    print(text)


if __name__ == "__main__":
    main()
