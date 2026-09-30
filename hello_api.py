import os
import sys

sys.stdout.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
from anthropic import Anthropic

load_dotenv()

api_key = os.environ.get("ANTHROPIC_API_KEY")
if not api_key:
    sys.exit(".env 파일에 ANTHROPIC_API_KEY가 설정되어 있지 않습니다.")

client = Anthropic(api_key=api_key)

company_info = (
    "한빛정밀은 자동차 부품을 생산하며, "
    "생산관리팀은 생산 계획과 라인 운영을 맡는다."
)

response = client.messages.create(
    model="claude-sonnet-5-5",
    max_tokens=200,
    messages=[
        {
            "role": "user",
            "content": f"{company_info}\n\n위 회사와 팀을 한 문장으로 소개해줘.",
        }
    ],
)

print(response.content[0].text)
