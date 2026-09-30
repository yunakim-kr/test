# Vercel 배포용 폴더

이 폴더(`vercel-sites`)만 Vercel에 배포합니다. 상위 폴더의 다른 파일(.env, Streamlit 앱 등)은 포함되지 않습니다.

| 주소 | 원본 |
|---|---|
| `/` | 바로가기 목차 |
| `/intro/` | ../소개페이지/index.html |
| `/jobs/` | ../정보보호_직무역량체계/index.html |

## 배포
```
cd vercel-sites
npx vercel        # 미리보기
npx vercel --prod # 운영 배포
```
Vercel 대시보드에서 Git 연동 시 Root Directory를 `vercel-sites`로 지정하세요.
원본을 수정하면 해당 index.html을 public 하위로 다시 복사해야 합니다.
