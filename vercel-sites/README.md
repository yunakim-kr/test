# Vercel 배포용 폴더

이 폴더(`vercel-sites`)만 Vercel에 배포합니다. 상위 폴더의 다른 파일(.env, weekly_report 코드 등)은 포함되지 않습니다.

| 주소 | 원본 |
|---|---|
| `/` | 바로가기 목차 |
| `/intro/` | ../소개페이지/index.html |
| `/jobs/` | ../정보보호_직무역량체계/index.html |

## 배포

### 방법 A. Vercel 대시보드에서 Git 연동 (권장)
1. https://vercel.com/new 에서 `yunakim-kr/test` 저장소를 Import
2. **Branch**: `main`
3. **Root Directory**: `vercel-sites`
4. Framework Preset은 "Other"로 두고 Deploy

이후 `main` 브랜치에 push할 때마다 자동으로 재배포됩니다.

### 방법 B. CLI로 직접 배포
```
cd vercel-sites
npx vercel        # 미리보기 배포 (최초 실행 시 로그인 필요)
npx vercel --prod # 운영 배포
```

## 원본 수정 시
원본(`../소개페이지/index.html`, `../정보보호_직무역량체계/index.html`)을 수정하면
이 폴더의 `public/intro/index.html`, `public/jobs/index.html`로 다시 복사해야 반영됩니다.
