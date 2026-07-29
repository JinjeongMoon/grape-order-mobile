# 그랑포도 주문서

모바일 손님용 포도 주문서입니다. 상품을 장바구니에 담으면 총액이 자동 계산되고, 주문자 정보는 Google Apps Script 웹앱을 통해 구글시트에 저장됩니다.

## 기능

- 상품 5개 이름/가격 관리자 입력
- 1박스 단위 수량 선택
- 장바구니 총액 자동 계산
- 이름, 전화번호, 요청사항 입력
- 계좌정보 복사 버튼과 토스트 안내
- 주문 중복 전송 방지
- 관리자 비밀번호 보호
- Google Apps Script 웹앱 URL로 주문 저장

## 개발

```bash
npm install
npm run dev
```

## 빌드

```bash
npm run build
```

빌드 결과는 `dist/`에 생성됩니다.

## Google Apps Script 연결

현재 기본 연결 URL은 `app/page.tsx`의 `defaultSettings.sheetEndpoint`에 들어 있습니다.

관리자 화면에서도 `Google Apps Script 웹앱 URL`을 수정한 뒤 `설정 저장하고 공개 링크 만들기`를 누르면, 해당 설정이 포함된 공유 링크를 만들 수 있습니다.

## GitHub Pages 배포

`.github/workflows/deploy-pages.yml`이 포함되어 있어 `main` 브랜치에 push하면 GitHub Actions로 자동 빌드됩니다.

GitHub 저장소에서 한 번만 설정하세요.

1. `Settings`로 이동
2. `Pages` 메뉴 선택
3. `Build and deployment`의 Source를 `GitHub Actions`로 선택

예상 주소:

```text
https://jinjeongmoon.github.io/grape-order-mobile/
```

## Vercel 또는 Netlify 배포

GitHub 저장소를 Vercel 또는 Netlify에 연결하면 됩니다.

- Build command: `npm run build`
- Output directory: `dist`
