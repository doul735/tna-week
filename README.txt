TNA 상단 로고/타이틀 정렬 + 월별 선택 활성/비활성 패치입니다.

수정 범위
1. 상단 타이틀 왼쪽에 46x46 로고가 붙도록 정렬
2. 파비콘/OG 이미지는 건드리지 않음
3. 조회 기준이 주간 조회일 때: 주간 선택 활성, 월별 선택 비활성
4. 조회 기준이 월별 조회일 때: 월별 선택 활성, 주간 선택 비활성
5. 기존 차트/KPI/테이블 디자인은 그대로 유지

업로드할 파일
- header-symbol.png
- tna-fix.css
- tna-fix.js

GitHub 적용 방법
1. 이 압축파일 안의 3개 파일을 GitHub 저장소 루트에 업로드하세요.
   루트는 index.html, app.js, styles.css가 있는 같은 위치입니다.

2. index.html에서 기존 styles.css 아래에 이 한 줄을 추가하세요.

<link rel="stylesheet" href="./tna-fix.css?v=20260622-header1" />

3. index.html에서 기존 app.js 아래에 이 한 줄을 추가하세요.

<script src="./tna-fix.js?v=20260622-header1" defer></script>

주의
- favicon.ico, favicon-32x32.png, apple-touch-icon.png는 건드리지 마세요.
- og-image.png도 그대로 두세요.
- header-symbol.png만 이번에 업로드한 46x46 로고로 교체됩니다.

적용 후 확인
- GitHub 저장 후 1~3분 기다리기
- 사이트에서 Ctrl + F5 강력 새로고침
- 그래도 이전 화면이면 시크릿 창에서 확인
