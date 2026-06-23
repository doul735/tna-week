TNA 대시보드 오류 수정본입니다.

수정한 문제
1. index.html에서 page-shell/header 구조가 깨져 있던 문제 수정
2. app.js 맨 아래에 들어간 HTML <script> 태그 제거
3. styles.css 하단에 중복/충돌하던 헤더 보정 CSS 정리
4. 상단 로고를 제목 왼쪽에 46x46 크기로 정렬
5. 주간/월별 조회 선택에 따라 주간 선택/월별 선택이 정상적으로 활성·비활성되도록 유지
6. 캐시 문제를 줄이기 위해 index.html의 app.js/styles.css 버전을 20260623-fix1로 변경

적용 방법
1. GitHub 저장소에서 기존 index.html, app.js, styles.css를 각각 이 파일들로 교체하세요.
2. header-symbol.png, favicon.ico, favicon-32x32.ico, apple-touch-icon.ico, og-image.png는 기존 파일 그대로 두세요.
3. 저장 후 GitHub Pages 반영까지 1~3분 기다리세요.
4. 사이트에서 Ctrl + Shift + R 또는 Ctrl + F5로 강력 새로고침하세요.

주의
- CSS 파일 안에는 <link ...> 태그를 넣으면 안 됩니다.
- JS 파일 안에는 <script ...> 태그를 넣으면 안 됩니다.
- index.html에만 <link>와 <script> 태그가 들어갑니다.
