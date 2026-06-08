배포 방법
1) GitHub 저장소에서 기존 index.html을 이 폴더의 index.html로 교체하세요.
2) styles.css, app.js 파일을 index.html과 같은 위치에 업로드하세요.
3) Netlify/GitHub Pages는 루트의 index.html을 자동으로 읽습니다.

수정 내용
- 대시보드 KPI 카드 크기/정렬 개선
- 월간 집계 미노출 오류 수정: MO_KPI → KPI_DEFS
- 구글시트 데이터 로딩 병렬화 및 5분 캐시 적용
- 엑셀 라이브러리 첫 화면 로딩 제외, 다운로드 시점 로딩
- 타이어앤 판매 평균 산식 설명 보강
- 최하 판매 매장 KPI 추가
