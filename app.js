// ─── 데이터 ───────────────────────────────────────────────────────
// 구글 시트 연동 시 이 URL을 Apps Script 웹 앱 URL로 교체하세요
const DATA_URL = 'https://script.google.com/macros/s/AKfycbw_TCR45muWiseITDdxHo_sYPKYxLS5CgRi_1LCouEgrapDkMQ7VE-HAj8zURoI2Uc/exec';
const TIRE_DATA_URL = DATA_URL + '?type=tire';

function cacheBustUrl(url) {
  return url + (url.includes('?') ? '&' : '?') + 'v=' + Date.now();
}

// 초기 로딩 속도 개선: 같은 세션에서는 캐시 데이터를 먼저 사용하고, 최신 조회 때만 강제 새로고침
async function fetchJsonWithCache(url, cacheKey, ttlMs = 5 * 60 * 1000, forceFresh = false) {
  const now = Date.now();

  if (!forceFresh) {
    try {
      const cached = localStorage.getItem(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.ts && Array.isArray(parsed.data) && now - parsed.ts < ttlMs) {
          return parsed.data;
        }
      }
    } catch (e) {
      console.warn('캐시 읽기 실패:', e);
    }
  }

  const res = await fetch(forceFresh ? cacheBustUrl(url) : url);
  if (!res.ok) throw new Error('데이터 요청 실패: ' + res.status);
  const data = await res.json();

  try {
    localStorage.setItem(cacheKey, JSON.stringify({ ts: now, data }));
  } catch (e) {
    console.warn('캐시 저장 실패:', e);
  }

  return data;
}

// XLSX는 다운로드할 때만 불러와서 첫 화면 로딩을 가볍게 유지
function ensureXlsx() {
  if (window.XLSX) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = resolve;
    script.onerror = () => reject(new Error('엑셀 라이브러리 로딩 실패'));
    document.head.appendChild(script);
  });
}


const RAW = [{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"한남점","team":"서울PC/LT팀","views":551,"conn":65,"miss":47,"res_in":33,"res_req":7,"review":3,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"방배점","team":"서울PC/LT팀","views":341,"conn":48,"miss":13,"res_in":22,"res_req":8,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"서초점","team":"서울PC/LT팀","views":409,"conn":77,"miss":21,"res_in":31,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"구리점","team":"서울PC/LT팀","views":340,"conn":59,"miss":18,"res_in":20,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"광교신도시점","team":"경기중부PC/LT팀","views":442,"conn":69,"miss":40,"res_in":59,"res_req":12,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"관양점","team":"경기중부PC/LT팀","views":173,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"홈플러스송도점","team":"경인PC/LT팀","views":584,"conn":112,"miss":116,"res_in":98,"res_req":16,"review":6,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"평택지산점","team":"경기중부PC/LT팀","views":145,"conn":12,"miss":6,"res_in":9,"res_req":1,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"오창점","team":"대전충청PC/LT팀","views":298,"conn":45,"miss":6,"res_in":20,"res_req":1,"review":14,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"서세종점","team":"대전충청PC/LT팀","views":401,"conn":66,"miss":12,"res_in":43,"res_req":10,"review":18,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"대전점","team":"대전충청PC/LT팀","views":280,"conn":25,"miss":13,"res_in":12,"res_req":2,"review":16,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"대전서구점","team":"대전충청PC/LT팀","views":334,"conn":48,"miss":15,"res_in":25,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"송천점","team":"광주전라PC/LT팀","views":366,"conn":64,"miss":6,"res_in":26,"res_req":3,"review":6,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"전주역점","team":"광주전라PC/LT팀","views":171,"conn":22,"miss":5,"res_in":13,"res_req":2,"review":29,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"남대구IC점","team":"대구경북PC/LT팀","views":225,"conn":25,"miss":9,"res_in":11,"res_req":2,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"수성점","team":"대구경북PC/LT팀","views":312,"conn":40,"miss":10,"res_in":16,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"부산거제점","team":"부산울산PC/LT팀","views":335,"conn":30,"miss":20,"res_in":33,"res_req":9,"review":6,"chat":0},{"yr":2025,"mo":5,"wk":1,"start":"2025-04-28","end":"2025-05-04","store":"부산수영점","team":"부산울산PC/LT팀","views":252,"conn":24,"miss":7,"res_in":11,"res_req":0,"review":8,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"한남점","team":"서울PC/LT팀","views":574,"conn":65,"miss":50,"res_in":22,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"방배점","team":"서울PC/LT팀","views":364,"conn":68,"miss":22,"res_in":26,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"서초점","team":"서울PC/LT팀","views":392,"conn":49,"miss":22,"res_in":26,"res_req":4,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"구리점","team":"서울PC/LT팀","views":323,"conn":30,"miss":14,"res_in":25,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"광교신도시점","team":"경기중부PC/LT팀","views":343,"conn":48,"miss":12,"res_in":35,"res_req":7,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"관양점","team":"경기중부PC/LT팀","views":156,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"홈플러스송도점","team":"경인PC/LT팀","views":442,"conn":75,"miss":75,"res_in":40,"res_req":8,"review":3,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"평택지산점","team":"경기중부PC/LT팀","views":210,"conn":17,"miss":13,"res_in":16,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"오창점","team":"대전충청PC/LT팀","views":377,"conn":40,"miss":6,"res_in":33,"res_req":7,"review":18,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"서세종점","team":"대전충청PC/LT팀","views":369,"conn":58,"miss":14,"res_in":39,"res_req":7,"review":21,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"대전점","team":"대전충청PC/LT팀","views":268,"conn":31,"miss":19,"res_in":20,"res_req":2,"review":11,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"대전서구점","team":"대전충청PC/LT팀","views":356,"conn":54,"miss":15,"res_in":21,"res_req":6,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"송천점","team":"광주전라PC/LT팀","views":325,"conn":40,"miss":8,"res_in":20,"res_req":1,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"전주역점","team":"광주전라PC/LT팀","views":179,"conn":26,"miss":9,"res_in":17,"res_req":2,"review":12,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"남대구IC점","team":"대구경북PC/LT팀","views":194,"conn":14,"miss":6,"res_in":13,"res_req":3,"review":8,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"수성점","team":"대구경북PC/LT팀","views":218,"conn":37,"miss":10,"res_in":14,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"부산거제점","team":"부산울산PC/LT팀","views":294,"conn":24,"miss":1,"res_in":23,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":2,"start":"2025-05-05","end":"2025-05-11","store":"부산수영점","team":"부산울산PC/LT팀","views":227,"conn":29,"miss":6,"res_in":6,"res_req":0,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"한남점","team":"서울PC/LT팀","views":500,"conn":60,"miss":27,"res_in":14,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"방배점","team":"서울PC/LT팀","views":341,"conn":48,"miss":15,"res_in":23,"res_req":3,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"서초점","team":"서울PC/LT팀","views":402,"conn":57,"miss":21,"res_in":26,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"구리점","team":"서울PC/LT팀","views":431,"conn":40,"miss":6,"res_in":23,"res_req":1,"review":7,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"광교신도시점","team":"경기중부PC/LT팀","views":365,"conn":49,"miss":30,"res_in":49,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"관양점","team":"경기중부PC/LT팀","views":167,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"홈플러스송도점","team":"경인PC/LT팀","views":491,"conn":115,"miss":84,"res_in":55,"res_req":6,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"평택지산점","team":"경기중부PC/LT팀","views":154,"conn":10,"miss":16,"res_in":10,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"오창점","team":"대전충청PC/LT팀","views":323,"conn":53,"miss":1,"res_in":28,"res_req":8,"review":18,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"서세종점","team":"대전충청PC/LT팀","views":365,"conn":53,"miss":9,"res_in":37,"res_req":10,"review":14,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"대전점","team":"대전충청PC/LT팀","views":227,"conn":38,"miss":15,"res_in":10,"res_req":1,"review":19,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"대전서구점","team":"대전충청PC/LT팀","views":371,"conn":52,"miss":15,"res_in":14,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"송천점","team":"광주전라PC/LT팀","views":216,"conn":40,"miss":3,"res_in":10,"res_req":0,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"전주역점","team":"광주전라PC/LT팀","views":172,"conn":26,"miss":1,"res_in":14,"res_req":2,"review":14,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"남대구IC점","team":"대구경북PC/LT팀","views":146,"conn":17,"miss":3,"res_in":6,"res_req":1,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"수성점","team":"대구경북PC/LT팀","views":279,"conn":44,"miss":6,"res_in":18,"res_req":4,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"부산거제점","team":"부산울산PC/LT팀","views":311,"conn":21,"miss":5,"res_in":34,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":3,"start":"2025-05-12","end":"2025-05-18","store":"부산수영점","team":"부산울산PC/LT팀","views":217,"conn":19,"miss":3,"res_in":15,"res_req":2,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"한남점","team":"서울PC/LT팀","views":480,"conn":48,"miss":16,"res_in":21,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"방배점","team":"서울PC/LT팀","views":362,"conn":53,"miss":17,"res_in":24,"res_req":5,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"서초점","team":"서울PC/LT팀","views":431,"conn":67,"miss":15,"res_in":29,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"구리점","team":"서울PC/LT팀","views":358,"conn":43,"miss":7,"res_in":22,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"광교신도시점","team":"경기중부PC/LT팀","views":403,"conn":56,"miss":30,"res_in":37,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"관양점","team":"경기중부PC/LT팀","views":139,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"홈플러스송도점","team":"경인PC/LT팀","views":424,"conn":62,"miss":56,"res_in":35,"res_req":3,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"평택지산점","team":"경기중부PC/LT팀","views":190,"conn":23,"miss":14,"res_in":11,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"오창점","team":"대전충청PC/LT팀","views":302,"conn":50,"miss":5,"res_in":26,"res_req":6,"review":22,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"서세종점","team":"대전충청PC/LT팀","views":305,"conn":48,"miss":8,"res_in":33,"res_req":8,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"대전점","team":"대전충청PC/LT팀","views":300,"conn":36,"miss":13,"res_in":15,"res_req":4,"review":30,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"대전서구점","team":"대전충청PC/LT팀","views":408,"conn":65,"miss":7,"res_in":23,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"송천점","team":"광주전라PC/LT팀","views":271,"conn":42,"miss":2,"res_in":11,"res_req":3,"review":4,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"전주역점","team":"광주전라PC/LT팀","views":142,"conn":13,"miss":3,"res_in":8,"res_req":2,"review":11,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"남대구IC점","team":"대구경북PC/LT팀","views":187,"conn":22,"miss":8,"res_in":7,"res_req":1,"review":7,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"수성점","team":"대구경북PC/LT팀","views":234,"conn":39,"miss":10,"res_in":10,"res_req":1,"review":7,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"부산거제점","team":"부산울산PC/LT팀","views":271,"conn":30,"miss":3,"res_in":10,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":4,"start":"2025-05-19","end":"2025-05-25","store":"부산수영점","team":"부산울산PC/LT팀","views":154,"conn":18,"miss":6,"res_in":10,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"한남점","team":"서울PC/LT팀","views":502,"conn":55,"miss":24,"res_in":24,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"방배점","team":"서울PC/LT팀","views":317,"conn":48,"miss":11,"res_in":34,"res_req":5,"review":3,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"서초점","team":"서울PC/LT팀","views":435,"conn":68,"miss":31,"res_in":33,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"구리점","team":"서울PC/LT팀","views":385,"conn":46,"miss":14,"res_in":27,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"광교신도시점","team":"경기중부PC/LT팀","views":426,"conn":56,"miss":9,"res_in":46,"res_req":11,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"관양점","team":"경기중부PC/LT팀","views":138,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"홈플러스송도점","team":"경인PC/LT팀","views":558,"conn":87,"miss":121,"res_in":57,"res_req":8,"review":6,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"평택지산점","team":"경기중부PC/LT팀","views":191,"conn":11,"miss":13,"res_in":15,"res_req":4,"review":8,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"오창점","team":"대전충청PC/LT팀","views":288,"conn":38,"miss":5,"res_in":19,"res_req":4,"review":11,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"서세종점","team":"대전충청PC/LT팀","views":366,"conn":67,"miss":7,"res_in":31,"res_req":11,"review":11,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"대전점","team":"대전충청PC/LT팀","views":305,"conn":52,"miss":14,"res_in":23,"res_req":3,"review":27,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"대전서구점","team":"대전충청PC/LT팀","views":363,"conn":52,"miss":11,"res_in":26,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"송천점","team":"광주전라PC/LT팀","views":266,"conn":56,"miss":3,"res_in":24,"res_req":1,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"전주역점","team":"광주전라PC/LT팀","views":140,"conn":21,"miss":2,"res_in":3,"res_req":0,"review":11,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"남대구IC점","team":"대구경북PC/LT팀","views":174,"conn":21,"miss":7,"res_in":14,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"수성점","team":"대구경북PC/LT팀","views":250,"conn":39,"miss":5,"res_in":13,"res_req":2,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"부산거제점","team":"부산울산PC/LT팀","views":587,"conn":28,"miss":3,"res_in":22,"res_req":2,"review":5,"chat":0},{"yr":2025,"mo":5,"wk":5,"start":"2025-05-26","end":"2025-06-01","store":"부산수영점","team":"부산울산PC/LT팀","views":205,"conn":21,"miss":13,"res_in":5,"res_req":1,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"한남점","team":"서울PC/LT팀","views":530,"conn":49,"miss":34,"res_in":30,"res_req":5,"review":3,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"방배점","team":"서울PC/LT팀","views":320,"conn":53,"miss":19,"res_in":19,"res_req":6,"review":3,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"서초점","team":"서울PC/LT팀","views":395,"conn":43,"miss":40,"res_in":28,"res_req":6,"review":6,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"구리점","team":"서울PC/LT팀","views":349,"conn":34,"miss":22,"res_in":19,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"광교신도시점","team":"경기중부PC/LT팀","views":406,"conn":43,"miss":15,"res_in":39,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"관양점","team":"경기중부PC/LT팀","views":136,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"홈플러스송도점","team":"경인PC/LT팀","views":573,"conn":89,"miss":68,"res_in":51,"res_req":4,"review":7,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"평택지산점","team":"경기중부PC/LT팀","views":185,"conn":11,"miss":16,"res_in":19,"res_req":6,"review":5,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"오창점","team":"대전충청PC/LT팀","views":336,"conn":47,"miss":6,"res_in":24,"res_req":3,"review":18,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"서세종점","team":"대전충청PC/LT팀","views":482,"conn":49,"miss":16,"res_in":44,"res_req":11,"review":14,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"대전점","team":"대전충청PC/LT팀","views":245,"conn":28,"miss":14,"res_in":12,"res_req":4,"review":10,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"대전서구점","team":"대전충청PC/LT팀","views":337,"conn":55,"miss":8,"res_in":21,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"송천점","team":"광주전라PC/LT팀","views":345,"conn":56,"miss":18,"res_in":28,"res_req":6,"review":15,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"전주역점","team":"광주전라PC/LT팀","views":175,"conn":21,"miss":5,"res_in":14,"res_req":3,"review":12,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"남대구IC점","team":"대구경북PC/LT팀","views":199,"conn":26,"miss":4,"res_in":13,"res_req":4,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"수성점","team":"대구경북PC/LT팀","views":305,"conn":42,"miss":4,"res_in":19,"res_req":5,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"부산거제점","team":"부산울산PC/LT팀","views":1010,"conn":27,"miss":5,"res_in":27,"res_req":5,"review":22,"chat":0},{"yr":2025,"mo":6,"wk":1,"start":"2025-06-02","end":"2025-06-08","store":"부산수영점","team":"부산울산PC/LT팀","views":157,"conn":12,"miss":15,"res_in":1,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"한남점","team":"서울PC/LT팀","views":598,"conn":55,"miss":34,"res_in":29,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"방배점","team":"서울PC/LT팀","views":378,"conn":62,"miss":19,"res_in":28,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"서초점","team":"서울PC/LT팀","views":458,"conn":80,"miss":29,"res_in":34,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"구리점","team":"서울PC/LT팀","views":298,"conn":29,"miss":27,"res_in":15,"res_req":4,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"광교신도시점","team":"경기중부PC/LT팀","views":441,"conn":57,"miss":47,"res_in":7,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"관양점","team":"경기중부PC/LT팀","views":156,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"홈플러스송도점","team":"경인PC/LT팀","views":701,"conn":97,"miss":144,"res_in":90,"res_req":10,"review":6,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"평택지산점","team":"경기중부PC/LT팀","views":179,"conn":20,"miss":10,"res_in":14,"res_req":2,"review":6,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"오창점","team":"대전충청PC/LT팀","views":278,"conn":34,"miss":4,"res_in":33,"res_req":6,"review":11,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"서세종점","team":"대전충청PC/LT팀","views":376,"conn":65,"miss":12,"res_in":29,"res_req":4,"review":13,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"대전점","team":"대전충청PC/LT팀","views":242,"conn":40,"miss":11,"res_in":14,"res_req":5,"review":19,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"대전서구점","team":"대전충청PC/LT팀","views":371,"conn":60,"miss":10,"res_in":21,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"송천점","team":"광주전라PC/LT팀","views":326,"conn":82,"miss":11,"res_in":29,"res_req":2,"review":7,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"전주역점","team":"광주전라PC/LT팀","views":155,"conn":26,"miss":4,"res_in":10,"res_req":2,"review":19,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"남대구IC점","team":"대구경북PC/LT팀","views":224,"conn":30,"miss":7,"res_in":3,"res_req":1,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"수성점","team":"대구경북PC/LT팀","views":281,"conn":38,"miss":5,"res_in":4,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"부산거제점","team":"부산울산PC/LT팀","views":806,"conn":39,"miss":10,"res_in":32,"res_req":6,"review":17,"chat":0},{"yr":2025,"mo":6,"wk":2,"start":"2025-06-09","end":"2025-06-15","store":"부산수영점","team":"부산울산PC/LT팀","views":175,"conn":15,"miss":11,"res_in":13,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"한남점","team":"서울PC/LT팀","views":697,"conn":66,"miss":37,"res_in":40,"res_req":11,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"방배점","team":"서울PC/LT팀","views":398,"conn":70,"miss":36,"res_in":26,"res_req":2,"review":6,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"서초점","team":"서울PC/LT팀","views":500,"conn":85,"miss":17,"res_in":37,"res_req":6,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"구리점","team":"서울PC/LT팀","views":338,"conn":44,"miss":18,"res_in":17,"res_req":3,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"광교신도시점","team":"경기중부PC/LT팀","views":445,"conn":57,"miss":22,"res_in":13,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"관양점","team":"경기중부PC/LT팀","views":144,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"홈플러스송도점","team":"경인PC/LT팀","views":617,"conn":92,"miss":70,"res_in":73,"res_req":11,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"평택지산점","team":"경기중부PC/LT팀","views":257,"conn":25,"miss":15,"res_in":19,"res_req":3,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"오창점","team":"대전충청PC/LT팀","views":290,"conn":49,"miss":5,"res_in":19,"res_req":3,"review":20,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"서세종점","team":"대전충청PC/LT팀","views":421,"conn":51,"miss":22,"res_in":44,"res_req":11,"review":9,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"대전점","team":"대전충청PC/LT팀","views":248,"conn":46,"miss":24,"res_in":11,"res_req":1,"review":17,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"대전서구점","team":"대전충청PC/LT팀","views":328,"conn":46,"miss":12,"res_in":18,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"송천점","team":"광주전라PC/LT팀","views":336,"conn":59,"miss":8,"res_in":32,"res_req":3,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"전주역점","team":"광주전라PC/LT팀","views":190,"conn":34,"miss":3,"res_in":11,"res_req":4,"review":8,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"남대구IC점","team":"대구경북PC/LT팀","views":187,"conn":32,"miss":2,"res_in":2,"res_req":0,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"수성점","team":"대구경북PC/LT팀","views":284,"conn":40,"miss":11,"res_in":4,"res_req":1,"review":4,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"부산거제점","team":"부산울산PC/LT팀","views":364,"conn":49,"miss":13,"res_in":20,"res_req":2,"review":8,"chat":0},{"yr":2025,"mo":6,"wk":3,"start":"2025-06-16","end":"2025-06-22","store":"부산수영점","team":"부산울산PC/LT팀","views":232,"conn":27,"miss":13,"res_in":12,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"한남점","team":"서울PC/LT팀","views":637,"conn":73,"miss":45,"res_in":44,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"방배점","team":"서울PC/LT팀","views":424,"conn":66,"miss":16,"res_in":40,"res_req":8,"review":8,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"서초점","team":"서울PC/LT팀","views":402,"conn":63,"miss":13,"res_in":34,"res_req":3,"review":6,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"구리점","team":"서울PC/LT팀","views":393,"conn":57,"miss":17,"res_in":30,"res_req":8,"review":2,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"광교신도시점","team":"경기중부PC/LT팀","views":415,"conn":53,"miss":22,"res_in":42,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"관양점","team":"경기중부PC/LT팀","views":139,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"홈플러스송도점","team":"경인PC/LT팀","views":623,"conn":107,"miss":135,"res_in":71,"res_req":15,"review":3,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"평택지산점","team":"경기중부PC/LT팀","views":221,"conn":29,"miss":4,"res_in":18,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"오창점","team":"대전충청PC/LT팀","views":270,"conn":47,"miss":1,"res_in":20,"res_req":0,"review":14,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"서세종점","team":"대전충청PC/LT팀","views":398,"conn":58,"miss":9,"res_in":49,"res_req":13,"review":10,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"대전점","team":"대전충청PC/LT팀","views":303,"conn":45,"miss":14,"res_in":32,"res_req":4,"review":27,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"대전서구점","team":"대전충청PC/LT팀","views":374,"conn":55,"miss":26,"res_in":27,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"송천점","team":"광주전라PC/LT팀","views":301,"conn":56,"miss":8,"res_in":23,"res_req":3,"review":7,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"전주역점","team":"광주전라PC/LT팀","views":178,"conn":24,"miss":5,"res_in":11,"res_req":1,"review":19,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"남대구IC점","team":"대구경북PC/LT팀","views":228,"conn":32,"miss":7,"res_in":15,"res_req":1,"review":3,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"수성점","team":"대구경북PC/LT팀","views":231,"conn":37,"miss":6,"res_in":14,"res_req":0,"review":6,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"부산거제점","team":"부산울산PC/LT팀","views":300,"conn":32,"miss":7,"res_in":24,"res_req":4,"review":13,"chat":0},{"yr":2025,"mo":6,"wk":4,"start":"2025-06-23","end":"2025-06-29","store":"부산수영점","team":"부산울산PC/LT팀","views":217,"conn":28,"miss":9,"res_in":11,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"한남점","team":"서울PC/LT팀","views":574,"conn":71,"miss":36,"res_in":38,"res_req":5,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"방배점","team":"서울PC/LT팀","views":348,"conn":40,"miss":16,"res_in":26,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"서초점","team":"서울PC/LT팀","views":457,"conn":64,"miss":16,"res_in":35,"res_req":6,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"구리점","team":"서울PC/LT팀","views":322,"conn":60,"miss":18,"res_in":27,"res_req":4,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"광교신도시점","team":"경기중부PC/LT팀","views":413,"conn":71,"miss":22,"res_in":47,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"관양점","team":"경기중부PC/LT팀","views":154,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"홈플러스송도점","team":"경인PC/LT팀","views":580,"conn":95,"miss":79,"res_in":78,"res_req":14,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"평택지산점","team":"경기중부PC/LT팀","views":253,"conn":29,"miss":13,"res_in":13,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"오창점","team":"대전충청PC/LT팀","views":326,"conn":35,"miss":7,"res_in":25,"res_req":5,"review":25,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"서세종점","team":"대전충청PC/LT팀","views":483,"conn":94,"miss":37,"res_in":41,"res_req":8,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"대전점","team":"대전충청PC/LT팀","views":316,"conn":41,"miss":18,"res_in":21,"res_req":4,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"대전서구점","team":"대전충청PC/LT팀","views":353,"conn":63,"miss":28,"res_in":29,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"송천점","team":"광주전라PC/LT팀","views":305,"conn":71,"miss":7,"res_in":18,"res_req":4,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"전주역점","team":"광주전라PC/LT팀","views":160,"conn":27,"miss":0,"res_in":15,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"남대구IC점","team":"대구경북PC/LT팀","views":230,"conn":27,"miss":7,"res_in":17,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"수성점","team":"대구경북PC/LT팀","views":272,"conn":27,"miss":21,"res_in":20,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"부산거제점","team":"부산울산PC/LT팀","views":324,"conn":28,"miss":12,"res_in":24,"res_req":4,"review":8,"chat":0},{"yr":2025,"mo":7,"wk":1,"start":"2025-06-30","end":"2025-07-06","store":"부산수영점","team":"부산울산PC/LT팀","views":312,"conn":24,"miss":17,"res_in":20,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"한남점","team":"서울PC/LT팀","views":619,"conn":71,"miss":55,"res_in":37,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"방배점","team":"서울PC/LT팀","views":446,"conn":49,"miss":17,"res_in":39,"res_req":7,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"서초점","team":"서울PC/LT팀","views":524,"conn":68,"miss":25,"res_in":50,"res_req":8,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"구리점","team":"서울PC/LT팀","views":384,"conn":51,"miss":13,"res_in":26,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"광교신도시점","team":"경기중부PC/LT팀","views":472,"conn":56,"miss":35,"res_in":59,"res_req":15,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"관양점","team":"경기중부PC/LT팀","views":182,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"홈플러스송도점","team":"경인PC/LT팀","views":606,"conn":88,"miss":91,"res_in":77,"res_req":9,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"평택지산점","team":"경기중부PC/LT팀","views":258,"conn":23,"miss":4,"res_in":13,"res_req":0,"review":4,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"오창점","team":"대전충청PC/LT팀","views":351,"conn":55,"miss":5,"res_in":28,"res_req":4,"review":20,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"서세종점","team":"대전충청PC/LT팀","views":507,"conn":79,"miss":9,"res_in":59,"res_req":14,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"대전점","team":"대전충청PC/LT팀","views":326,"conn":58,"miss":18,"res_in":33,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"대전서구점","team":"대전충청PC/LT팀","views":380,"conn":47,"miss":9,"res_in":20,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"송천점","team":"광주전라PC/LT팀","views":317,"conn":69,"miss":5,"res_in":32,"res_req":8,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"전주역점","team":"광주전라PC/LT팀","views":167,"conn":31,"miss":2,"res_in":10,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"남대구IC점","team":"대구경북PC/LT팀","views":312,"conn":43,"miss":9,"res_in":26,"res_req":4,"review":8,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"수성점","team":"대구경북PC/LT팀","views":292,"conn":39,"miss":5,"res_in":14,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"부산거제점","team":"부산울산PC/LT팀","views":296,"conn":28,"miss":9,"res_in":16,"res_req":5,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":2,"start":"2025-07-07","end":"2025-07-13","store":"부산수영점","team":"부산울산PC/LT팀","views":240,"conn":34,"miss":6,"res_in":6,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"한남점","team":"서울PC/LT팀","views":579,"conn":67,"miss":40,"res_in":32,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"방배점","team":"서울PC/LT팀","views":402,"conn":73,"miss":37,"res_in":34,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"서초점","team":"서울PC/LT팀","views":498,"conn":89,"miss":27,"res_in":38,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"구리점","team":"서울PC/LT팀","views":320,"conn":51,"miss":28,"res_in":31,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"광교신도시점","team":"경기중부PC/LT팀","views":511,"conn":73,"miss":20,"res_in":73,"res_req":15,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"관양점","team":"경기중부PC/LT팀","views":141,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"홈플러스송도점","team":"경인PC/LT팀","views":656,"conn":89,"miss":165,"res_in":73,"res_req":13,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"평택지산점","team":"경기중부PC/LT팀","views":228,"conn":23,"miss":4,"res_in":9,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"오창점","team":"대전충청PC/LT팀","views":372,"conn":45,"miss":10,"res_in":34,"res_req":7,"review":17,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"서세종점","team":"대전충청PC/LT팀","views":441,"conn":69,"miss":11,"res_in":46,"res_req":10,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"대전점","team":"대전충청PC/LT팀","views":305,"conn":45,"miss":33,"res_in":24,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"대전서구점","team":"대전충청PC/LT팀","views":378,"conn":38,"miss":17,"res_in":29,"res_req":7,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"송천점","team":"광주전라PC/LT팀","views":409,"conn":73,"miss":13,"res_in":36,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"전주역점","team":"광주전라PC/LT팀","views":162,"conn":30,"miss":6,"res_in":11,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"남대구IC점","team":"대구경북PC/LT팀","views":235,"conn":26,"miss":2,"res_in":13,"res_req":4,"review":9,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"수성점","team":"대구경북PC/LT팀","views":337,"conn":48,"miss":2,"res_in":19,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"부산거제점","team":"부산울산PC/LT팀","views":272,"conn":48,"miss":6,"res_in":21,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":3,"start":"2025-07-14","end":"2025-07-20","store":"부산수영점","team":"부산울산PC/LT팀","views":253,"conn":25,"miss":11,"res_in":10,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"한남점","team":"서울PC/LT팀","views":642,"conn":77,"miss":23,"res_in":37,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"방배점","team":"서울PC/LT팀","views":404,"conn":57,"miss":11,"res_in":23,"res_req":8,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"서초점","team":"서울PC/LT팀","views":502,"conn":75,"miss":15,"res_in":46,"res_req":9,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"구리점","team":"서울PC/LT팀","views":421,"conn":47,"miss":16,"res_in":25,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"광교신도시점","team":"경기중부PC/LT팀","views":572,"conn":57,"miss":39,"res_in":63,"res_req":11,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"관양점","team":"경기중부PC/LT팀","views":222,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"홈플러스송도점","team":"경인PC/LT팀","views":705,"conn":101,"miss":139,"res_in":71,"res_req":12,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"평택지산점","team":"경기중부PC/LT팀","views":288,"conn":24,"miss":21,"res_in":18,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"오창점","team":"대전충청PC/LT팀","views":354,"conn":53,"miss":15,"res_in":30,"res_req":7,"review":13,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"서세종점","team":"대전충청PC/LT팀","views":642,"conn":99,"miss":23,"res_in":76,"res_req":16,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"대전점","team":"대전충청PC/LT팀","views":370,"conn":57,"miss":21,"res_in":34,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"대전서구점","team":"대전충청PC/LT팀","views":463,"conn":67,"miss":31,"res_in":36,"res_req":8,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"송천점","team":"광주전라PC/LT팀","views":447,"conn":86,"miss":15,"res_in":30,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"전주역점","team":"광주전라PC/LT팀","views":231,"conn":45,"miss":9,"res_in":19,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"남대구IC점","team":"대구경북PC/LT팀","views":293,"conn":22,"miss":4,"res_in":13,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"수성점","team":"대구경북PC/LT팀","views":311,"conn":49,"miss":9,"res_in":17,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"부산거제점","team":"부산울산PC/LT팀","views":305,"conn":23,"miss":23,"res_in":21,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":4,"start":"2025-07-21","end":"2025-07-27","store":"부산수영점","team":"부산울산PC/LT팀","views":333,"conn":27,"miss":18,"res_in":19,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"한남점","team":"서울PC/LT팀","views":593,"conn":64,"miss":34,"res_in":33,"res_req":7,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"방배점","team":"서울PC/LT팀","views":404,"conn":54,"miss":32,"res_in":25,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"서초점","team":"서울PC/LT팀","views":463,"conn":60,"miss":14,"res_in":34,"res_req":7,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"구리점","team":"서울PC/LT팀","views":394,"conn":48,"miss":16,"res_in":27,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"광교신도시점","team":"경기중부PC/LT팀","views":640,"conn":78,"miss":60,"res_in":80,"res_req":13,"review":4,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"관양점","team":"경기중부PC/LT팀","views":241,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"홈플러스송도점","team":"경인PC/LT팀","views":807,"conn":141,"miss":163,"res_in":88,"res_req":14,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"평택지산점","team":"경기중부PC/LT팀","views":259,"conn":24,"miss":19,"res_in":12,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"오창점","team":"대전충청PC/LT팀","views":371,"conn":44,"miss":8,"res_in":24,"res_req":5,"review":9,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"서세종점","team":"대전충청PC/LT팀","views":601,"conn":97,"miss":19,"res_in":58,"res_req":11,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"대전점","team":"대전충청PC/LT팀","views":311,"conn":45,"miss":30,"res_in":18,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"대전서구점","team":"대전충청PC/LT팀","views":420,"conn":52,"miss":29,"res_in":25,"res_req":5,"review":3,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"송천점","team":"광주전라PC/LT팀","views":448,"conn":80,"miss":15,"res_in":47,"res_req":6,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"전주역점","team":"광주전라PC/LT팀","views":189,"conn":30,"miss":5,"res_in":14,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"남대구IC점","team":"대구경북PC/LT팀","views":311,"conn":39,"miss":13,"res_in":16,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"수성점","team":"대구경북PC/LT팀","views":393,"conn":48,"miss":12,"res_in":15,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"부산거제점","team":"부산울산PC/LT팀","views":369,"conn":28,"miss":30,"res_in":31,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":7,"wk":5,"start":"2025-07-28","end":"2025-08-03","store":"부산수영점","team":"부산울산PC/LT팀","views":286,"conn":32,"miss":9,"res_in":23,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"한남점","team":"서울PC/LT팀","views":667,"conn":82,"miss":26,"res_in":36,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"방배점","team":"서울PC/LT팀","views":391,"conn":58,"miss":18,"res_in":20,"res_req":6,"review":5,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"서초점","team":"서울PC/LT팀","views":495,"conn":68,"miss":18,"res_in":36,"res_req":5,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"구리점","team":"서울PC/LT팀","views":418,"conn":48,"miss":74,"res_in":32,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"광교신도시점","team":"경기중부PC/LT팀","views":614,"conn":79,"miss":26,"res_in":70,"res_req":15,"review":6,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"관양점","team":"경기중부PC/LT팀","views":222,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"홈플러스송도점","team":"경인PC/LT팀","views":721,"conn":88,"miss":178,"res_in":66,"res_req":11,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"평택지산점","team":"경기중부PC/LT팀","views":312,"conn":25,"miss":16,"res_in":18,"res_req":3,"review":6,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"오창점","team":"대전충청PC/LT팀","views":337,"conn":55,"miss":12,"res_in":41,"res_req":7,"review":17,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"서세종점","team":"대전충청PC/LT팀","views":567,"conn":79,"miss":18,"res_in":67,"res_req":15,"review":7,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"대전점","team":"대전충청PC/LT팀","views":372,"conn":53,"miss":27,"res_in":27,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"대전서구점","team":"대전충청PC/LT팀","views":491,"conn":56,"miss":43,"res_in":29,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"송천점","team":"광주전라PC/LT팀","views":392,"conn":69,"miss":6,"res_in":32,"res_req":5,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"전주역점","team":"광주전라PC/LT팀","views":227,"conn":36,"miss":3,"res_in":14,"res_req":1,"review":7,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"남대구IC점","team":"대구경북PC/LT팀","views":232,"conn":36,"miss":8,"res_in":15,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"수성점","team":"대구경북PC/LT팀","views":387,"conn":51,"miss":14,"res_in":20,"res_req":2,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"부산거제점","team":"부산울산PC/LT팀","views":336,"conn":25,"miss":24,"res_in":24,"res_req":5,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":1,"start":"2025-08-04","end":"2025-08-10","store":"부산수영점","team":"부산울산PC/LT팀","views":331,"conn":42,"miss":12,"res_in":20,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"한남점","team":"서울PC/LT팀","views":698,"conn":77,"miss":56,"res_in":37,"res_req":1,"review":5,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"방배점","team":"서울PC/LT팀","views":404,"conn":71,"miss":23,"res_in":28,"res_req":5,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"서초점","team":"서울PC/LT팀","views":603,"conn":96,"miss":23,"res_in":27,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"구리점","team":"서울PC/LT팀","views":367,"conn":44,"miss":26,"res_in":28,"res_req":6,"review":4,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"광교신도시점","team":"경기중부PC/LT팀","views":480,"conn":69,"miss":24,"res_in":48,"res_req":12,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"관양점","team":"경기중부PC/LT팀","views":227,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"홈플러스송도점","team":"경인PC/LT팀","views":854,"conn":97,"miss":222,"res_in":72,"res_req":12,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"평택지산점","team":"경기중부PC/LT팀","views":205,"conn":21,"miss":12,"res_in":14,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"오창점","team":"대전충청PC/LT팀","views":312,"conn":32,"miss":8,"res_in":33,"res_req":8,"review":9,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"서세종점","team":"대전충청PC/LT팀","views":526,"conn":61,"miss":16,"res_in":51,"res_req":14,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"대전점","team":"대전충청PC/LT팀","views":314,"conn":40,"miss":27,"res_in":23,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"대전서구점","team":"대전충청PC/LT팀","views":399,"conn":53,"miss":19,"res_in":25,"res_req":4,"review":5,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"송천점","team":"광주전라PC/LT팀","views":451,"conn":90,"miss":23,"res_in":22,"res_req":4,"review":9,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"전주역점","team":"광주전라PC/LT팀","views":228,"conn":46,"miss":4,"res_in":13,"res_req":1,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"남대구IC점","team":"대구경북PC/LT팀","views":236,"conn":22,"miss":14,"res_in":10,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"수성점","team":"대구경북PC/LT팀","views":302,"conn":45,"miss":9,"res_in":22,"res_req":4,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"부산거제점","team":"부산울산PC/LT팀","views":436,"conn":38,"miss":33,"res_in":31,"res_req":9,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":2,"start":"2025-08-11","end":"2025-08-17","store":"부산수영점","team":"부산울산PC/LT팀","views":275,"conn":27,"miss":12,"res_in":16,"res_req":1,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"한남점","team":"서울PC/LT팀","views":664,"conn":70,"miss":53,"res_in":46,"res_req":12,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"방배점","team":"서울PC/LT팀","views":378,"conn":53,"miss":19,"res_in":22,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"서초점","team":"서울PC/LT팀","views":433,"conn":64,"miss":12,"res_in":35,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"구리점","team":"서울PC/LT팀","views":321,"conn":43,"miss":30,"res_in":22,"res_req":2,"review":4,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"광교신도시점","team":"경기중부PC/LT팀","views":456,"conn":62,"miss":23,"res_in":55,"res_req":11,"review":7,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"관양점","team":"경기중부PC/LT팀","views":277,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"홈플러스송도점","team":"경인PC/LT팀","views":714,"conn":73,"miss":200,"res_in":73,"res_req":14,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"평택지산점","team":"경기중부PC/LT팀","views":226,"conn":25,"miss":3,"res_in":15,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"오창점","team":"대전충청PC/LT팀","views":409,"conn":54,"miss":13,"res_in":36,"res_req":9,"review":14,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"서세종점","team":"대전충청PC/LT팀","views":545,"conn":85,"miss":27,"res_in":47,"res_req":7,"review":10,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"대전점","team":"대전충청PC/LT팀","views":373,"conn":43,"miss":34,"res_in":32,"res_req":8,"review":16,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"대전서구점","team":"대전충청PC/LT팀","views":371,"conn":44,"miss":27,"res_in":24,"res_req":1,"review":8,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"송천점","team":"광주전라PC/LT팀","views":360,"conn":80,"miss":13,"res_in":26,"res_req":5,"review":11,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"전주역점","team":"광주전라PC/LT팀","views":212,"conn":31,"miss":2,"res_in":11,"res_req":2,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"남대구IC점","team":"대구경북PC/LT팀","views":210,"conn":31,"miss":10,"res_in":12,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"수성점","team":"대구경북PC/LT팀","views":379,"conn":54,"miss":8,"res_in":19,"res_req":2,"review":6,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"부산거제점","team":"부산울산PC/LT팀","views":364,"conn":40,"miss":25,"res_in":12,"res_req":1,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":3,"start":"2025-08-18","end":"2025-08-24","store":"부산수영점","team":"부산울산PC/LT팀","views":248,"conn":32,"miss":11,"res_in":4,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"한남점","team":"서울PC/LT팀","views":706,"conn":55,"miss":36,"res_in":40,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"방배점","team":"서울PC/LT팀","views":430,"conn":54,"miss":25,"res_in":37,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"서초점","team":"서울PC/LT팀","views":490,"conn":66,"miss":13,"res_in":29,"res_req":5,"review":5,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"구리점","team":"서울PC/LT팀","views":407,"conn":53,"miss":15,"res_in":28,"res_req":3,"review":10,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"광교신도시점","team":"경기중부PC/LT팀","views":511,"conn":52,"miss":26,"res_in":67,"res_req":26,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"관양점","team":"경기중부PC/LT팀","views":216,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"홈플러스송도점","team":"경인PC/LT팀","views":630,"conn":90,"miss":152,"res_in":52,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"평택지산점","team":"경기중부PC/LT팀","views":261,"conn":23,"miss":7,"res_in":9,"res_req":2,"review":15,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"오창점","team":"대전충청PC/LT팀","views":325,"conn":49,"miss":10,"res_in":25,"res_req":2,"review":6,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"서세종점","team":"대전충청PC/LT팀","views":447,"conn":73,"miss":11,"res_in":44,"res_req":8,"review":4,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"대전점","team":"대전충청PC/LT팀","views":345,"conn":47,"miss":13,"res_in":32,"res_req":2,"review":12,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"대전서구점","team":"대전충청PC/LT팀","views":469,"conn":40,"miss":23,"res_in":23,"res_req":4,"review":19,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"송천점","team":"광주전라PC/LT팀","views":340,"conn":51,"miss":8,"res_in":26,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"전주역점","team":"광주전라PC/LT팀","views":151,"conn":31,"miss":5,"res_in":7,"res_req":0,"review":4,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"남대구IC점","team":"대구경북PC/LT팀","views":254,"conn":35,"miss":9,"res_in":11,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"수성점","team":"대구경북PC/LT팀","views":309,"conn":43,"miss":9,"res_in":17,"res_req":3,"review":7,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"부산거제점","team":"부산울산PC/LT팀","views":324,"conn":43,"miss":21,"res_in":18,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":8,"wk":4,"start":"2025-08-25","end":"2025-08-31","store":"부산수영점","team":"부산울산PC/LT팀","views":303,"conn":32,"miss":16,"res_in":20,"res_req":1,"review":5,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"한남점","team":"서울PC/LT팀","views":839,"conn":72,"miss":52,"res_in":3,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"방배점","team":"서울PC/LT팀","views":474,"conn":63,"miss":26,"res_in":32,"res_req":10,"review":3,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"서초점","team":"서울PC/LT팀","views":578,"conn":65,"miss":16,"res_in":37,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"구리점","team":"서울PC/LT팀","views":440,"conn":71,"miss":21,"res_in":19,"res_req":0,"review":3,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"광교신도시점","team":"경기중부PC/LT팀","views":490,"conn":77,"miss":18,"res_in":47,"res_req":13,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"관양점","team":"경기중부PC/LT팀","views":227,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"홈플러스송도점","team":"경인PC/LT팀","views":958,"conn":83,"miss":267,"res_in":89,"res_req":25,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"평택지산점","team":"경기중부PC/LT팀","views":306,"conn":33,"miss":7,"res_in":19,"res_req":1,"review":12,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"오창점","team":"대전충청PC/LT팀","views":362,"conn":47,"miss":18,"res_in":30,"res_req":4,"review":11,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"서세종점","team":"대전충청PC/LT팀","views":598,"conn":74,"miss":13,"res_in":61,"res_req":13,"review":5,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"대전점","team":"대전충청PC/LT팀","views":334,"conn":48,"miss":30,"res_in":19,"res_req":7,"review":16,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"대전서구점","team":"대전충청PC/LT팀","views":535,"conn":80,"miss":31,"res_in":36,"res_req":9,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"송천점","team":"광주전라PC/LT팀","views":434,"conn":58,"miss":4,"res_in":26,"res_req":1,"review":6,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"전주역점","team":"광주전라PC/LT팀","views":197,"conn":32,"miss":3,"res_in":12,"res_req":1,"review":5,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"남대구IC점","team":"대구경북PC/LT팀","views":331,"conn":41,"miss":12,"res_in":17,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"수성점","team":"대구경북PC/LT팀","views":381,"conn":66,"miss":16,"res_in":18,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"부산거제점","team":"부산울산PC/LT팀","views":400,"conn":29,"miss":34,"res_in":25,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":1,"start":"2025-09-01","end":"2025-09-07","store":"부산수영점","team":"부산울산PC/LT팀","views":275,"conn":21,"miss":16,"res_in":16,"res_req":4,"review":11,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"한남점","team":"서울PC/LT팀","views":855,"conn":98,"miss":47,"res_in":9,"res_req":1,"review":3,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"방배점","team":"서울PC/LT팀","views":512,"conn":80,"miss":36,"res_in":37,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"서초점","team":"서울PC/LT팀","views":667,"conn":107,"miss":24,"res_in":29,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"구리점","team":"서울PC/LT팀","views":476,"conn":67,"miss":28,"res_in":22,"res_req":2,"review":10,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"광교신도시점","team":"경기중부PC/LT팀","views":611,"conn":75,"miss":25,"res_in":72,"res_req":19,"review":6,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"관양점","team":"경기중부PC/LT팀","views":249,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"홈플러스송도점","team":"경인PC/LT팀","views":871,"conn":90,"miss":211,"res_in":86,"res_req":19,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"평택지산점","team":"경기중부PC/LT팀","views":297,"conn":24,"miss":17,"res_in":21,"res_req":3,"review":9,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"오창점","team":"대전충청PC/LT팀","views":392,"conn":61,"miss":12,"res_in":38,"res_req":3,"review":8,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"서세종점","team":"대전충청PC/LT팀","views":585,"conn":99,"miss":12,"res_in":60,"res_req":12,"review":8,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"대전점","team":"대전충청PC/LT팀","views":381,"conn":55,"miss":35,"res_in":22,"res_req":1,"review":8,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"대전서구점","team":"대전충청PC/LT팀","views":597,"conn":63,"miss":26,"res_in":38,"res_req":7,"review":7,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"송천점","team":"광주전라PC/LT팀","views":423,"conn":73,"miss":7,"res_in":36,"res_req":4,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"전주역점","team":"광주전라PC/LT팀","views":218,"conn":39,"miss":5,"res_in":9,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"남대구IC점","team":"대구경북PC/LT팀","views":255,"conn":34,"miss":6,"res_in":9,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"수성점","team":"대구경북PC/LT팀","views":387,"conn":36,"miss":13,"res_in":12,"res_req":0,"review":9,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"부산거제점","team":"부산울산PC/LT팀","views":297,"conn":36,"miss":18,"res_in":10,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":2,"start":"2025-09-08","end":"2025-09-14","store":"부산수영점","team":"부산울산PC/LT팀","views":259,"conn":27,"miss":7,"res_in":16,"res_req":4,"review":6,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"한남점","team":"서울PC/LT팀","views":791,"conn":67,"miss":51,"res_in":0,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"방배점","team":"서울PC/LT팀","views":442,"conn":73,"miss":25,"res_in":43,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"서초점","team":"서울PC/LT팀","views":625,"conn":105,"miss":23,"res_in":46,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"구리점","team":"서울PC/LT팀","views":383,"conn":66,"miss":30,"res_in":22,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"광교신도시점","team":"경기중부PC/LT팀","views":601,"conn":74,"miss":39,"res_in":71,"res_req":17,"review":5,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"관양점","team":"경기중부PC/LT팀","views":261,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"홈플러스송도점","team":"경인PC/LT팀","views":868,"conn":88,"miss":223,"res_in":81,"res_req":13,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"평택지산점","team":"경기중부PC/LT팀","views":250,"conn":26,"miss":10,"res_in":18,"res_req":5,"review":7,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"오창점","team":"대전충청PC/LT팀","views":326,"conn":43,"miss":3,"res_in":23,"res_req":3,"review":11,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"서세종점","team":"대전충청PC/LT팀","views":510,"conn":80,"miss":19,"res_in":45,"res_req":4,"review":6,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"대전점","team":"대전충청PC/LT팀","views":335,"conn":51,"miss":14,"res_in":26,"res_req":3,"review":7,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"대전서구점","team":"대전충청PC/LT팀","views":490,"conn":61,"miss":29,"res_in":41,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"송천점","team":"광주전라PC/LT팀","views":358,"conn":72,"miss":14,"res_in":31,"res_req":5,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"전주역점","team":"광주전라PC/LT팀","views":203,"conn":40,"miss":3,"res_in":12,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"남대구IC점","team":"대구경북PC/LT팀","views":257,"conn":46,"miss":10,"res_in":16,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"수성점","team":"대구경북PC/LT팀","views":355,"conn":54,"miss":10,"res_in":23,"res_req":4,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"부산거제점","team":"부산울산PC/LT팀","views":311,"conn":33,"miss":11,"res_in":23,"res_req":6,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":3,"start":"2025-09-15","end":"2025-09-21","store":"부산수영점","team":"부산울산PC/LT팀","views":241,"conn":31,"miss":12,"res_in":12,"res_req":2,"review":5,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"한남점","team":"서울PC/LT팀","views":714,"conn":77,"miss":59,"res_in":0,"res_req":0,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"방배점","team":"서울PC/LT팀","views":513,"conn":77,"miss":26,"res_in":36,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"서초점","team":"서울PC/LT팀","views":585,"conn":96,"miss":22,"res_in":35,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"구리점","team":"서울PC/LT팀","views":360,"conn":50,"miss":29,"res_in":26,"res_req":2,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"광교신도시점","team":"경기중부PC/LT팀","views":508,"conn":65,"miss":38,"res_in":59,"res_req":11,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"관양점","team":"경기중부PC/LT팀","views":215,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"홈플러스송도점","team":"경인PC/LT팀","views":689,"conn":92,"miss":195,"res_in":70,"res_req":9,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"평택지산점","team":"경기중부PC/LT팀","views":297,"conn":32,"miss":3,"res_in":22,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"오창점","team":"대전충청PC/LT팀","views":338,"conn":45,"miss":7,"res_in":37,"res_req":6,"review":2,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"서세종점","team":"대전충청PC/LT팀","views":536,"conn":89,"miss":21,"res_in":47,"res_req":7,"review":7,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"대전점","team":"대전충청PC/LT팀","views":318,"conn":53,"miss":30,"res_in":15,"res_req":2,"review":5,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"대전서구점","team":"대전충청PC/LT팀","views":434,"conn":91,"miss":37,"res_in":33,"res_req":10,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"송천점","team":"광주전라PC/LT팀","views":379,"conn":84,"miss":6,"res_in":35,"res_req":8,"review":4,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"전주역점","team":"광주전라PC/LT팀","views":222,"conn":30,"miss":6,"res_in":12,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"남대구IC점","team":"대구경북PC/LT팀","views":242,"conn":42,"miss":4,"res_in":14,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"수성점","team":"대구경북PC/LT팀","views":269,"conn":33,"miss":9,"res_in":15,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"부산거제점","team":"부산울산PC/LT팀","views":248,"conn":29,"miss":1,"res_in":20,"res_req":4,"review":5,"chat":0},{"yr":2025,"mo":9,"wk":4,"start":"2025-09-22","end":"2025-09-28","store":"부산수영점","team":"부산울산PC/LT팀","views":274,"conn":25,"miss":4,"res_in":11,"res_req":2,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"한남점","team":"서울PC/LT팀","views":933,"conn":102,"miss":83,"res_in":13,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"방배점","team":"서울PC/LT팀","views":597,"conn":114,"miss":31,"res_in":57,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"서초점","team":"서울PC/LT팀","views":781,"conn":100,"miss":29,"res_in":51,"res_req":10,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"구리점","team":"서울PC/LT팀","views":518,"conn":99,"miss":33,"res_in":31,"res_req":6,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"광교신도시점","team":"경기중부PC/LT팀","views":759,"conn":148,"miss":81,"res_in":95,"res_req":14,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"관양점","team":"경기중부PC/LT팀","views":259,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"홈플러스송도점","team":"경인PC/LT팀","views":1036,"conn":138,"miss":180,"res_in":86,"res_req":18,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"평택지산점","team":"경기중부PC/LT팀","views":417,"conn":42,"miss":17,"res_in":26,"res_req":5,"review":19,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"오창점","team":"대전충청PC/LT팀","views":528,"conn":83,"miss":7,"res_in":54,"res_req":9,"review":16,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"서세종점","team":"대전충청PC/LT팀","views":799,"conn":117,"miss":38,"res_in":78,"res_req":21,"review":4,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"대전점","team":"대전충청PC/LT팀","views":452,"conn":98,"miss":34,"res_in":24,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"대전서구점","team":"대전충청PC/LT팀","views":689,"conn":68,"miss":45,"res_in":60,"res_req":14,"review":5,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"송천점","team":"광주전라PC/LT팀","views":707,"conn":114,"miss":27,"res_in":43,"res_req":4,"review":5,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"전주역점","team":"광주전라PC/LT팀","views":317,"conn":51,"miss":3,"res_in":22,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"남대구IC점","team":"대구경북PC/LT팀","views":359,"conn":67,"miss":10,"res_in":20,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"수성점","team":"대구경북PC/LT팀","views":516,"conn":76,"miss":17,"res_in":19,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"부산거제점","team":"부산울산PC/LT팀","views":445,"conn":59,"miss":18,"res_in":32,"res_req":7,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":1,"start":"2025-09-29","end":"2025-10-05","store":"부산수영점","team":"부산울산PC/LT팀","views":439,"conn":37,"miss":23,"res_in":30,"res_req":9,"review":6,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"한남점","team":"서울PC/LT팀","views":777,"conn":46,"miss":93,"res_in":6,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"방배점","team":"서울PC/LT팀","views":469,"conn":49,"miss":56,"res_in":42,"res_req":9,"review":4,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"서초점","team":"서울PC/LT팀","views":559,"conn":67,"miss":20,"res_in":30,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"구리점","team":"서울PC/LT팀","views":404,"conn":28,"miss":43,"res_in":30,"res_req":2,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"광교신도시점","team":"경기중부PC/LT팀","views":601,"conn":45,"miss":44,"res_in":72,"res_req":16,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"관양점","team":"경기중부PC/LT팀","views":279,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"홈플러스송도점","team":"경인PC/LT팀","views":905,"conn":67,"miss":174,"res_in":86,"res_req":19,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"평택지산점","team":"경기중부PC/LT팀","views":298,"conn":19,"miss":18,"res_in":17,"res_req":1,"review":6,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"오창점","team":"대전충청PC/LT팀","views":383,"conn":39,"miss":4,"res_in":29,"res_req":5,"review":7,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"서세종점","team":"대전충청PC/LT팀","views":604,"conn":49,"miss":16,"res_in":71,"res_req":9,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"대전점","team":"대전충청PC/LT팀","views":338,"conn":27,"miss":23,"res_in":30,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"대전서구점","team":"대전충청PC/LT팀","views":450,"conn":50,"miss":39,"res_in":28,"res_req":8,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"송천점","team":"광주전라PC/LT팀","views":457,"conn":76,"miss":18,"res_in":29,"res_req":5,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"전주역점","team":"광주전라PC/LT팀","views":268,"conn":32,"miss":11,"res_in":17,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"남대구IC점","team":"대구경북PC/LT팀","views":297,"conn":17,"miss":16,"res_in":22,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"수성점","team":"대구경북PC/LT팀","views":332,"conn":19,"miss":10,"res_in":21,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"부산거제점","team":"부산울산PC/LT팀","views":347,"conn":32,"miss":16,"res_in":20,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":2,"start":"2025-10-06","end":"2025-10-12","store":"부산수영점","team":"부산울산PC/LT팀","views":282,"conn":25,"miss":8,"res_in":7,"res_req":1,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"한남점","team":"서울PC/LT팀","views":713,"conn":86,"miss":75,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"방배점","team":"서울PC/LT팀","views":475,"conn":67,"miss":37,"res_in":40,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"서초점","team":"서울PC/LT팀","views":524,"conn":97,"miss":18,"res_in":38,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"구리점","team":"서울PC/LT팀","views":396,"conn":64,"miss":26,"res_in":33,"res_req":5,"review":4,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"광교신도시점","team":"경기중부PC/LT팀","views":618,"conn":68,"miss":53,"res_in":62,"res_req":17,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"관양점","team":"경기중부PC/LT팀","views":250,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"홈플러스송도점","team":"경인PC/LT팀","views":913,"conn":151,"miss":179,"res_in":87,"res_req":14,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"평택지산점","team":"경기중부PC/LT팀","views":246,"conn":26,"miss":5,"res_in":17,"res_req":1,"review":6,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"오창점","team":"대전충청PC/LT팀","views":337,"conn":56,"miss":11,"res_in":35,"res_req":6,"review":11,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"서세종점","team":"대전충청PC/LT팀","views":514,"conn":92,"miss":27,"res_in":39,"res_req":6,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"대전점","team":"대전충청PC/LT팀","views":348,"conn":38,"miss":29,"res_in":37,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"대전서구점","team":"대전충청PC/LT팀","views":448,"conn":63,"miss":24,"res_in":43,"res_req":10,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"송천점","team":"광주전라PC/LT팀","views":446,"conn":85,"miss":13,"res_in":43,"res_req":5,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"전주역점","team":"광주전라PC/LT팀","views":196,"conn":36,"miss":9,"res_in":8,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"남대구IC점","team":"대구경북PC/LT팀","views":234,"conn":29,"miss":1,"res_in":11,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"수성점","team":"대구경북PC/LT팀","views":346,"conn":53,"miss":13,"res_in":16,"res_req":5,"review":7,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"부산거제점","team":"부산울산PC/LT팀","views":308,"conn":48,"miss":7,"res_in":24,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":3,"start":"2025-10-13","end":"2025-10-19","store":"부산수영점","team":"부산울산PC/LT팀","views":198,"conn":23,"miss":14,"res_in":10,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"한남점","team":"서울PC/LT팀","views":820,"conn":100,"miss":83,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"방배점","team":"서울PC/LT팀","views":555,"conn":98,"miss":59,"res_in":57,"res_req":16,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"서초점","team":"서울PC/LT팀","views":684,"conn":121,"miss":51,"res_in":59,"res_req":6,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"구리점","team":"서울PC/LT팀","views":518,"conn":78,"miss":33,"res_in":34,"res_req":2,"review":10,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"광교신도시점","team":"경기중부PC/LT팀","views":817,"conn":78,"miss":189,"res_in":102,"res_req":10,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"관양점","team":"경기중부PC/LT팀","views":290,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"홈플러스송도점","team":"경인PC/LT팀","views":936,"conn":125,"miss":229,"res_in":87,"res_req":12,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"평택지산점","team":"경기중부PC/LT팀","views":345,"conn":29,"miss":28,"res_in":25,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"오창점","team":"대전충청PC/LT팀","views":483,"conn":68,"miss":15,"res_in":51,"res_req":12,"review":10,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"서세종점","team":"대전충청PC/LT팀","views":617,"conn":93,"miss":20,"res_in":56,"res_req":8,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"대전점","team":"대전충청PC/LT팀","views":483,"conn":53,"miss":57,"res_in":47,"res_req":5,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"대전서구점","team":"대전충청PC/LT팀","views":542,"conn":67,"miss":33,"res_in":35,"res_req":8,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"송천점","team":"광주전라PC/LT팀","views":487,"conn":115,"miss":30,"res_in":24,"res_req":2,"review":3,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"전주역점","team":"광주전라PC/LT팀","views":244,"conn":40,"miss":4,"res_in":25,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"남대구IC점","team":"대구경북PC/LT팀","views":246,"conn":34,"miss":19,"res_in":13,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"수성점","team":"대구경북PC/LT팀","views":371,"conn":44,"miss":15,"res_in":20,"res_req":0,"review":7,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"부산거제점","team":"부산울산PC/LT팀","views":371,"conn":52,"miss":22,"res_in":30,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":4,"start":"2025-10-20","end":"2025-10-26","store":"부산수영점","team":"부산울산PC/LT팀","views":229,"conn":15,"miss":9,"res_in":14,"res_req":0,"review":4,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"한남점","team":"서울PC/LT팀","views":729,"conn":75,"miss":87,"res_in":0,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"방배점","team":"서울PC/LT팀","views":593,"conn":98,"miss":45,"res_in":41,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"서초점","team":"서울PC/LT팀","views":790,"conn":143,"miss":54,"res_in":51,"res_req":5,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"구리점","team":"서울PC/LT팀","views":613,"conn":83,"miss":33,"res_in":35,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"광교신도시점","team":"경기중부PC/LT팀","views":1046,"conn":112,"miss":101,"res_in":107,"res_req":12,"review":4,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"관양점","team":"경기중부PC/LT팀","views":288,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"홈플러스송도점","team":"경인PC/LT팀","views":1085,"conn":113,"miss":419,"res_in":90,"res_req":13,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"평택지산점","team":"경기중부PC/LT팀","views":416,"conn":43,"miss":5,"res_in":20,"res_req":2,"review":5,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"오창점","team":"대전충청PC/LT팀","views":614,"conn":83,"miss":20,"res_in":45,"res_req":6,"review":12,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"서세종점","team":"대전충청PC/LT팀","views":849,"conn":137,"miss":43,"res_in":58,"res_req":13,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"대전점","team":"대전충청PC/LT팀","views":520,"conn":69,"miss":43,"res_in":30,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"대전서구점","team":"대전충청PC/LT팀","views":616,"conn":75,"miss":36,"res_in":39,"res_req":9,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"송천점","team":"광주전라PC/LT팀","views":679,"conn":155,"miss":34,"res_in":44,"res_req":4,"review":1,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"전주역점","team":"광주전라PC/LT팀","views":325,"conn":58,"miss":7,"res_in":17,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"남대구IC점","team":"대구경북PC/LT팀","views":260,"conn":49,"miss":8,"res_in":17,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"수성점","team":"대구경북PC/LT팀","views":627,"conn":85,"miss":20,"res_in":28,"res_req":2,"review":6,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"부산거제점","team":"부산울산PC/LT팀","views":494,"conn":47,"miss":22,"res_in":31,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":10,"wk":5,"start":"2025-10-27","end":"2025-11-02","store":"부산수영점","team":"부산울산PC/LT팀","views":365,"conn":33,"miss":31,"res_in":17,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"한남점","team":"서울PC/LT팀","views":873,"conn":108,"miss":148,"res_in":0,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"방배점","team":"서울PC/LT팀","views":616,"conn":144,"miss":60,"res_in":37,"res_req":5,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"서초점","team":"서울PC/LT팀","views":859,"conn":140,"miss":63,"res_in":54,"res_req":4,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"구로점","team":"경인PC/LT팀","views":109,"conn":17,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"구리점","team":"서울PC/LT팀","views":518,"conn":92,"miss":36,"res_in":42,"res_req":5,"review":5,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"광교신도시점","team":"경기중부PC/LT팀","views":1015,"conn":124,"miss":97,"res_in":94,"res_req":22,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"관양점","team":"경기중부PC/LT팀","views":251,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"홈플러스송도점","team":"경인PC/LT팀","views":1125,"conn":63,"miss":365,"res_in":101,"res_req":19,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"평택지산점","team":"경기중부PC/LT팀","views":385,"conn":39,"miss":30,"res_in":14,"res_req":3,"review":8,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"오창점","team":"대전충청PC/LT팀","views":544,"conn":103,"miss":19,"res_in":31,"res_req":2,"review":10,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"서세종점","team":"대전충청PC/LT팀","views":777,"conn":141,"miss":40,"res_in":52,"res_req":11,"review":4,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"대전점","team":"대전충청PC/LT팀","views":546,"conn":54,"miss":76,"res_in":26,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"대전서구점","team":"대전충청PC/LT팀","views":591,"conn":83,"miss":27,"res_in":42,"res_req":8,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"송천점","team":"광주전라PC/LT팀","views":592,"conn":121,"miss":16,"res_in":56,"res_req":6,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"전주역점","team":"광주전라PC/LT팀","views":304,"conn":53,"miss":4,"res_in":16,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"남대구IC점","team":"대구경북PC/LT팀","views":302,"conn":53,"miss":5,"res_in":16,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"수성점","team":"대구경북PC/LT팀","views":369,"conn":47,"miss":15,"res_in":20,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"부산거제점","team":"부산울산PC/LT팀","views":374,"conn":43,"miss":22,"res_in":17,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":1,"start":"2025-11-03","end":"2025-11-09","store":"부산수영점","team":"부산울산PC/LT팀","views":334,"conn":39,"miss":14,"res_in":16,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"한남점","team":"서울PC/LT팀","views":914,"conn":93,"miss":196,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"방배점","team":"서울PC/LT팀","views":608,"conn":142,"miss":56,"res_in":57,"res_req":9,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"서초점","team":"서울PC/LT팀","views":866,"conn":143,"miss":58,"res_in":35,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"구로점","team":"경인PC/LT팀","views":560,"conn":43,"miss":1,"res_in":22,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"구리점","team":"서울PC/LT팀","views":522,"conn":105,"miss":32,"res_in":35,"res_req":3,"review":6,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"광교신도시점","team":"경기중부PC/LT팀","views":876,"conn":144,"miss":104,"res_in":92,"res_req":20,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"관양점","team":"경기중부PC/LT팀","views":291,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"홈플러스송도점","team":"경인PC/LT팀","views":989,"conn":68,"miss":303,"res_in":109,"res_req":23,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"평택지산점","team":"경기중부PC/LT팀","views":426,"conn":33,"miss":17,"res_in":25,"res_req":1,"review":6,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"오창점","team":"대전충청PC/LT팀","views":473,"conn":54,"miss":3,"res_in":31,"res_req":6,"review":8,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"서세종점","team":"대전충청PC/LT팀","views":837,"conn":127,"miss":42,"res_in":83,"res_req":16,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"대전점","team":"대전충청PC/LT팀","views":475,"conn":94,"miss":40,"res_in":38,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"대전서구점","team":"대전충청PC/LT팀","views":532,"conn":82,"miss":44,"res_in":31,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"송천점","team":"광주전라PC/LT팀","views":551,"conn":102,"miss":24,"res_in":46,"res_req":6,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"전주역점","team":"광주전라PC/LT팀","views":324,"conn":70,"miss":7,"res_in":24,"res_req":1,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"남대구IC점","team":"대구경북PC/LT팀","views":228,"conn":32,"miss":9,"res_in":11,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"수성점","team":"대구경북PC/LT팀","views":437,"conn":63,"miss":20,"res_in":17,"res_req":3,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"부산거제점","team":"부산울산PC/LT팀","views":363,"conn":46,"miss":14,"res_in":20,"res_req":1,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":2,"start":"2025-11-10","end":"2025-11-16","store":"부산수영점","team":"부산울산PC/LT팀","views":324,"conn":25,"miss":17,"res_in":16,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"한남점","team":"서울PC/LT팀","views":888,"conn":99,"miss":147,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"방배점","team":"서울PC/LT팀","views":630,"conn":138,"miss":70,"res_in":52,"res_req":6,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"서초점","team":"서울PC/LT팀","views":748,"conn":138,"miss":50,"res_in":0,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"구로점","team":"경인PC/LT팀","views":642,"conn":87,"miss":0,"res_in":37,"res_req":3,"review":4,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"구리점","team":"서울PC/LT팀","views":512,"conn":94,"miss":72,"res_in":50,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"광교신도시점","team":"경기중부PC/LT팀","views":891,"conn":118,"miss":79,"res_in":122,"res_req":13,"review":5,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"관양점","team":"경기중부PC/LT팀","views":331,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"홈플러스송도점","team":"경인PC/LT팀","views":1128,"conn":89,"miss":427,"res_in":125,"res_req":21,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"평택지산점","team":"경기중부PC/LT팀","views":389,"conn":41,"miss":6,"res_in":19,"res_req":1,"review":4,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"오창점","team":"대전충청PC/LT팀","views":465,"conn":50,"miss":12,"res_in":44,"res_req":7,"review":8,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"서세종점","team":"대전충청PC/LT팀","views":622,"conn":112,"miss":26,"res_in":68,"res_req":14,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"대전점","team":"대전충청PC/LT팀","views":462,"conn":64,"miss":61,"res_in":38,"res_req":3,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"대전서구점","team":"대전충청PC/LT팀","views":571,"conn":72,"miss":57,"res_in":30,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"송천점","team":"광주전라PC/LT팀","views":534,"conn":118,"miss":19,"res_in":29,"res_req":4,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"전주역점","team":"광주전라PC/LT팀","views":300,"conn":54,"miss":10,"res_in":27,"res_req":1,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"남대구IC점","team":"대구경북PC/LT팀","views":269,"conn":50,"miss":15,"res_in":13,"res_req":1,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"수성점","team":"대구경북PC/LT팀","views":383,"conn":46,"miss":17,"res_in":25,"res_req":5,"review":6,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"부산거제점","team":"부산울산PC/LT팀","views":412,"conn":38,"miss":16,"res_in":24,"res_req":3,"review":10,"chat":0},{"yr":2025,"mo":11,"wk":3,"start":"2025-11-17","end":"2025-11-23","store":"부산수영점","team":"부산울산PC/LT팀","views":315,"conn":23,"miss":12,"res_in":12,"res_req":2,"review":12,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"한남점","team":"서울PC/LT팀","views":731,"conn":96,"miss":93,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"방배점","team":"서울PC/LT팀","views":687,"conn":114,"miss":62,"res_in":71,"res_req":5,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"서초점","team":"서울PC/LT팀","views":642,"conn":90,"miss":52,"res_in":3,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"구로점","team":"경인PC/LT팀","views":656,"conn":64,"miss":0,"res_in":50,"res_req":7,"review":7,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"구리점","team":"서울PC/LT팀","views":482,"conn":102,"miss":38,"res_in":50,"res_req":4,"review":4,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"광교신도시점","team":"경기중부PC/LT팀","views":668,"conn":99,"miss":67,"res_in":78,"res_req":8,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"관양점","team":"경기중부PC/LT팀","views":276,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"홈플러스송도점","team":"경인PC/LT팀","views":737,"conn":85,"miss":202,"res_in":78,"res_req":13,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"평택지산점","team":"경기중부PC/LT팀","views":342,"conn":29,"miss":11,"res_in":28,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"오창점","team":"대전충청PC/LT팀","views":387,"conn":83,"miss":10,"res_in":29,"res_req":8,"review":2,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"서세종점","team":"대전충청PC/LT팀","views":535,"conn":108,"miss":19,"res_in":61,"res_req":7,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"대전점","team":"대전충청PC/LT팀","views":402,"conn":59,"miss":57,"res_in":41,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"대전서구점","team":"대전충청PC/LT팀","views":486,"conn":65,"miss":63,"res_in":38,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"송천점","team":"광주전라PC/LT팀","views":361,"conn":90,"miss":7,"res_in":31,"res_req":7,"review":1,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"전주역점","team":"광주전라PC/LT팀","views":244,"conn":52,"miss":3,"res_in":16,"res_req":1,"review":4,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"남대구IC점","team":"대구경북PC/LT팀","views":215,"conn":27,"miss":3,"res_in":8,"res_req":2,"review":0,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"수성점","team":"대구경북PC/LT팀","views":330,"conn":42,"miss":13,"res_in":16,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"부산거제점","team":"부산울산PC/LT팀","views":301,"conn":28,"miss":18,"res_in":19,"res_req":2,"review":4,"chat":0},{"yr":2025,"mo":11,"wk":4,"start":"2025-11-24","end":"2025-11-30","store":"부산수영점","team":"부산울산PC/LT팀","views":291,"conn":35,"miss":5,"res_in":9,"res_req":2,"review":11,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"한남점","team":"서울PC/LT팀","views":1462,"conn":192,"miss":356,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"방배점","team":"서울PC/LT팀","views":1125,"conn":197,"miss":163,"res_in":117,"res_req":12,"review":4,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"서초점","team":"서울PC/LT팀","views":1147,"conn":206,"miss":126,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"구로점","team":"경인PC/LT팀","views":1039,"conn":175,"miss":1,"res_in":85,"res_req":11,"review":5,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"구리점","team":"서울PC/LT팀","views":799,"conn":169,"miss":117,"res_in":81,"res_req":12,"review":5,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"광교신도시점","team":"경기중부PC/LT팀","views":1630,"conn":204,"miss":352,"res_in":69,"res_req":7,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"관양점","team":"경기중부PC/LT팀","views":421,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"홈플러스송도점","team":"경인PC/LT팀","views":1638,"conn":334,"miss":257,"res_in":176,"res_req":15,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"평택지산점","team":"경기중부PC/LT팀","views":522,"conn":56,"miss":20,"res_in":35,"res_req":5,"review":5,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"오창점","team":"대전충청PC/LT팀","views":938,"conn":152,"miss":28,"res_in":89,"res_req":12,"review":14,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"서세종점","team":"대전충청PC/LT팀","views":819,"conn":156,"miss":41,"res_in":94,"res_req":14,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"대전점","team":"대전충청PC/LT팀","views":710,"conn":109,"miss":118,"res_in":52,"res_req":8,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"대전서구점","team":"대전충청PC/LT팀","views":729,"conn":109,"miss":95,"res_in":49,"res_req":8,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"송천점","team":"광주전라PC/LT팀","views":738,"conn":166,"miss":27,"res_in":70,"res_req":9,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"전주역점","team":"광주전라PC/LT팀","views":438,"conn":76,"miss":12,"res_in":37,"res_req":2,"review":4,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"남대구IC점","team":"대구경북PC/LT팀","views":380,"conn":52,"miss":24,"res_in":19,"res_req":3,"review":8,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"수성점","team":"대구경북PC/LT팀","views":546,"conn":81,"miss":26,"res_in":35,"res_req":3,"review":9,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"부산거제점","team":"부산울산PC/LT팀","views":431,"conn":50,"miss":29,"res_in":17,"res_req":2,"review":11,"chat":0},{"yr":2025,"mo":12,"wk":1,"start":"2025-12-01","end":"2025-12-07","store":"부산수영점","team":"부산울산PC/LT팀","views":470,"conn":48,"miss":23,"res_in":17,"res_req":1,"review":15,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"한남점","team":"서울PC/LT팀","views":1172,"conn":164,"miss":266,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"방배점","team":"서울PC/LT팀","views":823,"conn":176,"miss":152,"res_in":89,"res_req":11,"review":6,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"서초점","team":"서울PC/LT팀","views":930,"conn":217,"miss":82,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"구로점","team":"경인PC/LT팀","views":831,"conn":133,"miss":0,"res_in":57,"res_req":8,"review":18,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"구리점","team":"서울PC/LT팀","views":702,"conn":123,"miss":61,"res_in":47,"res_req":5,"review":6,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"광교신도시점","team":"경기중부PC/LT팀","views":1282,"conn":183,"miss":235,"res_in":0,"res_req":0,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"관양점","team":"경기중부PC/LT팀","views":362,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"홈플러스송도점","team":"경인PC/LT팀","views":910,"conn":221,"miss":126,"res_in":101,"res_req":10,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"평택지산점","team":"경기중부PC/LT팀","views":438,"conn":54,"miss":17,"res_in":28,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"오창점","team":"대전충청PC/LT팀","views":587,"conn":89,"miss":20,"res_in":46,"res_req":3,"review":12,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"서세종점","team":"대전충청PC/LT팀","views":735,"conn":149,"miss":45,"res_in":54,"res_req":9,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"대전점","team":"대전충청PC/LT팀","views":504,"conn":90,"miss":63,"res_in":49,"res_req":10,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"대전서구점","team":"대전충청PC/LT팀","views":636,"conn":108,"miss":103,"res_in":37,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"송천점","team":"광주전라PC/LT팀","views":553,"conn":121,"miss":20,"res_in":41,"res_req":7,"review":9,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"전주역점","team":"광주전라PC/LT팀","views":297,"conn":62,"miss":6,"res_in":23,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"남대구IC점","team":"대구경북PC/LT팀","views":308,"conn":38,"miss":5,"res_in":13,"res_req":2,"review":20,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"수성점","team":"대구경북PC/LT팀","views":448,"conn":76,"miss":21,"res_in":16,"res_req":3,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"부산거제점","team":"부산울산PC/LT팀","views":361,"conn":40,"miss":15,"res_in":19,"res_req":2,"review":7,"chat":0},{"yr":2025,"mo":12,"wk":2,"start":"2025-12-08","end":"2025-12-14","store":"부산수영점","team":"부산울산PC/LT팀","views":292,"conn":22,"miss":15,"res_in":14,"res_req":2,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"한남점","team":"서울PC/LT팀","views":929,"conn":125,"miss":158,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"방배점","team":"서울PC/LT팀","views":667,"conn":130,"miss":63,"res_in":56,"res_req":5,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"서초점","team":"서울PC/LT팀","views":810,"conn":158,"miss":74,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"구로점","team":"경인PC/LT팀","views":589,"conn":92,"miss":0,"res_in":19,"res_req":4,"review":10,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"구리점","team":"서울PC/LT팀","views":471,"conn":114,"miss":29,"res_in":32,"res_req":5,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"광교신도시점","team":"경기중부PC/LT팀","views":746,"conn":139,"miss":85,"res_in":1,"res_req":0,"review":4,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"관양점","team":"경기중부PC/LT팀","views":237,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"홈플러스송도점","team":"경인PC/LT팀","views":824,"conn":161,"miss":152,"res_in":64,"res_req":9,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"평택지산점","team":"경기중부PC/LT팀","views":362,"conn":44,"miss":20,"res_in":14,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"오창점","team":"대전충청PC/LT팀","views":488,"conn":89,"miss":7,"res_in":30,"res_req":3,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"서세종점","team":"대전충청PC/LT팀","views":502,"conn":85,"miss":29,"res_in":41,"res_req":8,"review":5,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"대전점","team":"대전충청PC/LT팀","views":419,"conn":70,"miss":59,"res_in":31,"res_req":6,"review":4,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"대전서구점","team":"대전충청PC/LT팀","views":479,"conn":100,"miss":58,"res_in":17,"res_req":2,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"송천점","team":"광주전라PC/LT팀","views":480,"conn":76,"miss":7,"res_in":18,"res_req":2,"review":9,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"전주역점","team":"광주전라PC/LT팀","views":253,"conn":38,"miss":6,"res_in":12,"res_req":1,"review":7,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"남대구IC점","team":"대구경북PC/LT팀","views":283,"conn":28,"miss":10,"res_in":16,"res_req":3,"review":20,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"수성점","team":"대구경북PC/LT팀","views":411,"conn":52,"miss":20,"res_in":18,"res_req":2,"review":7,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"부산거제점","team":"부산울산PC/LT팀","views":314,"conn":35,"miss":15,"res_in":9,"res_req":1,"review":6,"chat":0},{"yr":2025,"mo":12,"wk":3,"start":"2025-12-15","end":"2025-12-21","store":"부산수영점","team":"부산울산PC/LT팀","views":283,"conn":29,"miss":15,"res_in":12,"res_req":1,"review":6,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"한남점","team":"서울PC/LT팀","views":734,"conn":84,"miss":104,"res_in":0,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"방배점","team":"서울PC/LT팀","views":574,"conn":98,"miss":19,"res_in":46,"res_req":4,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"서초점","team":"서울PC/LT팀","views":648,"conn":96,"miss":24,"res_in":1,"res_req":0,"review":2,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"구로점","team":"경인PC/LT팀","views":436,"conn":73,"miss":0,"res_in":25,"res_req":1,"review":7,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"구리점","team":"서울PC/LT팀","views":345,"conn":49,"miss":33,"res_in":29,"res_req":2,"review":8,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"광교신도시점","team":"경기중부PC/LT팀","views":593,"conn":73,"miss":59,"res_in":0,"res_req":0,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"관양점","team":"경기중부PC/LT팀","views":241,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"홈플러스송도점","team":"경인PC/LT팀","views":802,"conn":122,"miss":132,"res_in":78,"res_req":9,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"평택지산점","team":"경기중부PC/LT팀","views":270,"conn":31,"miss":12,"res_in":12,"res_req":4,"review":7,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"오창점","team":"대전충청PC/LT팀","views":329,"conn":48,"miss":8,"res_in":39,"res_req":6,"review":0,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"서세종점","team":"대전충청PC/LT팀","views":540,"conn":84,"miss":25,"res_in":62,"res_req":12,"review":10,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"대전점","team":"대전충청PC/LT팀","views":381,"conn":58,"miss":39,"res_in":25,"res_req":4,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"대전서구점","team":"대전충청PC/LT팀","views":464,"conn":52,"miss":30,"res_in":30,"res_req":9,"review":1,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"송천점","team":"광주전라PC/LT팀","views":453,"conn":85,"miss":12,"res_in":32,"res_req":6,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"전주역점","team":"광주전라PC/LT팀","views":267,"conn":49,"miss":5,"res_in":17,"res_req":5,"review":7,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"남대구IC점","team":"대구경북PC/LT팀","views":225,"conn":23,"miss":6,"res_in":10,"res_req":4,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"수성점","team":"대구경북PC/LT팀","views":323,"conn":42,"miss":10,"res_in":18,"res_req":2,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"부산거제점","team":"부산울산PC/LT팀","views":345,"conn":46,"miss":13,"res_in":25,"res_req":5,"review":3,"chat":0},{"yr":2025,"mo":12,"wk":4,"start":"2025-12-22","end":"2025-12-28","store":"부산수영점","team":"부산울산PC/LT팀","views":240,"conn":27,"miss":12,"res_in":12,"res_req":2,"review":5,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"한남점","team":"서울PC/LT팀","views":788,"conn":56,"miss":66,"res_in":12,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"방배점","team":"서울PC/LT팀","views":592,"conn":87,"miss":36,"res_in":37,"res_req":7,"review":4,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"서초점","team":"서울PC/LT팀","views":673,"conn":81,"miss":19,"res_in":36,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"구로점","team":"경인PC/LT팀","views":515,"conn":89,"miss":0,"res_in":27,"res_req":4,"review":7,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"구리점","team":"서울PC/LT팀","views":394,"conn":72,"miss":23,"res_in":17,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"광교신도시점","team":"경기중부PC/LT팀","views":742,"conn":83,"miss":49,"res_in":38,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"관양점","team":"경기중부PC/LT팀","views":221,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"홈플러스송도점","team":"경인PC/LT팀","views":1234,"conn":146,"miss":257,"res_in":55,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"평택지산점","team":"경기중부PC/LT팀","views":308,"conn":43,"miss":11,"res_in":17,"res_req":2,"review":7,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"오창점","team":"대전충청PC/LT팀","views":380,"conn":53,"miss":9,"res_in":28,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"서세종점","team":"대전충청PC/LT팀","views":585,"conn":99,"miss":11,"res_in":55,"res_req":13,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"대전점","team":"대전충청PC/LT팀","views":454,"conn":54,"miss":32,"res_in":33,"res_req":4,"review":4,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"대전서구점","team":"대전충청PC/LT팀","views":556,"conn":66,"miss":33,"res_in":33,"res_req":6,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"송천점","team":"광주전라PC/LT팀","views":527,"conn":110,"miss":22,"res_in":40,"res_req":8,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"전주역점","team":"광주전라PC/LT팀","views":272,"conn":48,"miss":4,"res_in":8,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"남대구IC점","team":"대구경북PC/LT팀","views":333,"conn":43,"miss":5,"res_in":16,"res_req":1,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"수성점","team":"대구경북PC/LT팀","views":489,"conn":44,"miss":16,"res_in":21,"res_req":2,"review":5,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"부산거제점","team":"부산울산PC/LT팀","views":413,"conn":50,"miss":23,"res_in":23,"res_req":0,"review":6,"chat":0},{"yr":2026,"mo":1,"wk":1,"start":"2025-12-29","end":"2026-01-04","store":"부산수영점","team":"부산울산PC/LT팀","views":322,"conn":26,"miss":23,"res_in":16,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"한남점","team":"서울PC/LT팀","views":660,"conn":59,"miss":52,"res_in":34,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"방배점","team":"서울PC/LT팀","views":442,"conn":77,"miss":31,"res_in":35,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"서초점","team":"서울PC/LT팀","views":496,"conn":72,"miss":17,"res_in":35,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"구로점","team":"경인PC/LT팀","views":363,"conn":41,"miss":0,"res_in":19,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"구리점","team":"서울PC/LT팀","views":348,"conn":61,"miss":29,"res_in":23,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"광교신도시점","team":"경기중부PC/LT팀","views":555,"conn":81,"miss":44,"res_in":54,"res_req":9,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"관양점","team":"경기중부PC/LT팀","views":200,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"홈플러스송도점","team":"경인PC/LT팀","views":712,"conn":120,"miss":119,"res_in":62,"res_req":8,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"평택지산점","team":"경기중부PC/LT팀","views":208,"conn":11,"miss":17,"res_in":5,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"오창점","team":"대전충청PC/LT팀","views":305,"conn":40,"miss":8,"res_in":25,"res_req":3,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"서세종점","team":"대전충청PC/LT팀","views":488,"conn":78,"miss":15,"res_in":38,"res_req":6,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"대전점","team":"대전충청PC/LT팀","views":378,"conn":76,"miss":18,"res_in":31,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"대전서구점","team":"대전충청PC/LT팀","views":413,"conn":64,"miss":33,"res_in":23,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"송천점","team":"광주전라PC/LT팀","views":453,"conn":81,"miss":15,"res_in":39,"res_req":6,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"전주역점","team":"광주전라PC/LT팀","views":230,"conn":45,"miss":3,"res_in":10,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"남대구IC점","team":"대구경북PC/LT팀","views":211,"conn":29,"miss":5,"res_in":12,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"수성점","team":"대구경북PC/LT팀","views":346,"conn":56,"miss":9,"res_in":19,"res_req":4,"review":9,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"부산거제점","team":"부산울산PC/LT팀","views":260,"conn":33,"miss":14,"res_in":20,"res_req":5,"review":13,"chat":0},{"yr":2026,"mo":1,"wk":2,"start":"2026-01-05","end":"2026-01-11","store":"부산수영점","team":"부산울산PC/LT팀","views":204,"conn":21,"miss":2,"res_in":16,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"한남점","team":"서울PC/LT팀","views":689,"conn":65,"miss":123,"res_in":45,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"방배점","team":"서울PC/LT팀","views":418,"conn":60,"miss":22,"res_in":46,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"서초점","team":"서울PC/LT팀","views":527,"conn":85,"miss":22,"res_in":38,"res_req":8,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"구로점","team":"경인PC/LT팀","views":351,"conn":34,"miss":1,"res_in":12,"res_req":3,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"구리점","team":"서울PC/LT팀","views":296,"conn":61,"miss":21,"res_in":30,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"광교신도시점","team":"경기중부PC/LT팀","views":497,"conn":58,"miss":20,"res_in":59,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"관양점","team":"경기중부PC/LT팀","views":257,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"홈플러스송도점","team":"경인PC/LT팀","views":635,"conn":111,"miss":134,"res_in":65,"res_req":9,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"평택지산점","team":"경기중부PC/LT팀","views":221,"conn":15,"miss":8,"res_in":10,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"오창점","team":"대전충청PC/LT팀","views":225,"conn":34,"miss":10,"res_in":8,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"서세종점","team":"대전충청PC/LT팀","views":311,"conn":49,"miss":2,"res_in":43,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"대전점","team":"대전충청PC/LT팀","views":304,"conn":52,"miss":21,"res_in":25,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"대전서구점","team":"대전충청PC/LT팀","views":318,"conn":51,"miss":22,"res_in":19,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"송천점","team":"광주전라PC/LT팀","views":346,"conn":81,"miss":2,"res_in":23,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"전주역점","team":"광주전라PC/LT팀","views":188,"conn":45,"miss":1,"res_in":11,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"남대구IC점","team":"대구경북PC/LT팀","views":198,"conn":25,"miss":5,"res_in":15,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"수성점","team":"대구경북PC/LT팀","views":340,"conn":41,"miss":9,"res_in":19,"res_req":4,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"부산거제점","team":"부산울산PC/LT팀","views":256,"conn":26,"miss":10,"res_in":20,"res_req":3,"review":14,"chat":0},{"yr":2026,"mo":1,"wk":3,"start":"2026-01-12","end":"2026-01-18","store":"부산수영점","team":"부산울산PC/LT팀","views":194,"conn":24,"miss":17,"res_in":11,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"한남점","team":"서울PC/LT팀","views":704,"conn":76,"miss":61,"res_in":40,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"방배점","team":"서울PC/LT팀","views":592,"conn":102,"miss":41,"res_in":41,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"서초점","team":"서울PC/LT팀","views":493,"conn":81,"miss":13,"res_in":39,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"구로점","team":"경인PC/LT팀","views":442,"conn":54,"miss":0,"res_in":23,"res_req":2,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"구리점","team":"서울PC/LT팀","views":386,"conn":70,"miss":26,"res_in":23,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"광교신도시점","team":"경기중부PC/LT팀","views":567,"conn":72,"miss":31,"res_in":62,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"관양점","team":"경기중부PC/LT팀","views":239,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"홈플러스송도점","team":"경인PC/LT팀","views":862,"conn":114,"miss":229,"res_in":77,"res_req":6,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"평택지산점","team":"경기중부PC/LT팀","views":271,"conn":26,"miss":10,"res_in":11,"res_req":0,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"오창점","team":"대전충청PC/LT팀","views":334,"conn":39,"miss":2,"res_in":41,"res_req":6,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"서세종점","team":"대전충청PC/LT팀","views":321,"conn":54,"miss":10,"res_in":33,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"대전점","team":"대전충청PC/LT팀","views":301,"conn":33,"miss":4,"res_in":23,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"대전서구점","team":"대전충청PC/LT팀","views":401,"conn":56,"miss":29,"res_in":27,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"송천점","team":"광주전라PC/LT팀","views":394,"conn":79,"miss":4,"res_in":33,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"전주역점","team":"광주전라PC/LT팀","views":221,"conn":36,"miss":4,"res_in":14,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"남대구IC점","team":"대구경북PC/LT팀","views":275,"conn":45,"miss":16,"res_in":17,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"수성점","team":"대구경북PC/LT팀","views":341,"conn":43,"miss":13,"res_in":13,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"부산거제점","team":"부산울산PC/LT팀","views":326,"conn":36,"miss":17,"res_in":18,"res_req":3,"review":6,"chat":0},{"yr":2026,"mo":1,"wk":4,"start":"2026-01-19","end":"2026-01-25","store":"부산수영점","team":"부산울산PC/LT팀","views":235,"conn":21,"miss":6,"res_in":12,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"한남점","team":"서울PC/LT팀","views":833,"conn":85,"miss":45,"res_in":48,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"방배점","team":"서울PC/LT팀","views":579,"conn":64,"miss":31,"res_in":47,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"서초점","team":"서울PC/LT팀","views":558,"conn":64,"miss":12,"res_in":33,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"구로점","team":"경인PC/LT팀","views":483,"conn":65,"miss":0,"res_in":28,"res_req":5,"review":6,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"구리점","team":"서울PC/LT팀","views":433,"conn":55,"miss":30,"res_in":29,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"광교신도시점","team":"경기중부PC/LT팀","views":2974,"conn":365,"miss":232,"res_in":183,"res_req":38,"review":7,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"관양점","team":"경기중부PC/LT팀","views":291,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"홈플러스송도점","team":"경인PC/LT팀","views":858,"conn":136,"miss":172,"res_in":77,"res_req":15,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"평택지산점","team":"경기중부PC/LT팀","views":270,"conn":19,"miss":10,"res_in":15,"res_req":2,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"오창점","team":"대전충청PC/LT팀","views":314,"conn":35,"miss":3,"res_in":26,"res_req":8,"review":4,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"서세종점","team":"대전충청PC/LT팀","views":467,"conn":55,"miss":14,"res_in":37,"res_req":8,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"대전점","team":"대전충청PC/LT팀","views":307,"conn":32,"miss":17,"res_in":22,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"대전서구점","team":"대전충청PC/LT팀","views":431,"conn":47,"miss":20,"res_in":19,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"송천점","team":"광주전라PC/LT팀","views":440,"conn":84,"miss":4,"res_in":27,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"전주역점","team":"광주전라PC/LT팀","views":180,"conn":23,"miss":1,"res_in":8,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"남대구IC점","team":"대구경북PC/LT팀","views":231,"conn":28,"miss":6,"res_in":16,"res_req":3,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"수성점","team":"대구경북PC/LT팀","views":363,"conn":36,"miss":21,"res_in":19,"res_req":0,"review":3,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"부산거제점","team":"부산울산PC/LT팀","views":318,"conn":32,"miss":14,"res_in":21,"res_req":1,"review":5,"chat":0},{"yr":2026,"mo":1,"wk":5,"start":"2026-01-26","end":"2026-02-01","store":"부산수영점","team":"부산울산PC/LT팀","views":330,"conn":25,"miss":10,"res_in":20,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"한남점","team":"서울PC/LT팀","views":713,"conn":74,"miss":26,"res_in":31,"res_req":8,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"방배점","team":"서울PC/LT팀","views":426,"conn":55,"miss":23,"res_in":36,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"서초점","team":"서울PC/LT팀","views":468,"conn":57,"miss":34,"res_in":32,"res_req":0,"review":4,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"구로점","team":"경인PC/LT팀","views":391,"conn":60,"miss":1,"res_in":16,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"구리점","team":"서울PC/LT팀","views":378,"conn":67,"miss":33,"res_in":20,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"광교신도시점","team":"경기중부PC/LT팀","views":643,"conn":86,"miss":38,"res_in":71,"res_req":9,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"관양점","team":"경기중부PC/LT팀","views":219,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"홈플러스송도점","team":"경인PC/LT팀","views":608,"conn":109,"miss":69,"res_in":51,"res_req":9,"review":5,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"평택지산점","team":"경기중부PC/LT팀","views":244,"conn":31,"miss":22,"res_in":11,"res_req":0,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"오창점","team":"대전충청PC/LT팀","views":283,"conn":30,"miss":3,"res_in":20,"res_req":4,"review":13,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"서세종점","team":"대전충청PC/LT팀","views":336,"conn":59,"miss":11,"res_in":31,"res_req":5,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"대전점","team":"대전충청PC/LT팀","views":236,"conn":34,"miss":12,"res_in":20,"res_req":1,"review":4,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"대전서구점","team":"대전충청PC/LT팀","views":412,"conn":58,"miss":17,"res_in":26,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"송천점","team":"광주전라PC/LT팀","views":324,"conn":52,"miss":4,"res_in":30,"res_req":6,"review":4,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"전주역점","team":"광주전라PC/LT팀","views":191,"conn":28,"miss":4,"res_in":16,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"남대구IC점","team":"대구경북PC/LT팀","views":196,"conn":19,"miss":5,"res_in":16,"res_req":4,"review":4,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"수성점","team":"대구경북PC/LT팀","views":288,"conn":29,"miss":7,"res_in":11,"res_req":1,"review":5,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"부산거제점","team":"부산울산PC/LT팀","views":267,"conn":37,"miss":8,"res_in":12,"res_req":2,"review":8,"chat":0},{"yr":2026,"mo":2,"wk":1,"start":"2026-02-02","end":"2026-02-08","store":"부산수영점","team":"부산울산PC/LT팀","views":291,"conn":32,"miss":8,"res_in":15,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"한남점","team":"서울PC/LT팀","views":701,"conn":71,"miss":48,"res_in":21,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"방배점","team":"서울PC/LT팀","views":466,"conn":67,"miss":28,"res_in":33,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"서초점","team":"서울PC/LT팀","views":579,"conn":92,"miss":27,"res_in":37,"res_req":4,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"구로점","team":"경인PC/LT팀","views":491,"conn":64,"miss":0,"res_in":30,"res_req":7,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"구리점","team":"서울PC/LT팀","views":447,"conn":70,"miss":23,"res_in":36,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"광교신도시점","team":"경기중부PC/LT팀","views":688,"conn":72,"miss":36,"res_in":73,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"관양점","team":"경기중부PC/LT팀","views":259,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"홈플러스송도점","team":"경인PC/LT팀","views":699,"conn":138,"miss":69,"res_in":63,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"평택지산점","team":"경기중부PC/LT팀","views":221,"conn":14,"miss":6,"res_in":13,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"오창점","team":"대전충청PC/LT팀","views":303,"conn":29,"miss":6,"res_in":20,"res_req":4,"review":7,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"서세종점","team":"대전충청PC/LT팀","views":434,"conn":62,"miss":14,"res_in":44,"res_req":10,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"대전점","team":"대전충청PC/LT팀","views":369,"conn":56,"miss":16,"res_in":31,"res_req":4,"review":5,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"대전서구점","team":"대전충청PC/LT팀","views":388,"conn":48,"miss":24,"res_in":26,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"송천점","team":"광주전라PC/LT팀","views":361,"conn":71,"miss":8,"res_in":28,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"전주역점","team":"광주전라PC/LT팀","views":201,"conn":26,"miss":4,"res_in":15,"res_req":1,"review":5,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"남대구IC점","team":"대구경북PC/LT팀","views":269,"conn":31,"miss":5,"res_in":14,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"수성점","team":"대구경북PC/LT팀","views":373,"conn":37,"miss":10,"res_in":21,"res_req":4,"review":8,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"부산거제점","team":"부산울산PC/LT팀","views":345,"conn":49,"miss":19,"res_in":28,"res_req":5,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":2,"start":"2026-02-09","end":"2026-02-15","store":"부산수영점","team":"부산울산PC/LT팀","views":329,"conn":30,"miss":11,"res_in":17,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"한남점","team":"서울PC/LT팀","views":519,"conn":33,"miss":22,"res_in":38,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"방배점","team":"서울PC/LT팀","views":367,"conn":35,"miss":47,"res_in":37,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"서초점","team":"서울PC/LT팀","views":466,"conn":64,"miss":13,"res_in":38,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"구로점","team":"경인PC/LT팀","views":348,"conn":41,"miss":0,"res_in":33,"res_req":6,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"구리점","team":"서울PC/LT팀","views":361,"conn":35,"miss":37,"res_in":46,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"광교신도시점","team":"경기중부PC/LT팀","views":528,"conn":60,"miss":61,"res_in":55,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"관양점","team":"경기중부PC/LT팀","views":318,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"홈플러스송도점","team":"경인PC/LT팀","views":555,"conn":80,"miss":49,"res_in":56,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"평택지산점","team":"경기중부PC/LT팀","views":198,"conn":17,"miss":7,"res_in":7,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"오창점","team":"대전충청PC/LT팀","views":311,"conn":47,"miss":8,"res_in":22,"res_req":5,"review":7,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"서세종점","team":"대전충청PC/LT팀","views":582,"conn":71,"miss":10,"res_in":68,"res_req":13,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"대전점","team":"대전충청PC/LT팀","views":250,"conn":33,"miss":13,"res_in":22,"res_req":0,"review":5,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"대전서구점","team":"대전충청PC/LT팀","views":384,"conn":45,"miss":22,"res_in":30,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"송천점","team":"광주전라PC/LT팀","views":344,"conn":57,"miss":9,"res_in":22,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"전주역점","team":"광주전라PC/LT팀","views":299,"conn":42,"miss":5,"res_in":18,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"남대구IC점","team":"대구경북PC/LT팀","views":248,"conn":33,"miss":2,"res_in":14,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"수성점","team":"대구경북PC/LT팀","views":296,"conn":42,"miss":13,"res_in":18,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"부산거제점","team":"부산울산PC/LT팀","views":258,"conn":43,"miss":8,"res_in":30,"res_req":5,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":3,"start":"2026-02-16","end":"2026-02-22","store":"부산수영점","team":"부산울산PC/LT팀","views":261,"conn":27,"miss":6,"res_in":13,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"한남점","team":"서울PC/LT팀","views":773,"conn":50,"miss":27,"res_in":43,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"방배점","team":"서울PC/LT팀","views":430,"conn":76,"miss":24,"res_in":34,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"서초점","team":"서울PC/LT팀","views":553,"conn":88,"miss":19,"res_in":51,"res_req":8,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"구로점","team":"경인PC/LT팀","views":385,"conn":46,"miss":0,"res_in":27,"res_req":10,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"구리점","team":"서울PC/LT팀","views":377,"conn":55,"miss":12,"res_in":27,"res_req":7,"review":12,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"광교신도시점","team":"경기중부PC/LT팀","views":520,"conn":82,"miss":36,"res_in":51,"res_req":8,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"관양점","team":"경기중부PC/LT팀","views":231,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"홈플러스송도점","team":"경인PC/LT팀","views":575,"conn":123,"miss":58,"res_in":43,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"평택지산점","team":"경기중부PC/LT팀","views":225,"conn":24,"miss":10,"res_in":23,"res_req":1,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"오창점","team":"대전충청PC/LT팀","views":273,"conn":28,"miss":2,"res_in":18,"res_req":4,"review":6,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"서세종점","team":"대전충청PC/LT팀","views":412,"conn":56,"miss":11,"res_in":52,"res_req":15,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"대전점","team":"대전충청PC/LT팀","views":289,"conn":39,"miss":10,"res_in":27,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"대전서구점","team":"대전충청PC/LT팀","views":409,"conn":51,"miss":4,"res_in":25,"res_req":5,"review":3,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"송천점","team":"광주전라PC/LT팀","views":354,"conn":85,"miss":2,"res_in":26,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"전주역점","team":"광주전라PC/LT팀","views":331,"conn":38,"miss":9,"res_in":19,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"남대구IC점","team":"대구경북PC/LT팀","views":199,"conn":26,"miss":5,"res_in":7,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"수성점","team":"대구경북PC/LT팀","views":293,"conn":31,"miss":2,"res_in":12,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"부산거제점","team":"부산울산PC/LT팀","views":288,"conn":32,"miss":7,"res_in":14,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":2,"wk":4,"start":"2026-02-23","end":"2026-03-01","store":"부산수영점","team":"부산울산PC/LT팀","views":238,"conn":33,"miss":16,"res_in":14,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"한남점","team":"서울PC/LT팀","views":582,"conn":52,"miss":51,"res_in":22,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"방배점","team":"서울PC/LT팀","views":357,"conn":57,"miss":18,"res_in":27,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"서초점","team":"서울PC/LT팀","views":420,"conn":73,"miss":25,"res_in":33,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"구로점","team":"경인PC/LT팀","views":375,"conn":34,"miss":0,"res_in":21,"res_req":4,"review":6,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"구리점","team":"서울PC/LT팀","views":301,"conn":50,"miss":19,"res_in":23,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"광교신도시점","team":"경기중부PC/LT팀","views":436,"conn":73,"miss":37,"res_in":52,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"관양점","team":"경기중부PC/LT팀","views":214,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"홈플러스송도점","team":"경인PC/LT팀","views":478,"conn":70,"miss":69,"res_in":43,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"평택지산점","team":"경기중부PC/LT팀","views":169,"conn":19,"miss":15,"res_in":9,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"오창점","team":"대전충청PC/LT팀","views":258,"conn":39,"miss":4,"res_in":15,"res_req":2,"review":4,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"서세종점","team":"대전충청PC/LT팀","views":419,"conn":69,"miss":11,"res_in":50,"res_req":14,"review":6,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"대전점","team":"대전충청PC/LT팀","views":250,"conn":39,"miss":9,"res_in":26,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"대전서구점","team":"대전충청PC/LT팀","views":413,"conn":54,"miss":17,"res_in":26,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"송천점","team":"광주전라PC/LT팀","views":293,"conn":75,"miss":5,"res_in":34,"res_req":9,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"전주역점","team":"광주전라PC/LT팀","views":238,"conn":39,"miss":6,"res_in":19,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"남대구IC점","team":"대구경북PC/LT팀","views":149,"conn":17,"miss":15,"res_in":7,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"수성점","team":"대구경북PC/LT팀","views":272,"conn":45,"miss":4,"res_in":12,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"부산거제점","team":"부산울산PC/LT팀","views":232,"conn":25,"miss":10,"res_in":24,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":1,"start":"2026-03-02","end":"2026-03-08","store":"부산수영점","team":"부산울산PC/LT팀","views":197,"conn":27,"miss":5,"res_in":8,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"한남점","team":"서울PC/LT팀","views":674,"conn":71,"miss":36,"res_in":29,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"방배점","team":"서울PC/LT팀","views":392,"conn":58,"miss":29,"res_in":28,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"서초점","team":"서울PC/LT팀","views":497,"conn":84,"miss":31,"res_in":35,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"구로점","team":"경인PC/LT팀","views":376,"conn":45,"miss":0,"res_in":27,"res_req":5,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"구리점","team":"서울PC/LT팀","views":368,"conn":55,"miss":26,"res_in":34,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"광교신도시점","team":"경기중부PC/LT팀","views":510,"conn":55,"miss":21,"res_in":68,"res_req":8,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"관양점","team":"경기중부PC/LT팀","views":186,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"홈플러스송도점","team":"경인PC/LT팀","views":441,"conn":90,"miss":13,"res_in":35,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"평택지산점","team":"경기중부PC/LT팀","views":197,"conn":22,"miss":8,"res_in":16,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"오창점","team":"대전충청PC/LT팀","views":287,"conn":37,"miss":5,"res_in":23,"res_req":4,"review":8,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"서세종점","team":"대전충청PC/LT팀","views":405,"conn":47,"miss":13,"res_in":43,"res_req":10,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"대전점","team":"대전충청PC/LT팀","views":316,"conn":46,"miss":15,"res_in":24,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"대전서구점","team":"대전충청PC/LT팀","views":430,"conn":53,"miss":16,"res_in":29,"res_req":8,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"송천점","team":"광주전라PC/LT팀","views":265,"conn":57,"miss":3,"res_in":31,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"전주역점","team":"광주전라PC/LT팀","views":247,"conn":31,"miss":3,"res_in":16,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"남대구IC점","team":"대구경북PC/LT팀","views":216,"conn":21,"miss":17,"res_in":14,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"수성점","team":"대구경북PC/LT팀","views":342,"conn":40,"miss":6,"res_in":14,"res_req":0,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"부산거제점","team":"부산울산PC/LT팀","views":276,"conn":36,"miss":13,"res_in":15,"res_req":3,"review":3,"chat":0},{"yr":2026,"mo":3,"wk":2,"start":"2026-03-09","end":"2026-03-15","store":"부산수영점","team":"부산울산PC/LT팀","views":247,"conn":35,"miss":5,"res_in":4,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"한남점","team":"서울PC/LT팀","views":783,"conn":93,"miss":100,"res_in":41,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"방배점","team":"서울PC/LT팀","views":534,"conn":81,"miss":84,"res_in":47,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"서초점","team":"서울PC/LT팀","views":640,"conn":90,"miss":59,"res_in":51,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"구로점","team":"경인PC/LT팀","views":412,"conn":48,"miss":0,"res_in":28,"res_req":7,"review":3,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"구리점","team":"서울PC/LT팀","views":481,"conn":82,"miss":42,"res_in":41,"res_req":17,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"광교신도시점","team":"경기중부PC/LT팀","views":601,"conn":107,"miss":128,"res_in":75,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"관양점","team":"경기중부PC/LT팀","views":245,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"홈플러스송도점","team":"경인PC/LT팀","views":509,"conn":96,"miss":45,"res_in":55,"res_req":10,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"평택지산점","team":"경기중부PC/LT팀","views":190,"conn":19,"miss":10,"res_in":10,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"오창점","team":"대전충청PC/LT팀","views":328,"conn":35,"miss":4,"res_in":38,"res_req":8,"review":7,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"서세종점","team":"대전충청PC/LT팀","views":431,"conn":95,"miss":21,"res_in":46,"res_req":8,"review":5,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"대전점","team":"대전충청PC/LT팀","views":322,"conn":57,"miss":22,"res_in":28,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"대전서구점","team":"대전충청PC/LT팀","views":446,"conn":62,"miss":19,"res_in":37,"res_req":10,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"송천점","team":"광주전라PC/LT팀","views":338,"conn":55,"miss":11,"res_in":28,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"전주역점","team":"광주전라PC/LT팀","views":256,"conn":33,"miss":3,"res_in":19,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"남대구IC점","team":"대구경북PC/LT팀","views":271,"conn":29,"miss":6,"res_in":23,"res_req":8,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"수성점","team":"대구경북PC/LT팀","views":303,"conn":36,"miss":9,"res_in":17,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"부산거제점","team":"부산울산PC/LT팀","views":337,"conn":25,"miss":5,"res_in":24,"res_req":5,"review":19,"chat":0},{"yr":2026,"mo":3,"wk":3,"start":"2026-03-16","end":"2026-03-22","store":"부산수영점","team":"부산울산PC/LT팀","views":226,"conn":24,"miss":7,"res_in":13,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"한남점","team":"서울PC/LT팀","views":647,"conn":69,"miss":54,"res_in":30,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"방배점","team":"서울PC/LT팀","views":404,"conn":56,"miss":40,"res_in":46,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"서초점","team":"서울PC/LT팀","views":467,"conn":101,"miss":16,"res_in":34,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"구로점","team":"경인PC/LT팀","views":303,"conn":35,"miss":2,"res_in":31,"res_req":8,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"구리점","team":"서울PC/LT팀","views":316,"conn":80,"miss":47,"res_in":21,"res_req":9,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"광교신도시점","team":"경기중부PC/LT팀","views":494,"conn":88,"miss":20,"res_in":62,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"관양점","team":"경기중부PC/LT팀","views":308,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"홈플러스송도점","team":"경인PC/LT팀","views":488,"conn":92,"miss":41,"res_in":50,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"평택지산점","team":"경기중부PC/LT팀","views":178,"conn":20,"miss":9,"res_in":14,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"오창점","team":"대전충청PC/LT팀","views":246,"conn":43,"miss":3,"res_in":23,"res_req":8,"review":6,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"서세종점","team":"대전충청PC/LT팀","views":502,"conn":75,"miss":12,"res_in":69,"res_req":15,"review":1,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"대전점","team":"대전충청PC/LT팀","views":261,"conn":49,"miss":10,"res_in":29,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"대전서구점","team":"대전충청PC/LT팀","views":457,"conn":57,"miss":26,"res_in":28,"res_req":9,"review":3,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"송천점","team":"광주전라PC/LT팀","views":382,"conn":62,"miss":6,"res_in":31,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"전주역점","team":"광주전라PC/LT팀","views":310,"conn":40,"miss":4,"res_in":26,"res_req":1,"review":3,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"남대구IC점","team":"대구경북PC/LT팀","views":202,"conn":41,"miss":7,"res_in":9,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"수성점","team":"대구경북PC/LT팀","views":279,"conn":39,"miss":11,"res_in":17,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"부산거제점","team":"부산울산PC/LT팀","views":287,"conn":39,"miss":10,"res_in":15,"res_req":2,"review":7,"chat":0},{"yr":2026,"mo":3,"wk":4,"start":"2026-03-23","end":"2026-03-29","store":"부산수영점","team":"부산울산PC/LT팀","views":265,"conn":29,"miss":26,"res_in":8,"res_req":1,"review":5,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"한남점","team":"서울PC/LT팀","views":895,"conn":97,"miss":69,"res_in":43,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"방배점","team":"서울PC/LT팀","views":583,"conn":93,"miss":53,"res_in":49,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"서초점","team":"서울PC/LT팀","views":634,"conn":102,"miss":48,"res_in":40,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"구로점","team":"경인PC/LT팀","views":451,"conn":49,"miss":0,"res_in":32,"res_req":9,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"구리점","team":"서울PC/LT팀","views":459,"conn":95,"miss":35,"res_in":30,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"광교신도시점","team":"경기중부PC/LT팀","views":762,"conn":137,"miss":33,"res_in":88,"res_req":9,"review":1,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"관양점","team":"경기중부PC/LT팀","views":566,"conn":16,"miss":14,"res_in":0,"res_req":0,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"홈플러스송도점","team":"경인PC/LT팀","views":866,"conn":179,"miss":54,"res_in":76,"res_req":13,"review":43,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"평택지산점","team":"경기중부PC/LT팀","views":307,"conn":34,"miss":10,"res_in":19,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"오창점","team":"대전충청PC/LT팀","views":411,"conn":55,"miss":9,"res_in":31,"res_req":4,"review":14,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"서세종점","team":"대전충청PC/LT팀","views":708,"conn":105,"miss":19,"res_in":76,"res_req":21,"review":13,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"대전점","team":"대전충청PC/LT팀","views":375,"conn":65,"miss":4,"res_in":46,"res_req":8,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"대전서구점","team":"대전충청PC/LT팀","views":502,"conn":62,"miss":19,"res_in":20,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"송천점","team":"광주전라PC/LT팀","views":465,"conn":102,"miss":3,"res_in":51,"res_req":6,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"전주역점","team":"광주전라PC/LT팀","views":367,"conn":67,"miss":6,"res_in":24,"res_req":3,"review":11,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"남대구IC점","team":"대구경북PC/LT팀","views":299,"conn":44,"miss":20,"res_in":14,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"수성점","team":"대구경북PC/LT팀","views":391,"conn":44,"miss":6,"res_in":31,"res_req":5,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"부산거제점","team":"부산울산PC/LT팀","views":384,"conn":36,"miss":24,"res_in":11,"res_req":0,"review":11,"chat":0},{"yr":2026,"mo":4,"wk":1,"start":"2026-03-30","end":"2026-04-05","store":"부산수영점","team":"부산울산PC/LT팀","views":465,"conn":59,"miss":29,"res_in":22,"res_req":7,"review":3,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"한남점","team":"서울PC/LT팀","views":1340,"conn":124,"miss":80,"res_in":64,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"방배점","team":"서울PC/LT팀","views":776,"conn":118,"miss":60,"res_in":44,"res_req":6,"review":14,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"서초점","team":"서울PC/LT팀","views":850,"conn":126,"miss":38,"res_in":54,"res_req":6,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"구로점","team":"경인PC/LT팀","views":706,"conn":78,"miss":2,"res_in":34,"res_req":0,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"구리점","team":"서울PC/LT팀","views":780,"conn":94,"miss":46,"res_in":50,"res_req":14,"review":20,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"광교신도시점","team":"경기중부PC/LT팀","views":954,"conn":150,"miss":80,"res_in":110,"res_req":14,"review":8,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"관양점","team":"경기중부PC/LT팀","views":716,"conn":54,"miss":50,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"홈플러스송도점","team":"경인PC/LT팀","views":1126,"conn":190,"miss":68,"res_in":76,"res_req":8,"review":58,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"평택지산점","team":"경기중부PC/LT팀","views":406,"conn":46,"miss":22,"res_in":34,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"오창점","team":"대전충청PC/LT팀","views":546,"conn":92,"miss":0,"res_in":36,"res_req":8,"review":6,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"서세종점","team":"대전충청PC/LT팀","views":732,"conn":132,"miss":38,"res_in":84,"res_req":12,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"대전점","team":"대전충청PC/LT팀","views":770,"conn":60,"miss":30,"res_in":62,"res_req":14,"review":26,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"대전서구점","team":"대전충청PC/LT팀","views":820,"conn":106,"miss":38,"res_in":50,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"송천점","team":"광주전라PC/LT팀","views":624,"conn":150,"miss":22,"res_in":52,"res_req":10,"review":3,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"전주역점","team":"광주전라PC/LT팀","views":528,"conn":102,"miss":10,"res_in":38,"res_req":2,"review":10,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"남대구IC점","team":"대구경북PC/LT팀","views":464,"conn":50,"miss":18,"res_in":34,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"수성점","team":"대구경북PC/LT팀","views":580,"conn":88,"miss":22,"res_in":24,"res_req":4,"review":8,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"부산거제점","team":"부산울산PC/LT팀","views":516,"conn":62,"miss":12,"res_in":56,"res_req":12,"review":16,"chat":0},{"yr":2026,"mo":4,"wk":2,"start":"2026-04-06","end":"2026-04-12","store":"부산수영점","team":"부산울산PC/LT팀","views":456,"conn":52,"miss":8,"res_in":14,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"한남점","team":"서울PC/LT팀","views":1412,"conn":140,"miss":58,"res_in":54,"res_req":8,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"방배점","team":"서울PC/LT팀","views":754,"conn":110,"miss":44,"res_in":72,"res_req":6,"review":12,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"서초점","team":"서울PC/LT팀","views":890,"conn":148,"miss":26,"res_in":50,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"구로점","team":"경인PC/LT팀","views":660,"conn":82,"miss":2,"res_in":52,"res_req":10,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"구리점","team":"서울PC/LT팀","views":544,"conn":92,"miss":50,"res_in":48,"res_req":12,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"광교신도시점","team":"경기중부PC/LT팀","views":944,"conn":120,"miss":136,"res_in":100,"res_req":16,"review":6,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"관양점","team":"경기중부PC/LT팀","views":780,"conn":80,"miss":38,"res_in":48,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"홈플러스송도점","team":"경인PC/LT팀","views":1198,"conn":196,"miss":108,"res_in":104,"res_req":18,"review":62,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"평택지산점","team":"경기중부PC/LT팀","views":340,"conn":30,"miss":12,"res_in":12,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"오창점","team":"대전충청PC/LT팀","views":464,"conn":60,"miss":8,"res_in":34,"res_req":6,"review":12,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"서세종점","team":"대전충청PC/LT팀","views":696,"conn":118,"miss":20,"res_in":86,"res_req":18,"review":8,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"대전점","team":"대전충청PC/LT팀","views":708,"conn":86,"miss":24,"res_in":48,"res_req":8,"review":38,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"대전서구점","team":"대전충청PC/LT팀","views":742,"conn":78,"miss":16,"res_in":42,"res_req":8,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"송천점","team":"광주전라PC/LT팀","views":588,"conn":140,"miss":14,"res_in":66,"res_req":12,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"전주역점","team":"광주전라PC/LT팀","views":412,"conn":86,"miss":10,"res_in":16,"res_req":0,"review":10,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"남대구IC점","team":"대구경북PC/LT팀","views":414,"conn":60,"miss":16,"res_in":26,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"수성점","team":"대구경북PC/LT팀","views":622,"conn":74,"miss":16,"res_in":34,"res_req":6,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"부산거제점","team":"부산울산PC/LT팀","views":610,"conn":72,"miss":6,"res_in":24,"res_req":6,"review":8,"chat":0},{"yr":2026,"mo":4,"wk":3,"start":"2026-04-13","end":"2026-04-19","store":"부산수영점","team":"부산울산PC/LT팀","views":472,"conn":42,"miss":10,"res_in":16,"res_req":2,"review":14,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"한남점","team":"서울PC/LT팀","views":912,"conn":85,"miss":58,"res_in":31,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"방배점","team":"서울PC/LT팀","views":493,"conn":84,"miss":11,"res_in":31,"res_req":6,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"서초점","team":"서울PC/LT팀","views":654,"conn":93,"miss":21,"res_in":41,"res_req":8,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"구로점","team":"경인PC/LT팀","views":387,"conn":47,"miss":0,"res_in":26,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"구리점","team":"서울PC/LT팀","views":422,"conn":65,"miss":24,"res_in":25,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"광교신도시점","team":"경기중부PC/LT팀","views":565,"conn":72,"miss":23,"res_in":59,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"관양점","team":"경기중부PC/LT팀","views":526,"conn":40,"miss":39,"res_in":28,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"홈플러스송도점","team":"경인PC/LT팀","views":776,"conn":98,"miss":36,"res_in":63,"res_req":12,"review":30,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"평택지산점","team":"경기중부PC/LT팀","views":252,"conn":22,"miss":6,"res_in":11,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"오창점","team":"대전충청PC/LT팀","views":357,"conn":39,"miss":3,"res_in":25,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"서세종점","team":"대전충청PC/LT팀","views":504,"conn":89,"miss":12,"res_in":59,"res_req":14,"review":13,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"대전점","team":"대전충청PC/LT팀","views":439,"conn":56,"miss":14,"res_in":44,"res_req":11,"review":34,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"대전서구점","team":"대전충청PC/LT팀","views":498,"conn":55,"miss":26,"res_in":31,"res_req":9,"review":3,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"송천점","team":"광주전라PC/LT팀","views":355,"conn":72,"miss":5,"res_in":33,"res_req":5,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"전주역점","team":"광주전라PC/LT팀","views":297,"conn":63,"miss":8,"res_in":18,"res_req":2,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"남대구IC점","team":"대구경북PC/LT팀","views":234,"conn":24,"miss":11,"res_in":14,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"수성점","team":"대구경북PC/LT팀","views":459,"conn":59,"miss":4,"res_in":26,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"부산거제점","team":"부산울산PC/LT팀","views":307,"conn":33,"miss":7,"res_in":21,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":4,"wk":4,"start":"2026-04-20","end":"2026-04-26","store":"부산수영점","team":"부산울산PC/LT팀","views":309,"conn":25,"miss":14,"res_in":21,"res_req":4,"review":5,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"한남점","team":"서울PC/LT팀","views":708,"conn":46,"miss":38,"res_in":31,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"방배점","team":"서울PC/LT팀","views":253,"conn":35,"miss":7,"res_in":11,"res_req":1,"review":3,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"서초점","team":"서울PC/LT팀","views":393,"conn":58,"miss":26,"res_in":26,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"구로점","team":"경인PC/LT팀","views":336,"conn":35,"miss":0,"res_in":25,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"구리점","team":"서울PC/LT팀","views":232,"conn":42,"miss":7,"res_in":12,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"광교신도시점","team":"경기중부PC/LT팀","views":459,"conn":55,"miss":55,"res_in":52,"res_req":8,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"관양점","team":"경기중부PC/LT팀","views":323,"conn":8,"miss":4,"res_in":26,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"홈플러스송도점","team":"경인PC/LT팀","views":681,"conn":83,"miss":65,"res_in":50,"res_req":6,"review":22,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"평택지산점","team":"경기중부PC/LT팀","views":130,"conn":22,"miss":7,"res_in":4,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"오창점","team":"대전충청PC/LT팀","views":253,"conn":24,"miss":3,"res_in":19,"res_req":3,"review":10,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"서세종점","team":"대전충청PC/LT팀","views":361,"conn":44,"miss":3,"res_in":51,"res_req":10,"review":5,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"대전점","team":"대전충청PC/LT팀","views":313,"conn":38,"miss":11,"res_in":24,"res_req":2,"review":11,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"대전서구점","team":"대전충청PC/LT팀","views":349,"conn":32,"miss":15,"res_in":32,"res_req":8,"review":3,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"송천점","team":"광주전라PC/LT팀","views":282,"conn":49,"miss":3,"res_in":16,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"전주역점","team":"광주전라PC/LT팀","views":220,"conn":25,"miss":0,"res_in":17,"res_req":4,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"남대구IC점","team":"대구경북PC/LT팀","views":237,"conn":27,"miss":9,"res_in":6,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"수성점","team":"대구경북PC/LT팀","views":359,"conn":37,"miss":19,"res_in":12,"res_req":0,"review":4,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"부산거제점","team":"부산울산PC/LT팀","views":200,"conn":26,"miss":2,"res_in":24,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":4,"wk":5,"start":"2026-04-27","end":"2026-05-03","store":"부산수영점","team":"부산울산PC/LT팀","views":223,"conn":17,"miss":5,"res_in":11,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"한남점","team":"서울PC/LT팀","views":702,"conn":47,"miss":28,"res_in":22,"res_req":5,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"방배점","team":"서울PC/LT팀","views":315,"conn":41,"miss":9,"res_in":23,"res_req":1,"review":3,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"서초점","team":"서울PC/LT팀","views":365,"conn":51,"miss":19,"res_in":21,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"구로점","team":"경인PC/LT팀","views":408,"conn":31,"miss":0,"res_in":27,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"구리점","team":"서울PC/LT팀","views":280,"conn":39,"miss":11,"res_in":19,"res_req":0,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"광교신도시점","team":"경기중부PC/LT팀","views":408,"conn":45,"miss":21,"res_in":35,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"관양점","team":"경기중부PC/LT팀","views":517,"conn":18,"miss":3,"res_in":33,"res_req":7,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"홈플러스송도점","team":"경인PC/LT팀","views":607,"conn":83,"miss":46,"res_in":42,"res_req":2,"review":15,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"평택지산점","team":"경기중부PC/LT팀","views":187,"conn":20,"miss":9,"res_in":9,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"오창점","team":"대전충청PC/LT팀","views":205,"conn":44,"miss":7,"res_in":16,"res_req":1,"review":5,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"서세종점","team":"대전충청PC/LT팀","views":306,"conn":36,"miss":7,"res_in":31,"res_req":4,"review":4,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"대전점","team":"대전충청PC/LT팀","views":243,"conn":37,"miss":11,"res_in":23,"res_req":3,"review":8,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"대전서구점","team":"대전충청PC/LT팀","views":279,"conn":36,"miss":2,"res_in":16,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"송천점","team":"광주전라PC/LT팀","views":263,"conn":58,"miss":3,"res_in":20,"res_req":3,"review":5,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"전주역점","team":"광주전라PC/LT팀","views":187,"conn":21,"miss":5,"res_in":18,"res_req":4,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"남대구IC점","team":"대구경북PC/LT팀","views":208,"conn":25,"miss":6,"res_in":11,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"수성점","team":"대구경북PC/LT팀","views":269,"conn":31,"miss":4,"res_in":21,"res_req":4,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"부산거제점","team":"부산울산PC/LT팀","views":209,"conn":29,"miss":2,"res_in":27,"res_req":2,"review":8,"chat":0},{"yr":2026,"mo":5,"wk":1,"start":"2026-05-04","end":"2026-05-10","store":"부산수영점","team":"부산울산PC/LT팀","views":193,"conn":17,"miss":3,"res_in":6,"res_req":2,"review":4,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"한남점","team":"서울PC/LT팀","views":731,"conn":61,"miss":37,"res_in":34,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"방배점","team":"서울PC/LT팀","views":302,"conn":57,"miss":18,"res_in":24,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"서초점","team":"서울PC/LT팀","views":360,"conn":43,"miss":7,"res_in":16,"res_req":1,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"구로점","team":"경인PC/LT팀","views":377,"conn":37,"miss":1,"res_in":22,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"구리점","team":"서울PC/LT팀","views":274,"conn":35,"miss":19,"res_in":19,"res_req":2,"review":3,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"광교신도시점","team":"경기중부PC/LT팀","views":347,"conn":43,"miss":16,"res_in":33,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"관양점","team":"경기중부PC/LT팀","views":645,"conn":42,"miss":7,"res_in":51,"res_req":5,"review":7,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"홈플러스송도점","team":"경인PC/LT팀","views":638,"conn":78,"miss":41,"res_in":40,"res_req":2,"review":9,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"평택지산점","team":"경기중부PC/LT팀","views":117,"conn":13,"miss":1,"res_in":1,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"오창점","team":"대전충청PC/LT팀","views":185,"conn":24,"miss":0,"res_in":15,"res_req":2,"review":5,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"서세종점","team":"대전충청PC/LT팀","views":313,"conn":44,"miss":3,"res_in":41,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"대전점","team":"대전충청PC/LT팀","views":235,"conn":29,"miss":10,"res_in":25,"res_req":3,"review":12,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"대전서구점","team":"대전충청PC/LT팀","views":299,"conn":36,"miss":10,"res_in":16,"res_req":9,"review":6,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"송천점","team":"광주전라PC/LT팀","views":268,"conn":46,"miss":5,"res_in":24,"res_req":5,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"전주역점","team":"광주전라PC/LT팀","views":142,"conn":25,"miss":2,"res_in":10,"res_req":3,"review":4,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"남대구IC점","team":"대구경북PC/LT팀","views":159,"conn":18,"miss":6,"res_in":8,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"수성점","team":"대구경북PC/LT팀","views":243,"conn":27,"miss":13,"res_in":15,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"부산거제점","team":"부산울산PC/LT팀","views":188,"conn":25,"miss":4,"res_in":14,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":2,"start":"2026-05-11","end":"2026-05-17","store":"부산수영점","team":"부산울산PC/LT팀","views":229,"conn":22,"miss":5,"res_in":6,"res_req":3,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"한남점","team":"서울PC/LT팀","views":770,"conn":60,"miss":43,"res_in":37,"res_req":8,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"방배점","team":"서울PC/LT팀","views":324,"conn":64,"miss":20,"res_in":23,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"서초점","team":"서울PC/LT팀","views":336,"conn":57,"miss":6,"res_in":19,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"구로점","team":"경인PC/LT팀","views":346,"conn":42,"miss":0,"res_in":22,"res_req":3,"review":3,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"구리점","team":"서울PC/LT팀","views":312,"conn":31,"miss":8,"res_in":23,"res_req":3,"review":3,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"광교신도시점","team":"경기중부PC/LT팀","views":365,"conn":40,"miss":20,"res_in":41,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"관양점","team":"경기중부PC/LT팀","views":636,"conn":45,"miss":15,"res_in":39,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"홈플러스송도점","team":"경인PC/LT팀","views":597,"conn":76,"miss":71,"res_in":50,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"평택지산점","team":"경기중부PC/LT팀","views":159,"conn":15,"miss":5,"res_in":12,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"오창점","team":"대전충청PC/LT팀","views":264,"conn":45,"miss":5,"res_in":20,"res_req":6,"review":6,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"서세종점","team":"대전충청PC/LT팀","views":370,"conn":65,"miss":16,"res_in":42,"res_req":6,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"대전점","team":"대전충청PC/LT팀","views":281,"conn":37,"miss":18,"res_in":35,"res_req":5,"review":5,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"대전서구점","team":"대전충청PC/LT팀","views":311,"conn":42,"miss":32,"res_in":15,"res_req":5,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"송천점","team":"광주전라PC/LT팀","views":285,"conn":48,"miss":4,"res_in":22,"res_req":1,"review":4,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"전주역점","team":"광주전라PC/LT팀","views":212,"conn":37,"miss":3,"res_in":19,"res_req":1,"review":9,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"남대구IC점","team":"대구경북PC/LT팀","views":198,"conn":28,"miss":5,"res_in":15,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"수성점","team":"대구경북PC/LT팀","views":245,"conn":33,"miss":3,"res_in":13,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"부산거제점","team":"부산울산PC/LT팀","views":195,"conn":19,"miss":6,"res_in":12,"res_req":0,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":3,"start":"2026-05-18","end":"2026-05-24","store":"부산수영점","team":"부산울산PC/LT팀","views":284,"conn":21,"miss":3,"res_in":10,"res_req":1,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"한남점","team":"서울PC/LT팀","views":640,"conn":58,"miss":34,"res_in":28,"res_req":5,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"방배점","team":"서울PC/LT팀","views":307,"conn":37,"miss":14,"res_in":26,"res_req":5,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"서초점","team":"서울PC/LT팀","views":345,"conn":31,"miss":12,"res_in":20,"res_req":3,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"구로점","team":"경인PC/LT팀","views":342,"conn":30,"miss":0,"res_in":27,"res_req":3,"review":3,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"구리점","team":"서울PC/LT팀","views":327,"conn":30,"miss":17,"res_in":18,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"광교신도시점","team":"경기중부PC/LT팀","views":410,"conn":61,"miss":37,"res_in":40,"res_req":4,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"관양점","team":"경기중부PC/LT팀","views":645,"conn":41,"miss":7,"res_in":44,"res_req":7,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"홈플러스송도점","team":"경인PC/LT팀","views":503,"conn":63,"miss":37,"res_in":43,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"평택지산점","team":"경기중부PC/LT팀","views":130,"conn":10,"miss":7,"res_in":8,"res_req":2,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"오창점","team":"대전충청PC/LT팀","views":215,"conn":33,"miss":4,"res_in":15,"res_req":2,"review":2,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"서세종점","team":"대전충청PC/LT팀","views":413,"conn":57,"miss":12,"res_in":58,"res_req":10,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"대전점","team":"대전충청PC/LT팀","views":317,"conn":56,"miss":17,"res_in":32,"res_req":3,"review":4,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"대전서구점","team":"대전충청PC/LT팀","views":348,"conn":47,"miss":19,"res_in":24,"res_req":3,"review":0,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"송천점","team":"광주전라PC/LT팀","views":248,"conn":63,"miss":5,"res_in":11,"res_req":2,"review":5,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"전주역점","team":"광주전라PC/LT팀","views":214,"conn":36,"miss":6,"res_in":15,"res_req":1,"review":7,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"남대구IC점","team":"대구경북PC/LT팀","views":174,"conn":21,"miss":7,"res_in":4,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"수성점","team":"대구경북PC/LT팀","views":223,"conn":23,"miss":8,"res_in":12,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"부산거제점","team":"부산울산PC/LT팀","views":175,"conn":33,"miss":5,"res_in":6,"res_req":1,"review":1,"chat":0},{"yr":2026,"mo":5,"wk":4,"start":"2026-05-25","end":"2026-05-31","store":"부산수영점","team":"부산울산PC/LT팀","views":251,"conn":28,"miss":13,"res_in":13,"res_req":2,"review":1,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"한남점","team":"서울PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"방배점","team":"서울PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"서초점","team":"서울PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"구로점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"구리점","team":"서울PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"광교신도시점","team":"경기중부PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"관양점","team":"경기중부PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"홈플러스송도점","team":"경인PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"평택지산점","team":"경기중부PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"오창점","team":"대전충청PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"서세종점","team":"대전충청PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"대전점","team":"대전충청PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"대전서구점","team":"대전충청PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"송천점","team":"광주전라PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"전주역점","team":"광주전라PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"남대구IC점","team":"대구경북PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"수성점","team":"대구경북PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"부산거제점","team":"부산울산PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0},{"yr":2026,"mo":6,"wk":1,"start":"2026-06-01","end":"2026-06-07","store":"부산수영점","team":"부산울산PC/LT팀","views":0,"conn":0,"miss":0,"res_in":0,"res_req":0,"review":0,"chat":0}];

// ─── 주차 메타데이터 구축 ─────────────────────────────────────────
let weekMeta = {}; // "yr-mo-wk" → {start, end, label}
let weekKeys = [];
let STORES = [];
let TEAMS  = [];

function rebuildMetaFromRaw() {
  weekMeta = {};
  RAW.forEach(r => {
    const key = `${r.yr}-${r.mo}-${r.wk}`;
    if (!weekMeta[key]) {
      weekMeta[key] = {
        yr: r.yr, mo: r.mo, wk: r.wk,
        start: r.start, end: r.end,
        label: `${r.yr}년 ${r.mo}월 ${r.wk}주차`,
        short: `${r.mo}월${r.wk}주`
      };
    }
  });
  weekKeys = Object.keys(weekMeta).sort();
  STORES = [...new Set(RAW.map(r=>r.store).filter(Boolean))].sort();
  TEAMS  = [...new Set(RAW.map(r=>r.team).filter(Boolean))].sort();
}

rebuildMetaFromRaw();

// ─── 셀렉트 초기화 ───────────────────────────────────────────────
function initSelects() {
  const selWeek = document.getElementById('sel-week');
  selWeek.innerHTML = '';
  weekKeys.forEach(k => {
    const m = weekMeta[k];
    const opt = document.createElement('option');
    opt.value = k;
    opt.textContent = `${m.label}  (${m.start} ~ ${m.end})`;
    selWeek.appendChild(opt);
  });
  // 최신 주차 기본 선택
  if (weekKeys.length) selWeek.value = weekKeys[weekKeys.length - 1];

  const selStore = document.getElementById('sel-store');
  selStore.innerHTML = '<option value="ALL">전체 매장</option>';
  STORES.forEach(s => {
    const opt = document.createElement('option');
    opt.value = s; opt.textContent = s;
    selStore.appendChild(opt);
  });

  const selTeam = document.getElementById('sel-team');
  selTeam.innerHTML = '<option value="ALL">전체 팀</option>';
  TEAMS.forEach(t => {
    const opt = document.createElement('option');
    opt.value = t; opt.textContent = t;
    selTeam.appendChild(opt);
  });
}

// ─── 집계 헬퍼 ───────────────────────────────────────────────────
const METRICS = ['views','conn','miss','res_in','res_req','review','chat'];
function sumRows(rows) {
  const s = {views:0,conn:0,miss:0,res_in:0,res_req:0,review:0,chat:0};
  rows.forEach(r => METRICS.forEach(m => s[m] += r[m]));
  return s;
}

function getRows(yr, mo, wk, store='ALL', team='ALL') {
  return RAW.filter(r =>
    r.yr===yr && r.mo===mo && r.wk===wk &&
    (store==='ALL' || r.store===store) &&
    (team==='ALL'  || r.team===team)
  );
}

// 전월 동기: 같은 주차 번호, 바로 전 월
function prevMonthKey(yr, mo, wk) {
  let pmo = mo-1, pyr = yr;
  if (pmo < 1) { pmo = 12; pyr--; }
  const k = `${pyr}-${pmo}-${wk}`;
  return weekMeta[k] ? {yr:pyr, mo:pmo, wk} : null;
}
// 전년 동기: 같은 월+주, 1년 전
function prevYearKey(yr, mo, wk) {
  const k = `${yr-1}-${mo}-${wk}`;
  return weekMeta[k] ? {yr:yr-1, mo, wk} : null;
}

function delta(cur, prev, field) {
  if (!prev || prev[field]===0) return null;
  return ((cur[field]-prev[field])/prev[field]*100).toFixed(1);
}

// ─── KPI 카드 렌더 ───────────────────────────────────────────────
// 지표 그룹 정의
const KPI_GROUPS = [
  {
    id: 'exposure', label: '노출 지표', desc: '플레이스 노출 및 도달 성과',
    icon: `<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>`,
    keys: ['views'],
  },
  {
    id: 'convert', label: '전환 지표', desc: '콜 연결 및 예약 전환 성과',
    icon: `<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg>`,
    keys: ['conn', 'res_req'],
  },
  {
    id: 'manage', label: '관리 지표', desc: '미연결 개선 및 리뷰 관리 성과',
    icon: `<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>`,
    keys: ['miss', 'review'],
  },
];

const KPI_DEFS = [
  {key:'views',   label:'플레이스 조회수', cls:'views',   group:'exposure', higher:true,  icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>` },
  {key:'conn',    label:'연결콜',          cls:'conn',    group:'convert',  higher:true,  icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.69 12 19.79 19.79 0 0 1 1.61 3.41 2 2 0 0 1 3.6 1.22h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L7.91 8.84a16 16 0 0 0 6 6l.93-.93a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7a2 2 0 0 1 1.72 2.03z"/></svg>` },
  {key:'res_req', label:'예약 신청',       cls:'resreq',  group:'convert',  higher:true,  icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>` },
  {key:'miss',    label:'미연결콜',        cls:'miss',    group:'manage',   higher:false, icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="1" y1="1" x2="23" y2="23"/><path d="M16.72 11.06A10.94 10.94 0 0 1 19 12.55M5 19a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2h1.22M10.66 5H19a2 2 0 0 1 2 2v10.34M11 11a4 4 0 0 0 5.17 5.17"/></svg>` },
  {key:'review',  label:'리뷰',           cls:'review',  group:'manage',   higher:true,  icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>` },
  {key:'res_in',  label:'예약 유입',       cls:'resin',   group:'exposure', higher:true,  icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/><path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/></svg>` },
  {key:'chat',    label:'톡톡 상담',       cls:'chat',    group:'convert',  higher:true,  icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>` },
];

function renderKPI(cur, prevMo, prevYr) {
  const grid = document.getElementById('kpi-grid');

  // 증감률 계산
  function calcRate(v, ref, higher) {
    if (ref == null || ref === 0) return null;
    const d = (v - ref) / ref * 100;
    return { d, good: higher ? d >= 0 : d <= 0 };
  }

  function rateBadge(v, ref, higher, label) {
    const r = calcRate(v, ref, higher);
    if (!r) return `<span class="rate-item"><span class="rl">${label}</span><span class="rv rate-flat">—</span></span>`;
    const cls = r.good ? 'rate-up' : 'rate-dn';
    const arrow = r.d > 0 ? '▲' : r.d < 0 ? '▼' : '';
    return `<span class="rate-item"><span class="rl">${label}</span><span class="rv ${cls}">${arrow}${Math.abs(r.d).toFixed(1)}%</span></span>`;
  }

  // 그룹별 섹션 생성
  let html_out = '';
  KPI_GROUPS.forEach(grp => {
    const grpKeys = KPI_DEFS.filter(d => d.group === grp.id);

    // 그룹 합산 (조회수 제외 건수 합)
    const curTotal  = grpKeys.reduce((s,d) => s + (cur[d.key]||0), 0);
    const prevTotal = prevMo ? grpKeys.reduce((s,d) => s + (prevMo[d.key]||0), 0) : null;
    const grpRate   = prevTotal ? ((curTotal - prevTotal) / prevTotal * 100) : null;

    html_out += `<div class="kpi-group-wrap">
      <div class="kpi-group-header">
        <span class="kpi-group-badge ${grp.id}">${grp.icon} ${grp.label}</span>
        <span style="font-size:11px;color:var(--text3)">${grp.desc}</span>
        <div class="kpi-group-divider"></div>
        ${grpRate !== null ? `<div class="kpi-group-score">
          전월 대비 <span class="score-num" style="color:${grpRate>=0?'var(--green)':'var(--red)'}">${grpRate>=0?'▲':'▼'}${Math.abs(grpRate).toFixed(1)}%</span>
        </div>` : ''}
      </div>
      <div class="kpi-grid" style="margin-bottom:0">
        ${grpKeys.map(def => {
          const v  = cur[def.key] || 0;
          const pm = prevMo  ? prevMo[def.key]  : null;
          const py = prevYr  ? prevYr[def.key]  : null;
          const rMo = calcRate(v, pm, def.higher);
          const rYr = calcRate(v, py, def.higher);

          // 핵심 메시지
          const getMsg = (rate, label, isHigher) => {
            if (!rate) return '';
            const abs = Math.abs(rate.d).toFixed(1);
            if (def.key === 'miss') {
              return rate.d < 0
                ? `<div style="font-size:10px;color:var(--green);margin-top:6px;font-weight:600">${label} 미연결 개선 ▼${abs}%</div>`
                : rate.d > 0 ? `<div style="font-size:10px;color:var(--red);margin-top:6px;font-weight:600">${label} 미연결 증가 ▲${abs}%</div>` : '';
            }
            if (def.key === 'conn') {
              return rate.d > 0
                ? `<div style="font-size:10px;color:var(--green);margin-top:6px;font-weight:600">${label} 연결콜 증가 ▲${abs}%</div>`
                : rate.d < 0 ? `<div style="font-size:10px;color:var(--red);margin-top:6px;font-weight:600">${label} 연결콜 감소 ▼${abs}%</div>` : '';
            }
            return '';
          };

          return `<div class="kpi-card ${def.cls}">
            <div class="kpi-label">${def.icon} ${def.label}</div>
            <div class="kpi-value">${v.toLocaleString()}</div>
            <div class="kpi-compare">
              <div class="kpi-compare-row">
                <span class="compare-label">전월 동기</span>
                <span class="compare-val">${pm!=null?pm.toLocaleString():'—'}</span>
                ${rateBadge(v, pm, def.higher, '')}
              </div>
              <div class="kpi-compare-row">
                <span class="compare-label">전년 동기</span>
                <span class="compare-val">${py!=null?py.toLocaleString():'—'}</span>
                ${rateBadge(v, py, def.higher, '')}
              </div>
            </div>
            <div class="rate-row">
              ${rateBadge(v, pm, def.higher, '전월')}
              ${rateBadge(v, py, def.higher, '전년')}
            </div>
            ${rMo ? getMsg(rMo, '전월 대비', def.higher) : ''}
          </div>`;
        }).join('')}
      </div>
    </div>`;
  });

  grid.innerHTML = html_out;
}

// ─── 차트 ─────────────────────────────────────────────────────────
let charts = {};
const CHART_COLORS = {
  cur:   'rgba(245,166,35,1)',
  curBg: 'rgba(245,166,35,.15)',
  prev:  'rgba(75,124,243,1)',
  prevBg:'rgba(75,124,243,.12)',
  yoy:   'rgba(155,108,243,1)',
  yoyBg: 'rgba(155,108,243,.12)',
  miss:  'rgba(242,92,110,1)',
  missBg:'rgba(242,92,110,.12)',
  green: 'rgba(52,196,138,1)',
  greenBg:'rgba(52,196,138,.12)',
};

function chartDefaults() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { labels: { color:'#8b91a8', font:{family:'Pretendard',size:11}, boxWidth:12, padding:16 } },
      tooltip: { backgroundColor:'#1e2230', borderColor:'#2a2f42', borderWidth:1, titleColor:'#e8eaf0', bodyColor:'#8b91a8', padding:12 }
    },
    scales: {
      x: { ticks:{color:'#5a6080',font:{size:10}}, grid:{color:'rgba(42,47,66,.5)'} },
      y: { ticks:{color:'#5a6080',font:{size:10}}, grid:{color:'rgba(42,47,66,.5)'} }
    }
  };
}

function makeBarChart(id, labels, datasets) {
  if (charts[id]) charts[id].destroy();
  const ctx = document.getElementById(id).getContext('2d');
  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: { ...chartDefaults(), barPercentage:.65, categoryPercentage:.7 }
  });
}

function renderCharts(cur, prevMo, prevYr, meta) {
  const labs = ['당월 현재', '전월 동기', '전년 동기'];

  // 콜
  makeBarChart('chart-call', labs, [
    { label:'연결콜', data:[cur.conn, prevMo?.conn??0, prevYr?.conn??0],
      backgroundColor:[CHART_COLORS.cur, CHART_COLORS.prev, CHART_COLORS.yoy] },
    { label:'미연결콜', data:[cur.miss, prevMo?.miss??0, prevYr?.miss??0],
      backgroundColor:[CHART_COLORS.miss, 'rgba(242,92,110,.5)', 'rgba(242,92,110,.3)'] }
  ]);

  // 조회수
  makeBarChart('chart-views', labs, [
    { label:'조회수', data:[cur.views, prevMo?.views??0, prevYr?.views??0],
      backgroundColor:[CHART_COLORS.cur, CHART_COLORS.prev, CHART_COLORS.yoy] }
  ]);

  // 예약
  makeBarChart('chart-res', labs, [
    { label:'예약유입', data:[cur.res_in, prevMo?.res_in??0, prevYr?.res_in??0],
      backgroundColor:[CHART_COLORS.green, 'rgba(52,196,138,.5)', 'rgba(52,196,138,.3)'] },
    { label:'예약신청', data:[cur.res_req, prevMo?.res_req??0, prevYr?.res_req??0],
      backgroundColor:['#00b4d8','rgba(0,180,216,.5)','rgba(0,180,216,.3)'] }
  ]);

  // 리뷰+채팅
  makeBarChart('chart-etc', labs, [
    { label:'리뷰', data:[cur.review, prevMo?.review??0, prevYr?.review??0],
      backgroundColor:['#f9c74f','rgba(249,199,79,.5)','rgba(249,199,79,.3)'] },
    { label:'톡톡상담', data:[cur.chat, prevMo?.chat??0, prevYr?.chat??0],
      backgroundColor:[CHART_COLORS.yoy, 'rgba(155,108,243,.5)', 'rgba(155,108,243,.3)'] }
  ]);
}

// ─── 매장 테이블 ──────────────────────────────────────────────────
function renderStoreTable(yr, mo, wk, store, team) {
  const tbody = document.getElementById('store-tbody');
  const rows = getRows(yr, mo, wk, store, team);

  const totals = sumRows(rows);
  let html = rows.map(r => `
    <tr>
      <td>${r.store}</td><td>${r.team}</td>
      <td>${r.views.toLocaleString()}</td>
      <td>${r.conn.toLocaleString()}</td>
      <td>${r.miss.toLocaleString()}</td>
      <td>${r.res_in.toLocaleString()}</td>
      <td>${r.res_req.toLocaleString()}</td>
      <td>${r.review.toLocaleString()}</td>
      <td>${r.chat.toLocaleString()}</td>
    </tr>`).join('');

  html += `<tr class="tr-total">
    <td>합계</td><td>—</td>
    <td>${totals.views.toLocaleString()}</td>
    <td>${totals.conn.toLocaleString()}</td>
    <td>${totals.miss.toLocaleString()}</td>
    <td>${totals.res_in.toLocaleString()}</td>
    <td>${totals.res_req.toLocaleString()}</td>
    <td>${totals.review.toLocaleString()}</td>
    <td>${totals.chat.toLocaleString()}</td>
  </tr>`;
  tbody.innerHTML = html;
}

// ─── 상세 테이블 ─────────────────────────────────────────────────
function renderDetailTable(yr, mo, wk, store, team) {
  const tbody = document.getElementById('detail-tbody');
  const curRows = getRows(yr, mo, wk, store, team);
  const pmKey = prevMonthKey(yr, mo, wk);
  const pyKey = prevYearKey(yr, mo, wk);
  const pmRows = pmKey ? getRows(pmKey.yr, pmKey.mo, pmKey.wk, store, team) : [];
  const pyRows = pyKey ? getRows(pyKey.yr, pyKey.mo, pyKey.wk, store, team) : [];

  const storeSet = [...new Set(curRows.map(r=>r.store))];
  let html = '';

  storeSet.forEach(s => {
    const tr = curRows.find(r=>r.store===s);
    const pm = pmRows.find(r=>r.store===s);
    const py = pyRows.find(r=>r.store===s);
    const teamName = tr?.team||'';

    const row = (label, d, cls) => {
      if (!d) return '';
      return `<tr style="opacity:${cls?1:.65}">
        <td${cls?' style="font-weight:700"':''}>${cls?s:''}</td>
        <td>${cls?teamName:''}</td>
        <td><span class="compare-delta ${cls?'delta-up':'delta-flat'}" style="font-size:10px">${label}</span></td>
        <td>${d.views.toLocaleString()}</td>
        <td>${d.conn.toLocaleString()}</td>
        <td>${d.miss.toLocaleString()}</td>
        <td>${d.res_in.toLocaleString()}</td>
        <td>${d.res_req.toLocaleString()}</td>
        <td>${d.review.toLocaleString()}</td>
        <td>${d.chat.toLocaleString()}</td>
      </tr>`;
    };

    html += row('당 월', tr, true);
    html += row('전 월', pm, false);
    html += row('전 년', py, false);
  });

  tbody.innerHTML = html || '<tr><td colspan="10" style="text-align:center;padding:40px;color:var(--text3)">데이터가 없습니다</td></tr>';
}

// ─── 기간 헤더 렌더 ──────────────────────────────────────────────
function renderPeriodHeader(yr, mo, wk) {
  const meta = weekMeta[`${yr}-${mo}-${wk}`];
  const pmKey = prevMonthKey(yr, mo, wk);
  const pyKey = prevYearKey(yr, mo, wk);

  document.getElementById('ph-title').innerHTML =
    `<span>${yr}년 ${mo}월 ${wk}주차</span>`;

  let datesHtml = `<div class="period-date-row current">
    <span class="label">당 월</span>
    <span>${meta.start} (월) ~ ${meta.end} (일)</span>
  </div>`;

  if (pmKey) {
    const pm = weekMeta[`${pmKey.yr}-${pmKey.mo}-${pmKey.wk}`];
    if (pm) datesHtml += `<div class="period-date-row prev">
      <span class="label">전월동기</span>
      <span>${pm.start} ~ ${pm.end}  (${pm.yr}년 ${pm.mo}월 ${pm.wk}주)</span>
    </div>`;
  }
  if (pyKey) {
    const py = weekMeta[`${pyKey.yr}-${pyKey.mo}-${pyKey.wk}`];
    if (py) datesHtml += `<div class="period-date-row yoy">
      <span class="label">전년동기</span>
      <span>${py.start} ~ ${py.end}  (${py.yr}년 ${py.mo}월 ${py.wk}주)</span>
    </div>`;
  }

  document.getElementById('ph-dates').innerHTML = datesHtml;
  document.getElementById('period-header').style.display = 'flex';
  document.getElementById('period-display').textContent =
    `${meta.start} ~ ${meta.end}`;
}

// ─── 메인 조회 ───────────────────────────────────────────────────
function applyFilter() {
  const weekKey = document.getElementById('sel-week').value;
  const store   = document.getElementById('sel-store').value;
  const team    = document.getElementById('sel-team').value;
  const meta    = weekMeta[weekKey];
  if (!meta) return;

  const {yr, mo, wk} = meta;
  const pmKey = prevMonthKey(yr, mo, wk);
  const pyKey = prevYearKey(yr, mo, wk);

  const curSum  = sumRows(getRows(yr, mo, wk, store, team));
  const prevMoSum = pmKey ? sumRows(getRows(pmKey.yr, pmKey.mo, pmKey.wk, store, team)) : null;
  const prevYrSum = pyKey ? sumRows(getRows(pyKey.yr, pyKey.mo, pyKey.wk, store, team)) : null;

  renderPeriodHeader(yr, mo, wk);
  renderKPI(curSum, prevMoSum, prevYrSum);
  renderCharts(curSum, prevMoSum, prevYrSum, meta);
  renderStoreTable(yr, mo, wk, store, team);
  renderDetailTable(yr, mo, wk, store, team);
}

// ─── 엑셀 다운로드 ───────────────────────────────────────────────
async function downloadExcel() {
  await ensureXlsx();
  const weekKey = document.getElementById('sel-week').value;
  const store   = document.getElementById('sel-store').value;
  const team    = document.getElementById('sel-team').value;
  const meta    = weekMeta[weekKey];

  const rows = getRows(meta.yr, meta.mo, meta.wk, store, team);
  const pmKey = prevMonthKey(meta.yr, meta.mo, meta.wk);
  const pyKey = prevYearKey(meta.yr, meta.mo, meta.wk);
  const pmRows = pmKey ? getRows(pmKey.yr, pmKey.mo, pmKey.wk, store, team) : [];
  const pyRows = pyKey ? getRows(pyKey.yr, pyKey.mo, pyKey.wk, store, team) : [];

  const toSheet = (data, label) => data.map(r => ({
    구분: label, 매장명: r.store, 팀: r.team,
    년: r.yr, 월: r.mo, 주차: r.wk,
    시작일: r.start, 종료일: r.end,
    조회수: r.views, 연결콜: r.conn, 미연결콜: r.miss,
    예약유입: r.res_in, 예약신청: r.res_req,
    리뷰: r.review, 톡톡상담: r.chat
  }));

  const allData = [
    ...toSheet(rows, '당월'),
    ...toSheet(pmRows, '전월동기'),
    ...toSheet(pyRows, '전년동기'),
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(allData);
  ws['!cols'] = Array(15).fill({wch:12});
  XLSX.utils.book_append_sheet(wb, ws, '주간보고서');
  XLSX.writeFile(wb, `티스테이션_${meta.label}_${meta.start}_${meta.end}.xlsx`);
}

// ─── 탭 전환 ─────────────────────────────────────────────────────
function switchTab(id, btn) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-'+id).classList.add('active');
  btn.classList.add('active');

  // 숨겨진 탭 상태에서 차트를 만들면 캔버스 크기가 0으로 잡히는 경우가 있어, 탭 활성화 후 렌더링
  requestAnimationFrame(() => {
    if (id === 'tna') renderTna();
    if (id === 'monthly') renderMonthly();
  });
}

// ─── 구글 시트 연동 (선택사항) ───────────────────────────────────
async function loadFromGoogleSheets(forceFresh = false) {
  if (!DATA_URL) return false;
  try {
    const rows = await fetchJsonWithCache(DATA_URL, 'tna_place_rows_v2', 5 * 60 * 1000, forceFresh);
    // RAW 배열 교체
    RAW.length = 0;
    rows.forEach(r => {
      // 구글 시트 컬럼명 → 내부 키 매핑
      const base = new Date(1899,11,30);
      const toDate = s => {
        if (!s) return '';
        if (typeof s === 'number') {
          const d = new Date(base); d.setDate(base.getDate()+s);
          return d.toISOString().split('T')[0];
        }
        return String(s).split('T')[0];
      };
      RAW.push({
        yr: Number(r['년']||r.yr), mo: Number(r['월']||r.mo), wk: Number(r['주']||r.wk),
        start: toDate(r['시작일(월)']||r.start), end: toDate(r['종료일(일)']||r.end),
        store: r['매장명']||r.store, team: r['팀']||r.team,
        views: Number(r['조회수']||r.views||0), conn: Number(r['연결콜']||r.conn||0),
        miss: Number(r['미연결콜']||r.miss||0), res_in: Number(r['예약유입']||r.res_in||0),
        res_req: Number(r['예약신청']||r.res_req||0), review: Number(r['리뷰']||r.review||0),
        chat: Number(r['톡톡상담']||r.chat||0),
      });
    });
    return true;
  } catch(e) {
    console.warn('구글 시트 연동 실패, 내장 데이터 사용:', e);
    return false;
  }
}


// ─── 타이어앤 판매 데이터 ─────────────────────────────────────────
let TNA = {"stores":[{"store":"방배점","monthly":{"1":{"qty":12,"amt":1347720},"2":{"qty":0,"amt":0},"3":{"qty":4,"amt":1370160},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":16,"total_amt":2717880},{"store":"서초점","monthly":{"1":{"qty":4,"amt":424040},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":4,"total_amt":424040},{"store":"광교점","monthly":{"1":{"qty":4,"amt":477840},"2":{"qty":4,"amt":1137840},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":24,"amt":4447080},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":32,"total_amt":6062760},{"store":"대전서구점","monthly":{"1":{"qty":4,"amt":501600},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":8,"amt":1045440},"5":{"qty":2,"amt":231000},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":14,"total_amt":1778040},{"store":"송천점","monthly":{"1":{"qty":12,"amt":1515360},"2":{"qty":0,"amt":0},"3":{"qty":4,"amt":599280},"4":{"qty":10,"amt":1389960},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":26,"total_amt":3504600},{"store":"서세종점","monthly":{"1":{"qty":12,"amt":1599200},"2":{"qty":12,"amt":1386000},"3":{"qty":12,"amt":1424280},"4":{"qty":0,"amt":0},"5":{"qty":2,"amt":226040},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":38,"total_amt":4635520},{"store":"대전점","monthly":{"1":{"qty":0,"amt":0},"2":{"qty":0,"amt":0},"3":{"qty":4,"amt":454080},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":4,"total_amt":454080},{"store":"한남점","monthly":{"1":{"qty":16,"amt":1935120},"2":{"qty":4,"amt":747120},"3":{"qty":0,"amt":0},"4":{"qty":2,"amt":336600},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":22,"total_amt":3018840},{"store":"홈플러스송도점","monthly":{"1":{"qty":6,"amt":728640},"2":{"qty":8,"amt":799920},"3":{"qty":0,"amt":0},"4":{"qty":2,"amt":289080},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":16,"total_amt":1817640},{"store":"구리점","monthly":{"1":{"qty":0,"amt":0},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":4,"amt":543840},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":4,"total_amt":543840},{"store":"남대구IC점","monthly":{"1":{"qty":4,"amt":567600},"2":{"qty":2,"amt":184800},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":4,"amt":462000},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":10,"total_amt":1214400},{"store":"수성점","monthly":{"1":{"qty":0,"amt":0},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":0,"total_amt":0},{"store":"부산거제점","monthly":{"1":{"qty":6,"amt":918720},"2":{"qty":4,"amt":448800},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":10,"total_amt":1367520},{"store":"오창점","monthly":{"1":{"qty":0,"amt":0},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":0,"total_amt":0},{"store":"전주역점","monthly":{"1":{"qty":12,"amt":1401840},"2":{"qty":4,"amt":520080},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":4,"amt":634920},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":20,"total_amt":2556840},{"store":"구로점","monthly":{"1":{"qty":6,"amt":904520},"2":{"qty":0,"amt":0},"3":{"qty":4,"amt":264000},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":10,"total_amt":1168520},{"store":"부산수영점","monthly":{"1":{"qty":0,"amt":0},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":0,"total_amt":0},{"store":"관양점","monthly":{"1":{"qty":0,"amt":0},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":2,"amt":450120},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":2,"total_amt":450120},{"store":"평택지산점","monthly":{"1":{"qty":0,"amt":0},"2":{"qty":0,"amt":0},"3":{"qty":0,"amt":0},"4":{"qty":0,"amt":0},"5":{"qty":0,"amt":0},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}},"total_qty":0,"total_amt":0}],"monthly":{"1":{"qty":98,"amt":1232},"2":{"qty":38,"amt":522},"3":{"qty":28,"amt":411},"4":{"qty":22,"amt":306},"5":{"qty":42,"amt":700},"6":{"qty":0,"amt":0},"7":{"qty":0,"amt":0},"8":{"qty":0,"amt":0},"9":{"qty":0,"amt":0},"10":{"qty":0,"amt":0},"11":{"qty":0,"amt":0},"12":{"qty":0,"amt":0}}};

// 타이어앤 매장 셀렉트 초기화
function initTnaSelects() {
  const sel = document.getElementById('tna-sel-store');
  if (!sel) return;
  const prev = sel.value || 'ALL';
  sel.innerHTML = '<option value="ALL">전체 매장</option>';
  TNA.stores.forEach(d => {
    const o = document.createElement('option');
    o.value = d.store; o.textContent = d.store;
    sel.appendChild(o);
  });
  if ([...sel.options].some(o => o.value === prev)) sel.value = prev;

  const selYear = document.getElementById('tna-sel-year');
  if (selYear && !selYear.dataset.bound) {
    selYear.dataset.bound = '1';
    selYear.addEventListener('change', () => {
      rebuildTnaFromLiveRows();
      renderTna();
    });
  }
}

let tnaCharts = {};

function buildTnaFromRows(rows) {
  const MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];
  const byStore = {};
  const monthly = {};
  MONTHS.forEach(m => monthly[m] = { qty:0, amt:0 });

  rows.forEach(r => {
    const yr = Number(r.yr ?? r['년'] ?? 0);
    const mo = Number(r.mo ?? r['월'] ?? 0);
    const store = String(r.store ?? r['매장명'] ?? '').trim();
    const qty = Number(String(r.qty ?? r['판매개수'] ?? r['판매갯수'] ?? r['판매수량'] ?? 0).replace(/,/g,'')) || 0;
    const amt = Number(String(r.amount ?? r.amt ?? r['판매금액'] ?? r['판매 금액'] ?? 0).replace(/,/g,'')) || 0;

    if (!yr || !mo || !store) return;
    if (!byStore[store]) {
      byStore[store] = { store, monthly:{}, total_qty:0, total_amt:0 };
      MONTHS.forEach(m => byStore[store].monthly[m] = { qty:0, amt:0 });
    }
    byStore[store].monthly[mo].qty += qty;
    byStore[store].monthly[mo].amt += amt;
    byStore[store].total_qty += qty;
    byStore[store].total_amt += amt;

    monthly[mo].qty += qty;
    monthly[mo].amt += amt;
  });

  return {
    stores: Object.values(byStore).sort((a,b)=>a.store.localeCompare(b.store,'ko')),
    monthly
  };
}

let TNA_LIVE_ROWS = [];

function updateTnaYearSelectFromRows(rows) {
  const selYear = document.getElementById('tna-sel-year');
  if (!selYear || !Array.isArray(rows) || rows.length === 0) return;

  const years = [...new Set(rows.map(r => Number(r.yr ?? r['년'] ?? 0)).filter(Boolean))]
    .sort((a, b) => b - a);

  if (!years.length) return;

  const prev = selYear.value;
  selYear.innerHTML = '';
  years.forEach(yr => {
    const o = document.createElement('option');
    o.value = String(yr);
    o.textContent = yr + '년';
    selYear.appendChild(o);
  });

  if (years.map(String).includes(prev)) {
    selYear.value = prev;
  } else {
    selYear.value = String(years[0]);
  }
}

function getTnaSelectedYear() {
  const selYear = document.getElementById('tna-sel-year');
  return Number(selYear && selYear.value ? selYear.value : 2026);
}

function rebuildTnaFromLiveRows() {
  const selectedYear = getTnaSelectedYear();
  const filteredRows = Array.isArray(TNA_LIVE_ROWS) && TNA_LIVE_ROWS.length
    ? TNA_LIVE_ROWS.filter(r => Number(r.yr ?? r['년'] ?? 0) === selectedYear)
    : [];

  if (filteredRows.length) {
    const next = buildTnaFromRows(filteredRows);
    TNA.stores = next.stores;
    TNA.monthly = next.monthly;
  }

  initTnaSelects();
}

async function loadTnaFromGoogleSheets(forceFresh = false) {
  if (!TIRE_DATA_URL) return false;
  try {
    const rows = await fetchJsonWithCache(TIRE_DATA_URL, 'tna_tire_rows_v2', 5 * 60 * 1000, forceFresh);
    if (!Array.isArray(rows)) {
      throw new Error(rows && rows.error ? rows.error : '타이어앤 데이터 형식 오류');
    }

    TNA_LIVE_ROWS = rows;
    updateTnaYearSelectFromRows(rows);
    rebuildTnaFromLiveRows();

    console.log('타이어앤 데이터 연동 완료:', TNA.stores.length, '개 매장');
    return true;
  } catch(e) {
    console.warn('타이어앤 구글 시트 연동 실패, 내장 데이터 사용:', e);
    return false;
  }
}

async function refreshTnaAndRender() {
  const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim().includes('최신 조회'));
  const originalText = btn ? btn.innerHTML : '';
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '불러오는 중...';
  }

  await loadTnaFromGoogleSheets(true);
  initTnaSelects();
  renderTna();

  if (btn) {
    btn.disabled = false;
    btn.innerHTML = originalText || '최신 조회';
  }
}


function getTnaStores() {
  const sel = document.getElementById('tna-sel-store').value;
  return sel === 'ALL' ? TNA.stores : TNA.stores.filter(d => d.store === sel);
}

function renderTna() {
  const stores = getTnaStores();
  const MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];
  const MO_LBL = MONTHS.map(m => m+'월');
  const selectedYear = getTnaSelectedYear();
  const tableTitle = document.getElementById('tna-table-title');
  if (tableTitle) tableTitle.textContent = selectedYear + '년 타이어앤 판매 현황';

  // ── KPI 카드 ──
  const totalQty = stores.reduce((s,d)=>s+d.total_qty,0);
  const totalAmt = stores.reduce((s,d)=>s+d.total_amt,0);
  const amtMan = Math.round(totalAmt/10000);
  const activeMonths = MONTHS.filter(m => stores.some(d=>d.monthly[m]?.qty>0));
  const avgMonthQty = activeMonths.length ? Math.round(totalQty / activeMonths.length) : 0;
  const activeStores = stores.filter(d => d.total_qty > 0);
  const avgStoreQty = activeStores.length ? Math.round(totalQty / activeStores.length) : 0;
  const topStore = [...stores].sort((a,b)=>b.total_qty-a.total_qty)[0];
  const minQty = stores.length ? Math.min(...stores.map(d=>d.total_qty)) : 0;
  const bottomStores = stores.filter(d => d.total_qty === minQty);
  const bottomStore = bottomStores[0];

  document.getElementById('tna-kpi-grid').innerHTML = [
    {icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="3"/></svg>`, label:'총 판매 수량', val: totalQty.toLocaleString()+'개', sub:getTnaSelectedYear()+'년 누적', cls:'conn'},
    {icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>`, label:'총 판매 금액', val: amtMan.toLocaleString()+'만원', sub:getTnaSelectedYear()+'년 누적', cls:'views'},
    {icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`, label:'월 평균 판매', val: avgMonthQty.toLocaleString()+'개', sub:'판매 발생 월 기준 · ' + activeMonths.length + '개월', cls:'resin'},
    {icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M7 14l3-3 3 2 5-6"/></svg>`, label:'매장당 평균', val: avgStoreQty.toLocaleString()+'개', sub:'판매 발생 매장 기준 · ' + activeStores.length + '개점', cls:'resreq'},
    {icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="8 21 12 17 16 21"/><line x1="12" y1="17" x2="12" y2="12"/><path d="M7 4h10v4a5 5 0 0 1-10 0V4z"/><path d="M7 4c0 0-3 0-3 3s3 3 3 3"/><path d="M17 4c0 0 3 0 3 3s-3 3-3 3"/></svg>`, label:'최다 판매 매장', val: topStore?.store||'—', sub: (topStore?.total_qty||0)+'개', cls:'review'},
    {icon:`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="8 3 12 7 16 3"/><line x1="12" y1="7" x2="12" y2="12"/><path d="M7 20h10v-4a5 5 0 0 0-10 0v4z"/></svg>`, label:'최하 판매 매장', val: bottomStore?.store||'—', sub: minQty.toLocaleString()+'개' + (bottomStores.length > 1 ? ' · 외 ' + (bottomStores.length-1) + '개점' : ''), cls:'miss'},
  ].map(c=>`<div class="kpi-card ${c.cls}">
    <div class="kpi-label">${c.icon} ${c.label}</div>
    <div class="kpi-value" style="font-size:24px;margin-bottom:6px">${c.val}</div>
    <div style="font-size:11px;color:var(--text3)">${c.sub}</div>
  </div>`).join('');

  // ── 월별 추이 차트 ──
  const moQty = MONTHS.map(m => stores.reduce((s,d)=>s+(d.monthly[m]?.qty||0),0));
  const moAmt = MONTHS.map(m => Math.round(stores.reduce((s,d)=>s+(d.monthly[m]?.amt||0),0)/10000));

  if (tnaCharts.monthly) tnaCharts.monthly.destroy();
  tnaCharts.monthly = new Chart(document.getElementById('tna-chart-monthly').getContext('2d'), {
    data: {
      labels: MO_LBL,
      datasets: [
        { type:'bar', label:'판매금액(만원)', data:moAmt,
           backgroundColor:'rgba(75,124,243,.7)', borderRadius:4,
           yAxisID:'y1' },
        { type:'line', label:'판매갯수(개)', data:moQty,
           borderColor:'#f5a623', backgroundColor:'rgba(245,166,35,.15)',
           pointBackgroundColor:'#f5a623', pointRadius:5, tension:.3,
           yAxisID:'y2' }
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ labels:{ color:'#8b91a8', font:{family:'Pretendard',size:11}, boxWidth:12, padding:16 }},
        tooltip:{ backgroundColor:'#1e2230', borderColor:'#2a2f42', borderWidth:1, titleColor:'#e8eaf0', bodyColor:'#8b91a8', padding:12 } },
      scales:{
        x:{ ticks:{ color:'#5a6080', font:{size:11} }, grid:{ color:'rgba(42,47,66,.5)' } },
        y1:{ type:'linear', position:'left', ticks:{ color:'#4b7cf3', font:{size:10} }, grid:{ color:'rgba(42,47,66,.4)' },
              title:{ display:true, text:'판매금액(만원)', color:'#4b7cf3', font:{size:10} } },
        y2:{ type:'linear', position:'right', ticks:{ color:'#f5a623', font:{size:10} }, grid:{ drawOnChartArea:false },
              title:{ display:true, text:'판매갯수(개)', color:'#f5a623', font:{size:10} } }
      }
    }
  });

  // ── 매장별 갯수 순위 ──
  const sortedQty = [...stores].filter(d=>d.total_qty>0).sort((a,b)=>b.total_qty-a.total_qty);
  const sortedAmt = [...stores].filter(d=>d.total_amt>0).sort((a,b)=>b.total_amt-a.total_amt);

  if (tnaCharts.qtyRank) tnaCharts.qtyRank.destroy();
  tnaCharts.qtyRank = new Chart(document.getElementById('tna-chart-qty-rank').getContext('2d'), {
    type:'bar',
    data:{
      labels: sortedQty.map(d=>d.store),
      datasets:[{ label:'판매갯수(개)', data: sortedQty.map(d=>d.total_qty),
        backgroundColor:'rgba(75,124,243,.8)', borderRadius:4 }]
    },
    options:{
      responsive:true, maintainAspectRatio:false, indexAxis:'y',
      plugins:{ legend:{display:false},
        tooltip:{ backgroundColor:'#1e2230', borderColor:'#2a2f42', borderWidth:1, titleColor:'#e8eaf0', bodyColor:'#8b91a8', padding:10 },
        datalabels:{display:false}
      },
      scales:{
        x:{ ticks:{ color:'#5a6080',font:{size:10} }, grid:{ color:'rgba(42,47,66,.4)' } },
        y:{ ticks:{ color:'#8b91a8',font:{size:11} }, grid:{ display:false } }
      }
    }
  });

  if (tnaCharts.amtRank) tnaCharts.amtRank.destroy();
  tnaCharts.amtRank = new Chart(document.getElementById('tna-chart-amt-rank').getContext('2d'), {
    type:'bar',
    data:{
      labels: sortedAmt.map(d=>d.store),
      datasets:[{ label:'판매금액(만원)', data: sortedAmt.map(d=>Math.round(d.total_amt/10000)),
        backgroundColor:'rgba(245,166,35,.8)', borderRadius:4 }]
    },
    options:{
      responsive:true, maintainAspectRatio:false, indexAxis:'y',
      plugins:{ legend:{display:false},
        tooltip:{ backgroundColor:'#1e2230', borderColor:'#2a2f42', borderWidth:1, titleColor:'#e8eaf0', bodyColor:'#8b91a8', padding:10 } },
      scales:{
        x:{ ticks:{ color:'#5a6080',font:{size:10} }, grid:{ color:'rgba(42,47,66,.4)' } },
        y:{ ticks:{ color:'#8b91a8',font:{size:11} }, grid:{ display:false } }
      }
    }
  });

  // ── 히트맵 ──
  const maxVal = Math.max(...stores.flatMap(d=>MONTHS.map(m=>d.monthly[m]?.qty||0)));
  const heatRows = stores.filter(d=>d.total_qty>0).map(d => {
    const cells = MONTHS.map(m => {
      const v = d.monthly[m]?.qty||0;
      const intensity = maxVal > 0 ? v/maxVal : 0;
      const alpha = 0.08 + intensity*0.87;
      const bg = v>0 ? `rgba(75,124,243,${alpha.toFixed(2)})` : 'transparent';
      const txtColor = intensity>0.5 ? '#fff' : intensity>0.15 ? '#c8d4ff' : 'var(--text3)';
      return `<td style="text-align:center;font-size:12px;font-weight:${v>0?600:400};
        background:${bg};color:${txtColor};padding:8px 6px;border-bottom:1px solid rgba(42,47,66,.3);white-space:nowrap">
        ${v>0?v:''}</td>`;
    }).join('');
    return `<tr><td style="padding:8px 16px;font-weight:600;color:var(--text);font-size:12px;white-space:nowrap;border-bottom:1px solid rgba(42,47,66,.3)">${d.store}</td>${cells}</tr>`;
  }).join('');

  document.getElementById('tna-heatmap').innerHTML = `
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      <thead><tr>
        <th style="padding:10px 16px;text-align:left;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border)">매장명</th>
        ${MONTHS.map(m=>`<th style="padding:10px 8px;text-align:center;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border);white-space:nowrap">${m}월</th>`).join('')}
      </tr></thead>
      <tbody>${heatRows}</tbody>
    </table>`;

  // ── 상세 테이블 ──
  const thead = document.getElementById('tna-table-head');
  thead.innerHTML = `<th style="text-align:left;padding:10px 14px;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border)">매장명</th>`
    + MONTHS.flatMap(m=>[
      `<th style="padding:10px 8px;text-align:right;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border);white-space:nowrap">${m}월 개수</th>`,
      `<th style="padding:10px 8px;text-align:right;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border);white-space:nowrap">${m}월 금액(원)</th>`
    ]).join('')
    + `<th style="padding:10px 8px;text-align:right;background:rgba(245,166,35,.1);color:var(--accent);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border)">합계개수</th>`
    + `<th style="padding:10px 8px;text-align:right;background:rgba(245,166,35,.1);color:var(--accent);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border)">합계금액(원)</th>`;

  const tbody = document.getElementById('tna-table-body');
  let totQty=0, totAmt=0;
  const moTotQty = {}, moTotAmt = {};
  MONTHS.forEach(m=>{moTotQty[m]=0;moTotAmt[m]=0;});

  const tRows = stores.map(d => {
    totQty += d.total_qty; totAmt += d.total_amt;
    const cells = MONTHS.flatMap(m => {
      const v = d.monthly[m]; const qty=v?.qty||0, amt=v?.amt||0;
      moTotQty[m]+=qty; moTotAmt[m]+=amt;
      return [
        `<td style="padding:9px 8px;text-align:right;border-bottom:1px solid rgba(42,47,66,.4);color:${qty>0?'var(--text)':'var(--text3)'};font-size:12px">${qty>0?qty.toLocaleString():'-'}</td>`,
        `<td style="padding:9px 8px;text-align:right;border-bottom:1px solid rgba(42,47,66,.4);color:${amt>0?'var(--text2)':'var(--text3)'};font-size:12px">${amt>0?amt.toLocaleString():'-'}</td>`
      ];
    }).join('');
    return `<tr>
      <td style="padding:9px 14px;font-weight:600;color:var(--text);font-size:12px;border-bottom:1px solid rgba(42,47,66,.4);white-space:nowrap">${d.store}</td>
      ${cells}
      <td style="padding:9px 8px;text-align:right;border-bottom:1px solid rgba(42,47,66,.4);color:var(--accent);font-weight:700;font-size:12px">${d.total_qty>0?d.total_qty.toLocaleString():'-'}</td>
      <td style="padding:9px 8px;text-align:right;border-bottom:1px solid rgba(42,47,66,.4);color:var(--accent);font-weight:700;font-size:12px">${d.total_amt>0?d.total_amt.toLocaleString():'-'}</td>
    </tr>`;
  }).join('');

  const totCells = MONTHS.flatMap(m=>[
    `<td style="padding:9px 8px;text-align:right;background:var(--surface2);font-weight:700;color:var(--text);font-size:12px">${moTotQty[m]>0?moTotQty[m].toLocaleString():'-'}</td>`,
    `<td style="padding:9px 8px;text-align:right;background:var(--surface2);font-weight:700;color:var(--text);font-size:12px">${moTotAmt[m]>0?moTotAmt[m].toLocaleString():'-'}</td>`
  ]).join('');

  tbody.innerHTML = tRows + `<tr class="tr-total">
    <td style="padding:9px 14px;font-weight:700;color:var(--text);font-size:12px;background:var(--surface2)">합계</td>
    ${totCells}
    <td style="padding:9px 8px;text-align:right;background:rgba(245,166,35,.15);font-weight:800;color:var(--accent);font-size:13px">${totQty.toLocaleString()}</td>
    <td style="padding:9px 8px;text-align:right;background:rgba(245,166,35,.15);font-weight:800;color:var(--accent);font-size:13px">${totAmt.toLocaleString()}</td>
  </tr>`;
}

async function downloadTnaExcel() {
  await ensureXlsx();
  const stores = getTnaStores();
  const MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];
  const rows = stores.map(d => {
    const row = {매장명: d.store};
    MONTHS.forEach(m => {
      row[m+'월_개수'] = d.monthly[m]?.qty||0;
      row[m+'월_금액'] = d.monthly[m]?.amt||0;
    });
    row['합계_개수'] = d.total_qty;
    row['합계_금액'] = d.total_amt;
    return row;
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, '타이어앤_판매집계');
  XLSX.writeFile(wb, `티스테이션_타이어앤_판매현황_${getTnaSelectedYear()}년.xlsx`);
}


// ─── 월간 집계 ────────────────────────────────────────────────────

// 연/월 셀렉트 초기화
function initMonthlySelects() {
  // 연/월 조합 추출
  const ymSet = {};
  RAW.forEach(r => {
    if (!ymSet[r.yr]) ymSet[r.yr] = new Set();
    ymSet[r.yr].add(r.mo);
  });
  const years = Object.keys(ymSet).map(Number).sort((a,b)=>b-a);

  const selYr = document.getElementById('mo-sel-yr');
  const selMo = document.getElementById('mo-sel-mo');
  selYr.innerHTML = '';
  selMo.innerHTML = '';

  years.forEach(yr => {
    const o = document.createElement('option');
    o.value = yr; o.textContent = yr+'년';
    selYr.appendChild(o);
  });

  // 연도 변경 시 월 셀렉트 갱신
  function updateMonths() {
    const yr = Number(selYr.value);
    const months = ymSet[yr] ? [...ymSet[yr]].sort((a,b)=>a-b) : [];
    selMo.innerHTML = '';
    months.forEach(mo => {
      const o = document.createElement('option');
      o.value = mo; o.textContent = mo+'월';
      selMo.appendChild(o);
    });
    // 기본: 가장 최근 월
    if (months.length) selMo.value = months[months.length-1];
  }
  selYr.onchange = updateMonths;
  if (years.length) {
    selYr.value = years[0];
    updateMonths();
  }

  // 매장/팀 셀렉트
  const moSelStore = document.getElementById('mo-sel-store');
  const moSelTeam  = document.getElementById('mo-sel-team');
  moSelStore.innerHTML = '<option value="ALL">전체 매장</option>';
  moSelTeam.innerHTML  = '<option value="ALL">전체 팀</option>';
  STORES.forEach(s => {
    const o = document.createElement('option'); o.value=s; o.textContent=s;
    moSelStore.appendChild(o);
  });
  TEAMS.forEach(t => {
    const o = document.createElement('option'); o.value=t; o.textContent=t;
    moSelTeam.appendChild(o);
  });
}


// 월간 데이터 집계 헬퍼
function getMonthRows(yr, mo, store='ALL', team='ALL') {
  return RAW.filter(r =>
    r.yr === yr && r.mo === mo &&
    (store==='ALL' || r.store===store) &&
    (team==='ALL'  || r.team===team)
  );
}

function sumMonthRows(yr, mo, store, team) {
  return sumRows(getMonthRows(yr, mo, store, team));
}

// 전월 계산
function prevMonth(yr, mo) {
  return mo === 1 ? {yr: yr-1, mo: 12} : {yr, mo: mo-1};
}
// 전년 동월
function prevYear(yr, mo) {
  return {yr: yr-1, mo};
}

// 월간 차트 인스턴스
let moCharts = {};

function renderMonthly() {
  const yr    = Number(document.getElementById('mo-sel-yr').value);
  const mo    = Number(document.getElementById('mo-sel-mo').value);
  const store = document.getElementById('mo-sel-store').value;
  const team  = document.getElementById('mo-sel-team').value;

  const pm = prevMonth(yr, mo);
  const py = prevYear(yr, mo);

  const curSum    = sumMonthRows(yr,     mo,     store, team);
  const prevMoSum = sumMonthRows(pm.yr,  pm.mo,  store, team);
  const prevYrSum = sumMonthRows(py.yr,  py.mo,  store, team);

  // ── 기간 배지 ──
  function periodLabel(y, m) {
    const rows = getMonthRows(y, m, 'ALL', 'ALL');
    if (!rows.length) return null;
    const starts = rows.map(r=>r.start).filter(Boolean).sort();
    const ends   = rows.map(r=>r.end).filter(Boolean).sort();
    return (starts[0] && ends[ends.length-1])
      ? starts[0] + ' ~ ' + ends[ends.length-1]
      : y + '년 ' + m + '월';
  }

  const badges = document.getElementById('mo-period-badges');
  const badgeData = [
    {label:'당 월', y:yr,    m:mo,    cls:'rgba(245,166,35,.15)', tc:'var(--accent)'},
    {label:'전 월', y:pm.yr, m:pm.mo, cls:'rgba(75,124,243,.1)',  tc:'var(--blue)'},
    {label:'전 년', y:py.yr, m:py.mo, cls:'rgba(155,108,243,.1)', tc:'var(--purple)'},
  ];
  badges.innerHTML = badgeData.map(function(b) {
    const lbl = periodLabel(b.y, b.m);
    if (!lbl) return '';
    return '<div style="display:flex;align-items:center;gap:6px;padding:6px 12px;border-radius:20px;background:' + b.cls + ';border:1px solid ' + b.tc + '33;font-size:11px">'
      + '<span style="font-weight:700;color:' + b.tc + '">' + b.label + '</span>'
      + '<span style="color:var(--text2)">' + b.y + '년 ' + b.m + '월 · ' + lbl + '</span>'
      + '</div>';
  }).join('');

  document.getElementById('mo-table-title').textContent =
    yr + '년 ' + mo + '월 / ' + pm.yr + '년 ' + pm.mo + '월 / ' + py.yr + '년 ' + py.mo + '월 비교';

  // ── 증감률 계산 헬퍼 ──
  function calcRate(v, ref) {
    if (!ref || ref === 0) return null;
    return (v - ref) / ref * 100;
  }
  function rateHtml(v, ref, higher, label) {
    const d = calcRate(v, ref);
    if (d === null) return '<span class="rate-item"><span class="rl">' + label + '</span><span class="rv rate-flat">—</span></span>';
    const good = higher ? d >= 0 : d <= 0;
    const cls  = good ? 'rate-up' : 'rate-dn';
    const arrow = d > 0 ? '▲' : d < 0 ? '▼' : '';
    return '<span class="rate-item"><span class="rl">' + label + '</span><span class="rv ' + cls + '">' + arrow + Math.abs(d).toFixed(1) + '%</span></span>';
  }
  function msgHtml(key, v, ref, higher) {
    const d = calcRate(v, ref);
    if (d === null) return '';
    const good  = higher ? d >= 0 : d <= 0;
    const color = good ? 'var(--green)' : 'var(--red)';
    const arrow = d > 0 ? '▲' : d < 0 ? '▼' : '';
    const msgs  = {
      conn:    d > 0 ? '연결콜 증가' : '연결콜 감소',
      miss:    d < 0 ? '미연결 개선' : '미연결 증가',
      views:   d > 0 ? '조회수 증가' : '조회수 감소',
      res_req: d > 0 ? '예약신청 증가' : '예약신청 감소',
      review:  d > 0 ? '리뷰 증가' : '리뷰 감소',
    };
    const msg = msgs[key];
    if (!msg) return '';
    return '<div style="font-size:10px;font-weight:700;color:' + color + ';margin-top:6px">전월 대비 ' + msg + ' ' + arrow + Math.abs(d).toFixed(1) + '%</div>';
  }

  // ── KPI 섹션: 그룹별 렌더 ──
  const section = document.getElementById('mo-kpi-section');
  let kpiHtml = '<div>';

  KPI_GROUPS.forEach(function(grp) {
    const grpDefs = KPI_DEFS.filter(function(d){ return d.group === grp.id; });
    const curG  = grpDefs.reduce(function(s,d){ return s+(curSum[d.key]||0); }, 0);
    const prevG = grpDefs.reduce(function(s,d){ return s+(prevMoSum[d.key]||0); }, 0);
    const grpRate = prevG ? ((curG - prevG) / prevG * 100) : null;

    kpiHtml += '<div style="margin-bottom:24px">';
    kpiHtml += '<div class="kpi-group-header">';
    kpiHtml += '<span class="kpi-group-badge ' + grp.id + '">' + grp.icon + ' ' + grp.label + '</span>';
    kpiHtml += '<span style="font-size:11px;color:var(--text3)">' + grp.desc + '</span>';
    kpiHtml += '<div class="kpi-group-divider"></div>';
    if (grpRate !== null) {
      const gc = grpRate >= 0 ? 'var(--green)' : 'var(--red)';
      const ga = grpRate >= 0 ? '▲' : '▼';
      kpiHtml += '<div class="kpi-group-score">전월 대비 <span class="score-num" style="color:' + gc + '">' + ga + Math.abs(grpRate).toFixed(1) + '%</span></div>';
    }
    kpiHtml += '</div>';
    kpiHtml += '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:18px">';

    grpDefs.forEach(function(def) {
      const cv = curSum[def.key]||0, pv = prevMoSum[def.key]||0, yv = prevYrSum[def.key]||0;
      const higher = def.higher;

      function deltaHtml(cur, ref, label) {
        const d = calcRate(cur, ref);
        if (d === null) return '<div style="margin-top:5px"><span style="font-size:11px;color:var(--text3)">—</span></div>';
        const good = higher ? d >= 0 : d <= 0;
        const cls  = good ? 'delta-up' : 'delta-dn';
        const arrow = d > 0 ? '▲' : d < 0 ? '▼' : '';
        return '<div style="margin-top:5px"><span class="compare-delta ' + cls + '">' + arrow + ' ' + Math.abs(d).toFixed(1) + '%</span></div>';
      }

      kpiHtml += '<div class="kpi-card ' + def.cls + '" style="padding:18px">';
      kpiHtml += '<div class="kpi-label" style="margin-bottom:14px">' + def.label + '</div>';
      kpiHtml += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">';
      // 당월
      kpiHtml += '<div style="background:rgba(245,166,35,.08);border:1px solid rgba(245,166,35,.2);border-radius:8px;padding:12px;text-align:center">';
      kpiHtml += '<div style="font-size:9px;font-weight:700;color:var(--accent);letter-spacing:.05em;margin-bottom:6px">당 월</div>';
      kpiHtml += '<div style="font-size:22px;font-weight:800;color:var(--text);line-height:1">' + cv.toLocaleString() + '</div>';
      kpiHtml += '<div style="font-size:10px;color:var(--text3);margin-top:4px">' + yr + '년 ' + mo + '월</div>';
      kpiHtml += '</div>';
      // 전월
      kpiHtml += '<div style="background:rgba(75,124,243,.08);border:1px solid rgba(75,124,243,.2);border-radius:8px;padding:12px;text-align:center">';
      kpiHtml += '<div style="font-size:9px;font-weight:700;color:var(--blue);letter-spacing:.05em;margin-bottom:4px">전 월</div>';
      kpiHtml += '<div style="font-size:18px;font-weight:700;color:var(--text2);line-height:1">' + pv.toLocaleString() + '</div>';
      kpiHtml += '<div style="font-size:10px;color:var(--text3);margin-top:4px">' + pm.yr + '년 ' + pm.mo + '월</div>';
      kpiHtml += deltaHtml(cv, pv, '전월');
      kpiHtml += '</div>';
      // 전년
      kpiHtml += '<div style="background:rgba(155,108,243,.08);border:1px solid rgba(155,108,243,.2);border-radius:8px;padding:12px;text-align:center">';
      kpiHtml += '<div style="font-size:9px;font-weight:700;color:var(--purple);letter-spacing:.05em;margin-bottom:4px">전 년</div>';
      kpiHtml += '<div style="font-size:18px;font-weight:700;color:var(--text2);line-height:1">' + yv.toLocaleString() + '</div>';
      kpiHtml += '<div style="font-size:10px;color:var(--text3);margin-top:4px">' + py.yr + '년 ' + py.mo + '월</div>';
      kpiHtml += deltaHtml(cv, yv, '전년');
      kpiHtml += '</div>';
      kpiHtml += '</div>'; // grid
      kpiHtml += msgHtml(def.key, cv, pv, higher);
      kpiHtml += '</div>'; // kpi-card
    });

    kpiHtml += '</div></div>'; // grid + group
  });
  kpiHtml += '</div>';
  section.innerHTML = kpiHtml;

  // ── 차트 ──
  const labs = [yr+'년 '+mo+'월', pm.yr+'년 '+pm.mo+'월', py.yr+'년 '+py.mo+'월'];
  const C    = ['rgba(245,166,35,.85)','rgba(75,124,243,.75)','rgba(155,108,243,.75)'];

  function makeMonthChart(id, datasets) {
    if (moCharts[id]) moCharts[id].destroy();
    moCharts[id] = new Chart(document.getElementById(id).getContext('2d'), {
      type: 'bar',
      data: { labels: labs, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        barPercentage: .6, categoryPercentage: .65,
        plugins: {
          legend: { labels: { color:'#8b91a8', font:{family:'Pretendard',size:11}, boxWidth:12, padding:14 } },
          tooltip: { backgroundColor:'#1e2230', borderColor:'#2a2f42', borderWidth:1, titleColor:'#e8eaf0', bodyColor:'#8b91a8', padding:10 }
        },
        scales: {
          x: { ticks:{color:'#5a6080',font:{size:11}}, grid:{color:'rgba(42,47,66,.4)'} },
          y: { ticks:{color:'#5a6080',font:{size:10}}, grid:{color:'rgba(42,47,66,.4)'} }
        }
      }
    });
  }

  makeMonthChart('mo-chart-call', [
    { label:'연결콜',   data:[curSum.conn, prevMoSum.conn, prevYrSum.conn],
      backgroundColor: C },
    { label:'미연결콜', data:[curSum.miss, prevMoSum.miss, prevYrSum.miss],
      backgroundColor: ['rgba(242,92,110,.8)','rgba(242,92,110,.55)','rgba(242,92,110,.4)'] },
  ]);
  makeMonthChart('mo-chart-views', [
    { label:'조회수', data:[curSum.views, prevMoSum.views, prevYrSum.views],
      backgroundColor: C },
  ]);
  makeMonthChart('mo-chart-res', [
    { label:'예약유입', data:[curSum.res_in,  prevMoSum.res_in,  prevYrSum.res_in],
      backgroundColor: ['rgba(52,196,138,.85)','rgba(52,196,138,.6)','rgba(52,196,138,.4)'] },
    { label:'예약신청', data:[curSum.res_req, prevMoSum.res_req, prevYrSum.res_req],
      backgroundColor: ['rgba(0,180,216,.85)','rgba(0,180,216,.6)','rgba(0,180,216,.4)'] },
  ]);
  makeMonthChart('mo-chart-etc', [
    { label:'리뷰',     data:[curSum.review, prevMoSum.review, prevYrSum.review],
      backgroundColor: ['rgba(249,199,79,.85)','rgba(249,199,79,.6)','rgba(249,199,79,.4)'] },
    { label:'톡톡상담', data:[curSum.chat,   prevMoSum.chat,   prevYrSum.chat],
      backgroundColor: ['rgba(155,108,243,.85)','rgba(155,108,243,.6)','rgba(155,108,243,.4)'] },
  ]);

  // ── 매장별 비교 테이블 ──
  const allStoreRows = (store === 'ALL') ? STORES : STORES.filter(function(s){ return s === store; });
  const METRICS_KEYS = ['views','conn','miss','res_in','res_req','review','chat'];
  const PERIOD_SETS  = [
    {label:'당 월', y:yr,    m:mo,    tc:'var(--accent)',  bg:'rgba(245,166,35,.06)'},
    {label:'전 월', y:pm.yr, m:pm.mo, tc:'var(--blue)',    bg:'rgba(75,124,243,.04)'},
    {label:'전 년', y:py.yr, m:py.mo, tc:'var(--purple)',  bg:'rgba(155,108,243,.04)'},
  ];

  const thead = document.getElementById('mo-table-head');
  thead.innerHTML = '<th style="text-align:left;padding:10px 14px;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border);white-space:nowrap">매장명</th>'
    + '<th style="text-align:left;padding:10px 8px;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border)">팀</th>'
    + '<th style="padding:10px 8px;text-align:center;background:var(--surface2);color:var(--text3);font-size:10px;border-bottom:1px solid var(--border)">구분</th>'
    + ['조회수','연결콜','미연결콜','예약유입','예약신청','리뷰','톡톡상담'].map(function(h){
        return '<th style="padding:10px 8px;text-align:right;background:var(--surface2);color:var(--text3);font-size:10px;letter-spacing:.05em;border-bottom:1px solid var(--border);white-space:nowrap">'+h+'</th>';
      }).join('');

  const tbody = document.getElementById('mo-table-body');
  let tableHtml = '';

  allStoreRows.forEach(function(s) {
    const storeTeam = (RAW.find(function(r){ return r.store===s; }) || {}).team || '';
    if (team !== 'ALL' && storeTeam !== team) return;

    PERIOD_SETS.forEach(function(ps, pi) {
      const storeRows = getMonthRows(ps.y, ps.m, s, 'ALL');
      const d = sumRows(storeRows);
      const bd = pi===2 ? '.5' : '.2';
      tableHtml += '<tr style="background:' + ps.bg + '">';
      if (pi === 0) {
        tableHtml += '<td style="padding:8px 14px;font-weight:700;color:var(--text);font-size:12px;border-bottom:1px solid rgba(42,47,66,.3);white-space:nowrap;vertical-align:middle" rowspan="3">' + s + '</td>';
        tableHtml += '<td style="padding:8px 8px;font-size:10px;color:var(--text3);border-bottom:1px solid rgba(42,47,66,.3);vertical-align:middle" rowspan="3">' + storeTeam + '</td>';
      }
      tableHtml += '<td style="padding:7px 8px;text-align:center;border-bottom:1px solid rgba(42,47,66,.' + bd + ')"><span style="font-size:10px;font-weight:700;color:' + ps.tc + ';background:' + ps.bg + ';border:1px solid ' + ps.tc + '33;padding:2px 8px;border-radius:10px">' + ps.label + '</span></td>';
      METRICS_KEYS.forEach(function(k) {
        const v = d[k];
        tableHtml += '<td style="padding:7px 8px;text-align:right;border-bottom:1px solid rgba(42,47,66,.' + bd + ');color:' + (v>0?(pi===0?'var(--text)':'var(--text2)'):'var(--text3)') + ';font-size:12px;font-weight:' + (pi===0&&v>0?600:400) + '">' + (v>0?v.toLocaleString():'—') + '</td>';
      });
      tableHtml += '</tr>';
    });
  });

  // 합계 행
  PERIOD_SETS.forEach(function(ps, pi) {
    const d = sumMonthRows(ps.y, ps.m, store, team);
    const bt = pi===0 ? '2px' : '1px';
    tableHtml += '<tr style="background:var(--surface2)">';
    if (pi === 0) {
      tableHtml += '<td style="padding:10px 14px;font-weight:800;color:var(--text);font-size:12px;border-top:2px solid var(--border)" rowspan="3">합 계</td>';
      tableHtml += '<td style="font-size:10px;color:var(--text3);padding:10px 8px;border-top:2px solid var(--border)" rowspan="3">—</td>';
    }
    tableHtml += '<td style="padding:8px;text-align:center;border-top:' + bt + ' solid var(--border)"><span style="font-size:10px;font-weight:700;color:' + ps.tc + ';padding:2px 8px;border-radius:10px;background:' + ps.bg + ';border:1px solid ' + ps.tc + '33">' + ps.label + '</span></td>';
    METRICS_KEYS.forEach(function(k) {
      tableHtml += '<td style="padding:8px;text-align:right;border-top:' + bt + ' solid var(--border);font-weight:' + (pi===0?800:600) + ';color:' + (pi===0?'var(--accent)':'var(--text2)') + ';font-size:12px">' + d[k].toLocaleString() + '</td>';
    });
    tableHtml += '</tr>';
  });

  tbody.innerHTML = tableHtml;
}
async function downloadMonthlyExcel() {
  await ensureXlsx();
  const yr    = Number(document.getElementById('mo-sel-yr').value);
  const mo    = Number(document.getElementById('mo-sel-mo').value);
  const store = document.getElementById('mo-sel-store').value;
  const team  = document.getElementById('mo-sel-team').value;
  const pm = prevMonth(yr, mo), py = prevYear(yr, mo);

  const periods = [
    {label:'당월', y:yr, m:mo},
    {label:'전월', y:pm.yr, m:pm.mo},
    {label:'전년', y:py.yr, m:py.mo},
  ];

  const allData = [];
  periods.forEach(ps => {
    const rows = getMonthRows(ps.y, ps.m, store, team);
    rows.forEach(r => allData.push({
      구분: ps.label, 연도: ps.y, 월: ps.m, 주차: r.wk,
      기간: `${r.start}~${r.end}`,
      매장명: r.store, 팀: r.team,
      조회수: r.views, 연결콜: r.conn, 미연결콜: r.miss,
      예약유입: r.res_in, 예약신청: r.res_req, 리뷰: r.review, 톡톡상담: r.chat,
    }));
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(allData);
  ws['!cols'] = Array(14).fill({wch:12});
  XLSX.utils.book_append_sheet(wb, ws, '월간집계');
  XLSX.writeFile(wb, `티스테이션_월간집계_${yr}년${mo}월.xlsx`);
}

// ─── 초기화 ──────────────────────────────────────────────────────
async function init() {
  await Promise.all([
    loadFromGoogleSheets(),
    loadTnaFromGoogleSheets(),
  ]);
  rebuildMetaFromRaw();
  initSelects();
  initMonthlySelects();
  initTnaSelects();
  applyFilter();
}

init();
function openKpiModal(data) {
  document.getElementById('kpiModalTitle').textContent = data.title || '-';
  document.getElementById('kpiModalValue').textContent = data.current || '-';
  document.getElementById('kpiModalPrevMonth').textContent = data.prevMonth || '-';
  document.getElementById('kpiModalPrevYear').textContent = data.prevYear || '-';
  document.getElementById('kpiModalDeltaMonth').textContent = data.deltaMonth || '-';
  document.getElementById('kpiModalDeltaYear').textContent = data.deltaYear || '-';

  document.getElementById('kpiModal').style.display = 'block';
  document.body.style.overflow = 'hidden';
}

function closeKpiModal() {
  const modal = document.getElementById('kpiModal');
  if (!modal) return;

  modal.style.display = 'none';
  document.body.style.overflow = '';
}

function bindKpiCardEvents() {
  const cards = document.querySelectorAll('#kpi-grid .kpi-card');

  cards.forEach(card => {
    card.onclick = function () {
      const title = card.querySelector('.kpi-label')?.textContent?.trim() || '-';
      const current = card.querySelector('.kpi-value')?.textContent?.trim() || '-';

      const rows = [...card.querySelectorAll('.kpi-compare-row')];

      const prevMonth = rows[0]?.querySelector('.compare-val')?.textContent?.trim() || '-';
      const prevYear = rows[1]?.querySelector('.compare-val')?.textContent?.trim() || '-';

      const rates = [...card.querySelectorAll('.rate-item .rv')];

      const deltaMonth = rates[0]?.textContent?.trim() || '-';
      const deltaYear = rates[1]?.textContent?.trim() || '-';

      openKpiModal({
        title,
        current,
        prevMonth,
        prevYear,
        deltaMonth,
        deltaYear
      });
    };
  });
}

document.addEventListener('keydown', function (e) {
  if (e.key === 'Escape') {
    closeKpiModal();
  }
});
