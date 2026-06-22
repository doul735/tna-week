// ─────────────────────────────────────────────────────────────────────────────
// TNA 대시보드 app.js 전체 교체용
// 사용 방법: GitHub의 기존 app.js 내용을 전부 삭제하고, 이 코드 전체를 붙여넣으세요.
// ─────────────────────────────────────────────────────────────────────────────

const DATA_URL = 'https://script.google.com/macros/s/AKfycbw_TCR45muWiseITDdxHo_sYPKYxLS5CgRi_1LCouEgrapDkMQ7VE-HAj8zURoI2Uc/exec';

const RAW = [];

let weekMeta = {};
let weekKeys = [];
let monthKeys = [];
let STORES = [];
let TEAMS = [];

let charts = {};
let lastRenderData = null;

const METRICS = ['views', 'conn', 'miss', 'res_in', 'res_req', 'review', 'neg_review', 'chat'];

const METRIC_DEFS = {
  views: {
    label: '플레이스 조회수',
    group: '노출 관련 지표',
    groupKey: 'exposure',
    chip: '노출',
    icon: '◎',
    higher: true,
    desc: '네이버 플레이스에서 매장이 얼마나 많이 조회되었는지 확인하는 핵심 노출 지표입니다.'
  },
  conn: {
    label: '연결콜',
    group: '전환 관련 지표',
    groupKey: 'convert',
    chip: '전환',
    icon: '☎',
    higher: true,
    desc: '고객이 실제 전화 연결까지 이어진 수치입니다. 문의 전환 흐름을 확인할 수 있습니다.'
  },
  miss: {
    label: '미연결콜',
    group: '관리 관련 지표',
    groupKey: 'manage',
    chip: '관리',
    icon: '×',
    higher: false,
    desc: '연결되지 못한 전화입니다. 낮아질수록 응대 관리가 개선된 것으로 볼 수 있습니다.'
  },
  res_in: {
    label: '예약유입',
    group: '노출 관련 지표',
    groupKey: 'exposure',
    chip: '노출',
    icon: '↗',
    higher: true,
    desc: '예약 화면까지 유입된 고객 수입니다. 관심 고객의 행동 흐름을 볼 수 있습니다.'
  },
  res_req: {
    label: '예약신청',
    group: '전환 관련 지표',
    groupKey: 'convert',
    chip: '전환',
    icon: '▣',
    higher: true,
    desc: '실제 예약 신청까지 이어진 수치입니다. 예약 전환 성과를 보는 핵심 지표입니다.'
  },
  review: {
    label: '리뷰',
    group: '관리 관련 지표',
    groupKey: 'manage',
    chip: '관리',
    icon: '☆',
    higher: true,
    desc: '고객 리뷰 발생 수입니다. 매장 신뢰도와 운영 관리 상태를 함께 보여주는 지표입니다.'
  },
  neg_review: {
    label: '부정리뷰',
    group: '관리 관련 지표',
    groupKey: 'manage',
    chip: '관리',
    icon: '!',
    higher: false,
    desc: '부정적인 리뷰 또는 관리가 필요한 리뷰 수치입니다. 낮아질수록 고객 응대와 리뷰 관리 측면에서 긍정적으로 볼 수 있습니다.'
  },
  chat: {
    label: '톡톡상담',
    group: '관리 관련 지표',
    groupKey: 'manage',
    chip: '관리',
    icon: '●',
    higher: true,
    desc: '네이버 톡톡 상담 수치입니다. 전화 외 문의 채널의 반응을 확인할 수 있습니다.'
  }
};

// ─────────────────────────────────────────────────────────────────────────────
// 기본 유틸
// ─────────────────────────────────────────────────────────────────────────────

function $(id) {
  return document.getElementById(id);
}

function text(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function num(value) {
  if (value === null || value === undefined || value === '') return 0;
  return Number(String(value).replace(/,/g, '').trim()) || 0;
}

function fmt(value) {
  return num(value).toLocaleString();
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function formatDateLocal(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function dateText(value) {
  if (!value) return '';

  if (typeof value === 'number') {
    const base = new Date(1899, 11, 30);
    base.setDate(base.getDate() + value);
    return formatDateLocal(base);
  }

  const str = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.slice(0, 10);
  }

  if (/^\d{4}\.\d{1,2}\.\d{1,2}/.test(str)) {
    const parts = str.split(/[.\s]/).filter(Boolean);
    return `${parts[0]}-${pad2(parts[1])}-${pad2(parts[2])}`;
  }

  if (/^\d{4}\/\d{1,2}\/\d{1,2}/.test(str)) {
    const parts = str.split(/[\/\s]/).filter(Boolean);
    return `${parts[0]}-${pad2(parts[1])}-${pad2(parts[2])}`;
  }

  return str.split('T')[0];
}

function setTextIfExists(id, value) {
  const el = $(id);
  if (el) el.textContent = value;
}

function setHtmlIfExists(id, value) {
  const el = $(id);
  if (el) el.innerHTML = value;
}

function cacheBustUrl(url) {
  return url + (url.includes('?') ? '&' : '?') + 'v=' + Date.now();
}

// ─────────────────────────────────────────────────────────────────────────────
// CSS 보정
// styles.css를 부분 수정하지 않아도 모달/카드 기본 스타일이 적용되도록 JS에서 보정합니다.
// ─────────────────────────────────────────────────────────────────────────────

function injectRuntimeCss() {
  if ($('tna-runtime-fixes')) return;

  const style = document.createElement('style');
  style.id = 'tna-runtime-fixes';
  style.textContent = `
    #kpiModal[hidden],
    .kpi-modal[hidden] {
      display: none !important;
    }

    .kpi-modal {
      position: fixed;
      inset: 0;
      z-index: 9999;
      display: none;
      place-items: center;
      padding: 24px;
      background: rgba(0, 8, 14, .72);
      backdrop-filter: blur(12px);
    }

    .kpi-modal.is-open {
      display: grid !important;
    }

    .kpi-modal-panel {
      width: min(560px, 92vw);
      border: 1px solid rgba(255, 122, 0, .45);
      border-radius: 26px;
      background:
        radial-gradient(circle at top right, rgba(255, 122, 0, .18), transparent 36%),
        linear-gradient(135deg, rgba(15, 32, 46, .98), rgba(14, 21, 30, .98));
      box-shadow: 0 28px 90px rgba(0, 0, 0, .55);
      padding: 30px;
      color: #f8fafc;
    }

    .kpi-modal-head {
      display: flex;
      justify-content: space-between;
      gap: 20px;
      align-items: flex-start;
      margin-bottom: 22px;
    }

    .kpi-modal-group {
      color: #ff8a1c;
      font-size: 14px;
      font-weight: 800;
      margin-bottom: 8px;
    }

    .kpi-modal-title {
      font-size: 28px;
      font-weight: 900;
      letter-spacing: -0.04em;
      margin: 0;
    }

    .kpi-modal-close {
      width: 42px;
      height: 42px;
      border-radius: 50%;
      border: 1px solid rgba(255,255,255,.25);
      background: rgba(255,255,255,.08);
      color: #fff;
      font-size: 26px;
      line-height: 1;
      cursor: pointer;
    }

    .kpi-modal-value {
      font-size: 54px;
      font-weight: 900;
      letter-spacing: -0.05em;
      margin: 12px 0 24px;
    }

    .kpi-modal-compare {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 14px;
      margin-bottom: 20px;
    }

    .kpi-modal-box {
      border: 1px solid rgba(148, 163, 184, .2);
      border-radius: 18px;
      background: rgba(15, 23, 42, .52);
      padding: 18px;
    }

    .kpi-modal-box-label {
      display: block;
      color: #93a4b8;
      font-size: 13px;
      font-weight: 800;
      margin-bottom: 8px;
    }

    .kpi-modal-box-value {
      display: block;
      font-size: 24px;
      font-weight: 900;
      margin-bottom: 6px;
    }

    .kpi-modal-desc {
      color: #cbd5e1;
      line-height: 1.65;
      margin: 0;
    }

    #kpi-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 16px;
    }

    .kpi-card {
      position: relative;
      min-height: 154px;
      border: 1px solid rgba(148, 163, 184, .22);
      border-radius: 20px;
      padding: 20px;
      background:
        radial-gradient(circle at top left, rgba(255, 122, 0, .14), transparent 34%),
        linear-gradient(135deg, rgba(18, 34, 48, .96), rgba(13, 20, 29, .96));
      box-shadow: inset 0 1px 0 rgba(255,255,255,.05);
      cursor: pointer;
      transition: transform .18s ease, border-color .18s ease, box-shadow .18s ease;
      overflow: hidden;
    }

    .kpi-card:hover {
      transform: translateY(-2px);
      border-color: rgba(255, 122, 0, .45);
      box-shadow: 0 18px 40px rgba(0,0,0,.24);
    }

    .kpi-top {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 18px;
    }

    .kpi-head-left {
      display: flex;
      align-items: center;
      gap: 12px;
      min-width: 0;
    }

    .kpi-icon {
      width: 40px;
      height: 40px;
      flex: 0 0 auto;
      display: grid;
      place-items: center;
      border-radius: 999px;
      border: 1px solid rgba(255, 122, 0, .36);
      background: rgba(255, 122, 0, .12);
      color: #ff8a1c;
      font-weight: 900;
    }

    .kpi-title {
      margin: 0 0 4px;
      font-size: 17px;
      font-weight: 900;
      color: #f8fafc;
      letter-spacing: -0.03em;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .kpi-group {
      color: #9fb0c2;
      font-size: 12px;
      font-weight: 700;
    }

    .kpi-chip {
      flex: 0 0 auto;
      border: 1px solid rgba(148, 163, 184, .18);
      border-radius: 999px;
      padding: 5px 10px;
      color: #cbd5e1;
      font-size: 12px;
      font-weight: 800;
      background: rgba(2, 6, 23, .24);
    }

    .kpi-value {
      font-size: 38px;
      line-height: 1;
      font-weight: 900;
      letter-spacing: -0.06em;
      color: #f8fafc;
      margin-bottom: 18px;
    }

    .kpi-rates {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
    }

    .kpi-rate-box {
      border: 1px solid rgba(148, 163, 184, .14);
      border-radius: 14px;
      background: rgba(2, 6, 23, .28);
      padding: 10px;
    }

    .rate-label {
      display: block;
      margin-bottom: 4px;
      color: #a8b6c8;
      font-size: 12px;
      font-weight: 800;
    }

    .rate-up,
    .rate-good {
      color: #32d583;
      font-weight: 900;
    }

    .rate-down,
    .rate-bad {
      color: #fb7185;
      font-weight: 900;
    }

    .rate-flat {
      color: #94a3b8;
      font-weight: 900;
    }

    .empty-state {
      margin-top: 18px;
      border: 1px solid rgba(251, 113, 133, .28);
      border-radius: 18px;
      background: rgba(251, 113, 133, .08);
      padding: 18px;
      color: #fecdd3;
    }

    @media (max-width: 1080px) {
      #kpi-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
    }

    @media (max-width: 640px) {
      #kpi-grid,
      .kpi-modal-compare {
        grid-template-columns: 1fr;
      }

      .kpi-modal-panel {
        padding: 22px;
      }

      .kpi-modal-value {
        font-size: 44px;
      }
    }
  `;

  document.head.appendChild(style);
}

// ─────────────────────────────────────────────────────────────────────────────
// 데이터 로딩 / 정규화
// ─────────────────────────────────────────────────────────────────────────────

async function fetchJsonWithCache(url, cacheKey, ttlMs = 5 * 60 * 1000, forceFresh = false) {
  const now = Date.now();

  if (!forceFresh) {
    try {
      const cached = localStorage.getItem(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (parsed && parsed.ts && now - parsed.ts < ttlMs) {
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

function normalizePlaceRow(r) {
  return {
    yr: num(r['년'] ?? r.yr ?? r.year),
    mo: num(r['월'] ?? r.mo ?? r.month),
    wk: num(r['주'] ?? r['주차'] ?? r.wk ?? r.week),
    start: dateText(r['시작일(월)'] ?? r['시작일'] ?? r.start ?? r.start_date),
    end: dateText(r['종료일(일)'] ?? r['종료일'] ?? r.end ?? r.end_date),
    store: text(r['매장명'] ?? r.store ?? r.store_name),
    team: text(r['팀'] ?? r.team),
    views: num(r['조회수'] ?? r.views),
    conn: num(r['연결콜'] ?? r.conn),
    miss: num(r['미연결콜'] ?? r.miss),
    res_in: num(r['예약유입'] ?? r.res_in ?? r.resIn),
    res_req: num(r['예약신청'] ?? r.res_req ?? r.resReq),
    review: num(r['리뷰'] ?? r.review),
    neg_review: num(r['부정리뷰'] ?? r.neg_review ?? r.negative_review ?? r.bad_review),
    chat: num(r['톡톡상담'] ?? r.chat)
  };
}

function normalizeRows(data) {
  const rows =
    Array.isArray(data) ? data :
    Array.isArray(data?.data) ? data.data :
    Array.isArray(data?.rows) ? data.rows :
    [];

  return rows
    .map(normalizePlaceRow)
    .filter(r => r.yr && r.mo && r.wk && r.store);
}

async function loadData(forceFresh = false) {
  const data = await fetchJsonWithCache(DATA_URL, 'tna_place_dashboard_rows_v4', 5 * 60 * 1000, forceFresh);
  const rows = normalizeRows(data);

  RAW.length = 0;
  rows.forEach(row => RAW.push(row));

  rebuildMetaFromRaw();
}

// ─────────────────────────────────────────────────────────────────────────────
// 데이터 집계
// ─────────────────────────────────────────────────────────────────────────────

function rebuildMetaFromRaw() {
  weekMeta = {};

  RAW.forEach(r => {
    const key = `${r.yr}-${r.mo}-${r.wk}`;

    if (!weekMeta[key]) {
      weekMeta[key] = {
        key,
        yr: r.yr,
        mo: r.mo,
        wk: r.wk,
        start: r.start,
        end: r.end,
        label: `${r.yr}년 ${r.mo}월 ${r.wk}주차`,
        short: `${r.mo}월 ${r.wk}주`
      };
    }
  });

  weekKeys = Object.keys(weekMeta).sort((a, b) => {
    const ad = weekMeta[a].start ? new Date(weekMeta[a].start) : new Date(weekMeta[a].yr, weekMeta[a].mo - 1, weekMeta[a].wk);
    const bd = weekMeta[b].start ? new Date(weekMeta[b].start) : new Date(weekMeta[b].yr, weekMeta[b].mo - 1, weekMeta[b].wk);
    return ad - bd;
  });

  monthKeys = [...new Set(RAW.map(r => `${r.yr}-${r.mo}`))]
    .sort((a, b) => {
      const [ay, am] = a.split('-').map(Number);
      const [by, bm] = b.split('-').map(Number);
      return ay === by ? am - bm : ay - by;
    });

  STORES = [...new Set(RAW.map(r => r.store).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ko'));
  TEAMS = [...new Set(RAW.map(r => r.team).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ko'));
}

function sumRows(rows) {
  const s = {};
  METRICS.forEach(key => { s[key] = 0; });

  rows.forEach(r => {
    METRICS.forEach(key => {
      s[key] += num(r[key]);
    });
  });

  return s;
}

function hasAnyData(sum) {
  return METRICS.some(key => num(sum[key]) > 0);
}

function getRowsByWeekKey(weekKey, store = 'ALL', team = 'ALL') {
  const meta = weekMeta[weekKey];
  if (!meta) return [];

  return RAW.filter(r =>
    r.yr === meta.yr &&
    r.mo === meta.mo &&
    r.wk === meta.wk &&
    (store === 'ALL' || r.store === store) &&
    (team === 'ALL' || r.team === team)
  );
}

function getRowsByMonthKey(monthKey, store = 'ALL', team = 'ALL') {
  const [yr, mo] = String(monthKey).split('-').map(Number);

  return RAW.filter(r =>
    r.yr === yr &&
    r.mo === mo &&
    (store === 'ALL' || r.store === store) &&
    (team === 'ALL' || r.team === team)
  );
}

function prevMonthKey(yr, mo, wk) {
  let pmo = mo - 1;
  let pyr = yr;

  if (pmo < 1) {
    pmo = 12;
    pyr -= 1;
  }

  const key = `${pyr}-${pmo}-${wk}`;
  return weekMeta[key] ? key : null;
}

function prevYearKey(yr, mo, wk) {
  const key = `${yr - 1}-${mo}-${wk}`;
  return weekMeta[key] ? key : null;
}

function deltaRate(current, previous, higher = true) {
  if (previous === null || previous === undefined || num(previous) === 0) return null;

  const value = ((num(current) - num(previous)) / num(previous)) * 100;

  return {
    value,
    good: higher ? value >= 0 : value <= 0
  };
}

function rateHtml(current, previous, higher, label) {
  const rate = deltaRate(current, previous, higher);

  if (!rate) {
    return `
      <div class="kpi-rate-box">
        <span class="rate-label">${label}</span>
        <span class="rate-flat">—</span>
      </div>
    `;
  }

  const arrow = rate.value > 0 ? '▲' : rate.value < 0 ? '▼' : '–';
  const cls = rate.good ? 'rate-up' : 'rate-down';

  return `
    <div class="kpi-rate-box">
      <span class="rate-label">${label}</span>
      <span class="${cls}">${arrow} ${Math.abs(rate.value).toFixed(1)}%</span>
    </div>
  `;
}

// ─────────────────────────────────────────────────────────────────────────────
// 셀렉트 초기화
// ─────────────────────────────────────────────────────────────────────────────

function fillMetricSelect(selectId) {
  const select = $(selectId);
  if (!select) return;

  const currentValue = select.value || 'views';

  select.innerHTML = METRICS.map(key => {
    const def = METRIC_DEFS[key];
    return `<option value="${key}">${def.label}</option>`;
  }).join('');

  if (METRIC_DEFS[currentValue]) select.value = currentValue;
}

function initSelects() {
  const selStore = $('sel-store');
  const selTeam = $('sel-team');
  const selWeek = $('sel-week');
  const selMonth = $('sel-month');

  if (!selStore || !selTeam || !selWeek || !selMonth) {
    console.warn('필수 선택 박스 ID를 찾지 못했습니다.');
    return;
  }

  selStore.innerHTML = '<option value="ALL">전체 매장</option>';
  STORES.forEach(store => {
    selStore.innerHTML += `<option value="${store}">${store}</option>`;
  });

  selTeam.innerHTML = '<option value="ALL">전체 팀</option>';
  TEAMS.forEach(team => {
    selTeam.innerHTML += `<option value="${team}">${team}</option>`;
  });

  selWeek.innerHTML = '';
  weekKeys.forEach(key => {
    const m = weekMeta[key];
    selWeek.innerHTML += `<option value="${key}">${m.start} ~ ${m.end} · ${m.label}</option>`;
  });

  selMonth.innerHTML = '';
  monthKeys.forEach(key => {
    const [yr, mo] = key.split('-').map(Number);
    selMonth.innerHTML += `<option value="${key}">${yr}년 ${mo}월</option>`;
  });

  fillMetricSelect('trendMetricSelect');
  fillMetricSelect('compareMetricSelect');

  const latestDataWeek = [...weekKeys].reverse().find(key => hasAnyData(sumRows(getRowsByWeekKey(key))));

  if (latestDataWeek) {
    selWeek.value = latestDataWeek;
    const m = weekMeta[latestDataWeek];
    selMonth.value = `${m.yr}-${m.mo}`;
  } else if (weekKeys.length) {
    selWeek.value = weekKeys[weekKeys.length - 1];
    const m = weekMeta[selWeek.value];
    selMonth.value = `${m.yr}-${m.mo}`;
  }

  selStore.onchange = renderDashboard;
  selTeam.onchange = renderDashboard;

  selWeek.onchange = () => {
    const meta = weekMeta[selWeek.value];
    if (meta && [...selMonth.options].some(o => o.value === `${meta.yr}-${meta.mo}`)) {
      selMonth.value = `${meta.yr}-${meta.mo}`;
    }
    renderDashboard();
  };

  selMonth.onchange = () => {
    const [yr, mo] = selMonth.value.split('-').map(Number);

    const matchedWeeks = weekKeys.filter(key => {
      const m = weekMeta[key];
      return m.yr === yr && m.mo === mo;
    });

    const latestWeek = [...matchedWeeks].reverse().find(key => hasAnyData(sumRows(getRowsByWeekKey(key))));

    if (latestWeek) {
      selWeek.value = latestWeek;
    } else if (matchedWeeks.length) {
      selWeek.value = matchedWeeks[matchedWeeks.length - 1];
    }

    renderDashboard();
  };

  const trendMetricSelect = $('trendMetricSelect');
  const compareMetricSelect = $('compareMetricSelect');

  if (trendMetricSelect) trendMetricSelect.onchange = renderDashboard;
  if (compareMetricSelect) compareMetricSelect.onchange = renderDashboard;
}

// ─────────────────────────────────────────────────────────────────────────────
// 화면 렌더링
// ─────────────────────────────────────────────────────────────────────────────

function renderPeriodCards(meta, prevMoMeta, prevYrMeta) {
  setHtmlIfExists('period-current', `${meta.start} ~ ${meta.end}<br>선택 주차 기준`);
  setTextIfExists('header-period', `${meta.start} ~ ${meta.end} · ${meta.mo}월 ${meta.wk}주차`);

  setHtmlIfExists(
    'period-prev-month',
    prevMoMeta ? `${prevMoMeta.start} ~ ${prevMoMeta.end}<br>전월 같은 주차` : '비교 데이터 없음'
  );

  setHtmlIfExists(
    'period-prev-year',
    prevYrMeta ? `${prevYrMeta.start} ~ ${prevYrMeta.end}<br>전년 같은 기간` : '비교 데이터 없음'
  );
}

function renderKpiCards(cur, prevMo, prevYr) {
  const grid = $('kpi-grid');
  if (!grid) return;

  grid.innerHTML = METRICS.map(key => {
    const def = METRIC_DEFS[key];

    return `
      <article class="kpi-card ${def.groupKey}" data-metric="${key}" tabindex="0" role="button" aria-label="${def.label} 상세 보기">
        <div class="kpi-top">
          <div class="kpi-head-left">
            <div class="kpi-icon">${def.icon}</div>
            <div>
              <h3 class="kpi-title">${def.label}</h3>
              <div class="kpi-group">${def.group}</div>
            </div>
          </div>
          <div class="kpi-chip">${def.chip}</div>
        </div>

        <div class="kpi-value">${fmt(cur[key])}</div>

        <div class="kpi-rates">
          ${rateHtml(cur[key], prevMo ? prevMo[key] : null, def.higher, '전월 동기')}
          ${rateHtml(cur[key], prevYr ? prevYr[key] : null, def.higher, '전년 동기')}
        </div>
      </article>
    `;
  }).join('');

  document.querySelectorAll('#kpi-grid .kpi-card').forEach(card => {
    card.addEventListener('click', () => openKpiModal(card.dataset.metric));
    card.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        openKpiModal(card.dataset.metric);
      }
    });
  });
}

function makeChart(id, config) {
  if (charts[id]) charts[id].destroy();

  const canvas = $(id);
  if (!canvas || !window.Chart) return;

  charts[id] = new Chart(canvas.getContext('2d'), config);
}

function chartBaseOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: '#cbd5e1',
          font: { size: 13, family: 'Pretendard' },
          boxWidth: 12
        }
      },
      tooltip: {
        backgroundColor: '#111827',
        borderColor: 'rgba(255,255,255,.12)',
        borderWidth: 1,
        titleColor: '#f8fafc',
        bodyColor: '#cbd5e1',
        padding: 12
      }
    },
    scales: {
      x: {
        ticks: { color: '#cbd5e1', font: { size: 12, family: 'Pretendard' } },
        grid: { color: 'rgba(148, 163, 184, .12)' }
      },
      y: {
        beginAtZero: true,
        ticks: { color: '#cbd5e1', font: { size: 12, family: 'Pretendard' } },
        grid: { color: 'rgba(148, 163, 184, .12)' }
      }
    }
  };
}

function renderCharts(cur, prevMo, prevYr, trendKeys, store, team) {
  if (!window.Chart) return;

  const trendMetric = $('trendMetricSelect')?.value || 'views';
  const compareMetric = $('compareMetricSelect')?.value || 'views';

  const trendLabels = trendKeys.map(key => weekMeta[key].short);
  const trendValues = trendKeys.map(key => sumRows(getRowsByWeekKey(key, store, team))[trendMetric]);
  const reservationValues = trendKeys.map(key => sumRows(getRowsByWeekKey(key, store, team)).res_req);

  makeChart('chart-trend', {
    type: 'line',
    data: {
      labels: trendLabels,
      datasets: [{
        label: METRIC_DEFS[trendMetric]?.label || '조회수',
        data: trendValues,
        borderColor: '#ff7a00',
        backgroundColor: 'rgba(255, 122, 0, .18)',
        fill: true,
        tension: 0.35,
        pointRadius: 5,
        pointHoverRadius: 7,
        pointBackgroundColor: '#ff7a00'
      }]
    },
    options: chartBaseOptions()
  });

  const exposureTotal = cur.views + cur.res_in;
  const convertTotal = cur.conn + cur.res_req;
  const manageTotal = cur.miss + cur.review + cur.neg_review + cur.chat;

  makeChart('chart-groups', {
    type: 'doughnut',
    data: {
      labels: ['노출 관련 지표', '전환 관련 지표', '관리 관련 지표'],
      datasets: [{
        data: hasAnyData(cur) ? [exposureTotal, convertTotal, manageTotal] : [2, 2, 4],
        backgroundColor: ['#ff7a00', '#38a7ff', '#35d07f'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '58%',
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#111827',
          titleColor: '#f8fafc',
          bodyColor: '#cbd5e1',
          padding: 12
        }
      }
    }
  });

  makeChart('chart-reservation', {
    type: 'bar',
    data: {
      labels: trendLabels,
      datasets: [{
        label: '예약신청',
        data: reservationValues,
        backgroundColor: '#ff7a00',
        borderRadius: 8
      }]
    },
    options: chartBaseOptions()
  });

  makeChart('chart-compare', {
    type: 'bar',
    data: {
      labels: ['당월 / 당주', '전월 동기', '전년 동기'],
      datasets: [{
        label: METRIC_DEFS[compareMetric]?.label || '조회수',
        data: [
          cur[compareMetric] || 0,
          prevMo ? prevMo[compareMetric] || 0 : 0,
          prevYr ? prevYr[compareMetric] || 0 : 0
        ],
        backgroundColor: ['#ff7a00', '#38a7ff', '#8b5cf6'],
        borderRadius: 8
      }]
    },
    options: chartBaseOptions()
  });
}

function renderDetailTable(rows) {
  const tbody = $('detail-tbody');
  if (!tbody) return;

  const total = sumRows(rows);

  const html = rows.map(r => `
    <tr>
      <td>${r.store}</td>
      <td>${r.team}</td>
      <td>${fmt(r.views)}</td>
      <td>${fmt(r.conn)}</td>
      <td>${fmt(r.miss)}</td>
      <td>${fmt(r.res_in)}</td>
      <td>${fmt(r.res_req)}</td>
      <td>${fmt(r.review)}</td>
      <td>${fmt(r.chat)}</td>
    </tr>
  `).join('');

  tbody.innerHTML = html + `
    <tr class="total-row">
      <td>합계</td>
      <td>-</td>
      <td>${fmt(total.views)}</td>
      <td>${fmt(total.conn)}</td>
      <td>${fmt(total.miss)}</td>
      <td>${fmt(total.res_in)}</td>
      <td>${fmt(total.res_req)}</td>
      <td>${fmt(total.review)}</td>
      <td>${fmt(total.chat)}</td>
    </tr>
  `;
}

function renderInsight(cur, prevMo) {
  const insightEl = $('weekly-insight');
  if (!insightEl) return;

  if (!prevMo || !hasAnyData(prevMo)) {
    insightEl.textContent = '비교 가능한 전월 동기 데이터가 없어 이번 주 성과만 표시하고 있습니다.';
    return;
  }

  const viewsRate = deltaRate(cur.views, prevMo.views, true);
  const callRate = deltaRate(cur.conn, prevMo.conn, true);
  const reserveRate = deltaRate(cur.res_req, prevMo.res_req, true);
  const missRate = deltaRate(cur.miss, prevMo.miss, false);

  const messages = [];

  if (viewsRate) {
    messages.push(`플레이스 조회수는 전월 동기 대비 ${viewsRate.value >= 0 ? '상승' : '하락'}했습니다.`);
  }

  if (callRate) {
    messages.push(`연결콜은 ${callRate.value >= 0 ? '증가' : '감소'} 흐름입니다.`);
  }

  if (reserveRate) {
    messages.push(`예약 신청은 ${reserveRate.value >= 0 ? '개선' : '감소'}되었습니다.`);
  }

  if (missRate && missRate.good) {
    messages.push('미연결콜은 줄어 응대 관리 측면에서 긍정적입니다.');
  }

  insightEl.textContent = messages.join(' ') || '이번 주 핵심 지표를 확인해보세요.';
}

function renderDashboard() {
  const selWeek = $('sel-week');
  const selStore = $('sel-store');
  const selTeam = $('sel-team');

  if (!selWeek || !selStore || !selTeam) return;

  const weekKey = selWeek.value;
  const store = selStore.value || 'ALL';
  const team = selTeam.value || 'ALL';
  const meta = weekMeta[weekKey];

  if (!meta) return;

  const pmKey = prevMonthKey(meta.yr, meta.mo, meta.wk);
  const pyKey = prevYearKey(meta.yr, meta.mo, meta.wk);

  const prevMoMeta = pmKey ? weekMeta[pmKey] : null;
  const prevYrMeta = pyKey ? weekMeta[pyKey] : null;

  const curRows = getRowsByWeekKey(weekKey, store, team);
  const curSum = sumRows(curRows);

  const prevMoSum = pmKey ? sumRows(getRowsByWeekKey(pmKey, store, team)) : null;
  const prevYrSum = pyKey ? sumRows(getRowsByWeekKey(pyKey, store, team)) : null;

  const currentIndex = weekKeys.indexOf(weekKey);
  const trendKeys = weekKeys.slice(Math.max(0, currentIndex - 6), currentIndex + 1);

  lastRenderData = {
    meta,
    curRows,
    curSum,
    prevMoSum,
    prevYrSum
  };

  renderPeriodCards(meta, prevMoMeta, prevYrMeta);
  renderKpiCards(curSum, prevMoSum, prevYrSum);
  renderCharts(curSum, prevMoSum, prevYrSum, trendKeys, store, team);
  renderDetailTable(curRows);
  renderInsight(curSum, prevMoSum);
}

// ─────────────────────────────────────────────────────────────────────────────
// KPI 모달
// ─────────────────────────────────────────────────────────────────────────────

function ensureKpiModal() {
  let modal = $('kpiModal');

  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'kpiModal';
    document.body.appendChild(modal);
  }

  modal.className = 'kpi-modal';
  modal.hidden = true;
  modal.style.display = 'none';

  modal.innerHTML = `
    <div class="kpi-modal-panel" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
      <div class="kpi-modal-head">
        <div>
          <div id="modalGroup" class="kpi-modal-group">지표 상세</div>
          <h2 id="modalTitle" class="kpi-modal-title">-</h2>
        </div>
        <button id="modalCloseBtn" class="kpi-modal-close" type="button" aria-label="닫기">×</button>
      </div>

      <div id="modalValue" class="kpi-modal-value">-</div>

      <div class="kpi-modal-compare">
        <div class="kpi-modal-box">
          <span class="kpi-modal-box-label">전월 동기</span>
          <span id="modalPrevMonth" class="kpi-modal-box-value">-</span>
          <span id="modalPrevMonthRate" class="kpi-modal-box-rate">-</span>
        </div>

        <div class="kpi-modal-box">
          <span class="kpi-modal-box-label">전년 동기</span>
          <span id="modalPrevYear" class="kpi-modal-box-value">-</span>
          <span id="modalPrevYearRate" class="kpi-modal-box-rate">-</span>
        </div>
      </div>

      <p id="modalDesc" class="kpi-modal-desc">-</p>
    </div>
  `;

  const closeBtn = $('modalCloseBtn');
  if (closeBtn) closeBtn.addEventListener('click', closeKpiModal);

  modal.addEventListener('click', e => {
    if (e.target === modal) closeKpiModal();
  });
}

function openKpiModal(metricKey) {
  if (!lastRenderData) return;

  const def = METRIC_DEFS[metricKey];
  if (!def) return;

  const cur = lastRenderData.curSum;
  const prevMo = lastRenderData.prevMoSum;
  const prevYr = lastRenderData.prevYrSum;

  const currentValue = fmt(cur[metricKey]);
  const prevMonthValue = prevMo ? fmt(prevMo[metricKey]) : '-';
  const prevYearValue = prevYr ? fmt(prevYr[metricKey]) : '-';

  const moRate = prevMo ? deltaRate(cur[metricKey], prevMo[metricKey], def.higher) : null;
  const yrRate = prevYr ? deltaRate(cur[metricKey], prevYr[metricKey], def.higher) : null;

  const moRateHtml = moRate
    ? `<span class="${moRate.good ? 'rate-up' : 'rate-down'}">${moRate.value >= 0 ? '▲' : '▼'} ${Math.abs(moRate.value).toFixed(1)}%</span>`
    : `<span class="rate-flat">비교 불가</span>`;

  const yrRateHtml = yrRate
    ? `<span class="${yrRate.good ? 'rate-up' : 'rate-down'}">${yrRate.value >= 0 ? '▲' : '▼'} ${Math.abs(yrRate.value).toFixed(1)}%</span>`
    : `<span class="rate-flat">비교 불가</span>`;

  setTextIfExists('modalGroup', def.group);
  setTextIfExists('modalTitle', def.label);
  setTextIfExists('modalValue', currentValue);
  setTextIfExists('modalPrevMonth', prevMonthValue);
  setTextIfExists('modalPrevYear', prevYearValue);
  setHtmlIfExists('modalPrevMonthRate', moRateHtml);
  setHtmlIfExists('modalPrevYearRate', yrRateHtml);
  setTextIfExists('modalDesc', def.desc);

  const modal = $('kpiModal');
  if (!modal) return;

  modal.hidden = false;
  modal.style.display = 'grid';
  modal.classList.add('is-open');
  document.body.style.overflow = 'hidden';
}

function closeKpiModal() {
  const modal = $('kpiModal');
  if (!modal) return;

  modal.hidden = true;
  modal.style.display = 'none';
  modal.classList.remove('is-open');
  document.body.style.overflow = '';
}

// ─────────────────────────────────────────────────────────────────────────────
// 새로고침 / 초기 실행
// ─────────────────────────────────────────────────────────────────────────────

async function refreshData() {
  const btn = $('refreshBtn');
  const originalText = btn ? btn.textContent : '';

  try {
    if (btn) {
      btn.disabled = true;
      btn.textContent = '불러오는 중...';
    }

    await loadData(true);
    initSelects();
    renderDashboard();

    if ($('empty-state')) $('empty-state').hidden = true;
  } catch (e) {
    console.error(e);
    if ($('empty-state')) {
      $('empty-state').hidden = false;
      $('empty-state').textContent = '데이터를 불러오지 못했습니다. Apps Script 배포 URL 또는 데이터 형식을 확인해주세요.';
    }
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = originalText;
    }
  }
}

async function init() {
  injectRuntimeCss();
  ensureKpiModal();
  closeKpiModal();

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeKpiModal();
  });

  const refreshBtn = $('refreshBtn');
  if (refreshBtn) refreshBtn.addEventListener('click', refreshData);

  try {
    await loadData(false);

    if (!RAW.length) {
      if ($('empty-state')) {
        $('empty-state').hidden = false;
        $('empty-state').textContent = '표시할 데이터가 없습니다. Apps Script 응답 데이터와 시트 컬럼명을 확인해주세요.';
      }
      return;
    }

    initSelects();
    renderDashboard();

    if ($('empty-state')) $('empty-state').hidden = true;
    closeKpiModal();
  } catch (e) {
    console.error(e);

    if ($('empty-state')) {
      $('empty-state').hidden = false;
      $('empty-state').textContent = '데이터를 불러오지 못했습니다. Apps Script 배포 URL 또는 데이터 형식을 확인해주세요.';
    }

    closeKpiModal();
  }
}

document.addEventListener('DOMContentLoaded', init);
