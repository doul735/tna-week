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
    icon: '◎',
    higher: true,
    desc: '네이버 플레이스에서 매장이 얼마나 많이 조회되었는지 확인하는 핵심 노출 지표입니다.'
  },
  conn: {
    label: '연결콜',
    group: '전환 관련 지표',
    groupKey: 'convert',
    icon: '☎',
    higher: true,
    desc: '고객이 실제로 전화 연결까지 이어진 수치입니다. 문의 전환 흐름을 확인할 수 있습니다.'
  },
  miss: {
    label: '미연결콜',
    group: '관리 관련 지표',
    groupKey: 'manage',
    icon: '×',
    higher: false,
    desc: '연결되지 못한 전화입니다. 낮아질수록 응대 관리가 개선된 것으로 볼 수 있습니다.'
  },
  res_in: {
    label: '예약유입',
    group: '노출 관련 지표',
    groupKey: 'exposure',
    icon: '↗',
    higher: true,
    desc: '예약 화면까지 유입된 고객 수입니다. 관심 고객의 행동 흐름을 볼 수 있습니다.'
  },
  res_req: {
    label: '예약신청',
    group: '전환 관련 지표',
    groupKey: 'convert',
    icon: '▣',
    higher: true,
    desc: '실제 예약 신청까지 이어진 수치입니다. 예약 전환 성과를 보는 핵심 지표입니다.'
  },
  review: {
    label: '리뷰',
    group: '관리 관련 지표',
    groupKey: 'manage',
    icon: '☆',
    higher: true,
    desc: '고객 리뷰 발생 수입니다. 매장 신뢰도와 운영 관리 상태를 함께 보여주는 지표입니다.'
  },
  chat: {
    label: '톡톡상담',
    group: '관리 관련 지표',
    groupKey: 'manage',
    icon: '●',
    higher: true,
    desc: '네이버 톡톡 상담 수치입니다. 전화 외 문의 채널의 반응을 확인할 수 있습니다.'
  }
};

function $(id) {
  return document.getElementById(id);
}

function num(value) {
  if (value === null || value === undefined || value === '') return 0;
  return Number(String(value).replace(/,/g, '')) || 0;
}

function fmt(value) {
  return num(value).toLocaleString();
}

function toDateString(value) {
  if (!value) return '';

  if (typeof value === 'number') {
    const base = new Date(1899, 11, 30);
    base.setDate(base.getDate() + value);
    return base.toISOString().split('T')[0];
  }

  return String(value).split('T')[0];
}

function cacheBustUrl(url) {
  return url + (url.includes('?') ? '&' : '?') + 'v=' + Date.now();
}

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

function normalizeRows(data) {
  const rows = Array.isArray(data) ? data : Array.isArray(data?.data) ? data.data : [];

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
    neg_review: num(r['부정리뷰'] ?? r.neg_review ?? r.negative_review),
    chat: num(r['톡톡상담'] ?? r.chat),
  };
}

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
    return new Date(weekMeta[a].start) - new Date(weekMeta[b].start);
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
  const s = {
    views: 0,
    conn: 0,
    miss: 0,
    res_in: 0,
    res_req: 0,
    review: 0,
    neg_review: 0,
    chat: 0
  };
  rows.forEach(r => METRICS.forEach(m => { s[m] += num(r[m]); }));
  return s;
}

  return result;
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
  const [yr, mo] = monthKey.split('-').map(Number);

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

  const rate = ((num(current) - num(previous)) / num(previous)) * 100;

  return {
    value: rate,
    good: higher ? rate >= 0 : rate <= 0
  };
}

function rateHtml(current, previous, higher, label) {
  const rate = deltaRate(current, previous, higher);

  if (!rate) {
    return `<span><span class="rate-label">${label}</span><span class="rate-flat">—</span></span>`;
  }

  const arrow = rate.value > 0 ? '↑' : rate.value < 0 ? '↓' : '–';
  const cls = rate.good ? 'rate-up' : 'rate-down';

  return `<span><span class="rate-label">${label}</span><span class="${cls}">${arrow} ${Math.abs(rate.value).toFixed(1)}%</span></span>`;
}

function createSparkline(values, groupKey) {
  const clean = values.map(num);
  const max = Math.max(...clean, 1);
  const min = Math.min(...clean);
  const range = max - min || 1;

  const points = clean.map((v, i) => {
    const x = clean.length === 1 ? 100 : (i / (clean.length - 1)) * 100;
    const y = 34 - ((v - min) / range) * 28;
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  }).join(' ');

  return `
    <svg class="sparkline ${groupKey}" viewBox="0 0 100 44" preserveAspectRatio="none" aria-hidden="true">
      <polyline points="${points}"></polyline>
    </svg>
  `;
}

function initSelects() {
  const selStore = $('sel-store');
  const selTeam = $('sel-team');
  const selWeek = $('sel-week');
  const selMonth = $('sel-month');

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

  const latestDataWeek = [...weekKeys].reverse().find(key => hasAnyData(sumRows(getRowsByWeekKey(key))));
  if (latestDataWeek) {
    selWeek.value = latestDataWeek;
    const m = weekMeta[latestDataWeek];
    selMonth.value = `${m.yr}-${m.mo}`;
  } else if (weekKeys.length) {
    selWeek.value = weekKeys[weekKeys.length - 1];
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

  $('trendMetricSelect').onchange = renderDashboard;
  $('compareMetricSelect').onchange = renderDashboard;
}

function renderPeriodCards(meta, prevMoMeta, prevYrMeta) {
  const currentText = `${meta.start} ~ ${meta.end}<br>선택 주차 기준`;
  $('period-current').innerHTML = currentText;
  $('header-period').textContent = `${meta.start} ~ ${meta.end} · ${meta.mo}월 ${meta.wk}주차`;

  $('period-prev-month').innerHTML = prevMoMeta
    ? `${prevMoMeta.start} ~ ${prevMoMeta.end}<br>전월 같은 주차`
    : '비교 데이터 없음';

  $('period-prev-year').innerHTML = prevYrMeta
    ? `${prevYrMeta.start} ~ ${prevYrMeta.end}<br>전년 같은 기간`
    : '비교 데이터 없음';
}

function renderKpiCards(cur, prevMo, prevYr, trendKeys, store, team) {
  const grid = $('kpi-grid');

  grid.innerHTML = METRICS.map(key => {
    const def = METRIC_DEFS[key];

    const trendValues = trendKeys.map(weekKey => {
      return sumRows(getRowsByWeekKey(weekKey, store, team))[key];
    });

    return `
      <article class="kpi-card ${def.groupKey}" data-metric="${key}">
        <div class="kpi-top">
          <div class="kpi-icon">${def.icon}</div>
          <div>
            <h3 class="kpi-title">${def.label}</h3>
            <div class="kpi-group">${def.group}</div>
          </div>
        </div>

        <div class="kpi-value">${fmt(cur[key])}</div>

        <div class="kpi-rates">
          ${rateHtml(cur[key], prevMo ? prevMo[key] : null, def.higher, '전월 동기')}
          ${rateHtml(cur[key], prevYr ? prevYr[key] : null, def.higher, '전년 동기')}
        </div>

        ${createSparkline(trendValues, def.groupKey)}
      </article>
    `;
  }).join('');

  document.querySelectorAll('.kpi-card').forEach(card => {
    card.addEventListener('click', () => openKpiModal(card.dataset.metric));
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
  const trendMetric = $('trendMetricSelect').value;
  const compareMetric = $('compareMetricSelect').value;

  const trendLabels = trendKeys.map(key => weekMeta[key].short);
  const trendValues = trendKeys.map(key => sumRows(getRowsByWeekKey(key, store, team))[trendMetric]);
  const reservationValues = trendKeys.map(key => sumRows(getRowsByWeekKey(key, store, team)).res_req);

  makeChart('chart-trend', {
    type: 'line',
    data: {
      labels: trendLabels,
      datasets: [{
        label: METRIC_DEFS[trendMetric].label,
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

  makeChart('chart-groups', {
    type: 'doughnut',
    data: {
      labels: ['노출 관련 지표', '전환 관련 지표', '관리 관련 지표'],
      datasets: [{
        data: [2, 2, 3],
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
        label: METRIC_DEFS[compareMetric].label,
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
  if (!prevMo || !hasAnyData(prevMo)) {
    $('weekly-insight').textContent = '비교 가능한 전월 동기 데이터가 없어 이번 주 성과만 표시하고 있습니다.';
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

  $('weekly-insight').textContent = messages.join(' ') || '이번 주 핵심 지표를 확인해보세요.';
}

function renderDashboard() {
  const weekKey = $('sel-week').value;
  const store = $('sel-store').value;
  const team = $('sel-team').value;
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
    curSum,
    prevMoSum,
    prevYrSum
  };

  renderPeriodCards(meta, prevMoMeta, prevYrMeta);
  renderKpiCards(curSum, prevMoSum, prevYrSum, trendKeys, store, team);
  renderCharts(curSum, prevMoSum, prevYrSum, trendKeys, store, team);
  renderDetailTable(curRows);
  renderInsight(curSum, prevMoSum);
}

function openKpiModal(metricKey) {
  if (!lastRenderData) return;

  const def = METRIC_DEFS[metricKey];
  const cur = lastRenderData.curSum;
  const prevMo = lastRenderData.prevMoSum;
  const prevYr = lastRenderData.prevYrSum;

  $('modalGroup').textContent = def.group;
  $('modalTitle').textContent = def.label;
  $('modalValue').textContent = fmt(cur[metricKey]);
  $('modalPrevMonth').textContent = prevMo ? fmt(prevMo[metricKey]) : '-';
  $('modalPrevYear').textContent = prevYr ? fmt(prevYr[metricKey]) : '-';

  const moRate = prevMo ? deltaRate(cur[metricKey], prevMo[metricKey], def.higher) : null;
  const yrRate = prevYr ? deltaRate(cur[metricKey], prevYr[metricKey], def.higher) : null;

  $('modalPrevMonthRate').innerHTML = moRate
    ? `<span class="${moRate.good ? 'rate-up' : 'rate-down'}">${moRate.value >= 0 ? '↑' : '↓'} ${Math.abs(moRate.value).toFixed(1)}%</span>`
    : '<span class="rate-flat">비교 불가</span>';

  $('modalPrevYearRate').innerHTML = yrRate
    ? `<span class="${yrRate.good ? 'rate-up' : 'rate-down'}">${yrRate.value >= 0 ? '↑' : '↓'} ${Math.abs(yrRate.value).toFixed(1)}%</span>`
    : '<span class="rate-flat">비교 불가</span>';

  $('modalDesc').textContent = def.desc;

  $('kpiModal').hidden = false;
  document.body.style.overflow = 'hidden';
}

function closeKpiModal() {
  $('kpiModal').hidden = true;
  document.body.style.overflow = '';
}

async function loadData(forceFresh = false) {
  const data = await fetchJsonWithCache(DATA_URL, 'tna_place_dashboard_rows_v3', 5 * 60 * 1000, forceFresh);
  const rows = normalizeRows(data);

  RAW.length = 0;
  rows.forEach(row => RAW.push(row));

  rebuildMetaFromRaw();
}

async function refreshData() {
  const btn = $('refreshBtn');
  const originalText = btn.textContent;

  try {
    btn.disabled = true;
    btn.textContent = '불러오는 중...';

    await loadData(true);
    initSelects();
    renderDashboard();

    $('empty-state').hidden = true;
  } catch (e) {
    console.error(e);
    $('empty-state').hidden = false;
  } finally {
    btn.disabled = false;
    btn.textContent = originalText;
  }
}

async function init() {
  $('modalCloseBtn').addEventListener('click', closeKpiModal);

  $('kpiModal').addEventListener('click', e => {
    if (e.target.id === 'kpiModal') closeKpiModal();
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeKpiModal();
  });

  $('refreshBtn').addEventListener('click', refreshData);

  try {
    await loadData(false);

    if (!RAW.length) {
      $('empty-state').hidden = false;
      return;
    }

    initSelects();
    renderDashboard();
    $('empty-state').hidden = true;
  } catch (e) {
    console.error(e);
    $('empty-state').hidden = false;
  }
}

document.addEventListener('DOMContentLoaded', init);
// ─────────────────────────────────────────────
// KPI 모달 강제 수정 코드
// 사이트 첫 진입 시 모달 숨김 + 닫기 버튼 정상 작동
// ─────────────────────────────────────────────

function hideKpiModal() {
  const modal = document.getElementById('kpiModal');
  if (!modal) return;

  modal.setAttribute('hidden', '');
  modal.style.display = 'none';
  document.body.style.overflow = '';
}

function showKpiModal() {
  const modal = document.getElementById('kpiModal');
  if (!modal) return;

  modal.removeAttribute('hidden');
  modal.style.display = 'grid';
  document.body.style.overflow = 'hidden';
}

// 기존 openKpiModal 덮어쓰기
window.openKpiModal = function (data) {
  const title = document.getElementById('kpiModalTitle');
  const value = document.getElementById('kpiModalValue');
  const prevMonth = document.getElementById('kpiModalPrevMonth');
  const prevYear = document.getElementById('kpiModalPrevYear');
  const deltaMonth = document.getElementById('kpiModalDeltaMonth');
  const deltaYear = document.getElementById('kpiModalDeltaYear');

  if (title) title.textContent = data.title || '-';
  if (value) value.textContent = data.current || '-';
  if (prevMonth) prevMonth.textContent = data.prevMonth || '-';
  if (prevYear) prevYear.textContent = data.prevYear || '-';
  if (deltaMonth) deltaMonth.textContent = data.deltaMonth || '-';
  if (deltaYear) deltaYear.textContent = data.deltaYear || '-';

  showKpiModal();
};

// 기존 closeKpiModal 덮어쓰기
window.closeKpiModal = function () {
  hideKpiModal();
};

// 페이지 로딩 직후 무조건 모달 숨기기
window.addEventListener('load', function () {
  hideKpiModal();

  const modal = document.getElementById('kpiModal');
  if (!modal) return;

  const closeBtn = modal.querySelector('.modal-close, .kpi-modal-close, button[aria-label="닫기"], button');

  if (closeBtn) {
    closeBtn.onclick = function (e) {
      e.preventDefault();
      e.stopPropagation();
      hideKpiModal();
    };
  }

  modal.addEventListener('click', function (e) {
    if (e.target === modal) {
      hideKpiModal();
    }
  });

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      hideKpiModal();
    }
  });
});
