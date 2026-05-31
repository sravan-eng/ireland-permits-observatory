// ============================================================
// Ireland Work Permits Observatory — main.js
// All charts built from real DETE xlsx data via extract-data.py
// ============================================================

import {
  YEARS, totalByYear, byPermitType, bySector,
  nationalities2026, topNationalities2025, nationalityTrend,
  byCounty2025, events, SECTOR_COLORS, PERMIT_COLORS, PERMIT_LABELS
} from './data/permits.js';

import { COMPANIES, INDUSTRY_LABELS } from './data/companies.js';

// ── CHART.JS GLOBAL DEFAULTS ──────────────────────────────────
Chart.defaults.color = '#6b7280';
Chart.defaults.borderColor = 'rgba(0,0,0,0.07)';
Chart.defaults.font.family = "'DM Sans', sans-serif";
Chart.defaults.plugins.legend.display = false;
Chart.defaults.plugins.tooltip.backgroundColor = '#111827';
Chart.defaults.plugins.tooltip.borderColor = 'rgba(0,0,0,0.1)';
Chart.defaults.plugins.tooltip.borderWidth = 1;
Chart.defaults.plugins.tooltip.padding = 12;
Chart.defaults.plugins.tooltip.titleColor = '#f9fafb';
Chart.defaults.plugins.tooltip.bodyColor = '#9ca3af';
Chart.defaults.plugins.tooltip.cornerRadius = 8;

// ── STATE ─────────────────────────────────────────────────────
let selectedYear    = 2024;
let timelineMode    = 'total';
let showEventsFlag  = true;
let raceInterval    = null;
let raceYearIndex   = 0;
const raceYears     = [2019,2020,2021,2022,2023,2024,2025,2026];
let activeNats      = new Set(['India','Philippines','Brazil']);
let explorerChart   = null;
let searchQuery     = '';
let activeIndustry  = 'all';
let sortMode        = 'total_desc';
let visibleCount    = 12;

// ── HELPERS ───────────────────────────────────────────────────
const fmt  = n => (n||0).toLocaleString();
const pct  = (n, total) => total ? ((n/total)*100).toFixed(1) : '0.0';

// ── SECTION REVEAL ────────────────────────────────────────────
const revealObs = new IntersectionObserver(entries =>
  entries.forEach(e => { if (e.isIntersecting) e.target.classList.add('visible'); }),
  { threshold: 0.08 }
);
document.querySelectorAll('.section').forEach(s => revealObs.observe(s));

// ── NAV ACTIVE ────────────────────────────────────────────────
const navObs = new IntersectionObserver(entries => {
  entries.forEach(e => {
    if (e.isIntersecting) {
      document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
      const a = document.querySelector(`.nav-link[href="#${e.target.id}"]`);
      if (a) a.classList.add('active');
    }
  });
}, { threshold: 0.3 });
document.querySelectorAll('section[id]').forEach(s => navObs.observe(s));

// ── HERO COUNTERS ─────────────────────────────────────────────
function animateCounter(el, target, dur = 1800) {
  const start = performance.now();
  const tick  = now => {
    const p = Math.min((now - start) / dur, 1);
    el.textContent = Math.round((1 - Math.pow(1-p, 3)) * target).toLocaleString();
    if (p < 1) requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);
}
window.addEventListener('load', () =>
  document.querySelectorAll('[data-count]').forEach(el =>
    animateCounter(el, parseInt(el.dataset.count))
  )
);

// ── KPI CARDS ─────────────────────────────────────────────────
function updateKPIs(year) {
  const total = totalByYear[year] || 0;
  const types = byPermitType[year] || {};
  const prev  = totalByYear[year - 1];
  const chg   = prev ? ((total - prev) / prev * 100).toFixed(1) : null;

  document.getElementById('kpi-total').textContent   = fmt(total);
  document.getElementById('donut-total').textContent = fmt(total);
  document.getElementById('donut-year').textContent  = year;
  document.getElementById('sector-bar-year').textContent = year;

  const chgEl = document.getElementById('kpi-change');
  if (chg) {
    chgEl.textContent = `${+chg > 0 ? '▲' : '▼'} ${Math.abs(chg)}% vs ${year-1}`;
    chgEl.style.color = +chg > 0 ? '#00a878' : '#e53e3e';
  } else {
    chgEl.textContent = 'First year';
  }

  const cs = types.critical_skills || 0;
  const ge = types.general || 0;
  const ic = types.intra_company || 0;
  document.getElementById('kpi-cs').textContent       = fmt(cs);
  document.getElementById('kpi-cs-share').textContent = `${pct(cs,total)}% of total`;
  document.getElementById('kpi-ge').textContent       = fmt(ge);
  document.getElementById('kpi-ge-share').textContent = `${pct(ge,total)}% of total`;
  document.getElementById('kpi-ict').textContent      = fmt(ic);
  document.getElementById('kpi-ict-share').textContent= `${pct(ic,total)}% of total`;
}

// ── DONUT CHART ───────────────────────────────────────────────
let donutChart = null;
function buildDonut(year) {
  const types  = byPermitType[year] || {};
  const keys   = ['critical_skills','general','intra_company','dependant','other'];
  const data   = keys.map(k => types[k] || 0);
  const colors = keys.map(k => PERMIT_COLORS[k]);
  const ctx    = document.getElementById('donutChart').getContext('2d');
  if (donutChart) donutChart.destroy();

  donutChart = new Chart(ctx, {
    type: 'doughnut',
    data: { labels: keys.map(k => PERMIT_LABELS[k]), datasets: [{
      data, backgroundColor: colors,
      borderWidth: 2, borderColor: '#fff', hoverOffset: 8
    }]},
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '68%',
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: {
          label: c => ` ${c.label}: ${fmt(c.parsed)} (${pct(c.parsed, data.reduce((a,b)=>a+b,0))}%)`
        }}
      }
    }
  });

  document.getElementById('donut-legend').innerHTML = keys.map((k,i) => `
    <div class="legend-item">
      <span class="legend-dot" style="background:${colors[i]}"></span>
      <span>${PERMIT_LABELS[k]}</span>
    </div>`).join('');
}

// ── SECTOR BAR CHART ──────────────────────────────────────────
let sectorBarChart = null;
function buildSectorBar(year) {
  const raw    = bySector[year] || bySector[2024];
  const sorted = Object.entries(raw).sort((a,b) => b[1]-a[1]).slice(0, 12);
  const labels = sorted.map(([k]) => k);
  const values = sorted.map(([,v]) => v);
  const colors = sorted.map(([k]) => SECTOR_COLORS[k] || '#9ca3af');
  const ctx    = document.getElementById('sectorBarChart').getContext('2d');
  if (sectorBarChart) sectorBarChart.destroy();

  sectorBarChart = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ data: values, backgroundColor: colors, borderRadius: 4, borderSkipped: false }]},
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => fmt(v) }},
        y: { grid: { display: false }, ticks: { font: { size: 11 }}}
      },
      plugins: { legend: { display: false },
        tooltip: { callbacks: { label: c => ` ${fmt(c.parsed.x)} permits` }}
      }
    }
  });
}

// ── TIMELINE CHART ────────────────────────────────────────────
let timelineChart = null;
function buildTimeline() {
  const years = YEARS.filter(y => y >= 2009);
  const ctx   = document.getElementById('timelineChart').getContext('2d');
  if (timelineChart) timelineChart.destroy();

  let datasets;
  if (timelineMode === 'total') {
    datasets = [{
      label: 'Total permits',
      data:  years.map(y => totalByYear[y] || null),
      borderColor: '#00a878', backgroundColor: 'rgba(0,168,120,0.08)',
      borderWidth: 2.5, fill: true, tension: 0.4,
      pointRadius: 4, pointBackgroundColor: '#00a878',
      pointBorderColor: '#fff', pointBorderWidth: 2,
    }];
  } else {
    datasets = Object.keys(PERMIT_COLORS).map(t => ({
      label: PERMIT_LABELS[t],
      data:  years.map(y => byPermitType[y]?.[t] || null),
      borderColor: PERMIT_COLORS[t],
      backgroundColor: PERMIT_COLORS[t] + '18',
      borderWidth: 2, fill: false, tension: 0.4, pointRadius: 3,
    }));
  }

  timelineChart = new Chart(ctx, {
    type: 'line',
    data: { labels: years.map(String), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }}},
        y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => fmt(v) }}
      },
      plugins: {
        legend: { display: timelineMode !== 'total', position: 'top',
          labels: { usePointStyle: true, padding: 16, font: { size: 11 }}},
        tooltip: { callbacks: {
          title: c => `Year: ${c[0].label}`,
          afterBody: c => {
            const yr  = parseInt(c[0].label);
            const ev  = events.find(e => e.year === yr);
            return ev ? [`\n📌 ${ev.label}`] : [];
          }
        }}
      }
    }
  });
}

function renderEventPills() {
  document.getElementById('eventPills').innerHTML = events.map(ev =>
    `<div class="event-pill"><span>${ev.year}</span>${ev.label}</div>`
  ).join('');
}

// ── RACE CHART ────────────────────────────────────────────────
function renderRaceChart(year) {
  const data   = bySector[year] || bySector[2024];
  const sorted = Object.entries(data).sort((a,b) => b[1]-a[1]).slice(0,9);
  const maxVal = sorted[0][1];
  document.getElementById('raceYear').textContent = year;

  document.getElementById('raceChart').innerHTML = sorted.map(([name, val], i) => {
    const color = SECTOR_COLORS[name] || '#9ca3af';
    const w     = ((val / maxVal) * 100).toFixed(1);
    return `
      <div class="race-row">
        <span class="race-rank">${i+1}</span>
        <span class="race-sector-name">${name}</span>
        <div class="race-bar-wrap">
          <div class="race-bar" style="width:${w}%;background:${color}">
            <span class="race-bar-val">${fmt(val)}</span>
          </div>
        </div>
      </div>`;
  }).join('');
}

function buildRacePlay() {
  const playBtn  = document.getElementById('racePlay');
  const resetBtn = document.getElementById('raceReset');

  playBtn.addEventListener('click', () => {
    if (raceInterval) {
      clearInterval(raceInterval); raceInterval = null;
      playBtn.textContent = '▶ Play'; return;
    }
    playBtn.textContent = '⏸ Pause';
    raceInterval = setInterval(() => {
      raceYearIndex = (raceYearIndex + 1) % raceYears.length;
      renderRaceChart(raceYears[raceYearIndex]);
      if (raceYearIndex === raceYears.length - 1) {
        clearInterval(raceInterval); raceInterval = null;
        playBtn.textContent = '▶ Play';
      }
    }, 900);
  });

  resetBtn.addEventListener('click', () => {
    clearInterval(raceInterval); raceInterval = null;
    raceYearIndex = 0; playBtn.textContent = '▶ Play';
    renderRaceChart(raceYears[0]);
  });
}

// ── STACKED AREA ──────────────────────────────────────────────
let stackedAreaChart = null;
function buildStackedArea() {
  const years   = [2019,2020,2021,2022,2023,2024,2025,2026];
  const sectors = Object.keys(SECTOR_COLORS);
  const ctx     = document.getElementById('stackedAreaChart').getContext('2d');
  if (stackedAreaChart) stackedAreaChart.destroy();

  stackedAreaChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: years.map(String),
      datasets: sectors.map(s => ({
        label: s,
        data:  years.map(y => bySector[y]?.[s] || 0),
        borderColor: SECTOR_COLORS[s],
        backgroundColor: SECTOR_COLORS[s] + '55',
        fill: true, tension: 0.4, borderWidth: 1.5, pointRadius: 0,
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }},
        y: { stacked: true, grid: { color: 'rgba(0,0,0,0.05)' },
          ticks: { callback: v => fmt(v) }}
      },
      plugins: {
        legend: { display: true, position: 'bottom',
          labels: { usePointStyle: true, pointStyleWidth: 10,
            font: { size: 10 }, padding: 10 }}
      }
    }
  });
}

// ── NATIONALITY CHART ─────────────────────────────────────────
// Data from xlsx: { country, issued, refused, flag }
let natChart = null;
function buildNatChart() {
  const top15 = [...topNationalities2025]
    .sort((a,b) => b.issued - a.issued)
    .slice(0,15);

  const COLORS = ['#00a878','#2563eb','#7c3aed','#d97706','#e53e3e',
    '#5DCAA5','#97C459','#D4537E','#EF9F27','#378ADD',
    '#F0997B','#AFA9EC','#1D9E75','#D85A30','#E24B4A'];

  const ctx = document.getElementById('natChart').getContext('2d');
  if (natChart) natChart.destroy();

  natChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: top15.map(n => `${n.flag} ${n.country}`),
      datasets: [{
        data: top15.map(n => n.issued),            // ← REAL .issued field
        backgroundColor: COLORS,
        borderRadius: 4, borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => fmt(v) }},
        y: { grid: { display: false }, ticks: { font: { size: 11 }}}
      },
      plugins: { legend: { display: false },
        tooltip: { callbacks: {
          label: c => ` ${fmt(c.parsed.x)} permits issued`
        }}
      }
    }
  });
}

function buildNatBubbles() {
  const top12 = [...topNationalities2025]
    .sort((a,b) => b.issued - a.issued)
    .slice(0,12);
  document.getElementById('natBubbleGrid').innerHTML = top12.map(n => `
    <div class="nat-bubble">
      <span class="nat-flag">${n.flag}</span>
      <span class="nat-name">${n.country}</span>
      <span class="nat-count-badge">${fmt(n.issued)}</span>
    </div>`).join('');
}

// ── NATIONALITY TREND ─────────────────────────────────────────
let natTrendChart = null;
function buildNatToggles() {
  const nations = Object.keys(nationalityTrend);
  document.getElementById('natToggles').innerHTML = nations.map(n => `
    <button class="nat-toggle-btn ${activeNats.has(n) ? 'active' : ''}"
      data-nat="${n}">${n}</button>`).join('');
  document.querySelectorAll('.nat-toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const nat = btn.dataset.nat;
      activeNats.has(nat) ? activeNats.delete(nat) : activeNats.add(nat);
      btn.classList.toggle('active');
      buildNatTrendChart();
    });
  });
}

function buildNatTrendChart() {
  const years  = [2019,2020,2021,2022,2023,2024,2025,2026];
  const COLORS = ['#00a878','#2563eb','#7c3aed','#d97706'];
  const ctx    = document.getElementById('natTrendChart').getContext('2d');
  if (natTrendChart) natTrendChart.destroy();

  const datasets = Object.keys(nationalityTrend)
    .filter(n => activeNats.has(n))
    .map((n,i) => ({
      label: n,
      data:  years.map(y => nationalityTrend[n]?.[y] || null),
      borderColor: COLORS[i % COLORS.length],
      backgroundColor: COLORS[i % COLORS.length] + '20',
      borderWidth: 2, tension: 0.4, pointRadius: 4, fill: false,
    }));

  natTrendChart = new Chart(ctx, {
    type: 'line',
    data: { labels: years.map(String), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }},
        y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => fmt(v) }}
      },
      plugins: {
        legend: { display: true, position: 'top',
          labels: { usePointStyle: true, font: { size: 11 }, padding: 14 }}
      }
    }
  });
}

// ── COUNTY CHART ──────────────────────────────────────────────
// byCounty2025 is { "Dublin": 5515, "Cork": 1241, ... }
let countyChart = null;
function buildCountyChart() {
  const sorted = Object.entries(byCounty2025).sort((a,b) => b[1]-a[1]);
  const ctx    = document.getElementById('countyChart').getContext('2d');
  if (countyChart) countyChart.destroy();

  const numBars = sorted.length;
  const height  = Math.max(520, numBars * 24 + 80);
  document.getElementById('countyChart').parentElement.style.height = height + 'px';

  countyChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sorted.map(([k]) => k),
      datasets: [{
        data: sorted.map(([,v]) => v),       // ← plain integer values
        backgroundColor: sorted.map(([k]) =>
          k === 'Dublin' ? '#00a878' : k === 'Cork' ? '#2563eb' :
          k === 'Meath' || k === 'Kildare' || k === 'Limerick' ? '#7c3aed' :
          'rgba(37,99,235,0.35)'
        ),
        borderRadius: 4, borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { callback: v => fmt(v) }},
        y: { grid: { display: false }, ticks: { font: { size: 11 }}}
      },
      plugins: { legend: { display: false },
        tooltip: { callbacks: { label: c => ` ${fmt(c.parsed.x)} permits` }}
      }
    }
  });
}

// ── EXPLORER CHART ────────────────────────────────────────────
function buildExplorerChart() {
  const yFrom  = parseInt(document.getElementById('yearFrom').value);
  const yTo    = parseInt(document.getElementById('yearTo').value);
  const byWhat = document.getElementById('breakdownBy').value;
  const types  = [...document.querySelectorAll('#typeFilters input:checked')].map(i => i.value);
  const years  = YEARS.filter(y => y >= yFrom && y <= yTo);
  const ctx    = document.getElementById('explorerChart').getContext('2d');
  if (explorerChart) explorerChart.destroy();

  let datasets;
  if (byWhat === 'year') {
    datasets = [{ label: 'Permits', borderRadius: 6, borderSkipped: false,
      backgroundColor: '#00a878',
      data: years.map(y => types.reduce((s,t) => s + (byPermitType[y]?.[t] || 0), 0))
    }];
  } else if (byWhat === 'type') {
    datasets = types.map(t => ({
      label: PERMIT_LABELS[t], backgroundColor: PERMIT_COLORS[t],
      borderRadius: 4, borderSkipped: false,
      data: years.map(y => byPermitType[y]?.[t] || 0),
    }));
  } else {
    datasets = Object.keys(SECTOR_COLORS).map(s => ({
      label: s, backgroundColor: SECTOR_COLORS[s],
      borderRadius: 4, borderSkipped: false,
      data: years.map(y => bySector[y]?.[s] || 0),
    }));
  }

  document.getElementById('explorerTitle').textContent =
    `Permits ${yFrom}–${yTo} by ${byWhat}`;

  explorerChart = new Chart(ctx, {
    type: 'bar',
    data: { labels: years.map(String), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      scales: {
        x: { stacked: byWhat !== 'year', grid: { color: 'rgba(0,0,0,0.05)' }},
        y: { stacked: byWhat !== 'year', grid: { color: 'rgba(0,0,0,0.05)' },
          ticks: { callback: v => fmt(v) }}
      },
      plugins: {
        legend: { display: byWhat !== 'year', position: 'bottom',
          labels: { usePointStyle: true, font: { size: 10 }, padding: 10 }}
      }
    }
  });
}

// ── YEAR BUTTONS ──────────────────────────────────────────────
document.querySelectorAll('.year-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.year-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    selectedYear = parseInt(btn.dataset.year);
    updateKPIs(selectedYear);
    buildDonut(selectedYear);
    buildSectorBar(selectedYear);
  });
});

// ── TIMELINE TOGGLES ──────────────────────────────────────────
document.querySelectorAll('.toggle-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    timelineMode = btn.dataset.mode;
    buildTimeline();
  });
});
document.getElementById('showEvents').addEventListener('change', function() {
  showEventsFlag = this.checked;
  buildTimeline();
});

// ── EXPLORER ──────────────────────────────────────────────────
document.getElementById('applyExplorer').addEventListener('click', buildExplorerChart);

// ══════════════════════════════════════════════════════════════
// COMPANY SEARCH ENGINE
// Uses real COMPANIES array from companies.js (4,224 companies)
// ══════════════════════════════════════════════════════════════

const PERMIT_BAR_COLORS = { critical: '#2563eb', general: '#00a878', intra: '#7c3aed' };

function getTrend(c) {
  // companies.js has total (2026 Jan-Apr) — show monthly average vs last year estimate
  const avg = Math.round(c.total / 4);    // monthly average 2026
  return { label: `${fmt(c.total)} permits`, cls: 'trend-flat' };
}

function getFilteredSorted() {
  let list = [...COMPANIES];
  if (activeIndustry !== 'all') list = list.filter(c => c.industry === activeIndustry);
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    list = list.filter(c =>
      c.name.toLowerCase().includes(q) ||
      (c.desc && c.desc.toLowerCase().includes(q)) ||
      (c.county && c.county.toLowerCase().includes(q))
    );
  }
  switch (sortMode) {
    case 'total_desc':    list.sort((a,b) => b.total - a.total); break;
    case 'total_asc':     list.sort((a,b) => a.total - b.total); break;
    case 'critical_desc': list.sort((a,b) => (b.jan||0) - (a.jan||0)); break;
    case 'name_asc':      list.sort((a,b) => a.name.localeCompare(b.name)); break;
    case 'trend_desc':    list.sort((a,b) => b.total - a.total); break;
  }
  return list;
}

function renderCompanyCard(c) {
  const maxVal = Math.max(c.jan||0, c.feb||0, c.mar||0, c.apr||0, 1);
  const ind    = c.industry || 'other';
  const label  = INDUSTRY_LABELS[ind] || ind;
  const badgeClass = `badge-${ind}`;
  const county = c.county || 'Ireland';

  // monthly breakdown bars
  const months = [
    { label: 'January',  val: c.jan||0 },
    { label: 'February', val: c.feb||0 },
    { label: 'March',    val: c.mar||0 },
    { label: 'April',    val: c.apr||0 },
  ];
  const barsHtml = months.map(m => `
    <div class="permit-bar-row">
      <span class="pb-label">${m.label}</span>
      <div class="pb-track">
        <div class="pb-fill" style="width:${((m.val/maxVal)*100).toFixed(0)}%;background:#2563eb"></div>
      </div>
      <span class="pb-val">${m.val}</span>
    </div>`).join('');

  return `
    <div class="company-card ind-${ind}">
      <div class="card-top">
        <div class="company-name">${c.name}</div>
        <span class="company-sector-badge ${badgeClass}">${label}</span>
      </div>
      <p style="font-size:0.78rem;color:var(--text-secondary);margin-bottom:1rem;line-height:1.4">
        ${c.desc || label}
      </p>
      <div class="card-metrics">
        <div class="card-metric">
          <span class="cm-val">${fmt(c.total)}</span>
          <span class="cm-label">Jan–Apr 2026</span>
        </div>
        <div class="card-metric">
          <span class="cm-val">${Math.round((c.total||0)/4)}</span>
          <span class="cm-label">Monthly avg</span>
        </div>
        <div class="card-metric">
          <span class="cm-val">${Math.round((c.total||0)/12 * 100)}%</span>
          <span class="cm-label">YTD pace</span>
        </div>
      </div>
      <div class="card-permit-bars">${barsHtml}</div>
      <div class="card-footer">
        <span class="company-county">📍 <strong>${county}</strong></span>
        <span class="trend-badge trend-flat">${fmt(c.total)} total</span>
      </div>
    </div>`;
}

function renderGrid() {
  const list    = getFilteredSorted();
  const grid    = document.getElementById('companyGrid');
  const loadBtn = document.getElementById('loadMoreBtn');
  document.getElementById('resultCount').textContent = list.length.toLocaleString();

  if (list.length === 0) {
    grid.innerHTML = `
      <div class="no-results">
        <span class="nr-icon">🔍</span>
        <h3>No companies found</h3>
        <p>Try a different name or industry filter</p>
      </div>`;
    loadBtn.style.display = 'none';
    return;
  }

  grid.innerHTML = list.slice(0, visibleCount).map(renderCompanyCard).join('');
  loadBtn.style.display = list.length > visibleCount ? 'inline-block' : 'none';
}

// Search events
const searchInput = document.getElementById('companySearch');
const clearBtn    = document.getElementById('searchClear');

searchInput.addEventListener('input', () => {
  searchQuery = searchInput.value.trim();
  clearBtn.classList.toggle('visible', searchQuery.length > 0);
  visibleCount = 12;
  renderGrid();
});
clearBtn.addEventListener('click', () => {
  searchInput.value = ''; searchQuery = '';
  clearBtn.classList.remove('visible');
  visibleCount = 12; renderGrid();
});
document.querySelectorAll('.ind-pill').forEach(pill => {
  pill.addEventListener('click', () => {
    document.querySelectorAll('.ind-pill').forEach(p => p.classList.remove('active'));
    pill.classList.add('active');
    activeIndustry = pill.dataset.ind;
    visibleCount = 12; renderGrid();
  });
});
document.getElementById('sortSelect').addEventListener('change', function() {
  sortMode = this.value; renderGrid();
});
document.getElementById('loadMoreBtn').addEventListener('click', () => {
  visibleCount += 12; renderGrid();
});

// ── ALSO UPDATE SEARCH DESCRIPTION to reflect real data count ─
document.querySelector('#search .section-sub').textContent =
  `Search all ${COMPANIES.length.toLocaleString()} Irish permit-issuing employers from the 2026 DETE data — filter by industry`;

// ── INIT ──────────────────────────────────────────────────────
function init() {
  updateKPIs(selectedYear);
  buildDonut(selectedYear);
  buildSectorBar(selectedYear);
  buildTimeline();
  renderEventPills();
  renderRaceChart(raceYears[0]);
  buildRacePlay();
  buildStackedArea();
  buildNatChart();
  buildNatBubbles();
  buildNatToggles();
  buildNatTrendChart();
  buildCountyChart();
  buildExplorerChart();
  renderGrid();
}

init();
