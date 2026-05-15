/* ═══════════════════════════════════════════════════════
   SOMATRIN — QHSE Module — JavaScript
   Helpers: charts, table filter, pagination, month selector
   ═══════════════════════════════════════════════════════ */

/* ─── Helpers ────────────────────────────────────────── */
function qhseReadJson(id) {
  const el = document.getElementById(id);
  if (!el) return null;
  try { return JSON.parse(el.textContent); } catch { return null; }
}

const QHSE_COLORS = {
  primary:  '#1a2c4e',
  orange:   '#E87722',
  green:    '#10B981',
  red:      '#DC2626',
  blue:     '#3B82F6',
  yellow:   '#F59E0B',
  gray:     '#9CA3AF',
  teal:     '#0EA5E9',
  purple:   '#8B5CF6',
};
const CHART_PALETTE = [
  '#1a2c4e','#E87722','#10B981','#3B82F6','#F59E0B','#DC2626','#8B5CF6','#0EA5E9','#14B8A6','#F97316'
];

/* ─── Line Chart ─────────────────────────────────────── */
function qhseLineChart(canvasId, labels, values, label, color) {
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  return new Chart(el, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: label || 'Valeur',
        data: values,
        borderColor: color || QHSE_COLORS.primary,
        backgroundColor: (color || QHSE_COLORS.primary) + '22',
        tension: 0.3,
        fill: true,
        pointBackgroundColor: color || QHSE_COLORS.primary,
        pointRadius: 4,
        pointHoverRadius: 6,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });
}

/* ─── Line Chart (multi-dataset) ────────────────────── */
function qhseLineChartMulti(canvasId, labels, datasets) {
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  return new Chart(el, {
    type: 'line',
    data: {
      labels,
      datasets: datasets.map((d, i) => ({
        label: d.label,
        data: d.values,
        borderColor: d.color || CHART_PALETTE[i],
        backgroundColor: (d.color || CHART_PALETTE[i]) + '22',
        tension: 0.3,
        fill: false,
        pointRadius: 4,
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { labels: { font: { size: 11 } } } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });
}

/* ─── Bar Chart ──────────────────────────────────────── */
function qhseBarChart(canvasId, labels, values, label, color, horizontal) {
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  return new Chart(el, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: label || 'Valeur',
        data: values,
        backgroundColor: color || QHSE_COLORS.orange,
        borderRadius: 5,
        borderSkipped: false,
      }],
    },
    options: {
      indexAxis: horizontal ? 'y' : 'x',
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: !horizontal }, ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, grid: { display: horizontal }, ticks: { font: { size: 10 } } },
      },
    },
  });
}

/* ─── Bar Chart Multi ────────────────────────────────── */
function qhseBarChartMulti(canvasId, labels, datasets) {
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  return new Chart(el, {
    type: 'bar',
    data: {
      labels,
      datasets: datasets.map((d, i) => ({
        label: d.label,
        data: d.values,
        backgroundColor: d.color || CHART_PALETTE[i],
        borderRadius: 5,
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { labels: { font: { size: 11 } } } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });
}

/* ─── Area Chart ─────────────────────────────────────── */
function qhseAreaChart(canvasId, labels, values, label, color) {
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  return new Chart(el, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: label || 'Valeur',
        data: values,
        borderColor: color || QHSE_COLORS.orange,
        backgroundColor: (color || QHSE_COLORS.orange) + '3A',
        tension: 0.35,
        fill: true,
        pointRadius: 3,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 10 } } },
        y: { beginAtZero: true, ticks: { font: { size: 10 } } },
      },
    },
  });
}

/* ─── Pie / Doughnut Chart ───────────────────────────── */
function qhsePieChart(canvasId, labels, values, isDoughnut) {
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  return new Chart(el, {
    type: isDoughnut ? 'doughnut' : 'pie',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: CHART_PALETTE.slice(0, labels.length),
        borderWidth: 2,
        borderColor: '#fff',
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { font: { size: 10 }, padding: 8 },
        },
      },
      cutout: isDoughnut ? '55%' : '0%',
    },
  });
}

/* ─── Gauge (half-circle) ────────────────────────────── */
function qhseGaugeChart(canvasId, value, max, label, color) {
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  const pct = Math.min(100, Math.max(0, (value / max) * 100));
  const rest = 100 - pct;
  return new Chart(el, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [pct, rest],
        backgroundColor: [color || QHSE_COLORS.orange, '#f1f3f5'],
        borderWidth: 0,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      circumference: 180,
      rotation: -90,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false },
      },
      cutout: '72%',
    },
    plugins: [{
      id: 'gaugeText',
      afterDraw(chart) {
        const { ctx, chartArea } = chart;
        const cx = (chartArea.left + chartArea.right) / 2;
        const cy = chartArea.bottom - 10;
        ctx.save();
        ctx.font = 'bold 20px sans-serif';
        ctx.fillStyle = '#1a2c4e';
        ctx.textAlign = 'center';
        ctx.fillText(value, cx, cy);
        ctx.font = '11px sans-serif';
        ctx.fillStyle = '#6B7280';
        ctx.fillText(label || '', cx, cy + 16);
        ctx.restore();
      },
    }],
  });
}

/* ─── Table Search + Filter ──────────────────────────── */
function qhseInitTableSearch(inputId, tableId, countId) {
  const input = document.getElementById(inputId);
  const table = document.getElementById(tableId);
  if (!input || !table) return;
  const tbody = table.querySelector('tbody');
  const countEl = document.getElementById(countId);

  function updateCount() {
    const visible = Array.from(tbody.querySelectorAll('tr')).filter(r => r.style.display !== 'none').length;
    if (countEl) countEl.textContent = visible;
  }

  input.addEventListener('input', () => {
    const q = input.value.toLowerCase().trim();
    Array.from(tbody.querySelectorAll('tr')).forEach(row => {
      const text = row.textContent.toLowerCase();
      row.style.display = (!q || text.includes(q)) ? '' : 'none';
    });
    updateCount();
  });
  updateCount();
}

/* ─── Table Sort ─────────────────────────────────────── */
function qhseInitTableSort(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const headers = table.querySelectorAll('thead th[data-col]');
  headers.forEach(th => {
    th.addEventListener('click', () => {
      const col = parseInt(th.dataset.col);
      const asc = th.dataset.dir !== 'asc';
      th.dataset.dir = asc ? 'asc' : 'desc';
      headers.forEach(h => h.querySelector('i') && (h.querySelector('i').className = 'bi bi-chevron-expand'));
      th.querySelector('i') && (th.querySelector('i').className = asc ? 'bi bi-chevron-up' : 'bi bi-chevron-down');
      const tbody = table.querySelector('tbody');
      const rows = Array.from(tbody.querySelectorAll('tr'));
      rows.sort((a, b) => {
        const av = (a.cells[col]?.innerText || '').trim();
        const bv = (b.cells[col]?.innerText || '').trim();
        const an = parseFloat(av.replace(/[^\d.-]/g, ''));
        const bn = parseFloat(bv.replace(/[^\d.-]/g, ''));
        const num = !isNaN(an) && !isNaN(bn);
        let cmp = num ? an - bn : av.localeCompare(bv, 'fr', { sensitivity: 'base' });
        return asc ? cmp : -cmp;
      });
      rows.forEach(r => tbody.appendChild(r));
    });
  });
}

/* ─── Pagination ─────────────────────────────────────── */
function qhseInitPagination(tableId, pageSize, infoId, pagBtnsId) {
  const table   = document.getElementById(tableId);
  const infoEl  = document.getElementById(infoId);
  const pagEl   = document.getElementById(pagBtnsId);
  if (!table) return;
  const tbody = table.querySelector('tbody');
  let currentPage = 1;

  function getAllRows() {
    return Array.from(tbody.querySelectorAll('tr'));
  }
  function getVisibleRows() {
    return getAllRows().filter(r => r.style.display !== 'none');
  }
  function render() {
    const rows = getVisibleRows();
    const total = rows.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    currentPage = Math.min(currentPage, pages);
    const start = (currentPage - 1) * pageSize;
    const end   = start + pageSize;

    getAllRows().forEach(r => { if (r.style.display !== 'none') r.setAttribute('data-page-hidden', ''); });
    rows.forEach((r, i) => {
      r.removeAttribute('data-page-hidden');
      r.style.display = (i >= start && i < end) ? '' : 'none';
    });

    if (infoEl) {
      infoEl.innerHTML = `<strong>${Math.min(end, total)}</strong> / <strong>${total}</strong> lignes`;
    }
    if (pagEl) {
      pagEl.innerHTML = '';
      const addBtn = (txt, page, active, disabled) => {
        const btn = document.createElement('button');
        btn.className = 'pag-btn' + (active ? ' active' : '') + (disabled ? ' disabled' : '');
        btn.innerHTML = txt;
        btn.disabled = disabled;
        btn.addEventListener('click', () => { currentPage = page; render(); });
        pagEl.appendChild(btn);
      };
      addBtn('<i class="bi bi-chevron-left"></i>', currentPage - 1, false, currentPage <= 1);
      const maxBtns = 5;
      let startP = Math.max(1, currentPage - 2);
      let endP   = Math.min(pages, startP + maxBtns - 1);
      if (endP - startP < maxBtns - 1) startP = Math.max(1, endP - maxBtns + 1);
      for (let p = startP; p <= endP; p++) addBtn(p, p, p === currentPage, false);
      addBtn('<i class="bi bi-chevron-right"></i>', currentPage + 1, false, currentPage >= pages);
    }
  }
  render();
  return { refresh: render, goTo: (p) => { currentPage = p; render(); } };
}

/* ─── Month Selector ─────────────────────────────────── */
function qhseInitMonthSelector(containerId, onChange) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.querySelectorAll('.month-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      container.querySelectorAll('.month-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (onChange) onChange(btn.dataset.month);
    });
  });
}

/* ─── Body Zone Hover Highlight ───────────────────────── */
function qhseInitBodyZones() {
  document.querySelectorAll('.body-zone-item').forEach(item => {
    item.addEventListener('mouseenter', () => {
      const zone = item.dataset.zone;
      const svgEl = document.querySelector(`.bz-${zone}`);
      if (svgEl) {
        svgEl.style.fill = '#E87722';
        svgEl.style.opacity = '0.9';
      }
    });
    item.addEventListener('mouseleave', () => {
      document.querySelectorAll('[class^="bz-"]').forEach(el => {
        el.style.fill = '';
        el.style.opacity = '';
      });
    });
  });
}

/* ─── Export table to CSV ────────────────────────────── */
function qhseExportCSV(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const rows = [];
  const headers = Array.from(table.querySelectorAll('thead th')).map(th => `"${th.innerText.replace(/"/g,'""')}"`);
  rows.push(headers.join(','));
  table.querySelectorAll('tbody tr').forEach(tr => {
    const cells = Array.from(tr.querySelectorAll('td')).map(td => `"${td.innerText.replace(/"/g,'""').trim()}"`);
    rows.push(cells.join(','));
  });
  const blob = new Blob(['﻿' + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename || 'export.csv';
  a.click();
}
