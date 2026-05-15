/* ═══════════════════════════════════════════════════════
   SOMATRIN — Production JS
   ═══════════════════════════════════════════════════════ */

/* ─── Chart helpers (existing) ──────────────────────── */
function somaReadJsonScript(id) {
  const node = document.getElementById(id);
  if (!node) return null;
  try { return JSON.parse(node.textContent); } catch (e) { return null; }
}

function somaLineChart(canvasId, labels, values, label) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return;
  new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [{ label, data: values, borderColor: '#1a2c4e', backgroundColor: 'rgba(26,44,78,0.15)', tension: 0.3, fill: true }],
    },
  });
}

function somaBarChart(canvasId, labels, values, label, color) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return;
  new Chart(canvas, {
    type: 'bar',
    data: { labels, datasets: [{ label, data: values, backgroundColor: color || '#e87722', borderRadius: 6 }] },
  });
}

function somaPieChart(canvasId, labels, values) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return;
  new Chart(canvas, {
    type: 'pie',
    data: { labels, datasets: [{ data: values, backgroundColor: ['#1a2c4e', '#e87722', '#2f5597'] }] },
  });
}

function somaNormalizeCellValue(text) {
  if (text == null) return "";
  return String(text).replace(/\s+/g, " ").trim();
}

function somaSortTable(table, columnIndex, direction) {
  const tbody = table.querySelector("tbody");
  if (!tbody) return;
  const rows = Array.from(tbody.querySelectorAll("tr")).filter((r) => !r.classList.contains("no-sort"));
  const sorted = rows.sort((a, b) => {
    const av = somaNormalizeCellValue(a.children[columnIndex]?.innerText || "");
    const bv = somaNormalizeCellValue(b.children[columnIndex]?.innerText || "");
    const an = Number(av.replace(/[^\d.-]/g, ""));
    const bn = Number(bv.replace(/[^\d.-]/g, ""));
    const bothNumeric = !Number.isNaN(an) && !Number.isNaN(bn) && av !== "" && bv !== "";
    let cmp = 0;
    if (bothNumeric) cmp = an - bn;
    else cmp = av.localeCompare(bv, "fr", { sensitivity: "base" });
    return direction === "asc" ? cmp : -cmp;
  });
  sorted.forEach((row) => tbody.appendChild(row));
}

function somaBuildMobileCards(table) {
  const shell = table.closest(".prod-table-shell");
  if (!shell) return;
  let cards = shell.querySelector(".prod-mobile-cards");
  if (!cards) {
    cards = document.createElement("div");
    cards.className = "prod-mobile-cards";
    shell.appendChild(cards);
  }
  cards.innerHTML = "";
  const headers = Array.from(table.querySelectorAll("thead th")).map((th) => somaNormalizeCellValue(th.innerText));
  const rows = Array.from(table.querySelectorAll("tbody tr"));
  rows.forEach((row) => {
    const cells = Array.from(row.querySelectorAll("td"));
    if (!cells.length) return;
    const card = document.createElement("article");
    card.className = "prod-mobile-card";
    const head = document.createElement("div");
    head.className = "prod-mobile-head";
    head.innerHTML = `<span>${somaNormalizeCellValue(cells[0]?.innerText || "")}</span><span>${somaNormalizeCellValue(cells[1]?.innerText || "")}</span>`;
    card.appendChild(head);
    cells.slice(2).forEach((cell, idx) => {
      const line = document.createElement("div");
      line.className = "prod-mobile-line";
      line.innerHTML = `<span class="label">${headers[idx + 2] || ""}</span><span class="value">${somaNormalizeCellValue(cell.innerText)}</span>`;
      card.appendChild(line);
    });
    cards.appendChild(card);
  });
}

function initProductionTables() {
  const tables = document.querySelectorAll(".prod-data-table");
  if (!tables.length) return;

  tables.forEach((table) => {
    const searchInput = table.closest(".prod-table-shell")?.querySelector(".prod-table-search");
    const clearBtn = table.closest(".prod-table-shell")?.querySelector(".prod-clear-search");
    const sortableHeaders = table.querySelectorAll("th.prod-sortable");

    sortableHeaders.forEach((th, idx) => {
      th.addEventListener("click", () => {
        const current = th.dataset.sortDir || "none";
        const next = current === "asc" ? "desc" : "asc";
        sortableHeaders.forEach((h) => {
          h.dataset.sortDir = "none";
          h.classList.remove("sort-asc", "sort-desc");
        });
        th.dataset.sortDir = next;
        th.classList.add(next === "asc" ? "sort-asc" : "sort-desc");
        somaSortTable(table, idx, next);
        somaBuildMobileCards(table);
      });
    });

    if (searchInput) {
      searchInput.addEventListener("input", () => {
        const q = somaNormalizeCellValue(searchInput.value).toLowerCase();
        table.querySelectorAll("tbody tr").forEach((row) => {
          const txt = somaNormalizeCellValue(row.innerText).toLowerCase();
          row.style.display = txt.includes(q) ? "" : "none";
        });
        somaBuildMobileCards(table);
      });
    }

    if (clearBtn && searchInput) {
      clearBtn.addEventListener("click", () => {
        searchInput.value = "";
        searchInput.dispatchEvent(new Event("input"));
      });
    }

    somaBuildMobileCards(table);
  });
}

document.addEventListener("DOMContentLoaded", initProductionTables);

/* Backward compatible helper used by existing templates */
function initProdTable(config) {
  const table = document.getElementById(config.tableId);
  if (!table) return;
  const rowsPerPageSelect = config.rppId ? document.getElementById(config.rppId) : null;
  const searchInput = config.searchId ? document.getElementById(config.searchId) : null;
  const countNode = config.countId ? document.getElementById(config.countId) : null;
  const pagerNode = config.pagerId ? document.getElementById(config.pagerId) : null;

  const allRows = Array.from(table.querySelectorAll("tbody tr[data-row], tbody tr")).filter((r) => !r.querySelector(".pt-empty"));
  const sortableHeaders = table.querySelectorAll("th.pt-sort");
  let pageSize = Number(config.pageSize || 25);
  let page = 1;
  let query = "";

  function filteredRows() {
    if (!query) return allRows;
    return allRows.filter((r) => somaNormalizeCellValue(r.innerText).toLowerCase().includes(query));
  }

  function renderPager(totalPages) {
    if (!pagerNode) return;
    pagerNode.innerHTML = "";
    const mkBtn = (label, p, disabled, active) => {
      const b = document.createElement("button");
      b.className = "prod-page-btn" + (active ? " pt-active" : "") + (disabled ? " pt-disabled" : "");
      b.innerHTML = label;
      b.disabled = disabled;
      if (!disabled) b.addEventListener("click", () => { page = p; refresh(); });
      return b;
    };
    pagerNode.appendChild(mkBtn('<i class="bi bi-chevron-left"></i>', Math.max(1, page - 1), page <= 1, false));
    for (let p = 1; p <= totalPages; p += 1) {
      if (p === page || (p >= page - 2 && p <= page + 2)) pagerNode.appendChild(mkBtn(String(p), p, false, p === page));
    }
    pagerNode.appendChild(mkBtn('<i class="bi bi-chevron-right"></i>', Math.min(totalPages, page + 1), page >= totalPages, false));
  }

  function refresh() {
    const rows = filteredRows();
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / pageSize));
    if (page > totalPages) page = totalPages;
    const start = (page - 1) * pageSize;
    const end = pageSize >= 9999 ? total : start + pageSize;

    allRows.forEach((r) => { r.style.display = "none"; });
    rows.slice(start, end).forEach((r) => { r.style.display = ""; });
    if (countNode) {
      const startDisplay = total ? start + 1 : 0;
      const endDisplay = Math.min(end, total);
      countNode.innerHTML = `Affichage ${startDisplay}-${endDisplay} de ${total} lignes`;
    }
    renderPager(totalPages);
    somaBuildMobileCards(table);
  }

  if (searchInput) {
    searchInput.addEventListener("input", () => {
      query = somaNormalizeCellValue(searchInput.value).toLowerCase();
      page = 1;
      refresh();
    });
    const clear = searchInput.parentElement?.querySelector(".prod-search-clear, .prod-clear-search");
    if (clear) {
      clear.addEventListener("click", () => {
        searchInput.value = "";
        query = "";
        refresh();
      });
    }
  }

  if (rowsPerPageSelect) {
    rowsPerPageSelect.value = String(pageSize);
    rowsPerPageSelect.addEventListener("change", () => {
      pageSize = Number(rowsPerPageSelect.value || 25);
      page = 1;
      refresh();
    });
  }

  sortableHeaders.forEach((th) => {
    th.addEventListener("click", () => {
      const col = Number(th.dataset.col || 0);
      const nextDir = th.dataset.sortDir === "asc" ? "desc" : "asc";
      sortableHeaders.forEach((h) => { h.dataset.sortDir = ""; h.classList.remove("sort-asc", "sort-desc"); });
      th.dataset.sortDir = nextDir;
      th.classList.add(nextDir === "asc" ? "sort-asc" : "sort-desc");
      const rows = allRows.slice().sort((a, b) => {
        const av = somaNormalizeCellValue(a.children[col]?.innerText || "");
        const bv = somaNormalizeCellValue(b.children[col]?.innerText || "");
        const an = Number(av.replace(/[^\d.-]/g, ""));
        const bn = Number(bv.replace(/[^\d.-]/g, ""));
        const bothNumeric = !Number.isNaN(an) && !Number.isNaN(bn) && av !== "" && bv !== "";
        const cmp = bothNumeric ? an - bn : av.localeCompare(bv, "fr", { sensitivity: "base" });
        return nextDir === "asc" ? cmp : -cmp;
      });
      const tbody = table.querySelector("tbody");
      rows.forEach((r) => tbody.appendChild(r));
      page = 1;
      refresh();
    });
  });

  refresh();
}

/* ═══════════════════════════════════════════════════════
   Production Table: Sort · Search · Paginate · Mobile
   ═══════════════════════════════════════════════════════
   Usage:
     initProdTable({
       tableId:  'myTable',
       searchId: 'mySearch',   // optional
       countId:  'myCount',    // optional
       pagerId:  'myPager',    // optional
       rppId:    'myRpp',      // optional  (rows-per-page <select>)
       pageSize: 25            // default
     });
   ─────────────────────────────────────────────────────── */
function initProdTable(opts) {
  var pageSize = opts.pageSize || 25;
  var table = document.getElementById(opts.tableId);
  if (!table) return;

  var allRows = Array.from(table.querySelectorAll('tbody tr[data-row]'));
  var visible  = allRows.slice();
  var page     = 1;
  var sortCol  = -1;
  var sortDir  = 1; // 1 = asc, -1 = desc

  /* ── Column sorting ── */
  var ths = Array.from(table.querySelectorAll('thead th.pt-sort'));
  ths.forEach(function(th) {
    th.addEventListener('click', function() {
      var colIdx = parseInt(th.getAttribute('data-col') !== null ? th.getAttribute('data-col') : th.cellIndex, 10);
      if (sortCol === colIdx) { sortDir *= -1; }
      else { sortCol = colIdx; sortDir = 1; }
      ths.forEach(function(h) { h.classList.remove('pt-asc', 'pt-desc'); });
      th.classList.add(sortDir === 1 ? 'pt-asc' : 'pt-desc');
      visible.sort(function(a, b) {
        var ca = a.cells[colIdx], cb = b.cells[colIdx];
        var va = ca ? ca.textContent.trim() : '';
        var vb = cb ? cb.textContent.trim() : '';
        var na = parseFloat(va.replace(/\s/g, '').replace(',', '.'));
        var nb = parseFloat(vb.replace(/\s/g, '').replace(',', '.'));
        if (!isNaN(na) && !isNaN(nb)) return sortDir * (na - nb);
        return sortDir * va.localeCompare(vb, 'fr');
      });
      page = 1;
      render();
    });
  });

  /* ── Search ── */
  var searchInput = opts.searchId ? document.getElementById(opts.searchId) : null;
  var clearBtn = searchInput ? searchInput.parentElement.querySelector('.prod-search-clear') : null;

  if (searchInput) {
    searchInput.addEventListener('input', function() {
      if (clearBtn) clearBtn.style.display = searchInput.value ? 'block' : 'none';
      applyFilter();
    });
  }

  if (clearBtn) {
    clearBtn.addEventListener('click', function() {
      searchInput.value = '';
      clearBtn.style.display = 'none';
      applyFilter();
    });
  }

  /* ── Rows per page ── */
  var rppSel = opts.rppId ? document.getElementById(opts.rppId) : null;
  if (rppSel) {
    rppSel.addEventListener('change', function() {
      pageSize = parseInt(rppSel.value, 10) || 25;
      page = 1;
      render();
    });
  }

  /* ── Filter logic ── */
  function applyFilter() {
    var q = searchInput ? searchInput.value.toLowerCase().trim() : '';
    visible = allRows.filter(function(tr) {
      if (!q) return true;
      return tr.textContent.toLowerCase().indexOf(q) !== -1;
    });
    page = 1;
    render();
  }

  /* ── Render ── */
  function render() {
    var total = visible.length;
    var pages = Math.max(1, Math.ceil(total / pageSize));
    if (page > pages) page = pages;

    allRows.forEach(function(tr) { tr.style.display = 'none'; });
    visible.slice((page - 1) * pageSize, page * pageSize).forEach(function(tr) { tr.style.display = ''; });

    /* Count label */
    var countEl = opts.countId ? document.getElementById(opts.countId) : null;
    if (countEl) {
      var start = total ? (page - 1) * pageSize + 1 : 0;
      var end   = Math.min(page * pageSize, total);
      countEl.innerHTML = '<strong>' + start + '–' + end + '</strong> de <strong>' + total + '</strong> ligne(s)';
    }

    /* Pager */
    var pagerEl = opts.pagerId ? document.getElementById(opts.pagerId) : null;
    if (!pagerEl) return;
    pagerEl.innerHTML = '';

    buildPager(pagerEl, page, pages);
  }

  function buildPager(el, cur, total) {
    /* Prev */
    var prev = mkBtn('<i class="bi bi-chevron-left"></i>', cur === 1);
    prev.onclick = function() { if (cur > 1) { page = cur - 1; render(); } };
    el.appendChild(prev);

    /* Page numbers */
    pageRange(cur, total).forEach(function(p) {
      if (p === '…') {
        var sp = document.createElement('span');
        sp.textContent = '…';
        sp.style.cssText = 'padding:0 4px;color:#9CA3AF;align-self:center;font-size:12px;';
        el.appendChild(sp);
      } else {
        var btn = mkBtn(p, false, p === cur);
        (function(pg) {
          btn.onclick = function() { page = pg; render(); };
        })(p);
        el.appendChild(btn);
      }
    });

    /* Next */
    var next = mkBtn('<i class="bi bi-chevron-right"></i>', cur === total);
    next.onclick = function() { if (cur < total) { page = cur + 1; render(); } };
    el.appendChild(next);
  }

  function mkBtn(html, disabled, active) {
    var btn = document.createElement('button');
    btn.className = 'prod-page-btn' +
      (active   ? ' pt-active'   : '') +
      (disabled ? ' pt-disabled' : '');
    btn.innerHTML = String(html);
    if (disabled) btn.disabled = true;
    return btn;
  }

  function pageRange(cur, total) {
    if (total <= 7) {
      var r = [];
      for (var i = 1; i <= total; i++) r.push(i);
      return r;
    }
    var r = [1];
    if (cur > 3) r.push('…');
    for (var i = Math.max(2, cur - 1); i <= Math.min(total - 1, cur + 1); i++) r.push(i);
    if (cur < total - 2) r.push('…');
    r.push(total);
    return r;
  }

  /* ── Init ── */
  applyFilter();

  return { refresh: applyFilter };
}

/* ─── Filter badge removal helper ───────────────────── */
function prodRemoveFilter(key) {
  var url = new URL(window.location.href);
  url.searchParams.delete(key);
  window.location.href = url.toString();
}

/* ─── Page navigation helper ─────────────────────────── */
function prodGoToPage(page) {
  var url = new URL(window.location.href);
  url.searchParams.set('page', page);
  window.location.href = url.toString();
}

/* ═══════════════════════════════════════════════════════
   PRODUCTION CHARTS — initChartsProduction + helpers
   ═══════════════════════════════════════════════════════ */

/* Shared SOMATRIN palette */
var SOMA_COLORS = {
  navy:    '#1a2c4e',
  orange:  '#e87722',
  teal:    '#0d9488',
  violet:  '#7c3aed',
  green:   '#10b981',
  red:     '#dc2626',
  yellow:  '#f59e0b',
  blue:    '#2563eb',
  gray:    '#6b7280',
  navyAlpha: 'rgba(26,44,78,0.15)',
  orangeAlpha: 'rgba(232,119,34,0.15)',
};

/* Registry for active Chart.js instances (keyed by canvas ID) */
var _prodChartInstances = {};

function _destroyChart(canvasId) {
  if (_prodChartInstances[canvasId]) {
    _prodChartInstances[canvasId].destroy();
    delete _prodChartInstances[canvasId];
  }
}

function _register(canvasId, instance) {
  _destroyChart(canvasId);
  _prodChartInstances[canvasId] = instance;
  return instance;
}

/* ── Low-level chart factories ──────────────────────── */

function prodLineChart(canvasId, labels, datasets, opts) {
  var canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return null;
  opts = opts || {};
  var chartDatasets = datasets.map(function(ds) {
    return Object.assign({
      tension: 0.35,
      fill: false,
      pointRadius: 4,
      pointHoverRadius: 6,
      borderWidth: 2,
    }, ds);
  });
  return _register(canvasId, new Chart(canvas, {
    type: 'line',
    data: { labels: labels, datasets: chartDatasets },
    options: Object.assign({
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { font: { size: 11 }, color: '#374151' } },
        tooltip: { mode: 'index', intersect: false }
      },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { color: '#6b7280', font: { size: 11 } } },
        y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { color: '#6b7280', font: { size: 11 } },
             beginAtZero: true }
      }
    }, opts.chartOptions || {})
  }));
}

function prodBarChart(canvasId, labels, datasets, opts) {
  var canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return null;
  opts = opts || {};
  var chartDatasets = datasets.map(function(ds) {
    return Object.assign({ borderRadius: 5, borderWidth: 0 }, ds);
  });
  return _register(canvasId, new Chart(canvas, {
    type: opts.horizontal ? 'bar' : 'bar',
    data: { labels: labels, datasets: chartDatasets },
    options: Object.assign({
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: opts.horizontal ? 'y' : 'x',
      plugins: {
        legend: { position: 'top', labels: { font: { size: 11 }, color: '#374151' } },
        tooltip: { mode: 'index', intersect: false }
      },
      scales: {
        x: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { color: '#6b7280', font: { size: 11 } } },
        y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { color: '#6b7280', font: { size: 11 } },
             beginAtZero: true }
      }
    }, opts.chartOptions || {})
  }));
}

function prodDoughnutChart(canvasId, labels, values, colors, opts) {
  var canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return null;
  opts = opts || {};
  var palette = colors || [SOMA_COLORS.navy, SOMA_COLORS.orange, SOMA_COLORS.teal,
                            SOMA_COLORS.violet, SOMA_COLORS.green, SOMA_COLORS.yellow];
  return _register(canvasId, new Chart(canvas, {
    type: opts.pie ? 'pie' : 'doughnut',
    data: {
      labels: labels,
      datasets: [{ data: values, backgroundColor: palette, borderWidth: 2, borderColor: '#fff' }]
    },
    options: Object.assign({
      responsive: true,
      maintainAspectRatio: false,
      cutout: opts.pie ? 0 : '60%',
      plugins: {
        legend: { position: opts.legendPosition || 'right', labels: { font: { size: 11 }, color: '#374151', padding: 12 } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              var total = ctx.dataset.data.reduce(function(a, b) { return a + b; }, 0);
              var pct = total ? ((ctx.parsed / total) * 100).toFixed(1) : '0';
              return ' ' + ctx.label + ': ' + ctx.formattedValue + ' (' + pct + '%)';
            }
          }
        }
      }
    }, opts.chartOptions || {})
  }));
}

function prodRadarChart(canvasId, labels, datasets) {
  var canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return null;
  return _register(canvasId, new Chart(canvas, {
    type: 'radar',
    data: { labels: labels, datasets: datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { font: { size: 11 }, color: '#374151' } } },
      scales: {
        r: {
          grid: { color: 'rgba(0,0,0,0.08)' },
          ticks: { color: '#6b7280', font: { size: 10 }, backdropColor: 'transparent' },
          pointLabels: { color: '#374151', font: { size: 11 } }
        }
      }
    }
  }));
}

/* ── Page-specific chart initializers ───────────────── */

function _initDashboardCharts() {
  /* 1. Production mensuelle (line) */
  var monthly = somaReadJsonScript('data-monthly');
  if (monthly && monthly.labels) {
    prodLineChart('chart-production-mensuelle', monthly.labels,
      [{ label: 'Tonnage (t)', data: monthly.tonnage || [], borderColor: SOMA_COLORS.navy, backgroundColor: SOMA_COLORS.navyAlpha, fill: true },
       { label: 'Objectif (t)', data: monthly.objectif || [], borderColor: SOMA_COLORS.orange, borderDash: [5,4] }]);
  }

  /* 2. Répartition coûts (doughnut) */
  var couts = somaReadJsonScript('data-couts-pie');
  if (couts && couts.labels) {
    prodDoughnutChart('chart-couts-repartition', couts.labels, couts.values);
  }

  /* 3. Gasoil par site (bar) */
  var gasoil = somaReadJsonScript('data-gasoil-sites');
  if (gasoil && gasoil.labels) {
    prodBarChart('chart-gasoil-sites', gasoil.labels,
      [{ label: 'Consommé (L)', data: gasoil.consomme || [], backgroundColor: SOMA_COLORS.orange },
       { label: 'Cible (L)',    data: gasoil.cible || [],    backgroundColor: SOMA_COLORS.navy + '66' }]);
  }

  /* 4. IPC / Rendement (bar) */
  var ipc = somaReadJsonScript('data-ipc-sites');
  if (ipc && ipc.labels) {
    prodBarChart('chart-ipc-sites', ipc.labels,
      [{ label: 'IPC', data: ipc.values || [], backgroundColor: [SOMA_COLORS.navy, SOMA_COLORS.orange, SOMA_COLORS.teal] }]);
  }
}

function _initGasoilCharts() {
  /* Consommation par site (grouped bar) */
  var data = somaReadJsonScript('data-gasoil');
  if (data && data.labels) {
    prodBarChart('chart-gasoil-bar', data.labels,
      [{ label: 'Consommé (L)', data: data.consomme || [], backgroundColor: SOMA_COLORS.orange },
       { label: 'Cible (L)',    data: data.cible || [],    backgroundColor: SOMA_COLORS.navy + '99' }]);
  }

  /* Évolution mensuelle (line) */
  var evo = somaReadJsonScript('data-gasoil-evo');
  if (evo && evo.labels) {
    prodLineChart('chart-gasoil-evo', evo.labels,
      [{ label: 'Consommation (L)', data: evo.values || [], borderColor: SOMA_COLORS.orange,
         backgroundColor: SOMA_COLORS.orangeAlpha, fill: true }]);
  }
}

function _initProductionCharts() {
  /* Tonnage par site (bar) */
  var data = somaReadJsonScript('data-production');
  if (data && data.labels) {
    prodBarChart('chart-tonnage', data.labels,
      [{ label: 'Tonnage (t)',   data: data.tonnage || [],   backgroundColor: SOMA_COLORS.navy },
       { label: 'Rendement (%)', data: data.rendement || [], backgroundColor: SOMA_COLORS.orange, yAxisID: 'y1' }],
      { chartOptions: { scales: { y1: { position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, ticks: { color: SOMA_COLORS.orange } } } } });
  }
}

function _initPointagesCharts() {
  var data = somaReadJsonScript('data-pointages');
  if (data && data.labels) {
    prodBarChart('chart-pointages', data.labels,
      [{ label: 'Heures travaillées', data: data.heures || [], backgroundColor: SOMA_COLORS.navy },
       { label: 'Effectif',           data: data.effectif || [], backgroundColor: SOMA_COLORS.orange }]);
  }
}

function _initMachinesCharts() {
  var data = somaReadJsonScript('data-machines');
  if (data && data.labels) {
    prodBarChart('chart-machines-util', data.labels,
      [{ label: 'Utilisation (%)', data: data.utilisation || [],
         backgroundColor: (data.utilisation || []).map(function(v) {
           return v >= 80 ? SOMA_COLORS.green : v >= 50 ? SOMA_COLORS.yellow : SOMA_COLORS.red;
         }) }],
      { horizontal: true });
  }
}

function _initCoutsCharts() {
  var pie = somaReadJsonScript('data-couts-pie');
  if (pie && pie.labels) {
    prodDoughnutChart('chart-couts-pie', pie.labels, pie.values, null, { pie: false, legendPosition: 'bottom' });
  }

  var evo = somaReadJsonScript('data-couts-evo');
  if (evo && evo.labels) {
    prodLineChart('chart-couts-evo', evo.labels,
      (evo.series || []).map(function(s, i) {
        var colors = [SOMA_COLORS.navy, SOMA_COLORS.orange, SOMA_COLORS.teal, SOMA_COLORS.violet, SOMA_COLORS.green];
        return { label: s.label, data: s.data, borderColor: colors[i % colors.length] };
      }));
  }
}

function _initVentesCharts() {
  var data = somaReadJsonScript('data-ventes');
  if (data && data.labels) {
    prodLineChart('chart-ca-mensuel', data.labels,
      [{ label: 'CA (MAD)', data: data.ca || [], borderColor: SOMA_COLORS.navy,
         backgroundColor: SOMA_COLORS.navyAlpha, fill: true }]);
  }
}

function _initRentabiliteCharts() {
  var data = somaReadJsonScript('data-rentabilite');
  if (data && data.labels) {
    /* CA vs Coûts line */
    prodLineChart('chart-ca-couts', data.labels,
      [{ label: 'CA (MAD)',    data: data.ca || [],    borderColor: SOMA_COLORS.navy, fill: false },
       { label: 'Coûts (MAD)', data: data.couts || [], borderColor: SOMA_COLORS.red,  fill: false }]);
    /* Marge mensuelle bar */
    prodBarChart('chart-marge', data.labels,
      [{ label: 'Marge (%)', data: data.marge || [],
         backgroundColor: (data.marge || []).map(function(v) {
           return v >= 0 ? SOMA_COLORS.green : SOMA_COLORS.red;
         }) }]);
  }

  /* Répartition par site (doughnut) */
  var sites = somaReadJsonScript('data-sites-repartition');
  if (sites && sites.labels) {
    prodDoughnutChart('chart-sites-repartition', sites.labels, sites.ca || []);
  }
}

function _initSitesCharts() {
  var data = somaReadJsonScript('data-sites-radar');
  if (data && data.axes) {
    var siteLabels = data.site_labels || ['LH BENSLIMANE', 'LH OUJDA', 'SME AIT BAHA'];
    prodRadarChart('chart-sites-radar', data.axes,
      [{ label: siteLabels[0],        data: data.site1 || [], borderColor: SOMA_COLORS.navy,   backgroundColor: SOMA_COLORS.navyAlpha },
       { label: siteLabels[1],    data: data.site2 || [], borderColor: SOMA_COLORS.orange, backgroundColor: SOMA_COLORS.orangeAlpha },
       { label: siteLabels[2],     data: data.site3 || [], borderColor: SOMA_COLORS.teal,   backgroundColor: 'rgba(13,148,136,0.12)' }]);
  }

  var bar = somaReadJsonScript('data-sites-bar');
  if (bar && bar.labels) {
    prodBarChart('chart-sites-tonnage', bar.labels,
      [{ label: 'Tonnage (t)',   data: bar.tonnage || [],   backgroundColor: SOMA_COLORS.navy },
       { label: 'Rendement (%)', data: bar.rendement || [], backgroundColor: SOMA_COLORS.orange }]);
  }
}

function _initRatiosCharts() {
  var data = somaReadJsonScript('data-ratios');
  if (data && data.labels) {
    prodLineChart('chart-ratios', data.labels,
      (data.series || []).map(function(s, i) {
        var colors = [SOMA_COLORS.navy, SOMA_COLORS.orange, SOMA_COLORS.teal, SOMA_COLORS.violet];
        return { label: s.label, data: s.data, borderColor: colors[i % colors.length] };
      }));
  }
}

function _initIpcCharts() {
  var data = somaReadJsonScript('data-ipc');
  if (data && data.labels) {
    prodBarChart('chart-ipc', data.labels,
      [{ label: 'IPC', data: data.values || [],
         backgroundColor: [SOMA_COLORS.navy, SOMA_COLORS.orange, SOMA_COLORS.teal] }]);
  }
}

/* ── Page detector ───────────────────────────────────── */
var _PAGE_INIT_MAP = {
  'dashboard':  _initDashboardCharts,
  'gasoil':     _initGasoilCharts,
  'detail':     _initProductionCharts,
  'production': _initProductionCharts,
  'pointages':  _initPointagesCharts,
  'machines':   _initMachinesCharts,
  'couts':      _initCoutsCharts,
  'ventes':     _initVentesCharts,
  'rentabilite':_initRentabiliteCharts,
  'sites':      _initSitesCharts,
  'ratios':     _initRatiosCharts,
  'ipc':        _initIpcCharts,
};

/**
 * Main entry point — auto-detect page from body[data-page] or URL segment,
 * then run the matching chart initializer plus global table setup.
 */
function initChartsProduction() {
  if (typeof Chart === 'undefined') return;

  /* ── Default Chart.js global defaults ── */
  Chart.defaults.font.family = "'Inter', 'Segoe UI', sans-serif";
  Chart.defaults.color = '#6b7280';
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(26,44,78,0.92)';
  Chart.defaults.plugins.tooltip.titleColor = '#fff';
  Chart.defaults.plugins.tooltip.bodyColor = '#e5e7eb';
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 6;

  /* ── Detect current page ── */
  var pageKey = (document.body.dataset.page || '').toLowerCase().trim();
  if (!pageKey) {
    var segs = window.location.pathname.replace(/\/$/, '').split('/');
    pageKey = segs[segs.length - 1] || segs[segs.length - 2] || '';
  }

  /* Run exact match or fuzzy match */
  var initFn = _PAGE_INIT_MAP[pageKey];
  if (!initFn) {
    Object.keys(_PAGE_INIT_MAP).forEach(function(key) {
      if (!initFn && pageKey.indexOf(key) !== -1) initFn = _PAGE_INIT_MAP[key];
    });
  }
  if (initFn) initFn();

  /* ── Also try to init any chart canvas that has data-* attributes inline ── */
  _initInlineCharts();
}

/* ── Inline chart data (canvas with data-chart-type / data-labels / data-values) */
function _initInlineCharts() {
  document.querySelectorAll('canvas[data-chart-type]').forEach(function(canvas) {
    if (_prodChartInstances[canvas.id]) return;
    var type   = canvas.dataset.chartType;
    var labels = _tryParse(canvas.dataset.labels) || [];
    var values = _tryParse(canvas.dataset.values) || [];
    var label  = canvas.dataset.label || '';
    var color  = canvas.dataset.color || SOMA_COLORS.navy;
    if (type === 'line') {
      prodLineChart(canvas.id, labels, [{ label: label, data: values, borderColor: color, backgroundColor: color.replace(')', ',0.15)').replace('rgb', 'rgba') }]);
    } else if (type === 'bar') {
      prodBarChart(canvas.id, labels, [{ label: label, data: values, backgroundColor: color }]);
    } else if (type === 'doughnut' || type === 'pie') {
      prodDoughnutChart(canvas.id, labels, values, null, { pie: type === 'pie' });
    }
  });
}

function _tryParse(str) {
  if (!str) return null;
  try { return JSON.parse(str); } catch (e) { return null; }
}

/* ── selectSite(siteId) ─────────────────────────────── */
/**
 * Filter the current page by site — updates URL param and re-renders.
 * Expects <select id="filter-site"> or buttons with data-site-id.
 */
function selectSite(siteId) {
  var url = new URL(window.location.href);
  if (siteId && siteId !== 'all' && siteId !== '') {
    url.searchParams.set('site', siteId);
  } else {
    url.searchParams.delete('site');
  }
  url.searchParams.delete('page');
  window.location.href = url.toString();
}

/* ── selectMonth(month) ─────────────────────────────── */
/**
 * Filter the current page by month (YYYY-MM) — updates URL param.
 * Expects <select id="filter-month"> or buttons with data-month.
 */
function selectMonth(month) {
  var url = new URL(window.location.href);
  if (month && month !== 'all' && month !== '') {
    url.searchParams.set('mois', month);
  } else {
    url.searchParams.delete('mois');
  }
  url.searchParams.delete('page');
  window.location.href = url.toString();
}

/* ── exportCSV() ────────────────────────────────────── */
/**
 * Export the first visible .prod-data-table to CSV.
 * Falls back to any <table> if no prod-data-table found.
 */
function exportCSV(filename) {
  var table = document.querySelector('.prod-data-table') || document.querySelector('table');
  if (!table) { console.warn('exportCSV: no table found'); return; }
  filename = filename || ('production_export_' + _todayStr() + '.csv');

  var rows = [];

  /* Headers */
  var headers = Array.from(table.querySelectorAll('thead th')).map(function(th) {
    return _csvCell(th.innerText || th.textContent);
  });
  if (headers.length) rows.push(headers.join(';'));

  /* Body (only visible rows) */
  table.querySelectorAll('tbody tr').forEach(function(tr) {
    if (tr.style.display === 'none') return;
    var cells = Array.from(tr.querySelectorAll('td')).map(function(td) {
      return _csvCell(td.innerText || td.textContent);
    });
    if (cells.length) rows.push(cells.join(';'));
  });

  var bom    = '﻿'; /* UTF-8 BOM for Excel */
  var blob   = new Blob([bom + rows.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  var url    = URL.createObjectURL(blob);
  var anchor = document.createElement('a');
  anchor.href     = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  setTimeout(function() { URL.revokeObjectURL(url); }, 2000);
}

function _csvCell(text) {
  var val = String(text || '').replace(/\s+/g, ' ').trim();
  if (val.indexOf(';') !== -1 || val.indexOf('"') !== -1 || val.indexOf('\n') !== -1) {
    val = '"' + val.replace(/"/g, '""') + '"';
  }
  return val;
}

function _todayStr() {
  var d = new Date();
  return d.getFullYear() + '-' +
    String(d.getMonth() + 1).padStart(2, '0') + '-' +
    String(d.getDate()).padStart(2, '0');
}

/* ── printPage() ────────────────────────────────────── */
/**
 * Print the page with clean SOMATRIN header.
 * Adds a temporary print title and removes elements with .no-print.
 */
function printPage(title) {
  var pageTitle = title ||
    (document.querySelector('.prod-page-title, h1, .page-title') || {}).innerText ||
    document.title;

  /* Inject a print-only header */
  var printHeader = document.getElementById('_prod-print-header');
  if (!printHeader) {
    printHeader = document.createElement('div');
    printHeader.id = '_prod-print-header';
    printHeader.style.cssText = 'display:none;position:fixed;top:0;left:0;right:0;padding:8px 16px;' +
      'background:#1a2c4e;color:#fff;font-family:sans-serif;font-size:12px;z-index:9999;' +
      'border-bottom:3px solid #e87722;';
    printHeader.innerHTML =
      '<strong>SOMATRIN</strong> — ' + pageTitle +
      '<span style="float:right">Imprimé le ' + new Date().toLocaleDateString('fr-FR') + '</span>';
    document.body.prepend(printHeader);
  }

  var style = document.getElementById('_prod-print-style');
  if (!style) {
    style = document.createElement('style');
    style.id = '_prod-print-style';
    style.textContent =
      '@media print {' +
      '  #_prod-print-header { display:block !important; }' +
      '  .no-print, .prod-filters-bar, .prod-pager-bar, nav, .sidebar { display:none !important; }' +
      '  .prod-table-shell { overflow:visible !important; }' +
      '  table { font-size:10px; }' +
      '  body { padding-top:48px; }' +
      '}';
    document.head.appendChild(style);
  }

  window.print();
}

/* ── updateCharts(newData) ──────────────────────────── */
/**
 * Update all registered chart instances with new data.
 * newData: { canvasId: { labels, datasets } | { labels, values } }
 */
function updateCharts(newData) {
  if (!newData || typeof Chart === 'undefined') return;
  Object.keys(newData).forEach(function(id) {
    var instance = _prodChartInstances[id];
    if (!instance) return;
    var payload = newData[id];
    if (payload.labels) {
      instance.data.labels = payload.labels;
    }
    if (payload.datasets) {
      payload.datasets.forEach(function(ds, i) {
        if (instance.data.datasets[i]) {
          Object.assign(instance.data.datasets[i], ds);
        }
      });
    } else if (payload.values && instance.data.datasets[0]) {
      instance.data.datasets[0].data = payload.values;
    }
    instance.update('active');
  });
}

/* ── Auto-wire filter controls ──────────────────────── */
function _wireFilters() {
  /* Site selector */
  var siteSel = document.getElementById('filter-site');
  if (siteSel) {
    siteSel.addEventListener('change', function() { selectSite(siteSel.value); });
  }

  /* Month selector */
  var monthSel = document.getElementById('filter-month');
  if (monthSel) {
    monthSel.addEventListener('change', function() { selectMonth(monthSel.value); });
  }

  /* Year selector */
  var yearSel = document.getElementById('filter-year');
  if (yearSel) {
    yearSel.addEventListener('change', function() {
      var url = new URL(window.location.href);
      if (yearSel.value) url.searchParams.set('annee', yearSel.value);
      else url.searchParams.delete('annee');
      url.searchParams.delete('page');
      window.location.href = url.toString();
    });
  }

  /* Site toggle buttons (data-site-id) */
  document.querySelectorAll('[data-site-id]').forEach(function(btn) {
    btn.addEventListener('click', function() { selectSite(btn.dataset.siteId); });
  });

  /* Month toggle buttons (data-month) */
  document.querySelectorAll('[data-month]').forEach(function(btn) {
    btn.addEventListener('click', function() { selectMonth(btn.dataset.month); });
  });

  /* Export CSV buttons */
  document.querySelectorAll('[data-action="export-csv"]').forEach(function(btn) {
    btn.addEventListener('click', function() { exportCSV(btn.dataset.filename); });
  });

  /* Export Excel — triggers CSV with .xlsx filename hint */
  document.querySelectorAll('[data-action="export-excel"]').forEach(function(btn) {
    btn.addEventListener('click', function() {
      exportCSV((btn.dataset.filename || 'export').replace(/\.[^.]+$/, '') + '.csv');
    });
  });

  /* Print buttons */
  document.querySelectorAll('[data-action="print"]').forEach(function(btn) {
    btn.addEventListener('click', function() { printPage(btn.dataset.title); });
  });
}

/* ── KPI number animation ───────────────────────────── */
function _animateKpis() {
  document.querySelectorAll('[data-kpi-value]').forEach(function(el) {
    var target  = parseFloat(el.dataset.kpiValue);
    var decimals = parseInt(el.dataset.kpiDecimals || '0', 10);
    var suffix  = el.dataset.kpiSuffix || '';
    if (isNaN(target)) return;
    var start   = 0;
    var duration = 900;
    var startTime = null;
    function step(ts) {
      if (!startTime) startTime = ts;
      var pct = Math.min((ts - startTime) / duration, 1);
      var ease = 1 - Math.pow(1 - pct, 3);
      el.textContent = (start + ease * (target - start)).toFixed(decimals) + suffix;
      if (pct < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  });
}

/* ── Gauge progress bars ────────────────────────────── */
function _initGauges() {
  document.querySelectorAll('[data-gauge]').forEach(function(el) {
    var pct = Math.min(100, Math.max(0, parseFloat(el.dataset.gauge) || 0));
    var color = pct >= 80 ? SOMA_COLORS.green : pct >= 50 ? SOMA_COLORS.orange : SOMA_COLORS.red;
    el.style.setProperty('--gauge-pct', pct + '%');
    el.style.setProperty('--gauge-color', color);
    /* For conic-gradient gauges */
    var shell = el.querySelector('.gauge-core');
    if (shell) {
      shell.style.background =
        'conic-gradient(' + color + ' ' + pct + '%, #e5e7eb ' + pct + '% 100%)';
    }
    /* For progress-bar gauges */
    var bar = el.querySelector('.gauge-fill, .prod-progress-fill');
    if (bar) {
      bar.style.width = pct + '%';
      bar.style.background = color;
    }
  });
}

/* ── Responsive table toggle ────────────────────────── */
function _handleResize() {
  var isMobile = window.innerWidth <= 767;
  document.querySelectorAll('.prod-data-table').forEach(function(table) {
    var shell = table.closest('.prod-table-shell');
    if (!shell) return;
    var cards = shell.querySelector('.prod-mobile-cards');
    if (cards) cards.style.display = isMobile ? 'block' : 'none';
    table.style.display = isMobile ? 'none' : '';
  });
}

/* ── DOMContentLoaded bootstrap ─────────────────────── */
document.addEventListener('DOMContentLoaded', function() {
  initChartsProduction();
  _wireFilters();
  _animateKpis();
  _initGauges();
  _handleResize();

  var resizeTimer;
  window.addEventListener('resize', function() {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(_handleResize, 150);
  });
});

