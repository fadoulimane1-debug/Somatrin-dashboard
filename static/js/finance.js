/* ===== SOMATRIN — Finance & Comptabilité — Charts & Filters ===== */

const FIN_COLORS = {
  navy:     '#1a2c4e',
  navyDk:   '#0f1d2d',
  navyMd:   '#3d5a7f',
  orange:   '#E87722',
  orangeLt: '#F5A860',
  orangeDk: '#C85E0F',
  gray:     '#E8EBF0',
  grayMd:   '#6B7280',
};

const FIN_PALETTE = [
  FIN_COLORS.navy, FIN_COLORS.orange,
  FIN_COLORS.navyMd, FIN_COLORS.grayMd,
  FIN_COLORS.orangeLt, FIN_COLORS.orangeDk,
];

Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
Chart.defaults.color = '#64748b';

function finPie(canvasId, data, title) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !Object.keys(data).length) return null;
  return new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: Object.keys(data),
      datasets: [{
        data: Object.values(data),
        backgroundColor: FIN_PALETTE,
        borderWidth: 2,
        borderColor: '#fff',
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '60%',
      plugins: {
        legend: { position: 'bottom', labels: { padding: 12, font: { size: 11 } } },
        title: { display: !!title, text: title || '', font: { size: 12, weight: '600' }, color: FIN_COLORS.navy },
        tooltip: {
          callbacks: {
            label: function(c) {
              var total = c.dataset.data.reduce(function(a,b){return a+b;},0);
              var pct = total ? Math.round(c.raw*100/total) : 0;
              return ' '+c.label+': '+c.raw.toLocaleString('fr-FR')+' ('+pct+'%)';
            },
          },
        },
      },
    },
  });
}

function finBar(canvasId, data, title, color, horizontal) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !Object.keys(data).length) return null;
  var colors = Array.isArray(color)
    ? color
    : Object.keys(data).map(function(_, i){ return color || FIN_PALETTE[i % FIN_PALETTE.length]; });
  return new Chart(ctx, {
    type: 'bar',
    data: {
      labels: Object.keys(data),
      datasets: [{ data: Object.values(data), backgroundColor: colors, borderRadius: 5, borderSkipped: false }],
    },
    options: {
      indexAxis: horizontal ? 'y' : 'x',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: !!title, text: title||'', font:{size:12,weight:'600'}, color:FIN_COLORS.navy },
        tooltip: {
          callbacks: { label: function(c){ return ' '+c.raw.toLocaleString('fr-FR')+' MAD'; } },
        },
      },
      scales: {
        x: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
      },
    },
  });
}

function finLine(canvasId, data, title, color) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !Object.keys(data).length) return null;
  var c = color || FIN_COLORS.orange;
  return new Chart(ctx, {
    type: 'line',
    data: {
      labels: Object.keys(data),
      datasets: [{
        data: Object.values(data),
        borderColor: c,
        backgroundColor: c.replace(')', ',.10)').replace('rgb', 'rgba'),
        fill: true, tension: 0.4,
        pointBackgroundColor: c, pointRadius: 4, borderWidth: 2,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: !!title, text: title||'', font:{size:12,weight:'600'}, color:FIN_COLORS.navy },
        tooltip: {
          callbacks: { label: function(c){ return ' '+c.raw.toLocaleString('fr-FR')+' MAD'; } },
        },
      },
      scales: {
        x: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
      },
    },
  });
}

function finBarMulti(canvasId, datasets, labels, title) {
  var ctx = document.getElementById(canvasId);
  if (!ctx) return null;
  return new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: datasets,
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 11 } } },
        title: { display: !!title, text: title||'', font:{size:12,weight:'600'}, color:FIN_COLORS.navy },
        tooltip: {
          callbacks: { label: function(c){ return ' '+c.dataset.label+': '+c.raw.toLocaleString('fr-FR')+' MAD'; } },
        },
      },
      scales: {
        x: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
      },
    },
  });
}

/* ── Gauge (semi-circle, health score) ── */
function finGauge(canvasId, score) {
  var ctx = document.getElementById(canvasId);
  if (!ctx) return null;
  score = Math.max(0, Math.min(100, score || 0));
  var col = score >= 80 ? FIN_COLORS.navy : score >= 60 ? FIN_COLORS.orange : score >= 40 ? FIN_COLORS.orangeDk : '#7A3800';
  return new Chart(ctx, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [score, 100 - score],
        backgroundColor: [col, '#e5e7eb'],
        borderWidth: 0,
        hoverOffset: 0,
      }],
    },
    options: {
      circumference: 180,
      rotation: -90,
      cutout: '68%',
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { enabled: false } },
      animation: { animateRotate: true, duration: 900 },
    },
  });
}

/* ── Line chart with dashed target line ── */
function finLineTarget(canvasId, data, title, pctMode) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !data.labels) return null;
  var target = data.target !== undefined ? data.target : null;
  var datasets = [{
    label: title || 'Valeur',
    data: data.data,
    borderColor: FIN_COLORS.orange,
    backgroundColor: 'rgba(232,119,34,.10)',
    fill: true, tension: 0.4,
    pointBackgroundColor: FIN_COLORS.orange, pointRadius: 4, borderWidth: 2,
  }];
  if (target !== null) {
    datasets.push({
      label: 'Cible',
      data: Array(data.labels.length).fill(target),
      borderColor: FIN_COLORS.navyMd,
      borderDash: [6, 4],
      borderWidth: 2, pointRadius: 0,
      fill: false,
    });
  }
  var suffix = pctMode ? '%' : ' j';
  return new Chart(ctx, {
    type: 'line',
    data: { labels: data.labels, datasets: datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 11 } } },
        title: { display: !!title, text: title || '', font: { size: 12, weight: '600' }, color: FIN_COLORS.navy },
        tooltip: { callbacks: { label: function(c) { return ' ' + c.dataset.label + ': ' + c.raw + suffix; } } },
      },
      scales: {
        x: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 }, callback: function(v) { return v + suffix; } } },
      },
    },
  });
}

/* ── Bar chart with target reference ── */
function finBarTarget(canvasId, data, title) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !data.labels) return null;
  var target = data.target !== undefined ? data.target : null;
  var datasets = [{
    label: 'Réel',
    data: data.data,
    backgroundColor: FIN_COLORS.navy,
    borderRadius: 5, borderSkipped: false,
  }];
  if (target !== null) {
    datasets.push({
      label: 'Cible (' + target + 'j)',
      data: Array(data.labels.length).fill(target),
      type: 'line',
      borderColor: FIN_COLORS.orange,
      borderDash: [6, 4],
      borderWidth: 2, pointRadius: 0,
      fill: false,
    });
  }
  return new Chart(ctx, {
    type: 'bar',
    data: { labels: data.labels, datasets: datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 11 } } },
        title: { display: !!title, text: title || '', font: { size: 12, weight: '600' }, color: FIN_COLORS.navy },
        tooltip: { callbacks: { label: function(c) { return ' ' + c.dataset.label + ': ' + c.raw + ' j'; } } },
      },
      scales: {
        x: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 } } },
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 10 }, callback: function(v) { return v + ' j'; } } },
      },
    },
  });
}

/* ── Generic client-side filter + pagination ── */
function setupFinFilter(opts) {
  /* opts: {tableId, filterId (array of input IDs), pagerEl, countEl, pageSize} */
  var pageSize = opts.pageSize || 50;
  var allRows = Array.from(document.querySelectorAll('#' + opts.tableId + ' tbody tr[data-row]'));
  var visible = allRows.slice();
  var page = 1;

  function filter() {
    visible = allRows.filter(function(tr) {
      return (opts.filters || []).every(function(fn){ return fn(tr); });
    });
    page = 1;
    render();
  }

  function render() {
    var total = visible.length;
    var pages = Math.max(1, Math.ceil(total / pageSize));
    if (page > pages) page = pages;
    allRows.forEach(function(tr){ tr.style.display = 'none'; });
    visible.slice((page-1)*pageSize, page*pageSize).forEach(function(tr){ tr.style.display=''; });
    var cEl = opts.countEl ? document.getElementById(opts.countEl) : null;
    if (cEl) cEl.textContent = total + ' ligne(s) affichée(s)';
    var pEl = opts.pagerEl ? document.getElementById(opts.pagerEl) : null;
    if (pEl) {
      pEl.innerHTML = '';
      if (pages > 1) {
        for (var i = 1; i <= pages; i++) {
          (function(p){
            var btn = document.createElement('button');
            btn.className = 'btn btn-sm ' + (p===page ? 'btn-dark' : 'btn-outline-secondary');
            btn.textContent = p;
            btn.onclick = function(){ page=p; render(); };
            pEl.appendChild(btn);
          })(i);
        }
      }
    }
  }

  filter();
  return { filter: filter, render: render };
}
