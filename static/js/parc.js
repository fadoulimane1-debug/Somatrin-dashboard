/* ===== SOMATRIN — Parc & Maintenance — Charts & Interactions ===== */

const PARC_COLORS = {
  navy:    '#1a2c4e',
  orange:  '#E87722',
  green:   '#16a34a',
  red:     '#dc2626',
  blue:    '#0ea5e9',
  purple:  '#7c3aed',
  teal:    '#0d9488',
  amber:   '#d97706',
  slate:   '#64748b',
};

const PALETTE = [
  PARC_COLORS.navy, PARC_COLORS.orange, PARC_COLORS.green,
  PARC_COLORS.blue,  PARC_COLORS.purple, PARC_COLORS.teal,
  PARC_COLORS.amber, PARC_COLORS.red,    PARC_COLORS.slate,
];

Chart.defaults.font.family = "'Segoe UI', system-ui, sans-serif";
Chart.defaults.color = '#64748b';

function parcPie(canvasId, data, title) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !Object.keys(data).length) return;
  return new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: Object.keys(data),
      datasets: [{
        data: Object.values(data),
        backgroundColor: PALETTE,
        borderWidth: 2,
        borderColor: '#fff',
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '60%',
      plugins: {
        legend: { position: 'bottom', labels: { padding: 14, font: { size: 12 } } },
        title: { display: !!title, text: title, font: { size: 13, weight: '600' }, color: PARC_COLORS.navy },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              var total = ctx.dataset.data.reduce(function(a, b) { return a + b; }, 0);
              var pct = total ? Math.round(ctx.raw * 100 / total) : 0;
              return ' ' + ctx.label + ': ' + ctx.raw.toLocaleString('fr-FR') + ' (' + pct + '%)';
            },
          },
        },
      },
    },
  });
}

function parcBar(canvasId, data, title, color, horizontal) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !Object.keys(data).length) return;
  var colors = Array.isArray(color)
    ? color
    : Object.keys(data).map(function(_, i) { return color || PALETTE[i % PALETTE.length]; });
  return new Chart(ctx, {
    type: 'bar',
    data: {
      labels: Object.keys(data),
      datasets: [{
        data: Object.values(data),
        backgroundColor: colors,
        borderRadius: 6,
        borderSkipped: false,
      }],
    },
    options: {
      indexAxis: horizontal ? 'y' : 'x',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: !!title, text: title, font: { size: 13, weight: '600' }, color: PARC_COLORS.navy },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              return ' ' + ctx.raw.toLocaleString('fr-FR');
            },
          },
        },
      },
      scales: {
        x: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 11 } } },
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 11 } } },
      },
    },
  });
}

function parcLine(canvasId, data, title) {
  var ctx = document.getElementById(canvasId);
  if (!ctx || !data || !Object.keys(data).length) return;
  return new Chart(ctx, {
    type: 'line',
    data: {
      labels: Object.keys(data),
      datasets: [{
        data: Object.values(data),
        borderColor: PARC_COLORS.orange,
        backgroundColor: 'rgba(232,119,34,0.10)',
        fill: true,
        tension: 0.4,
        pointBackgroundColor: PARC_COLORS.orange,
        pointRadius: 4,
        borderWidth: 2,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        title: { display: !!title, text: title, font: { size: 13, weight: '600' }, color: PARC_COLORS.navy },
      },
      scales: {
        x: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 11 } } },
        y: { grid: { color: '#f1f5f9' }, ticks: { font: { size: 11 } } },
      },
    },
  });
}
