import { useEffect, useRef } from 'react';
import { Chart, registerables } from 'chart.js';
import type { DataRow, ColumnMapping } from '../types';
import { toNumber, CHART_COLORS, parseToDate, fmtVND } from '../utils';

Chart.register(...registerables);

const NONE = '__none__';

interface ChartsSectionProps {
  rows: DataRow[];
  mapping: ColumnMapping;
  theme: string;
}

function chartDefaults(theme: string) {
  const textColor = theme === 'dark' ? '#7068a0' : '#8580aa';
  const gridColor = theme === 'dark' ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.07)';
  return { textColor, gridColor };
}

// Helper to aggregate data
function aggregateBy(rows: DataRow[], groupCol: string, valueCol: string | null): [string, number][] {
  const agg: Record<string, number> = {};
  rows.forEach(r => {
    const key = r[groupCol] === null || r[groupCol] === undefined || r[groupCol] === ''
      ? 'Không xác định'
      : String(r[groupCol]);
    const val = valueCol ? toNumber(r[valueCol]) : 1;
    agg[key] = (agg[key] || 0) + val;
  });
  return Object.entries(agg).sort((a, b) => b[1] - a[1]);
}

function aggregateByDate(rows: DataRow[], dateCol: string, valueCol: string | null, costCol?: string): {
  labels: string[];
  revenueData: number[];
  costData: number[];
  groupByMonth: boolean;
} {
  const parsedRows = rows
    .map(r => {
      const d = parseToDate(r[dateCol]);
      return d ? { ...r, _parsedDate: d } : null;
    })
    .filter((r): r is DataRow & { _parsedDate: Date } => r !== null);

  const distinctDays = new Set(parsedRows.map(r => r._parsedDate.toISOString().slice(0, 10)));
  const groupByMonth = distinctDays.size > 60;

  const revAgg: Record<string, number> = {};
  const costAgg: Record<string, number> = {};
  
  parsedRows.forEach(r => {
    const d = r._parsedDate;
    const key = groupByMonth ? d.toISOString().slice(0, 7) : d.toISOString().slice(0, 10);
    revAgg[key] = (revAgg[key] || 0) + (valueCol ? toNumber(r[valueCol]) : 1);
    if (costCol && costCol !== NONE) {
      costAgg[key] = (costAgg[key] || 0) + toNumber(r[costCol]);
    }
  });

  const labels = Object.keys(revAgg).sort();
  return {
    labels,
    revenueData: labels.map(l => revAgg[l]),
    costData: labels.map(l => costAgg[l] || 0),
    groupByMonth,
  };
}

export const ChartsSection: React.FC<ChartsSectionProps> = ({ rows, mapping, theme }) => {
  const trendRef = useRef<HTMLCanvasElement>(null);
  const doughnutRef = useRef<HTMLCanvasElement>(null);
  const barRef = useRef<HTMLCanvasElement>(null);

  const trendInstance = useRef<Chart | null>(null);
  const doughnutInstance = useRef<Chart | null>(null);
  const barInstance = useRef<Chart | null>(null);

  const hasDate = !!mapping.dateCol && mapping.dateCol !== NONE;
  const hasCategory = !!mapping.categoryCol && mapping.categoryCol !== NONE;
  const hasRevenue = !!mapping.revenueCol && mapping.revenueCol !== NONE;
  const hasCost = !!mapping.costCol && mapping.costCol !== NONE;

  useEffect(() => {
    const { textColor, gridColor } = chartDefaults(theme);

    const destroy = (ref: React.MutableRefObject<Chart | null>) => {
      if (ref.current) { ref.current.destroy(); ref.current = null; }
    };

    // ─── Trend Chart (Line) ───────────────────────────────
    destroy(trendInstance);
    if (hasDate && trendRef.current) {
      const { labels, revenueData, costData } = aggregateByDate(
        rows, mapping.dateCol, hasRevenue ? mapping.revenueCol : null, hasCost ? mapping.costCol : undefined
      );
      const datasets = [];
      if (hasRevenue) {
        datasets.push({
          label: 'Doanh thu',
          data: revenueData,
          borderColor: '#8b5cf6',
          backgroundColor: 'rgba(139,92,246,0.15)',
          fill: true,
          tension: 0.4,
          pointRadius: revenueData.length < 30 ? 4 : 2,
          pointHoverRadius: 6,
          pointBackgroundColor: '#8b5cf6',
          borderWidth: 2,
        });
      }
      if (hasCost && costData.some(v => v > 0)) {
        datasets.push({
          label: 'Chi phí',
          data: costData,
          borderColor: '#f43f5e',
          backgroundColor: 'rgba(244,63,94,0.10)',
          fill: true,
          tension: 0.4,
          pointRadius: costData.length < 30 ? 3 : 1,
          pointHoverRadius: 5,
          pointBackgroundColor: '#f43f5e',
          borderWidth: 2,
          borderDash: [4, 4],
        });
      }

      const ctx = trendRef.current.getContext('2d')!;
      trendInstance.current = new Chart(ctx, {
        type: 'line',
        data: { labels, datasets },
        options: {
          responsive: true, maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: {
              display: datasets.length > 1,
              labels: { color: textColor, font: { family: 'Inter', size: 11 }, boxWidth: 12, padding: 16 },
            },
            tooltip: {
              backgroundColor: theme === 'dark' ? '#1a1630' : '#ffffff',
              titleColor: theme === 'dark' ? '#f0eefb' : '#1a1636',
              bodyColor: theme === 'dark' ? '#b5aed8' : '#4b4570',
              borderColor: theme === 'dark' ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
              borderWidth: 1,
              callbacks: {
                label: (ctx) => ` ${ctx.dataset.label}: ${fmtVND(ctx.parsed.y ?? 0, mapping.currency)}`,
              },
            },
          },
          scales: {
            x: {
              ticks: { color: textColor, maxRotation: 40, font: { size: 11 } },
              grid: { color: gridColor },
            },
            y: {
              ticks: {
                color: textColor, font: { size: 11 },
                callback: (v: any) => {
                  const isUsd = mapping.currency === 'USD';
                  const abs = Math.abs(v);
                  const isNeg = v < 0;
                  let valStr = '';
                  if (abs >= 1e9) valStr = (abs / 1e9).toFixed(1) + (isUsd ? 'B' : ' tỷ');
                  else if (abs >= 1e6) valStr = (abs / 1e6).toFixed(1) + (isUsd ? 'M' : ' tr');
                  else if (abs >= 1e3) valStr = (abs / 1e3).toFixed(0) + 'k';
                  else valStr = String(abs);
                  
                  if (isUsd) {
                    return (isNeg ? '-$' : '$') + valStr;
                  } else {
                    return (isNeg ? '-' : '') + valStr + ' VNĐ';
                  }
                },
              },
              grid: { color: gridColor },
            },
          },
        },
      });
    }

    // ─── Doughnut Chart (Category Split) ──────────────────
    destroy(doughnutInstance);
    if (hasCategory && doughnutRef.current) {
      const entries = aggregateBy(rows, mapping.categoryCol, hasRevenue ? mapping.revenueCol : null).slice(0, 8);
      const ctx = doughnutRef.current.getContext('2d')!;
      doughnutInstance.current = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: entries.map(e => e[0]),
          datasets: [{
            data: entries.map(e => e[1]),
            backgroundColor: CHART_COLORS.map(c => c + 'cc'),
            borderColor: CHART_COLORS,
            borderWidth: 2,
            hoverOffset: 6,
          }],
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          cutout: '60%',
          plugins: {
            legend: {
              position: 'right',
              labels: { color: textColor, font: { family: 'Inter', size: 11 }, boxWidth: 10, padding: 12, pointStyle: 'circle', usePointStyle: true },
            },
            tooltip: {
              backgroundColor: theme === 'dark' ? '#1a1630' : '#ffffff',
              titleColor: theme === 'dark' ? '#f0eefb' : '#1a1636',
              bodyColor: theme === 'dark' ? '#b5aed8' : '#4b4570',
              borderColor: theme === 'dark' ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
              borderWidth: 1,
              callbacks: {
                label: (ctx) => {
                  const valStr = hasRevenue ? fmtVND(ctx.parsed, mapping.currency) : ctx.parsed.toLocaleString('vi-VN');
                  return ` ${ctx.label}: ${valStr} (${((ctx.parsed / (ctx.dataset.data as number[]).reduce((a, b) => a + b, 0)) * 100).toFixed(1)}%)`;
                },
              },
            },
          },
        },
      });
    }

    // ─── Bar Chart (Top Channels by Revenue) ──────────────
    destroy(barInstance);
    if (hasCategory && barRef.current) {
      const allEntries = aggregateBy(rows, mapping.categoryCol, hasRevenue ? mapping.revenueCol : null);
      let top = allEntries.slice(0, 8);
      const rest = allEntries.slice(8);
      if (rest.length) top.push(['Khác', rest.reduce((a, e) => a + e[1], 0)]);

      const ctx = barRef.current.getContext('2d')!;
      barInstance.current = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: top.map(e => e[0]),
          datasets: [{
            data: top.map(e => e[1]),
            backgroundColor: CHART_COLORS.map(c => c + '99'),
            borderColor: CHART_COLORS,
            borderWidth: 1,
            borderRadius: 6,
            borderSkipped: false,
            maxBarThickness: 42,
          }],
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          indexAxis: top.length > 6 ? 'y' : 'x',
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: theme === 'dark' ? '#1a1630' : '#ffffff',
              titleColor: theme === 'dark' ? '#f0eefb' : '#1a1636',
              bodyColor: theme === 'dark' ? '#b5aed8' : '#4b4570',
              borderColor: theme === 'dark' ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)',
              borderWidth: 1,
              callbacks: {
                label: (ctx) => {
                  const v = ctx.parsed[top.length > 6 ? 'x' : 'y'];
                  if (v == null) return ' 0';
                  const valStr = hasRevenue ? fmtVND(v, mapping.currency) : v.toLocaleString('vi-VN');
                  return ` ${valStr}`;
                },
              },
            },
          },
          scales: {
            x: { ticks: { color: textColor, font: { size: 11 }, maxRotation: top.length > 6 ? 0 : 30 }, grid: { color: gridColor } },
            y: { ticks: { color: textColor, font: { size: 11 } }, grid: { color: gridColor } },
          },
        },
      });
    }

    return () => {
      destroy(trendInstance);
      destroy(doughnutInstance);
      destroy(barInstance);
    };
  }, [rows, mapping, theme]);

  return (
    <>
      {/* Row 1: Trend + Doughnut */}
      <div className="charts-row">
        <div className="panel">
          <div className="panel-header">
            <div>
              <div className="panel-title">
                {hasRevenue && hasCost ? 'Doanh thu & Chi phí theo thời gian' : hasRevenue ? 'Doanh thu theo thời gian' : 'Xu hướng theo thời gian'}
              </div>
              <div className="panel-subtitle">
                {hasDate ? 'Nhóm theo ngày / tháng' : 'Không có cột thời gian'}
              </div>
            </div>
          </div>
          {hasDate ? (
            <div className="chart-holder">
              <canvas ref={trendRef} />
            </div>
          ) : (
            <div className="panel-empty">
              <div className="empty-icon">📅</div>
              <span>Chưa chọn cột thời gian</span>
            </div>
          )}
          {hasRevenue && hasCost && (
            <div className="chart-legend">
              <div className="legend-item"><div className="legend-dot" style={{ background: '#8b5cf6' }} />Doanh thu</div>
              <div className="legend-item"><div className="legend-dot" style={{ background: '#f43f5e' }} />Chi phí</div>
            </div>
          )}
        </div>

        <div className="panel">
          <div className="panel-header">
            <div>
              <div className="panel-title">
                {hasRevenue ? `Doanh thu theo ${mapping.categoryCol !== '__none__' ? mapping.categoryCol : 'kênh'}` : 'Phân bổ theo kênh'}
              </div>
              <div className="panel-subtitle">Top 8 kênh</div>
            </div>
          </div>
          {hasCategory ? (
            <div className="chart-holder">
              <canvas ref={doughnutRef} />
            </div>
          ) : (
            <div className="panel-empty">
              <div className="empty-icon">🏷️</div>
              <span>Chưa chọn cột kênh / danh mục</span>
            </div>
          )}
        </div>
      </div>

      {/* Row 2: Bar Chart */}
      <div className="panel" style={{ minHeight: 280 }}>
        <div className="panel-header">
          <div>
            <div className="panel-title">
              {hasRevenue ? `So sánh ${mapping.categoryCol !== '__none__' ? mapping.categoryCol : 'kênh'} theo Doanh thu` : `Số lượng theo ${mapping.categoryCol !== '__none__' ? mapping.categoryCol : 'kênh'}`}
            </div>
            <div className="panel-subtitle">Tất cả kênh / danh mục</div>
          </div>
        </div>
        {hasCategory ? (
          <div className="chart-holder" style={{ minHeight: 220 }}>
            <canvas ref={barRef} />
          </div>
        ) : (
          <div className="panel-empty">
            <div className="empty-icon">📊</div>
            <span>Chưa chọn cột kênh / danh mục</span>
          </div>
        )}
      </div>
    </>
  );
};
