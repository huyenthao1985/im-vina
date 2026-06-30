import React, { useEffect, useMemo } from 'react';
import type { DataRow } from '../types';

interface SalesJuneDashboardProps {
  rows: DataRow[];
  theme: 'light' | 'dark';
  lang: 'vi' | 'en' | 'ko';
  onToggleTheme: () => void;
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
}

export const SalesJuneDashboard: React.FC<SalesJuneDashboardProps> = ({
  rows,
  theme,
  lang,
  onToggleTheme,
  setLang
}) => {
  const isDark = theme === 'dark';
  const chartTextColor = isDark ? '#f3f4f6' : '#1e293b';
  const chartGridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';

  // Filter records to June 2026 specifically
  const juneRows = useMemo(() => {
    return rows.filter(r => r.year === 2026 && String(r.month).toUpperCase() === 'JUNE');
  }, [rows]);

  // Compute KPIs
  const kpiSales = useMemo(() => {
    return juneRows.filter(r => r.division === 'sales').reduce((sum, r) => sum + (r.value || 0), 0);
  }, [juneRows]);

  const kpiShipment = useMemo(() => {
    return juneRows.filter(r => r.division === 'shipment').reduce((sum, r) => sum + (r.value || 0), 0);
  }, [juneRows]);

  const kpiProduction = useMemo(() => {
    return juneRows.filter(r => r.division === 'production').reduce((sum, r) => sum + (r.value || 0), 0);
  }, [juneRows]);

  // Format Helper
  const fmt = (n: number) => Math.round(n).toLocaleString('vi-VN');

  // Labels based on language
  const labels = {
    title: lang === 'vi' ? 'Báo cáo doanh số — Tháng 6/2026' : lang === 'en' ? 'Sales Report — June 2026' : '2026년 6월 매출 보고서',
    totalSales: lang === 'vi' ? 'Doanh số (sales)' : lang === 'en' ? 'Sales Amt' : '매출액 (Sales)',
    totalShipment: lang === 'vi' ? 'Sản lượng (shipment)' : lang === 'en' ? 'Shipment Qty' : '출하량 (Shipment)',
    totalProduction: lang === 'vi' ? 'Sản xuất (production)' : lang === 'en' ? 'Production Qty' : '생산량 (Production)',
    topModelsTitle: lang === 'vi' ? 'Top 5 Model theo Doanh số (K$)' : lang === 'en' ? 'Top 5 Models by Sales (K$)' : '매출 Top 5 모델 (K$)',
    customerShareTitle: lang === 'vi' ? 'Tỷ trọng theo Khách hàng' : lang === 'en' ? 'Customer Sales Share' : '고객사별 매출 비중',
    noData: lang === 'vi' ? 'Không có dữ liệu trong Tháng 6/2026' : lang === 'en' ? 'No data available for June 2026' : '2026년 6월 데이터가 없습니다.'
  };

  useEffect(() => {
    if (juneRows.length === 0 || typeof window.Plotly === 'undefined') return;

    // 1. Top 5 Models Bar Chart
    const salesByModel: Record<string, number> = {};
    juneRows.filter(r => r.division === 'sales').forEach(r => {
      const model = r.model || 'Unknown';
      salesByModel[model] = (salesByModel[model] || 0) + r.value;
    });

    const topModels = Object.entries(salesByModel)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);

    const modelEl = document.getElementById('juneModelsChart');
    if (modelEl) {
      if (topModels.length === 0) {
        modelEl.innerHTML = `<div class="panel-empty">${labels.noData}</div>`;
      } else {
        modelEl.innerHTML = '';
        const modelNames = topModels.map(e => e[0]);
        const modelVals = topModels.map(e => e[1]);

        const trace = [{
          x: modelNames,
          y: modelVals,
          type: 'bar',
          marker: {
            color: '#8b5cf6'
          },
          text: modelVals.map(v => Math.round(v).toLocaleString('vi-VN')),
          textposition: 'outside',
          textfont: { color: chartTextColor, weight: 'bold' }
        }];

        const layout = {
          paper_bgcolor: 'rgba(0,0,0,0)',
          plot_bgcolor: 'rgba(0,0,0,0)',
          font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
          margin: { l: 50, r: 16, t: 30, b: 36 },
          xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
          yaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor } },
          hovermode: 'x unified'
        };

        window.Plotly.react('juneModelsChart', trace as any, layout as any, { displayModeBar: false, responsive: true });
      }
    }

    // 2. Customer Share Donut Chart
    const targetCustomers = ['GAOXIN', 'Q-TECH', 'SEMV', 'SUNNY'];
    const salesByCustomer: Record<string, number> = {};
    targetCustomers.forEach(c => { salesByCustomer[c] = 0; });

    juneRows.filter(r => r.division === 'sales').forEach(r => {
      const c = String(r.origin || '').toUpperCase().trim();
      if (targetCustomers.includes(c)) {
        salesByCustomer[c] += r.value;
      }
    });

    const donutLabelsArr = targetCustomers.filter(c => salesByCustomer[c] > 0);
    const donutValuesArr = donutLabelsArr.map(c => salesByCustomer[c]);
    const totalSalesJune = donutValuesArr.reduce((a, b) => a + b, 0);

    const donutEl = document.getElementById('juneCustomerDonut');
    if (donutEl) {
      if (totalSalesJune === 0) {
        donutEl.innerHTML = `<div class="panel-empty">${labels.noData}</div>`;
      } else {
        donutEl.innerHTML = '';
        const customerColors: Record<string, string> = {
          'GAOXIN': '#0891b2',
          'Q-TECH': '#e11d48',
          'SEMV': '#059669',
          'SUNNY': '#7c3aed'
        };
        const colors = donutLabelsArr.map(c => customerColors[c] || '#cbd5e1');

        const traceDonut = [{
          labels: donutLabelsArr,
          values: donutValuesArr,
          type: 'pie',
          hole: 0.58,
          marker: { colors: colors, line: { color: isDark ? '#131026' : '#ffffff', width: 2.5 } },
          textinfo: 'percent',
          textfont: { size: 11, color: '#fff', weight: 'bold' },
          hoverinfo: 'label+value+percent',
          domain: { x: [0, 0.7], y: [0.02, 0.98] }
        }];

        const layoutDonut = {
          paper_bgcolor: 'rgba(0,0,0,0)',
          font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
          margin: { l: 5, r: 5, t: 5, b: 5 },
          showlegend: true,
          legend: { orientation: 'v', font: { size: 10, color: chartTextColor }, x: 0.72, y: 0.5, yanchor: 'middle' },
          annotations: [
            {
              font: { size: 34 },
              showarrow: false,
              text: '<span style="text-shadow: 0 0 8px rgba(251, 191, 36, 0.7)">💡</span>',
              x: 0.35,
              y: 0.5
            }
          ]
        };

        window.Plotly.react('juneCustomerDonut', traceDonut as any, layoutDonut as any, { displayModeBar: false, responsive: true });
      }
    }
  }, [juneRows, chartTextColor, chartGridColor, isDark, labels.noData]);

  return (
    <div className="sales-dashboard second-dashboard" style={{ padding: '24px', boxSizing: 'border-box' }}>
      {/* Header */}
      <div className="hero" style={{ position: 'relative', padding: '16px 20px 4px', margin: '0 0 16px' }}>
        <h1 style={{ fontSize: '26px', margin: '0 0 6px', fontWeight: '700' }}>{labels.title}</h1>
        <div className="header-line" style={{ margin: '0 auto 8px' }}></div>
        
        {/* Toolbar items floating in header */}
        <div style={{ position: 'absolute', top: 0, right: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
          <select value={lang} onChange={e => setLang(e.target.value as any)} className="lang-select">
            <option value="vi">Tiếng Việt</option>
            <option value="en">English</option>
            <option value="ko">한국어</option>
          </select>
          <button
            className="theme-toggle"
            onClick={onToggleTheme}
            style={{ position: 'static' }}
            title={theme === 'dark' ? 'Chuyển sang Sáng' : 'Chuyển sang Tối'}
          >
            {theme === 'dark' ? '☀️' : '🌙'}
          </button>
        </div>
      </div>

      {/* Content wrapper */}
      <div className="dash-container" style={{ animation: 'fadeUp 0.6s ease both' }}>
        {juneRows.length === 0 ? (
          <div className="panel-empty" style={{ padding: '60px', textAlign: 'center', background: 'var(--surface)', border: '1px solid var(--border-soft)', borderRadius: '12px' }}>
            {labels.noData}
          </div>
        ) : (
          <>
            {/* KPI Row (3 cards) */}
            <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(3, 1fr)', gap: '16px', marginBottom: '24px' }}>
              <div className="kpi-card" style={{ '--kpi-bg': 'var(--purple-soft)' } as any}>
                <div className="kpi-card-header">
                  <div className="kpi-card-icon" style={{ background: 'var(--purple-soft)', color: 'var(--purple)' }}>📈</div>
                  <div className="kpi-card-label">{labels.totalSales}</div>
                </div>
                <div className="kpi-card-value">{fmt(kpiSales)} K$</div>
              </div>

              <div className="kpi-card" style={{ '--kpi-bg': 'var(--cyan-soft)' } as any}>
                <div className="kpi-card-header">
                  <div className="kpi-card-icon" style={{ background: 'var(--cyan-soft)', color: 'var(--cyan)' }}>📦</div>
                  <div className="kpi-card-label">{labels.totalShipment}</div>
                </div>
                <div className="kpi-card-value">{fmt(kpiShipment)} Kea</div>
              </div>

              <div className="kpi-card" style={{ '--kpi-bg': 'var(--green-soft)' } as any}>
                <div className="kpi-card-header">
                  <div className="kpi-card-icon" style={{ background: 'var(--green-soft)', color: 'var(--green)' }}>🏭</div>
                  <div className="kpi-card-label">{labels.totalProduction}</div>
                </div>
                <div className="kpi-card-value">{fmt(kpiProduction)} Kea</div>
              </div>
            </div>

            {/* Charts Row */}
            <div className="charts-row row-1" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
              <div className="panel">
                <div className="panel-head">
                  <h3>{labels.topModelsTitle}</h3>
                </div>
                <div className="chart-holder" id="juneModelsChart" style={{ minHeight: '320px' }}></div>
              </div>

              <div className="panel">
                <div className="panel-head">
                  <h3>{labels.customerShareTitle}</h3>
                </div>
                <div className="chart-holder" id="juneCustomerDonut" style={{ minHeight: '320px' }}></div>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
};
