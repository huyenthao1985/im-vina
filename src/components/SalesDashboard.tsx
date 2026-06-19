import React, { useState, useEffect, useMemo } from 'react';
import type { DataRow, ColumnMapping, FilterState } from '../types';
import { translations } from '../translations';
import { supabase } from '../lib/supabase';

interface SalesDashboardProps {
  rows: DataRow[];
  mapping: ColumnMapping;
  theme: string;
  onBack: () => void;
  onToggleTheme: () => void;
  filters: FilterState;
  onFilterChange: (f: FilterState) => void;
  minDate: Date | null;
  maxDate: Date | null;
  allCategories: string[];
  lang: 'vi' | 'en' | 'ko';
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
  onFileSelected: (file: File, wb: any) => void;
}

declare global {
  interface Window {
    Plotly?: any;
    XLSX?: any;
  }
}

export const SalesDashboard: React.FC<SalesDashboardProps> = ({
  rows,
  theme,
  onToggleTheme,
  lang,
  setLang,
  onFileSelected
}) => {
  // Normalize raw rows to matches expected format
  const normalizedRows = useMemo(() => {
    if (rows.length === 0) return [];
    const keys = Object.keys(rows[0]);
    const findKey = (names: string[]) => {
      for (const n of names) {
        const found = keys.find(k => k.toLowerCase() === n.toLowerCase());
        if (found) return found;
      }
      return null;
    };

    const keyModel = findKey(['model']);
    const keyOrigin = findKey(['출하지', 'origin']);
    const keyCustomer = findKey(['custom', 'customer']);
    const keyType = findKey(['type']);
    const keyDivision = findKey(['division']);
    const keyYear = findKey(['year']);
    const keyMonth = findKey(['month']);
    const keyValue = findKey(["q`ty/amt", "q'ty/amt", 'qty/amt', 'qtyamt', 'value', 'val']);

    return rows.map(r => {
      const model = keyModel ? String(r[keyModel] || '').trim() : '';
      const divUpper = keyDivision ? String(r[keyDivision] || '').toUpperCase() : '';
      let division: 'production' | 'shipment' | 'sales' | null = null;
      if (divUpper.includes('PROD')) division = 'production';
      else if (divUpper.includes('SHIP')) division = 'shipment';
      else if (divUpper.includes('SALES')) division = 'sales';

      const year = keyYear ? Number(r[keyYear]) : 0;
      let month = keyMonth ? String(r[keyMonth] || '').trim().toUpperCase() : '';
      if (month === 'JUN') month = 'JUNE';
      if (month === 'JUL') month = 'JULY';
      const value = keyValue ? Number(r[keyValue]) || 0 : 0;

      return {
        model,
        origin: keyOrigin && r[keyOrigin] ? String(r[keyOrigin]).trim() : '',
        customer: keyCustomer && r[keyCustomer] ? String(r[keyCustomer]).trim() : '',
        type: keyType && r[keyType] ? String(r[keyType]).trim() : '',
        division,
        year,
        month,
        value
      };
    }).filter(r => r.model !== '' && r.division !== null && r.year !== 0);
  }, [rows]);

  // Extract filter dimensions
  const years = useMemo(() => {
    return [...new Set(normalizedRows.map(r => r.year))].sort((a, b) => a - b);
  }, [normalizedRows]);

  const origins = useMemo(() => {
    return [...new Set(normalizedRows.map(r => r.origin).filter(Boolean))].sort();
  }, [normalizedRows]);

  const models = useMemo(() => {
    return [...new Set(normalizedRows.map(r => r.model).filter(Boolean))].sort();
  }, [normalizedRows]);

  // Filters State
  const [yearFrom, setYearFrom] = useState<number>(0);
  const [yearTo, setYearTo] = useState<number>(0);
  const [origin, setOrigin] = useState<string>('');
  const [selModel, setSelModel] = useState<string>('');

  const [currentTime, setCurrentTime] = useState<Date>(new Date());
  const [isSaving, setIsSaving] = useState(false);

  const handleSaveToCloud = async () => {
    if (!supabase) return;
    setIsSaving(true);
    try {
      await supabase.from('sales_data').delete().neq('id', '00000000-0000-0000-0000-000000000000');
      
      const BATCH_SIZE = 500;
      for (let i = 0; i < normalizedRows.length; i += BATCH_SIZE) {
        const batch = normalizedRows.slice(i, i + BATCH_SIZE);
        const { error } = await supabase.from('sales_data').insert(batch);
        if (error) {
          console.error("Save error:", error);
          alert('Lỗi lưu: ' + error.message);
          setIsSaving(false);
          return;
        }
      }
      alert('Đã đồng bộ thành công dữ liệu lên Supabase Cloud!');
    } catch (err: any) {
      alert('Lỗi: ' + err.message);
    }
    setIsSaving(false);
  };

  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  const formattedTime = currentTime.toLocaleString(lang === 'vi' ? 'vi-VN' : lang === 'ko' ? 'ko-KR' : 'en-US', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  });

  // Set default years once they are loaded
  useEffect(() => {
    if (years.length > 0) {
      setYearFrom(years[0]);
      setYearTo(years[years.length - 1]);
    }
  }, [years]);

  // Apply filters
  const filteredRecords = useMemo(() => {
    if (years.length === 0) return [];
    const yFrom = yearFrom || years[0];
    const yTo = yearTo || years[years.length - 1];
    return normalizedRows.filter(r =>
      r.year >= Math.min(yFrom, yTo) && r.year <= Math.max(yFrom, yTo) &&
      (!origin || r.origin === origin) &&
      (!selModel || r.model === selModel)
    );
  }, [normalizedRows, years, yearFrom, yearTo, origin, selModel]);

  // Check how many filters are active
  const activeFilterCount = useMemo(() => {
    let count = 0;
    if (origin) count++;
    if (selModel) count++;
    if (years.length > 0) {
      if (yearFrom !== years[0] || yearTo !== years[years.length - 1]) count++;
    }
    return count;
  }, [years, yearFrom, yearTo, origin, selModel]);

  const handleResetFilters = () => {
    if (years.length > 0) {
      setYearFrom(years[0]);
      setYearTo(years[years.length - 1]);
    }
    setOrigin('');
    setSelModel('');
  };

  // Aggregated Values for KPI Cards
  const ttl = useMemo(() => filteredRecords.filter(r => r.month === 'TTL'), [filteredRecords]);
  
  const sumBy = (arr: any[], div: string) => arr.filter(r => r.division === div).reduce((s, r) => s + r.value, 0);
  const totalSales = useMemo(() => sumBy(ttl, 'sales'), [ttl]);
  const totalShipment = useMemo(() => sumBy(ttl, 'shipment'), [ttl]);
  const totalProduction = useMemo(() => sumBy(ttl, 'production'), [ttl]);
  const ratio = totalProduction > 0 ? (totalShipment / totalProduction * 100) : 0;

  // YoY Growth logic
  const growthData = useMemo(() => {
    const salesByYear: Record<number, number> = {};
    ttl.filter(r => r.division === 'sales').forEach(r => {
      salesByYear[r.year] = (salesByYear[r.year] || 0) + r.value;
    });
    const yrs = Object.keys(salesByYear).map(Number).sort((a, b) => a - b);

    // Find YoY growth comparing last two complete years
    const monthsPresentByYear: Record<number, Set<string>> = {};
    filteredRecords.filter(r => r.division === 'sales' && r.month !== 'TTL').forEach(r => {
      if (!monthsPresentByYear[r.year]) monthsPresentByYear[r.year] = new Set();
      monthsPresentByYear[r.year].add(r.month);
    });
    const completeYears = yrs.filter(y => (monthsPresentByYear[y] ? monthsPresentByYear[y].size : 0) >= 12);
    
    let growth: number | null = null;
    let growthYears: string | null = null;
    if (completeYears.length >= 2) {
      const last = completeYears[completeYears.length - 1];
      const prev = completeYears[completeYears.length - 2];
      if (salesByYear[prev] > 0) {
        growth = ((salesByYear[last] - salesByYear[prev]) / salesByYear[prev]) * 100;
        growthYears = lang === 'vi' ? `${last} so với ${prev}` : lang === 'en' ? `${last} vs ${prev}` : `${last} 대비 ${prev}`;
      }
    }
    return { growth, growthYears };
  }, [ttl, filteredRecords, lang]);

  // Format Helper
  const fmt = (n: number) => Math.round(n).toLocaleString('vi-VN');

  const t = translations[lang];

  // KPI card configuration
  const kpis = [
    { label: t.totalSales, value: fmt(totalSales) + ' K$', color: 'var(--purple)', tint: 'var(--purple-soft)', icon: '📈' },
    { label: t.totalShipment, value: fmt(totalShipment) + ' Kea', color: 'var(--cyan)', tint: 'var(--cyan-soft)', icon: '📦' },
    { label: t.totalProduction, value: fmt(totalProduction) + ' Kea', color: 'var(--green)', tint: 'var(--green-soft)', icon: '🏭' },
    { label: t.conversionRatio, value: ratio.toFixed(1) + '%', color: 'var(--rose)', tint: 'var(--rose-soft)', icon: '📊' },
    {
      label: t.salesGrowthYoY,
      value: growthData.growth === null ? 'N/A' : (growthData.growth >= 0 ? '+' : '') + growthData.growth.toFixed(1) + '%',
      color: 'var(--orange)',
      tint: 'var(--amber-soft)',
      icon: '⚡',
      sub: growthData.growthYears ? growthData.growthYears : (growthData.growth === null ? t.insufficientYears : '')
    }
  ];

  // Render Plotly charts dynamically
  useEffect(() => {
    if (filteredRecords.length === 0 || typeof window.Plotly === 'undefined') return;

    // 1a. ALL YEAR Trend Chart Data
    const salesByYear: Record<number, number> = {};
    const shipByYear: Record<number, number> = {};
    
    // Aggregate by year from the TTL rows (month === 'TTL')
    ttl.filter(r => r.division === 'sales').forEach(r => {
      salesByYear[r.year] = (salesByYear[r.year] || 0) + r.value;
    });
    ttl.filter(r => r.division === 'shipment').forEach(r => {
      shipByYear[r.year] = (shipByYear[r.year] || 0) + r.value;
    });
    
    const yrsList = Object.keys(salesByYear).map(Number).sort((a, b) => a - b);
    const chartYears = yrsList.length ? yrsList : Object.keys(shipByYear).map(Number).sort((a, b) => a - b);

    const yShipYear = chartYears.map(y => shipByYear[y] || 0);
    const ySalesYear = chartYears.map(y => salesByYear[y] || 0);

    // Calculate YoY Growth for Year Sales
    const yGrowthYear = chartYears.map((y, idx) => {
      if (idx === 0) return null;
      const prevY = chartYears[idx - 1];
      const prevVal = salesByYear[prevY] || 0;
      const currVal = salesByYear[y] || 0;
      if (prevVal > 0) {
        return currVal / prevVal; // e.g. 1.31, 2.16
      }
      return 0;
    });

    const isDark = theme === 'dark';
    const chartTextColor = isDark ? '#f3f4f6' : '#1e293b';
    const chartGridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';

    // Build traces for ALL YEAR Chart
    const tracesYear = [
      {
        x: chartYears,
        y: yShipYear,
        type: 'bar',
        name: lang === 'vi' ? 'XUẤT HÀNG (K)' : lang === 'en' ? 'SHIPMENT (K)' : '출하 (K)',
        marker: { color: '#2d7f96' },
        yaxis: 'y',
        text: yShipYear.map(v => v > 0 ? Math.round(v).toLocaleString('vi-VN') : ''),
        textposition: 'auto',
        textfont: { color: isDark ? '#ffffff' : '#000000' }
      },
      {
        x: chartYears,
        y: ySalesYear,
        type: 'scatter',
        mode: 'lines+markers+text',
        name: lang === 'vi' ? 'DOANH SỐ (K$)' : lang === 'en' ? 'SALES AMT (K$)' : '매출 (K$)',
        line: { color: '#00a65a', width: 4 },
        marker: { size: 8, color: '#00a65a', symbol: 'circle' },
        yaxis: 'y',
        text: ySalesYear.map(v => v > 0 ? Math.round(v).toLocaleString('vi-VN') : ''),
        textposition: 'top center',
        textfont: { color: chartTextColor, weight: 'bold' }
      },
      {
        x: chartYears,
        y: yGrowthYear,
        type: 'scatter',
        mode: 'lines+markers+text',
        name: lang === 'vi' ? 'Tr.trưởng YoY' : lang === 'en' ? 'YoY Growth' : '전년 대비 증감',
        line: { color: '#f39c12', width: 3, dash: 'dash' },
        marker: { size: 8, color: '#f39c12', symbol: 'circle' },
        yaxis: 'y2',
        text: yGrowthYear.map(v => v !== null && v > 0 ? Math.round(v * 100) + '%' : ''),
        textposition: 'top center',
        textfont: { color: '#e67e22', weight: 'bold' }
      }
    ];

    const layoutYear = {
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
      margin: { l: 54, r: 54, t: 30, b: 36 },
      legend: { orientation: 'h', y: 1.18, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
      xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
      yaxis: {
        title: { text: lang === 'vi' ? 'Xuất hàng / Doanh số' : lang === 'en' ? 'Shipment / Sales' : '출하 / 매출', font: { size: 11, color: chartTextColor } },
        gridcolor: chartGridColor,
        tickfont: { size: 10, color: chartTextColor }
      },
      yaxis2: {
        title: { text: lang === 'vi' ? 'Tăng trưởng YoY' : lang === 'en' ? 'YoY Growth' : '증감', font: { size: 11, color: chartTextColor } },
        overlaying: 'y',
        side: 'right',
        gridcolor: 'rgba(0,0,0,0)',
        tickfont: { size: 10, color: chartTextColor },
        tickformat: '.0%',
        range: [-0.5, 1.5]
      },
      hovermode: 'x unified'
    };

    const maxGrowthY = Math.max(...yGrowthYear.filter(v => v !== null).map(Number), 1.5);
    layoutYear.yaxis2.range = [-0.5, Math.ceil(maxGrowthY * 1.25 * 2) / 2];

    window.Plotly.react('trendChartYear', tracesYear as any, layoutYear as any, { displayModeBar: false, responsive: true });

    // 1b. MONTH Trend Chart Data
    const monthsOrder = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE', 'JULY', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    const salesByMonth: Record<string, number> = {};
    const shipByMonth: Record<string, number> = {};

    filteredRecords.filter(r => r.month !== 'TTL').forEach(r => {
      const m = String(r.month || '').toUpperCase().trim();
      if (r.division === 'sales') {
        salesByMonth[m] = (salesByMonth[m] || 0) + r.value;
      } else if (r.division === 'shipment') {
        shipByMonth[m] = (shipByMonth[m] || 0) + r.value;
      }
    });

    const activeMonths = monthsOrder.filter(m => (salesByMonth[m] || 0) > 0 || (shipByMonth[m] || 0) > 0);

    const yShipMonth = activeMonths.map(m => shipByMonth[m] || 0);
    const ySalesMonth = activeMonths.map(m => salesByMonth[m] || 0);

    const yGrowthMonth = activeMonths.map((m, idx) => {
      if (idx === 0) return null;
      const prevM = activeMonths[idx - 1];
      const prevVal = salesByMonth[prevM] || 0;
      const currVal = salesByMonth[m] || 0;
      if (prevVal > 0) {
        return currVal / prevVal; // e.g. 0.70, 1.14
      }
      return 0;
    });

    const tracesMonth = [
      {
        x: activeMonths,
        y: yShipMonth,
        type: 'bar',
        name: lang === 'vi' ? 'XUẤT HÀNG (K)' : lang === 'en' ? 'SHIPMENT (K)' : '출하 (K)',
        marker: { color: '#2d7f96' },
        yaxis: 'y',
        text: yShipMonth.map(v => v > 0 ? Math.round(v).toLocaleString('vi-VN') : ''),
        textposition: 'auto',
        textfont: { color: isDark ? '#ffffff' : '#000000' }
      },
      {
        x: activeMonths,
        y: ySalesMonth,
        type: 'scatter',
        mode: 'lines+markers+text',
        name: lang === 'vi' ? 'DOANH SỐ (K$)' : lang === 'en' ? 'SALES AMT (K$)' : '매출 (K$)',
        line: { color: '#00a65a', width: 4 },
        marker: { size: 8, color: '#00a65a', symbol: 'circle' },
        yaxis: 'y',
        text: ySalesMonth.map(v => v > 0 ? Math.round(v).toLocaleString('vi-VN') : ''),
        textposition: 'top center',
        textfont: { color: chartTextColor, weight: 'bold' }
      },
      {
        x: activeMonths,
        y: yGrowthMonth,
        type: 'scatter',
        mode: 'lines+markers+text',
        name: lang === 'vi' ? 'Tr.trưởng MoM' : lang === 'en' ? 'MoM Growth' : '전월 대비 증감',
        line: { color: '#f39c12', width: 3, dash: 'dash' },
        marker: { size: 8, color: '#f39c12', symbol: 'circle' },
        yaxis: 'y2',
        text: yGrowthMonth.map(v => v !== null && v > 0 ? Math.round(v * 100) + '%' : ''),
        textposition: 'top center',
        textfont: { color: '#e67e22', weight: 'bold' }
      }
    ];

    const layoutMonth = {
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
      margin: { l: 54, r: 54, t: 30, b: 36 },
      legend: { orientation: 'h', y: 1.18, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
      xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
      yaxis: {
        title: { text: lang === 'vi' ? 'Xuất hàng / Doanh số' : lang === 'en' ? 'Shipment / Sales' : '출하 / 매출', font: { size: 11, color: chartTextColor } },
        gridcolor: chartGridColor,
        tickfont: { size: 10, color: chartTextColor }
      },
      yaxis2: {
        title: { text: lang === 'vi' ? 'Tăng trưởng MoM' : lang === 'en' ? 'Growth' : '증감', font: { size: 11, color: chartTextColor } },
        overlaying: 'y',
        side: 'right',
        gridcolor: 'rgba(0,0,0,0)',
        tickfont: { size: 10, color: chartTextColor },
        tickformat: '.0%',
        range: [-0.5, 1.5]
      },
      hovermode: 'x unified'
    };
    
    const maxGrowthM = Math.max(...yGrowthMonth.filter(v => v !== null).map(Number), 1.5);
    layoutMonth.yaxis2.range = [-0.5, Math.ceil(maxGrowthM * 1.25 * 2) / 2];

    window.Plotly.react('trendChartMonth', tracesMonth as any, layoutMonth as any, { displayModeBar: false, responsive: true });

    // 5. Custom Row 2 Chart 1: ALL YEAR Top Models (Now respects filters)
    const salesByModelAllYears: Record<string, number> = {};
    ttl.filter(r => r.division === 'sales').forEach(r => {
      salesByModelAllYears[r.model] = (salesByModelAllYears[r.model] || 0) + r.value;
    });
    const topModelsAllYearsList = Object.entries(salesByModelAllYears)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8);

    if (topModelsAllYearsList.length === 0) {
      document.getElementById('topModelsAllYear')!.innerHTML = `<div class="panel-empty">${lang === 'vi' ? 'Không có dữ liệu.' : 'No data.'}</div>`;
    } else {
      const modelLabels = topModelsAllYearsList.map(e => e[0]);
      const modelValues = topModelsAllYearsList.map(e => e[1]);
      const tracesModelAll = [
        {
          x: [null], y: [null], type: 'bar',
          name: 'BEST SALES TTL',
          marker: { color: '#0891b2' }
        },
        {
          x: [null], y: [null], type: 'bar',
          name: 'SALES (K$)',
          marker: { color: '#d97706' }
        },
        {
          x: modelLabels,
          y: modelValues,
          type: 'bar',
          marker: {
            color: modelValues.map((_, idx) => idx < 3 ? '#0891b2' : '#d97706')
          },
          text: modelValues.map(v => Math.round(v).toLocaleString('vi-VN')),
          textposition: 'outside',
          textfont: {
            color: modelValues.map((_, idx) => idx < 3 ? (isDark ? '#fb7185' : '#e11d48') : chartTextColor),
            weight: 'bold'
          },
          showlegend: false
        }
      ];
      const layoutModelAll = {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
        font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
        margin: { l: 50, r: 16, t: 30, b: 36 },
        legend: { orientation: 'h', y: 1.18, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
        xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
        yaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor } },
        hovermode: 'x unified'
      };
      window.Plotly.react('topModelsAllYear', tracesModelAll as any, layoutModelAll as any, { displayModeBar: false, responsive: true });
    }

    // 6. Custom Row 2 Chart 2: MONTH Top Models (respects all filters)
    const salesByModelFiltered: Record<string, number> = {};
    filteredRecords.filter(r => r.division === 'sales' && r.month !== 'TTL').forEach(r => {
      salesByModelFiltered[r.model] = (salesByModelFiltered[r.model] || 0) + r.value;
    });
    const topModelsFilteredList = Object.entries(salesByModelFiltered)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8);

    if (topModelsFilteredList.length === 0) {
      document.getElementById('topModelsFiltered')!.innerHTML = `<div class="panel-empty">${lang === 'vi' ? 'Không có dữ liệu.' : 'No data.'}</div>`;
    } else {
      const modelLabels = topModelsFilteredList.map(e => e[0]);
      const modelValues = topModelsFilteredList.map(e => e[1]);
      const bestSalesMonthText = yearFrom === yearTo 
        ? `BEST SALES YR${String(yearFrom).slice(-2)}` 
        : `BEST SALES YR${String(yearFrom).slice(-2)}-${String(yearTo).slice(-2)}`;

      const tracesModelFiltered = [
        {
          x: [null], y: [null], type: 'bar',
          name: bestSalesMonthText,
          marker: { color: '#0891b2' }
        },
        {
          x: [null], y: [null], type: 'bar',
          name: 'SALES (K$)',
          marker: { color: '#d97706' }
        },
        {
          x: modelLabels,
          y: modelValues,
          type: 'bar',
          marker: {
            color: modelValues.map((_, idx) => idx < 3 ? '#0891b2' : '#d97706')
          },
          text: modelValues.map(v => Math.round(v).toLocaleString('vi-VN')),
          textposition: 'outside',
          textfont: {
            color: modelValues.map((_, idx) => idx < 3 ? (isDark ? '#fb7185' : '#e11d48') : chartTextColor),
            weight: 'bold'
          },
          showlegend: false
        }
      ];
      const layoutModelFiltered = {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
        font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
        margin: { l: 50, r: 16, t: 30, b: 36 },
        legend: { orientation: 'h', y: 1.18, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
        xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
        yaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor } },
        hovermode: 'x unified'
      };
      window.Plotly.react('topModelsFiltered', tracesModelFiltered as any, layoutModelFiltered as any, { displayModeBar: false, responsive: true });
    }

    // 7. Custom Row 2 Chart 3: ALL YEAR Customer Sales Stacked Bar
    const targetCustomers = ['GAOXIN', 'Q-TECH', 'SEMV', 'SUNNY'];
    const activeYears = years.filter(y => y >= (yearFrom || years[0]) && y <= (yearTo || years[years.length - 1]));
    
    const salesByYearCust: Record<number, Record<string, number>> = {};
    activeYears.forEach(y => {
      salesByYearCust[y] = {};
      targetCustomers.forEach(c => {
        salesByYearCust[y][c] = 0;
      });
    });

    ttl.filter(r => r.division === 'sales').forEach(r => {
      const y = r.year;
      const c = String(r.origin || '').toUpperCase().trim();
      if (activeYears.includes(y) && targetCustomers.includes(c)) {
        salesByYearCust[y][c] += r.value;
      }
    });

    const customerColors: Record<string, string> = {
      'GAOXIN': '#0891b2',
      'Q-TECH': '#e11d48',
      'SEMV': '#059669',
      'SUNNY': '#7c3aed'
    };

    const tracesCustomerStacked = targetCustomers.map(c => {
      const yVals = activeYears.map(y => salesByYearCust[y][c]);
      return {
        x: activeYears,
        y: yVals,
        name: c,
        type: 'bar',
        marker: { color: customerColors[c] || '#cbd5e1' },
        text: yVals.map(v => v > 0 ? Math.round(v).toLocaleString('vi-VN') : ''),
        textposition: 'inside',
        textfont: { color: '#ffffff', weight: 'bold', size: 9 },
        insidetextanchor: 'middle'
      };
    });

    const layoutCustomerStacked = {
      paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
      font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
      margin: { l: 50, r: 16, t: 30, b: 36 },
      barmode: 'stack',
      legend: { orientation: 'h', y: 1.18, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
      xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
      yaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor } }
    };
    window.Plotly.react('customerSalesStacked', tracesCustomerStacked as any, layoutCustomerStacked as any, { displayModeBar: false, responsive: true });

    // 8. Custom Row 2 Chart 4: MONTH (CUSTOM) Customer Sales Donut
    const customerSalesFiltered: Record<string, number> = {};
    targetCustomers.forEach(c => { customerSalesFiltered[c] = 0; });
    ttl.filter(r => r.division === 'sales').forEach(r => {
      const c = String(r.origin || '').toUpperCase().trim();
      if (targetCustomers.includes(c)) {
        customerSalesFiltered[c] += r.value;
      }
    });

    const donutLabels = targetCustomers.filter(c => customerSalesFiltered[c] > 0);
    const donutValues = donutLabels.map(c => customerSalesFiltered[c]);

    if (donutValues.length === 0) {
      document.getElementById('customerSalesDonut')!.innerHTML = `<div class="panel-empty">${lang === 'vi' ? 'Không có dữ liệu.' : 'No data.'}</div>`;
    } else {
      const donutColors = donutLabels.map(c => customerColors[c] || '#cbd5e1');
      const traceCustomerDonut = [{
        labels: donutLabels,
        values: donutValues,
        type: 'pie',
        hole: 0.6,
        marker: { colors: donutColors, line: { color: isDark ? '#131026' : '#ffffff', width: 2 } },
        textinfo: 'value',
        texttemplate: '%{value:,.0f}',
        textfont: { size: 10, color: '#fff', weight: 'bold' },
        hoverinfo: 'label+value+percent'
      }];
      const layoutCustomerDonut = {
        paper_bgcolor: 'rgba(0,0,0,0)',
        font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
        margin: { l: 6, r: 6, t: 6, b: 6 },
        showlegend: true,
        legend: { orientation: 'v', font: { size: 10, color: chartTextColor } },
        annotations: [
          {
            font: { size: 28 },
            showarrow: false,
            text: '💡',
            x: 0.5,
            y: 0.5
          }
        ]
      };
      window.Plotly.react('customerSalesDonut', traceCustomerDonut as any, layoutCustomerDonut as any, { displayModeBar: false, responsive: true });
    }
  }, [filteredRecords, theme, lang, t]);


  // Detail table data
  const tableRows = useMemo(() => {
    const models: Record<string, any> = {};
    ttl.forEach(r => {
      if (!models[r.model]) {
        models[r.model] = { model: r.model, origin: r.origin, customer: r.customer, type: r.type, production: 0, shipment: 0, sales: 0 };
      }
      models[r.model][r.division!] += r.value;
      if (!models[r.model].origin && r.origin) models[r.model].origin = r.origin;
      if (!models[r.model].customer && r.customer) models[r.model].customer = r.customer;
      if (!models[r.model].type && r.type) models[r.model].type = r.type;
    });
    const rowsList = Object.values(models).map(m => ({
      ...m,
      ratio: m.production > 0 ? (m.shipment / m.production * 100) : 0
    }));
    rowsList.sort((a, b) => b.sales - a.sales);
    return rowsList.slice(0, 15);
  }, [ttl]);

  // Full Database table state and logic
  const [dbSearch, setDbSearch] = useState('');
  const [dbSort, setDbSort] = useState<{ column: string | null; direction: 'asc' | 'desc' | null }>({ column: null, direction: null });
  const [dbPage, setDbPage] = useState(1);
  const [dbPageSize, setDbPageSize] = useState(25);

  const [dbSelModel, setDbSelModel] = useState('');
  const [dbSelYear, setDbSelYear] = useState('');
  const [dbSelMonth, setDbSelMonth] = useState('');
  const [dbSelCustomer, setDbSelCustomer] = useState('');

  const dbUniqueModels = useMemo(() => {
    return [...new Set(filteredRecords.map(r => r.model).filter(Boolean))].sort();
  }, [filteredRecords]);

  const dbUniqueYears = useMemo(() => {
    return [...new Set(filteredRecords.map(r => r.year).filter(Boolean))].sort((a, b) => a - b);
  }, [filteredRecords]);

  const dbUniqueMonths = useMemo(() => {
    const monthOrder: Record<string, number> = {
      'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6, 'JUNE': 6,
      'JUL': 7, 'JULY': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12, 'TTL': 13
    };
    return [...new Set(filteredRecords.map(r => r.month).filter(Boolean))].sort((a, b) => {
      const oa = monthOrder[a.toUpperCase()] || 99;
      const ob = monthOrder[b.toUpperCase()] || 99;
      return oa - ob;
    });
  }, [filteredRecords]);

  const dbUniqueCustomers = useMemo(() => {
    return [...new Set(filteredRecords.map(r => r.customer).filter(Boolean))].sort();
  }, [filteredRecords]);

  const getDivisionLabel = (div: 'production' | 'shipment' | 'sales' | null) => {
    if (!div) return '';
    if (div === 'production') return lang === 'vi' ? 'Sản xuất' : lang === 'en' ? 'Production' : '생산';
    if (div === 'shipment') return lang === 'vi' ? 'Xuất hàng' : lang === 'en' ? 'Shipment' : '출하';
    return lang === 'vi' ? 'Doanh số' : lang === 'en' ? 'Sales' : '매출';
  };

  const dbFiltered = useMemo(() => {
    let result = filteredRecords;

    if (dbSelModel) {
      result = result.filter(r => r.model === dbSelModel);
    }
    if (dbSelYear) {
      result = result.filter(r => String(r.year) === dbSelYear);
    }
    if (dbSelMonth) {
      result = result.filter(r => r.month.toUpperCase() === dbSelMonth.toUpperCase());
    }
    if (dbSelCustomer) {
      result = result.filter(r => r.customer === dbSelCustomer);
    }

    const q = dbSearch.trim().toLowerCase();
    if (!q) return result;

    return result.filter(r => {
      const divLabel = getDivisionLabel(r.division).toLowerCase();
      return (
        r.model.toLowerCase().includes(q) ||
        r.origin.toLowerCase().includes(q) ||
        r.customer.toLowerCase().includes(q) ||
        r.type.toLowerCase().includes(q) ||
        divLabel.includes(q) ||
        String(r.year).includes(q) ||
        r.month.toLowerCase().includes(q)
      );
    });
  }, [filteredRecords, dbSearch, dbSelModel, dbSelYear, dbSelMonth, dbSelCustomer, lang]);

  const dbSorted = useMemo(() => {
    if (!dbSort.column || !dbSort.direction) return dbFiltered;
    const { column, direction } = dbSort;
    return [...dbFiltered].sort((a: any, b: any) => {
      let va = a[column];
      let vb = b[column];

      if (column === 'division') {
        va = getDivisionLabel(a.division);
        vb = getDivisionLabel(b.division);
      }

      const isNumA = typeof va === 'number';
      const isNumB = typeof vb === 'number';

      if (isNumA && isNumB) {
        return direction === 'asc' ? va - vb : vb - va;
      }
      
      const sa = String(va || '').toLowerCase();
      const sb = String(vb || '').toLowerCase();
      return direction === 'asc' ? sa.localeCompare(sb) : sb.localeCompare(sa);
    });
  }, [dbFiltered, dbSort, lang]);

  const dbTotalPages = Math.max(1, Math.ceil(dbSorted.length / dbPageSize));
  const dbSafePage = Math.min(dbPage, dbTotalPages);
  const dbPaged = useMemo(() => {
    const start = (dbSafePage - 1) * dbPageSize;
    return dbSorted.slice(start, start + dbPageSize);
  }, [dbSorted, dbSafePage, dbPageSize]);

  // Reset page when search or filters change
  useEffect(() => {
    setDbPage(1);
  }, [dbSearch, dbSelModel, dbSelYear, dbSelMonth, dbSelCustomer, yearFrom, yearTo, origin, selModel]);

  const handleDbSort = (col: string) => {
    setDbSort(prev => {
      if (prev.column !== col) return { column: col, direction: 'asc' };
      if (prev.direction === 'asc') return { column: col, direction: 'desc' };
      return { column: null, direction: null };
    });
    setDbPage(1);
  };

  const getDbSortIcon = (col: string) => {
    if (dbSort.column !== col) return <span className="sort-icon">⇅</span>;
    return <span className="sort-icon">{dbSort.direction === 'asc' ? '↑' : '↓'}</span>;
  };

  const dbPageNums = useMemo(() => {
    const nums: (number | '...')[] = [];
    if (dbTotalPages <= 7) {
      for (let i = 1; i <= dbTotalPages; i++) nums.push(i);
    } else {
      nums.push(1);
      if (dbSafePage > 3) nums.push('...');
      for (let i = Math.max(2, dbSafePage - 1); i <= Math.min(dbTotalPages - 1, dbSafePage + 1); i++) nums.push(i);
      if (dbSafePage < dbTotalPages - 2) nums.push('...');
      nums.push(dbTotalPages);
    }
    return nums;
  }, [dbTotalPages, dbSafePage]);

  const renderShowingRowsText = (start: number, end: number, total: number) => {
    return t.showingRows
      .replace('{start}', String(start))
      .replace('{end}', String(end))
      .replace('{total}', String(total));
  };

  return (
    <div className="sales-dashboard" style={{ position: 'relative', zIndex: 1 }}>
      {/* Centered Hero Header */}
      <div className="hero" style={{ position: 'relative' }}>
        <h1>{t.mainTitleDash}</h1>
        <div className="header-line"></div>
        {/* Floating Theme Toggle and Lang select in top-right area */}
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

      <div className="dash-container">
        {/* Filter bar exactly matching reference layout */}
        <div className="topbar-dash">
          <span className="pill-rows">{formattedTime}</span>
          <div className="topbar-filters">
            <div className="filter-field">
              <label>{t.fromYear}</label>
              <select value={yearFrom} onChange={e => setYearFrom(Number(e.target.value))}>
                {years.map(y => <option key={y} value={y}>{y}</option>)}
              </select>
            </div>
            <div className="filter-field">
              <label>{t.toYear}</label>
              <select value={yearTo} onChange={e => setYearTo(Number(e.target.value))}>
                {years.map(y => <option key={y} value={y}>{y}</option>)}
              </select>
            </div>
            <div className="filter-field">
              <label>{t.partner}</label>
              <select value={origin} onChange={e => setOrigin(e.target.value)}>
                <option value="">{t.allOption}</option>
                {origins.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
            <div className="filter-field">
              <label>{t.colModel}</label>
              <select value={selModel} onChange={e => setSelModel(e.target.value)}>
                <option value="">{t.allOption}</option>
                {models.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
            </div>
            {activeFilterCount > 0 && (
              <span className="filter-badge">{activeFilterCount} {t.filtersApplied}</span>
            )}
            <button className="btn-text" onClick={handleResetFilters}>{t.resetBtn}</button>
          </div>
          
          <div style={{ display: 'flex', gap: '8px' }}>
            {supabase && (
              <button 
                className="btn-outline" 
                type="button" 
                onClick={handleSaveToCloud} 
                disabled={isSaving}
                style={{ borderColor: 'var(--green)', color: 'var(--green)' }}
              >
                {isSaving ? '⏳ Đang lưu...' : t.saveToCloudBtn}
              </button>
            )}
            <label className="btn-outline" style={{ cursor: 'pointer', margin: 0, display: 'inline-flex', alignItems: 'center', gap: '8px', color: 'var(--text)', border: '1px solid var(--border)' }}>
              <input 
                type="file" 
                accept=".xlsx, .xls" 
                style={{ display: 'none' }} 
                onChange={async (e) => {
                  const file = e.target.files?.[0];
                  if (!file) return;
                  const reader = new FileReader();
                  reader.onload = async (evt) => {
                    try {
                      const data = new Uint8Array(evt.target!.result as ArrayBuffer);
                      // Import XLSX dynamically or expect it to be available globally/passed
                      // Wait, we need to import XLSX in SalesDashboard.tsx!
                      const XLSX = await import('xlsx');
                      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                      onFileSelected(file, workbook);
                    } catch (err) {
                      alert('Error reading Excel file');
                    }
                  };
                  reader.readAsArrayBuffer(file);
                  e.target.value = '';
                }} 
              />
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15">
                <path d="M21 12a9 9 0 0 1-9 9c-2.52 0-4.93-1-6.74-2.74L3 16" />
                <path d="M3 12a9 9 0 0 1 9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
                <path d="M3 16v5h5" />
                <path d="M16 3h5v5" />
              </svg>
              {t.loadExcelBtn}
            </label>
          </div>
        </div>

        {/* KPI Grid */}
        <div className="kpi-grid" style={{ marginBottom: '22px' }}>
          {kpis.map((c, i) => (
            <div className="kpi-card" key={i} style={{
              border: `1px solid ${c.tint}`,
              borderLeft: `4px solid ${c.color}`,
              background: `linear-gradient(135deg, ${c.tint} 0%, rgba(30, 41, 59, 0.4) 100%)`,
              animationDelay: `${i * 0.05}s`
            }}>
              <div className="kpi-card-header">
                <div className="kpi-card-icon" style={{ background: c.tint, color: c.color, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '34px', height: '34px', borderRadius: '10px', fontSize: '18px', flexShrink: 0 }}>{c.icon}</div>
                <div className="kpi-card-label" style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-2)', textTransform: 'uppercase', letterSpacing: '.04em' }}>{c.label}</div>
              </div>
              <div className="kpi-card-value" style={{ fontFamily: 'monospace', fontSize: '28px', fontWeight: '500', color: 'var(--text-0)' }}>{c.value}</div>
              {c.sub && <div style={{ fontSize: '12px', color: 'var(--text-2)', marginTop: '4px' }}>{c.sub}</div>}
            </div>
          ))}
        </div>

        {/* Charts Row 1 */}
        <div className="charts-row row-1">
          {/* ALL YEAR Trend Chart */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3>{t.chart1Title}</h3>
                <p>{t.chart1Sub}</p>
              </div>
              <span style={{
                background: '#1a1630',
                color: '#ffff00',
                fontWeight: 'bold',
                fontSize: '11px',
                padding: '3px 8px',
                borderRadius: '4px',
                border: '1px solid #ffff00',
                fontFamily: 'monospace',
                letterSpacing: '1px',
                marginTop: '2px'
              }}>
                ALL YEAR
              </span>
            </div>
            <div className="chart-holder" id="trendChartYear"></div>
          </div>

          {/* MONTH Trend Chart */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3>{t.chartMonthTitle}</h3>
                <p>{t.chartMonthSub}</p>
              </div>
              <span style={{
                background: '#1a1630',
                color: '#ffff00',
                fontWeight: 'bold',
                fontSize: '11px',
                padding: '3px 8px',
                borderRadius: '4px',
                border: '1px solid #ffff00',
                fontFamily: 'monospace',
                letterSpacing: '1px',
                marginTop: '2px'
              }}>
                MONTH
              </span>
            </div>
            <div className="chart-holder" id="trendChartMonth"></div>
          </div>
        </div>

        {/* Charts Row 2 */}
        <div className="charts-row row-2">
          {/* Top Models ALL YEAR */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3>{t.chartModelAllYearTitle}</h3>
                <p>{t.chartModelAllYearSub}</p>
              </div>
              <span style={{
                background: '#1a1630',
                color: '#ffff00',
                fontWeight: 'bold',
                fontSize: '11px',
                padding: '3px 8px',
                borderRadius: '4px',
                border: '1px solid #ffff00',
                fontFamily: 'monospace',
                letterSpacing: '1px',
                marginTop: '2px'
              }}>
                ALL YEAR
              </span>
            </div>
            <div className="chart-holder" id="topModelsAllYear"></div>
          </div>

          {/* Top Models MONTH */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3>{t.chartModelFilteredTitle}</h3>
                <p>{t.chartModelFilteredSub}</p>
              </div>
              <span style={{
                background: '#1a1630',
                color: '#ffff00',
                fontWeight: 'bold',
                fontSize: '11px',
                padding: '3px 8px',
                borderRadius: '4px',
                border: '1px solid #ffff00',
                fontFamily: 'monospace',
                letterSpacing: '1px',
                marginTop: '2px'
              }}>
                MONTH
              </span>
            </div>
            <div className="chart-holder" id="topModelsFiltered"></div>
          </div>

          {/* Customer Stacked ALL YEAR (CUSTOM) */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3>{t.chartCustomerStackedTitle}</h3>
                <p>{t.chartCustomerStackedSub}</p>
              </div>
              <span style={{
                background: '#1a1630',
                color: '#ffff00',
                fontWeight: 'bold',
                fontSize: '11px',
                padding: '3px 8px',
                borderRadius: '4px',
                border: '1px solid #ffff00',
                fontFamily: 'monospace',
                letterSpacing: '1px',
                marginTop: '2px'
              }}>
                ALL YEAR (CUSTOM)
              </span>
            </div>
            <div className="chart-holder" id="customerSalesStacked"></div>
          </div>

          {/* Customer Donut MONTH (CUSTOM) */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3>{t.chartCustomerDonutTitle}</h3>
                <p>{t.chartCustomerDonutSub}</p>
              </div>
              <span style={{
                background: '#1a1630',
                color: '#ffff00',
                fontWeight: 'bold',
                fontSize: '11px',
                padding: '3px 8px',
                borderRadius: '4px',
                border: '1px solid #ffff00',
                fontFamily: 'monospace',
                letterSpacing: '1px',
                marginTop: '2px'
              }}>
                MONTH (CUSTOM)
              </span>
            </div>
            <div className="chart-holder" id="customerSalesDonut"></div>
          </div>
        </div>




        {/* Detail Table */}
        <div className="table-panel">
          <div className="table-head">
            <h3>{t.tableTitle}</h3>
          </div>
          <div className="table-scroll">
            <table>
              <thead>
                <tr>
                  <th>{t.colModel}</th><th>{t.colPartner}</th><th>{t.colCustomer}</th><th>{t.colType}</th>
                  <th className="num">{t.colProd}</th><th className="num">{t.colShip}</th><th className="num">{t.colSales}</th><th className="num">{t.colRatio}</th>
                </tr>
              </thead>
              <tbody>
                {tableRows.length === 0 ? (
                  <tr>
                    <td colSpan={8} style={{ textAlign: 'center', padding: '20px', color: 'var(--text-2)', fontFamily: 'var(--font-body)' }}>
                      {t.tableEmpty}
                    </td>
                  </tr>
                ) : (
                  tableRows.map((r, idx) => (
                    <tr key={idx}>
                      <td className="label-cell">{r.model}</td>
                      <td>{r.origin || '—'}</td>
                      <td>{r.customer || '—'}</td>
                      <td>{r.type || '—'}</td>
                      <td className="num">{fmt(r.production)}</td>
                      <td className="num">{fmt(r.shipment)}</td>
                      <td className="num">{fmt(r.sales)}</td>
                      <td className="num">{r.ratio.toFixed(1)}%</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* Full Database Table */}
        <div className="table-panel" style={{ marginTop: '24px' }}>
          <div className="table-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px', marginBottom: '16px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <h3>{t.dbTitle}</h3>
              <span className="pill-rows" style={{ margin: 0, fontSize: '11px', padding: '2px 8px' }}>{formattedTime}</span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              {/* Model Select */}
              <select
                value={dbSelModel}
                onChange={e => { setDbSelModel(e.target.value); setDbPage(1); }}
                className="lang-select"
                style={{ fontSize: '12.5px', padding: '5px 10px', height: '34px' }}
              >
                <option value="">{t.colModel}: {t.allOption}</option>
                {dbUniqueModels.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
              
              {/* Year Select */}
              <select
                value={dbSelYear}
                onChange={e => { setDbSelYear(e.target.value); setDbPage(1); }}
                className="lang-select"
                style={{ fontSize: '12.5px', padding: '5px 10px', height: '34px' }}
              >
                <option value="">{t.colYear}: {t.allOption}</option>
                {dbUniqueYears.map(y => <option key={y} value={y}>{y}</option>)}
              </select>

              {/* Month Select */}
              <select
                value={dbSelMonth}
                onChange={e => { setDbSelMonth(e.target.value); setDbPage(1); }}
                className="lang-select"
                style={{ fontSize: '12.5px', padding: '5px 10px', height: '34px' }}
              >
                <option value="">{t.colMonth}: {t.allOption}</option>
                {dbUniqueMonths.map(m => <option key={m} value={m}>{m}</option>)}
              </select>

              {/* Customer/Custom Select */}
              <select
                value={dbSelCustomer}
                onChange={e => { setDbSelCustomer(e.target.value); setDbPage(1); }}
                className="lang-select"
                style={{ fontSize: '12.5px', padding: '5px 10px', height: '34px' }}
              >
                <option value="">{t.colCustomer}: {t.allOption}</option>
                {dbUniqueCustomers.map(c => <option key={c} value={c}>{c}</option>)}
              </select>

              {/* Search Box */}
              <input
                type="text"
                className="search-box"
                placeholder={t.dbSearchPlace}
                value={dbSearch}
                onChange={e => { setDbSearch(e.target.value); setDbPage(1); }}
                style={{ width: '220px', height: '34px', padding: '6px 12px' }}
              />
            </div>
          </div>
          <div className="table-scroll" style={{ maxHeight: '500px' }}>
            <table>
              <thead>
                <tr>
                  <th className="sortable" onClick={() => handleDbSort('model')}>{t.colModel} {getDbSortIcon('model')}</th>
                  <th className="sortable" onClick={() => handleDbSort('origin')}>{t.colPartner} {getDbSortIcon('origin')}</th>
                  <th className="sortable" onClick={() => handleDbSort('customer')}>{t.colCustomer} {getDbSortIcon('customer')}</th>
                  <th className="sortable" onClick={() => handleDbSort('type')}>{t.colType} {getDbSortIcon('type')}</th>
                  <th className="sortable" onClick={() => handleDbSort('division')}>{t.colDivision} {getDbSortIcon('division')}</th>
                  <th className="sortable num" onClick={() => handleDbSort('year')}>{t.colYear} {getDbSortIcon('year')}</th>
                  <th className="sortable" onClick={() => handleDbSort('month')}>{t.colMonth} {getDbSortIcon('month')}</th>
                  <th className="sortable num" onClick={() => handleDbSort('value')}>{t.colValue} {getDbSortIcon('value')}</th>
                </tr>
              </thead>
              <tbody>
                {dbPaged.length === 0 ? (
                  <tr>
                    <td colSpan={8} style={{ textAlign: 'center', padding: '30px', color: 'var(--text-2)', fontFamily: 'var(--font-body)' }}>
                      {t.tableEmpty}
                    </td>
                  </tr>
                ) : (
                  dbPaged.map((r, idx) => (
                    <tr key={idx}>
                      <td className="label-cell">{r.model}</td>
                      <td>{r.origin || '—'}</td>
                      <td>{r.customer || '—'}</td>
                      <td>{r.type || '—'}</td>
                      <td>{getDivisionLabel(r.division)}</td>
                      <td className="num">{r.year}</td>
                      <td>{r.month}</td>
                      <td className="num" style={{ color: r.division === 'sales' ? 'var(--purple-light)' : r.division === 'shipment' ? 'var(--cyan)' : 'var(--green)' }}>
                        {fmt(r.value)} {r.division === 'sales' ? 'K$' : 'K'}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
          
          {/* Pagination Controls */}
          <div className="pagination-bar" style={{ marginTop: '12px' }}>
            <div className="pagination-info">
              {renderShowingRowsText(
                dbFiltered.length === 0 ? 0 : (dbSafePage - 1) * dbPageSize + 1,
                Math.min(dbSafePage * dbPageSize, dbFiltered.length),
                dbFiltered.length
              )}
            </div>
            <div className="pagination-controls">
              <button className="page-btn" onClick={() => setDbPage(1)} disabled={dbSafePage === 1}>«</button>
              <button className="page-btn" onClick={() => setDbPage(p => Math.max(1, p - 1))} disabled={dbSafePage === 1}>‹</button>
              {dbPageNums.map((n, i) => n === '...'
                ? <span key={`ellipsis-${i}`} style={{ padding: '0 4px', color: 'var(--text-3)' }}>…</span>
                : <button key={n} className={`page-btn ${dbSafePage === n ? 'active' : ''}`} onClick={() => setDbPage(n as number)}>{n}</button>
              )}
              <button className="page-btn" onClick={() => setDbPage(p => Math.min(dbTotalPages, p + 1))} disabled={dbSafePage === dbTotalPages}>›</button>
              <button className="page-btn" onClick={() => setDbPage(dbTotalPages)} disabled={dbSafePage === dbTotalPages}>»</button>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <span style={{ fontSize: '11.5px', color: 'var(--text-2)' }}>{t.rowsPerPage}</span>
              <select
                className="page-size-select"
                value={dbPageSize}
                onChange={e => { setDbPageSize(Number(e.target.value)); setDbPage(1); }}
              >
                {[10, 25, 50, 100, 200].map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
