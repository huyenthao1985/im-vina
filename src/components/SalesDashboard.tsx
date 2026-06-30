import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import type { DataRow, ColumnMapping, FilterState } from '../types';
import { translations } from '../translations';
import { supabase } from '../lib/supabase';
import { CustomSelect } from './CustomSelect';

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
  // ── Stable rows cache ──────────────────────────────────────────────────────
  // Mục 1 nhận rows từ allRows global. Khi user upload file ở Mục 3 (Test_3),
  // allRows bị thay thế bởi dữ liệu Manpower không có cột 'division'
  // → normalizedRows về [] → charts bị reset.
  // Fix: cache snapshot cuối cùng có dữ liệu Sales hợp lệ; nếu rows mới
  // không có cột 'division' thì giữ lại snapshot cũ thay vì reset charts.
  const stableRowsRef = useRef<DataRow[]>([]);

  const effectiveRows = useMemo(() => {
    if (rows.length === 0) return stableRowsRef.current;
    // Nhận dạng: dữ liệu Sales/Mục 1 có cột 'division' (SALES/SHIP/PROD)
    const firstRowKeys = Object.keys(rows[0]).map(k => k.toLowerCase());
    const hasDivisionCol = firstRowKeys.some(k => k === 'division' || k === 'div');
    if (hasDivisionCol) {
      // Dữ liệu hợp lệ cho Mục 1 → cập nhật cache
      stableRowsRef.current = rows;
      return rows;
    }
    // Dữ liệu từ dashboard khác (Test_3, Test_2 OIS...) → dùng lại snapshot cũ
    return stableRowsRef.current.length > 0 ? stableRowsRef.current : rows;
  }, [rows]);
  // ──────────────────────────────────────────────────────────────────────────

  // Normalize raw rows to matches expected format
  const normalizedRows = useMemo(() => {
    if (effectiveRows.length === 0) return [];
    const keys = Object.keys(effectiveRows[0]);
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

    return effectiveRows.map(r => {
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
  }, [effectiveRows]);

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

  const [viewMode, setViewMode] = useState<'month' | 'quarter' | 'year'>('year');
  const [selectedPeriod, setSelectedPeriod] = useState<string>('');
  const [growthTimeframe, setGrowthTimeframe] = useState<'year' | 'quarter'>('year');

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

  const getQuarterLabel = (y: number, q: number) => {
    return lang === 'vi' ? `Q${q}/${String(y).slice(-2)}` : lang === 'ko' ? `${y}년 Q${q}` : `Q${q}/${String(y).slice(-2)}`;
  };

  const getMonthLabel = (y: number, m: string) => {
    return `${m} ${String(y).slice(-2)}`;
  };

  // Global mapping to check quarter completeness
  const globalMonthsByYear = useMemo(() => {
    const map: Record<number, Set<string>> = {};
    normalizedRows.forEach(r => {
      if (r.month === 'TTL') return;
      if (!map[r.year]) map[r.year] = new Set<string>();
      map[r.year].add(r.month.toUpperCase());
    });
    return map;
  }, [normalizedRows]);

  const isQuarterCompleteGlobal = (year: number, qNum: number) => {
    const presentMonths = globalMonthsByYear[year];
    if (!presentMonths) return false;
    const qMonths = 
      qNum === 1 ? ['JAN', 'FEB', 'MAR'] :
      qNum === 2 ? ['APR', 'MAY', 'JUNE'] :
      qNum === 3 ? ['JULY', 'AUG', 'SEP'] :
      ['OCT', 'NOV', 'DEC'];
    return qMonths.every(m => presentMonths.has(m));
  };

  const getQuarterFromMonth = (month: string): number => {
    const m = month.toUpperCase();
    if (['JAN', 'FEB', 'MAR'].includes(m)) return 1;
    if (['APR', 'MAY', 'JUNE'].includes(m)) return 2;
    if (['JULY', 'AUG', 'SEP'].includes(m)) return 3;
    if (['OCT', 'NOV', 'DEC'].includes(m)) return 4;
    return 0;
  };

  const periodOptions = useMemo(() => {
    if (viewMode === 'year') {
      return [
        { value: 'ALL', label: lang === 'vi' ? 'Tất cả' : lang === 'ko' ? '전체' : 'All' },
        ...years.map(y => ({ value: String(y), label: String(y) }))
      ];
    } else if (viewMode === 'quarter') {
      const options: { value: string; label: string }[] = [];
      const yStart = Math.min(yearFrom, yearTo) || years[0];
      const yEnd = Math.max(yearFrom, yearTo) || years[years.length - 1];
      for (let y = yStart; y <= yEnd; y++) {
        for (let q = 1; q <= 4; q++) {
          const monthsForQ = q === 1 ? ['JAN', 'FEB', 'MAR'] :
                             q === 2 ? ['APR', 'MAY', 'JUNE'] :
                             q === 3 ? ['JULY', 'AUG', 'SEP'] :
                             ['OCT', 'NOV', 'DEC'];
          const hasData = normalizedRows.some(r => r.year === y && monthsForQ.includes(r.month.toUpperCase()));
          if (hasData) {
            options.push({
              value: `${y}-Q${q}`,
              label: getQuarterLabel(y, q)
            });
          }
        }
      }
      return options;
    } else {
      // viewMode === 'month'
      const options: { value: string; label: string }[] = [];
      const yStart = Math.min(yearFrom, yearTo) || years[0];
      const yEnd = Math.max(yearFrom, yearTo) || years[years.length - 1];
      const monthsOrder = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE', 'JULY', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      for (let y = yStart; y <= yEnd; y++) {
        monthsOrder.forEach(m => {
          const hasData = normalizedRows.some(r => r.year === y && r.month.toUpperCase() === m);
          if (hasData) {
            options.push({
              value: `${y}-${m}`,
              label: getMonthLabel(y, m)
            });
          }
        });
      }
      return options;
    }
  }, [viewMode, years, yearFrom, yearTo, normalizedRows, lang]);

  useEffect(() => {
    if (periodOptions.length > 0) {
      if (viewMode === 'year') {
        const yearOpts = periodOptions.filter(o => o.value !== 'ALL');
        if (yearOpts.length > 0) {
          setSelectedPeriod(yearOpts[yearOpts.length - 1].value);
        } else {
          setSelectedPeriod('ALL');
        }
      } else if (viewMode === 'quarter') {
        let defaultVal = '';
        for (let i = periodOptions.length - 1; i >= 0; i--) {
          const opt = periodOptions[i];
          const [yStr, qStr] = opt.value.split('-Q');
          const y = Number(yStr);
          const q = Number(qStr);
          if (isQuarterCompleteGlobal(y, q)) {
            defaultVal = opt.value;
            break;
          }
        }
        if (!defaultVal && periodOptions.length > 0) {
          defaultVal = periodOptions[periodOptions.length - 1].value;
        }
        setSelectedPeriod(defaultVal);
      } else {
        setSelectedPeriod(periodOptions[periodOptions.length - 1]?.value || '');
      }
    } else {
      setSelectedPeriod('');
    }
  }, [periodOptions, viewMode]);

  const selectedPeriodLabel = useMemo(() => {
    const found = periodOptions.find(o => o.value === selectedPeriod);
    return found ? found.label : selectedPeriod;
  }, [periodOptions, selectedPeriod]);

  // Aggregated Values for KPI Cards
  const ttl = useMemo(() => filteredRecords.filter(r => r.month === 'TTL'), [filteredRecords]);
  
  const sumBy = (arr: any[], div: string) => arr.filter(r => r.division === div).reduce((s, r) => s + r.value, 0);
  const totalSales = useMemo(() => sumBy(ttl, 'sales'), [ttl]);
  const totalShipment = useMemo(() => sumBy(ttl, 'shipment'), [ttl]);
  const totalProduction = useMemo(() => sumBy(ttl, 'production'), [ttl]);

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

  // Global list of quarters to find the latest complete quarter
  const globalQuarters = useMemo(() => {
    const qList: { year: number; q: number }[] = [];
    years.forEach(y => {
      for (let q = 1; q <= 4; q++) {
        const monthsForQ = q === 1 ? ['JAN', 'FEB', 'MAR'] :
                           q === 2 ? ['APR', 'MAY', 'JUNE'] :
                           q === 3 ? ['JULY', 'AUG', 'SEP'] :
                           ['OCT', 'NOV', 'DEC'];
        const hasData = normalizedRows.some(r => r.year === y && monthsForQ.includes(r.month.toUpperCase()));
        if (hasData) {
          qList.push({ year: y, q });
        }
      }
    });
    qList.sort((a, b) => a.year !== b.year ? a.year - b.year : a.q - b.q);
    return qList;
  }, [years, normalizedRows]);

  const latestCompleteQuarter = useMemo(() => {
    for (let i = globalQuarters.length - 1; i >= 0; i--) {
      const { year, q } = globalQuarters[i];
      if (isQuarterCompleteGlobal(year, q)) {
        return { year, q, index: i };
      }
    }
    return null;
  }, [globalQuarters, globalMonthsByYear]);

  const qoqGrowthData = useMemo(() => {
    if (!latestCompleteQuarter) {
      return { growth: null, text: lang === 'vi' ? '(chưa đủ dữ liệu)' : '(insufficient data)' };
    }
    const currY = latestCompleteQuarter.year;
    const currQ = latestCompleteQuarter.q;

    let prevY = currY;
    let prevQ = currQ - 1;
    if (prevQ === 0) {
      prevY = currY - 1;
      prevQ = 4;
    }

    const currComplete = isQuarterCompleteGlobal(currY, currQ);
    const prevComplete = isQuarterCompleteGlobal(prevY, prevQ);

    if (!currComplete || !prevComplete) {
      return { growth: null, text: lang === 'vi' ? '(chưa đủ dữ liệu)' : '(insufficient data)' };
    }

    const monthsCurr = currQ === 1 ? ['JAN', 'FEB', 'MAR'] :
                       currQ === 2 ? ['APR', 'MAY', 'JUNE'] :
                       currQ === 3 ? ['JULY', 'AUG', 'SEP'] :
                       ['OCT', 'NOV', 'DEC'];
    
    const monthsPrev = prevQ === 1 ? ['JAN', 'FEB', 'MAR'] :
                       prevQ === 2 ? ['APR', 'MAY', 'JUNE'] :
                       prevQ === 3 ? ['JULY', 'AUG', 'SEP'] :
                       ['OCT', 'NOV', 'DEC'];

    const currSales = filteredRecords
      .filter(r => r.division === 'sales' && r.year === currY && monthsCurr.includes(r.month.toUpperCase()))
      .reduce((sum, r) => sum + r.value, 0);

    const prevSales = filteredRecords
      .filter(r => r.division === 'sales' && r.year === prevY && monthsPrev.includes(r.month.toUpperCase()))
      .reduce((sum, r) => sum + r.value, 0);

    if (prevSales <= 0) {
      return { growth: null, text: 'N/A' };
    }

    const growth = ((currSales - prevSales) / prevSales) * 100;
    const labelCurr = `Q${currQ} ${currY}`;
    const labelPrev = `Q${prevQ} ${prevY}`;
    const text = lang === 'vi' ? `${labelCurr} so với ${labelPrev}` : lang === 'en' ? `${labelCurr} vs ${labelPrev}` : `${labelCurr} 대비 ${labelPrev}`;

    return { growth, text };
  }, [latestCompleteQuarter, filteredRecords, lang]);

  // Format Helper
  const fmt = (n: number) => Math.round(n).toLocaleString('vi-VN');

  const t = translations[lang];

  const finalGrowthValue = useMemo(() => {
    if (growthTimeframe === 'year') {
      return growthData.growth === null ? 'N/A' : (growthData.growth >= 0 ? '+' : '') + growthData.growth.toFixed(1) + '%';
    } else {
      return qoqGrowthData.growth === null ? 'N/A' : (qoqGrowthData.growth >= 0 ? '+' : '') + qoqGrowthData.growth.toFixed(1) + '%';
    }
  }, [growthTimeframe, growthData, qoqGrowthData]);

  const finalGrowthSub = useMemo(() => {
    if (growthTimeframe === 'year') {
      return growthData.growthYears ? growthData.growthYears : (growthData.growth === null ? t.insufficientYears : '');
    } else {
      return qoqGrowthData.text;
    }
  }, [growthTimeframe, growthData, qoqGrowthData, t]);

  const finalGrowthLabel = useMemo(() => {
    if (growthTimeframe === 'year') {
      return lang === 'vi' ? 'Tăng trưởng Doanh số (YoY)' : lang === 'en' ? 'Sales Growth (YoY)' : '매출 증감 (YoY)';
    } else {
      return lang === 'vi' ? 'Tăng trưởng Doanh số (QoQ)' : lang === 'en' ? 'Sales Growth (QoQ)' : '매출 증감 (QoQ)';
    }
  }, [growthTimeframe, lang]);

  // KPI card configuration (4 cards — conversionRatio removed)
  const kpis = [
    { label: t.totalSales, value: fmt(totalSales) + ' K$', color: 'var(--purple)', tint: 'var(--purple-soft)', icon: '📈' },
    { label: t.totalShipment, value: fmt(totalShipment) + ' Kea', color: 'var(--cyan)', tint: 'var(--cyan-soft)', icon: '📦' },
    { label: t.totalProduction, value: fmt(totalProduction) + ' Kea', color: 'var(--green)', tint: 'var(--green-soft)', icon: '🏭' },
    {
      label: finalGrowthLabel,
      value: finalGrowthValue,
      color: 'var(--orange)',
      tint: 'var(--amber-soft)',
      icon: '⚡',
      sub: finalGrowthSub
    }
  ];

  /**
   * SINGLE SOURCE OF TRUTH for charts and tables that need per-model / per-customer breakdown.
   * Rules:
   *  - ALWAYS excludes month='TTL' rows (TTL rows never carry model/origin metadata).
   *  - viewMode=year, ALL → all non-TTL rows within the yearFrom-yearTo range.
   *  - viewMode=year, specific year → non-TTL rows for that year only.
   *  - viewMode=quarter → non-TTL rows for the selected quarter's months.
   *  - viewMode=month → non-TTL rows for the selected month.
   */
  const chartFilteredRecords = useMemo(() => {
    // Base: always exclude TTL rows
    const nonTtl = filteredRecords.filter(r => r.month !== 'TTL');

    if (!selectedPeriod) return nonTtl;

    if (viewMode === 'year') {
      if (selectedPeriod === 'ALL') {
        // ALL → keep everything in the yearFrom-yearTo window (already handled by filteredRecords)
        return nonTtl;
      }
      const y = Number(selectedPeriod);
      return nonTtl.filter(r => r.year === y);
    } else if (viewMode === 'quarter') {
      const [yStr, qStr] = selectedPeriod.split('-Q');
      const y = Number(yStr);
      const q = Number(qStr);
      if (!y || !q) return nonTtl;
      const qMonths =
        q === 1 ? ['JAN', 'FEB', 'MAR'] :
        q === 2 ? ['APR', 'MAY', 'JUNE'] :
        q === 3 ? ['JULY', 'AUG', 'SEP'] :
                  ['OCT', 'NOV', 'DEC'];
      return nonTtl.filter(r => r.year === y && qMonths.includes(r.month.toUpperCase()));
    } else {
      // viewMode === 'month'
      if (selectedPeriod === 'ALL') return nonTtl;
      const parts = selectedPeriod.split('-');
      const y = Number(parts[0]);
      const m = parts.slice(1).join('-'); // handle month names that may contain '-'
      return nonTtl.filter(r => r.year === y && r.month.toUpperCase() === m.toUpperCase());
    }
  }, [viewMode, selectedPeriod, filteredRecords]);

  // Alias for detail table (same data, kept separate so the variable name stays clear)
  const tableFilteredRecords = chartFilteredRecords;


  // Render Plotly charts dynamically
  useEffect(() => {
    if (filteredRecords.length === 0 || typeof window.Plotly === 'undefined') return;

    try {
      const isDark = theme === 'dark';
      const chartTextColor = isDark ? '#f3f4f6' : '#1e293b';
      const chartGridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';

      // 1. Unified Trend Chart
      let xList: string[] = [];
      let yShip: number[] = [];
      let ySales: number[] = [];
      let yGrowth: (number | null)[] = [];
      let growthName = '';

      if (viewMode === 'year') {
        growthName = lang === 'vi' ? 'Tr.trưởng YoY' : lang === 'en' ? 'YoY Growth' : '전년 대비 증감';
        const salesByYear: Record<number, number> = {};
        const shipByYear: Record<number, number> = {};
        ttl.filter(r => r.division === 'sales').forEach(r => {
          salesByYear[r.year] = (salesByYear[r.year] || 0) + r.value;
        });
        ttl.filter(r => r.division === 'shipment').forEach(r => {
          shipByYear[r.year] = (shipByYear[r.year] || 0) + r.value;
        });
        const yrsList = Object.keys(salesByYear).map(Number).sort((a, b) => a - b);
        const chartYears = yrsList.length ? yrsList : Object.keys(shipByYear).map(Number).sort((a, b) => a - b);
        xList = chartYears.map(String);
        yShip = chartYears.map(y => shipByYear[y] || 0);
        ySales = chartYears.map(y => salesByYear[y] || 0);
        yGrowth = chartYears.map((y, idx) => {
          if (idx === 0) return null;
          const prevY = chartYears[idx - 1];
          const prevVal = salesByYear[prevY] || 0;
          const currVal = salesByYear[y] || 0;
          return prevVal > 0 ? (currVal / prevVal) : null;
        });
      } else if (viewMode === 'quarter') {
        growthName = lang === 'vi' ? 'Tr.trưởng QoQ' : lang === 'en' ? 'QoQ Growth' : '전분기 대비 증감';
        const yStart = Math.min(yearFrom, yearTo) || years[0];
        const yEnd = Math.max(yearFrom, yearTo) || years[years.length - 1];
        
        const quartersInFilter: { year: number; q: number; key: string; label: string }[] = [];
        for (let y = yStart; y <= yEnd; y++) {
          for (let q = 1; q <= 4; q++) {
            const monthsForQ = q === 1 ? ['JAN', 'FEB', 'MAR'] :
                               q === 2 ? ['APR', 'MAY', 'JUNE'] :
                               q === 3 ? ['JULY', 'AUG', 'SEP'] :
                               ['OCT', 'NOV', 'DEC'];
            const hasData = normalizedRows.some(r => r.year === y && monthsForQ.includes(r.month.toUpperCase()));
            if (hasData) {
              quartersInFilter.push({
                year: y,
                q,
                key: `${y}-Q${q}`,
                label: getQuarterLabel(y, q)
              });
            }
          }
        }

        xList = quartersInFilter.map(item => item.label);
        
        const salesByQ: Record<string, number> = {};
        const shipByQ: Record<string, number> = {};
        
        filteredRecords.filter(r => r.month !== 'TTL').forEach(r => {
          const q = getQuarterFromMonth(r.month);
          if (q > 0) {
            const key = `${r.year}-Q${q}`;
            if (r.division === 'sales') {
              salesByQ[key] = (salesByQ[key] || 0) + r.value;
            } else if (r.division === 'shipment') {
              shipByQ[key] = (shipByQ[key] || 0) + r.value;
            }
          }
        });

        yShip = quartersInFilter.map(item => shipByQ[item.key] || 0);
        ySales = quartersInFilter.map(item => salesByQ[item.key] || 0);

        yGrowth = quartersInFilter.map((item) => {
          const currY = item.year;
          const currQ = item.q;
          let prevY = currY;
          let prevQ = currQ - 1;
          if (prevQ === 0) {
            prevY = currY - 1;
            prevQ = 4;
          }
          
          const currComplete = isQuarterCompleteGlobal(currY, currQ);
          const prevComplete = isQuarterCompleteGlobal(prevY, prevQ);
          
          if (!currComplete || !prevComplete) {
            return null;
          }

          const monthsPrev = prevQ === 1 ? ['JAN', 'FEB', 'MAR'] :
                             prevQ === 2 ? ['APR', 'MAY', 'JUNE'] :
                             prevQ === 3 ? ['JULY', 'AUG', 'SEP'] :
                             ['OCT', 'NOV', 'DEC'];
          const prevVal = filteredRecords
            .filter(r => r.division === 'sales' && r.year === prevY && monthsPrev.includes(r.month.toUpperCase()))
            .reduce((sum, r) => sum + r.value, 0);

          const currVal = salesByQ[item.key] || 0;
          return prevVal > 0 ? (currVal / prevVal) : null;
        });
      } else {
        growthName = lang === 'vi' ? 'Tr.trưởng MoM' : lang === 'en' ? 'MoM Growth' : '전월 대비 증감';
        const yStart = Math.min(yearFrom, yearTo) || years[0];
        const yEnd = Math.max(yearFrom, yearTo) || years[years.length - 1];
        const monthsOrder = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE', 'JULY', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        
        const monthsInFilter: { year: number; m: string; key: string; label: string }[] = [];
        for (let y = yStart; y <= yEnd; y++) {
          monthsOrder.forEach(m => {
            const hasData = normalizedRows.some(r => r.year === y && r.month.toUpperCase() === m);
            if (hasData) {
              monthsInFilter.push({
                year: y,
                m,
                key: `${y}-${m}`,
                label: getMonthLabel(y, m)
              });
            }
          });
        }

        xList = monthsInFilter.map(item => item.label);

        const salesByM: Record<string, number> = {};
        const shipByM: Record<string, number> = {};

        filteredRecords.filter(r => r.month !== 'TTL').forEach(r => {
          const m = r.month.toUpperCase();
          const key = `${r.year}-${m}`;
          if (r.division === 'sales') {
            salesByM[key] = (salesByM[key] || 0) + r.value;
          } else if (r.division === 'shipment') {
            shipByM[key] = (shipByM[key] || 0) + r.value;
          }
        });

        yShip = monthsInFilter.map(item => shipByM[item.key] || 0);
        ySales = monthsInFilter.map(item => salesByM[item.key] || 0);

        yGrowth = monthsInFilter.map((item) => {
          const currY = item.year;
          const currM = item.m;
          
          let prevY = currY;
          let prevMIndex = monthsOrder.indexOf(currM) - 1;
          if (prevMIndex < 0) {
            prevY = currY - 1;
            prevMIndex = 11;
          }
          const prevM = monthsOrder[prevMIndex];
          
          const prevVal = filteredRecords
            .filter(r => r.division === 'sales' && r.year === prevY && r.month.toUpperCase() === prevM)
            .reduce((sum, r) => sum + r.value, 0);

          const currVal = salesByM[item.key] || 0;
          return prevVal > 0 ? (currVal / prevVal) : null;
        });
      }

      const tracesUnified = [
        {
          x: xList,
          y: yShip,
          type: 'bar',
          name: lang === 'vi' ? 'XUẤT HÀNG (K)' : lang === 'en' ? 'SHIPMENT (K)' : '출하 (K)',
          marker: { color: '#2d7f96' },
          yaxis: 'y',
          text: yShip.map(v => v > 0 ? Math.round(v).toLocaleString('vi-VN') : ''),
          textposition: 'auto',
          textfont: { color: isDark ? '#ffffff' : '#000000' }
        },
        {
          x: xList,
          y: ySales,
          type: 'scatter',
          mode: 'lines+markers+text',
          name: lang === 'vi' ? 'DOANH SỐ (K$)' : lang === 'en' ? 'SALES AMT (K$)' : '매출 (K$)',
          line: { color: '#00a65a', width: 4 },
          marker: { size: 8, color: '#00a65a', symbol: 'circle' },
          yaxis: 'y',
          text: ySales.map(v => v > 0 ? Math.round(v).toLocaleString('vi-VN') : ''),
          textposition: 'top center',
          textfont: { color: chartTextColor, weight: 'bold' }
        },
        {
          x: xList,
          y: yGrowth,
          type: 'scatter',
          mode: 'lines+markers+text',
          name: growthName,
          line: { color: '#f39c12', width: 3, dash: 'dash' },
          marker: { size: 8, color: '#f39c12', symbol: 'circle' },
          yaxis: 'y2',
          text: yGrowth.map(v => v !== null && v > 0 ? Math.round(v * 100) + '%' : ''),
          textposition: 'top center',
          textfont: { color: '#e67e22', weight: 'bold' }
        }
      ];

      const layoutUnified = {
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor: 'rgba(0,0,0,0)',
        font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
        margin: { l: 54, r: 54, t: 10, b: 36 },
        legend: { orientation: 'h', y: 1.12, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
        xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
        yaxis: {
          title: { text: lang === 'vi' ? 'Xuất hàng / Doanh số' : lang === 'en' ? 'Shipment / Sales' : '출하 / 매출', font: { size: 11, color: chartTextColor } },
          gridcolor: chartGridColor,
          tickfont: { size: 10, color: chartTextColor }
        },
        yaxis2: {
          title: { text: growthName, font: { size: 11, color: chartTextColor } },
          overlaying: 'y',
          side: 'right',
          gridcolor: 'rgba(0,0,0,0)',
          tickfont: { size: 10, color: chartTextColor },
          tickformat: '.0%',
          range: [-0.5, 1.5]
        },
        hovermode: 'x unified'
      }

      const maxGrowthUnified = Math.max(...yGrowth.filter(v => v !== null).map(Number), 1.5);
      layoutUnified.yaxis2.range = [-0.5, Math.ceil(maxGrowthUnified * 1.25 * 2) / 2];

      const unifiedTrendEl = document.getElementById('trendChartUnified');
      if (unifiedTrendEl) {
        window.Plotly.react('trendChartUnified', tracesUnified as any, layoutUnified as any, { displayModeBar: false, responsive: true });
      }

      // 2. Top 10 Model theo Doanh số — uses chartFilteredRecords (non-TTL, period-aware)
      const salesByModel: Record<string, number> = {};
      chartFilteredRecords
        .filter(r => r.division === 'sales')
        .forEach(r => {
          if (r.model) salesByModel[r.model] = (salesByModel[r.model] || 0) + r.value;
        });
      const topModelsList = Object.entries(salesByModel)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

      console.log('[TopModel] selectedPeriod=', selectedPeriod, 'typeof=', typeof selectedPeriod,
        'chartFilteredRecords.length=', chartFilteredRecords.length,
        'salesByModel keys=', Object.keys(salesByModel).length,
        'topModelsList=', JSON.stringify(topModelsList));

      const topModelsUnifiedEl = document.getElementById('topModelsUnified');
      if (topModelsUnifiedEl) {
        // ALWAYS clear first — prevents stale 'no data' text persisting on re-render
        window.Plotly.purge('topModelsUnified');
        topModelsUnifiedEl.innerHTML = '';
        if (topModelsList.length === 0) {
          topModelsUnifiedEl.innerHTML = `<div class="panel-empty">${lang === 'vi' ? 'Không có dữ liệu.' : 'No data.'}</div>`;
        } else {
          const modelLabels = topModelsList.map(e => e[0]);
          const modelValues = topModelsList.map(e => e[1]);

          const tracesModelUnified = [
            {
              x: modelLabels,
              y: modelValues,
              type: 'bar',
              name: `SALES (K$) — ${selectedPeriodLabel}`,
              marker: {
                color: modelValues.map((_, idx) => idx < 3 ? '#0891b2' : '#d97706')
              },
              text: modelValues.map(v => Math.round(v).toLocaleString('vi-VN')),
              textposition: 'outside',
              textfont: {
                color: modelValues.map((_, idx) => idx < 3 ? (isDark ? '#fb7185' : '#e11d48') : chartTextColor),
                size: 11,
                weight: 'bold'
              },
              showlegend: true
            }
          ];
          const layoutModelUnified = {
            paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
            font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
            margin: { l: 50, r: 16, t: 10, b: 36 },
            legend: { orientation: 'h', y: 1.12, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
            xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
            yaxis: {
              gridcolor: chartGridColor,
              tickfont: { size: 10, color: chartTextColor },
              tickformat: '',
              title: { text: 'K$', font: { size: 10, color: chartTextColor } }
            },
            hovermode: 'x unified'
          };
          window.Plotly.newPlot('topModelsUnified', tracesModelUnified as any, layoutModelUnified as any, { displayModeBar: false, responsive: true });
        }
      }

      // 3. Customer Sales Stacked (Unified)
      const targetCustomers = ['GAOXIN', 'Q-TECH', 'SEMV', 'SUNNY'];
      const salesByCustKey: Record<string, Record<string, number>> = {};
      let custKeys: { key: string; label: string }[] = [];

      if (viewMode === 'year') {
        const yrsList = years.filter(y => y >= (yearFrom || years[0]) && y <= (yearTo || years[years.length - 1]));
        custKeys = yrsList.map(y => ({ key: String(y), label: String(y) }));
      } else if (viewMode === 'quarter') {
        const yStart = Math.min(yearFrom, yearTo) || years[0];
        const yEnd = Math.max(yearFrom, yearTo) || years[years.length - 1];
        for (let y = yStart; y <= yEnd; y++) {
          for (let q = 1; q <= 4; q++) {
            const monthsForQ = q === 1 ? ['JAN', 'FEB', 'MAR'] :
                               q === 2 ? ['APR', 'MAY', 'JUNE'] :
                               q === 3 ? ['JULY', 'AUG', 'SEP'] :
                               ['OCT', 'NOV', 'DEC'];
            const hasData = normalizedRows.some(r => r.year === y && monthsForQ.includes(r.month.toUpperCase()));
            if (hasData) {
              custKeys.push({ key: `${y}-Q${q}`, label: getQuarterLabel(y, q) });
            }
          }
        }
      } else {
        const yStart = Math.min(yearFrom, yearTo) || years[0];
        const yEnd = Math.max(yearFrom, yearTo) || years[years.length - 1];
        const monthsOrder = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE', 'JULY', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        for (let y = yStart; y <= yEnd; y++) {
          monthsOrder.forEach(m => {
            const hasData = normalizedRows.some(r => r.year === y && r.month.toUpperCase() === m);
            if (hasData) {
              custKeys.push({ key: `${y}-${m}`, label: getMonthLabel(y, m) });
            }
          });
        }
      }

      custKeys.forEach(item => {
        salesByCustKey[item.key] = {};
        targetCustomers.forEach(c => {
          salesByCustKey[item.key][c] = 0;
        });
      });

      if (viewMode === 'year') {
        ttl.filter(r => r.division === 'sales').forEach(r => {
          const yStr = String(r.year);
          const c = String(r.origin || '').toUpperCase().trim();
          if (salesByCustKey[yStr] && targetCustomers.includes(c)) {
            salesByCustKey[yStr][c] += r.value;
          }
        });
      } else if (viewMode === 'quarter') {
        filteredRecords.filter(r => r.month !== 'TTL' && r.division === 'sales').forEach(r => {
          const q = getQuarterFromMonth(r.month);
          if (q > 0) {
            const key = `${r.year}-Q${q}`;
            const c = String(r.origin || '').toUpperCase().trim();
            if (salesByCustKey[key] && targetCustomers.includes(c)) {
              salesByCustKey[key][c] += r.value;
            }
          }
        });
      } else {
        filteredRecords.filter(r => r.month !== 'TTL' && r.division === 'sales').forEach(r => {
          const m = r.month.toUpperCase();
          const key = `${r.year}-${m}`;
          const c = String(r.origin || '').toUpperCase().trim();
          if (salesByCustKey[key] && targetCustomers.includes(c)) {
            salesByCustKey[key][c] += r.value;
          }
        });
      }

      const customerColors: Record<string, string> = {
        'GAOXIN': '#0891b2',
        'Q-TECH': '#e11d48',
        'SEMV': '#059669',
        'SUNNY': '#7c3aed'
      };

      const tracesCustomerStacked = targetCustomers.map(c => {
        const yVals = custKeys.map(item => salesByCustKey[item.key][c]);
        return {
          x: custKeys.map(item => item.label),
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
        margin: { l: 50, r: 16, t: 10, b: 36 },
        barmode: 'stack',
        legend: { orientation: 'h', y: 1.12, x: 0.5, xanchor: 'center', font: { color: chartTextColor } },
        xaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor }, type: 'category' },
        yaxis: { gridcolor: chartGridColor, tickfont: { size: 10, color: chartTextColor } }
      };

      const stackedCustomerEl = document.getElementById('customerSalesStackedUnified');
      if (stackedCustomerEl) {
        window.Plotly.react('customerSalesStackedUnified', tracesCustomerStacked as any, layoutCustomerStacked as any, { displayModeBar: false, responsive: true });
      }
      // 4. Customer Sales Donut (Unified)
      const customerSalesFiltered: Record<string, number> = {};
      targetCustomers.forEach(c => { customerSalesFiltered[c] = 0; });
      chartFilteredRecords
        .filter(r => r.division === 'sales')
        .forEach(r => {
          const c = String(r.origin || '').toUpperCase().trim();
          if (targetCustomers.includes(c)) {
            customerSalesFiltered[c] += r.value;
          }
        });

      const donutLabels = targetCustomers.filter(c => customerSalesFiltered[c] > 0);
      const donutValues = donutLabels.map(c => customerSalesFiltered[c]);
      const totalSales = donutValues.reduce((sum, val) => sum + val, 0);

      const donutEl = document.getElementById('customerSalesDonutUnified');
      if (donutEl) {
        window.Plotly.purge('customerSalesDonutUnified');
        donutEl.innerHTML = '';
        if (totalSales === 0) {
          donutEl.innerHTML = `<div class="panel-empty">${lang === 'vi' ? 'Không có dữ liệu.' : 'No data.'}</div>`;
        } else {
          const donutColors = donutLabels.map(c => customerColors[c] || '#cbd5e1');
          const traceCustomerDonut = [{
            labels: donutLabels,
            values: donutValues,
            type: 'pie',
            hole: 0.58,
            marker: { colors: donutColors, line: { color: isDark ? '#1e293b' : '#ffffff', width: 2 } },
            textinfo: 'value',
            texttemplate: '%{value:,.0f}',
            textfont: { size: 10, color: '#ffffff', weight: 'bold' },
            hoverinfo: 'label+value+percent'
          }];
          const layoutCustomerDonut = {
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
            margin: { l: 5, r: 5, t: 5, b: 5 },
            showlegend: true,
            legend: {
              orientation: 'v',
              x: 0.85,
              y: 0.5,
              yanchor: 'middle',
              font: { size: 10, color: chartTextColor }
            },
            annotations: [
              {
                font: { size: 34 },
                showarrow: false,
                text: '<span style="text-shadow: 0 0 8px rgba(251, 191, 36, 0.7)">💡</span>',
                x: 0.5,
                y: 0.5
              }
            ]
          };
          window.Plotly.react('customerSalesDonutUnified', traceCustomerDonut as any, layoutCustomerDonut as any, { displayModeBar: false, responsive: true });
        }
      }
    } catch (err: any) {
      console.error("PLOTLY_RENDER_ERROR_OCCURRED:", err.message, err.stack);
      const errorDiv = document.createElement('div');
      errorDiv.className = 'plotly-error-banner';
      errorDiv.style.color = 'red';
      errorDiv.style.padding = '10px';
      errorDiv.style.background = '#fee2e2';
      errorDiv.style.border = '1px solid #fca5a5';
      errorDiv.style.margin = '10px';
      errorDiv.style.borderRadius = '4px';
      errorDiv.innerText = "Error rendering charts: " + err.message + "\nStack: " + err.stack;
      document.querySelector('.dash-container')?.prepend(errorDiv);
    }
  }, [filteredRecords, chartFilteredRecords, viewMode, theme, lang, t, selectedPeriod, selectedPeriodLabel]);

  // Detail table data
  const tableRows = useMemo(() => {
    const modelsMap: Record<string, any> = {};
    tableFilteredRecords.forEach(r => {
      if (!modelsMap[r.model]) {
        modelsMap[r.model] = { model: r.model, origin: r.origin, customer: r.customer, type: r.type, production: 0, shipment: 0, sales: 0 };
      }
      if (r.division) {
        modelsMap[r.model][r.division] += r.value;
      }
      if (!modelsMap[r.model].origin && r.origin) modelsMap[r.model].origin = r.origin;
      if (!modelsMap[r.model].customer && r.customer) modelsMap[r.model].customer = r.customer;
      if (!modelsMap[r.model].type && r.type) modelsMap[r.model].type = r.type;
    });
    const rowsList = Object.values(modelsMap).map((m: any) => ({
      ...m,
      ratio: m.production > 0 ? (m.shipment / m.production * 100) : 0
    }));
    rowsList.sort((a, b) => b.sales - a.sales);
    return rowsList.slice(0, 15);
  }, [tableFilteredRecords]);

  // Full Database table state and logic
  const [dbSearch, setDbSearch] = useState('');
  const [dbSort, setDbSort] = useState<{ column: string | null; direction: 'asc' | 'desc' | null }>({ column: null, direction: null });
  const [dbPage, setDbPage] = useState(1);
  const [dbPageSize, setDbPageSize] = useState(25);

  const [dbSelModel, setDbSelModel] = useState('');
  const [dbSelYear, setDbSelYear] = useState('');
  const [dbSelMonth, setDbSelMonth] = useState('');
  const [dbSelQuarter, setDbSelQuarter] = useState('');
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
    if (dbSelQuarter) {
      result = result.filter(r => {
        const q = getQuarterFromMonth(r.month);
        return `Q${q}` === dbSelQuarter;
      });
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
  }, [filteredRecords, dbSearch, dbSelModel, dbSelYear, dbSelMonth, dbSelQuarter, dbSelCustomer, lang]);

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
  }, [dbSearch, dbSelModel, dbSelYear, dbSelMonth, dbSelQuarter, dbSelCustomer, yearFrom, yearTo, origin, selModel]);

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
      <style>{`
        .sales-dashboard:not(.second-dashboard) .hero {
          padding: 16px 20px 4px !important;
        }
        .sales-dashboard:not(.second-dashboard) .hero h1 {
          margin: 0 0 6px !important;
          font-size: 26px !important;
        }
        .sales-dashboard:not(.second-dashboard) .header-line {
          margin-bottom: 12px !important;
        }
      `}</style>

      {/* Centered Hero Header */}
      <div className="hero" style={{ position: 'relative', textAlign: 'center' }}>
        <h1 style={{ textAlign: 'center' }}>{t.mainTitleDash}</h1>
        <div className="header-line" style={{ margin: '0 auto' }}></div>
      </div>

      <div className="dash-container">
        {/* Filter bar exactly matching reference layout */}
        <div className="topbar-dash" style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '16px' }}>
          {/* Dòng 1 (labels) */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
            {/* Cụm trái dòng 1: Đồng hồ */}
            <div style={{ width: '170px', display: 'flex', alignItems: 'center', flexShrink: 0, fontSize: '13px', color: 'var(--text-2)', fontWeight: '700', whiteSpace: 'nowrap' }}>
              {formattedTime}
            </div>
            {/* Cụm giữa dòng 1: Labels */}
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px' }}>
              <span style={{ width: '90px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: 'var(--text-2)', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.fromYear}</span>
              <span style={{ width: '90px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: 'var(--text-2)', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.toYear}</span>
              <span style={{ width: '110px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: 'var(--text-2)', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.partner}</span>
              <span style={{ width: '110px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: 'var(--text-2)', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.colModel}</span>
              <span style={{ width: '180px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: 'var(--text-2)', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{lang === 'vi' ? 'XEM THEO' : lang === 'ko' ? '보기 방식' : 'VIEW BY'}</span>
              <span style={{ width: '140px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: 'var(--text-2)', textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{lang === 'vi' ? 'CHI TIẾT' : lang === 'ko' ? '상세 선택' : 'DETAIL'}</span>
              {activeFilterCount > 0 && <span style={{ width: '130px', flexShrink: 0 }}></span>}
              {activeFilterCount > 0 && <span style={{ width: '60px', flexShrink: 0 }}></span>}
            </div>
            {/* Cụm phải dòng 1: Spacers matching Dòng 2 (Tải tệp lên + Save to Cloud) */}
            <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
              <div style={{ width: '130px' }}></div>
              {supabase && <div style={{ width: '120px' }}></div>}
            </div>
          </div>

          {/* Dòng 2 (values/controls) */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
            {/* Cụm trái dòng 2: Trống để giữ căn lề */}
            <div style={{ width: '170px', flexShrink: 0 }}></div>
            {/* Cụm giữa dòng 2: Controls */}
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px', alignItems: 'center' }}>
              <CustomSelect
                value={String(yearFrom)}
                onChange={v => setYearFrom(Number(v))}
                options={years.map(y => ({ value: String(y), label: String(y) }))}
                style={{ width: '90px', height: '38px' }}
              />
              <CustomSelect
                value={String(yearTo)}
                onChange={v => setYearTo(Number(v))}
                options={years.map(y => ({ value: String(y), label: String(y) }))}
                style={{ width: '90px', height: '38px' }}
              />
              <CustomSelect
                value={origin}
                onChange={setOrigin}
                options={[
                  { value: '', label: t.allOption },
                  ...origins.map(o => ({ value: o, label: o }))
                ]}
                style={{ width: '110px', height: '38px' }}
              />
              <CustomSelect
                value={selModel}
                onChange={setSelModel}
                options={[
                  { value: '', label: t.allOption },
                  ...models.map(m => ({ value: m, label: m }))
                ]}
                style={{ width: '110px', height: '38px' }}
              />
              <div style={{ display: 'flex', gap: '0px', height: '38px', width: '180px', flexShrink: 0 }}>
                {(['month', 'quarter', 'year'] as const).map(mode => (
                  <button
                    key={mode}
                    onClick={() => setViewMode(mode)}
                    style={{
                      flex: 1,
                      padding: '6px 0',
                      fontSize: '13px',
                      fontWeight: 600,
                      borderRadius: mode === 'month' ? '6px 0 0 6px' : mode === 'year' ? '0 6px 6px 0' : '0',
                      border: '1px solid var(--border)',
                      borderRight: mode !== 'year' ? 'none' : '1px solid var(--border)',
                      background: viewMode === mode ? '#2e7d8c' : 'var(--bg-card)',
                      color: viewMode === mode ? '#ffffff' : 'var(--text)',
                      cursor: 'pointer',
                      transition: 'all 0.15s ease',
                      height: '100%',
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      whiteSpace: 'nowrap'
                    }}
                  >
                    {mode === 'month'
                      ? (lang === 'vi' ? 'Tháng' : lang === 'ko' ? '월별' : 'Month')
                      : mode === 'quarter'
                        ? (lang === 'vi' ? 'Quý' : lang === 'ko' ? '분기' : 'Quarter')
                        : (lang === 'vi' ? 'Năm' : lang === 'ko' ? '연별' : 'Year')
                    }
                  </button>
                ))}
              </div>
              <CustomSelect
                value={selectedPeriod}
                onChange={setSelectedPeriod}
                options={periodOptions}
                style={{ width: '140px', height: '38px' }}
              />
              {activeFilterCount > 0 && (
                <span className="filter-badge" style={{ width: '130px', height: '38px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', boxSizing: 'border-box', margin: 0, flexShrink: 0 }}>
                  {activeFilterCount} {t.filtersApplied}
                </span>
              )}
              {activeFilterCount > 0 && (
                <button 
                  className="btn-text" 
                  onClick={handleResetFilters}
                  style={{ width: '60px', height: '38px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', boxSizing: 'border-box', padding: 0, flexShrink: 0 }}
                >
                  {t.resetBtn}
                </button>
              )}
            </div>
            {/* Cụm phải dòng 2: Tải file lên + Save to Cloud */}
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexShrink: 0 }}>
              {/* Lang select + Theme toggle đã di chuyển lên GlobalHeaderControls
                  ở góc trên-phải toàn trang trong App.tsx */}

              {/* Nút TẢI FILE LÊN (upload) — Mục 1 trước đây KHÔNG có nút này
                  (onFileSelected prop nhận vào nhưng chưa từng được gọi) → user
                  không có cách nào tải dữ liệu mới trực tiếp tại Mục 1.
                  Bổ sung theo đúng pattern của TargetActualDashboard. */}
              <label
                className="btn-outline"
                style={{ cursor: 'pointer', margin: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px', color: 'var(--text)', border: '1px solid var(--border)', height: '38px', width: '130px', boxSizing: 'border-box', fontSize: '13px' }}
              >
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
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15" style={{ flexShrink: 0 }}>
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <path d="M17 8l-5-5-5 5" />
                  <path d="M12 3v12" />
                </svg>
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {lang === 'vi' ? 'Tải tệp lên' : lang === 'ko' ? '파일 업로드' : 'Upload File'}
                </span>
              </label>

              {supabase && (
                <button 
                  className="btn-outline" 
                  type="button" 
                  onClick={handleSaveToCloud} 
                  disabled={isSaving}
                  style={{ borderColor: 'var(--green)', color: 'var(--green)', height: '38px', width: '120px', boxSizing: 'border-box', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', margin: 0 }}
                >
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{isSaving ? '⏳...' : t.saveToCloudBtn}</span>
                </button>
              )}
            </div>
          </div>
        </div>

        {/* KPI Grid — 4 cards, compact style matching Dashboard mới */}
        {/*
          Layout strategy:
          - Cards 1-3: flex 1 (equal share), min-width 160px
          - Card 4 (growth): flex 1.4 + min-width 240px so the long Vi title
            "TĂNG TRƯỞNG DOANH SỐ (YOY)" + toggle always fits on one line.
          - Title uses white-space:nowrap + clamp() font-size for resilience.
        */}
        <div style={{
          display: 'flex',
          flexDirection: 'row',
          gap: '12px',
          marginBottom: '16px',
          flexWrap: 'nowrap',
          alignItems: 'stretch'
        }}>
          {kpis.map((c, i) => {
            const isGrowth = i === 3; // last card is the growth card
            return (
              <div key={i} style={{
                flex: isGrowth ? '1.4 1 240px' : '1 1 160px',
                minWidth: isGrowth ? '240px' : '160px',
                border: `1px solid ${c.tint}`,
                borderLeft: `4px solid ${c.color}`,
                borderRadius: '10px',
                background: `linear-gradient(135deg, ${c.tint} 0%, rgba(30, 41, 59, 0.4) 100%)`,
                padding: '12px 14px',
                animationDelay: `${i * 0.05}s`,
                position: 'relative',
                display: 'flex',
                flexDirection: 'column',
                gap: '6px',
                boxSizing: 'border-box'
              }}>
                {/* Header row: icon + label (nowrap) + toggle */}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', gap: '4px', minWidth: 0 }}>
                  {/* Left: icon + label */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: '5px', minWidth: 0, flex: 1, overflow: 'hidden' }}>
                    <div style={{
                      background: c.tint,
                      color: c.color,
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      width: '26px',
                      height: '26px',
                      borderRadius: '50%',
                      fontSize: '13px',
                      flexShrink: 0
                    }}>{c.icon}</div>
                    <div style={{
                      /* clamp: shrinks from 13.5px down to 11px as container narrows */
                      fontSize: 'clamp(11px, 1.8vw, 13.5px)',
                      fontWeight: '700',
                      color: 'var(--text-0)',
                      textTransform: 'uppercase',
                      letterSpacing: '.02em',
                      lineHeight: 1.2,
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis'
                    }}>{c.label}</div>
                  </div>
                  {/* Right: toggle (growth card only) */}
                  {isGrowth && (
                    <div style={{ display: 'flex', background: 'var(--border)', borderRadius: '4px', padding: '1px', flexShrink: 0, marginLeft: '4px' }}>
                      {(['year', 'quarter'] as const).map(tf => (
                        <button
                          key={tf}
                          onClick={(e) => { e.stopPropagation(); setGrowthTimeframe(tf); }}
                          style={{
                            background: growthTimeframe === tf ? '#2e7d8c' : 'transparent',
                            color: growthTimeframe === tf ? '#ffffff' : 'var(--text-2)',
                            border: 'none',
                            borderRadius: '3px',
                            padding: '2px 5px',
                            fontSize: '10px',
                            fontWeight: 600,
                            cursor: 'pointer',
                            whiteSpace: 'nowrap'
                          }}
                        >
                          {tf === 'year' ? (lang === 'vi' ? 'NĂM' : 'YEAR') : (lang === 'vi' ? 'QUÝ' : 'QTR')}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                {/* Value row: big number + inline sub-text for growth card */}
                <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px', flexWrap: 'wrap' }}>
                  <div style={{ fontFamily: 'monospace', fontSize: '26px', fontWeight: '800', color: 'var(--text-0)', lineHeight: 1 }}>
                    {c.value}
                  </div>
                  {isGrowth && c.sub && (
                    <span style={{ fontSize: '11.5px', color: 'var(--text-2)', fontWeight: 500, whiteSpace: 'nowrap', lineHeight: 1 }}>
                      {c.sub}
                    </span>
                  )}
                </div>
              </div>
            );
          })}
        </div>


        {/* Charts Row 1 */}
        <div className="charts-row row-1">
          {/* Unified Trend Chart */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3 style={{ fontSize: '15px', fontWeight: '600' }}>
                  {viewMode === 'year' 
                    ? t.chart1Title 
                    : viewMode === 'quarter' 
                      ? (lang === 'vi' ? 'Xuất hàng & Doanh số theo Quý' : lang === 'en' ? 'Shipment & Sales by Quarter' : '분기별 출하 & 매출')
                      : t.chartMonthTitle
                  }
                </h3>
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
                marginTop: '2px',
                textTransform: 'uppercase'
              }}>
                {viewMode}
              </span>
            </div>
            <div className="chart-holder" id="trendChartUnified"></div>
          </div>

          {/* Unified Top Models Chart */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3 style={{ fontSize: '15px', fontWeight: '600' }}>
                  {lang === 'vi' 
                    ? `Top 10 Model theo Doanh số - ${selectedPeriodLabel}` 
                    : lang === 'en' 
                      ? `Top 10 Models by Sales - ${selectedPeriodLabel}` 
                      : `매출 Top 10 모델 - ${selectedPeriodLabel}`}
                </h3>
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
                marginTop: '2px',
                textTransform: 'uppercase'
              }}>
                {selectedPeriod}
              </span>
            </div>
            <div className="chart-holder" id="topModelsUnified"></div>
          </div>
        </div>

        {/* Charts Row 2 */}
        <div className="charts-row row-2">
          {/* Customer Stacked (Left) */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
              <div>
                <h3 style={{ fontSize: '15px', fontWeight: '600' }}>
                  {lang === 'vi' 
                    ? 'Doanh số theo Khách hàng (CUSTOM)' 
                    : lang === 'en' 
                      ? 'Customer Sales Trend (CUSTOM)' 
                      : '고객사별 매출 (CUSTOM)'}
                </h3>
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
                marginTop: '2px',
                textTransform: 'uppercase'
              }}>
                {viewMode} (CUSTOM)
              </span>
            </div>
            <div className="chart-holder" id="customerSalesStackedUnified"></div>
          </div>

            {/* Customer Pinwheel (Right) */}
            <div className="panel">
              <div className="panel-head" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div>
                  <h3 style={{ fontSize: '15px', fontWeight: '600' }}>
                    {lang === 'vi'
                      ? `Tỷ trọng theo Khách hàng - ${selectedPeriodLabel}`
                      : lang === 'en'
                        ? `Customer Sales Share - ${selectedPeriodLabel}`
                        : `고객사별 매출 비중 - ${selectedPeriodLabel}`}
                  </h3>
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
                  marginTop: '2px',
                  textTransform: 'uppercase'
                }}>
                  {selectedPeriod} (CUSTOM)
                </span>
              </div>
              <div className="chart-holder" id="customerSalesDonutUnified"></div>
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

              {/* Quarter Select */}
              <select
                value={dbSelQuarter}
                onChange={e => {
                  setDbSelQuarter(e.target.value);
                  setDbSelMonth(''); // mutually exclusive
                  setDbPage(1);
                }}
                className="lang-select"
                style={{ fontSize: '12.5px', padding: '5px 10px', height: '34px' }}
              >
                <option value="">{lang === 'vi' ? 'Quý: Tất cả' : lang === 'ko' ? '분기: 전체' : 'Quarter: All'}</option>
                <option value="Q1">Q1</option>
                <option value="Q2">Q2</option>
                <option value="Q3">Q3</option>
                <option value="Q4">Q4</option>
              </select>

              {/* Month Select */}
              <select
                value={dbSelMonth}
                onChange={e => {
                  setDbSelMonth(e.target.value);
                  setDbSelQuarter(''); // mutually exclusive
                  setDbPage(1);
                }}
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
