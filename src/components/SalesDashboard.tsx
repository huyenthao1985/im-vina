import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import type { DataRow, ColumnMapping, FilterState } from '../types';
import { translations } from '../translations';
import { CustomSelect } from './CustomSelect';
import { NeonButton } from './NeonButton';

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

// FIX (EPCC-vanilla-toolbar): màu chữ nhãn cố định (không đổi theo theme)
// dùng riêng cho thanh filter nền "Vanilla" (#FFF4D6) — cam đậm để đủ tương
// phản trên nền vàng nhạt, tương tự ảnh tham chiếu "Vanilla / HEX #FFF4D6".
const VANILLA_LABEL_COLOR = '#C0EF6A';

/* ═══════════════════════════════════════════════════════════════════════
 * FIX (header-legend-merge + header-color-match-template, EPCC) — áp dụng
 * lại đúng pattern đã dùng ở Mục 2 (TargetActualDashboard.tsx):
 *  1. Xóa badge vàng viền (YEAR/2026/CUSTOM...) ở góc phải mỗi khung —
 *     đây là phần người dùng khoanh đỏ + gạch chéo yêu cầu xóa.
 *  2. Gộp legend (đang do Plotly tự vẽ RỜI bên trong chart, tạo khoảng
 *     trắng thừa giữa tiêu đề và biểu đồ — đúng vùng người dùng khoanh đỏ
 *     + mũi tên chỉ lên) LÊN CHUNG 1 hàng với tiêu đề trong panel-head,
 *     tắt showlegend của Plotly để không còn lặp lại.
 *  3. Tô nền panel-head bằng đúng 4 màu tham chiếu lấy từ
 *     TargetActualDashboard.tsx (CHART_HEADER_THEME, thứ tự 1-4):
 *     tím #8b5cf6 → cyan #06b6d4 → cam #f59e0b → cam-đỏ #f97316.
 * ═══════════════════════════════════════════════════════════════════════ */
const CHART_HEADER_THEME = [
  { accent: '#8b5cf6', bgLight: '#d6c6fc', bgDark: 'rgba(139,92,246,0.14)' },
  { accent: '#06b6d4', bgLight: '#a8e5f0', bgDark: 'rgba(6,182,212,0.14)' },
  { accent: '#f59e0b', bgLight: '#fcddaa', bgDark: 'rgba(255,255,255,0.05)' },
  { accent: '#f97316', bgLight: '#fdcead', bgDark: 'rgba(249,115,22,0.14)' },
] as const;

// PHẢI khớp nguyên văn với `targetCustomers`/`customerColors` dùng trong
// effect vẽ chart (Chart 3 & 4) để legend header và màu cột/lát cắt luôn
// đồng bộ dù danh sách khách hàng thay đổi theo dữ liệu.
const CUSTOMER_LEGEND_ORDER = ['GAOXIN', 'Q-TECH', 'SEMV', 'SUNNY'] as const;
const CUSTOMER_COLORS: Record<string, string> = {
  'GAOXIN': '#0891b2',
  'Q-TECH': '#e11d48',
  'SEMV': '#059669',
  'SUNNY': '#7c3aed'
};

const SDLegendItem: React.FC<{ type: 'bar' | 'line' | 'dot'; color: string; label: string }> = ({ type, color, label }) => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '10.5px', fontWeight: 500, whiteSpace: 'nowrap', textTransform: 'none' }}>
    {type === 'bar' && <span style={{ width: '9px', height: '9px', borderRadius: '2px', background: color, flexShrink: 0 }} />}
    {type === 'line' && <span style={{ width: '12px', height: 0, borderTop: `2px dashed ${color}`, flexShrink: 0 }} />}
    {type === 'dot' && <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: color, flexShrink: 0 }} />}
    {label}
  </span>
);

export const SalesDashboard: React.FC<SalesDashboardProps> = ({
  rows,
  theme,
  onToggleTheme: _onToggleTheme,
  lang,
  setLang: _setLang,
  onFileSelected
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isLightMode = theme === 'light';
  // FIX (header-color-match-template, EPCC): style nền panel-head 4 khung
  // Mục 1, lấy nền sáng/tối (bgLight/bgDark) + viền trái accent từ
  // CHART_HEADER_THEME, idx = 0-3 tương ứng đúng thứ tự 4 khung.
  const chartHeaderStyle = (idx: number): React.CSSProperties => {
    const c = CHART_HEADER_THEME[idx % CHART_HEADER_THEME.length];
    return {
      background: isLightMode ? c.bgLight : c.bgDark,
      borderLeft: `4px solid ${c.accent}`,
      borderRadius: '8px 8px 0 0',
      padding: '10px 14px',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      flexWrap: 'wrap' as const,
      gap: '8px'
    };
  };
  // ── Fix: "biểu đồ hiện khung trống, không có hình/số" khi mới mở trang ──
  // Root cause: script Plotly load từ CDN (bất đồng bộ) có thể CHƯA sẵn sàng
  // tại thời điểm effect vẽ chart chạy lần đầu. Guard cũ chỉ kiểm tra
  // `typeof window.Plotly === 'undefined'` rồi bail ra — và vì không có gì
  // trigger effect chạy lại sau khi Plotly tải xong, chart bị TRỐNG VĨNH VIỄN
  // dù data đã có sẵn. Fix: poll tới khi Plotly sẵn sàng rồi set state để
  // effect vẽ chart tự chạy lại.
  const [plotlyReady, setPlotlyReady] = useState<boolean>(
    typeof window !== 'undefined' && !!window.Plotly
  );
  useEffect(() => {
    if (plotlyReady) return;
    const id = setInterval(() => {
      if (typeof window !== 'undefined' && window.Plotly) {
        setPlotlyReady(true);
        clearInterval(id);
      }
    }, 150);
    return () => clearInterval(id);
  }, [plotlyReady]);

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
  // FIX: đã xóa handleSaveToCloud/isSaving/nút "Lên mây" — file này bị sót
  // lại khi dọn dẹp double-write path ở 3 dashboard khác (Manpower/PerCapita/
  // TargetActual). App.tsx đã tự động đồng bộ bucket Sales lên Supabase qua
  // syncBucketToSupabase() (dùng upsert onConflict đúng SALES_DATA_KEY) —
  // giữ thêm 1 đường ghi thủ công delete()+insert() thô ở đây sẽ tái lập
  // đúng bug gốc "ghi trùng dữ liệu qua 2 đường" đã fix ở phần đầu.

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
  // Fallback khi không có dòng TTL (vd dữ liệu Supabase chỉ có monthly rows):
  // KPI card và chart không bị hiện 0 trống nữa mà tựng hợp từ monthly.
  const totalSales = useMemo(() => {
    const v = sumBy(ttl, 'sales');
    if (v > 0) return v;
    return filteredRecords.filter(r => r.division === 'sales' && r.month !== 'TTL').reduce((s, r) => s + r.value, 0);
  }, [ttl, filteredRecords]);
  const totalShipment = useMemo(() => {
    const v = sumBy(ttl, 'shipment');
    if (v > 0) return v;
    return filteredRecords.filter(r => r.division === 'shipment' && r.month !== 'TTL').reduce((s, r) => s + r.value, 0);
  }, [ttl, filteredRecords]);
  const totalProduction = useMemo(() => {
    const v = sumBy(ttl, 'production');
    if (v > 0) return v;
    return filteredRecords.filter(r => r.division === 'production' && r.month !== 'TTL').reduce((s, r) => s + r.value, 0);
  }, [ttl, filteredRecords]);

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

  // FIX (header-legend-merge, EPCC): label legend "Tr.trưởng ..." hiển thị
  // trong panel-head (bên cạnh tiêu đề) — đồng bộ đúng điều kiện viewMode
  // dùng để đặt tên growthName bên trong effect vẽ chart (year→YoY,
  // quarter→QoQ, month→MoM).
  const trendGrowthLegendLabel = viewMode === 'year'
    ? (lang === 'vi' ? 'Tr.trưởng YoY' : lang === 'en' ? 'YoY Growth' : '전년 대비 증감')
    : viewMode === 'quarter'
      ? (lang === 'vi' ? 'Tr.trưởng QoQ' : lang === 'en' ? 'QoQ Growth' : '전분기 대비 증감')
      : (lang === 'vi' ? 'Tr.trưởng MoM' : lang === 'en' ? 'MoM Growth' : '전월 대비 증감');

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

  // FIX (EPCC-unify-kpi-colors-all-sections): đồng nhất bộ màu KPI card với
  // mục 2 (TargetActualDashboard) và mục 3 (PerCapitaTab) — cùng 4 màu theo
  // đúng thứ tự trái→phải: teal #2e7d8c → green #10b981 → purple #8b5cf6 →
  // orange #f59e0b. Trước đây dùng biến theme var(--purple)/var(--cyan)/
  // var(--green)/var(--orange) (đổi theo theme) — nay đổi sang màu cố định
  // giống hệt 2 mục kia để cả 3 mục nhất quán một bộ màu (không đổi theo theme).
  // FIX (EPCC-unify-kpi-icon-size-muc1-theo-muc2): mục 2 (TargetActualDashboard,
  // thẻ chuẩn) dùng SVG line-icon phẳng 16x16px, KHÔNG có khung tròn nền màu.
  // Trước đây mục 1 dùng div badge tròn 26x26px chứa emoji, khiến hàng header
  // (icon+label) cao hơn ~8-10px so với chuẩn, kéo theo tổng chiều cao card
  // bị lệch. Thay bằng SVG line-icon 16x16 giống hệt mục 2 để chiều cao khớp.
  const renderKpiIcon = (key: string, color: string) => {
    const common = {
      viewBox: '0 0 24 24',
      fill: 'none',
      stroke: color,
      strokeWidth: 2.5,
      strokeLinecap: 'round' as const,
      strokeLinejoin: 'round' as const,
      style: { width: '16px', height: '16px', flexShrink: 0 }
    };
    switch (key) {
      case 'trending-up':
        return (
          <svg {...common}>
            <polyline points="23 6 13.5 15.5 8.5 10.5 1 18" />
            <polyline points="17 6 23 6 23 12" />
          </svg>
        );
      case 'package':
        return (
          <svg {...common}>
            <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
            <polyline points="3.27 6.96 12 12.01 20.73 6.96" />
            <line x1="12" y1="22.08" x2="12" y2="12" />
          </svg>
        );
      case 'factory':
        return (
          <svg {...common}>
            <path d="M2 20a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V8l-7 5V8l-7 5V4l-6 5Z" />
            <path d="M17 18h1" />
            <path d="M12 18h1" />
            <path d="M7 18h1" />
          </svg>
        );
      case 'zap':
        return (
          <svg {...common}>
            <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" />
          </svg>
        );
      default:
        return null;
    }
  };

  const kpis = [
    { label: t.totalSales, value: fmt(totalSales) + ' K$', color: '#2e7d8c', tint: 'rgba(46,125,140,0.1)', icon: 'trending-up' },
    { label: t.totalShipment, value: fmt(totalShipment) + ' Kea', color: '#10b981', tint: 'rgba(16,185,129,0.1)', icon: 'package' },
    { label: t.totalProduction, value: fmt(totalProduction) + ' Kea', color: '#8b5cf6', tint: 'rgba(139,92,246,0.1)', icon: 'factory' },
    {
      label: finalGrowthLabel,
      value: finalGrowthValue,
      color: '#f59e0b',
      tint: 'rgba(245,158,11,0.1)',
      icon: 'zap',
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
    if (filteredRecords.length === 0 || !plotlyReady) return;

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
        // Fallback khi không có dòng TTL: tổng hợp từ monthly non-TTL
        if (Object.keys(salesByYear).length === 0) {
          filteredRecords.filter(r => r.division === 'sales' && r.month !== 'TTL').forEach(r => {
            salesByYear[r.year] = (salesByYear[r.year] || 0) + r.value;
          });
        }
        if (Object.keys(shipByYear).length === 0) {
          filteredRecords.filter(r => r.division === 'shipment' && r.month !== 'TTL').forEach(r => {
            shipByYear[r.year] = (shipByYear[r.year] || 0) + r.value;
          });
        }
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

      // EPCC (chart1-max-12-periods) — giới hạn Chart 1 (Xuất hàng & Doanh số)
      // tối đa 12 điểm GẦN NHẤT, bất kể đang xem theo Tháng/Quý/Năm. Trước đây
      // vẽ TOÀN BỘ khoảng năm lọc (VD 2017-2026 theo Quý = ~40 cột) khiến trục
      // X dồn nén, nhãn số đè lên nhau, không đọc được (đúng vùng khoanh đỏ
      // trong ảnh lỗi người dùng gửi). Luôn ưu tiên giai đoạn GẦN NHẤT vì đó
      // là thông tin người dùng quan tâm nhất khi xem xu hướng ngắn hạn.
      const MAX_TREND_POINTS = 12;
      if (xList.length > MAX_TREND_POINTS) {
        xList = xList.slice(-MAX_TREND_POINTS);
        yShip = yShip.slice(-MAX_TREND_POINTS);
        ySales = ySales.slice(-MAX_TREND_POINTS);
        yGrowth = yGrowth.slice(-MAX_TREND_POINTS);
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
          line: { color: '#00a65a', width: 4, shape: 'spline', smoothing: 1.1 },
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
          line: { color: '#f39c12', width: 3, dash: 'dash', shape: 'spline', smoothing: 1.1 },
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
        margin: { l: 54, r: 54, t: 16, b: 36 },
        showlegend: false,
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

      // EPCC (unified-chart-ylabel-headroom) — thêm khoảng đệm phía trên trục Y chính
      // để nhãn số (text) của điểm cao nhất (VD năm 2023) không bị kẹt sát mép/mất chữ
      const maxYUnified = Math.max(...yShip, ...ySales, 1);
      (layoutUnified.yaxis as any).range = [0, maxYUnified * 1.3];

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
              showlegend: false
            }
          ];
          const layoutModelUnified = {
            paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
            font: { family: 'Plus Jakarta Sans, sans-serif', color: chartTextColor, size: 11 },
            margin: { l: 50, r: 16, t: 16, b: 36 },
            showlegend: false,
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

      // EPCC (chart1-max-12-periods) — cùng lý do với Chart 1 phía trên: giới
      // hạn Chart 3 (Doanh số theo Khách hàng) tối đa 12 giai đoạn GẦN NHẤT,
      // tránh trục X dồn quá nhiều cột khi khoảng năm lọc rộng.
      if (custKeys.length > MAX_TREND_POINTS) {
        custKeys = custKeys.slice(-MAX_TREND_POINTS);
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
        margin: { l: 50, r: 16, t: 16, b: 36 },
        barmode: 'stack',
        showlegend: false,
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
            domain: { x: [0.06, 0.94], y: [0.06, 0.94] },
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
            margin: { l: 10, r: 10, t: 10, b: 10 },
            showlegend: false,
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
  }, [filteredRecords, chartFilteredRecords, viewMode, theme, lang, t, selectedPeriod, selectedPeriodLabel, plotlyReady]);

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
      {/* Header ngang hàng — Lang+Theme đã chuyển vào Sidebar (dùng chung
          cho Mục 1-4), không lặp lại riêng ở đây nữa. */}
      <div className="dashboard-header-grid">
        <div className="dashboard-header-left" style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-2)', fontWeight: 700, whiteSpace: 'nowrap' }}>
          <span aria-hidden="true">🕐</span>
          {formattedTime}
        </div>
        <h1 className="dashboard-header-title">
          {t.mainTitleDash}
        </h1>
        <div className="dashboard-header-right" />
      </div>

      <div className="dash-container">
        {/* Filter bar exactly matching reference layout
            FIX (EPCC-vanilla-toolbar): nền thanh filter đổi từ tông tối
            (topbar-dash mặc định) sang "Vanilla" (#FFF4D6) theo ảnh tham
            chiếu — áp dụng cho Mục 1/2/3. Các nút ĐANG ĐƯỢC CHỌN (Tháng/Quý/
            Năm...) vẫn giữ nguyên màu xanh/teal hiện có để phân biệt rõ lựa
            chọn hiện tại, không đổi theo nền vanilla này. */}
        <div className="topbar-dash" style={{
          display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '10px',
          background: '#2F3A1D',
          borderRadius: '14px', padding: '10px 14px',
          border: '1px solid rgba(0,0,0,0.18)',
          boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
        }}>
          {/* Dòng 1 (labels) */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
            {/* Cụm trái dòng 1: đồng hồ đã chuyển vào Sidebar — giữ spacer
                trống cùng bề rộng để không lệch layout các cột nhãn bên phải. */}
            <div style={{ width: '170px', flexShrink: 0 }} />
            {/* Cụm giữa dòng 1: Labels */}
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px' }}>
              <span style={{ width: '90px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: VANILLA_LABEL_COLOR, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.fromYear}</span>
              <span style={{ width: '90px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: VANILLA_LABEL_COLOR, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.toYear}</span>
              <span style={{ width: '110px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: VANILLA_LABEL_COLOR, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.partner}</span>
              <span style={{ width: '110px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: VANILLA_LABEL_COLOR, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.colModel}</span>
              <span style={{ width: '180px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: VANILLA_LABEL_COLOR, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{lang === 'vi' ? 'XEM THEO' : lang === 'ko' ? '보기 방식' : 'VIEW BY'}</span>
              <span style={{ width: '140px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: VANILLA_LABEL_COLOR, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{lang === 'vi' ? 'CHI TIẾT' : lang === 'ko' ? '상세 선택' : 'DETAIL'}</span>
              {activeFilterCount > 0 && <span style={{ width: '130px', flexShrink: 0 }}></span>}
              {activeFilterCount > 0 && <span style={{ width: '60px', flexShrink: 0 }}></span>}
            </div>
            {/* Cụm phải dòng 1: Spacers matching Dòng 2 (Tải tệp lên + Save to Cloud) */}
            <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
              <div style={{ width: '130px' }}></div>
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
                      // FIX (EPCC-vanilla-toolbar-contrast): nút KHÔNG active
                      // trước dùng var(--border)/var(--bg-card)/var(--text) —
                      // các biến này đổi theo theme và có lúc gần trùng màu
                      // nền Vanilla khiến chữ gần như biến mất (VD "Tuần",
                      // "Quý" mờ hẳn khi đổi theme). Chuyển sang màu CỐ ĐỊNH,
                      // không phụ thuộc theme, đảm bảo luôn đọc được trên nền
                      // Vanilla ở cả Light lẫn Dark.
                      border: '1px solid rgba(0,0,0,0.18)',
                      borderRight: mode !== 'year' ? 'none' : '1px solid rgba(0,0,0,0.18)',
                      background: viewMode === mode ? '#2e7d8c' : 'rgba(255,255,255,0.55)',
                      color: viewMode === mode ? '#ffffff' : '#7A5A2E',
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
                <NeonButton 
                  className="btn btn-ghost btn-sm" 
                  onClick={handleResetFilters}
                  style={{ height: '38px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}
                >
                  {t.resetBtn}
                </NeonButton>
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
              <input
                ref={fileInputRef}
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
              <NeonButton
                className="btn btn-outline btn-sm"
                onClick={() => fileInputRef.current?.click()}
                style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px', height: '38px', width: '130px', boxSizing: 'border-box', fontSize: '13px' }}
              >
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15" style={{ flexShrink: 0 }}>
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <path d="M17 8l-5-5-5 5" />
                  <path d="M12 3v12" />
                </svg>
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {lang === 'vi' ? 'Tải tệp lên' : lang === 'ko' ? '파일 업로드' : 'Upload File'}
                </span>
              </NeonButton>
            </div>
          </div>
        </div>

        {/* KPI Grid — 4 cards, compact style matching Dashboard mới */}
        {/*
          FIX (EPCC-kpi-row-fill-width): trước đây flex-basis đặt cứng theo px
          (160px / 240px) khiến phần không gian dư ra sau khi cấp basis không
          được chia hết theo đúng tỉ lệ flex-grow trong một số trường hợp độ
          rộng container, để lại khoảng trống trắng bên phải hàng KPI (xem ảnh
          khoanh đỏ). Đổi flex-basis về '0%' (co giãn thuần theo tỉ lệ grow,
          minWidth vẫn giữ làm sàn chống co quá nhỏ) + ép container width:100%
          để hàng KPI luôn lấp đầy hết chiều rộng, không còn ô trống phía phải.
          FIX (quy-chuan-4-card-deu-nhau): trước đây card 4 (growth) có
          flex-grow 1.4 + min-width 240px trong khi 3 card còn lại chỉ có
          flex-grow 1 + min-width 160px → card 4 luôn rộng hơn hẳn 3 card
          kia (lệch chiều dài rõ rệt, xem ảnh khoanh đỏ). Đổi cả 4 card về
          cùng flex: '1 1 0%' + minWidth: '200px' để chia đều 100% chiều
          rộng container, đảm bảo 4 card luôn bằng nhau về chiều dài bất kể
          kích thước màn hình.
          - Title vẫn dùng white-space:nowrap + textOverflow:ellipsis nên
            nếu màn hình quá hẹp, tiêu đề dài (card 4) sẽ tự cắt gọn bằng
            "…" thay vì đẩy card rộng ra — không còn lệch layout.
        */}
        <div style={{
          display: 'flex',
          flexDirection: 'row',
          gap: '12px',
          marginBottom: '10px',
          flexWrap: 'nowrap',
          alignItems: 'stretch',
          width: '100%'
        }}>
          {kpis.map((c, i) => {
            const isGrowth = i === 3; // last card is the growth card
            return (
              // FIX (EPCC-add-kpi-card-class-muc1): mục 1 trước đây là div
              // TRẦN, không có className="kpi-card" — nên KHÔNG thừa hưởng
              // được các hiệu ứng mặc định của class chuẩn (box-shadow/glow
              // phía trên, transition...) mà mục 2 & mục 3 đều có (vì cả 2
              // đều dùng className="kpi-card"). Thêm đúng className="kpi-card"
              // ở đây — các style inline (border/background/padding/radius)
              // vẫn override đúng phần cần custom theo màu từng card, còn
              // phần chưa set inline (shadow, transition...) sẽ tự động lấy
              // đúng mặc định giống hệt mục 2, đảm bảo đồng nhất hoàn toàn.
              <div key={i} className="kpi-card" style={{
                flex: '1 1 0%',
                minWidth: '200px',
                border: `1px solid ${c.tint}`,
                borderLeft: `4px solid ${c.color}`,
                // FIX (EPCC-unify-kpi-size-all-sections): copy đúng kích thước
                // (padding + border-radius) từ className="kpi-card" chuẩn mà
                // mục 2 và mục 3 đang dùng (index.css: padding 7px 10px,
                // border-radius var(--radius-md) = 12px) — trước đây mục 1 tự
                // đặt padding '12px 14px' / borderRadius '10px' riêng, lệch
                // chuẩn so với 2 mục kia.
                borderRadius: 'var(--radius-md)',
                background: `linear-gradient(135deg, ${c.tint} 0%, rgba(30, 41, 59, 0.4) 100%)`,
                padding: '7px 10px',
                animationDelay: `${i * 0.05}s`,
                position: 'relative',
                display: 'flex',
                flexDirection: 'column',
                gap: '0px',
                boxSizing: 'border-box'
              }}>
                {/* Header row: icon + label (nowrap) + toggle
                    FIX (EPCC-unify-kpi-typography-all-sections): copy đúng
                    margin-bottom: 8px (thay cho gap:6px ở card cha) + cỡ chữ
                    label cố định 13.5px (bỏ clamp responsive) từ mục 2
                    (TargetActualDashboard) để 3 mục có cỡ chữ trên-dưới và
                    khoảng cách header→value giống hệt nhau. */}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%', gap: '4px', minWidth: 0, marginBottom: '8px' }}>
                  {/* Left: icon + label */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px', minWidth: 0, flex: 1, overflow: 'hidden' }}>
                    {renderKpiIcon(c.icon, c.color)}
                    <div style={{
                      fontSize: '13.5px',
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
            <div className="panel-head" style={chartHeaderStyle(0)}>
              <h3 style={{ fontSize: '15px', fontWeight: '600', margin: 0 }}>
                {viewMode === 'year' 
                  ? t.chart1Title 
                  : viewMode === 'quarter' 
                    ? (lang === 'vi' ? 'Xuất hàng & Doanh số theo Quý' : lang === 'en' ? 'Shipment & Sales by Quarter' : '분기별 출하 & 매출')
                    : t.chartMonthTitle
                }
              </h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <SDLegendItem type="bar" color="#2d7f96" label={lang === 'vi' ? 'XUẤT HÀNG (K)' : lang === 'en' ? 'SHIPMENT (K)' : '출하 (K)'} />
                <SDLegendItem type="line" color="#00a65a" label={lang === 'vi' ? 'DOANH SỐ (K$)' : lang === 'en' ? 'SALES AMT (K$)' : '매출 (K$)'} />
                <SDLegendItem type="line" color="#f39c12" label={trendGrowthLegendLabel} />
              </div>
            </div>
            <div className="chart-holder" id="trendChartUnified"></div>
          </div>

          {/* Unified Top Models Chart */}
          <div className="panel">
            <div className="panel-head" style={chartHeaderStyle(1)}>
              <h3 style={{ fontSize: '15px', fontWeight: '600', margin: 0 }}>
                {lang === 'vi' 
                  ? `Top 10 Model theo Doanh số - ${selectedPeriodLabel}` 
                  : lang === 'en' 
                    ? `Top 10 Models by Sales - ${selectedPeriodLabel}` 
                    : `매출 Top 10 모델 - ${selectedPeriodLabel}`}
              </h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <SDLegendItem type="bar" color="#0891b2" label={`SALES (K$) — ${selectedPeriodLabel}`} />
              </div>
            </div>
            <div className="chart-holder" id="topModelsUnified"></div>
          </div>
        </div>

        {/* Charts Row 2 */}
        <div className="charts-row row-2">
          {/* Customer Stacked (Left) */}
          <div className="panel">
            <div className="panel-head" style={chartHeaderStyle(2)}>
              <h3 style={{ fontSize: '15px', fontWeight: '600', margin: 0 }}>
                {lang === 'vi' 
                  ? 'Doanh số theo Khách hàng (CUSTOM)' 
                  : lang === 'en' 
                    ? 'Customer Sales Trend (CUSTOM)' 
                    : '고객사별 매출 (CUSTOM)'}
              </h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                {CUSTOMER_LEGEND_ORDER.map(c => (
                  <SDLegendItem key={c} type="bar" color={CUSTOMER_COLORS[c]} label={c} />
                ))}
              </div>
            </div>
            <div className="chart-holder" id="customerSalesStackedUnified"></div>
          </div>

            {/* Customer Pinwheel (Right) */}
            <div className="panel">
              <div className="panel-head" style={chartHeaderStyle(3)}>
                <h3 style={{ fontSize: '15px', fontWeight: '600', margin: 0 }}>
                  {lang === 'vi'
                    ? `Tỷ trọng theo Khách hàng - ${selectedPeriodLabel}`
                    : lang === 'en'
                      ? `Customer Sales Share - ${selectedPeriodLabel}`
                      : `고객사별 매출 비중 - ${selectedPeriodLabel}`}
                </h3>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  {CUSTOMER_LEGEND_ORDER.map(c => (
                    <SDLegendItem key={c} type="dot" color={CUSTOMER_COLORS[c]} label={c} />
                  ))}
                </div>
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

      {/* FIX (charts-full-size + header-flush-edges, EPCC):
          1. "dòng đầu tiên (panel-head) để top=0, left=0, right=0" — panel
             bao ngoài (class .panel, CSS global) đang có padding riêng
             khiến thanh màu tiêu đề bị thụt vào, không chạm sát 3 cạnh
             trên/trái/phải (đúng 3 mũi tên khoanh đỏ ở góc mỗi khung trong
             ảnh). Fix: đưa padding của .panel về 0 (ghi đè bằng specificity
             cao ".sales-dashboard" x3, không sửa CSS global dùng chung Mục
             khác) để panel-head — vốn là block-level div đầu tiên trong
             .panel — tự nhiên nằm sát top:0/left:0/right:0. Dồn phần
             padding đã bỏ đi sang .chart-holder để nội dung biểu đồ bên
             trong vẫn có khoảng đệm trái/phải/dưới như cũ, không bị dính
             sát mép.
          2. "nới rộng biểu đồ to hết mức có thể" — tăng tiếp chiều cao
             chart-holder (520px → 640px) so với lần chỉnh trước. */}
      <style>{`
        .sales-dashboard.sales-dashboard.sales-dashboard .charts-row .panel {
          height: auto !important;
          max-height: none !important;
          padding: 0 !important;
          overflow: hidden;
        }
        .sales-dashboard.sales-dashboard.sales-dashboard .charts-row .panel-head {
          margin: 0 !important;
        }
        .sales-dashboard.sales-dashboard.sales-dashboard .charts-row .chart-holder {
          width: 100% !important;
          height: 640px !important;
          max-height: none !important;
          box-sizing: border-box;
          padding: 14px 16px 16px 16px !important;
          margin: 0 !important;
        }
        @media (max-width: 1024px) {
          .sales-dashboard.sales-dashboard.sales-dashboard .charts-row .chart-holder {
            height: 480px !important;
          }
        }
        /* FIX (donut-center, EPCC): hack cũ dùng margin-top âm để "nhích"
           khung donut đã đẩy nó vượt ra ngoài vùng padding của
           .chart-holder — vì .panel cha có overflow:hidden nên phần rìa
           trên của vòng donut bị cắt/che mờ ở góc. Thay bằng flex-center
           chuẩn: khung #customerSalesDonutUnified tự canh giữa theo cả 2
           trục trong không gian sẵn có, kết hợp domain 6% đệm ở trace pie
           (xem traceCustomerDonut) để vòng tròn không bao giờ chạm biên.
           Chỉ áp dụng riêng khung donut này, không đụng 3 biểu đồ còn lại. */
        #customerSalesDonutUnified {
          display: flex !important;
          align-items: center !important;
          justify-content: center !important;
          width: 100% !important;
          height: 100% !important;
        }
      `}</style>
    </div>
  );
};
