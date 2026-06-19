import { useState, useEffect, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import type {
  DataRow, ColumnMapping, FilterState, KPIData, ThemeMode, ScreenState
} from './types';
import { supabase } from './lib/supabase';
import {
  detectColumns, toNumber, getDateBounds, resetChannelColors, parseRowToDate
} from './utils';

import { KPIGrid } from './components/KPIGrid';
import { ChartsSection } from './components/ChartsSection';
import { FilterBar } from './components/FilterBar';
import { DetailTable } from './components/DetailTable';
import { SalesDashboard } from './components/SalesDashboard';
import './App.css';

const NONE = '__none__';



function computeKPI(rows: DataRow[], mapping: ColumnMapping): KPIData {
  const hasRevenue = mapping.revenueCol !== NONE;
  const hasCost = mapping.costCol !== NONE;
  const hasCategory = mapping.categoryCol !== NONE;

  const totalRevenue = hasRevenue ? rows.reduce((a, r) => a + toNumber(r[mapping.revenueCol]), 0) : 0;
  const totalCost = hasCost ? rows.reduce((a, r) => a + toNumber(r[mapping.costCol]), 0) : 0;
  const totalProfit = totalRevenue - totalCost;
  const roi = totalCost > 0 ? ((totalProfit / totalCost) * 100) : 0;
  const uniqueCategories = hasCategory
    ? new Set(rows.map(r => r[mapping.categoryCol]).filter(v => v !== null && v !== undefined && v !== '')).size
    : 0;

  return {
    totalRevenue,
    totalCost,
    totalProfit,
    roi,
    totalRows: rows.length,
    uniqueCategories,
    hasCost: hasCost && totalCost > 0,
  };
}

function applyFilters(rows: DataRow[], mapping: ColumnMapping, filters: FilterState): DataRow[] {
  let result = rows;
  const hasDate = mapping.dateCol !== NONE;
  const hasCategory = mapping.categoryCol !== NONE;

  if (hasDate && filters.dateStart) {
    const start = new Date(filters.dateStart);
    result = result.filter(r => {
      const d = parseRowToDate(r, mapping.dateCol);
      return d !== null && d >= start;
    });
  }
  if (hasDate && filters.dateEnd) {
    const end = new Date(filters.dateEnd);
    end.setHours(23, 59, 59, 999);
    result = result.filter(r => {
      const d = parseRowToDate(r, mapping.dateCol);
      return d !== null && d <= end;
    });
  }
  if (hasCategory && filters.categories.length > 0) {
    result = result.filter(r => filters.categories.includes(String(r[mapping.categoryCol])));
  }
  return result;
}

function countActiveFilters(filters: FilterState): number {
  let count = 0;
  if (filters.dateStart) count++;
  if (filters.dateEnd) count++;
  if (filters.categories.length > 0) count++;
  return count;
}

export default function App() {
  const [theme, setTheme] = useState<ThemeMode>('light');
  const [lang, setLang] = useState<'vi' | 'en' | 'ko'>('vi');
  const [screen, setScreen] = useState<ScreenState>('dashboard');
  const [filename, setFilename] = useState('');

  // Data state
  const [allRows, setAllRows] = useState<DataRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<ColumnMapping>({
    dateCol: NONE, categoryCol: NONE, revenueCol: NONE, costCol: NONE, currency: 'VND',
  });

  // Load from Supabase on mount
  useEffect(() => {
    async function loadSupabaseData() {
      let finalData = null;
      let hasError = false;

      if (supabase) {
        try {
          let allData: any[] = [];
          let from = 0;
          const PAGE_SIZE = 1000;
          let keepFetching = true;

          while (keepFetching) {
            const { data, error } = await supabase
              .from('sales_data')
              .select('*')
              .order('year', { ascending: true })
              .range(from, from + PAGE_SIZE - 1);
              
            if (error) {
              console.error('Supabase fetch error:', error);
              hasError = true;
              break;
            } else if (data && data.length > 0) {
              allData = [...allData, ...data];
              from += PAGE_SIZE;
              if (data.length < PAGE_SIZE) {
                keepFetching = false;
              }
            } else {
              keepFetching = false;
            }
          }

          if (!hasError && allData.length > 0) {
            finalData = allData;
            // Cache to localStorage
            try { localStorage.setItem('cached_sales_data', JSON.stringify(finalData)); } catch (e) { console.error('Cache limit exceeded'); }
          }
        } catch (err) {
          console.error('Supabase exception:', err);
          hasError = true;
        }
      } else {
        hasError = true;
      }

      // If Supabase failed (e.g. project paused), try loading from local cache
      if ((hasError || !finalData) && typeof window !== 'undefined') {
        try {
          const cached = localStorage.getItem('cached_sales_data');
          if (cached) {
            finalData = JSON.parse(cached);
            console.log('Loaded from local cache (Supabase paused or offline)');
          }
        } catch (e) { console.error(e); }
      }
      
      if (finalData && finalData.length > 0) {
        setAllRows(finalData);
        setHeaders(['model', 'origin', 'customer', 'type', 'division', 'year', 'month', 'value']);
        setFilename(hasError ? 'Cached Local Database' : 'Supabase Cloud Database');
        setScreen('dashboard');
      }
    }
    loadSupabaseData();
  }, []);

  // Filters
  const [filters, setFilters] = useState<FilterState>({ dateStart: '', dateEnd: '', categories: [] });

  // Theme sync
  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  const toggleTheme = () => setTheme(t => t === 'dark' ? 'light' : 'dark');

  // Parse sheet from workbook
  const parseSheet = useCallback((wb: any, sheetName: string) => {
    const XLSX_lib = XLSX;
    const sheet = wb.Sheets[sheetName];
    const rawHeaders = XLSX_lib.utils.sheet_to_json(sheet, { header: 1, defval: null })[0] as any[] || [];
    const filteredHeaders = rawHeaders
      .filter((h: any) => h !== null && h !== undefined && String(h).trim() !== '')
      .map((h: any) => String(h).trim());
    const parsedRowsRaw = XLSX_lib.utils.sheet_to_json(sheet, { defval: null }) as DataRow[];
    const parsedRows = parsedRowsRaw.map(row => {
      const trimmedRow: DataRow = {};
      Object.keys(row).forEach(key => {
        trimmedRow[key.trim()] = row[key];
      });
      return trimmedRow;
    });
    const types = detectColumns(parsedRows, filteredHeaders);

    setHeaders(filteredHeaders);
    setAllRows(parsedRows);
    setAllRows(parsedRows);
    
    // Also save uploaded data to cache so it persists even if Supabase is offline
    try { localStorage.setItem('cached_sales_data', JSON.stringify(parsedRows)); } catch (e) { console.error('Cache limit exceeded'); }

    // Auto-detect mapping
    const allColNames = types.map(t => t.name);

    // 1. Date Col
    let autoDate = types.find(t => t.type === 'date')?.name || NONE;
    if (autoDate === NONE) {
      const dateKeywords = ['ngày', 'date', 'thời gian', 'time', 'year', 'năm', 'nam', 'month', 'tháng', 'thang'];
      const found = allColNames.find(name => dateKeywords.some(k => name.toLowerCase().includes(k)));
      if (found) autoDate = found;
    }

    // 2. Category Col
    let autoCategory = types.find(t => t.type === 'text')?.name || NONE;
    if (autoCategory === NONE || autoCategory.toLowerCase().includes('ngày') || autoCategory.toLowerCase().includes('date')) {
      const catKeywords = ['kênh', 'channel', 'danh mục', 'category', 'set', 'model', 'chiến dịch', 'campaign', 'nhóm', 'group'];
      const found = allColNames.find(name => catKeywords.some(k => name.toLowerCase().includes(k)));
      if (found) autoCategory = found;
    }

    // 3. Revenue & Cost Cols
    const metricKeywords = ['doanh thu', 'revenue', 'value', 'amount', 'sản lượng', 'thành tiền', 'sales', 'tiền', 'price'];
    const costKeywords = ['chi phí', 'cost', 'expense', 'fee', 'chi'];

    const possibleMetrics = types.filter(t => {
      const nameLower = t.name.toLowerCase();
      const isTimeOrDimension = ['year', 'năm', 'nam', 'month', 'tháng', 'thang', 'ngày', 'date', 'time', 'id', 'code', 'stt'].some(k => nameLower.includes(k));
      return t.type === 'numeric' && !isTimeOrDimension;
    }).map(t => t.name);

    let autoRevenue = NONE;
    let autoCost = NONE;

    const foundRev = types.find(t => metricKeywords.some(k => t.name.toLowerCase().includes(k)))?.name;
    const foundCost = types.find(t => costKeywords.some(k => t.name.toLowerCase().includes(k)))?.name;

    if (foundRev) {
      autoRevenue = foundRev;
    } else if (possibleMetrics.length > 0) {
      autoRevenue = possibleMetrics[0];
    }

    if (foundCost) {
      autoCost = foundCost;
    } else if (possibleMetrics.length > 1) {
      autoCost = possibleMetrics[1];
    } else if (possibleMetrics.length > 0 && possibleMetrics[0] !== autoRevenue) {
      autoCost = possibleMetrics[0];
    }

    const autoMap: ColumnMapping = {
      dateCol: autoDate,
      categoryCol: autoCategory,
      revenueCol: autoRevenue,
      costCol: autoCost,
      currency: 'VND',
    };
    setMapping(autoMap);
  }, []);

  const handleFileSelected = useCallback((file: File, wb: any) => {
    setFilename(file.name);
    const sheet = wb.SheetNames[0];
    parseSheet(wb, sheet);
    // Skip mapping screen and go directly to dashboard
    resetChannelColors();
    setFilters({ dateStart: '', dateEnd: '', categories: [] });
    setScreen('dashboard');
  }, [parseSheet]);

  const handleReset = useCallback(() => {
    setScreen('upload');
    setAllRows([]);
    setHeaders([]);
    setMapping({ dateCol: NONE, categoryCol: NONE, revenueCol: NONE, costCol: NONE });
    setFilters({ dateStart: '', dateEnd: '', categories: [] });
  }, []);

  // Filtered rows
  const filteredRows = useMemo(() => applyFilters(allRows, mapping, filters), [allRows, mapping, filters]);

  // KPIs from filtered rows
  const kpi = useMemo(() => computeKPI(filteredRows, mapping), [filteredRows, mapping]);

  // Date bounds from all rows
  const dateBounds = useMemo(() => {
    if (mapping.dateCol === NONE) return { min: null, max: null };
    return getDateBounds(allRows, mapping.dateCol);
  }, [allRows, mapping.dateCol]);

  // All categories
  const allCategories = useMemo(() => {
    if (mapping.categoryCol === NONE) return [];
    return Array.from(
      new Set(allRows.map(r => r[mapping.categoryCol]).filter(v => v !== null && v !== undefined && v !== '').map(String))
    ).sort();
  }, [allRows, mapping.categoryCol]);

  const activeFilterCount = useMemo(() => countActiveFilters(filters), [filters]);

  const resetFilters = useCallback(() => {
    setFilters({ dateStart: '', dateEnd: '', categories: [] });
  }, []);

  const dashboardMode = useMemo(() => {
    const headersStr = headers.join(' ').toLowerCase();
    const isSales = headersStr.includes('division') || headersStr.includes('model') || headersStr.includes('qty') || headersStr.includes("q'ty") || headersStr.includes('value');
    return isSales ? 'sales' : 'marketing';
  }, [headers]);

  return (
    <>
      <div className="bg-aura" />







      {screen === 'dashboard' && dashboardMode === 'sales' && (
        <SalesDashboard
          rows={filteredRows}
          mapping={mapping}
          theme={theme}
          onBack={handleReset}
          onToggleTheme={toggleTheme}
          filters={filters}
          onFilterChange={setFilters}
          minDate={dateBounds.min}
          maxDate={dateBounds.max}
          allCategories={allCategories}
          lang={lang}
          setLang={setLang}
          onFileSelected={handleFileSelected}
        />
      )}

      {screen === 'dashboard' && dashboardMode === 'marketing' && (
        <div className="dashboard-screen">
          <header className="topbar">
            <div className="topbar-left">
              <div className="brandmark">
                <div className="brandmark-dot">📊</div>
                <span className="brandmark-text">Marketing Insights</span>
              </div>
            </div>
            <div className="topbar-right">
              <div className="filename-pill">
                <span>📁</span>
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{filename}</span>
              </div>

              <button className="btn btn-ghost btn-sm" onClick={handleReset}>
                ← Tải file khác
              </button>
              <button className="theme-toggle" onClick={toggleTheme}>
                {theme === 'dark' ? '☀️' : '🌙'}
              </button>
            </div>
          </header>

          <main className="dash-body">
            <FilterBar
              filters={filters}
              onFilterChange={setFilters}
              categories={allCategories}
              minDate={dateBounds.min}
              maxDate={dateBounds.max}
              mapping={mapping}
              activeCount={activeFilterCount}
              onReset={resetFilters}
            />

            <KPIGrid kpi={kpi} mapping={mapping} />

            <ChartsSection rows={filteredRows} mapping={mapping} theme={theme} />

            <DetailTable rows={filteredRows} headers={headers} mapping={mapping} />
          </main>
        </div>
      )}
    </>
  );
}
