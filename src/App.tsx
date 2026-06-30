import { useState, useEffect, useMemo, useCallback, useRef } from 'react';
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
import { TargetActualDashboard } from './components/TargetActualDashboard';
import { SalesDashboard } from './components/SalesDashboard';
import { Sidebar } from './components/Sidebar';
import { GlobalHeaderControls } from './components/GlobalHeaderControls';
// ── Manpower + Per Capita tabs (Tab 2 nằm trong ManpowerDashboard) ──────────
import { ManpowerDashboard } from './components/ManpowerDashboard';
// ─────────────────────────────────────────────────────
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
  const [activeViewId, setActiveViewId] = useState<string>('overview');
  const [sidebarCollapsed, setSidebarCollapsed] = useState<boolean>(true);
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

      // ── Nếu đã có dữ liệu upload tay (Manpower/PerCapita...) thì dùng cache local,
      //    KHÔNG fetch Supabase đè lên — tránh mất data khi reload/restart dev server.
      let useLocalOnly = false;
      try {
        useLocalOnly = localStorage.getItem('manual_upload_flag') === '1';
      } catch (e) { /* ignore */ }

      if (useLocalOnly && typeof window !== 'undefined') {
        try {
          const cached = localStorage.getItem('cached_sales_data');
          if (cached) {
            finalData = JSON.parse(cached);
            console.log('Loaded from manual-upload cache (skip Supabase to preserve uploaded data)');
          }
        } catch (e) { console.error(e); }
      }

      if (!finalData && supabase) {
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
            try { localStorage.setItem('cached_sales_data', JSON.stringify(finalData)); } catch (e) { console.error('Cache limit exceeded'); }
          }
        } catch (err) {
          console.error('Supabase exception:', err);
          hasError = true;
        }
      } else if (!finalData) {
        hasError = true;
      }

      if (!finalData && (hasError) && typeof window !== 'undefined') {
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
        const firstRowKeys = Object.keys(finalData[0]).map(k => k.trim());
        setHeaders(firstRowKeys);
        setFilename(useLocalOnly ? 'Manual Upload (Cached)' : hasError ? 'Cached Local Database' : 'Supabase Cloud Database');
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
    
    try {
      localStorage.setItem('cached_sales_data', JSON.stringify(parsedRows));
      localStorage.setItem('manual_upload_flag', '1');
    } catch (e) { console.error('Cache limit exceeded'); }


    let autoDate = types.find(t => t.type === 'date')?.name || NONE;
    // ... (rest of auto-detect mapping logic unchanged)
    setMapping({
      dateCol: autoDate,
      categoryCol: NONE,
      revenueCol: NONE,
      costCol: NONE,
      currency: 'VND',
    });
    setScreen('dashboard');
  }, []);

  const handleFileSelected = useCallback((file: File, wb: any) => {
    setFilename(file.name);
    if (wb.SheetNames.length > 0) {
      parseSheet(wb, wb.SheetNames[0]);
    }
  }, [parseSheet]);

  const handleReset = useCallback(() => {
    setScreen('upload');
    setAllRows([]);
    setHeaders([]);
    resetChannelColors?.();
    try {
      localStorage.removeItem('manual_upload_flag');
      localStorage.removeItem('cached_sales_data');
    } catch (e) { /* ignore */ }
  }, []);

  const dateBounds = useMemo(() => getDateBounds(allRows, mapping.dateCol), [allRows, mapping.dateCol]);

  const filteredRows = useMemo(
    () => applyFilters(allRows, mapping, filters),
    [allRows, mapping, filters]
  );

  const kpi = useMemo(() => computeKPI(filteredRows, mapping), [filteredRows, mapping]);

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

  const isTargetActual = useMemo(() => {
    const headersLower = headers.map(h => h.toLowerCase().trim());
    const headersStr = headersLower.join(' ');
    
    const hasTarget = headersStr.includes('target') || headersStr.includes('목표');
    const hasActual = headersStr.includes('actual') || headersStr.includes('실적');
    if (hasTarget && hasActual) return true;

    const hasModel = headersLower.includes('model');
    const hasDivision = headersLower.includes('division');
    const hasValue = headersLower.includes('value');
    const hasDate = headersLower.includes('date');
    const hasType = headersLower.includes('type') || headersLower.includes('type1');
    const hasCustom = headersLower.includes('custom') || headersLower.includes('type2');

    if (hasModel && hasDivision && hasValue && hasDate && hasType && hasCustom) {
      return true;
    }

    return false;
  }, [headers]);

  // ── detect manpower data (Test_3 style) ─────────────────────────────────
  // Test_3.xlsx: Type = "TTL ManPower AVG" → chứa "manpower" → route sang ManpowerDashboard.
  // Per Capita (Tab 2) nằm trong ManpowerDashboard, không cần detect riêng.
  const isManpower = useMemo(() => {
    const headersLower = headers.map(h => h.toLowerCase().trim());
    const hasModel = headersLower.includes('model');
    const hasType  = headersLower.includes('type');
    const hasDate  = headersLower.includes('date');
    const hasValue = headersLower.includes('value');
    if (hasModel && hasType && hasDate && hasValue) {
      return allRows.some(r =>
        String(r.type || r.Type || '').toLowerCase().includes('manpower')
      );
    }
    return false;
  }, [headers, allRows]);
  // ─────────────────────────────────────────────────────────────────────────

  const dashboardMode = useMemo(() => {
    if (isManpower) return 'manpower';
    if (isTargetActual) return 'target_actual';
    const headersStr = headers.join(' ').toLowerCase();
    const isSales = headersStr.includes('division') || headersStr.includes('model') || headersStr.includes('qty') || headersStr.includes("q'ty") || headersStr.includes('value');
    return isSales ? 'sales' : 'marketing';
  }, [headers, isTargetActual, isManpower]);

  // ── Auto-switch view chỉ khi lần đầu load data (allRows từ rỗng → có data).
  //    KHÔNG override khi user đang xem Mục 1 rồi upload file Mục 3,
  //    vì điều đó sẽ làm Mục 1 bị đẩy sang Mục 3 (bug cũ).
  const prevRowsLengthRef = useRef<number>(0);
  useEffect(() => {
    const wasEmpty = prevRowsLengthRef.current === 0;
    const nowHasData = allRows.length > 0;
    prevRowsLengthRef.current = allRows.length;

    // Chỉ tự động chuyển view khi đây là lần đầu có data (initial load)
    if (wasEmpty && nowHasData) {
      if (dashboardMode === 'manpower') {
        setActiveViewId('manpower');
      } else if (dashboardMode === 'target_actual') {
        setActiveViewId('target_actual');
      } else {
        setActiveViewId('overview');
      }
    }
    // Nếu user đã chọn view thủ công (sidebar click) → giữ nguyên, không override
  }, [allRows, dashboardMode]);

  return (
    <>
      <div className="bg-aura" />

      {screen === 'dashboard' ? (
        <div className={`app-layout ${sidebarCollapsed ? 'sidebar-collapsed' : ''}`}>

          <Sidebar
            activeViewId={activeViewId}
            onSelectView={(id) => {
              setActiveViewId(id);
              setSidebarCollapsed(true);
            }}
            collapsed={sidebarCollapsed}
            onToggleCollapse={() => setSidebarCollapsed(!sidebarCollapsed)}
            lang={lang}
          />
          
          <main className="app-content" style={{ paddingTop: '48px', position: 'relative' }}>

          {/* ── Global header bar: Lang + Theme ─────────────────────────────────
              Dùng position:fixed nhưng giới hạn left = sidebar width, để nó
              chỉ nổi trên vùng content (không đè sidebar). app-content có
              paddingTop = đúng chiều cao bar này → nội dung KHÔNG BAO GIỜ
              bị che dù scroll bao nhiêu, vì bar không nằm trong luồng scroll
              của content nhưng content đã "nhường" đúng khoảng trống cho nó. */}
          <div
            style={{
              position: 'fixed',
              top: 0,
              left: sidebarCollapsed ? 68 : 260,
              right: 0,
              height: '48px',
              zIndex: 500,
              display: 'flex',
              justifyContent: 'flex-end',
              alignItems: 'center',
              padding: '6px 20px',
              boxSizing: 'border-box',
              background: 'var(--surface, #fff)',
              borderBottom: '1px solid var(--border-soft, #e5e7eb)',
              boxShadow: '0 1px 3px rgba(0,0,0,0.07)',
              transition: 'left 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
            }}
          >
            <GlobalHeaderControls
              lang={lang}
              setLang={setLang}
              isDark={theme === 'dark'}
              onToggleTheme={toggleTheme}
            />
          </div>

            {/* ── MENU 1: Sales Dashboard ──
                Chỉ render khi user chủ động chọn 'overview' (Mục 1).
                KHÔNG dùng dashboardMode để quyết định — tránh bug hiện Mục 3 khi
                allRows là Manpower nhưng user đang ở Mục 1. */}
            {activeViewId === 'overview' && (
              <SalesDashboard
                rows={allRows}
                mapping={mapping}
                theme={theme}
                onBack={() => {}}
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

            {/* ── MENU 3: Manpower Dashboard (có Tab 2 Per Capita bên trong) ── */}
            {activeViewId === 'manpower' && (
              <ManpowerDashboard
                rows={allRows}
                theme={theme}
                lang={lang}
                onToggleTheme={toggleTheme}
                setLang={setLang}
                onFileSelected={handleFileSelected}
              />
            )}

            {(activeViewId === 'target_actual') && (
              <TargetActualDashboard
                rows={allRows}
                theme={theme}
                onToggleTheme={toggleTheme}
                lang={lang}
                setLang={setLang}
                onFileSelected={handleFileSelected}
              />
            )}

            {activeViewId === 'placeholder' && (
              <div style={{ padding: '40px', background: 'var(--surface)', margin: '24px', borderRadius: '12px', border: '1px solid var(--border-soft)' }}>
                <h2 style={{ color: 'var(--text-0)' }}>
                  {lang === 'vi' ? '4. Cấu hình & Khác' : lang === 'en' ? '4. Config & Others' : '4. 설정 및 기타'}
                </h2>
                <p style={{ color: 'var(--text-2)', marginTop: '8px' }}>
                  {lang === 'vi' 
                    ? 'Trang này hiện chưa có nội dung biểu đồ. Bạn có thể dễ dàng thêm mới các module khác tại đây.' 
                    : lang === 'en' 
                      ? 'This page does not have chart content yet. You can easily add other modules here.'
                      : '이 페이지에는 아직 차트 콘텐츠가 없습니다. 여기에 다른 모듈을 쉽게 추가할 수 있습니다.'}
                </p>
              </div>
            )}

            {dashboardMode === 'marketing' && (
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
                  <ChartsSection rows={allRows} mapping={mapping} theme={theme} />
                  <DetailTable rows={filteredRows} headers={headers} mapping={mapping} />
                </main>
              </div>
            )}
          </main>
        </div>
      ) : (
        screen === 'upload' ? (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '100vh', padding: '20px' }}>
            <div style={{ background: 'var(--surface)', padding: '40px', borderRadius: '12px', border: '1px solid var(--border-soft)', maxWidth: '480px', width: '100%', textAlign: 'center', boxShadow: 'var(--shadow-sm)' }}>
              <div style={{ fontSize: '48px', marginBottom: '16px' }}>📊</div>
              <h2 style={{ color: 'var(--text-0)', marginBottom: '8px' }}>IM-VINA Dashboards</h2>
              <p style={{ color: 'var(--text-2)', marginBottom: '24px' }}>
                {lang === 'vi' ? 'Vui lòng tải lên tệp Excel dữ liệu để bắt đầu.' : 'Please upload an Excel data file to start.'}
              </p>
              <label className="btn btn-outline" style={{ cursor: 'pointer', display: 'inline-flex', padding: '12px 24px', fontSize: '15px' }}>
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
                        handleFileSelected(file, workbook);
                      } catch (err) {
                        alert('Error reading Excel file');
                      }
                    };
                    reader.readAsArrayBuffer(file);
                  }}
                />
                📁 {lang === 'vi' ? 'Tải tệp Excel' : 'Upload Excel'}
              </label>
            </div>
          </div>
        ) : null
      )}
    </>
  );
}
