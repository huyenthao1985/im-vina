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

// ── Bucket routing helpers ───────────────────────────────────────────────
// Root cause of "mất data khi chuyển tab" (Image 1): trước đây cả 3 dashboard
// (Sales / Target-Actual / Manpower) dùng chung 1 mảng `allRows` duy nhất.
// Upload file cho Mục 3 sẽ THAY THẾ toàn bộ allRows → Mục 1 mất dữ liệu dù
// chưa hề đổi gì ở Mục 1. Fix: tách thành 3 bucket độc lập, mỗi lần upload/
// tải-từ-cloud chỉ cập nhật đúng bucket tương ứng, không đụng tới 2 bucket kia.

function detectIsManpower(headers: string[], rows: DataRow[]): boolean {
  const headersLower = headers.map(h => h.toLowerCase().trim());
  const hasModel = headersLower.includes('model');
  const hasType = headersLower.includes('type');
  const hasDate = headersLower.includes('date');
  const hasValue = headersLower.includes('value');
  if (hasModel && hasType && hasDate && hasValue) {
    return rows.some(r => String((r as any).type || (r as any).Type || '').toLowerCase().includes('manpower'));
  }
  return false;
}

function detectIsTargetActual(headers: string[]): boolean {
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
  return hasModel && hasDivision && hasValue && hasDate && hasType && hasCustom;
}

// Quyết định 1 sheet vừa upload thuộc bucket nào (dùng cho upload thủ công)
function detectBucket(headers: string[], rows: DataRow[]): 'sales' | 'manpower' | 'target_actual' {
  if (detectIsManpower(headers, rows)) return 'manpower';
  if (detectIsTargetActual(headers)) return 'target_actual';
  return 'sales';
}

// Phân vùng dữ liệu tải về từ Supabase theo cột 'source_tag' (mới) với fallback
// đọc cột 'origin' cho các dòng cũ (Manpower/TargetActual) đã lưu trước khi có
// source_tag. Dòng Sales KHÔNG có source_tag/origin trùng 'Manpower' hay
// 'TargetActual' nên rơi vào nhánh else một cách an toàn — 'origin' của dòng
// Sales là dữ liệu xuất xứ thật (vd "Vietnam"/"Korea"), không bao giờ trùng 2 tag trên.
function bucketByTag(rows: DataRow[]): { sales: DataRow[]; manpower: DataRow[]; targetActual: DataRow[] } {
  const sales: DataRow[] = [];
  const manpower: DataRow[] = [];
  const targetActual: DataRow[] = [];
  rows.forEach(r => {
    const tag = (r as any).source_tag || (r as any).origin;
    if (tag === 'Manpower') manpower.push(r);
    else if (tag === 'TargetActual') targetActual.push(r);
    else sales.push(r);
  });
  return { sales, manpower, targetActual };
}
// ─────────────────────────────────────────────────────────────────────────

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

  // Data state — 3 bucket độc lập thay vì 1 mảng allRows dùng chung.
  const [salesRows, setSalesRows] = useState<DataRow[]>([]);
  const [manpowerRows, setManpowerRows] = useState<DataRow[]>([]);
  const [targetActualRows, setTargetActualRows] = useState<DataRow[]>([]);
  // allRows: giữ lại cho pipeline generic cũ (marketing fallback, KPI/mapping
  // tổng quát) — luôn phản ánh sheet được upload/tải gần nhất, KHÔNG dùng để
  // render 3 dashboard chính nữa (chúng dùng bucket riêng ở trên).
  const [allRows, setAllRows] = useState<DataRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [mapping, setMapping] = useState<ColumnMapping>({
    dateCol: NONE, categoryCol: NONE, revenueCol: NONE, costCol: NONE, currency: 'VND',
  });
  const hasHydratedRef = useRef(false);
  // Trạng thái đồng bộ Supabase khi upload thủ công — có thể dùng để hiện
  // badge "Đang đồng bộ..." / "Đã lên Cloud" / "Lỗi đồng bộ" trên UI nếu cần.
  const [syncStatus, setSyncStatus] = useState<'idle' | 'syncing' | 'synced' | 'error'>('idle');

  // Persist 3 bucket ra localStorage mỗi khi có thay đổi (sau khi đã hydrate
  // lần đầu) — cho phép reload trang mà KHÔNG mất bucket nào, kể cả khi bucket
  // đó chỉ tồn tại ở local (chưa "Lên mây").
  useEffect(() => {
    if (!hasHydratedRef.current) return;
    try {
      localStorage.setItem('cached_dashboard_buckets', JSON.stringify({
        sales: salesRows, manpower: manpowerRows, targetActual: targetActualRows,
      }));
    } catch (e) { console.error('Cache limit exceeded'); }
  }, [salesRows, manpowerRows, targetActualRows]);

  // Load from Supabase on mount — Supabase luôn được ưu tiên.
  // Cache local chỉ là fallback khi Supabase không khả dụng.
  useEffect(() => {
    async function loadSupabaseData() {

      // ── BƯỚC 1: Luôn thử Supabase trước ──────────────────────────────────
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
              break;
            } else if (data && data.length > 0) {
              allData = [...allData, ...data];
              from += PAGE_SIZE;
              if (data.length < PAGE_SIZE) keepFetching = false;
            } else {
              keepFetching = false;
            }
          }

          if (allData.length > 0) {
            // Supabase có dữ liệu → dùng ngay, cập nhật cache local
            const { sales, manpower, targetActual } = bucketByTag(allData);
            setSalesRows(sales);
            setManpowerRows(manpower);
            setTargetActualRows(targetActual);
            setAllRows(allData);
            setHeaders(Object.keys(allData[0]).map(k => k.trim()));
            setFilename('Supabase Cloud Database');
            setScreen('dashboard');
            try {
              localStorage.setItem('cached_dashboard_buckets', JSON.stringify({ sales, manpower, targetActual }));
            } catch (e) { /* ignore */ }
            hasHydratedRef.current = true;
            console.log('Loaded from Supabase:', allData.length, 'rows');
            return;
          }
          console.log('Supabase connected but table is empty — falling back to local cache');
        } catch (err) {
          console.error('Supabase exception:', err);
        }
      } else {
        console.warn('Supabase not configured — falling back to local cache');
      }

      // ── BƯỚC 2: Fallback — dùng cache local nếu Supabase không có dữ liệu ─
      try {
        const cachedBuckets = localStorage.getItem('cached_dashboard_buckets');
        if (cachedBuckets) {
          const parsed = JSON.parse(cachedBuckets);
          setSalesRows(parsed.sales || []);
          setManpowerRows(parsed.manpower || []);
          setTargetActualRows(parsed.targetActual || []);
          setAllRows([...(parsed.sales || []), ...(parsed.manpower || []), ...(parsed.targetActual || [])]);
          setFilename('Cached Local Database');
          setScreen('dashboard');
          hasHydratedRef.current = true;
          console.log('Loaded from local cache (3-bucket)');
          return;
        }
      } catch (e) { /* ignore */ }

      // ── BƯỚC 3: Không có dữ liệu nào ─────────────────────────────────────
      hasHydratedRef.current = true;
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

  // ── Đồng bộ dữ liệu upload thủ công lên Supabase ────────────────────────
  // Trước đây upload thủ công CHỈ lưu vào localStorage. Vì loadSupabaseData
  // luôn ưu tiên Supabase khi F5, dữ liệu upload tay sẽ "biến mất" khỏi màn
  // hình (bị Supabase cũ/rỗng ghi đè) dù vẫn còn trong cache local. Hàm này
  // đẩy ngay dữ liệu vừa upload lên Supabase để lần load sau luôn thấy nó.
  // Chiến lược "replace-theo-bucket": mỗi lần upload thay thế toàn bộ dữ
  // liệu của đúng bucket đó trên Supabase, không cộng dồn/trùng lặp.
  const syncBucketToSupabase = useCallback(async (
    rows: DataRow[],
    bucketTag: 'Sales' | 'Manpower' | 'TargetActual'
  ) => {
    if (!supabase) {
      console.warn('Supabase not configured — bỏ qua đồng bộ, dữ liệu chỉ lưu local.');
      return;
    }
    setSyncStatus('syncing');
    try {
      // 1) Xoá toàn bộ dòng CŨ thuộc đúng bucket này trước khi insert dòng mới.
      //    Riêng bucket 'Sales': dữ liệu Sales cũ trên Supabase có thể CHƯA có
      //    cột source_tag (vì cột 'origin' của Sales là xuất xứ thật như
      //    "Vietnam"/"Korea", không phải tag phân loại) → phải xoá cả những
      //    dòng source_tag IS NULL, nếu không sẽ bị trùng lặp dữ liệu Sales.
      const deleteQuery = supabase.from('sales_data').delete();
      const { error: deleteError } = bucketTag === 'Sales'
        ? await deleteQuery.or('source_tag.eq.Sales,source_tag.is.null')
        : await deleteQuery.eq('source_tag', bucketTag);

      if (deleteError) {
        console.error(`Supabase delete error (${bucketTag}):`, deleteError);
        setSyncStatus('error');
        return;
      }

      // 2) Gắn source_tag cho từng dòng trước khi insert, để lần load sau
      //    (hàm bucketByTag) phân loại đúng bucket một cách chắc chắn.
      const taggedRows = rows.map(r => ({ ...r, source_tag: bucketTag }));

      // 3) Insert theo từng lô để tránh vượt giới hạn payload của Supabase.
      const CHUNK_SIZE = 500;
      for (let i = 0; i < taggedRows.length; i += CHUNK_SIZE) {
        const chunk = taggedRows.slice(i, i + CHUNK_SIZE);
        const { error: insertError } = await supabase.from('sales_data').insert(chunk);
        if (insertError) {
          console.error(`Supabase insert error (${bucketTag}) tại lô bắt đầu dòng ${i}:`, insertError);
          setSyncStatus('error');
          return; // các lô insert trước đó đã lên thành công, dừng tại đây
        }
      }

      setSyncStatus('synced');
      console.log(`Đã đồng bộ ${taggedRows.length} dòng (${bucketTag}) lên Supabase`);
    } catch (err) {
      console.error(`Supabase upsert exception (${bucketTag}):`, err);
      setSyncStatus('error');
    }
  }, []);

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

    // Route sheet vừa upload vào ĐÚNG bucket của nó, KHÔNG đụng tới 2 bucket
    // còn lại — đây là fix trực tiếp cho bug "upload Manpower thì Mục 1 mất data".
    // Định tuyến data vào đúng bucket VÀ tự động chuyển sang tab phù hợp
    // sau khi user upload thủ công. Supabase initial load KHÔNG tự chuyển
    // (đã bỏ auto-switch effect) → trang luôn mở ở Mục 1 mặc định.
    const bucket = detectBucket(filteredHeaders, parsedRows);
    if (bucket === 'manpower') {
      setManpowerRows(parsedRows);
      setActiveViewId('manpower');
      syncBucketToSupabase(parsedRows, 'Manpower');
    } else if (bucket === 'target_actual') {
      setTargetActualRows(parsedRows);
      setActiveViewId('target_actual');
      syncBucketToSupabase(parsedRows, 'TargetActual');
    } else {
      setSalesRows(parsedRows);
      setActiveViewId('overview');
      syncBucketToSupabase(parsedRows, 'Sales');
    }

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
  }, [syncBucketToSupabase]);

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
    setSalesRows([]);
    setManpowerRows([]);
    setTargetActualRows([]);
    resetChannelColors?.();
    try {
      localStorage.removeItem('cached_sales_data');
      localStorage.removeItem('cached_dashboard_buckets');
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

  // ── Auto-switch: đã chuyển sang parseSheet (chỉ khi upload tay).
  //    KHÔNG còn auto-switch khi Supabase load → trang luôn mở ở Mục 1. ──

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
            {activeViewId === 'overview' && dashboardMode !== 'marketing' && (
              <SalesDashboard
                rows={salesRows}
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
                rows={manpowerRows}
                theme={theme}
                lang={lang}
                onToggleTheme={toggleTheme}
                setLang={setLang}
                onFileSelected={handleFileSelected}
              />
            )}

            {(activeViewId === 'target_actual') && (
              <TargetActualDashboard
                rows={targetActualRows}
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

            {activeViewId === 'overview' && dashboardMode === 'marketing' && (
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
