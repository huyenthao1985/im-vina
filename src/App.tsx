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
  const [, setSyncStatus] = useState<'idle' | 'syncing' | 'synced' | 'error'>('idle');

  // Cờ loading cho lần tải dữ liệu đầu tiên từ Supabase/cache. Trong lúc
  // headers vẫn còn rỗng, dashboardMode sẽ mặc định rơi vào 'marketing' —
  // nếu không có cờ này, người dùng sẽ thấy màn "Marketing Insights" (rỗng)
  // lóe lên trước khi dashboard thật hiện ra. true = đang tải lần đầu.
  const [isInitialDataLoading, setIsInitialDataLoading] = useState(true);

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
    // ── BƯỚC 0: Hiển thị NGAY dữ liệu cache local (nếu có) trong lúc chờ
    // Supabase tải bản mới nhất ở nền — kiểu "stale-while-revalidate". Trước
    // đây mỗi lần F5 đều phải đợi Supabase tải xong (nhiều giây) mới hiện gì
    // đó, dù máy này đã từng tải thành công và có sẵn cache. Giờ cache hiện
    // ra gần như tức thì, Supabase âm thầm cập nhật lại phía sau khi xong —
    // người dùng thấy dữ liệu (có thể hơi cũ vài giây) ngay lập tức thay vì
    // nhìn spinner 12-15s mỗi lần.
    try {
      const cachedBuckets = localStorage.getItem('cached_dashboard_buckets');
      if (cachedBuckets) {
        const parsed = JSON.parse(cachedBuckets);
        const sales = parsed.sales || [];
        const manpower = parsed.manpower || [];
        const targetActual = parsed.targetActual || [];
        const combined = [...sales, ...manpower, ...targetActual];
        if (combined.length > 0) {
          setSalesRows(sales);
          setManpowerRows(manpower);
          setTargetActualRows(targetActual);
          setAllRows(combined);
          setHeaders(Object.keys(combined[0]).map(k => k.trim()));
          setFilename('Cached Local Database');
          setScreen('dashboard');
          hasHydratedRef.current = true;
          setIsInitialDataLoading(false);
          console.log('Hiện ngay từ cache local, đang đồng bộ Supabase ở nền...');
        }
      }
    } catch (e) { /* ignore */ }

    async function loadSupabaseData() {

      // ── BƯỚC 1: Luôn tải Supabase để lấy bản mới nhất — nếu BƯỚC 0 đã
      // hiện cache, việc này chạy NGẦM phía sau, không che dashboard đang
      // hiển thị (chỉ âm thầm thay dữ liệu khi có bản mới hơn).
      if (supabase) {
        // Gán ra const cục bộ: TypeScript chỉ giữ được narrowing (khác null)
        // của `supabase` cho tới khi hết block try/catch NGOÀI cùng — bên
        // trong closure lồng (Array.from callback bên dưới), TS coi lại là
        // "có thể null" vì biến gốc có thể bị gán lại trước khi closure chạy.
        // `db` là const nên narrowing được giữ nguyên trong mọi closure.
        const db = supabase;
        try {
          let allData: any[] = [];
          const PAGE_SIZE = 1000;

          // TIMEOUT 10 giây: nếu Supabase mất hơn 10s, bỏ qua và dùng cache.
          // Đây là nguyên nhân chính khiến F5 màn hình chờ >1 phút khi có
          // hàng chục nghìn dòng hoặc kết nối chậm.
          const SUPABASE_TIMEOUT_MS = 10_000;
          // Dùng `any` thay vì generic <T> để tránh lỗi TSX parse "<T>" → JSX
          const withTimeout = (promise: Promise<any>): Promise<any> =>
            Promise.race([
              promise,
              new Promise<null>((_resolve, reject) =>
                setTimeout(() => reject(new Error('Supabase timeout')), SUPABASE_TIMEOUT_MS)
              ),
            ]).catch((err: Error) => {
              console.warn('Supabase bị timeout hoặc lỗi:', err.message);
              return null;
            });

          // Lấy tổng số dòng trước (head request).
          const countResult = await withTimeout(
            Promise.resolve(db.from('sales_data').select('*', { count: 'exact', head: true }))
          );

          if (!countResult) {
            // Timeout → dùng cache đang hiển thị, không block nữa.
            console.warn('Supabase count timeout — giữ cache local.');
            hasHydratedRef.current = true;
            setIsInitialDataLoading(false);
            return;
          }

          const { count, error: countError } = countResult;
          if (countError) {
            console.error('Supabase count error:', countError);
          } else if (count && count > 0) {
            const totalPages = Math.ceil(count / PAGE_SIZE);

            const fetchPage = (i: number) => {
              const from = i * PAGE_SIZE;
              return withTimeout(
                Promise.resolve(
                  db
                    .from('sales_data')
                    .select('*')
                    .order('id', { ascending: true })
                    .range(from, from + PAGE_SIZE - 1)
                )
              );
            };

            const CONCURRENCY = 4;
            const pageIndexes = Array.from({ length: totalPages }, (_, i) => i);
            let timedOut = false;
            for (let i = 0; i < pageIndexes.length; i += CONCURRENCY) {
              if (timedOut) break;
              const batch = pageIndexes.slice(i, i + CONCURRENCY);
              const results = await Promise.all(batch.map(fetchPage));
              for (const result of results) {
                if (!result) { timedOut = true; break; }
                const { data, error } = result as any;
                if (error) {
                  console.error('Supabase fetch error (1 trang):', error);
                  continue;
                }
                if (data) {
                  allData = allData.concat(
                    data.map((d: any) => ({
                      ...(d.payload && typeof d.payload === 'object' ? d.payload : d),
                      source_tag: d.source_tag,
                      date: d.date ?? (d.payload?.date) ?? d.month ?? (d.payload?.month),
                    }))
                  );
                }
              }
            }
            if (timedOut) {
              console.warn(`Supabase page fetch timeout tại trang ${Math.floor(allData.length / PAGE_SIZE)} — dùng cache.`);
              hasHydratedRef.current = true;
              setIsInitialDataLoading(false);
              return;
            }
          }

          if (allData.length > 0) {
            // Supabase có dữ liệu → dùng ngay (thay thế cache nếu đang hiện),
            // cập nhật lại cache local cho lần F5 sau.
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
            setIsInitialDataLoading(false);
            console.log('Đã đồng bộ dữ liệu mới nhất từ Supabase:', allData.length, 'rows');
            return;
          }
          console.log('Supabase connected but table is empty — giữ nguyên cache (nếu có) hoặc để trống');
        } catch (err) {
          console.error('Supabase exception:', err);
        }
      } else {
        console.warn('Supabase not configured — giữ nguyên cache (nếu có) hoặc để trống');
      }

      // Nếu BƯỚC 0 đã hiện cache thành công mà Supabase lại lỗi/rỗng, KHÔNG
      // xoá dashboard đang hiển thị — cứ để nguyên dữ liệu cache đó.
      // Nếu chưa hề có cache nào (lần đầu mở máy này) thì để trống.
      hasHydratedRef.current = true;
      setIsInitialDataLoading(false);
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

  // ── Đồng bộ dữ liệu upload lên Supabase ─────────────────────────────────
  // Chiến lược "replace-theo-bucket": mỗi lần upload thay thế toàn bộ dữ
  // liệu của đúng bucket đó trên Supabase, không cộng dồn/trùng lặp.
  //
  // ── FIX ROOT CAUSE "biểu đồ Mục 3 không hiện sau F5" ────────────────────
  // Phiên bản cũ dùng { source_tag, payload: r } nhưng bảng sales_data có
  // NOT NULL trên model/division/year/month/value. Mọi INSERT đều lỗi
  // "null value in column 'model' violates not-null constraint" — dữ liệu
  // KHÔNG BAO GIỜ lên được Supabase, biểu đồ F5 luôn trống.
  // Fix: map Excel columns → schema columns đúng cột, đúng NOT NULL.
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
      const deleteQuery = supabase.from('sales_data').delete();
      const { error: deleteError } = bucketTag === 'Sales'
        ? await deleteQuery.or('source_tag.eq.Sales,source_tag.is.null')
        : await deleteQuery.eq('source_tag', bucketTag);

      if (deleteError) {
        console.error(`Supabase delete error (${bucketTag}):`, deleteError);
        setSyncStatus('error');
        return;
      }

      // 2) Helper: lấy giá trị field case-insensitive (Excel thường Capitalize đầu)
      const MONTH_ABBRS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
      function guessYear(rawDate: any): number {
        if (rawDate instanceof Date) return rawDate.getFullYear();
        const s = String(rawDate ?? '').trim().toUpperCase();
        if (MONTH_ABBRS.includes(s) || /^W\d{1,2}$/.test(s) || /^\d{1,2}[/-]\d{1,2}$/.test(s)) {
          return new Date().getFullYear();
        }
        const n = parseInt(s, 10);
        return (!isNaN(n) && n > 2000 && n < 2100) ? n : new Date().getFullYear();
      }

      function gv(r: any, ...keys: string[]): any {
        for (const k of keys) {
          const v = r[k] ?? r[k.charAt(0).toUpperCase() + k.slice(1)] ?? r[k.toLowerCase()];
          if (v !== undefined && v !== null && v !== '') return v;
        }
        return undefined;
      }

      // 3) Map Excel rows → schema columns theo từng bucket
      let taggedRows: Record<string, any>[];

      if (bucketTag === 'Manpower') {
        // Chỉ lưu dòng nhân lực (type chứa 'manpower') và dòng sản xuất
        // (division = SUB1/SUB2/MAIN và type chứa 'plan'/'actual').
        // Lọc bỏ toàn bộ dòng rác không thuộc 2 loại này để tránh database bị phình.
        const filteredRows = rows.filter(r => {
          const ts  = String(gv(r, 'type', 'Type') ?? '').toLowerCase();
          const div = String(gv(r, 'division', 'Division') ?? '').trim().toUpperCase();
          const isManpowerType = ts.includes('manpower') || ts.includes('인당생산수');
          const isProdRow = ['SUB1','SUB2','MAIN'].includes(div) &&
                            (ts.includes('plan') || ts.includes('actual'));
          return isManpowerType || isProdRow;
        });
        taggedRows = filteredRows.map(r => {
          const rawDate = gv(r, 'date', 'Date', 'month', 'Month') ?? '';
          const divVal  = String(gv(r, 'division', 'Division') ?? 'production').trim();
          return {
            source_tag: 'Manpower',
            model:    String(gv(r, 'model', 'Model') ?? '').trim() || 'N/A',
            origin:   'Manpower',
            customer: String(gv(r, 'customer', 'Customer') ?? '').trim(),
            type:     String(gv(r, 'type', 'Type') ?? '').trim(),
            division: divVal || 'production',
            year:     guessYear(rawDate),
            month:    String(rawDate).trim() || 'N/A',
            value:    Number(gv(r, 'value', 'Value') ?? 0),
          };
        });
      } else if (bucketTag === 'TargetActual') {
        taggedRows = rows.map(r => {
          const rawDate = gv(r, 'date', 'Date', 'month', 'Month') ?? '';
          return {
            source_tag: 'TargetActual',
            model:    String(gv(r, 'model', 'Model') ?? '').trim() || 'N/A',
            origin:   'TargetActual',
            customer: String(gv(r, 'customer', 'Customer') ?? '').trim(),
            type:     String(gv(r, 'type', 'Type', 'type1', 'Type1') ?? '').trim(),
            division: String(gv(r, 'division', 'Division') ?? '').trim() || 'sales',
            year:     Number(gv(r, 'year', 'Year') ?? guessYear(rawDate)),
            month:    String(rawDate).trim() || 'N/A',
            value:    Number(gv(r, 'value', 'Value') ?? 0),
          };
        });
      } else {
        // Sales bucket
        taggedRows = rows.map(r => ({
          source_tag: 'Sales',
          model:    String(gv(r, 'model', 'Model') ?? '').trim() || 'N/A',
          origin:   String(gv(r, 'origin', 'Origin', '출하지') ?? '').trim(),
          customer: String(gv(r, 'customer', 'Customer', 'custom', 'Custom', 'CUSTOM') ?? '').trim(),
          type:     String(gv(r, 'type', 'Type', 'type1', 'Type1') ?? '').trim(),
          division: String(gv(r, 'division', 'Division') ?? '').trim() || 'sales',
          year:     Number(gv(r, 'year', 'Year') ?? new Date().getFullYear()),
          month:    String(gv(r, 'month', 'Month') ?? '').trim() || 'N/A',
          value:    Number(gv(r, 'value', 'Value', "q'ty", "Q'TY") ?? 0),
        }));
      }

      // 4) Insert theo từng lô để tránh vượt giới hạn payload Supabase.
      const CHUNK_SIZE = 500;
      for (let i = 0; i < taggedRows.length; i += CHUNK_SIZE) {
        const chunk = taggedRows.slice(i, i + CHUNK_SIZE);
        const { error: insertError } = await supabase.from('sales_data').insert(chunk);
        if (insertError) {
          console.error(`Supabase insert error (${bucketTag}) tại lô dòng ${i}:`, insertError);
          alert(
            `⚠️ Đồng bộ Cloud thất bại (${bucketTag}, dòng ${i}–${i + chunk.length}).\n` +
            `Dữ liệu vẫn hiển thị tạm trên máy này nhưng sẽ MẤT khi F5.\n` +
            `Chi tiết lỗi: ${insertError.message}`
          );
          setSyncStatus('error');
          return;
        }
      }

      setSyncStatus('synced');
      console.log(`Đã đồng bộ ${taggedRows.length} dòng (${bucketTag}) lên Supabase`);
    } catch (err) {
      console.error(`Supabase upsert exception (${bucketTag}):`, err);
      setSyncStatus('error');
    }
  }, []);

  // Parse TOÀN BỘ workbook (mọi sheet), không chỉ sheet đầu tiên.
  // ── FIX "Mục 3 không hiện dữ liệu ở Sheet 2" ────────────────────────────
  // Root cause: hàm cũ (parseSheet) chỉ đọc đúng 1 sheet duy nhất —
  // wb.SheetNames[0] — do handleFileSelected luôn gọi parseSheet(wb,
  // wb.SheetNames[0]). File Manpower thường có Sheet 1 chứa dòng
  // type="...ManPower AVG" (nuôi Tab 1 - 근무 인력 현황) và Sheet 2 chứa dòng
  // type="인당생산수" (nuôi Tab 2 - Per Capita) — Sheet 2 bị bỏ qua hoàn
  // toàn nên PerCapitaTab luôn thấy rows rỗng → "Không có dữ liệu nhân lực"
  // dù Tab 1 vẫn có dữ liệu bình thường (vì Tab 1 chỉ cần Sheet 1).
  // Fix: duyệt qua TẤT CẢ wb.SheetNames, gộp toàn bộ rows (và union headers)
  // của mọi sheet lại làm MỘT trước khi phân loại bucket — vì trong thực tế
  // 1 lần upload luôn là 1 file cho 1 dashboard/bucket, chỉ khác nhau ở việc
  // dữ liệu được tách ra nhiều sheet cho gọn (theo tầng ngày/tuần/tháng hay
  // theo loại chỉ số), không phải nhiều bucket khác nhau trong cùng 1 file.
  const parseWorkbook = useCallback((wb: any) => {
    const XLSX_lib = XLSX;
    const sheetNames: string[] = wb.SheetNames || [];

    let combinedHeaders: string[] = [];
    let combinedRows: DataRow[] = [];

    sheetNames.forEach((sheetName: string) => {
      const sheet = wb.Sheets[sheetName];
      if (!sheet) return;
      const rawHeaders = XLSX_lib.utils.sheet_to_json(sheet, { header: 1, defval: null })[0] as any[] || [];
      const filteredHeaders = rawHeaders
        .filter((h: any) => h !== null && h !== undefined && String(h).trim() !== '')
        .map((h: any) => String(h).trim());
      if (filteredHeaders.length === 0) return; // sheet rỗng/không có header → bỏ qua

      const parsedRowsRaw = XLSX_lib.utils.sheet_to_json(sheet, { defval: null }) as DataRow[];
      const parsedRows = parsedRowsRaw.map(row => {
        const trimmedRow: DataRow = {};
        Object.keys(row).forEach(key => {
          trimmedRow[key.trim()] = row[key];
        });
        return trimmedRow;
      });
      if (parsedRows.length === 0) return;

      // Union headers: giữ thêm cột mới nếu sheet sau có cột mà sheet trước
      // chưa có (VD Sheet2 chỉ thêm vài cột phụ so với Sheet1).
      filteredHeaders.forEach(h => { if (!combinedHeaders.includes(h)) combinedHeaders.push(h); });
      combinedRows = combinedRows.concat(parsedRows);
    });

    return { headers: combinedHeaders, rows: combinedRows };
  }, []);

  // Giữ tên `parseSheet` cho phần còn lại của pipeline (mapping, bucket
  // routing...) không đổi — chỉ thay nguồn headers/rows đầu vào từ "1 sheet"
  // sang "toàn bộ workbook đã gộp".
  const parseSheet = useCallback((wb: any) => {
    const { headers: filteredHeaders, rows: parsedRows } = parseWorkbook(wb);
    const types = detectColumns(parsedRows, filteredHeaders);

    setHeaders(filteredHeaders);
    setAllRows(parsedRows);

    // Route dữ liệu (đã gộp mọi sheet) vào ĐÚNG bucket của nó, KHÔNG đụng tới
    // 2 bucket còn lại — đây là fix trực tiếp cho bug "upload Manpower thì
    // Mục 1 mất data". Định tuyến data vào đúng bucket VÀ tự động chuyển
    // sang tab phù hợp sau khi user upload thủ công. Supabase initial load
    // KHÔNG tự chuyển (đã bỏ auto-switch effect) → trang luôn mở ở Mục 1
    // mặc định.
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
  }, [parseWorkbook, syncBucketToSupabase]);

  const handleFileSelected = useCallback((file: File, wb: any) => {
    setFilename(file.name);
    if (wb.SheetNames.length > 0) {
      parseSheet(wb);
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
          
          <main className="app-content" style={{ position: 'relative' }}>

          {/* ── Global header bar: Lang + Theme ─────────────────────────────────
              LỊCH SỬ FIX:
              (1) Bar gốc dùng position:fixed + height CỨNG 48px khớp tay với
                  paddingTop CỨNG 48px trên app-content → 2 số lệch nhau làm
                  control bên trong tràn xuống, đè lên hàng đầu mọi tab.
              (2) Đã đổi sang position:'sticky' để tự "chiếm chỗ" đúng theo
                  chiều cao thật — hết lỗi tràn (1). NHƯNG sticky vẫn giữ
                  hành vi "dính ở đầu khi cuộn" (giống fixed): nội dung phía
                  dưới cuộn LÊN thì vẫn chui XUỐNG DƯỚI bar này về mặt thị
                  giác → đúng như phản ánh "cuộn xuống là đè lên các mục
                  khác". Người dùng không cần bar này dính lại khi cuộn.
              FIX CUỐI: bỏ hẳn position đặc biệt — để bar nằm trong luồng
              layout bình thường (static/relative mặc định) như mọi phần tử
              khác. Nó chỉ hiện ở đầu trang và CUỘN ĐI CÙNG nội dung, không
              bao giờ còn nằm "nổi" trên bất cứ thứ gì nữa. */}
          {activeViewId === 'placeholder' && (
            <div
              style={{
                display: 'flex',
                justifyContent: 'flex-end',
                alignItems: 'center',
                flexWrap: 'wrap',
                padding: '6px 20px',
                boxSizing: 'border-box',
                background: 'var(--surface, #fff)',
                borderBottom: '1px solid var(--border-soft, #e5e7eb)',
              }}
            >
              <GlobalHeaderControls
                lang={lang}
                setLang={setLang}
                isDark={theme === 'dark'}
                onToggleTheme={toggleTheme}
              />
            </div>
          )}

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
              <div style={{ padding: '40px', background: 'var(--surface)', margin: '24px', borderRadius: '12px', border: '1px solid var(--border-soft)', textAlign: 'center' }}>
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

            {activeViewId === 'overview' && isInitialDataLoading && (
              <div
                style={{
                  display: 'flex', flexDirection: 'column', alignItems: 'center',
                  justifyContent: 'center', minHeight: '60vh', gap: '16px',
                }}
              >
                <div className="spinner" style={{
                  width: '36px', height: '36px', borderRadius: '50%',
                  border: '3px solid var(--border-soft, #e5e7eb)',
                  borderTopColor: 'var(--primary, #6366f1)',
                  animation: 'spin 0.8s linear infinite',
                }} />
                <p style={{ color: 'var(--text-2)' }}>
                  {lang === 'vi' ? 'Đang tải dữ liệu...' : lang === 'en' ? 'Loading data...' : '데이터 로딩 중...'}
                </p>
                <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
              </div>
            )}

            {activeViewId === 'overview' && !isInitialDataLoading && dashboardMode === 'marketing' && (
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
