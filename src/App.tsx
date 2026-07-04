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

// ── Khoá nghiệp vụ (business key) cho bảng sales_data ───────────────────────
// Đây là bộ cột XÁC ĐỊNH DUY NHẤT 1 dòng dữ liệu (không tính `value` — value
// là số liệu, được phép thay đổi/ghi đè khi cập nhật). Dùng bộ cột NÀY cho cả:
//   1. UNIQUE CONSTRAINT phía Supabase (supabase_dedupe_and_constraint.sql)
//   2. onConflict của upsert() phía client (syncBucketToSupabase bên dưới)
// PHẢI giữ 2 nơi này khớp nhau tuyệt đối, nếu không upsert sẽ báo lỗi
// "there is no unique or exclusion constraint matching the ON CONFLICT".
const SALES_DATA_KEY = [
  'source_tag', 'model', 'origin', 'customer', 'type', 'division', 'year', 'month',
] as const;

// ── Cache 3 bucket dữ liệu bằng IndexedDB (KHÔNG dùng localStorage) ─────────
// FIX ROOT CAUSE "Mục 2/Mục 3 ra sai dữ liệu ngay sau F5": cache stale-while-
// revalidate (BƯỚC 0 bên dưới) trước đây dùng localStorage.setItem() để lưu
// RAW toàn bộ 3 bucket. localStorage mỗi origin chỉ có quota ~5-10MB — với
// ~100.000+ dòng dữ liệu hiện tại, JSON.stringify(...) vượt xa quota này nên
// setItem() luôn ném QuotaExceededError (log "Cache limit exceeded" trong
// console) và KHÔNG BAO GIỜ ghi cache mới được nữa — cache bị "đứng hình"
// vĩnh viễn ở snapshot CŨ NHẤT còn lọt quota (từ thời dữ liệu còn nhỏ, TRƯỚC
// các bản fix gv()/created_at/dedupe gần đây). Mỗi lần F5, BƯỚC 0 lại hiện
// NGAY snapshot cũ SAI đó trong vài giây (thời gian Supabase tải lại ~100k
// dòng ở nền) — đúng hiện tượng người dùng gặp ở Mục 2 (%Rate bất thường)
// và Mục 3 (dữ liệu tụt bất thường ở label cuối).
// Fix: chuyển cache sang IndexedDB — quota thực tế hàng trăm MB, KHÔNG còn
// bị giới hạn bởi quy mô dữ liệu hiện tại (và cả khi tăng thêm nhiều lần nữa).
const IDB_DB_NAME = 'imvina_dashboard_cache';
const IDB_STORE = 'kv';

function idbOpenCacheDb(): Promise<IDBDatabase | null> {
  return new Promise(resolve => {
    if (typeof indexedDB === 'undefined') { resolve(null); return; }
    try {
      const req = indexedDB.open(IDB_DB_NAME, 1);
      req.onupgradeneeded = () => { req.result.createObjectStore(IDB_STORE); };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    } catch (e) { resolve(null); }
  });
}

async function idbGetCache(key: string): Promise<string | null> {
  const db = await idbOpenCacheDb();
  if (!db) return null;
  return new Promise(resolve => {
    try {
      const tx = db.transaction(IDB_STORE, 'readonly');
      const req = tx.objectStore(IDB_STORE).get(key);
      req.onsuccess = () => resolve((req.result as string) ?? null);
      req.onerror = () => resolve(null);
    } catch (e) { resolve(null); }
  });
}

async function idbSetCache(key: string, value: string): Promise<boolean> {
  const db = await idbOpenCacheDb();
  if (!db) return false;
  return new Promise(resolve => {
    try {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    } catch (e) { resolve(false); }
  });
}

async function idbDelCache(key: string): Promise<void> {
  const db = await idbOpenCacheDb();
  if (!db) return;
  return new Promise(resolve => {
    try {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    } catch (e) { resolve(); }
  });
}

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

// (Đã bỏ hàm bucketByTag() phân loại-sau-khi-tải: bản fix "Manpower biến mất
// sau F5" giờ lọc RIÊNG theo bucket ngay tại query Supabase — xem
// applyBucketFilter()/fetchBucketRows() trong loadSupabaseData() bên dưới —
// nên không cần bước phân loại lại toàn bộ allData sau khi tải nữa.)
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

  // Persist 3 bucket ra cache (sau khi đã hydrate lần đầu) — cho phép reload
  // trang mà KHÔNG mất bucket nào, kể cả khi bucket đó chỉ tồn tại ở local
  // (chưa "Lên mây"). Dùng idbSetCache (IndexedDB) thay cho localStorage —
  // xem giải thích chi tiết ở khối comment cạnh idbOpenCacheDb() phía trên:
  // localStorage quota ~5-10MB không đủ chứa 100k+ dòng, khiến cache cũ
  // "đứng hình" vĩnh viễn và hiện sai dữ liệu ở mọi lần F5.
  useEffect(() => {
    if (!hasHydratedRef.current) return;
    idbSetCache('cached_dashboard_buckets', JSON.stringify({
      sales: salesRows, manpower: manpowerRows, targetActual: targetActualRows,
    })).catch(() => { /* IndexedDB không khả dụng (private mode...) — bỏ qua, không ảnh hưởng dashboard đang hiển thị */ });
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
    // Dọn cache localStorage kiểu CŨ (nếu còn sót lại từ trước bản fix này) —
    // key này không còn được dùng để đọc/viết nữa, xoá cho sạch, tránh gây
    // nhầm lẫn khi debug về sau.
    try { localStorage.removeItem('cached_dashboard_buckets'); } catch (e) { /* ignore */ }

    let cancelled = false;

    async function readLocalCacheAndShow() {
      try {
        const cachedBuckets = await idbGetCache('cached_dashboard_buckets');
        if (cancelled || !cachedBuckets) return;
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
          console.log('Hiện ngay từ cache local (IndexedDB), đang đồng bộ Supabase ở nền...');
        }
      } catch (e) { /* ignore */ }
    }

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
          const PAGE_SIZE = 1000;
          // ── FIX ROOT CAUSE "Manpower biến mất sau F5" ───────────────────
          // Bản cũ: 1 query DUY NHẤT order by id desc, cắt cứng 5000 dòng
          // cho TOÀN BẢNG (3 bucket gộp chung). Khi Sales/TargetActual có
          // id lớn hơn (ghi gần đây/nhiều hơn), 5000 dòng "mới nhất theo id"
          // bị Sales/TargetActual chiếm hết → Manpower bị loại sạch dù data
          // vẫn còn trong DB (đây là nguyên nhân thật của bug, KHÔNG phải do
          // ghi thất bại hay do sự cố hạ tầng Supabase).
          // Fix: fetch RIÊNG theo từng bucket (source_tag), mỗi bucket có
          // ngân sách MAX_PAGES*PAGE_SIZE dòng riêng, không bị bucket lớn
          // "ăn hết" quota của bucket nhỏ.
          const MAX_PAGES = 100; // tối đa 100,000 dòng / bucket (đủ cho mọi trường hợp thực tế)

          // ─── TIMEOUT 10 giây cho mỗi request ────────────────────────────
          const TIMEOUT_MS = 10_000;
          const withTimeoutMs = (ms: number, promise: Promise<any>): Promise<any> =>
            Promise.race([
              promise,
              new Promise<null>((_r, reject) =>
                setTimeout(() => reject(new Error('Supabase timeout')), ms)
              ),
            ]).catch((err: Error) => {
              console.warn('Supabase bị timeout hoặc lỗi:', err.message);
              return null;
            });

          type BucketKind = 'sales' | 'manpower' | 'target_actual';

          // Áp filter đúng bucket lên 1 query builder Supabase (dùng lại
          // cho cả head-count request và các trang range() bên dưới).
          const applyBucketFilter = (kind: BucketKind, q: any) => {
            if (kind === 'manpower') return q.eq('source_tag', 'Manpower');
            if (kind === 'target_actual') return q.eq('source_tag', 'TargetActual');
            // Sales: KHÔNG mang source_tag Manpower/TargetActual (comment
            // gốc ở bucketByTag() vẫn đúng — origin của Sales là dữ liệu
            // xuất xứ thật, không bao giờ trùng 2 tag trên).
            return q.not('source_tag', 'in', '("Manpower","TargetActual")');
          };

          // Tải toàn bộ (tối đa MAX_PAGES trang) của 1 bucket.
          // Trả về null nếu gặp lỗi/timeout → phía gọi sẽ giữ cache cũ.
          async function fetchBucketRows(kind: BucketKind): Promise<any[] | null> {
            const countResult = await withTimeoutMs(
              TIMEOUT_MS,
              Promise.resolve(
                applyBucketFilter(kind, db.from('sales_data').select('*', { count: 'exact', head: true }))
              )
            );
            if (!countResult) {
              console.warn(`Supabase count timeout [${kind}] — giữ cache local.`);
              return null;
            }
            const { count, error: countError } = countResult;
            if (countError) {
              console.error(`Supabase count error [${kind}]:`, countError);
              return null;
            }
            if (!count || count === 0) return [];

            const totalPages = Math.min(Math.ceil(count / PAGE_SIZE), MAX_PAGES);
            if (count > 10_000) {
              console.warn(`Bucket [${kind}] có ${count} dòng. Đang tải toàn bộ (tối đa ${MAX_PAGES * PAGE_SIZE} dòng) — cân nhắc dọn dẹp database nếu tốc độ tải chậm.`);
            }

            let bucketData: any[] = [];
            const CONCURRENCY = 3;
            const pageIndexes = Array.from({ length: totalPages }, (_, i) => i);
            for (let i = 0; i < pageIndexes.length; i += CONCURRENCY) {
              const batch = pageIndexes.slice(i, i + CONCURRENCY);
              const results = await Promise.all(
                batch.map(pi => {
                  const from = pi * PAGE_SIZE;
                  return withTimeoutMs(
                    TIMEOUT_MS,
                    Promise.resolve(
                      applyBucketFilter(kind, db.from('sales_data').select('*'))
                        // ── FIX ROOT CAUSE "347% Target/Actual lệch bất thường" ──
                        // Bản cũ order by 'id' — nhưng `id` là UUID (ngẫu nhiên,
                        // KHÔNG tuần tự theo thời gian ghi). order('id', desc)
                        // không hề lấy "5000 dòng mới nhất" như comment cũ mô tả,
                        // mà lấy ra 1 tập con NGẪU NHIÊN theo thứ tự chuỗi UUID —
                        // khiến tỉ lệ dòng Plan/Target vs Actual trong 5000 dòng
                        // được chọn hoàn toàn tình cờ, không phản ánh dữ liệu thật.
                        // Fix: sắp xếp theo `created_at` (thời điểm ghi thật) để
                        // "5000 dòng mới nhất" đúng nghĩa như tên gọi.
                        .order('created_at', { ascending: false })
                        .range(from, from + PAGE_SIZE - 1)
                    )
                  );
                })
              );
              for (const result of results) {
                if (!result) {
                  console.warn(`Bucket [${kind}] page fetch timeout — dùng cache local.`);
                  return null;
                }
                const { data, error } = result as any;
                if (error) {
                  console.warn(`Bucket [${kind}] fetch error, dừng và dùng cache:`, error.code, error.message);
                  return null;
                }
                if (data) {
                  bucketData = bucketData.concat(
                    data.map((d: any) => ({
                      ...(d.payload && typeof d.payload === 'object' ? d.payload : d),
                      source_tag: d.source_tag,
                      date: d.date ?? (d.payload?.date) ?? d.month ?? (d.payload?.month),
                    }))
                  );
                }
              }
            }
            return bucketData;
          }

          // Fetch 3 bucket độc lập — mỗi bucket có quota riêng, không tranh
          // nhau. Chạy tuần tự để tránh dồn quá nhiều request cùng lúc lên
          // Supabase Free tier (mỗi bucket bên trong đã tự chạy CONCURRENCY=3).
          const salesResult = await fetchBucketRows('sales');
          const manpowerResult = await fetchBucketRows('manpower');
          const targetActualResult = await fetchBucketRows('target_actual');

          if (salesResult === null && manpowerResult === null && targetActualResult === null) {
            // Cả 3 bucket đều lỗi/timeout → giữ cache đang hiển thị (nếu có).
            hasHydratedRef.current = true;
            setIsInitialDataLoading(false);
            return;
          }

          const sales = salesResult ?? [];
          const manpower = manpowerResult ?? [];
          const targetActual = targetActualResult ?? [];
          const allData = [...sales, ...manpower, ...targetActual];

          if (salesResult === null) console.warn('Bucket Sales tải thất bại — giữ dữ liệu Sales cache cũ (nếu có), 2 bucket khác vẫn cập nhật.');
          if (manpowerResult === null) console.warn('Bucket Manpower tải thất bại — giữ dữ liệu Manpower cache cũ (nếu có), 2 bucket khác vẫn cập nhật.');
          if (targetActualResult === null) console.warn('Bucket TargetActual tải thất bại — giữ dữ liệu TargetActual cache cũ (nếu có), 2 bucket khác vẫn cập nhật.');

          if (allData.length > 0) {
            // Chỉ ghi đè bucket nào tải THÀNH CÔNG (khác null) — bucket lỗi
            // giữ nguyên state hiện có, tránh xoá mất dữ liệu đang hiển thị.
            if (salesResult !== null) setSalesRows(sales);
            if (manpowerResult !== null) setManpowerRows(manpower);
            if (targetActualResult !== null) setTargetActualRows(targetActual);
            setAllRows(allData);
            setHeaders(Object.keys(allData[0]).map(k => k.trim()));
            setFilename('Supabase Cloud Database');
            setScreen('dashboard');
            idbSetCache('cached_dashboard_buckets', JSON.stringify({ sales, manpower, targetActual }))
              .catch(() => { /* IndexedDB không khả dụng — bỏ qua, dashboard vẫn dùng state React hiện tại */ });
            hasHydratedRef.current = true;
            setIsInitialDataLoading(false);
            console.log(
              'Đã đồng bộ dữ liệu mới nhất từ Supabase — Sales:', sales.length,
              '| Manpower:', manpower.length,
              '| TargetActual:', targetActual.length,
              '| Tổng:', allData.length, 'rows'
            );
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
    readLocalCacheAndShow().finally(() => { if (!cancelled) loadSupabaseData(); });
    return () => { cancelled = true; };
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

      // 2) Helper: lấy giá trị field case-insensitive.
      // ── FIX ROOT CAUSE "TargetActual/Manpower ra toàn số 0 sau khi sync" ──
      // Bản cũ: `r[k] ?? r[Capitalize(k)] ?? r[k.toLowerCase()]` — biến thể
      // thứ 3 TRÙNG LẶP VÔ NGHĨA với biến thể 1 (vì `k` truyền vào luôn đã
      // là lowercase, ví dụ 'division') — ý định ban đầu rõ ràng là check
      // thêm biến thể VIẾT HOA TOÀN BỘ ('DIVISION') nhưng bị viết nhầm.
      // Hệ quả: nếu Excel dùng header viết hoa toàn bộ (rất phổ biến với
      // file xuất từ hệ thống nội bộ — MODEL/DIVISION/TYPE1/TYPE2/DATE/VALUE),
      // gv() không tìm thấy cột nào → rơi vào fallback ('sales', '', 0...).
      // Với TargetActual, `division` fallback 'sales' không khớp bất kỳ điều
      // kiện SHIPMENT/OIS/AMT* nào trong TargetActualDashboard → mọi
      // qtyTarget/qtyActual/amtTarget/amtActual luôn = 0.
      // Fix: quét TOÀN BỘ key thật của dòng dữ liệu, so khớp không phân biệt
      // hoa/thường — bắt đúng bất kể Excel dùng casing nào.
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
          const target = k.toLowerCase();
          // Quét toàn bộ key thật của dòng dữ liệu, so khớp không phân biệt
          // hoa/thường/khoảng trắng đầu-cuối — bắt đúng dù Excel dùng
          // 'division', 'Division', hay 'DIVISION'.
          const matchKey = Object.keys(r).find(rk => rk.trim().toLowerCase() === target);
          const v = matchKey !== undefined ? r[matchKey] : undefined;
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
            // 'customer' thật trong file Excel gốc thường nằm ở cột Custom/
            // Type2 (xem detectIsTargetActual + nhánh đọc file thô trong
            // TargetActualDashboard.tsx: custCol tìm 'type2'/'custom'), không
            // phải cột tên đúng chữ 'Customer' — thêm fallback để không bị
            // rơi về 'Unknown' khi hiển thị donut chart theo khách hàng.
            customer: String(gv(r, 'customer', 'Customer', 'custom', 'Custom', 'type2', 'Type2') ?? '').trim(),
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

      // 4) Khử trùng NGAY TRONG cùng 1 lần upload trước khi gửi lên Supabase.
      // Lý do bắt buộc: bước 5 dùng upsert(...).onConflict trên bộ cột khoá
      // nghiệp vụ (SALES_DATA_KEY) — nếu 1 file Excel lỡ có 2 dòng trùng y
      // hệt khoá đó, Postgres sẽ báo lỗi "ON CONFLICT DO UPDATE command
      // cannot affect row a second time" vì cùng 1 câu lệnh không được đụng
      // 1 dòng đích 2 lần. Dòng xuất hiện SAU trong file được giữ lại, ghi
      // đè dòng trước — coi như "giá trị cập nhật gần nhất thắng".
      const dedupedMap = new Map<string, Record<string, any>>();
      for (const row of taggedRows) {
        const key = SALES_DATA_KEY.map(k => String(row[k] ?? '')).join('||');
        dedupedMap.set(key, row); // Set lại đè giá trị cũ nếu trùng khoá
      }
      const dedupedRows = Array.from(dedupedMap.values());
      if (dedupedRows.length < taggedRows.length) {
        console.warn(
          `${bucketTag}: phát hiện ${taggedRows.length - dedupedRows.length} dòng trùng khoá ` +
          `ngay trong file vừa tải lên — đã gộp, chỉ giữ giá trị cuối cùng cho mỗi khoá.`
        );
      }

      // 5) Upsert theo từng lô thay vì insert thô — mỗi lô là "insert nếu
      // chưa có khoá, update nếu đã có" (idempotent). Kết hợp với UNIQUE
      // CONSTRAINT phía Supabase (supabase_dedupe_and_constraint.sql), đây
      // là lớp bảo vệ triệt để: DÙ hàm này bị gọi trùng nhiều lần (double
      // click, race condition, script cũ nào lỡ chạy lại, môi trường dev
      // vô tình trỏ nhầm vào DB thật...), dữ liệu trong bảng KHÔNG BAO GIỜ
      // bị nhân bản nữa — chỉ được ghi đè bằng giá trị mới nhất.
      const CHUNK_SIZE = 500;
      for (let i = 0; i < dedupedRows.length; i += CHUNK_SIZE) {
        const chunk = dedupedRows.slice(i, i + CHUNK_SIZE);
        const { error: upsertError } = await supabase
          .from('sales_data')
          .upsert(chunk, { onConflict: SALES_DATA_KEY.join(',') });
        if (upsertError) {
          console.error(`Supabase upsert error (${bucketTag}) tại lô dòng ${i}:`, upsertError);
          alert(
            `⚠️ Đồng bộ Cloud thất bại (${bucketTag}, dòng ${i}–${i + chunk.length}).\n` +
            `Dữ liệu vẫn hiển thị tạm trên máy này nhưng sẽ MẤT khi F5.\n` +
            `Chi tiết lỗi: ${upsertError.message}\n\n` +
            `Nếu lỗi nhắc tới "unique constraint" hoặc "ON CONFLICT", hãy chạy file ` +
            `supabase_dedupe_and_constraint.sql trong Supabase SQL Editor trước.`
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
    idbDelCache('cached_dashboard_buckets').catch(() => { /* ignore */ });
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
