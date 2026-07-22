import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { NeonButton } from './NeonButton';

// ─────────────────────────────────────────────────────────────────────────
// MENU 5 — Dashboard Hiệu suất sản xuất theo Model (dữ liệu thật từ Test5.xlsx)
// Nhân bản từ DashboardTemplate.tsx, thay demo/mock data bằng dữ liệu thật.
// Bucket dữ liệu: độc lập hoàn toàn, KHÔNG đụng tới sales/manpower/target_actual.
//
// EPCC (menu5-direct-upload-no-prebuilt-json) - FIX ROOT CAUSE "Vite báo lỗi
// 'Failed to resolve import ../data/test5Data.json ... Does the file exist?'
// và app CRASH TOÀN BỘ ngay khi mở Mục 5": bản cũ dùng `import test5Raw from
// '../data/test5Data.json'` — đây là STATIC IMPORT, được Vite resolve ở
// BUILD-TIME, không phải runtime. Nếu file .json này không có sẵn trong dự
// án (nó vốn chỉ là artifact tổng hợp SẴN từ 1 lần chạy Python, không phải
// file luôn đi kèm mã nguồn) thì toàn bộ app sập ngay từ lúc transform, kể
// cả không hề vào Mục 5 — không có cách nào try/catch được lỗi import tĩnh.
// Fix: bỏ hẳn phụ thuộc vào file .json dựng sẵn. Mục 5 giờ hoạt động ĐÚNG
// như Mục 1/2/3 — bắt đầu ở trạng thái rỗng, người dùng tự tải trực tiếp
// Test5.xlsx lên qua nút "Tải lên Excel mới" (đã hỗ trợ gộp mọi sheet, xem
// aggregateWorkbookToTest5Data() bên dưới); kết quả được cache qua
// IndexedDB nên F5 không mất dữ liệu đã tải.
// ─────────────────────────────────────────────────────────────────────────

type Lang = 'vi' | 'en' | 'ko';
type ThemeMode = 'dark' | 'light';
type ViewBy = 'month' | 'quarter' | 'year';

interface SeriesPoint { y: number; m: number; v: number; }
interface ModelEntry {
  model: string;
  rowCount: number;
  series: Record<string, SeriesPoint[]>;
  radarType: { type: string; value: number }[];
  radarCustom: { custom: string; value: number }[];
  // EPCC (menu5-4-new-charts-store-ttl-rows) - FIX ROOT CAUSE "không thể tái
  // tạo đúng số TTL trong Excel (VD tổng theo năm của F-Cost%/F-cost($)) chỉ
  // bằng cách cộng/trung bình lại 12 tháng đã lưu": trước đây accumulateSheetRows
  // CHỦ Ý bỏ qua mọi dòng MONTH="TTL" (isTtl=true không lưu vào `monthly`).
  // Với item dạng số lượng (SCRAP, SHIPMENT...) TTL đúng bằng tổng 12 tháng nên
  // không sao, nhưng với F-Cost% một số model có vài dòng tháng dữ liệu lỗi/bẩn
  // (VD giá trị "8", "1" thay vì %) khiến suy ngược lại từ trung bình tháng lệch
  // hẳn so với TTL gốc trong Excel. Fix: lưu thêm nguyên văn giá trị Ở DÒNG TTL
  // theo (item, year), dùng trực tiếp cho 4 biểu đồ tổng hợp cross-model mới,
  // không suy luận lại từ dữ liệu tháng nữa.
  ttlByItemYear: Record<string, Record<number, number>>;
}
interface Test5Data {
  bestModel: string;
  models: ModelEntry[];
  // EPCC (menu5-revenue-tab-sales-by-customer) - customer -> year -> tổng
  // SALES AMT (TTL), dùng cho tab mới "TÌNH HÌNH DOANH THU".
  salesByCustomer: Record<string, Record<number, number>>;
  // EPCC (menu5-revenue-chart3-viewby-monthly-customer) - customer -> year ->
  // month -> tổng SALES AMT THÁNG (không phải TTL), dùng cho Chart 3 khi
  // XEM THEO = Quý/Tháng — xem giải thích đầy đủ ở AggAccumulator phía trên.
  salesByCustomerMonthly: Record<string, Record<number, Record<number, number>>>;
}

// Trạng thái rỗng ban đầu — thay cho test5Data.json dựng sẵn trước đây.
const EMPTY_DATA: Test5Data = { bestModel: '', models: [], salesByCustomer: {}, salesByCustomerMonthly: {} };


// ── Cache dữ liệu Excel đã tải lên bằng IndexedDB (KHÔNG dùng localStorage) ─
// Namespace riêng cho Mục 5, không đụng key của sales/manpower/target_actual.
const IDB_DB_NAME = 'imvina_dashboard_cache';
const IDB_STORE = 'kv';
// EPCC (menu5-fcost-precision-cache-bust) - FIX ROOT CAUSE "đã sửa lỗi làm
// tròn F-Cost% (menu5-fcost-precision-fix) nhưng Target/Actual trên chart
// VẪN sai y hệt số cũ (1.05%→1.10%, 0.71%→0.70%...)": không phải do fix
// trước sai — parse logic đã đúng — mà do dữ liệu ĐÃ TỪNG tải lên TRƯỚC khi
// fix vẫn còn nằm nguyên trong cache IndexedDB (key cũ 'test5_agg_v1'), và
// useEffect lúc mở trang ĐỌC THẲNG cache này lên chart, KHÔNG re-parse lại
// Excel — nên dù code parse mới đã đúng, người dùng vẫn thấy số liệu cache
// CŨ (parse bằng logic làm tròn sai trước đó) cho tới khi tự tay bấm "Tải
// Excel" để ghi đè cache. Fix tận gốc, không phụ thuộc trí nhớ người dùng:
// bump version key 'v1' → 'v2' — cache cũ lập tức "vô hình" với code mới
// (đọc bằng key v2, không tồn tại), buộc app hiện màn hình rỗng yêu cầu tải
// lại Excel đúng 1 lần duy nhất, từ đó luôn dùng logic parse mới. Áp dụng
// quy ước này cho MỌI lần sửa logic parse sau này: đổi logic tính toán ⇒
// PHẢI bump version cache key kèm theo, nếu không cache cũ sẽ luôn thắng.
// EPCC (menu5-ttl-overwrite-cache-bust) - FIX ROOT CAUSE "Chart 'XUẤT HÀNG &
// DOANH SỐ THEO NĂM' (Chart 1) vẫn hiển thị số liệu SAI (thấp hơn Excel gốc
// 20-65% tuỳ năm ở giai đoạn 2017-2022), dù logic parse (accumulateSheetRows,
// xem comment menu5-ttl-overwrite-multi-type ở dưới) ĐÃ được sửa đúng để CỘNG
// DỒN (+=) thay vì GHI ĐÈ (.set()) các dòng TTL trùng (model, item, năm) do
// khác TYPE/CUSTOM": bản thân logic parse mới hoàn toàn đúng — đối chiếu trực
// tiếp Test5.xlsx bằng script độc lập cho kết quả khớp 100%. Vấn đề y hệt bug
// đã từng gặp và tự ghi lại làm quy ước ở EPCC (menu5-fcost-precision-cache-
// bust) ngay phía trên: fix menu5-ttl-overwrite-multi-type ĐỔI LOGIC TÍNH TOÁN
// (thay đổi cách cộng dồn TTL) nhưng KHÔNG bump version cache key kèm theo —
// vi phạm đúng quy ước "MỌI lần sửa logic parse sau này: đổi logic tính toán
// ⇒ PHẢI bump version cache key" đã tự đề ra. Hậu quả: người dùng đã tải
// Test5.xlsx lên TRƯỚC khi fix ttl-overwrite được áp dụng vẫn có dữ liệu tổng
// hợp CŨ (tính bằng logic ghi đè sai) nằm nguyên trong IndexedDB dưới key
// 'test5_agg_v2' — useEffect lúc mở trang đọc thẳng cache này lên chart,
// KHÔNG re-parse lại Excel, nên vẫn thấy số liệu SAI y hệt trước fix, dù code
// hiện tại đã đúng. Fix tận gốc, không phụ thuộc trí nhớ người dùng: bump
// version key 'v2' → 'v3' — cache cũ lập tức "vô hình" với code mới, buộc app
// hiện màn hình rỗng yêu cầu tải lại Excel đúng 1 lần duy nhất, từ đó luôn
// dùng logic cộng dồn TTL mới (đúng). Nhắc lại quy ước cho các lần sửa sau:
// đổi CÁCH TÍNH trong accumulateSheetRows/buildTest5DataFromAccumulator ⇒
// LUÔN bump version cache key kèm theo trong CÙNG 1 lần sửa, không tách rời.
// EPCC (menu5-monthly-fallback-cache-bust) - bump lần nữa 'v3' → 'v4' vì 2
// fix trong lượt này (menu5-monthly-overwrite-multi-type +
// menu5-ttl-missing-row-fallback) tiếp tục đổi logic tính toán — đúng quy
// ước đã tự đặt ra ở comment v1→v2 phía trên: đổi logic ⇒ luôn bump cache.
// EPCC (menu5-revenue-chart3-viewby-monthly-customer-cache-bust) - bump lần
// nữa 'v4' → 'v5' vì lần sửa này THÊM MỚI field `salesByCustomerMonthly`
// (đổi logic tích luỹ trong accumulateSheetRows/buildTest5DataFromAccumulator)
// — đúng quy ước đã tự đặt ra ở các lần bump trước: đổi logic tính toán ⇒
// LUÔN bump version cache key kèm theo. Không bump sẽ khiến cache cũ (thiếu
// hẳn field mới) vẫn được đọc thẳng lên Chart 3, khiến nhánh Quý/Tháng mới
// thêm không có dữ liệu cho tới khi người dùng tự tay tải lại Excel.
const IDB_KEY_MENU5_DATA = 'menu5:test5_agg_v5';
const IDB_KEY_MENU5_META = 'menu5:test5_agg_v5_meta';

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
// EPCC (menu5-remove-clear-reupload-buttons): đã bỏ hẳn nút "Xoá dữ liệu"/
// "Tải lên Excel mới" khỏi filter bar theo yêu cầu — idbDelCache() không
// còn nơi nào gọi tới (trước đây chỉ dùng trong handleClearData đã xoá),
// nên dọn bỏ luôn theo đúng quy ước "xoá code chết mỗi lần fix". Cache vẫn
// ghi qua idbSetCache() ở handleUploadFile và đọc qua idbGetCache() khi mở
// trang — chỉ mất khả năng XOÁ cache qua UI, không ảnh hưởng luồng tải lên/
// đọc cache còn lại.

// ── Parse + tổng hợp file Excel upload (JS, chạy trong trình duyệt) ────────
// Tương đương logic Python đã dùng để tạo test5Data.json ban đầu, nhưng nhận
// diện cột THEO TÊN header (không theo vị trí cứng) để chịu được file có thứ
// tự cột hơi khác. Yêu cầu tối thiểu các cột: MODEL, TYPE, ITEM, CUSTOM,
// YEAR, MONTH, và 1 cột số lượng (QTY / Q'TY / Q`TY).
const MONTH_MAP: Record<string, number> = {
  JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6,
  JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12,
};

function normalizeHeader(h: unknown): string {
  return String(h ?? '').toUpperCase().replace(/[^A-Z]/g, '');
}

function toNum(v: unknown): number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const s = String(v).trim();
  if (s === '' || s === '-' || s.toUpperCase() === 'N/A' || s.toUpperCase() === 'NA') return null;
  const n = parseFloat(s.replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

class Test5ParseError extends Error {}

// ── Kiểu tích luỹ dùng chung khi gộp NHIỀU sheet vào 1 bộ dữ liệu ───────────
interface AggAccumulator {
  monthly: Map<string, Map<string, Map<string, Map<number, number>>>>;
  radarType: Map<string, Map<string, number>>;
  radarCustom: Map<string, Map<string, number>>;
  rowCount: Map<string, number>;
  // model -> item -> year -> giá trị nguyên văn ở dòng MONTH="TTL"
  ttl: Map<string, Map<string, Map<number, number>>>;
  // EPCC (menu5-revenue-tab-sales-by-customer) - THÊM MỚI cho tab "TÌNH HÌNH
  // DOANH THU": customer -> year -> tổng SALES AMT (lấy nguyên văn dòng
  // MONTH="TTL", giống nguyên tắc `ttl` ở trên — KHÔNG cộng dồn lại từ 12
  // tháng để tránh sai số như đã ghi chú ở menu5-4-new-charts-store-ttl-rows),
  // cộng dồn qua MỌI model (khác `ttl`/`radarCustom` vốn tách riêng theo
  // từng model) — vì tab doanh thu cần góc nhìn "toàn công ty theo khách
  // hàng", không phải theo model.
  salesByCustomer: Map<string, Map<number, number>>;
  // EPCC (menu5-revenue-chart3-viewby-monthly-customer) - THÊM MỚI: customer
  // -> year (string key) -> month -> tổng SALES AMT THÁNG (cộng dồn từ MỌI
  // dòng THÁNG thật, KHÔNG lấy dòng TTL — khác `salesByCustomer` ở trên vốn
  // CHỦ Ý chỉ có TTL/năm, không đủ chi tiết để vẽ trục Quý/Tháng). Dùng
  // riêng cho Chart 3 (Doanh số theo Khách hàng) khi người dùng đổi "XEM
  // THEO" sang Quý/Tháng — y hệt nguyên tắc `monthly` (theo model) ở trên,
  // chỉ khác đơn vị nhóm là khách hàng (`custom`) thay vì model.
  salesByCustomerMonthly: Map<string, Map<string, Map<number, number>>>;
}

function createAccumulator(): AggAccumulator {
  return {
    monthly: new Map(),
    radarType: new Map(),
    radarCustom: new Map(),
    rowCount: new Map(),
    ttl: new Map(),
    salesByCustomer: new Map(),
    salesByCustomerMonthly: new Map(),
  };
}

// Dò dòng tiêu đề (chứa đủ các cột bắt buộc) trong 1 sheet đã đọc dạng
// mảng-2-chiều. Trả về null nếu sheet này không có bảng dữ liệu hợp lệ
// (ví dụ sheet tổng hợp/pivot/ghi chú phụ) — KHÔNG coi là lỗi ở cấp sheet,
// chỉ bỏ qua sheet đó khi gộp toàn workbook.
function findHeaderRow(rows: unknown[][]): { headerRowIdx: number; colIdx: Record<string, number> } | null {
  for (let r = 0; r < Math.min(rows.length, 20); r++) {
    const row = rows[r] ?? [];
    const map: Record<string, number> = {};
    row.forEach((cell, i) => {
      const key = normalizeHeader(cell);
      if (key && !(key in map)) map[key] = i;
    });
    const hasCore = ['MODEL', 'TYPE', 'ITEM', 'CUSTOM', 'YEAR', 'MONTH'].every(k => k in map);
    const qtyKey = Object.keys(map).find(k => k.includes('QTY'));
    if (hasCore && qtyKey) {
      return { headerRowIdx: r, colIdx: { ...map, QTY: map[qtyKey] } };
    }
  }
  return null;
}

// EPCC (menu5-multisheet-upload) - FIX ROOT CAUSE "dữ liệu/biểu đồ Mục 5
// không đầy đủ sau khi tải Excel mới": handleUploadFile bản cũ chỉ đọc
// wb.SheetNames[0] (đúng 1 sheet đầu tiên) rồi gọi thẳng
// aggregateRowsToTest5Data(rows) trên sheet đó. Nếu file người dùng tải lên
// có nhiều sheet (ví dụ tách theo năm/khách hàng, hoặc có thêm sheet phụ),
// toàn bộ model/tháng nằm ở sheet thứ 2 trở đi bị ÂM THẦM BỎ QUA — không có
// thông báo lỗi nào, biểu đồ/KPI/radar chỉ đơn giản thiếu dữ liệu so với kỳ
// vọng (khác hẳn dữ liệu mẫu test5Data.json vốn được tổng hợp sẵn từ TOÀN
// BỘ Test5.xlsx). Đây đúng dạng bug đã từng gặp và fix ở App.tsx
// (parseWorkbook gộp mọi sheet) nhưng fix đó chưa áp dụng lại cho component
// độc lập này.
// Fix: gộp accumulateSheetRows() vào 1 bộ Map dùng chung cho MỌI sheet có
// đủ cột bắt buộc trong workbook, thay vì chỉ đọc sheet đầu tiên.
function accumulateSheetRows(
  rows: unknown[][],
  colIdx: Record<string, number>,
  headerRowIdx: number,
  acc: AggAccumulator,
): void {
  const { monthly, radarType, radarCustom, rowCount, ttl, salesByCustomer, salesByCustomerMonthly } = acc;
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r] ?? [];
    const modelRaw = row[colIdx.MODEL];
    const model = modelRaw != null ? String(modelRaw).trim() : '';
    if (!model) continue;
    const typ = row[colIdx.TYPE] != null ? String(row[colIdx.TYPE]).trim() : '';
    const item = row[colIdx.ITEM] != null ? String(row[colIdx.ITEM]).trim() : '';
    const custom = row[colIdx.CUSTOM] != null ? String(row[colIdx.CUSTOM]).trim() : '';
    const yearRaw = row[colIdx.YEAR];
    const year = yearRaw != null ? Number(yearRaw) : NaN;
    const monthRaw = row[colIdx.MONTH];
    const qty = toNum(row[colIdx.QTY]);
    if (qty === null) continue;

    rowCount.set(model, (rowCount.get(model) ?? 0) + 1);

    let monthNum: number | null = null;
    let isTtl = false;
    if (typeof monthRaw === 'number') {
      monthNum = monthRaw;
    } else {
      const ms = String(monthRaw ?? '').toUpperCase().trim();
      if (ms === 'TTL') isTtl = true;
      else if (ms in MONTH_MAP) monthNum = MONTH_MAP[ms];
    }

    // EPCC (menu5-fcost-zero-as-missing) - FIX ROOT CAUSE "biểu đồ ALL YEAR
    // (và mọi biểu đồ/KPI/bảng khác dùng F-Cost %) hiển thị TARGET/ACTUAL =
    // 0.0% liền mạch cho các tháng JUN→DEC dù người dùng khẳng định 'file
    // database có số liệu thực tế'": đã đối chiếu trực tiếp Test5.xlsx —
    // KHÔNG PHẢI lỗi code đọc/vẽ sai. App đọc ĐÚNG dữ liệu nó nhận được:
    // dòng MODEL=Target/Actual (và nhiều model khác, tổng 101 dòng toàn
    // file) có ô Q`TY = số 0 NGUYÊN VĂN cho các tháng đó (VD năm 2024,
    // model Target/Actual, JUN..DEC đều =0), trong khi CHÍNH file này ở
    // các tháng thật sự "chưa diễn ra" (VD JUL..DEC/2026) lại dùng ký hiệu
    // "-" (được toNum() coi là NULL/không có dữ liệu ở dòng 202 phía trên).
    // Hai cách biểu diễn "chưa có số liệu" khác nhau (0 vs "-") ngay trong
    // CÙNG 1 sheet khiến các tháng đã qua nhưng bị bỏ trống lại bị hiểu lầm
    // thành "đã đạt 0% F-Cost" — một kết quả GẦN NHƯ KHÔNG THỂ xảy ra thật
    // với chỉ số F-Cost % (khác các ITEM dạng số lượng như SHIPMENT/SCRAP,
    // nơi 0 là giá trị hợp lệ bình thường). Fix tại ĐÚNG 1 điểm gốc dùng
    // chung cho MỌI biểu đồ/KPI/bảng (thay vì vá riêng từng chart): khi
    // item = F-Cost % (ITEM_FCOST) và giá trị tháng = 0 (không phải dòng
    // TTL), coi như CHƯA CÓ SỐ LIỆU — bỏ qua, không ghi vào `monthly` — để
    // hành vi giống hệt cách "-" đã được xử lý (đường biểu đồ ngắt đoạn,
    // KPI/trung bình không bị kéo lệch, bảng hiện "—" thay vì "0.00%").
    // Ảnh hưởng dây chuyền tự động tới TẤT CẢ nơi dùng ITEM_FCOST: Chart 1
    // (ALL YEAR: bars F-cost($)/(억원) + line Target/Actual), Chart 2 (F-COST
    // THEO MODEL, cả 3 chế độ Tháng/Quý/Năm), 4 thẻ KPI (TB F-Cost%, Tháng
    // đạt Target, so với năm trước), và cột "F-Cost %" trong bảng chi tiết.
    // Chart 3/4 (Scrap) không đụng tới vì dùng ITEM khác, không liên quan.
    const isFcostZeroPlaceholder = item === ITEM_FCOST && qty === 0;
    if (!isTtl && monthNum && year && !isFcostZeroPlaceholder) {
      if (!monthly.has(model)) monthly.set(model, new Map());
      const byItem = monthly.get(model)!;
      if (!byItem.has(item)) byItem.set(item, new Map());
      const byYear = byItem.get(item)!;
      const yKey = String(year);
      if (!byYear.has(yKey)) byYear.set(yKey, new Map());
      // EPCC (menu5-fcost-precision-fix) - FIX ROOT CAUSE "Target/Actual (và
      // mọi F-Cost % khác) bị SAI SỐ so với Excel gốc (VD Target JAN 1.05%
      // trong Excel → hiển thị 1.10% trên chart)": trước đây MỌI item (kể cả
      // F-Cost %, vốn là giá trị PHÂN SỐ như 0.0105 = 1.05%) đều bị làm tròn
      // chung về 3 chữ số thập phân của phân số (`Math.round(qty*1000)/1000`).
      // Với ITEM dạng số lượng (SHIPMENT, SCRAP...) làm tròn 0.001 đơn vị là
      // vô hại, nhưng với F-Cost % thì 0.001 phân số = 0.1 ĐIỂM PHẦN TRĂM —
      // quá thô, đủ để đổi hẳn số hiển thị (1.05%→1.10%, 0.71%→0.70%...).
      // Fix: F-Cost % dùng độ chính xác cao hơn (6 chữ số thập phân, cùng
      // quy ước với dòng TTL ở dưới) để giữ NGUYÊN đúng giá trị gốc trong
      // Excel, không tự ý làm tròn; các ITEM số lượng khác giữ nguyên hành vi
      // cũ (không đổi, tránh ảnh hưởng ngoài ý muốn).
      // EPCC (menu5-monthly-overwrite-multi-type) - FIX ROOT CAUSE "sau khi
      // sửa menu5-ttl-overwrite-multi-type (dòng TTL), Chart 1 (Xuất hàng &
      // Doanh số theo Năm, tab TÌNH HÌNH DOANH THU) vẫn lệch với ảnh tham
      // chiếu ở đúng những năm/model có NHIỀU dòng TYPE (VD model SO1C30 S25,
      // năm 2024, TYPE=OIS ASSY)": đối chiếu Test5.xlsx phát hiện ĐÚNG BUG Y
      // HỆT `menu5-ttl-overwrite-multi-type` nhưng ở TẦNG DỮ LIỆU THÁNG thay
      // vì TTL — nhiều dòng TYPE/CUSTOM khác nhau có thể cùng (item, năm,
      // tháng), nhưng code cũ dùng `.set()` GHI ĐÈ mỗi khi gặp thêm dòng cùng
      // tháng, chỉ dòng TYPE đọc sau cùng còn tồn tại. Với 4 ITEM SỐ LƯỢNG
      // (SHIPMENT/PROD'N/SALES AMT/SCRAP) các dòng TYPE khác nhau PHẢI cộng
      // lại — F-Cost % vẫn giữ nguyên ghi đè vì là tỷ lệ, không được cộng dồn.
      const isQuantityItemMonthly = item === ITEM_SHIPMENT || item === ITEM_PRODUCTION
        || item === ITEM_SALES || item === ITEM_SCRAP;
      const prevMonthQty = byYear.get(yKey)!.get(monthNum) ?? 0;
      const rawQty = isQuantityItemMonthly ? prevMonthQty + qty : qty;
      const roundedQty = item === ITEM_FCOST
        ? Math.round(rawQty * 1000000) / 1000000
        : Math.round(rawQty * 1000) / 1000;
      byYear.get(yKey)!.set(monthNum, roundedQty);
    }

    // EPCC (menu5-4-new-charts-store-ttl-rows): lưu song song giá trị dòng
    // TTL nguyên văn — xem giải thích root-cause ở khai báo ttlByItemYear.
    // EPCC (menu5-ttl-overwrite-multi-type) - FIX ROOT CAUSE "Chart 1 XUẤT
    // HÀNG & DOANH SỐ THEO NĂM (+ bảng chi tiết/Top model khi gặp năm có
    // model nhiều dòng TYPE) thấp hơn thực tế 20-65% tuỳ năm": bản cũ dùng
    // `.set()` GHI ĐÈ mỗi khi gặp dòng TTL — nhưng 1 model thường có NHIỀU
    // dòng TTL cùng (item, năm) do khác TYPE/CUSTOM (VD SUB1, SUB2, OIS ASSY,
    // OIS LENS, OIS IR, VCMC — đối chiếu trực tiếp Test5.xlsx: model SO3850
    // năm 2017 có TYPE=SUB1 TTL=1.190,7 VÀ TYPE=SUB2 TTL=1.195,3 cho cùng
    // SHIPMENT, đúng ra phải cộng lại = 2.386 nhưng dòng sau ghi đè dòng
    // trước nên chỉ còn 1.195,3). Toàn file có 136/528 tổ hợp (model, item,
    // năm) rơi vào tình huống này, khiến SHIPMENT thiếu 25,9%, PROD'N thiếu
    // 27,6%, SALES AMT thiếu 18,8%, SCRAP thiếu 8,7% so với tổng đúng — trong
    // khi Chart 3/4 (dùng `salesByCustomer`, vốn đã cộng dồn bằng `+=` ngay
    // từ đầu) không hề bị ảnh hưởng, nên 2 nhóm chart lệch nhau dù cùng biểu
    // diễn 1 con số "tổng doanh số công ty".
    // Fix: CỘNG DỒN (thay vì ghi đè) cho 4 ITEM dạng SỐ LƯỢNG
    // (SHIPMENT/PROD'N/SALES AMT/SCRAP) — đúng bản chất có thể cộng nhiều
    // dòng TYPE/CUSTOM lại. F-Cost % là ITEM dạng TỶ LỆ (không được cộng dồn)
    // nên GIỮ NGUYÊN hành vi ghi đè/lấy dòng cuối như cũ, tránh phá logic
    // ALL YEAR/F-COST THEO MODEL đang dùng đúng.
    if (isTtl && year) {
      if (!ttl.has(model)) ttl.set(model, new Map());
      const byItemTtl = ttl.get(model)!;
      if (!byItemTtl.has(item)) byItemTtl.set(item, new Map());
      const byYearTtl = byItemTtl.get(item)!;
      const isQuantityItem = item === ITEM_SHIPMENT || item === ITEM_PRODUCTION
        || item === ITEM_SALES || item === ITEM_SCRAP;
      const prevTtl = byYearTtl.get(year) ?? 0;
      const newTtl = isQuantityItem ? prevTtl + qty : qty;
      byYearTtl.set(year, Math.round(newTtl * 1000000) / 1000000);
    }

    if (item === ITEM_SHIPMENT && isTtl) {
      if (typ) {
        if (!radarType.has(model)) radarType.set(model, new Map());
        const mm = radarType.get(model)!;
        mm.set(typ, (mm.get(typ) ?? 0) + qty);
      }
      if (custom) {
        if (!radarCustom.has(model)) radarCustom.set(model, new Map());
        const mm = radarCustom.get(model)!;
        mm.set(custom, (mm.get(custom) ?? 0) + qty);
      }
    }

    // EPCC (menu5-revenue-tab-sales-by-customer) - tích luỹ SALES AMT theo
    // (khách hàng, năm) từ dòng TTL — dùng riêng cho tab "TÌNH HÌNH DOANH
    // THU" (chart "Doanh số theo Khách hàng" + donut "Tỷ trọng theo Khách
    // hàng"), không đụng gì tới `radarCustom` (SHIPMENT, không tách năm).
    if (item === ITEM_SALES && isTtl && custom && year) {
      if (!salesByCustomer.has(custom)) salesByCustomer.set(custom, new Map());
      const byYear = salesByCustomer.get(custom)!;
      byYear.set(year, (byYear.get(year) ?? 0) + qty);
    }

    // EPCC (menu5-revenue-chart3-viewby-monthly-customer) - tích luỹ song
    // song SALES AMT theo (khách hàng, năm, THÁNG) từ các dòng THÁNG thật
    // (isTtl=false) — dùng riêng cho Chart 3 khi XEM THEO = Quý/Tháng.
    // Cộng dồn (+=) giống các ITEM số lượng khác (nhiều dòng TYPE khác nhau
    // có thể cùng (khách hàng, năm, tháng), phải cộng lại — xem cùng lý do
    // ở menu5-monthly-overwrite-multi-type phía trên), không lấy dòng TTL.
    if (item === ITEM_SALES && !isTtl && monthNum && year && custom) {
      if (!salesByCustomerMonthly.has(custom)) salesByCustomerMonthly.set(custom, new Map());
      const byYearC = salesByCustomerMonthly.get(custom)!;
      const yKeyC = String(year);
      if (!byYearC.has(yKeyC)) byYearC.set(yKeyC, new Map());
      const byMonthC = byYearC.get(yKeyC)!;
      const prevC = byMonthC.get(monthNum) ?? 0;
      byMonthC.set(monthNum, Math.round((prevC + qty) * 1000) / 1000);
    }
  }
}

// Dựng Test5Data cuối cùng từ bộ Map đã tích luỹ (có thể đến từ 1 hoặc
// NHIỀU sheet gộp lại) — logic giữ nguyên 100% so với bản gốc, chỉ đổi
// nguồn đầu vào từ "3 Map cục bộ trong 1 lần gọi" sang "acc dùng chung".
function buildTest5DataFromAccumulator(acc: AggAccumulator): Test5Data {
  const { monthly, radarType, radarCustom, rowCount, ttl, salesByCustomer, salesByCustomerMonthly } = acc;

  if (monthly.size === 0) {
    throw new Test5ParseError('Không đọc được dòng dữ liệu hợp lệ nào sau dòng tiêu đề.');
  }

  const models: ModelEntry[] = Array.from(monthly.keys()).sort().map(model => {
    const series: Record<string, SeriesPoint[]> = {};
    for (const [item, byYear] of monthly.get(model)!.entries()) {
      const pts: SeriesPoint[] = [];
      for (const [yKey, byMonth] of byYear.entries()) {
        for (const [m, v] of byMonth.entries()) pts.push({ y: Number(yKey), m, v });
      }
      pts.sort((a, b) => ym(a.y, a.m) - ym(b.y, b.m));
      series[item] = pts;
    }
    const ttlByItemYear: Record<string, Record<number, number>> = {};
    for (const [item, byYear] of ttl.get(model)?.entries() ?? []) {
      ttlByItemYear[item] = Object.fromEntries(byYear.entries());
    }
    // EPCC (menu5-ttl-missing-row-fallback) - FIX ROOT CAUSE "sau khi sửa cả
    // menu5-ttl-overwrite-multi-type (dòng TTL) và menu5-monthly-overwrite-
    // multi-type (dòng tháng), Chart 1 (Xuất hàng & Doanh số theo Năm) VẪN
    // lệch với ảnh tham chiếu, riêng năm 2024": đối chiếu trực tiếp Test5.xlsx
    // bằng script độc lập tìm ra nguyên nhân THỨ 3, khác hẳn 2 bug trước —
    // model 'SO1C30 S25' (TYPE='OIS ASSY', năm 2024) có ĐỦ 12 dòng THÁNG cho
    // SHIPMENT/PROD'N/SALES AMT (cộng lại đúng ~3.688/~3.879/~9.098 tương
    // ứng) nhưng KHÔNG HỀ CÓ dòng TTL nào cho tổ hợp này — không phải ghi đè
    // (2 bug trước) mà là THIẾU HẲN dòng TTL ngay trong Excel gốc. Vì `ttl`
    // chỉ cộng những dòng TTL THỰC SỰ TỒN TẠI, tổ hợp này không được cộng vào
    // đâu cả → SHIPMENT/PROD'N/SALES AMT năm 2024 (tổng toàn công ty) thấp
    // hơn thực tế 8-11% tuỳ chỉ số, đúng bằng phần đóng góp bị thiếu của
    // SO1C30 S25/OIS ASSY (và vài tổ hợp nhỏ khác cùng dạng).
    // Fix: với 4 ITEM SỐ LƯỢNG (SHIPMENT/PROD'N/SALES AMT/SCRAP), sau khi có
    // `series[item]` (dữ liệu THÁNG đã cộng dồn đúng ở trên), tính lại tổng
    // theo năm từ 12 tháng rồi lấy MAX so với giá trị TTL nguyên văn hiện có
    // — TTL vẫn được ưu tiên khi 2 số khớp nhau (trường hợp bình thường,
    // không đổi hành vi cũ), chỉ khi TTL THIẾU/THẤP HƠN tổng tháng thật (dấu
    // hiệu thiếu dòng TTL trong Excel) mới dùng tổng tháng để không bỏ sót dữ
    // liệu có thật. Không áp dụng cho F-Cost % (tỷ lệ, không được cộng dồn
    // theo cách này — giữ nguyên logic TTL-only đã có).
    for (const item of [ITEM_SHIPMENT, ITEM_PRODUCTION, ITEM_SALES, ITEM_SCRAP]) {
      const pts = series[item];
      if (!pts || !pts.length) continue;
      const monthlySumByYear = new Map<number, number>();
      for (const p of pts) monthlySumByYear.set(p.y, (monthlySumByYear.get(p.y) ?? 0) + p.v);
      const byYearOut = ttlByItemYear[item] ?? (ttlByItemYear[item] = {});
      for (const [y, monthlySum] of monthlySumByYear.entries()) {
        const roundedMonthlySum = Math.round(monthlySum * 1000000) / 1000000;
        const existingTtl = byYearOut[y];
        if (existingTtl == null || roundedMonthlySum > existingTtl) {
          byYearOut[y] = roundedMonthlySum;
        }
      }
    }
    return {
      model,
      rowCount: rowCount.get(model) ?? 0,
      series,
      radarType: Array.from(radarType.get(model)?.entries() ?? []).sort((a, b) => a[0].localeCompare(b[0]))
        .map(([type, value]) => ({ type, value: Math.round(value * 10) / 10 })),
      radarCustom: Array.from(radarCustom.get(model)?.entries() ?? []).sort((a, b) => a[0].localeCompare(b[0]))
        .map(([custom, value]) => ({ custom, value: Math.round(value * 10) / 10 })),
      ttlByItemYear,
    };
  });

  // EPCC (menu5-bestmodel-prefer-recent-year) - FIX ROOT CAUSE "Chart 1 & 2
  // (vừa liên kết với bộ lọc MODEL) trống trơn ngay khi mở trang, dù dữ
  // liệu Excel có thật": bestModel trước đây chỉ so F-Cost% TRUNG BÌNH TOÀN
  // BỘ LỊCH SỬ (>=6 điểm), không quan tâm dữ liệu đó cũ hay mới — 1 model đã
  // ngừng sản xuất từ lâu (VD chỉ có dữ liệu 2020-2022) vẫn có thể "thắng"
  // nếu vài năm đó F-Cost% trung bình thấp, và bị chọn làm ⭐ mặc định. Từ
  // khi Chart 1/2 gắn thẳng với model đang chọn (menu5-chart1-chart2-link-
  // selected-model), bộ lọc NĂM mặc định lại là năm gần nhất trong dữ liệu
  // → model mặc định không có dòng nào ở năm đó → 2 chart trống ngay từ đầu,
  // trong khi trước đây (Chart 1/2 là dữ liệu tổng hợp) không ai nhận ra vấn
  // đề này. Fix: ưu tiên model có dữ liệu F-Cost% trong ĐÚNG NĂM GẦN NHẤT
  // xuất hiện trong toàn bộ file (`globalMaxYear`) trước, rồi mới so trung
  // bình F-Cost% của năm đó để chọn thấp nhất — nếu không model nào có dữ
  // liệu năm gần nhất (file toàn dữ liệu cũ), fallback về logic gốc (trung
  // bình toàn lịch sử) để không vỡ hành vi cũ.
  let globalMaxYear = -Infinity;
  for (const m of models) {
    for (const p of m.series[ITEM_FCOST] ?? []) {
      if (p.y > globalMaxYear) globalMaxYear = p.y;
    }
  }
  let bestModel = models[0]?.model ?? '';
  let bestAvg = Infinity;
  if (globalMaxYear > -Infinity) {
    for (const m of models) {
      const vals = (m.series[ITEM_FCOST] ?? []).filter(p => p.y === globalMaxYear).map(p => p.v);
      if (vals.length >= 1) {
        const avg = vals.reduce((a, v) => a + v, 0) / vals.length;
        if (avg < bestAvg) { bestAvg = avg; bestModel = m.model; }
      }
    }
  }
  if (bestAvg === Infinity) {
    // Model "tốt nhất" = F-Cost % trung bình thấp nhất (>=6 điểm dữ liệu) —
    // logic gốc, chỉ dùng khi KHÔNG có model nào có dữ liệu ở năm gần nhất.
    for (const m of models) {
      const vals = (m.series[ITEM_FCOST] ?? []).map(p => p.v);
      if (vals.length >= 6) {
        const avg = vals.reduce((a, v) => a + v, 0) / vals.length;
        if (avg < bestAvg) { bestAvg = avg; bestModel = m.model; }
      }
    }
  }

  // EPCC (menu5-revenue-tab-sales-by-customer) - chuyển Map lồng nhau sang
  // object thuần (JSON-serializable, để IndexedDB cache qua JSON.stringify
  // không mất dữ liệu — xem idbSetCache(IDB_KEY_MENU5_DATA, ...)).
  const salesByCustomerOut: Record<string, Record<number, number>> = {};
  for (const [custom, byYear] of salesByCustomer.entries()) {
    salesByCustomerOut[custom] = Object.fromEntries(
      Array.from(byYear.entries()).map(([y, v]) => [y, Math.round(v * 1000) / 1000])
    );
  }

  // EPCC (menu5-revenue-chart3-viewby-monthly-customer) - chuyển
  // `salesByCustomerMonthly` (Map lồng nhau) sang object thuần, cùng
  // nguyên tắc JSON-serializable với `salesByCustomerOut` ở trên.
  const salesByCustomerMonthlyOut: Record<string, Record<number, Record<number, number>>> = {};
  for (const [custom, byYear] of salesByCustomerMonthly.entries()) {
    const yearsOut: Record<number, Record<number, number>> = {};
    for (const [yKey, byMonth] of byYear.entries()) {
      yearsOut[Number(yKey)] = Object.fromEntries(byMonth.entries());
    }
    salesByCustomerMonthlyOut[custom] = yearsOut;
  }

  return { bestModel, models, salesByCustomer: salesByCustomerOut, salesByCustomerMonthly: salesByCustomerMonthlyOut };
}

// ── Điểm vào chính: gộp TOÀN BỘ sheet trong workbook trước khi tổng hợp ────
// Trước đây handleUploadFile tự đọc wb.SheetNames[0] rồi gọi thẳng hàm dò
// header + đổ dữ liệu — nghĩa là chỉ 1 sheet đầu tiên được xét. Hàm này thay
// thế điểm vào đó: lặp mọi sheet, sheet nào có đủ cột bắt buộc thì gộp vào
// CÙNG 1 acc; sheet không khớp (VD sheet pivot/ghi chú phụ) thì bỏ qua êm,
// không làm hỏng cả file.
function aggregateWorkbookToTest5Data(wb: XLSX.WorkBook): Test5Data {
  const acc = createAccumulator();
  let matchedAnySheet = false;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][];
    const found = findHeaderRow(rows);
    if (!found) continue; // sheet này không có bảng dữ liệu hợp lệ -> bỏ qua, không phải lỗi cả file
    matchedAnySheet = true;
    accumulateSheetRows(rows, found.colIdx, found.headerRowIdx, acc);
  }

  if (!matchedAnySheet) {
    throw new Test5ParseError(
      'Không tìm thấy đủ cột bắt buộc (MODEL, TYPE, ITEM, CUSTOM, YEAR, MONTH, Q\'TY) trong bất kỳ sheet nào của file.'
    );
  }

  return buildTest5DataFromAccumulator(acc);
}

// Ghi chú EPCC (adapt-viewby-thang): dữ liệu Test5 chỉ có độ chi tiết
// NĂM + THÁNG (không có ngày/tuần) — nên "Xem theo" ở đây là Tháng/Quý/Năm
// thay vì Ngày/Tuần/Tháng như bản demo gốc, để không hiển thị sai lệch dữ liệu.
const ITEM_SHIPMENT = 'SHIPMENT (KEA)';
const ITEM_PRODUCTION = "PROD'N (KEA)";
const ITEM_SALES = 'SALES AMT';
const ITEM_SCRAP = 'SCRAP';
const ITEM_FCOST = 'F-Cost %';
// EPCC (menu5-chart5-6-mass-spl-fcost) - THÊM MỚI theo yêu cầu người dùng
// "thêm 2 biểu đồ Mass F-Cost% và SPL F-Cost%, y hệt Chart 2, làm biểu đồ
// 5,6": 2 item này đã tồn tại sẵn trong Excel gốc (cột ITEM = 'Mass F-Cost %'
// / 'SPL F-Cost %'), song song với 'F-Cost %' — không cần sửa logic đọc
// Excel vì `series` đã là Record<string, SeriesPoint[]> generic theo tên
// item, chỉ cần khai báo hằng số để tái dùng đúng hàm build Chart 2.
const ITEM_MASS_FCOST = 'Mass F-Cost %';
const ITEM_SPL_FCOST = 'SPL F-Cost %';

// EPCC (menu5-kpi-target-from-real-target-row) - đã xoá hằng số
// `FCOST_TARGET = 0.05` (5% viết cứng, không đọc từ Excel) — mọi nơi từng
// dùng làm "mục tiêu" (Card 1 + ngưỡng đạt/không-đạt Card 3) giờ tra thẳng
// dòng MODEL='Target' thật trong file (xem `targetModel`/`kpiTargetTtl`/
// `targetMonthlyMap` ở khối tính KPI).

// Palette tái sử dụng đúng 4 biến thể neon đã có sẵn trong App.css
// (không bịa màu mới) — quy ước dùng lại cho radar + bar/line chart.
const NEON = {
  violet: '#a855f7',
  cyan: '#03e9f4',
  amber: '#f5a524',
  pink: '#e8399a',
};

// ── 4 "model" đặc biệt trong cột MODEL của Test5.xlsx — không phải model sản
// phẩm thật, mà là dòng tổng hợp F-Cost cấp OIS ASSY (Target/Actual theo %,
// F-cost($)/F-Cost(억원) theo giá trị tiền) dùng riêng cho biểu đồ tổng quan
// "ALL YEAR" (chart 1 mới). Phải loại các tên này ra khi tính top-N model
// thật theo F-Cost% (chart 2/4 mới), nếu không sẽ lẫn vào danh sách model.
const SPECIAL_FCOST_MODELS = ['Target', 'Actual', 'F-cost($)', 'F- Cost (억원)'];
// EPCC (menu5-model-all-option) - sentinel riêng cho option "Tất cả" trong
// dropdown MODEL, không trùng với tên model thật nào trong Excel.
const ALL_MODEL_VALUE = '__ALL_MODELS__';
// EPCC (menu5-revenue-custom-filter) - hằng số "Tất cả" cho bộ lọc CUSTOM mới
// thêm vào thanh toolbar (tab TÌNH HÌNH DOANH THU), cùng quy ước với
// ALL_MODEL_VALUE ở trên.
const ALL_CUSTOM_VALUE = '__ALL_CUSTOMS__';

// EPCC (menu5-4-new-charts-palette) - bảng màu MỞ RỘNG riêng cho biểu đồ
// nhiều chuỗi (top 9 model trong "MASS F-COST%"): đây là màu MÃ HOÁ DỮ LIỆU
// (data-encoding categorical palette, kiểu Chart.js/D3), không phải màu
// thương hiệu/UI — nên không thuộc phạm vi "chỉ dùng 4 biến thể neon sẵn có"
// của quy ước nút bấm/theme. 4 màu đầu vẫn lấy đúng từ NEON hiện có để tối đa
// tái sử dụng, 5 màu sau bổ sung thêm để đủ phân biệt 9 chuỗi.
const CHART_PALETTE_9 = [
  NEON.pink, NEON.cyan, '#22c55e', NEON.violet, '#f43f5e',
  NEON.amber, '#3b82f6', '#94a3b8', '#facc15', '#14b8a6',
];

// EPCC (menu5-chart256-top10-by-fcost) - số model tối đa vẽ ở Chart 2/5/6
// (F-Cost%/Mass F-Cost%/SPL F-Cost% THEO MODEL) khi bộ lọc MODEL = "Tất cả",
// sắp theo F-Cost giảm dần — xem `filterModelsWithDataInRange()`.
const CHART_TOP_N_MODELS = 10;

// EPCC (menu5-chart-header-match-sales) - copy NGUYÊN VĂN bảng màu thanh
// header 4 khung biểu đồ từ SalesDashboard.tsx (biến CHART_HEADER_THEME,
// comment gốc "header-color-match-template") theo đúng khoanh đỏ ảnh mẫu
// người dùng cung cấp — không tự bịa màu mới, giữ đúng thứ tự 1-4:
// tím #8b5cf6 → cyan #06b6d4 → cam #f59e0b → cam-đỏ #f97316.
const CHART_HEADER_THEME = [
  { accent: '#8b5cf6', bgLight: '#d6c6fc', bgDark: 'rgba(139,92,246,0.14)' },
  { accent: '#06b6d4', bgLight: '#a8e5f0', bgDark: 'rgba(6,182,212,0.14)' },
  { accent: '#f59e0b', bgLight: '#fcddaa', bgDark: 'rgba(255,255,255,0.05)' },
  { accent: '#f97316', bgLight: '#fdcead', bgDark: 'rgba(249,115,22,0.14)' },
] as const;

// EPCC (menu5-chart1-sales-colors) - FIX theo yêu cầu người dùng "lấy màu
// biểu đồ đường và cột như SalesDashboard chart1, áp dụng cho chart 1
// menu5db": không có sẵn source SalesDashboard.tsx trong phạm vi sửa lần
// này, nên 3 mã màu dưới đây được ĐO TRỰC TIẾP từ ảnh chụp màn hình
// EPCC (menu5-chart1-target-actual-color-ref-img) - FIX theo yêu cầu người
// dùng "đổi màu Target/Actual ở ALL YEAR (Chart 1) theo đúng màu tham chiếu
// ở ảnh 1": đo trực tiếp pixel màu chữ số trong ảnh người dùng cung cấp —
// dòng số dao động theo % Actual là màu xanh rêu #0F766E, dòng số cố định
// ~96.4% là màu đỏ #EF4444. Đổi TARGET (đường đứt, cũ #F39C12 cam) sang đỏ
// #EF4444, ACTUAL (đường liền, cũ #00A65A xanh lá) sang xanh rêu #0F766E —
// khớp đúng 2 màu đo được, không tự bịa màu mới. Cột "XUẤT HÀNG (K)"/
// SHIPMENT (`bar`) không nằm trong yêu cầu, giữ nguyên.
// EPCC (menu5-actual-color-purple-ref-img-v2) - FIX theo yêu cầu người dùng
// "Màu Actual biểu đồ đường chuyển sang màu như ảnh tham chiếu": ảnh tham
// chiếu mới cung cấp là swatch màu tím mã HEX #4A26AB — đổi ACTUAL (đường
// liền, cũ xanh rêu #0F766E) sang đúng mã màu này. TARGET (đường đứt, đỏ
// #EF4444) không nằm trong yêu cầu lần này, giữ nguyên.
// EPCC (menu5-shipment-bar-color-ref-img) - FIX theo yêu cầu người dùng
// "Biểu đồ Shipment hình cột Chart 1 đổi sang màu như ảnh, áp dụng cho
// chuyển đổi theme và truy vấn các mục năm/model": ảnh tham chiếu là swatch
// mã HEX #01877E. Đổi `bar` (cột SHIPMENT, cũ #2D7F96) sang màu này. Yêu
// cầu "đồng nhất khi truy vấn năm/model" đồng nghĩa KHÔNG được có ngoại lệ
// đổi màu riêng theo năm/model nào — trong khi đó `highlightBarColors=
// {{ '2020': '#17C3B2' }}` (khai báo ở nơi gọi <ComboBarDualLineChart>) lại
// CHỦ Ý tô riêng cột năm 2020 một màu khác hẳn `bar`. Giữ nguyên override đó
// sẽ trực tiếp vi phạm yêu cầu lần này, nên gỡ bỏ hẳn `highlightBarColors`
// khỏi lời gọi Chart 1 — mọi cột SHIPMENT (mọi năm/model, cả 2 theme, vì
// `bar` không phụ thuộc `theme`) giờ LUÔN dùng chung 1 màu #01877E.
const SALES_CHART1_COLORS = {
  bar: '#01877E',    // cột "XUẤT HÀNG (K)" / SHIPMENT (KEA) — theo ảnh tham chiếu người dùng
  line: '#4A26AB',   // ACTUAL — tím, theo ảnh tham chiếu người dùng cung cấp
  lineDashed: '#EF4444', // TARGET — đỏ, giữ nguyên (không thuộc yêu cầu lần này)
};

const TEXT: Record<Lang, Record<string, string>> = {
  vi: {
    title: 'TÌNH HÌNH DOANH THU - F-COST',
    tabOverview: 'TÌNH HÌNH F-COST',
    tabDetail: 'Chi tiết & dữ liệu',
    // EPCC (menu5-revenue-tab) - THÊM MỚI tab "TÌNH HÌNH DOANH THU", nhân bản
    // y hệt bố cục/màu sắc SalesDashboard.tsx (Mục 1) nhưng dữ liệu lấy từ
    // Test5.xlsx (đã tải sẵn ở Mục 5) — xem khối JSX `tab === 'revenue'`.
    tabRevenue: 'TÌNH HÌNH DOANH THU',
    revKpiSales: 'TỔNG DOANH SỐ',
    revKpiShipment: 'TỔNG XUẤT HÀNG',
    revKpiProduction: 'TỔNG SẢN XUẤT',
    revKpiYoy: 'TĂNG TRƯỞNG DOANH SỐ (YOY)',
    // EPCC (menu5-chart1-2-3-title-one-line-shorten) - FIX theo yêu cầu
    // người dùng "Chart 1 Sửa XUẤT HÀNG & DOANH SỐ THEO NĂM = DOANH SỐ THEO
    // NĂM (cùng 1 dòng). Chart2 di chuyển dòng dưới lên trên thành 1 dòng...
    // Chart3 cũng cho lên thành 1 dòng": Chart1 rút gọn bỏ phần "XUẤT HÀNG &"
    // chỉ còn "DOANH SỐ THEO NĂM" trên 1 dòng; Chart2/Chart3 gộp 2 dòng cũ
    // thành 1 dòng liền (bỏ ký tự xuống dòng `\n`), áp dụng đồng bộ cả 3 ngôn
    // ngữ VI/EN/KO.
    revChart1Title: 'DOANH SỐ THEO NĂM', revChart1Sub: 'Xuất hàng & Doanh số theo năm',
    revChart2Title: 'TOP 10 MODEL THEO DOANH SỐ', revChart2Sub: 'Doanh số theo model',
    revChart3Title: 'DOANH SỐ THEO KHÁCH HÀNG', revChart3Sub: 'Doanh số theo khách hàng qua các năm',
    revChart4Title: 'TỶ TRỌNG THEO KHÁCH HÀNG', revChart4Sub: 'Tỷ trọng khách hàng',
    // EPCC (menu5-revchart5-6-rtv) - THÊM MỚI 2 tiêu đề Chart 5,6 "RTV THEO
    // THÁNG NĂM" / "RTV THEO MODEL", nhân bản đúng ảnh tham chiếu người dùng
    // cung cấp (2 khung "RTV THEO THÁNG NĂM" / "RTV THEO MODEL").
    // EPCC (menu5-chart4-5-6-title-one-line-compact-header) - FIX theo yêu
    // cầu người dùng "Cardtop cho di chuyển chữ như hình vẽ thành 1 dòng và
    // thu gọn thanh Card để biểu đồ phóng to rộng hơn" (áp dụng đúng 3 card
    // có mũi tên đỏ trong ảnh tham chiếu: TỶ TRỌNG THEO KHÁCH HÀNG / RTV THEO
    // THÁNG NĂM / RTV THEO MODEL — card "DOANH SỐ THEO KHÁCH HÀNG" không có
    // mũi tên nên giữ nguyên 2 dòng như cũ): bỏ ký tự xuống dòng `\n` khỏi 3
    // title này để hiển thị liền 1 dòng đúng như ảnh vẽ.
    revChart5Title: 'RTV THEO THÁNG NĂM', revChart5Sub: 'RTV theo tháng trong năm',
    revChart6Title: 'RTV THEO MODEL', revChart6Sub: 'RTV theo model',
    revTableTitle: 'Toàn bộ Model theo Doanh số',
    revColModel: 'MODEL', revColCustomer: 'KHÁCH HÀNG', revColType: 'TYPE',
    revColProd: 'SẢN XUẤT (K)', revColShip: 'XUẤT HÀNG (K)', revColSales: 'DOANH SỐ (K$)', revColRatio: 'TỶ LỆ XH/SX',
    revColRtv: 'RTV (K)', revColRtvRate: '% RATE', revColGrowth: 'TĂNG TRƯỞNG DOANH SỐ', revColScrap: 'SCRAP (K)',
    revColPeriod: 'NĂM/KỲ',
    revFilterModelLabel: 'Model:', revFilterYearLabel: 'Năm:', revFilterQuarterLabel: 'Quý:',
    revFilterMonthLabel: 'Tháng:', revFilterCustomerLabel: 'Khách hàng:',
    revFilterSearchPlaceholder: 'Tìm kiếm Model, Đối tác, Khách hàng',
    revLegendShip: 'XUẤT HÀNG (K)', revLegendProd: 'SẢN XUẤT (K)', revLegendSales: 'DOANH SỐ (K$)', revLegendYoy: 'Tr.trưởng YoY',
    revLegendShipKea: 'SHIPMENT(KEA)', revLegendRtvKea: 'RTV(KEA)', revLegendRtvRate: '% Rate',
    revYoyPanelTitle: 'Tăng giảm giữa\ncác năm (K$)',
    revYoyVs: 'so với',
    startMonth: 'THÁNG BẮT ĐẦU',
    endMonth: 'THÁNG KẾT THÚC',
    // EPCC (menu5-year-quick-filter) - FIX theo yêu cầu người dùng "bổ sung
    // thêm 1 ô truy vấn theo năm: tất cả/2017~2026": thêm nhãn cho ô lọc
    // nhanh theo năm (đứng cạnh nút Tháng/Quý/Năm ở "XEM THEO") — khác với
    // key `byYear` (chữ trên nút bấm chế độ xem), đây là nhãn tiêu đề CỘT
    // + option "Tất cả" trong dropdown.
    yearFilterLabel: 'NĂM',
    yearFilterAll: 'Tất cả',
    // EPCC (menu5-model-all-option) - thêm nhãn cho option "Tất cả" trong
    // dropdown MODEL: khi chọn, KPI + bảng chi tiết + Chart 1 & 2 đều dùng
    // thẳng dòng "Actual" (tổng hợp toàn công ty) đã có sẵn trong Excel,
    // thay vì dữ liệu của 1 model lẻ. Xem chi tiết ở modelData/chart2Models.
    modelFilterAll: 'Tất cả',
    model: 'MODEL',
    // EPCC (menu5-revenue-custom-filter) - nhãn + option "Tất cả" cho bộ lọc
    // CUSTOM mới thêm vào toolbar, đặt tên y hệt cột CUSTOM trong Excel
    // (theo đúng ảnh tham chiếu người dùng cung cấp), không dịch, cùng quy
    // ước literal như nhãn "MODEL" ở dòng dưới.
    customFilter: 'CUSTOM',
    customFilterAll: 'Tất cả',
    viewBy: 'XEM THEO',
    byMonth: 'Tháng',
    byQuarter: 'Quý',
    byYear: 'Năm',
    // EPCC (menu5-toolbar-import-not-export) - đổi tên key: nút này BẤM VÀO
    // MỞ HỘP THOẠI CHỌN FILE ĐỂ TẢI LÊN (import), không phải xuất/tải xuống
    // báo cáo. Tên cũ "exportExcel" gắn với handleExport() (ghi file mới)
    // gây hiểu sai đúng như người dùng phản ánh — xem chỗ dùng ở toolbar.
    importExcelBtn: 'Tải tệp lên',
    updatedToLabel: 'DỮ LIỆU CẬP NHẬT ĐẾN',
    kpi1: 'F-COST TB',
    kpi2: 'SCRAP AMT',
    kpi3: 'THÁNG ĐẠT MỤC TIÊU',
    kpi4: 'CHÊNH LỆCH SO VỚI NĂM TRƯỚC',
    kpiTarget: 'Mục tiêu',
    kpiActual: 'Thực tế',
    chart3: 'SHIPMENT THEO THÁNG',
    chart3sub: 'Sản lượng xuất hàng',
    chart4: "PROD'N THEO THÁNG",
    chart4sub: 'Sản lượng sản xuất',
    chart5: 'SALES AMT THEO THÁNG',
    chart5sub: 'Doanh số bán hàng',
    chart6: 'SCRAP THEO THÁNG',
    chart6sub: 'Sản lượng phế phẩm',
    noData: 'Không có dữ liệu trong khoảng đã chọn',
    nc1Title: 'F-COST % THEO NĂM', nc1Sub: 'F-Cost theo năm — Target vs Actual',
    nc2Title: 'F-COST %\nTHEO MODEL', nc2Sub: 'F-Cost% theo tháng — Top 10 model',
    nc7Title: 'F-COST %\nTHEO MODEL (NHỆN)', nc7Sub: 'F-Cost% theo tháng — Top 10 model (dạng radar)',
    nc3Title: 'F-COST(K$) LŨY KẾ THEO NĂM', nc3Sub: '',
    nc4Title: 'F-COST(K$) THEO NĂM', nc4Sub: '',
    nc5Title: 'MASS F-COST %\nTHEO MODEL', nc5Sub: 'Mass F-Cost% theo tháng — Top 10 model',
    nc6Title: 'SPL F-COST %\nTHEO MODEL', nc6Sub: 'SPL F-Cost% theo tháng — Top 10 model',
    legendFcost: 'F-Cost ($)', legendFcostWon: 'F-Cost (억원)', legendShipment: 'SHIPMENT (KEA)', legendTarget: 'TARGET', legendActual: 'ACTUAL', colMonth: 'Tháng', colShipment: 'Shipment', colProd: "Prod'n",
    colSales: 'Sales Amt', colScrap: 'Scrap', colFcost: 'F-Cost %',
    uploadExcel: 'Tải lên Excel mới',
    parsing: 'Đang xử lý...',
    noDataYet: '📁 Chưa có dữ liệu, vui lòng tải Excel lên',
    emptyDesc: 'Vui lòng tải lên tệp Excel Test5 để bắt đầu xem báo cáo.',
    dataUpdatedAt: '✅ Đã cập nhật lúc',
  },
  en: {
    title: 'REVENUE - F-COST STATUS',
    tabOverview: 'Overview',
    tabDetail: 'Detail & Data',
    tabRevenue: 'Revenue Overview',
    revKpiSales: 'TOTAL SALES',
    revKpiShipment: 'TOTAL SHIPMENT',
    revKpiProduction: 'TOTAL PRODUCTION',
    revKpiYoy: 'SALES GROWTH (YOY)',
    revChart1Title: 'SALES BY YEAR', revChart1Sub: 'Shipment & Sales by year',
    revChart2Title: 'TOP 10 MODELS BY SALES', revChart2Sub: 'Sales by model',
    revChart3Title: 'SALES BY CUSTOMER', revChart3Sub: 'Sales by customer over years',
    revChart4Title: 'SHARE BY CUSTOMER', revChart4Sub: 'Customer share',
    revChart5Title: 'RTV BY MONTH/YEAR', revChart5Sub: 'RTV by month in year',
    revChart6Title: 'RTV BY MODEL', revChart6Sub: 'RTV by model',
    revTableTitle: 'All Models by Sales',
    revColModel: 'MODEL', revColCustomer: 'CUSTOMER', revColType: 'TYPE',
    revColProd: "PROD'N (K)", revColShip: 'SHIPMENT (K)', revColSales: 'SALES (K$)', revColRatio: 'SHIP/PROD RATIO',
    revColRtv: 'RTV (K)', revColRtvRate: '% RATE', revColGrowth: 'SALES GROWTH', revColScrap: 'SCRAP (K)',
    revColPeriod: 'YEAR/PERIOD',
    revFilterModelLabel: 'Model:', revFilterYearLabel: 'Year:', revFilterQuarterLabel: 'Quarter:',
    revFilterMonthLabel: 'Month:', revFilterCustomerLabel: 'Customer:',
    revFilterSearchPlaceholder: 'Search Model, Partner, Customer',
    revLegendShip: 'SHIPMENT (K)', revLegendProd: "PROD'N (K)", revLegendSales: 'SALES AMT (K$)', revLegendYoy: 'YoY Growth',
    revLegendShipKea: 'SHIPMENT(KEA)', revLegendRtvKea: 'RTV(KEA)', revLegendRtvRate: '% Rate',
    revYoyPanelTitle: 'Change between\nyears (K$)',
    revYoyVs: 'vs',
    endMonth: 'END MONTH',
    yearFilterLabel: 'YEAR',
    yearFilterAll: 'All',
    modelFilterAll: 'All',
    model: 'MODEL',
    customFilter: 'CUSTOM',
    customFilterAll: 'All',
    viewBy: 'VIEW BY',
    byMonth: 'Month',
    byQuarter: 'Quarter',
    byYear: 'Year',
    importExcelBtn: 'Upload File',
    updatedToLabel: 'DATA UPDATED TO',
    kpi1: 'AVG F-COST',
    kpi2: 'SCRAP AMT',
    kpi3: 'MONTHS ON TARGET',
    kpi4: 'CHANGE VS LAST YEAR',
    kpiTarget: 'Target',
    kpiActual: 'Actual',
    chart3: 'SHIPMENT BY MONTH',
    chart3sub: 'Shipment volume',
    chart4: "PROD'N BY MONTH",
    chart4sub: 'Production volume',
    chart5: 'SALES AMT BY MONTH',
    chart5sub: 'Sales amount',
    chart6: 'SCRAP BY MONTH',
    chart6sub: 'Scrap volume',
    noData: 'No data in the selected range',
    nc1Title: 'ALL YEAR', nc1Sub: 'F-Cost by year — Target vs Actual',
    nc2Title: 'F-COST\nBY MODEL', nc2Sub: 'Monthly F-Cost% — Top 10 models',
    nc7Title: 'F-COST %\nBY MODEL (RADAR)', nc7Sub: 'Monthly F-Cost% — Top 10 models (radar view)',
    nc3Title: 'F-COST(K$) CUMULATIVE BY YEAR', nc3Sub: '',
    nc4Title: 'F-COST(K$) BY YEAR', nc4Sub: '',
    nc5Title: 'MASS F-COST %\nBY MODEL', nc5Sub: 'Monthly Mass F-Cost% — Top 10 models',
    nc6Title: 'SPL F-COST %\nBY MODEL', nc6Sub: 'Monthly SPL F-Cost% — Top 10 models',
    legendFcost: 'F-Cost ($)', legendFcostWon: 'F-Cost (억원)', legendShipment: 'SHIPMENT (KEA)', legendTarget: 'TARGET', legendActual: 'ACTUAL',
    colYear: 'Year', colMonth: 'Month', colShipment: 'Shipment', colProd: "Prod'n",
    colSales: 'Sales Amt', colScrap: 'Scrap', colFcost: 'F-Cost %',
    uploadExcel: 'Upload new Excel',
    parsing: 'Processing...',
    noDataYet: '📁 No data yet, please upload an Excel file',
    emptyDesc: 'Please upload the Test5 Excel file to start viewing the report.',
    dataUpdatedAt: '✅ Updated at',
  },
  ko: {
    title: '매출 - F-COST 현황',
    tabOverview: '전체 현황',
    tabDetail: '상세 데이터',
    tabRevenue: '매출 현황',
    revKpiSales: '총 매출',
    revKpiShipment: '총 출하',
    revKpiProduction: '총 생산',
    revKpiYoy: '매출 증감 (YOY)',
    revChart1Title: '연도별 매출', revChart1Sub: '연도별 출하 & 매출',
    revChart2Title: '매출 상위 10개 모델', revChart2Sub: '모델별 매출',
    revChart3Title: '고객사별 매출', revChart3Sub: '연도별 고객사 매출',
    revChart4Title: '고객사별 비중', revChart4Sub: '고객사 비중',
    revChart5Title: '월별 RTV', revChart5Sub: '연중 월별 RTV',
    revChart6Title: '모델별 RTV', revChart6Sub: '모델별 RTV',
    revTableTitle: '전체 모델 매출 현황',
    revColModel: '모델', revColCustomer: '고객사', revColType: '타입',
    revColProd: '생산 (K)', revColShip: '출하 (K)', revColSales: '매출 (K$)', revColRatio: '출하/생산 비율',
    revColRtv: 'RTV (K)', revColRtvRate: '% RATE', revColGrowth: '매출 성장', revColScrap: 'SCRAP (K)',
    revColPeriod: '연도/기간',
    revFilterModelLabel: '모델:', revFilterYearLabel: '연도:', revFilterQuarterLabel: '분기:',
    revFilterMonthLabel: '월:', revFilterCustomerLabel: '고객사:',
    revFilterSearchPlaceholder: '모델, 파트너, 고객 검색',
    revLegendShip: '출하 (K)', revLegendProd: '생산 (K)', revLegendSales: '매출 (K$)', revLegendYoy: '전년 대비 증감',
    revLegendShipKea: 'SHIPMENT(KEA)', revLegendRtvKea: 'RTV(KEA)', revLegendRtvRate: '% Rate',
    revYoyPanelTitle: '연도별 증감\n(K$)',
    revYoyVs: 'vs',
    startMonth: '시작 월',
    endMonth: '종료 월',
    yearFilterLabel: '연도',
    yearFilterAll: '전체',
    modelFilterAll: '전체',
    model: '모델',
    customFilter: 'CUSTOM',
    customFilterAll: '전체',
    viewBy: '보기 기준',
    byMonth: '월',
    byQuarter: '분기',
    byYear: '년',
    importExcelBtn: '파일 업로드',
    updatedToLabel: '데이터 업데이트 기준일',
    kpi1: '평균 F-Cost',
    kpi2: 'SCRAP 금액',
    kpi3: '목표 달성 월수',
    kpi4: '전년 대비 변화',
    kpiTarget: '목표',
    kpiActual: '실제',
    chart3: '월별 Shipment',
    chart3sub: '출하량',
    chart4: '월별 생산량',
    chart4sub: '생산량',
    chart5: '월별 매출액',
    chart5sub: '매출액',
    chart6: '월별 Scrap',
    chart6sub: '불량 수량',
    noData: '선택한 기간에 데이터가 없습니다',
    nc1Title: 'ALL YEAR', nc1Sub: '연도별 F-Cost — Target vs Actual',
    nc2Title: '모델별\nF-COST', nc2Sub: '월별 F-Cost% — 상위 10개 모델',
    nc7Title: '모델별\nF-COST % (레이더)', nc7Sub: '월별 F-Cost% — 상위 10개 모델 (레이더 차트)',
    nc3Title: '연도별 누적 F-COST(K$)', nc3Sub: '',
    nc4Title: '연도별 F-COST(K$)', nc4Sub: '',
    nc5Title: '모델별\nMASS F-COST %', nc5Sub: '월별 Mass F-Cost% — 상위 10개 모델',
    nc6Title: '모델별\nSPL F-COST %', nc6Sub: '월별 SPL F-Cost% — 상위 10개 모델',
    legendFcost: 'F-Cost ($)', legendFcostWon: 'F-Cost (억원)', legendShipment: 'SHIPMENT (KEA)', legendTarget: 'TARGET', legendActual: 'ACTUAL',
    colYear: '년', colMonth: '월', colShipment: 'Shipment', colProd: '생산량',
    colSales: '매출액', colScrap: 'Scrap', colFcost: 'F-Cost %',
    uploadExcel: '새 엑셀 업로드',
    parsing: '처리 중...',
    noDataYet: '📁 데이터가 없습니다. 엑셀 파일을 업로드하세요',
    emptyDesc: 'Test5 엑셀 파일을 업로드하여 리포트를 시작하세요.',
    dataUpdatedAt: '✅ 업데이트:',
  },
};

function ym(y: number, m: number) { return y * 12 + (m - 1); }

function fmtNum(n: number, digits = 1) {
  return n.toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: 0 });
}

// EPCC (menu5-dead-code-cleanup) - FIX theo yêu cầu người dùng "loại bỏ code
// thừa": 2 component `BarLineChart` và `RadarChart` từng tồn tại ở đây
// KHÔNG còn được gọi ở bất kỳ đâu trong file (đã grep xác nhận: 0 lần gọi
// ngoài chính định nghĩa) — là code chết còn sót lại từ giai đoạn phát triển
// trước, không phục vụ tính năng nào đang hiển thị cho người dùng. Xoá hẳn
// để giảm ~100 dòng không cần thiết, không ảnh hưởng bất kỳ chart nào đang
// chạy (4 chart mới dùng `ComboBarDualLineChart`/`MultiLineChart`/
// `RoseChart`, hoàn toàn tách biệt).
// ═══════════════════════════════════════════════════════════════════════
// EPCC (menu5-4-new-charts) - 3 component biểu đồ mới, nhân bản đúng hình
// mẫu người dùng cung cấp, dùng chung SVG thuần (không thư viện ngoài) như
// các chart hiện có. Xem giải thích mapping dữ liệu ở phần tính toán trong
// component chính (comment "== 4 BIỂU ĐỒ MỚI ==").
// ═══════════════════════════════════════════════════════════════════════

// ── Chart 1: Bar (F-Cost $) + 2 line trục phải (TARGET/ACTUAL %) ──────────
// EPCC (menu5-chart1-2-smooth-axis-gap30) - FIX theo yêu cầu người dùng
// "gap width của biểu đồ cột còn 30%, biểu đồ đường chuyển thành smoothed,
// phóng to hai bên trái phải, hiển thị trục X/Y cho cả 2 biểu đồ, ẩn giá trị
// khi chưa có của các tháng ở biểu đồ 1": trước đây 2 đường (target/actual ở
// Chart 1, các series F-Cost% ở Chart 2) vẽ bằng <polyline> nối thẳng từng
// điểm — không có cách "làm mượt" polyline thuần SVG. Thêm hàm dùng chung
// buildSmoothLinePath(): tách mảng giá trị thành từng đoạn liên tục (ngắt
// đoạn tại điểm null, giữ đúng hành vi "đứt đường khi chưa có dữ liệu" như
// cũ), rồi nối các điểm trong từng đoạn bằng Catmull-Rom → Cubic Bezier
// (tangent = (p[i+1]-p[i-1])/6, chuẩn suy ra từ centripetal Catmull-Rom) để
// ra đường cong mượt, sau đó ghép nhiều đoạn path bằng nhiều lệnh "M" trong
// cùng 1 thuộc tính `d`. Áp dụng cho cả 2 biểu đồ, thay <polyline> bằng
// <path>.
function smoothPathFromPoints(pts: { x: number; y: number }[]): string {
  if (pts.length === 0) return '';
  if (pts.length === 1) return `M${pts[0].x},${pts[0].y}`;
  let d = `M${pts[0].x},${pts[0].y}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] || pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] || p2;
    const c1x = p1.x + (p2.x - p0.x) / 6;
    const c1y = p1.y + (p2.y - p0.y) / 6;
    const c2x = p2.x - (p3.x - p1.x) / 6;
    const c2y = p2.y - (p3.y - p1.y) / 6;
    d += ` C${c1x},${c1y} ${c2x},${c2y} ${p2.x},${p2.y}`;
  }
  return d;
}
function buildSmoothLinePath(
  vals: (number | null)[],
  xOf: (i: number) => number,
  yOf: (v: number) => number,
): string {
  const segments: { x: number; y: number }[][] = [];
  let current: { x: number; y: number }[] = [];
  vals.forEach((v, i) => {
    if (v == null) {
      if (current.length) segments.push(current);
      current = [];
    } else {
      current.push({ x: xOf(i), y: yOf(v) });
    }
  });
  if (current.length) segments.push(current);
  return segments.map(seg => smoothPathFromPoints(seg)).join(' ');
}

function ComboBarDualLineChart({
  bars, bars2, bar2Color, target, actual, barColor, targetColor, actualColor, height = 240, highlightBarColors, theme,
}: {
  bars: { label: string; value: number | null }[];
  bars2?: (number | null)[];
  bar2Color?: string;
  target: (number | null)[];
  actual: (number | null)[];
  barColor: string; targetColor: string; actualColor: string;
  height?: number;
  // EPCC (menu5-chart1-bar-highlight-color) - FIX theo yêu cầu người dùng
  // "Chuyển hình cột khoanh đỏ sang màu #17C3B2" (cột năm 2020 trong Chart
  // 1): trước đây MỌI cột F-cost($) dùng chung 1 màu `barColor` cố định,
  // không có cách đổi riêng 1 cột theo `label` mà không ảnh hưởng các cột
  // khác. Thêm prop tuỳ chọn `highlightBarColors` — map `label` (VD "2020")
  // sang màu riêng; cột nào có label khớp key trong map thì dùng màu đó
  // thay cho `barColor` mặc định, các cột còn lại không đổi.
  highlightBarColors?: Record<string, string>;
  // EPCC (menu5-chart-axis-text-theme-aware) - FIX theo yêu cầu người dùng
  // "màu chữ hiển thị trong biểu đồ cũng thay đổi khi đổi theme": thêm prop
  // `theme` để tính màu chữ trục Y trái/phải + nhãn tháng/năm trục X — trước
  // đây các nhãn này cố định `fill="#000000"`, ở theme dark khó đọc trên nền
  // tối. Dùng chung biến `axisTextColor` bên dưới cho mọi nhãn thuộc nhóm
  // này trong component.
  theme: ThemeMode;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const width = 620;
  // EPCC (menu5-chart1-2-restore-axis-gutter) - FIX theo yêu cầu người dùng
  // "dữ liệu số hiển thị hết ra, cột xanh vào trong inside, trục X/Y cho số
  // ra ngoài, và để left/right=0": lượt trước đặt padL=padR=0 làm nhãn số
  // trục Y ("1,755", "878", "4%", "2%"...) rơi RA NGOÀI viewBox 0..width nên
  // bị SVG cắt mất (đúng như ảnh người dùng khoanh đỏ). Không thể vừa để
  // cột/đường chạm sát mép TUYỆT ĐỐI vừa hiện đủ chữ số cạnh đó — 2 yêu cầu
  // này đối lập nhau về mặt hình học. Cách dung hoà đúng ý người dùng: dành
  // riêng 1 dải "gutter" NHỎ (không phải margin to như bản gốc 44px) chỉ đủ
  // rộng cho chữ số trục Y, đặt nhãn NẰM TRONG dải gutter đó (ngoài vùng cột/
  // đường — đúng nghĩa "trục cho số ra ngoài"), còn cột (bar) được kéo "vào
  // trong" (inset) đúng bằng đúng bề rộng gutter đó thay vì chạm mép tuyệt
  // đối 0. Khoảng trắng thừa kiểu margin lớn (44px cũ) thì bỏ hẳn — đó mới
  // là phần "để left/right=0" người dùng muốn xoá.
  // EPCC (menu5-chart1-2-axis-label-clip-fix) - FIX ROOT CAUSE "số trục Y
  // trái bị ẩn/cắt mất (VD ',962' thay vì '2,962')": gutter cũ (padL=30) chỉ
  // đủ chỗ cho số 3 chữ số ngắn; khi dữ liệu lớn hơn (4-5 ký tự kể cả dấu
  // phẩy như '2,962') thì bề rộng chữ (fontSize 10.2, in đậm) vượt quá 24px
  // (padL-6) sẵn có, phần đầu chữ bị đẩy ra ngoài x=0 của viewBox nên bị cắt
  // (SVG không tự động co chữ). Tăng padL 30→42 để đủ chỗ cho số dài hơn,
  // không còn bị cắt ở mọi trường hợp dữ liệu.
  const padL = 42, padR = 28, padT = 26, padB = 26;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;
  const n = Math.max(1, bars.length);
  const step = innerW / n;
  // EPCC (menu5-chart1-gap20) - FIX theo yêu cầu người dùng "Gap width của
  // biểu đồ cột Shipment và F-cost để còn 20%": bề rộng cột (barW) dùng
  // chung cho CẢ 2 cột (bars=SHIPMENT, bars2=F-Cost(억원) chồng lên trên) vì
  // cùng vẽ tại 1 vị trí `xOf(i)` — tăng barW từ 70%→80% mỗi "step" để gap
  // (khoảng trống giữa các cột) giảm tương ứng 30%→20%.
  const barW = step * 0.8;

  // EPCC (menu5-chart1-fcost-won-stack) - maxBar phải cộng luôn bars2 (nếu
  // có) vào giá trị lớn nhất, nếu không phần chồng thêm có thể bị đẩy vượt
  // quá đỉnh biểu đồ ở năm có cả 2 series cùng lúc.
  const maxBar = Math.max(1, ...bars.map((b, i) => (b.value ?? 0) + (bars2?.[i] ?? 0)));
  const maxPct = Math.max(0.01, ...target.filter((v): v is number => v != null),
    ...actual.filter((v): v is number => v != null));

  const xOf = (i: number) => padL + step * i + step / 2;
  const _yBar = (v: number) => padT + innerH - (v / maxBar) * innerH; void _yBar;
  const yPct = (v: number) => padT + innerH - (v / maxPct) * innerH;

  return (
    // EPCC (menu5-chart1-2-svg-aspect-ratio-gap) - FIX ROOT CAUSE "vẫn còn
    // khoảng trống 2 bên trục X/Y dù padL/padR đã giảm tối đa": nguyên nhân
    // KHÔNG nằm ở padL/padR trong code (đã đúng), mà ở thuộc tính SVG mặc
    // định `preserveAspectRatio="xMidYMid meet"` — vì `width="100%"` (co
    // giãn theo card) nhưng `height={height}` là SỐ PX CỐ ĐỊNH, khi khung
    // card thực tế rộng hơn tỉ lệ viewBox (620:height), trình duyệt KHÔNG
    // kéo giãn nội dung theo chiều ngang mà giữ nguyên tỉ lệ gốc rồi CĂN GIỮA
    // — phần dư ra chính là 2 dải trống trái/phải nhìn thấy trong ảnh. Đây
    // không phải "code thừa" cần xoá, mà là 1 thuộc tính CÒN THIẾU: thêm
    // `preserveAspectRatio="none"` để buộc SVG kéo giãn khớp đúng 100% khung
    // thực tế theo cả 2 chiều, xoá hẳn khoảng trống này.
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} preserveAspectRatio="none" role="img">
      {[0, 0.5, 1].map((tt, i) => (
        <line key={i} x1={padL} x2={width - padR} y1={padT + innerH * (1 - tt)} y2={padT + innerH * (1 - tt)}
          stroke="var(--border-soft, #2a2f3a)" strokeWidth={1} strokeDasharray="3,4" />
      ))}
      {/* EPCC (menu5-chart1-2-smooth-axis-gap30) - trục Y trái (thang cột
          F-cost) + trục Y phải (thang %) + trục X, cùng nhãn số tại 3 mốc
          0/50%/100% của mỗi thang — trước đây chỉ có gridline chấm chấm,
          không có đường trục lẫn nhãn số. */}
      <line x1={padL} x2={padL} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={width - padR} x2={width - padR} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={padL} x2={width - padR} y1={padT + innerH} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      {/* EPCC (menu5-circled-text-bold-black-20pct) - FIX theo yêu cầu người
          dùng "các vùng khoanh đỏ hãy cho chữ đen đậm và tăng cỡ chữ 20%":
          nhãn trục Y trái (1,755 / 878 / 0) — tăng fontSize 8.5→10.2
          (+20%), đổi màu từ `var(--text-2)` (xám nhạt) sang đen cố định,
          thêm `fontWeight={700}` để in đậm. */}
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yl${i}`} x={padL - 6} y={padT + innerH * (1 - tt) + 3} fontSize={10.2} fontWeight={700} textAnchor="end" fill={axisTextColor}>
          {fmtNum(maxBar * tt, 0)}
        </text>
      ))}
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yr${i}`} x={width - padR + 6} y={padT + innerH * (1 - tt) + 3} fontSize={10.2} fontWeight={700} textAnchor="start" fill={axisTextColor}>
          {(maxPct * tt * 100).toFixed(0)}%
        </text>
      ))}
      {bars.map((b, i) => {
        if (b.value == null) return null;
        const h = (b.value / maxBar) * innerH;
        const x = xOf(i) - barW / 2;
        const y = padT + innerH - h;
        return <rect key={i} x={x} y={y} width={barW} height={Math.max(h, 1)} rx={2} fill={highlightBarColors?.[b.label] ?? barColor} opacity={0.75} />;
      })}
      {/* EPCC (menu5-chart1-fcost-won-stack) - đoạn chồng F-Cost(억원) lên
          NGAY TRÊN đỉnh bar F-cost($), cùng trục/thang đo (đơn vị khác nhau
          nên đoạn này có thể rất mỏng ở các năm có cả 2 — đúng bản chất dữ
          liệu, xem giải thích ở nơi tính chart1). */}
      {bars2 && bars2.map((v, i) => {
        if (v == null) return null;
        const base = bars[i]?.value ?? 0;
        const h = (v / maxBar) * innerH;
        const x = xOf(i) - barW / 2;
        const y = padT + innerH - (base / maxBar) * innerH - h;
        return <rect key={`w${i}`} x={x} y={y} width={barW} height={Math.max(h, 1.5)} rx={2} fill={bar2Color ?? barColor} opacity={0.9} />;
      })}
      {/* EPCC (menu5-chart1-sales-colors): đường target đổi sang NÉT ĐỨT,
          khớp đúng phong cách đường "Tr.trưởng YoY" trong ảnh tham chiếu
          SalesDashboard chart1 — đường actual giữ nét liền. */}
      <path d={buildSmoothLinePath(target, xOf, yPct)} fill="none" stroke={targetColor} strokeWidth={2} strokeDasharray="6,4" strokeLinejoin="round" strokeLinecap="round" />
      <path d={buildSmoothLinePath(actual, xOf, yPct)} fill="none" stroke={actualColor} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round" />
      {target.map((v, i) => v == null ? null : (
        <g key={`tg${i}`}>
          <circle cx={xOf(i)} cy={yPct(v)} r={2.5} fill={targetColor} />
          {/* EPCC (menu5-chart1-line-label-precision-size) - FIX theo yêu
              cầu người dùng "2 biểu đồ đường Chart 1 thêm 1 số thập phân,
              tăng cỡ chữ 30%": đổi `toFixed(1)` → `toFixed(2)` (VD "1.2%" →
              "1.20%") và `fontSize` 9 → 11.7 (9 * 1.3, đúng +30%) cho CẢ 2
              nhãn Target/Actual — không đổi màu/vị trí (y-7/y+13) đã có. */}
          <text x={xOf(i)} y={yPct(v) - 7} fontSize={11.7} fontWeight={700} textAnchor="middle" fill={targetColor}>
            {(v * 100).toFixed(2)}%
          </text>
        </g>
      ))}
      {actual.map((v, i) => v == null ? null : (
        <g key={`ac${i}`}>
          <circle cx={xOf(i)} cy={yPct(v)} r={2.5} fill={actualColor} />
          {/* EPCC (menu5-chart1-bar-inside-actual-labels) - FIX theo yêu cầu
              người dùng "biểu đồ đường cũng cho hiển thị số ra": trước đây
              chỉ đường TARGET có nhãn %, đường ACTUAL chỉ có chấm tròn không
              có số (2 dấu "?" trong ảnh người dùng khoanh đỏ). Thêm nhãn %
              cho ACTUAL, đặt DƯỚI điểm (y+13, ngược hướng với nhãn TARGET đặt
              TRÊN điểm ở y-7) để 2 nhãn không đè lên nhau khi 2 đường gần
              nhau. */}
          <text x={xOf(i)} y={yPct(v) + 13} fontSize={11.7} fontWeight={700} textAnchor="middle" fill={actualColor}>
            {(v * 100).toFixed(2)}%
          </text>
        </g>
      ))}
      {/* Nhãn F-Cost(억원) — đặt NỔI BẬT (bold, màu riêng) ngay trên đỉnh cột
          tổng, vì đoạn bar chồng có thể quá mỏng để tự đọc được số. */}
      {bars2 && bars2.map((v, i) => {
        if (v == null) return null;
        const base = bars[i]?.value ?? 0;
        const topY = padT + innerH - ((base + v) / maxBar) * innerH;
        return (
          <text key={`wl${i}`} x={xOf(i)} y={Math.max(topY - 6, padT + 8)} fontSize={9.5} fontWeight={700}
            textAnchor="middle" fill={bar2Color ?? barColor}>
            {fmtNum(v, 1)}
          </text>
        );
      })}
      {/* EPCC (menu5-chart1-bar-inside-actual-labels) - FIX theo yêu cầu
          người dùng "cho số vào inside biểu đồ cột": trước đây nhãn giá trị
          đặt Ở DƯỚI, ngoài vùng cột (height-padB+14) — trùng vị trí với nhãn
          năm/tháng bên dưới trục X, gây đè chữ lên nhau (khoanh đỏ trong
          ảnh). Dời nhãn vào NGAY BÊN TRONG đỉnh cột (barTop + 14), đổi màu
          chữ sang trắng để tương phản rõ trên nền cột teal; dùng Math.min để
          nhãn không tụt ra khỏi đáy cột với các cột rất thấp (VD 27, 240).
      */}
      {/* EPCC (menu5-shipment-label-move-to-bottom) - FIX theo yêu cầu người
          dùng "Chart1 phần Shipment để giá trị xuống đáy hộp cho số liệu
          biểu đồ đường nhìn rõ hơn": đặt nhãn ở gần ĐỈNH cột (barTop+14, xem
          comment trên) khiến với các cột cao (gần chạm đỉnh khung), nhãn số
          SHIPMENT color trắng nằm chen ngay cạnh vùng nhãn % của 2 đường
          TARGET/ACTUAL phía trên — gây rối mắt, khó đọc số % (đúng khoanh đỏ
          trong ảnh người dùng). Đổi sang vị trí CỐ ĐỊNH ở đáy hộp (padT +
          innerH - 6, sát trục X, bên trong đáy cột) cho MỌI cột bất kể chiều
          cao, tách hẳn khỏi vùng nhãn % phía trên — đường biểu đồ Target/
          Actual giờ có trọn khoảng trống phía trên để hiển thị rõ.
      */}
      {bars.map((b, i) => {
        if (b.value == null) return null;
        const ty = padT + innerH - 6;
        return (
          <text key={`bv${i}`} x={xOf(i)} y={ty} fontSize={9.5} fontWeight={700} textAnchor="middle" fill="#ffffff">
            {fmtNum(b.value, 0)}
          </text>
        );
      })}
      {bars.map((b, i) => (
        <text key={`lb${i}`} x={xOf(i)} y={height - 4} fontSize={12} fontWeight={700} textAnchor="middle" fill={axisTextColor}>
          {b.label}
        </text>
      ))}
    </svg>
  );
}

// ── Chart 2: nhiều đường F-Cost% theo top-N model, chung 1 trục X (TTL+12 tháng) ──
// EPCC (menu5-chart34-y-axis-fixed-vs-auto) - hàm dùng chung tính trần trục
// Y TỰ ĐỘNG cho MultiLineChart khi không truyền `maxV` cố định (hiện chỉ
// Chart 6/SPL F-Cost% dùng nhánh này): lấy giá trị lớn nhất thật trong toàn
// bộ series, cộng thêm 10% khoảng đệm phía trên (để điểm cao nhất không dính
// sát mép trên cùng), rồi làm tròn LÊN mốc 0.5% (0.005) gần nhất để 4 mốc
// gridline (0/25/50/75/100%) vẫn ra số tròn dễ đọc thay vì số lẻ ngẫu nhiên.
function computeAutoMaxV(series: { values: (number | null)[] }[]): number {
  const allVals = series.flatMap(s => s.values.filter((v): v is number => v != null));
  if (!allVals.length) return 0.01;
  const dataMax = Math.max(...allVals);
  const step = 0.005;
  return Math.max(step, Math.ceil((dataMax * 1.1) / step) * step);
}

// EPCC (menu5-chart234-dual-y-axis-5pct) - FIX theo yêu cầu người dùng "tại
// Chart2,3,4 đối với các giá trị model nào lớn hơn 5% hãy chuyển sang trục Y
// [phụ] rồi thể hiện ra trục Y": trước đây CHỈ CÓ 1 trục Y bên trái dùng
// chung cho MỌI model — model nào có đỉnh vượt trần (VD SO1C2EH ~28% ở Chart
// 2 trần 3.5%, hay 68.4%/28.1% ở SPL) đều bị KẸP (clamp) sát mép trên, mất
// hẳn hình dạng đường và không đọc được số thật trên trục. Fix: bất kỳ model
// nào có ÍT NHẤT 1 điểm dữ liệu > 5% (SECONDARY_AXIS_THRESHOLD) được coi là
// "model biến động lớn" — TOÀN BỘ đường của model đó chuyển hẳn sang đo theo
// TRỤC Y PHỤ bên phải (thang riêng, tự tính theo đúng giá trị thật lớn nhất
// của riêng nhóm này, không còn bị kẹp), trục Y CHÍNH bên trái giữ nguyên
// hành vi cũ (fixedMaxV nếu có, hoặc auto) chỉ tính theo các model còn lại
// (≤5%) — nhờ vậy nhóm model chiếm đa số (0-5%) không còn bị dồn sát đáy.
// Trục phụ chỉ được VẼ RA (đường trục + nhãn %) khi có ít nhất 1 model thuộc
// nhóm này, để không đổi giao diện các trường hợp không có biến động lớn.
const SECONDARY_AXIS_THRESHOLD = 0.05;

function MultiLineChart({
  categories, series, height = 220, theme, maxV: fixedMaxV,
}: {
  categories: string[];
  series: { label: string; color: string; values: (number | null)[] }[];
  height?: number;
  // EPCC (menu5-chart-axis-text-theme-aware) - xem giải thích chi tiết ở
  // ComboBarDualLineChart(); áp dụng cùng cách cho nhãn trục % + tên tháng
  // của Chart 2.
  theme: ThemeMode;
  // EPCC (menu5-chart34-y-axis-fixed-vs-auto) - FIX theo yêu cầu người dùng
  // "Chart 3 (Mass F-Cost%, vị trí thứ 3) trục % 3.5% --> 5.0%, Chart 4 (SPL
  // F-Cost%, vị trí thứ 4) trục % 3.5% --> Auto theo số liệu thực": trước
  // đây `maxV` LUÔN cố định cứng 0.035 (3.5%) cho MỌI chart dùng chung
  // component này (Chart 2/5/6), không thể tách riêng trần trục theo từng
  // chart. Thêm prop `maxV` (đổi tên tham số nhận vào thành `fixedMaxV` để
  // không đụng biến `maxV` tính bên dưới) — nếu người gọi truyền vào thì
  // dùng ĐÚNG giá trị cố định đó (Chart 2 giữ 0.035, Chart 5/Mass đổi sang
  // 0.05); nếu KHÔNG truyền (Chart 6/SPL) thì tự tính auto từ dữ liệu thực
  // qua `computeAutoMaxV()` bên dưới, để điểm đột biến (VD SPL 33.7%) không
  // còn bị kẹp mất ở mép 3.5% như trước.
  // EPCC (menu5-chart234-dual-y-axis-5pct) - `fixedMaxV` giờ CHỈ áp dụng cho
  // trục CHÍNH (model ≤5%); trục PHỤ (model >5%) luôn tự tính auto theo đúng
  // dữ liệu thật của riêng nhóm đó (xem `maxV2` bên dưới), không nhận prop
  // cố định vì mỗi chart có mức đột biến khác nhau, không có 1 trần chung
  // hợp lý.
  maxV?: number;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const width = 620;
  // EPCC (menu5-chart234-dual-y-axis-5pct) - tách series thành 2 nhóm theo
  // ngưỡng SECONDARY_AXIS_THRESHOLD: nhóm `secondary` (có ít nhất 1 điểm
  // >5%) đo theo trục phụ bên phải, nhóm `primary` (còn lại) đo theo trục
  // chính bên trái như hành vi cũ. `hasSecondary` quyết định có vẽ trục phụ
  // (đường trục + nhãn %) hay không — giữ nguyên giao diện cũ khi mọi model
  // đều ≤5%.
  const secondarySeries = series.filter(s => s.values.some(v => v != null && v > SECONDARY_AXIS_THRESHOLD));
  const primarySeries = series.filter(s => !secondarySeries.includes(s));
  const hasSecondary = secondarySeries.length > 0;
  // EPCC (menu5-chart1-2-restore-axis-gutter) - giống Chart 1: chỉ cần
  // gutter TRÁI cho nhãn % (chart này không có trục phải) — padR giữ gần 0
  // vì không có nhãn nào cần chỗ bên phải, đường dữ liệu vẫn kéo gần sát mép
  // phải như yêu cầu "left/right=0" cho phần không cần nhãn.
  // EPCC (menu5-chart1-2-axis-label-clip-fix) - cùng nguyên nhân/cách fix
  // như Chart 1 (xem comment chi tiết ở ComboBarDualLineChart): tăng padL
  // 30→42 để nhãn % dài hơn (VD '11.1%') không bị cắt mất ký tự đầu.
  // EPCC (menu5-chart234-dual-y-axis-5pct) - khi có trục phụ, padR cần đủ
  // rộng (42, bằng padL) để chứa nhãn % của trục phải; khi không có trục phụ
  // giữ nguyên padR=6 như cũ (không đổi giao diện các chart không cần trục
  // phụ).
  const padL = 42, padR = hasSecondary ? 42 : 6, padT = 16, padB = 24;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;
  const n = Math.max(1, categories.length);
  const step = innerW / Math.max(1, n - 1);
  // EPCC (menu5-chart2-y-axis-cap-3-5pct) - FIX theo yêu cầu người dùng "giới
  // hạn cao nhất trục Y (trục %) để 3.5% thôi để nhìn rõ biểu đồ các model":
  // trước đây `maxV` LUÔN lấy giá trị lớn nhất thật sự trong dữ liệu
  // (`Math.max(0.01, ...allVals)`) — khi có 1 model đột biến (VD SO1C2EH
  // ~28%), toàn bộ thang trục kéo giãn theo đỉnh đó, ép mọi model còn lại
  // (đa số ở mức 0-3%) dồn sát đáy, gần như không phân biệt được. Cố định
  // trần trục Y = 3.5% (0.035) theo đúng yêu cầu — điểm nào vượt ngưỡng này
  // (như đỉnh 28% nói trên) sẽ được KẸP (clamp) lên đúng mép trên của khung
  // biểu đồ ở hàm `yOf` bên dưới, thay vì kéo giãn cả thang đo theo nó.
  // EPCC (menu5-chart34-y-axis-fixed-vs-auto) - giờ chỉ CÒN LÀ TRẦN MẶC ĐỊNH
  // khi không có chart nào truyền `maxV` hay dữ liệu rỗng; Chart 2/5 truyền
  // hẳn số cố định riêng, Chart 6 bỏ trống để rơi vào nhánh auto bên dưới.
  // EPCC (menu5-chart234-dual-y-axis-5pct) - trần trục CHÍNH giờ chỉ tính
  // trên `primarySeries` (model ≤5%) — vì nhóm >5% đã tách sang trục phụ,
  // không còn kéo giãn trần trục chính lên nữa. Nếu KHÔNG có model nào ≤5%
  // (hiếm, toàn bộ đều >5%) thì rơi về tính trên `series` gốc để tránh chia
  // cho 0 / trục chính trống vô nghĩa.
  const maxV = fixedMaxV ?? computeAutoMaxV(primarySeries.length ? primarySeries : series);
  // EPCC (menu5-chart234-dual-y-axis-5pct) - trần trục PHỤ luôn tự tính auto
  // theo đúng dữ liệu thật của nhóm >5% (không nhận fixedMaxV, không kẹp).
  const maxV2 = hasSecondary ? computeAutoMaxV(secondarySeries) : 1;
  const xOf = (i: number) => padL + step * i;
  const yOf = (v: number) => Math.max(padT, padT + innerH - (v / maxV) * innerH);
  const yOf2 = (v: number) => Math.max(padT, padT + innerH - (v / maxV2) * innerH);
  // EPCC (menu5-chart234-dual-y-axis-5pct) - chọn đúng hàm quy đổi toạ độ Y
  // (trục chính hay trục phụ) cho từng series, dựa trên việc series đó có
  // nằm trong nhóm `secondarySeries` hay không.
  const yOfForSeries = (s: any) => (secondarySeries.includes(s) ? yOf2 : yOf);

  return (
    // EPCC (menu5-chart1-2-svg-aspect-ratio-gap) - cùng nguyên nhân/cách fix
    // như Chart 1: xem comment chi tiết ở ComboBarDualLineChart().
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} preserveAspectRatio="none" role="img">
      {[0, 0.25, 0.5, 0.75, 1].map((tt, i) => (
        <line key={i} x1={padL} x2={width - padR} y1={padT + innerH * (1 - tt)} y2={padT + innerH * (1 - tt)}
          stroke="var(--border-soft, #2a2f3a)" strokeWidth={1} strokeDasharray="3,4" />
      ))}
      {/* EPCC (menu5-chart1-2-smooth-axis-gap30) - trục Y (%) + trục X, cùng
          nhãn số tại các mốc gridline — trước đây không có đường trục/nhãn số. */}
      <line x1={padL} x2={padL} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={padL} x2={width - padR} y1={padT + innerH} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      {[0, 0.25, 0.5, 0.75, 1].map((tt, i) => (
        <text key={`yl${i}`} x={padL - 6} y={padT + innerH * (1 - tt) + 3} fontSize={10.2} fontWeight={700} textAnchor="end" fill={axisTextColor}>
          {(maxV * tt * 100).toFixed(1)}%
        </text>
      ))}
      {/* EPCC (menu5-chart234-dual-y-axis-5pct) - TRỤC Y PHỤ bên phải, chỉ vẽ
          khi `hasSecondary` (có model >5%). Dùng đường trục riêng (nét đứt để
          phân biệt với trục chính) + nhãn % lấy theo `maxV2` (thang riêng của
          nhóm model biến động lớn), đặt sát mép phải khung vẽ
          (`width - padR`). */}
      {hasSecondary && (
        <>
          <line x1={width - padR} x2={width - padR} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} strokeDasharray="2,2" />
          {[0, 0.25, 0.5, 0.75, 1].map((tt, i) => (
            <text key={`yl2-${i}`} x={width - padR + 6} y={padT + innerH * (1 - tt) + 3} fontSize={10.2} fontWeight={700} textAnchor="start" fill={axisTextColor}>
              {(maxV2 * tt * 100).toFixed(1)}%
            </text>
          ))}
        </>
      )}
      {series.map((s, si) => (
        <path key={si} d={buildSmoothLinePath(s.values, xOf, yOfForSeries(s))} fill="none" stroke={s.color} strokeWidth={1.8}
          strokeLinejoin="round" strokeLinecap="round" strokeDasharray={secondarySeries.includes(s) ? '6,3' : undefined} />
      ))}
      {series.map((s, si) => s.values.map((v, i) => v == null ? null : (
        <circle key={`${si}-${i}`} cx={xOf(i)} cy={yOfForSeries(s)(v)} r={2.5} fill={s.color} />
      )))}
      {/* EPCC (menu5-chart2-point-value-labels) - FIX theo yêu cầu người
          dùng "cho hiển thị số liệu của các biểu đồ đường này": trước đây
          Chart 2 chỉ có chấm tròn tại mỗi điểm, không có con số % đi kèm
          (khác Chart 1 đã có nhãn Target/Actual) — người dùng phải tự đoán
          giá trị qua vị trí trên trục Y. Thêm nhãn % tại mỗi điểm có dữ
          liệu, màu trùng màu đường (`s.color`) để phân biệt giữa nhiều
          model, đặt phía TRÊN chấm tròn (y-6); cỡ chữ nhỏ hơn Chart 1 (8.5
          thay vì 11.7) vì Chart 2 có tới 8 đường cùng lúc, cần tiết kiệm
          không gian để tránh chồng chéo quá mức. */}
      {series.map((s, si) => s.values.map((v, i) => v == null ? null : (
        <text key={`vl${si}-${i}`} x={xOf(i)} y={yOfForSeries(s)(v) - 6} fontSize={8.5} fontWeight={700}
          textAnchor="middle" fill={s.color}>
          {(v * 100).toFixed(2)}%
        </text>
      )))}
      {/* EPCC (menu5-chart1-2-zero-lr-padding) - với padL/padR=0, điểm đầu/
          cuối nằm ĐÚNG tại x=0 và x=width; nhãn tháng textAnchor="middle" cũ
          sẽ bị cắt một nửa ở 2 mép. Đổi riêng nhãn đầu/cuối sang anchor
          start/end (các nhãn giữa giữ nguyên middle). */}
      {categories.map((c, i) => (
        <text key={i} x={xOf(i)} y={height - 6} fontSize={11.4} fontWeight={700}
          textAnchor={i === 0 ? 'start' : i === categories.length - 1 ? 'end' : 'middle'}
          fill={axisTextColor}>
          {c}
        </text>
      ))}
    </svg>
  );
}

// ── Chart 7: biểu đồ NHỆN (radar/spider) — NHÂN BẢN Y HỆT DỮ LIỆU của Chart 2
// (categories = TTL + 12 tháng, series = F-Cost% từng model), chỉ khác kiểu vẽ:
// thay vì trục X ngang + nhiều đường, mỗi "tháng" (category) trở thành 1 trục
// tia toả ra từ tâm, mỗi model là 1 đa giác khép kín nối các điểm trên các
// trục đó — đúng dạng "spider chart" trong ảnh tham chiếu người dùng mô tả,
// dùng để so sánh hình dạng biến động F-Cost% giữa nhiều model cùng lúc.
function RadarChart({
  categories, series, size = 460, theme, maxV: fixedMaxV, rings = 4,
}: {
  categories: string[];
  series: { label: string; color: string; values: (number | null)[] }[];
  size?: number;
  theme: ThemeMode;
  maxV?: number;
  rings?: number;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  // EPCC (menu5-radar-lightmode-grid-visible) - FIX ROOT CAUSE "đường tròn
  // nhện gần như biến mất ở light mode": lưới cũ dùng thẳng biến CSS
  // `var(--border-soft, #2a2f3a)` — biến này được định nghĩa để hợp với nền
  // TỐI của các dashboard khác trong app, nhưng card F-Cost (ChartCard) có
  // nền TRẮNG ở light mode, khiến `--border-soft` (vốn đã nhạt) gần như
  // trong suốt trên nền trắng. Tách riêng màu lưới cho radar theo `theme`
  // (không phụ thuộc `--border-soft`), cùng quy ước "đen mờ ở light / trắng
  // mờ ở dark" như `axisTextColor` phía trên — đảm bảo đổi theme vẫn đúng
  // màu ở cả 2 chế độ.
  const gridStroke = theme === 'light' ? 'rgba(15,23,42,0.28)' : 'rgba(226,232,240,0.22)';
  const gridStrokeOuter = theme === 'light' ? 'rgba(15,23,42,0.42)' : 'rgba(226,232,240,0.34)';
  const axisLineColor = theme === 'light' ? 'rgba(15,23,42,0.32)' : 'rgba(226,232,240,0.26)';
  // EPCC (menu5-radar-inner-wash) - THÊM MỚI lớp phủ mờ phía trong (như ảnh
  // tham chiếu RTY Total) — ĐÂY LÀ NỀN TRANG TRÍ TĨNH, không tính từ dữ liệu
  // (khác RTY Total, nơi hiệu ứng "nền nhạt" tự nhiên sinh ra từ nhiều
  // series dữ liệu cao/đều chồng lớp — xem giải thích trước đó). Dùng
  // radial gradient nhạt dần ra biên, theo đúng yêu cầu người dùng "có thể
  // làm phủ mờ nền phía trong", màu trung tính đổi theo theme.
  const washColor = theme === 'light' ? 'rgba(99,102,241,0.07)' : 'rgba(129,140,248,0.11)';
  const gradientIdRef = useRef(`radarWash-${Math.random().toString(36).slice(2)}`);
  const gradientId = gradientIdRef.current;
  const cx = size / 2, cy = size / 2;
  // padding đủ rộng để chứa nhãn tháng (category) + nhãn % ở các vòng lưới,
  // cùng nguyên tắc padL/padR của MultiLineChart (chừa chỗ cho text, không để
  // đè lên khung vẽ).
  const padding = size * 0.2;
  const radius = size / 2 - padding;
  const n = Math.max(1, categories.length);
  const angleStep = (2 * Math.PI) / n;
  const angleOf = (i: number) => -Math.PI / 2 + angleStep * i; // bắt đầu từ đỉnh (12h), giống categories[0]=TTL ở trên cùng

  // EPCC: cùng cách tính maxV như MultiLineChart (fixedMaxV nếu có, không thì
  // auto theo dữ liệu thực +10% đệm, làm tròn lên mốc 0.5%) — để Chart 7 "y
  // hệt dữ liệu" Chart 2 cũng đọc đúng thang % tương đương.
  const allVals = series.flatMap(s => s.values.filter((v): v is number => v != null));
  const dataMax = allVals.length ? Math.max(...allVals) : 0.01;
  const autoMaxV = Math.max(0.005, Math.ceil((dataMax * 1.1) / 0.005) * 0.005);
  const maxV = fixedMaxV ?? autoMaxV;

  const radiusOf = (v: number) => Math.max(0, Math.min(1, v / maxV)) * radius;
  const pointAt = (i: number, r: number): [number, number] => {
    const a = angleOf(i);
    return [cx + r * Math.cos(a), cy + r * Math.sin(a)];
  };

  // EPCC (menu5-radar-circular-grid) - FIX theo yêu cầu người dùng "phân
  // tích tại sao vẫn không giống hình 2 có nền màu nhạt bên trong và có
  // vòng tròn bên ngoài": nguyên nhân gốc là lưới cũ vẽ dạng ĐA GIÁC khớp
  // đúng số trục (7 cạnh cho 7 trục TTL+6 tháng), trong khi Plotly
  // `radialaxis` (dùng ở RTY Total, plotSpiderPanel()) MẶC ĐỊNH vẽ lưới là
  // VÒNG TRÒN ĐỒNG TÂM THẬT SỰ bất kể số trục — khiến 2 chart nhìn khác hẳn
  // dù cùng là "radar". Đổi ringLevels sang vẽ bằng <circle> (đường tròn
  // hoàn chỉnh) thay vì <polygon>, đúng hành vi mặc định của Plotly.
  const ringLevels = Array.from({ length: rings }, (_, ri) => (ri + 1) / rings);

  return (
    <svg viewBox={`0 0 ${size} ${size}`} width="100%" height={size} role="img">
      <defs>
        <radialGradient id={gradientId} cx="50%" cy="50%" r="65%">
          <stop offset="0%" stopColor={washColor} stopOpacity={1} />
          <stop offset="100%" stopColor={washColor} stopOpacity={0} />
        </radialGradient>
      </defs>
      {/* Lớp phủ mờ nền phía trong — vẽ TRƯỚC lưới/trục/dữ liệu (nằm dưới cùng) */}
      <circle cx={cx} cy={cy} r={radius} fill={`url(#${gradientId})`} />
      {/* Lưới vòng tròn đồng tâm (không còn là đa giác) — các vòng phụ bên
          trong vẫn nét đứt mảnh như cũ. */}
      {ringLevels.map((lv, ri) => (
        <circle key={ri} cx={cx} cy={cy} r={radius * lv} fill="none"
          stroke={gridStroke} strokeWidth={1} strokeDasharray="3,4" />
      ))}
      {/* Vòng ngoài cùng — nét liền đậm hơn, làm viền khung radar (đúng
          "vòng tròn bên ngoài" trong ảnh tham chiếu RTY Total), tách biệt
          rõ với các vòng lưới phụ nét đứt bên trong. */}
      <circle cx={cx} cy={cy} r={radius} fill="none" stroke={gridStrokeOuter} strokeWidth={1.4} />
      {/* Trục tia từ tâm ra từng category */}
      {categories.map((_, i) => {
        const [x2, y2] = pointAt(i, radius);
        return <line key={i} x1={cx} y1={cy} x2={x2} y2={y2} stroke={axisLineColor} strokeWidth={1} />;
      })}
      {/* Nhãn % dọc trục đầu tiên (giống cột nhãn trục Y bên trái của Chart 2) */}
      {ringLevels.map((lv, ri) => {
        const [x, y] = pointAt(0, radius * lv);
        return (
          <text key={`rl${ri}`} x={x + 6} y={y - 4} fontSize={9.6} fontWeight={700} textAnchor="start" fill={axisTextColor}>
            {(maxV * lv * 100).toFixed(1)}%
          </text>
        );
      })}
      {/* Đa giác dữ liệu từng model — giá trị thiếu (null, model chưa có số
          liệu tháng đó) coi như 0 để đa giác vẫn khép kín, đồng dạng cách
          RoseChart xử lý model không có dữ liệu ở top-N kỳ hiện tại.
          EPCC (menu5-radar-unify-rtytotal-style) - FIX theo yêu cầu người
          dùng "lấy cách vẽ biểu đồ nhện từ RTY Total để làm giống": đồng bộ
          fillOpacity 0.08→0.1 và strokeWidth 1.8→2, đúng thông số
          plotSpiderPanel() trong RtyTotalTab.tsx (fillcolor `${color}1A`
          ≈ alpha 0.10, line.width: 2). */}
      {series.map((s, si) => {
        const pts = s.values.map((v, i) => pointAt(i, radiusOf(v ?? 0)).join(',')).join(' ');
        return (
          <polygon key={si} points={pts} fill={s.color} fillOpacity={0.1} stroke={s.color} strokeWidth={2} strokeLinejoin="round" />
        );
      })}
      {/* Chấm tròn tại từng điểm có dữ liệu thật (bỏ qua null) — marker.size
          đồng bộ RTY Total (r=3 ~ tương đương marker size 5-6px của Plotly). */}
      {series.map((s, si) => s.values.map((v, i) => {
        if (v == null) return null;
        const [x, y] = pointAt(i, radiusOf(v));
        return <circle key={`${si}-${i}`} cx={x} cy={y} r={3} fill={s.color} stroke="var(--surface, #0f1420)" strokeWidth={0.8} />;
      }))}
      {/* EPCC (menu5-radar-unify-rtytotal-style) - THÊM MỚI nhãn số liệu ngay
          tại từng điểm dữ liệu, đúng "cách vẽ" plotSpiderPanel() trong
          RtyTotalTab.tsx (Plotly scatterpolar dùng
          `mode: 'lines+markers+text', textposition: 'top center'`, màu chữ
          = màu series). SVG không có textposition tự động nên mô phỏng bằng
          cách đặt nhãn lùi ra xa tâm thêm 1 khoảng cố định (labelGap) dọc
          đúng trục tia của điểm đó — cùng hướng "hướng ra ngoài" như Plotly
          "top center" trên biểu đồ cực. Chỉ vẽ khi có dữ liệu thật (v != null),
          giữ đúng quy tắc "0/thiếu dữ liệu không vẽ" đã áp dụng cho polygon/marker. */}
      {series.map((s, si) => s.values.map((v, i) => {
        if (v == null) return null;
        const labelGap = size * 0.032;
        const [lx, ly] = pointAt(i, radiusOf(v) + labelGap);
        return (
          <text key={`lbl-${si}-${i}`} x={lx} y={ly} fontSize={9.5} fontWeight={700}
            textAnchor="middle" dominantBaseline="middle" fill={s.color}
            style={{ paintOrder: 'stroke', stroke: 'var(--surface, #0f1420)', strokeWidth: 2.5, strokeLinejoin: 'round' }}>
            {(v * 100).toFixed(1)}
          </text>
        );
      }))}
      {/* Nhãn category (TTL/tháng) tại đầu mút mỗi trục */}
      {categories.map((c, i) => {
        const [x, y] = pointAt(i, radius + size * 0.06);
        return (
          <text key={`cl${i}`} x={x} y={y} fontSize={10.8} fontWeight={700} textAnchor="middle" dominantBaseline="middle" fill={axisTextColor}>
            {c}
          </text>
        );
      })}
    </svg>
  );
}

function LegendDot({ color, label, theme }: { color: string; label: string; theme: ThemeMode }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
      <span style={{ width: 10, height: 10, borderRadius: 2, background: color, flexShrink: 0 }} />
      {/* EPCC (menu5-legend-text-theme-aware) - FIX theo yêu cầu người dùng
          "vẫn còn bỏ sót 2 mục khoanh đỏ này chưa đổi màu khi đổi theme":
          trước đây (menu5-circled-text-bold-black-20pct) đã cố định màu chữ
          nhãn legend (F-Cost($)/F-Cost(억원)/TARGET/ACTUAL ở Chart 1, tên
          model ở Chart 2) sang đen — lúc đó chỉ tính cho theme light, quên
          mất theme dark khiến chữ đen gần như biến mất trên nền tối. Đổi
          sang tính theo `theme`, cùng quy ước với tiêu đề card
          (menu5-header-title-color-theme-aware): đen ở light, trắng ở dark. */}
      <span style={{ fontSize: 12.6, fontWeight: 700, color: theme === 'light' ? '#000000' : '#ffffff' }}>{label}</span>
    </div>
  );
}

// ── Chart 3 & 4: Rose / Nightingale chart — vòng ngoài xám ghi nhãn model,
// cánh hoa (radial bar) từ tâm dài theo giá trị. `value: null` => vẽ nhãn
// "#N/A" xám mờ, không có cánh hoa (dùng cho Chart 4: model không thuộc
// top-N được chọn cho kỳ hiện tại). ───────────────────────────────────────
// EPCC (menu5-rose-chart-taller-v3) - FIX theo yêu cầu người dùng "tăng chiều
// cao như hình tham chiếu để dễ nhìn và rõ chữ hơn, áp dụng cho cả 2 biểu đồ
// 3,4": size mặc định trước đó (320) khiến khối SVG (vốn hình vuông, cao =
// size) bị giới hạn chiều cao thấp hơn hẳn bề rộng cột lưới 2 cột, làm cánh
// hoa + chip nhãn bị co nhỏ theo tỉ lệ "meet". Tăng size mặc định 320→420 để
// khối SVG cao hơn hẳn (chiều cao là kích thước bị giới hạn thực tế), đồng
// thời đổi fontSize từ hằng số cố định (8.6) sang SCALE THEO size (size *
// 0.027) — để khi tăng size, chữ trong chip nhãn cũng to rõ theo, không còn
// bị "nhỏ tương đối" so với tổng thể biểu đồ như trước.
function RoseChart({
  items, barColor, size = 420, rotateLabels90 = false,
}: {
  items: { label: string; value: number | null }[];
  barColor: string;
  size?: number;
  // EPCC (menu5-rose-chart3-labels-vertical) - FIX theo yêu cầu người dùng
  // "để các chữ xoay 90 độ so với hiện tại, hướng quay là từ phải hướng
  // lên, áp dụng cho Chart 3 thôi": trước đây mọi nhãn (cả Chart 3 & 4)
  // dùng chung 1 công thức `rot` bám theo tiếp tuyến vòng tròn (chữ đọc
  // xuôi theo đường cong). Thêm prop `rotateLabels90` (mặc định false, chỉ
  // Chart 3 truyền true) — khi bật, trừ thêm 90° vào góc xoay hiện có. Trừ
  // (không phải cộng) vì SVG `rotate()` dương là xoay CÙNG chiều kim đồng
  // hồ; xoay ngược chiều kim đồng hồ (trừ 90°) đưa cạnh PHẢI của chữ hướng
  // LÊN trên — đúng chiều "từ phải hướng lên" người dùng mô tả. Chart 4
  // không truyền prop này nên giữ nguyên hành vi cũ.
  rotateLabels90?: boolean;
}) {
  const cx = size / 2, cy = size / 2;
  // EPCC (menu5-rose-chart-fill-in-v2) - FIX ROOT CAUSE "cánh hoa vẫn quá nhỏ
  // so với vòng nhãn dù đã giảm ringThickness lần trước (0.075→0.05)": lần
  // sửa trước mới giảm NHẸ ringThickness, chưa đủ mạnh — với dữ liệu model
  // top chênh lệch lớn (VD 77 vs 9), dù đã dùng scale căn bậc hai, model nhỏ
  // nhất vẫn chỉ đạt sqrt(9/77)≈34% bán kính, nhìn vẫn "lọt thỏm". Tăng mạnh
  // hơn nữa lần này: (1) giảm tiếp ringThickness 0.05→0.032 (vòng mỏng hơn,
  // dành thêm bán kính cho cánh hoa); (2) thêm SÀN TỐI THIỂU 18% bán kính cho
  // mọi cánh có dữ liệu (kể cả giá trị rất nhỏ so với max) rồi mới scale phần
  // còn lại theo căn bậc hai — đảm bảo cánh nào cũng "vươn" ra rõ ràng, không
  // còn cánh nào chỉ là 1 vệt mờ gần tâm, đúng dáng đầy đặn trong hình vẽ tay.
  // EPCC (menu5-rose-ring-label-inside-v4) - FIX theo yêu cầu người dùng
  // "cho label ring to hơn để tên model nằm luôn trong đó như ảnh tham
  // chiếu, và cho màu nền cho label ring, áp dụng cả chart 3 và 4": trước
  // đây `ringThickness` chỉ 0.032*size (gần như 1 sợi chỉ, opacity 0.5, màu
  // tối) — không đủ chỗ chứa chữ, nên chữ phải đẩy ra thành 1 CHIP NỔI riêng
  // ngoài `outerR`. Đổi cách tiếp cận: tăng `ringThickness` lên 0.13*size để
  // dải ring đủ rộng chứa cả dòng "MODEL, giá trị" bên trong, đồng thời tăng
  // `outerR` lên 0.47*size (an toàn vì không còn chip nhô ra ngoài outerR
  // nữa) để tận dụng hết khung vuông. Vì đổi ở component `RoseChart` dùng
  // chung, chart 3 (ALL YEAR) và chart 4 (MONTH) tự động nhận đồng thời.
  const outerR = size * 0.47;      // bán kính ngoài cùng của vòng nhãn (ring)
  const ringThickness = size * 0.13; // dày hơn hẳn để chữ nằm gọn bên trong ring
  const innerR = outerR - ringThickness; // bán kính tối đa của cánh hoa
  const FLOOR_RATIO = 0.18; // sàn tối thiểu cho cánh có dữ liệu (kể cả giá trị rất nhỏ)
  const n = Math.max(1, items.length);
  const maxV = Math.max(1, ...items.map(it => it.value ?? 0));
  const gapDeg = 1.4;
  const stepDeg = 360 / n;
  const fontSize = size * 0.027; // scale theo size (320px cũ ≈ 8.64, khớp giá trị cũ)

  const toRad = (deg: number) => (deg - 90) * (Math.PI / 180);
  const polar = (r: number, deg: number) => ({ x: cx + r * Math.cos(toRad(deg)), y: cy + r * Math.sin(toRad(deg)) });

  function arcPath(r0: number, r1: number, a0: number, a1: number) {
    const p0 = polar(r1, a0), p1 = polar(r1, a1), p2 = polar(r0, a1), p3 = polar(r0, a0);
    const large = a1 - a0 > 180 ? 1 : 0;
    return `M ${p0.x} ${p0.y} A ${r1} ${r1} 0 ${large} 1 ${p1.x} ${p1.y} L ${p2.x} ${p2.y} A ${r0} ${r0} 0 ${large} 0 ${p3.x} ${p3.y} Z`;
  }

  return (
    <svg viewBox={`0 0 ${size} ${size}`} width="100%" height={size} role="img">
      {/* EPCC (menu5-rose-white-hub-fixed-theme) - FIX theo yêu cầu người
          dùng "phần ở giữa labelring và biểu đồ bên trong có khoảng trống
          luôn để mặc định là màu trắng dù đổi theme": trước đây vùng giữa
          cánh hoa (bán kính thay đổi theo giá trị, luôn <= innerR) và vòng
          nhãn (bắt đầu từ innerR) không có gì vẽ ở đó — SVG để trong suốt,
          nên lộ ra màu nền card `var(--surface)` vốn đổi theo theme (tối/
          sáng). Vẽ 1 hình tròn nền màu TRẮNG CỐ ĐỊNH (không dùng biến CSS
          theme) bán kính = innerR, vẽ TRƯỚC cánh hoa (nằm dưới), để khoảng
          trống này luôn trắng bất kể theme. */}
      {/* EPCC (menu5-rose-hub-color-ref-img1) - FIX theo yêu cầu người dùng
          "cho đường kính của đường vòng tròn này màu như hình 1 tham chiếu":
          đổi nền vòng tròn giữa (hub) từ trắng cố định sang màu đo từ ảnh
          swatch người dùng cung cấp (#BBC6D9), thay cho #ffffff cũ (xem
          menu5-rose-white-hub-fixed-theme phía trên — quyết định "luôn
          trắng" đó nay được người dùng chủ động ghi đè bằng màu mới). */}
      {/* EPCC (menu5-rose-gap-background-ref-img) - FIX theo yêu cầu người
          dùng "cho gapdeg có nền Background là màu như ảnh": trước đây các
          mảnh labelring chỉ vẽ riêng lẻ từng path theo (a0,a1), khe hở
          `gapDeg` giữa 2 mảnh liền kề để trống → lộ màu nền card tối phía
          sau (xem đoạn giải thích "các đường khoanh đỏ" ở trên). Vẽ 1 vòng
          tròn ĐẦY (không khe hở) bán kính = outerR, TRƯỚC hub và TRƯỚC các
          mảnh labelring (nằm dưới cùng), màu đo từ ảnh swatch người dùng
          cung cấp (#A0C9C3) — nhờ đó khe hở gapDeg trong dải labelring giờ
          lộ ra đúng màu này thay vì màu nền tối. Hub (#BBC6D9) vẽ NGAY SAU,
          đè lên phần giữa (0→innerR) như cũ, không ảnh hưởng khe hở cánh
          hoa (vẫn lộ màu hub BBC6D9 như trước, không phải A0C9C3). */}
      <circle cx={cx} cy={cy} r={outerR} fill="#A0C9C3" />
      <circle cx={cx} cy={cy} r={innerR} fill="#BBC6D9" />
      {items.map((it, i) => {
        const a0 = i * stepDeg + gapDeg / 2, a1 = (i + 1) * stepDeg - gapDeg / 2;
        const mid = (a0 + a1) / 2;
        const hasVal = it.value != null;
        const ratio = hasVal ? Math.min(1, (it.value as number) / maxV) : 0;
        const r1 = hasVal ? innerR * (FLOOR_RATIO + (1 - FLOOR_RATIO) * Math.sqrt(ratio)) : innerR * 0.03;
        const isBottomHalf = mid > 90 && mid < 270;
        // EPCC (menu5-rose-chart3-labels-anchor-fix) - FIX ROOT CAUSE "chữ
        // vẫn tràn ra ngoài viền labelring ở đúng nhãn trên cùng (mid≈0)":
        // lần sửa trước (menu5-rose-chart3-labels-inward) chỉ dịch tâm neo
        // `labelR` vào gần `innerR` hơn nhưng vẫn dùng `textAnchor="middle"`
        // — với neo giữa, chữ luôn tràn ĐỀU cả 2 hướng dọc theo trục chữ.
        // Ở nửa "trên" vòng tròn (mid ∈ [-90,90], tức không rơi vào
        // `isBottomHalf`), trục chữ (sau khi xoay `rot`) trùng hướng LY TÂM
        // (ra ngoài) — nghĩa là "tên model" (đầu chuỗi) nằm gần tâm neo,
        // còn "con số" (cuối chuỗi, sau dấu phẩy) bị đẩy ra phía ngoài
        // cùng, đúng phần bị lộ ra ngoài labelring trong ảnh chụp. Ở nửa
        // "dưới" (`isBottomHalf`), trục chữ lại trùng hướng HƯỚNG TÂM (vào
        // trong) nên không có vấn đề (đầu chuỗi ra ngoài, con số vào trong).
        // ROOT FIX theo đúng yêu cầu người dùng "sắp xếp tên model trước
        // rồi con số... để không bị lộ số ra ngoài": GIỮ NGUYÊN thứ tự
        // chuỗi "MODEL, số" (không đảo chữ), chỉ đổi CÁCH NEO (`textAnchor`)
        // theo từng nửa để điểm neo luôn nằm SÁT MÉP TRONG AN TOÀN của
        // outerR, và toàn bộ chữ (kể cả con số) mọc dần VÀO TRONG từ đó:
        // - Nửa trên (trục chữ ra ngoài): neo tại `outerR - margin`,
        //   `textAnchor="end"` → điểm neo là KÝ TỰ CUỐI (số), chữ mọc
        //   ngược vào trong (đúng hướng "ngược kim đồng hồ" theo trục chữ)
        //   → toàn bộ chữ, kể cả số, không còn vượt quá outerR.
        // - Nửa dưới (trục chữ vào trong): neo cũng tại `outerR - margin`,
        //   `textAnchor="start"` → điểm neo là KÝ TỰ ĐẦU (model), chữ mọc
        //   tiếp vào trong theo đúng chiều trục sẵn có → cùng nguyên lý,
        //   không đổi hành vi hiển thị so với trước (vốn đã an toàn).
        // Chart 4 (rotateLabels90=false) không đụng tới, vẫn neo giữa dải
        // ring như nguyên bản.
        const ringSafeMargin = size * 0.012;
        const labelR = rotateLabels90 ? outerR - ringSafeMargin : innerR + ringThickness / 2;
        const lp = polar(labelR, mid);
        const textAnchorAttr: 'start' | 'middle' | 'end' = rotateLabels90
          ? (isBottomHalf ? 'start' : 'end')
          : 'middle';
        // Xoay nhãn cho luôn đọc xuôi: nửa dưới vòng tròn thì lật 180°.
        const baseRot = isBottomHalf ? mid + 180 : mid;
        const rot = rotateLabels90 ? baseRot - 90 : baseRot;
        const labelText = `${it.label}${hasVal ? `, ${fmtNum(it.value as number, 0)}` : ', #N/A'}`;
        return (
          <g key={i}>
            <path d={arcPath(0, Math.max(r1, 2), a0, a1)} fill={hasVal ? barColor : 'var(--border-soft, #384152)'} opacity={hasVal ? 0.85 : 0.35} />
            {/* EPCC (menu5-rose-ring-color-noborder) - FIX theo yêu cầu người
                dùng "thay màu ảnh 1 cho labelring, đường phía trong bỏ đi":
                (1) đổi nền ring sang màu vàng nhạt đo từ ảnh swatch mới
                (#FEFDA1), thay hồng pastel cũ (#F8DAD0); (2) bỏ hẳn `stroke`
                — đây chính là đường viền quanh mỗi mảnh ring (gồm cả cạnh
                cung TRONG giáp hub trắng, bị khoanh đỏ trong ảnh mẫu). Việc
                bỏ stroke KHÔNG làm mất khoảng cách giữa các mảnh ring, vì
                khoảng cách đó vốn đến từ chính hình học path (gapDeg chèn
                giữa a0/a1 của mỗi mảnh), không phụ thuộc vào stroke. Áp dụng
                chung cho cả Chart 3 & 4 vì cùng 1 component RoseChart. */}
            {/* EPCC (menu5-rose-ring-no-dark-seam) - FIX ROOT CAUSE "đường
                label ring bo ngoài bị vòng tròn đen khi để dark mode": ring
                trước đây vẽ từ `innerR + 1` (không phải `innerR`), cố ý chừa
                1px hở giữa hub trắng và ring vàng. Ở dark mode, 1px hở này
                lộ thẳng nền `--surface` (tối) phía sau SVG ra ngoài, tạo
                thành 1 đường viền đen mảnh chạy vòng đúng ranh giới hub/ring
                — ở light mode gần như vô hình vì nền sáng nên trước đây
                không phát hiện ra. Bỏ khoảng hở (`innerR` thay vì `innerR +
                1`) để ring vẽ khít liền vào hub, không còn khe hở lộ nền ra
                ngoài ở bất kỳ theme nào. */}
            {/* EPCC (menu5-rose-labelring-color-ref-img2) - FIX theo yêu cầu
                người dùng "Labelring Background chuyển thành màu như hình 2
                tham chiếu": đổi nền labelring từ vàng nhạt #FEFDA1 sang màu
                đo từ ảnh swatch mới người dùng cung cấp (#E2B4B7). */}
            <path d={arcPath(innerR, outerR, a0, a1)} fill="#E2B4B7" opacity={0.95} />
            <g transform={`rotate(${rot}, ${lp.x}, ${lp.y})`}>
              <text x={lp.x} y={lp.y} fontSize={fontSize} fontWeight={700} textAnchor={textAnchorAttr} dominantBaseline="middle" fill="#1c2b3a">
                {labelText}
              </text>
            </g>
          </g>
        );
      })}
    </svg>
  );
}

// ═══════════════════════════════════════════════════════════════════════
// EPCC (menu5-revenue-tab) - 4 component biểu đồ MỚI cho tab "TÌNH HÌNH
// DOANH THU", nhân bản kiểu dáng/màu sắc của SalesDashboard.tsx (dùng
// Plotly) nhưng vẽ bằng SVG thuần theo đúng "ngôn ngữ" chart sẵn có của
// Mục 5 (ComboBarDualLineChart/MultiLineChart/RoseChart ở trên) — tránh
// phụ thuộc thêm `window.Plotly` (không được Mục 5 tải sẵn) trong khi vẫn
// giữ ĐÚNG bảng màu gốc từ SalesDashboard: bar #2d7f96, line Sales #00a65a,
// line YoY #f39c12 (nét đứt), top-3 model #0891b2 (chữ #e11d48), rest
// #d97706, khách hàng lấy màu theo CHART_PALETTE_9.
// ═══════════════════════════════════════════════════════════════════════

// ── Chart 1: Cột SHIPMENT + đường SALES (cùng trục trái) + đường đứt YoY% (trục phải) ──
// EPCC (menu5-revchart1-add-prodn) - thêm prop `prodVals` (cột SẢN XUẤT,
// đứng cạnh cột XUẤT HÀNG, cùng trục trái) theo đúng ảnh tham chiếu người
// dùng cung cấp — 2 cột giờ vẽ dạng "grouped bar" (chia đôi bề rộng `barW`
// cũ cho mỗi cột) thay vì 1 cột full-width như trước.
// EPCC (menu5-revchart1-link-model-year-viewby) - FIX theo yêu cầu người
// dùng "liên kết thanh lựa chọn model/năm sẽ thay đổi theo đúng giá trị
// của model và năm đó... khi thay đổi năm thì trục dưới sẽ thành tháng,
// tức có thể thay đổi theo cả tính năng lựa chọn tháng/quý/năm nữa": trước
// đây prop `years: number[]` CHỈ chấp nhận nhãn dạng năm (2017..2026) —
// không thể biểu diễn nhãn "Q1..Q4" hay "JAN..DEC" khi bộ lọc XEM THEO đổi
// sang Quý/Tháng (giống ĐÚNG cơ chế `chart1`/`buildFcostByModelChart` ở
// tab F-COST đã dùng cho Chart 1/2 của tab đó). Đổi tên + kiểu prop thành
// `labels: (string | number)[]` để nhận được cả 3 dạng nhãn năm/quý/tháng
// từ `revenueChart1` (xem useMemo tương ứng phía dưới, giờ build theo
// đúng 3 nhánh viewBy y hệt `chart1`).
function RevenueTrendChart({
  labels, shipVals, prodVals, salesVals, yoyVals, theme,
}: {
  labels: (string | number)[];
  shipVals: (number | null)[];
  prodVals: (number | null)[];
  salesVals: (number | null)[];
  yoyVals: (number | null)[];
  theme: ThemeMode;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const width = 620, height = 260;
  // EPCC (menu5-revchart1-labels-above-bars-gap15) - FIX theo yêu cầu người
  // dùng "số hiện ra ở cột xuất hàng và sản xuất hãy chuyển lên Above": nhãn
  // số của 2 cột SHIPMENT/PROD'N trước đây nằm ĐÈ BÊN TRONG cột (xoay -90°,
  // neo gần đáy cột) — nay chuyển ra NGOÀI, phía TRÊN đỉnh mỗi cột (giống
  // cách đặt nhãn của đường DOANH SỐ/YoY). Vì nhãn giờ nằm phía trên đỉnh
  // cột cao nhất (gần mép trên chart), tăng `padT` từ 26 → 34 để có đủ chỗ,
  // tránh bị cắt chữ khi cột chạm gần đỉnh trục.
  const padL = 46, padR = 46, padT = 34, padB = 26;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;
  const n = Math.max(1, labels.length);
  const step = innerW / n;
  // EPCC (menu5-revchart1-labels-above-bars-gap15) - FIX theo yêu cầu người
  // dùng "Sửa Gap width là 15%": trước đó 2 cột SHIPMENT+PROD'N gộp lại
  // (`groupW`) chiếm 80% bề rộng mỗi năm (`step`), để hở (gap) 20% — nay đổi
  // thành chiếm 85% (`step * 0.85`), phần hở giữa 2 nhóm cột của 2 năm liền
  // kề còn đúng 15% như yêu cầu mới.
  const groupW = step * 0.85;
  const barW = groupW / 2;
  const maxBar = Math.max(1, ...shipVals.map(v => v ?? 0), ...prodVals.map(v => v ?? 0), ...salesVals.map(v => v ?? 0));
  // EPCC (menu5-revchart1-labels-above-bars-gap15) - FIX theo yêu cầu người
  // dùng "số hiện ra ở cột xuất hàng và sản xuất hãy chuyển lên Above": nhãn
  // giờ nằm NGOÀI, phía TRÊN đỉnh cột (trên nền trang/chart), KHÔNG còn đè
  // lên nền cột màu đậm như trước — nên `#ffffff` cứng cũ không còn phù hợp
  // (chữ trắng trên nền sáng ở theme light sẽ gần như vô hình). Đổi sang
  // dùng đúng màu của mỗi cột (`#2d7f96` cho SHIPMENT, `#F17988` cho
  // PROD'N — xem `shipLabelColor`/`prodLabelColor` khai báo riêng ngay dưới)
  // để chữ vẫn liên kết trực quan với đúng cột của nó, đọc rõ ở cả 2 theme.
  const shipLabelColor = '#2d7f96';
  // EPCC (menu5-revchart1-prod-label-inside-end-rotate) - FIX theo yêu cầu
  // người dùng "Data Label của sản xuất... khi chuyển theme màu sắc dễ nhìn
  // nhận diện được": nhãn số cột PROD'N giờ nằm ĐÈ BÊN TRONG cột hồng
  // (#F17988) — màu hồng cố định cũ (`#F17988`, trùng màu nền cột) sẽ gần
  // như hòa vào nền cột (hồng trên hồng), không đọc được ở bất kỳ theme
  // nào. Đổi sang cặp màu tương phản mạnh với nền cột hồng, đồng thời khác
  // nhau theo `theme` để luôn nổi rõ: theme light dùng màu tối gần đen/nâu
  // đậm (`#3a0d12`), theme dark dùng trắng (`#ffffff`).
  const prodInsideLabelColor = theme === 'light' ? '#3a0d12' : '#ffffff';
  // EPCC (menu5-revchart1-yoy-label-theme) - FIX theo yêu cầu người dùng
  // "Đường màu cam khó nhìn số nên màu chữ cần thay đổi theo Theme": nhãn %
  // của đường Tr.trưởng YoY trước đó dùng `fill="#e67e22"` CỐ ĐỊNH — cùng
  // tông cam với chính đường/chấm tròn nên khó phân biệt số với nền ở cả 2
  // theme. Nay đổi màu chữ theo `theme`, tương phản rõ với nền trang (khác
  // hẳn `barLabelColor` ở trên vốn đè lên nền CỘT màu đậm cố định, không
  // phải nền trang nên không cần đổi theo theme): theme sáng dùng cam sậm
  // gần nâu (`#a34d05`) cho dễ đọc trên nền trắng, theme tối dùng cam nhạt
  // sáng (`#ffcc80`) cho dễ đọc trên nền tối.
  const yoyLabelColor = theme === 'light' ? '#a34d05' : '#ffcc80';
  // EPCC (menu5-revchart1-fontsize) - FIX theo yêu cầu người dùng "Cỡ chữ
  // toàn bộ biểu đồ tăng 20% so với hiện tại": thay vì sửa từng con số
  // fontSize rải rác, tách thành 1 hệ số nhân dùng chung — mọi fontSize
  // trong SVG này đều nhân với `FONT_SCALE`, tăng đúng 20% so với bản gốc,
  // đồng thời dễ chỉnh lại 1 chỗ duy nhất nếu sau này cần đổi tiếp.
  const FONT_SCALE = 1.2;
  // EPCC (menu5-revchart1-yoy-ratio) - đổi trục % từ đối xứng ±maxPct (kiểu
  // diff, có thể âm) sang trục 0→maxPct (kiểu tỷ lệ curr/prev, luôn ≥ 0),
  // cùng cách vẽ trục % đã dùng cho nc1 (F-Cost Target/Actual) ở trên.
  const maxPct = Math.max(0.1, ...yoyVals.filter((v): v is number => v != null));
  const xOf = (i: number) => padL + step * i + step / 2;
  const yBar = (v: number) => padT + innerH - (v / maxBar) * innerH;
  const yPct = (v: number) => padT + innerH - (v / maxPct) * innerH;

  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} preserveAspectRatio="none" role="img">
      {[0, 0.5, 1].map((tt, i) => (
        <line key={i} x1={padL} x2={width - padR} y1={padT + innerH * (1 - tt)} y2={padT + innerH * (1 - tt)}
          stroke="var(--border-soft, #2a2f3a)" strokeWidth={1} strokeDasharray="3,4" />
      ))}
      <line x1={padL} x2={padL} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={width - padR} x2={width - padR} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={padL} x2={width - padR} y1={padT + innerH} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yl${i}`} x={padL - 6} y={padT + innerH * (1 - tt) + 3} fontSize={10 * FONT_SCALE} fontWeight={700} textAnchor="end" fill={axisTextColor}>
          {fmtNum(maxBar * tt, 0)}
        </text>
      ))}
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yr${i}`} x={width - padR + 6} y={padT + innerH * (1 - tt) + 3} fontSize={10 * FONT_SCALE} fontWeight={700} textAnchor="start" fill={axisTextColor}>
          {(maxPct * tt * 100).toFixed(0)}%
        </text>
      ))}
      {shipVals.map((v, i) => {
        if (v == null) return null;
        const h = (v / maxBar) * innerH;
        const x = xOf(i) - groupW / 2;
        const y = padT + innerH - h;
        return <rect key={i} x={x} y={y} width={barW} height={Math.max(h, 1)} rx={2} fill="#2d7f96" opacity={0.9} />;
      })}
      {/* EPCC (menu5-revchart1-labels-above-bars-gap15) - FIX theo yêu cầu
          người dùng "chuyển lên Above": nhãn số cột SHIPMENT trước đây nằm
          ĐÈ BÊN TRONG cột (xoay -90°, neo gần đáy) — nay đặt NGANG (không
          xoay), NGAY PHÍA TRÊN đỉnh cột (`y(v) - 6`), giống cách đặt nhãn
          của đường DOANH SỐ/YoY bên dưới, để không còn đè lên chính cột. */}
      {shipVals.map((v, i) => {
        if (v == null) return null;
        const cx = xOf(i) - groupW / 2 + barW / 2;
        const cy = yBar(v) - 6;
        return (
          <text key={`bv${i}`} x={cx} y={cy} fontSize={8.5 * FONT_SCALE} fontWeight={700} textAnchor="middle" fill={shipLabelColor}>
            {fmtNum(v, 0)}
          </text>
        );
      })}
      {/* EPCC (menu5-revchart1-add-prodn) - cột PROD'N (KEA), đứng ngay cạnh
          phải cột SHIPMENT trong cùng 1 "group" mỗi năm, cùng trục trái/cùng
          thang `maxBar` — đúng bố cục 2 cột cạnh nhau trong ảnh tham chiếu.
          EPCC (menu5-revchart1-bar-style) - FIX theo yêu cầu người dùng đổi
          màu cột SẢN XUẤT sang màu hồng trong ảnh mẫu (mã màu #F17988). */}
      {prodVals.map((v, i) => {
        if (v == null) return null;
        const h = (v / maxBar) * innerH;
        const x = xOf(i) - groupW / 2 + barW;
        const y = padT + innerH - h;
        return <rect key={`pd${i}`} x={x} y={y} width={barW} height={Math.max(h, 1)} rx={2} fill="#F17988" opacity={0.9} />;
      })}
      {/* EPCC (menu5-revchart1-prod-label-inside-end-rotate) - FIX theo yêu
          cầu người dùng "Data Label của sản xuất để xoay 90 và inside end":
          nhãn số cột PROD'N trước đây nằm NGANG, phía TRÊN đỉnh cột (giống
          SHIPMENT) — nay đổi riêng cho cột này: xoay dọc 90° (`rotate(-90)`)
          và đặt kiểu "inside end" — điểm neo (`cy`) đặt SÁT ĐỈNH cột (bên
          trong, lùi vào 4px thay vì đẩy ra ngoài như trước), `textAnchor=
          "end"` khiến chữ bắt đầu ngay tại đỉnh và CHẠY DẦN XUỐNG vào bên
          trong thân cột (đúng hành vi "inside end" — nhãn nằm ở đầu mút của
          cột nhưng ở PHÍA TRONG, không tràn ra ngoài đỉnh). Màu chữ dùng
          `prodInsideLabelColor` (đổi theo theme) để luôn tương phản rõ với
          nền cột hồng ở cả 2 theme. */}
      {prodVals.map((v, i) => {
        if (v == null) return null;
        const cx = xOf(i) - groupW / 2 + barW + barW / 2;
        const cy = yBar(v) + 4;
        return (
          <text key={`pv${i}`} x={cx} y={cy} fontSize={8.5 * FONT_SCALE} fontWeight={700} textAnchor="end" fill={prodInsideLabelColor}
            transform={`rotate(-90 ${cx} ${cy})`}>
            {fmtNum(v, 0)}
          </text>
        );
      })}
      <path d={buildSmoothLinePath(salesVals, xOf, yBar)} fill="none" stroke="#00a65a" strokeWidth={3} strokeLinejoin="round" strokeLinecap="round" />
      {salesVals.map((v, i) => v == null ? null : (
        <g key={`sv${i}`}>
          <circle cx={xOf(i)} cy={yBar(v)} r={3} fill="#00a65a" />
          <text x={xOf(i)} y={yBar(v) - 8} fontSize={9.5 * FONT_SCALE} fontWeight={700} textAnchor="middle" fill="#00a65a">{fmtNum(v, 0)}</text>
        </g>
      ))}
      <path d={buildSmoothLinePath(yoyVals, xOf, yPct)} fill="none" stroke="#f39c12" strokeWidth={2.5} strokeDasharray="6,4" strokeLinejoin="round" strokeLinecap="round" />
      {yoyVals.map((v, i) => v == null ? null : (
        <g key={`gv${i}`}>
          <circle cx={xOf(i)} cy={yPct(v)} r={3} fill="#f39c12" />
          <text x={xOf(i)} y={yPct(v) - 8} fontSize={9.5 * FONT_SCALE} fontWeight={700} textAnchor="middle" fill={yoyLabelColor}>{(v * 100).toFixed(0)}%</text>
        </g>
      ))}
      {labels.map((lb, i) => (
        <text key={`lb${i}`} x={xOf(i)} y={height - 4} fontSize={11 * FONT_SCALE} fontWeight={700} textAnchor="middle" fill={axisTextColor}>{lb}</text>
      ))}
    </svg>
  );
}

// ── Chart 2: cột đơn giản Top 10 model theo Doanh số — top-3 tô đậm/chữ đỏ ──
function RevenueTopModelBarChart({
  items, theme,
}: {
  items: { label: string; value: number }[];
  theme: ThemeMode;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const width = 620, height = 260;
  const padL = 40, padR = 12, padT = 26, padB = 30;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;
  const n = Math.max(1, items.length);
  const step = innerW / n;
  // EPCC (menu5-chart2-3-gap20) - FIX theo yêu cầu người dùng "Chart 2,3 Gap
  // Width để còn 20%": trước đó bar chỉ chiếm 62% bề rộng mỗi item (`step`),
  // để hở (gap) 38% — nay đổi thành chiếm 80% (`step * 0.8`), gap còn đúng
  // 20% như yêu cầu.
  const barW = step * 0.8;
  const maxV = Math.max(1, ...items.map(x => x.value));
  const xOf = (i: number) => padL + step * i + step / 2;
  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} preserveAspectRatio="none" role="img">
      {[0, 0.5, 1].map((tt, i) => (
        <line key={i} x1={padL} x2={width - padR} y1={padT + innerH * (1 - tt)} y2={padT + innerH * (1 - tt)}
          stroke="var(--border-soft, #2a2f3a)" strokeWidth={1} strokeDasharray="3,4" />
      ))}
      <line x1={padL} x2={padL} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={padL} x2={width - padR} y1={padT + innerH} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yl${i}`} x={padL - 6} y={padT + innerH * (1 - tt) + 3} fontSize={10} fontWeight={700} textAnchor="end" fill={axisTextColor}>
          {fmtNum(maxV * tt, 0)}
        </text>
      ))}
      {items.map((it, i) => {
        const h = (it.value / maxV) * innerH;
        const x = xOf(i) - barW / 2;
        const y = padT + innerH - h;
        const isTop3 = i < 3;
        return (
          <g key={i}>
            <rect x={x} y={y} width={barW} height={Math.max(h, 1)} rx={2} fill={isTop3 ? '#0891b2' : '#d97706'} />
            <text x={xOf(i)} y={y - 6} fontSize={10} fontWeight={700} textAnchor="middle"
              fill={isTop3 ? (theme === 'light' ? '#e11d48' : '#fb7185') : axisTextColor}>
              {fmtNum(it.value, 1)}
            </text>
          </g>
        );
      })}
      {items.map((it, i) => (
        <text key={`lb${i}`} x={xOf(i)} y={height - 8} fontSize={9.5} fontWeight={700} textAnchor="middle" fill={axisTextColor}>
          {it.label}
        </text>
      ))}
    </svg>
  );
}

// ── Chart 5/6: cột kép SHIPMENT(KEA)/RTV(KEA) + 1 đường %Rate — dùng chung
// cho cả "RTV THEO THÁNG NĂM" (Chart 5, trục X = tháng) và "RTV THEO MODEL"
// (Chart 6, trục X = model) ──
// EPCC (menu5-revchart5-6-rtv) - THÊM MỚI theo yêu cầu người dùng "dựa theo
// ảnh 1 tham chiếu để tạo thêm 2 chart 5,6 ... lấy code màu sắc của chart1 để
// làm đúng như vậy không suy luận thêm": nhân bản CÙNG quy ước SVG thuần với
// ComboBarDualLineChart (Chart 1) — cùng padding/gutter trục Y trái-phải,
// cùng hàm buildSmoothLinePath() cho đường mượt, cùng cách tô màu chữ trục
// theo theme (`axisTextColor`). Khác biệt duy nhất so với Chart 1: 2 cột
// SHIPMENT/RTV vẽ CẠNH NHAU (clustered) thay vì chồng lên nhau, và chỉ có 1
// đường %Rate (RTV/Shipment) thay vì 2 đường Target/Actual. KHÔNG bịa màu
// mới — dùng lại NGUYÊN VẸN 3 mã màu đã có sẵn trong `SALES_CHART1_COLORS`
// (đo từ ảnh tham chiếu Chart 1): `bar` (#01877E, teal) cho cột SHIPMENT,
// `lineDashed` (#EF4444, đỏ) cho cột RTV, `line` (#4A26AB, tím) cho đường
// %Rate.
function RevenueRtvComboChart({
  labels, shipVals, rtvVals, ratePct, height = 240, theme,
}: {
  labels: string[];
  shipVals: (number | null)[];
  rtvVals: (number | null)[];
  ratePct: (number | null)[];
  height?: number;
  theme: ThemeMode;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const width = 620;
  const padL = 42, padR = 34, padT = 26, padB = 26;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;
  const n = Math.max(1, labels.length);
  const step = innerW / n;
  // EPCC (menu5-revchart5-6-stacked) - FIX theo yêu cầu người dùng "Chuyển
  // biểu đồ cột Chart 5,6 thành Stacked Column": trước đây SHIPMENT/RTV vẽ 2
  // cột CẠNH NHAU (clustered, mỗi cột nửa bề rộng step*0.8) — đổi sang 1 cột
  // DUY NHẤT mỗi category, RTV CHỒNG lên NGAY TRÊN đỉnh SHIPMENT, cùng quy
  // ước "Gap width 20%" (barW = step*0.8) như Chart 1/2, và cùng cách chồng
  // bar2-lên-bar đã có sẵn ở ComboBarDualLineChart (`bars2`) — tái dùng đúng
  // pattern, không phát minh cách vẽ mới.
  const barW = step * 0.8;
  // maxBar giờ phải là TỔNG (SHIPMENT + RTV) của từng cột vì 2 phần chồng
  // lên nhau dùng chung 1 thang đo — nếu chỉ lấy max riêng từng series như
  // trước, cột nào có cả 2 giá trị lớn sẽ bị đẩy vượt đỉnh khung.
  const maxBar = Math.max(1, ...shipVals.map((v, i) => (v ?? 0) + (rtvVals[i] ?? 0)));
  // EPCC (menu5-revchart5-6-headroom) - FIX theo yêu cầu người dùng "nhãn
  // RTV ở khoanh đỏ không nằm Above" (thực chất: cột cao nhất — 2021/SO2701
  // — chạm SÁT đỉnh khung vẽ vì `maxBar` = đúng giá trị lớn nhất, khiến nhãn
  // "Above" của nó bị clamp lên `padT + 9`, kẹt sát mép trên, trông như bị
  // cắt/không đúng vị trí). THÊM `barScaleMax` = maxBar * 1.18 (18% headroom)
  // CHỈ dùng để tính chiều cao/toạ độ CỘT (không đổi số hiển thị trên trục Y
  // — vẫn giữ `maxBar` gốc cho tick text, đúng giá trị thật) — cột cao nhất
  // giờ chỉ cao ~85% khung, chừa đủ khoảng trống phía trên cho nhãn RTV/
  // %Rate không còn bị kẹt sát viền. Đây là cách "auto headroom" tiêu chuẩn
  // của các thư viện chart (Chart.js/Highcharts đều mặc định chừa top margin
  // tương tự), không phải số bịa tuỳ tiện.
  const barScaleMax = maxBar * 1.18;
  const maxPct = Math.max(1, ...ratePct.filter((v): v is number => v != null));
  const xOf = (i: number) => padL + step * i + step / 2;
  const yPct = (v: number) => padT + innerH - (v / maxPct) * innerH;
  // EPCC (menu5-revchart5-6-darkmode-rate-color) - FIX theo yêu cầu người
  // dùng "đường %Rate màu tím tại Darkmode khó nhìn": đổi màu đường/điểm/nhãn
  // %Rate SANG '#a855f7' (NEON.violet — mã màu tím SÁNG đã có sẵn trong file,
  // dùng cho các accent khác ở Dark Mode) CHỈ khi theme==='dark'; Light Mode
  // giữ NGUYÊN '#4A26AB' (SALES_CHART1_COLORS.line) như cũ — KHÔNG đổi biến
  // dùng chung SALES_CHART1_COLORS.line vì biến này còn dùng cho đường
  // ACTUAL ở Chart 1, đổi chung sẽ ảnh hưởng ngoài phạm vi yêu cầu.
  const rateLineColor = theme === 'dark' ? '#a855f7' : SALES_CHART1_COLORS.line;

  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} preserveAspectRatio="none" role="img">
      {[0, 0.5, 1].map((tt, i) => (
        <line key={i} x1={padL} x2={width - padR} y1={padT + innerH * (1 - tt)} y2={padT + innerH * (1 - tt)}
          stroke="var(--border-soft, #2a2f3a)" strokeWidth={1} strokeDasharray="3,4" />
      ))}
      <line x1={padL} x2={padL} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={width - padR} x2={width - padR} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={padL} x2={width - padR} y1={padT + innerH} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yl${i}`} x={padL - 6} y={padT + innerH * (1 - tt) + 3} fontSize={10.2} fontWeight={700} textAnchor="end" fill={axisTextColor}>
          {fmtNum(maxBar * tt, 0)}
        </text>
      ))}
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yr${i}`} x={width - padR + 6} y={padT + innerH * (1 - tt) + 3} fontSize={10.2} fontWeight={700} textAnchor="start" fill={axisTextColor}>
          {(maxPct * tt).toFixed(1)}%
        </text>
      ))}
      {/* Đoạn đáy (base) của cột chồng — SHIPMENT, bắt đầu từ trục X (y=0). */}
      {shipVals.map((v, i) => {
        if (v == null) return null;
        const h = (v / barScaleMax) * innerH;
        const x = xOf(i) - barW / 2;
        const y = padT + innerH - h;
        return <rect key={`sb${i}`} x={x} y={y} width={barW} height={Math.max(h, 1)} rx={2} fill={SALES_CHART1_COLORS.bar} opacity={0.85} />;
      })}
      {/* Đoạn chồng phía trên — RTV, đặt NGAY TRÊN đỉnh đoạn SHIPMENT cùng
          cột (base = shipVals[i]), giống cách `bars2` chồng lên `bars` ở
          ComboBarDualLineChart (Chart 1). */}
      {rtvVals.map((v, i) => {
        if (v == null) return null;
        const base = shipVals[i] ?? 0;
        const h = (v / barScaleMax) * innerH;
        const x = xOf(i) - barW / 2;
        const y = padT + innerH - (base / barScaleMax) * innerH - h;
        return <rect key={`rb${i}`} x={x} y={y} width={barW} height={Math.max(h, 1.5)} rx={2} fill={SALES_CHART1_COLORS.lineDashed} opacity={0.9} />;
      })}
      <path d={buildSmoothLinePath(ratePct, xOf, yPct)} fill="none" stroke={rateLineColor} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round" />
      {ratePct.map((v, i) => v == null ? null : (
        <g key={`rt${i}`}>
          <circle cx={xOf(i)} cy={yPct(v)} r={2.8} fill={rateLineColor} />
          {/* EPCC (menu5-revchart5-6-label-legibility) - THEO yêu cầu người
              dùng "số hiển thị đang không nhìn rõ, phóng to 30% + đổi theme
              phải đổi màu phù hợp dễ nhìn": fontSize 10.5→13.7 (x1.3, làm
              tròn), thêm viền halo NGOÀI chữ (paintOrder="stroke") màu NỀN
              theo theme — tái dùng đúng cặp '#ffffff'/'#1e293b' đã dùng làm
              viền donut (RoseChart) trong file này — để số %Rate luôn tách
              khỏi lưới/đường bên dưới bất kể theme sáng/tối. Màu chữ dùng
              `rateLineColor` (tím sáng #a855f7 ở Dark Mode, tím đậm gốc ở
              Light Mode) khớp màu đường/điểm ngay phía trên. */}
          <text x={xOf(i)} y={yPct(v) - 8} fontSize={13.7} fontWeight={700} textAnchor="middle"
            fill={rateLineColor} stroke={theme === 'light' ? '#ffffff' : '#1e293b'}
            strokeWidth={3} paintOrder="stroke">
            {v.toFixed(1)}%
          </text>
        </g>
      ))}
      {/* EPCC (menu5-revchart5-6-ship-label-inside) - REVERT theo yêu cầu
          người dùng "Cột Shipment để chữ vào inside bar khi đổi theme cũng
          thay đổi phù hợp": đưa nhãn số SHIPMENT trở lại BÊN TRONG cột teal
          (không còn "Above" như RTV/%Rate nữa) — nhưng thay vì cố định sát
          ĐÁY như bản gốc ban đầu (đọc khó vì luôn dính mép trục X bất kể cột
          cao/thấp), CANH GIỮA theo chiều dọc đúng đoạn teal (điểm giữa từ
          đỉnh đoạn ship `topShipY` đến đáy cột) để nhãn luôn nằm gọn trong
          lòng cột dù cột cao hay thấp. Giữ NGUYÊN halo nền tương phản theo
          theme (rgba tối hơn ở theme tối) đã thêm trước đó — đây chính là
          phần "đổi theme thì màu cũng đổi phù hợp" người dùng yêu cầu. */}
      {shipVals.map((v, i) => {
        if (v == null) return null;
        const topShipY = padT + innerH - (v / barScaleMax) * innerH;
        const baselineY = padT + innerH;
        const midY = (topShipY + baselineY) / 2 + 4;
        return (
          <text key={`sv${i}`} x={xOf(i)} y={midY} fontSize={11.2} fontWeight={700} textAnchor="middle"
            fill="#ffffff" stroke={theme === 'light' ? 'rgba(15,23,42,0.55)' : 'rgba(0,0,0,0.65)'}
            strokeWidth={2.5} paintOrder="stroke">
            {fmtNum(v, 0)}
          </text>
        );
      })}
      {/* Nhãn số RTV — đặt ngay TRÊN đỉnh đoạn chồng (đoạn RTV thường mỏng
          hơn nhiều so với SHIPMENT nên không đặt bên trong được), giống cách
          Chart 1 đặt nhãn `bars2` phía trên đỉnh cột tổng. Cùng halo nền
          theo theme như nhãn %Rate ở trên để số đỏ luôn nổi rõ. */}
      {rtvVals.map((v, i) => {
        if (v == null) return null;
        const base = shipVals[i] ?? 0;
        const topY = padT + innerH - ((base + v) / barScaleMax) * innerH;
        return (
          <text key={`rv${i}`} x={xOf(i)} y={Math.max(topY - 7, padT + 9)} fontSize={11.2} fontWeight={700}
            textAnchor="middle" fill={SALES_CHART1_COLORS.lineDashed}
            stroke={theme === 'light' ? '#ffffff' : '#1e293b'} strokeWidth={3} paintOrder="stroke">
            {fmtNum(v, 0)}
          </text>
        );
      })}
      {labels.map((lb, i) => (
        <text key={`lb${i}`} x={xOf(i)} y={height - 4} fontSize={10.5} fontWeight={700} textAnchor="middle" fill={axisTextColor}>
          {lb}
        </text>
      ))}
    </svg>
  );
}

// ── Chart 3: cột chồng Doanh số theo Khách hàng qua các năm ──
// EPCC (menu5-revchart3-viewby-link) - đổi prop `years: number[]` sang
// `labels: (string | number)[]`, cùng lý do với RevenueTrendChart phía trên
// (cần nhận nhãn "Q1..Q4"/"JAN..DEC" khi XEM THEO đổi sang Quý/Tháng).
function RevenueCustomerStackedChart({
  labels, series, theme,
}: {
  labels: (string | number)[];
  series: { label: string; color: string; values: number[] }[];
  theme: ThemeMode;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const width = 620, height = 260;
  const padL = 44, padR = 12, padT = 26, padB = 30;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;
  const n = Math.max(1, labels.length);
  const step = innerW / n;
  // EPCC (menu5-chart2-3-gap20) - FIX theo yêu cầu người dùng "Chart 2,3 Gap
  // Width để còn 20%": trước đó bar chỉ chiếm 60% bề rộng mỗi năm (`step`),
  // để hở (gap) 40% — nay đổi thành chiếm 80% (`step * 0.8`), gap còn đúng
  // 20% như yêu cầu.
  const barW = step * 0.8;
  const totals = labels.map((_, i) => series.reduce((a, s) => a + (s.values[i] ?? 0), 0));
  const maxV = Math.max(1, ...totals);
  const xOf = (i: number) => padL + step * i + step / 2;
  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} preserveAspectRatio="none" role="img">
      {[0, 0.5, 1].map((tt, i) => (
        <line key={i} x1={padL} x2={width - padR} y1={padT + innerH * (1 - tt)} y2={padT + innerH * (1 - tt)}
          stroke="var(--border-soft, #2a2f3a)" strokeWidth={1} strokeDasharray="3,4" />
      ))}
      <line x1={padL} x2={padL} y1={padT} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      <line x1={padL} x2={width - padR} y1={padT + innerH} y2={padT + innerH} stroke="var(--text-2, #9aa3b2)" strokeWidth={1} />
      {[0, 0.5, 1].map((tt, i) => (
        <text key={`yl${i}`} x={padL - 6} y={padT + innerH * (1 - tt) + 3} fontSize={10} fontWeight={700} textAnchor="end" fill={axisTextColor}>
          {fmtNum(maxV * tt, 0)}
        </text>
      ))}
      {labels.map((_, i) => {
        let acc = 0;
        const x = xOf(i) - barW / 2;
        return (
          <g key={i}>
            {series.map((s, si) => {
              const v = s.values[i] ?? 0;
              if (v <= 0) return null;
              const h = (v / maxV) * innerH;
              const y = padT + innerH - (acc / maxV) * innerH - h;
              acc += v;
              return <rect key={si} x={x} y={y} width={barW} height={Math.max(h, 0.5)} fill={s.color} />;
            })}
            {totals[i] > 0 && (
              <text x={xOf(i)} y={padT + innerH - (acc / maxV) * innerH - 6} fontSize={9.5} fontWeight={700}
                textAnchor="middle" fill={axisTextColor}>
                {fmtNum(totals[i], 0)}
              </text>
            )}
          </g>
        );
      })}
      {labels.map((lb, i) => (
        <text key={`lb${i}`} x={xOf(i)} y={height - 8} fontSize={11} fontWeight={700} textAnchor="middle" fill={axisTextColor}>{lb}</text>
      ))}
    </svg>
  );
}

// EPCC (menu5-chart4-donut-enlarge) - FIX theo yêu cầu người dùng "Chart 4
// phóng to hết cỡ đường tròn để nhìn rõ hơn (cả ring bên ngoài lẫn ring bên
// trong)": trước đây `rOuter = size * 0.38` và `rInner = size * 0.24` khiến
// donut chỉ chiếm ~76% bề rộng khung SVG (viền ngoài) với vòng khá mỏng —
// để lại nhiều khoảng trống thừa quanh donut. Tăng cả 2 bán kính lên
// `rOuter = size * 0.47` (gần sát mép khung, "hết cỡ") và `rInner = size *
// 0.30` (nới rộng độ dày ring theo cùng tỉ lệ) — donut giờ to hơn rõ rệt ở
// cả viền ngoài lẫn viền trong, dễ nhìn hơn hẳn, vẫn chừa đủ chỗ giữa tâm
// cho biểu tượng 💡.
// EPCC (menu5-chart4-donut-ring-thick-pct-icon) - FIX theo yêu cầu người
// dùng "cho rộng vào trong hơn thu nhỏ nhất có thể, hiển thị % như ảnh 1
// tham chiếu, icon giữa thay bằng như ảnh 1": (1) giảm `rInner` từ 0.30*size
// XUỐNG 0.16*size (thu nhỏ lỗ trong GẦN NHƯ TỐI ĐA, chỉ vừa đủ chỗ cho icon
// mới) trong khi GIỮ NGUYÊN `rOuter=0.47*size` — ring dày hẳn ra, khớp tỉ lệ
// dày/mỏng ở ảnh 1 tham chiếu. (2) Nhãn đổi từ chỉ có `value` sang
// `"value, xx%"` (đúng định dạng ảnh 1, VD "335,557, 54%"): lát cắt đủ lớn
// (>= LEADER_THRESHOLD) vẫn đặt nhãn NGAY GIỮA lát cắt như cũ; lát cắt QUÁ
// NHỎ (dưới ngưỡng, không đủ chỗ chứa chữ) chuyển sang nhãn NGOÀI ring kèm
// 1 đường leader-line mảnh nối từ mép ngoài lát cắt ra nhãn — đúng cách ảnh
// 1 xử lý lát "4,845 , 1%". (3) Thay hẳn icon 💡 (emoji, không khớp phong
// cách ảnh 1) bằng `DonutCenterNetworkIcon` — icon trang trí dạng "mạng lưới
// nút" (các chấm tròn nối bằng đường mảnh trên nền tròn tối) vẽ mới bằng SVG
// thuần, không dùng icon/ảnh có sẵn của bên thứ ba.
// EPCC (menu5-donut-center-icon-3d) - FIX theo yêu cầu người dùng "Hình phía
// trong vòng tròn hãy tạo Icon 3D dạng sống động hơn": icon mạng lưới cũ
// (menu5-chart4-donut-ring-thick-pct-icon) hoàn toàn PHẲNG — nền tròn 1 màu
// đặc + chấm/đường mảnh cùng độ sáng, không có cảm giác chiều sâu. Viết lại
// TOÀN BỘ phần vẽ bằng SVG thuần (không icon/ảnh bên thứ ba), tạo hiệu ứng
// "3D" đúng kỹ thuật chuẩn cho SVG tĩnh (không có WebGL/3D thật):
//  1) Quả cầu nền dùng `radialGradient` lệch tâm (highlight ở góc trên-trái,
//     tối dần ra rìa) để mắt đọc ra hình khối cầu lồi, thay vì hình tròn dẹt.
//  2) Viền ngoài (bezel) có gradient kim loại sáng/tối xen kẽ, mô phỏng ánh
//     kim loại phản chiếu quanh mép — tăng cảm giác "vật thể" thay vì hình vẽ.
//  3) Mỗi nút mạng lưới cũng là 1 quả cầu nhỏ có gradient + `feGaussianBlur`
//     (glow) riêng, thay vì chấm tròn phẳng — nút trung tâm rực sáng nhất,
//     đúng vai trò "nguồn sáng" của toàn khối.
//  4) `feDropShadow` cho khối line nối, tạo cảm giác các đường "nổi" nhẹ trên
//     mặt cầu thay vì dán phẳng lên trên.
// `uid` (từ `useId()`) đảm bảo id gradient/filter không trùng nếu component
// được render nhiều lần trên cùng 1 trang.
// EPCC (menu5-donut-center-im-logo) - FIX theo yêu cầu người dùng "cho chữ
// logo đúng như ảnh 1 vào vị trí khoanh đỏ, bỏ icon đó đi và thay bằng logo
// này": bỏ HẲN icon "mạng lưới nút 3D" cũ (DonutCenterNetworkIcon, xem lịch
// sử EPCC phía trên) — thay bằng logo "iM" (oval viền xanh navy, chữ "iM"
// đậm màu xanh navy, dấu chấm chữ i tách rời) đúng theo ảnh logo công ty
// người dùng cung cấp. Vẽ lại bằng SVG thuần (không dùng ảnh/icon bên thứ
// ba) để giữ nét sắc nét ở mọi kích thước donut — nền tròn trắng (đổi màu
// theo `theme` để tương phản đủ với nền dashboard tối), oval viền dày vẽ
// bằng <ellipse> chỉ có stroke (không fill) đúng dạng viền rỗng của logo
// gốc, chữ "iM" đặt giữa oval bằng font đậm; dấu chấm chữ "i" tách thành 1
// <circle> riêng (không dùng dấu chấm mặc định của font) để khớp hình khối
// tròn đặc trưng trong logo gốc thay vì chấm nhỏ thông thường.
function DonutCenterLogo({ cx, cy, r, theme }: { cx: number; cy: number; r: number; theme: ThemeMode }) {
  // EPCC (menu5-donut-logo-exact-font-color) - (lịch sử fix trước) đo màu
  // navy trực tiếp từ ảnh gốc → `#04508f`.
  // EPCC (menu5-donut-logo-industrial-rounded-font) - FIX theo yêu cầu người
  // dùng "sửa chữ theo đúng phong cách kiểu chữ: Industrial Geometric
  // Rounded Logotype": bản trước ưu tiên `'Century Gothic'` — ĐÚNG là 1 font
  // hình học (geometric) nhưng góc/đầu nét lại KHÁ SẮC (không "rounded"),
  // không đúng 3 tiêu chí người dùng nêu ra cùng lúc (industrial + geometric
  // + rounded). Đổi hẳn thứ tự ưu tiên font sang nhóm font ĐÚNG cả 3 tiêu
  // chí — "industrial geometric rounded logotype" là mô tả kinh điển của
  // dòng font logo công nghiệp châu Âu thập niên 70-90 (VAG Rounded, dùng
  // cho logo Volkswagen/Audi cũ, Eurostile Rounded — dùng nhiều cho logo
  // hãng xe/công nghiệp/công nghệ): thân chữ dày đều tăm tắp (industrial),
  // khung chữ dựng trên hình khối tròn/vuông đơn giản (geometric), mọi đầu
  // nét và góc đều bo tròn thay vì vuông sắc (rounded) — ĐÚNG 3 đặc điểm đã
  // thấy ở ảnh logo gốc. Thứ tự fallback mới: `'VAG Rounded'` (khớp nhất cả
  // 3 tiêu chí, phổ biến trên máy có cài font hãng) → `'Eurostile Rounded'`
  // → `'Century Gothic'` (geometric, fallback gần nhất nếu máy không có 2
  // font trên) → `'Segoe UI Rounded'`/`'SF Pro Rounded'` (rounded, có sẵn
  // theo hệ điều hành Windows/macOS) → `Verdana`/sans-serif (an toàn cuối).
  // Thêm `letterSpacing` nhẹ cho chữ "M" — đặc trưng thường thấy ở logotype
  // công nghiệp (chữ giãn nhẹ, không sát khít như văn bản thường); bo tròn
  // thêm góc thân chữ "i" (tăng `rx` bo góc) cho khớp cảm giác "rounded"
  // đồng bộ với oval viền ngoài.
  const logoBlue = '#04508f';
  const logoFont = "'VAG Rounded', 'Eurostile Rounded', 'Century Gothic', 'Segoe UI Rounded', 'SF Pro Rounded', Verdana, sans-serif";
  const rx = r * 0.98;
  const ry = r * 0.6;
  const strokeW = Math.max(ry * 0.16, 1.4);
  return (
    <g>
      {/* Nền tròn phía sau oval — trắng ở theme light, xám xanh rất tối ở
          theme dark để oval trắng-navy vẫn nổi rõ trên nền dashboard tối. */}
      <circle cx={cx} cy={cy} r={r} fill={theme === 'light' ? '#ffffff' : '#101826'} />
      {/* Oval viền — chỉ có stroke, không fill, đúng kiểu logo gốc */}
      <ellipse cx={cx} cy={cy} rx={rx} ry={ry} fill="none" stroke={logoBlue} strokeWidth={strokeW} />
      {/* Chữ "M" */}
      <text x={cx + rx * 0.24} y={cy + ry * 0.42} fontSize={ry * 1.5} fontWeight={900}
        fontFamily={logoFont} letterSpacing={ry * 0.02} fill={logoBlue} textAnchor="middle">M</text>
      {/* Thân chữ "i" (không dùng dấu chấm mặc định của font) — góc bo tròn
          hơn để khớp phong cách "rounded" của font logotype công nghiệp */}
      <rect x={cx - rx * 0.34 - ry * 0.09} y={cy - ry * 0.18} width={ry * 0.18} height={ry * 0.7} rx={ry * 0.09} fill={logoBlue} />
      {/* Dấu chấm chữ "i" — chấm tròn đặc, tách rời phía trên thân chữ */}
      <circle cx={cx - rx * 0.34} cy={cy - ry * 0.52} r={ry * 0.16} fill={logoBlue} />
    </g>
  );
}

function RevenueDonutChart({
  items, size = 260, theme,
}: {
  items: { label: string; value: number; color: string }[];
  size?: number;
  theme: ThemeMode;
}) {
  const total = items.reduce((a, x) => a + x.value, 0);
  const cx = size / 2, cy = size / 2;
  const rOuter = size * 0.47, rInner = size * 0.16;
  const LEADER_THRESHOLD = 0.06; // lát cắt nhỏ hơn 6% → đưa nhãn ra ngoài kèm leader-line
  let angle = -Math.PI / 2;
  const arcs = items.map(it => {
    const frac = total > 0 ? it.value / total : 0;
    const start = angle;
    const end = angle + frac * Math.PI * 2;
    angle = end;
    const mid = (start + end) / 2;
    const largeArc = end - start > Math.PI ? 1 : 0;
    const p = (r: number, a: number) => [cx + r * Math.cos(a), cy + r * Math.sin(a)];
    const [x1, y1] = p(rOuter, start), [x2, y2] = p(rOuter, end);
    const [x3, y3] = p(rInner, end), [x4, y4] = p(rInner, start);
    const d = `M ${x1} ${y1} A ${rOuter} ${rOuter} 0 ${largeArc} 1 ${x2} ${y2} L ${x3} ${y3} A ${rInner} ${rInner} 0 ${largeArc} 0 ${x4} ${y4} Z`;
    const [lx, ly] = p((rOuter + rInner) / 2, mid);
    return { ...it, d, frac, mid, lx, ly };
  });
  return (
    <svg viewBox={`0 0 ${size} ${size}`} width="100%" height={size} role="img">
      {arcs.map((a, i) => (
        <path key={i} d={a.d} fill={a.color} stroke={theme === 'light' ? '#ffffff' : '#1e293b'} strokeWidth={2} />
      ))}
      {arcs.filter(a => a.frac >= LEADER_THRESHOLD).map((a, i) => (
        <text key={`t${i}`} x={a.lx} y={a.ly} fontSize={11.5} fontWeight={700} textAnchor="middle" fill="#ffffff">
          {fmtNum(a.value, 0)}, {(a.frac * 100).toFixed(0)}%
        </text>
      ))}
      {arcs.filter(a => a.frac > 0 && a.frac < LEADER_THRESHOLD).map((a, i) => {
        const p = (r: number, ang: number) => [cx + r * Math.cos(ang), cy + r * Math.sin(ang)];
        const [ex, ey] = p(rOuter, a.mid);
        const [lx2, ly2] = p(rOuter + 16, a.mid);
        const goRight = Math.cos(a.mid) >= 0;
        const tx = lx2 + (goRight ? 4 : -4);
        return (
          <g key={`lo${i}`}>
            <line x1={ex} y1={ey} x2={lx2} y2={ly2} stroke={a.color} strokeWidth={1.2} />
            <text x={tx} y={ly2 + 3} fontSize={10} fontWeight={700} textAnchor={goRight ? 'start' : 'end'} fill={theme === 'light' ? '#000000' : '#e5e7eb'}>
              {fmtNum(a.value, 0)}, {(a.frac * 100).toFixed(0)}%
            </text>
          </g>
        );
      })}
      <DonutCenterLogo cx={cx} cy={cy} r={rInner * 0.86} theme={theme} />
    </svg>
  );
}

// ── Panel "Tăng giảm giữa các năm" — lồng cạnh donut Chart 4, nhân bản đúng
// bố cục ảnh tham chiếu người dùng cung cấp ──
// EPCC (menu5-revchart4-yoy-panel) - THÊM MỚI: mỗi năm 1 hàng, năm liền
// trước không xác định (năm đầu tiên) để TRỐNG (không vẽ số/thanh) — đúng ô
// "2017" trống trong ảnh mẫu. Năm TĂNG (diff>0): vẽ 1 thanh ngang màu xanh
// dương với số trắng bên trong (khớp các năm 2018-2021 trong ảnh). Năm GIẢM
// (diff<0): CHỈ hiện số màu đỏ, KHÔNG vẽ thanh — đúng theo ảnh mẫu (2022,
// 2023, 2025 chỉ có chữ đỏ, không có thanh), giữ nguyên chủ ý "nhấn chữ đỏ
// cảnh báo suy giảm thay vì vẽ thanh chiếm diện tích" của ảnh gốc. KHÔNG bịa
// hex mới — dùng lại `#3b82f6` (đã có sẵn trong CHART_PALETTE_9) cho thanh
// tăng và `SALES_CHART1_COLORS.lineDashed` (#EF4444, đã dùng làm màu đỏ ở
// Chart 1/5/6) cho số giảm.
function YoyChangePanel({
  rows, title, theme,
}: {
  rows: { year: number; diff: number | null }[];
  title: string;
  theme: ThemeMode;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const rowH = 24;
  // EPCC (menu5-yoy-panel-year-gap-fix) - width tăng 230→250 để bù lại phần
  // yearGap tăng thêm bên dưới (chữ năm & số âm/dương đang đè sát nhau, xem
  // comment yearGap ngay dưới) mà KHÔNG làm hẹp lại độ dài thanh tăng dương.
  const width = 250;
  const height = rows.length * rowH + 6;
  const maxAbs = Math.max(1, ...rows.map(r => Math.abs(r.diff ?? 0)));
  // EPCC (menu5-yoy-panel-diverging-layout) - THAY layout "1 cột dồn phải"
  // bằng layout "diverging 2 chiều": năm nằm CHÍNH GIỮA làm trục 0, số ÂM
  // (đỏ, không vẽ thanh — giữ nguyên chủ ý cũ) đẩy sang BÊN TRÁI trục, thanh
  // DƯƠNG (xanh, kèm số trắng bên trong) đẩy sang BÊN PHẢI trục — đúng yêu
  // cầu người dùng "âm bên trái / dương bên phải". Không bịa màu mới, dùng
  // lại '#3b82f6' và SALES_CHART1_COLORS.lineDashed như cũ.
  const zeroX = width / 2;
  // EPCC (menu5-yoy-panel-year-gap-fix) - FIX theo yêu cầu người dùng "chữ
  // hiển thị bên trái và phải của các năm đang sát năm quá rất khó nhìn":
  // sau khi tăng fontSize 20% (menu5-yoy-panel-title-oneline-fontup20), nhãn
  // năm (VD "2023") và số chênh lệch bên cạnh (VD "-1,425") rộng hơn hẳn
  // nhưng `yearGap=15` cũ giữ nguyên → 2 khối chữ đè/dính sát vào nhau (đúng
  // hiện tượng khoanh đỏ trong ảnh "20213,248"/"20198,849"). Tăng yearGap
  // 15→26 để chừa khoảng trống rõ ràng 2 bên nhãn năm, không còn dính chữ.
  const yearGap = 26; // khoảng trống 2 bên trục cho nhãn năm không bị đè
  const leftEdge = zeroX - yearGap; // số âm neo phải tại đây, kéo dài sang trái
  const rightStart = zeroX + yearGap; // thanh dương bắt đầu từ đây, kéo dài sang phải
  const barMaxW = width / 2 - yearGap - 6;
  // EPCC (menu5-yoy-panel-title-oneline-fontup20) - FIX theo yêu cầu người
  // dùng "cho phần tiêu đề Tăng giảm giữa các năm (K$) nằm trên 1 hàng, top=
  // 0.2mm, và phóng to chữ phần này lên to hơn 20% nữa": (1) title trước đây
  // cố ý xuống 2 dòng bằng '\n' + whiteSpace:'pre-line' (xem TEXT.revYoyPanelTitle
  // ở 3 ngôn ngữ) — đổi whiteSpace sang 'nowrap' để LUÔN hiển thị trên đúng 1
  // hàng bất kể chuỗi còn '\n' hay không (an toàn cho cả 3 ngôn ngữ). (2) Padding
  // trên cùng đổi từ '6px 4px' (padding đều 4 phía) sang paddingTop riêng
  // '0.2mm' (CSS hỗ trợ thẳng đơn vị mm, không cần quy đổi px) + giữ nguyên
  // padding dưới/trái/phải 6px/4px. (3) MỌI fontSize trong panel này (tiêu đề +
  // nhãn năm + số trong thanh xanh + số đỏ) tăng thêm đúng 20% so với gốc:
  // 10.5→12.6, 11→13.2, 9→10.8, 10→12 — không đổi màu/logic khác.
  return (
    <div style={{ width, flexShrink: 0 }}>
      <div style={{
        fontSize: 12.6, fontWeight: 800, textAlign: 'center', whiteSpace: 'nowrap', lineHeight: 1.2,
        paddingTop: '0.2mm', paddingBottom: 6, paddingLeft: 4, paddingRight: 4,
        color: theme === 'light' ? '#92400e' : '#facc15',
        background: theme === 'light' ? '#fde68a' : 'rgba(250,204,21,0.12)',
        borderRadius: 8, marginBottom: 4,
      }}>
        {title.replace(/\n/g, ' ')}
      </div>
      <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} role="img">
        {rows.map((r, i) => {
          const cy = i * rowH + rowH / 2;
          if (r.diff == null) {
            return (
              <text key={i} x={zeroX} y={cy + 4} fontSize={13.2} fontWeight={700} textAnchor="middle" fill={axisTextColor} opacity={0.6}>
                {r.year}
              </text>
            );
          }
          if (r.diff > 0) {
            const barW = Math.max((r.diff / maxAbs) * barMaxW, 2);
            // EPCC (menu5-yoy-panel-label-inside-base) - FIX theo yêu cầu người
            // dùng "chỗ phần khoanh để thành Inside Base": trước đây số (VD
            // "43,344") neo tại MÉP PHẢI thanh (`x = rightStart + barW - 4`,
            // `textAnchor='end'` — kiểu nhãn "Inside End"). Đổi sang neo TẠI
            // MÉP TRÁI/GỐC của thanh — đúng vị trí thanh bắt đầu mọc ra từ
            // trục 0 (`x = rightStart + 4`, `textAnchor='start'`) — kiểu nhãn
            // "Inside Base" của các thư viện chart: chữ số luôn nằm NGAY SÁT
            // mép trong bên TRÁI (gốc) thanh xanh, đọc từ trái sang phải theo
            // đúng hướng thanh mọc, thay vì bám mép phải như trước.
            return (
              <g key={i}>
                <text x={zeroX} y={cy + 4} fontSize={13.2} fontWeight={700} textAnchor="middle" fill={axisTextColor}>
                  {r.year}
                </text>
                <rect x={rightStart} y={cy - 8} width={barW} height={16} rx={3} fill="#3b82f6" />
                {/* EPCC (menu5-yoy-panel-bar-label-theme-aware) - FIX theo yêu
                    cầu người dùng "chữ màu trắng không phù hợp với theme
                    Lightmode, cần chuyển màu đen/trắng khi chuyển theme":
                    trước đây số bên trong thanh xanh LUÔN `fill="#ffffff"`
                    bất kể theme — ở theme light, nền thanh `#3b82f6` là màu
                    xanh dương khá nhạt/sáng so với nền trắng xung quanh, chữ
                    trắng chìm vào thanh, khó đọc (đúng vùng khoanh đỏ trong
                    ảnh người dùng gửi). Đổi sang tính theo `theme`: giữ
                    nguyên trắng ở theme dark (nền tối, chữ trắng tương phản
                    tốt với cả thanh lẫn nền xung quanh), đổi sang chữ đen ở
                    theme light (tương phản rõ hơn hẳn với thanh xanh nhạt
                    trên nền sáng). */}
                <text x={rightStart + 4} y={cy + 4} fontSize={10.8} fontWeight={700} textAnchor="start" fill={theme === 'light' ? '#0f172a' : '#ffffff'}>
                  {/* EPCC (menu5-yoy-panel-arrow-icons) - THÊM MỚI theo yêu cầu
                      người dùng "số âm có thêm mũi tên xuống phía sau, số
                      dương mũi tên đi lên" (đúng ảnh tham chiếu): nối thêm
                      ký tự mũi tên NGAY SAU chữ số, cùng 1 <text>/fill với
                      số — không cần tính lại vị trí x riêng, vì mũi tên chỉ
                      là 1 ký tự nối tiếp cuối chuỗi, tự động bám sát ngay sau
                      số theo hướng đọc trái→phải (giống "6,248⇧" trong ảnh).
                      Dùng ký tự mũi tên đậm sẵn có trong Unicode ('⬆'/'⬇'),
                      không cần thêm path SVG mới. */}
                  {fmtNum(r.diff, 0)} ⬆
                </text>
              </g>
            );
          }
          return (
            <g key={i}>
              <text x={leftEdge} y={cy + 4} fontSize={12} fontWeight={800} textAnchor="end" fill={SALES_CHART1_COLORS.lineDashed}>
                {/* EPCC (menu5-yoy-panel-arrow-icons) - cùng nguyên tắc ở
                    nhánh dương phía trên: nối thêm '⬇' NGAY SAU số âm. Vì
                    <text> này dùng textAnchor="end" (toàn chuỗi neo mép PHẢI
                    tại `leftEdge`), mũi tên nối cuối chuỗi vẫn hiển thị ở
                    VỊ TRÍ PHẢI CÙNG (ngay sau số, sát trục năm) — đúng
                    "phía sau" số âm như ảnh tham chiếu ("-31,800⬇"), không
                    bị đẩy lệch ra ngoài mép trái như nếu đặt trước số.
                    Cùng màu đỏ với số (`SALES_CHART1_COLORS.lineDashed`). */}
                {fmtNum(r.diff, 0)} ⬇
              </text>
              <text x={zeroX} y={cy + 4} fontSize={13.2} fontWeight={700} textAnchor="middle" fill={axisTextColor}>
                {r.year}
              </text>
            </g>
          );
        })}
      </svg>
    </div>
  );
}

interface Menu5Props {
  theme: ThemeMode;
  onToggleTheme?: () => void;
  lang: Lang;
  setLang?: (l: Lang) => void;
}

export default function Menu5ModelDashboard({ theme, lang }: Menu5Props) {
  const t = TEXT[lang] ?? TEXT.vi;
  // Màu nhãn toolbar — lấy nguyên văn từ TargetActualDashboard.tsx
  // (filterLabelColor) để 2 dashboard đồng bộ, không tự bịa màu mới.
  const filterLabelColor = '#C0EF6A';
  const [tab, setTab] = useState<'overview' | 'detail' | 'revenue'>('revenue');
  const [viewBy, setViewBy] = useState<ViewBy>('month');

  // EPCC (menu5-header-tabs-match-rty) - FIX ROOT CAUSE "thanh header/tab
  // Mục 5 dùng class 'topbar' chung + NeonButton, khác hẳn màu sắc/kích
  // thước với Mục 4 (RTY)": copy NGUYÊN VĂN state đồng hồ sống từ
  // RtyDashboard.tsx (formattedTime, cập nhật mỗi giây) để thanh header
  // Mục 5 hiển thị đúng định dạng "HH:mm:ss dd/mm/yyyy" giống hệt RTY,
  // thay cho `now.toLocaleTimeString()` tĩnh (không tự cập nhật) cũ.
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  const formattedTime = useMemo(() => {
    const p = (n: number) => String(n).padStart(2, '0');
    return `${p(now.getHours())}:${p(now.getMinutes())}:${p(now.getSeconds())} ${p(now.getDate())}/${p(now.getMonth() + 1)}/${now.getFullYear()}`;
  }, [now]);

  // ── Nguồn dữ liệu: mặc định = mẫu nhúng sẵn, có thể thay bằng file Excel
  // người dùng tải lên (cache qua IndexedDB để không mất khi F5). ──────────
  const [dataSource, setDataSource] = useState<Test5Data>(EMPTY_DATA);
  const [lastUpdated, setLastUpdated] = useState<string | null>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    (async () => {
      const cached = await idbGetCache(IDB_KEY_MENU5_DATA);
      const meta = await idbGetCache(IDB_KEY_MENU5_META);
      if (cached) {
        try {
          const parsed = JSON.parse(cached);
          // EPCC (menu5-revenue-tab-sales-by-customer) - cache cũ (trước khi
          // thêm tab "TÌNH HÌNH DOANH THU") sẽ không có field này → fallback
          // rỗng thay vì để undefined làm vỡ các chart mới khi đọc lại từ IDB.
          if (!parsed.salesByCustomer) parsed.salesByCustomer = {};
          // EPCC (menu5-revenue-chart3-viewby-monthly-customer) - fallback
          // rỗng tương tự cho field mới, phòng hờ dù đã bump version cache.
          if (!parsed.salesByCustomerMonthly) parsed.salesByCustomerMonthly = {};
          setDataSource(parsed);
          if (meta) setLastUpdated(meta);
        } catch { /* cache hỏng, bỏ qua, dùng mẫu mặc định */ }
      }
    })();
  }, []);

  async function handleUploadFile(file: File) {
    setIsParsing(true);
    setUploadError(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
      // EPCC (menu5-multisheet-upload): gộp TẤT CẢ sheet trong workbook,
      // không chỉ sheet đầu tiên — xem chi tiết root cause ở comment tại
      // định nghĩa aggregateWorkbookToTest5Data() phía trên.
      const parsed = aggregateWorkbookToTest5Data(wb);
      const nowIso = new Date().toISOString();
      setDataSource(parsed);
      setLastUpdated(nowIso);
      await idbSetCache(IDB_KEY_MENU5_DATA, JSON.stringify(parsed));
      await idbSetCache(IDB_KEY_MENU5_META, nowIso);
    } catch (err) {
      setUploadError(err instanceof Test5ParseError ? err.message : 'Không đọc được file Excel này.');
    } finally {
      setIsParsing(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  }

  const allModels = dataSource.models;
  const bestModel = dataSource.bestModel;
  // EPCC (menu5-revenue-tab-sales-by-customer) - dữ liệu nguồn cho tab mới
  // "TÌNH HÌNH DOANH THU" (Chart "Doanh số theo Khách hàng" + donut "Tỷ
  // trọng theo Khách hàng"). Fallback `?? {}` cho dữ liệu cache cũ.
  const salesByCustomer = dataSource.salesByCustomer ?? {};
  // EPCC (menu5-revenue-chart3-viewby-monthly-customer) - dữ liệu THÁNG theo
  // khách hàng, dùng cho Chart 3 khi XEM THEO = Quý/Tháng.
  const salesByCustomerMonthly = dataSource.salesByCustomerMonthly ?? {};
  // EPCC (menu5-model-all-option) - GHI CHÚ LỊCH SỬ: từng có `firstRealModel`/
  // `safeBestModel` ở đây làm "lưới an toàn" cho giá trị mặc định của
  // `selectedModel` (phòng khi mặc định lỡ trỏ vào 'Target'/'Actual' — 2
  // model đặc biệt bị ẩn khỏi <select>). Từ khi đổi mặc định `selectedModel`
  // sang thẳng `ALL_MODEL_VALUE` (xem menu5-model-default-all), rủi ro đó
  // không còn tồn tại nữa (luôn khớp đúng option "Tất cả"), nên đã xoá 2
  // biến không còn dùng tới thay vì giữ lại code chết.

  // EPCC (menu5-model-all-chart1-empty) - FIX ROOT CAUSE "chọn 'Tất cả' ở
  // MODEL thì Chart 1 (ALL YEAR) hiện 'Không có dữ liệu trong khoảng đã
  // chọn', dù Chart 2 vẫn vẽ được đường F-Cost%": đối chiếu trực tiếp Excel
  // xác nhận dòng MODEL=Actual (và Target) trong file CHỈ có các dòng với
  // ITEM = "F-Cost %" — KHÔNG có dòng nào cho SHIPMENT/PROD'N/SALES
  // AMT/SCRAP. Bản fix trước (menu5-model-all-option) trỏ thẳng `modelData`
  // vào dòng 'Actual' này khi chọn "Tất cả", đúng cho F-Cost% (nên Chart 2 —
  // vốn chỉ dùng F-Cost% — vẫn chạy đúng), nhưng Chart 1 còn cần cột bar
  // SHIPMENT + bảng chi tiết còn cần cả 4 cột số lượng (Shipment/Prod'n/
  // Sales/Scrap) — tất cả đều RỖNG ở dòng 'Actual' → bars toàn `null` →
  // card hiện "Không có dữ liệu".
  // Fix: khi chọn "Tất cả", KHÔNG dùng nguyên dòng 'Actual' làm `modelData`
  // nữa, mà DỰNG 1 ModelEntry tổng hợp riêng:
  //  - F-Cost % (series + ttlByItemYear): lấy Y NGUYÊN từ dòng 'Actual' có
  //    sẵn (KHÔNG cộng dồn lại từ real model — đây chính là số liệu tổng
  //    hợp công ty đã có sẵn trong Excel, theo đúng yêu cầu gốc của người
  //    dùng).
  //  - SHIPMENT/PROD'N/SALES AMT/SCRAP (series + ttlByItemYear): CỘNG DỒN
  //    trực tiếp từ mọi real model (loại 4 model đặc biệt) theo từng
  //    (năm, tháng)/(item, năm) — đây là các item dạng SỐ LƯỢNG nên cộng dồn
  //    đúng bản chất (khác F-Cost% là TỶ LỆ, không được cộng dồn).
  // Nhờ vậy Chart 1 (bar Shipment + line Target/Actual), KPI, và bảng chi
  // tiết đều có đủ dữ liệu khi chọn "Tất cả", trong khi Chart 2 (chỉ dùng
  // F-Cost%, tra thẳng theo tên model 'Actual' trong `allModels` gốc, không
  // qua object tổng hợp này) không bị ảnh hưởng gì, vẫn đúng như trước.
  const realModelsForAgg = useMemo(
    () => allModels.filter(m => !SPECIAL_FCOST_MODELS.includes(m.model)),
    [allModels]
  );
  const allModelsAggregate = useMemo<ModelEntry | null>(() => {
    const actualEntry = allModels.find(m => m.model === 'Actual') ?? null;
    if (!realModelsForAgg.length && !actualEntry) return null;

    const sumItems = [ITEM_SHIPMENT, ITEM_PRODUCTION, ITEM_SALES, ITEM_SCRAP];
    const series: Record<string, SeriesPoint[]> = {};
    for (const item of sumItems) {
      const byKey = new Map<number, number>();
      for (const m of realModelsForAgg) {
        for (const p of m.series[item] ?? []) {
          const key = ym(p.y, p.m);
          byKey.set(key, (byKey.get(key) ?? 0) + p.v);
        }
      }
      series[item] = Array.from(byKey.entries())
        .map(([key, v]) => ({ y: Math.floor(key / 12), m: (key % 12) + 1, v }))
        .sort((a, b) => ym(a.y, a.m) - ym(b.y, b.m));
    }
    // F-Cost % giữ NGUYÊN dòng 'Actual' gốc — không suy luận lại.
    series[ITEM_FCOST] = actualEntry?.series[ITEM_FCOST] ?? [];

    const ttlByItemYear: Record<string, Record<number, number>> = {};
    for (const item of sumItems) {
      const byYear = new Map<number, number>();
      for (const m of realModelsForAgg) {
        for (const [yearStr, v] of Object.entries(m.ttlByItemYear[item] ?? {})) {
          const year = Number(yearStr);
          byYear.set(year, (byYear.get(year) ?? 0) + v);
        }
      }
      ttlByItemYear[item] = Object.fromEntries(byYear.entries());
    }
    ttlByItemYear[ITEM_FCOST] = actualEntry?.ttlByItemYear[ITEM_FCOST] ?? {};

    return {
      model: 'Actual',
      rowCount: realModelsForAgg.reduce((a, m) => a + m.rowCount, 0),
      series,
      radarType: [],
      radarCustom: [],
      ttlByItemYear,
    };
  }, [allModels, realModelsForAgg]);

  // Biên độ năm/tháng toàn dữ liệu, dùng làm giá trị mặc định cho bộ lọc.
  const globalBounds = useMemo(() => {
    let minYm = Infinity, maxYm = -Infinity;
    for (const md of allModels) {
      for (const pts of Object.values(md.series)) {
        for (const p of pts) {
          const v = ym(p.y, p.m);
          if (v < minYm) minYm = v;
          if (v > maxYm) maxYm = v;
        }
      }
    }
    if (!isFinite(minYm)) { minYm = ym(2017, 1); maxYm = ym(2026, 12); }
    return { minYm, maxYm };
  }, [allModels]);

  // EPCC (menu5-year-quick-filter) - FIX theo yêu cầu người dùng "bổ sung
  // thêm 1 ô truy vấn theo năm: tất cả/2017~2026 và sau này có thêm dữ liệu
  // thì cứ thêm tự động như bình thường. Lấy dữ liệu theo file đã tải lên để
  // tham chiếu": danh sách năm hiển thị trong dropdown KHÔNG hardcode
  // 2017..2026, mà dò trực tiếp từ `allModels` (dữ liệu đã tải lên/cache) —
  // gộp cả năm xuất hiện ở điểm dữ liệu theo tháng (`series`) LẪN năm chỉ có
  // dòng TTL (`ttlByItemYear`, phòng trường hợp 1 năm chỉ có dòng tổng, không
  // có dòng tháng chi tiết nào). Nhờ vậy khi người dùng tải thêm file Excel
  // có năm mới (VD 2027), dropdown tự động có thêm option năm đó — không cần
  // sửa code.
  const allDataYears = useMemo(() => {
    const yearsSet = new Set<number>();
    for (const md of allModels) {
      for (const pts of Object.values(md.series)) {
        for (const p of pts) yearsSet.add(p.y);
      }
      for (const byYear of Object.values(md.ttlByItemYear)) {
        for (const yKey of Object.keys(byYear)) yearsSet.add(Number(yKey));
      }
    }
    return Array.from(yearsSet).sort((a, b) => a - b);
  }, [allModels]);

  const toMonthInput = (v: number) => {
    const y = Math.floor(v / 12), m = (v % 12) + 1;
    return `${y}-${String(m).padStart(2, '0')}`;
  };
  const fromMonthInput = (s: string) => {
    const [y, m] = s.split('-').map(Number);
    return ym(y, m);
  };

  // EPCC (menu5-model-default-all) - FIX theo yêu cầu người dùng "để mặc
  // định Tất cả cho MODEL/CUSTOM/NĂM khi mở app, áp dụng chung cho cả tình
  // hình F-COST": trước đây MODEL mặc định trỏ vào 1 model cụ thể
  // (`safeBestModel || firstRealModel`, model có F-Cost cao nhất) — khác
  // hẳn CUSTOM (`ALL_CUSTOM_VALUE`) và NĂM (`globalBounds.minYm/maxYm`, tự
  // suy ra "all") vốn ĐÃ mặc định "Tất cả" sẵn. Đổi sang `ALL_MODEL_VALUE`
  // để 3 bộ lọc đồng nhất "Tất cả" ngay khi mở app — vì `selectedModel` là
  // 1 state DÙNG CHUNG cho cả 2 tab (F-COST và Doanh Thu, xem `modelData`),
  // đổi ở đây tự động áp dụng cho cả 2 tab, không cần sửa thêm nơi nào khác.
  const [selectedModel, setSelectedModel] = useState<string>(ALL_MODEL_VALUE);
  // EPCC (menu5-revenue-custom-filter) - state cho bộ lọc CUSTOM mới, mặc
  // định "Tất cả" (không lọc gì, giữ nguyên hành vi cũ của Chart 3/4).
  const [selectedCustom, setSelectedCustom] = useState<string>(ALL_CUSTOM_VALUE);
  // EPCC (menu5-revenue-table-full-list-filters) - THÊM MỚI theo yêu cầu
  // người dùng "liệt kê tất cả số liệu có từ file excel lên và trên đầu Card
  // top để dạng có lựa chọn như ảnh 1 tham chiếu": 6 state lọc RIÊNG cho bảng
  // "Toàn bộ Model theo Doanh số" (KHÔNG dùng chung `selectedModel`/
  // `selectedCustom` ở toolbar trên cùng — 2 bộ lọc đó chi phối TOÀN BỘ 2 tab
  // F-COST/DOANH THU, trong khi bảng này cần lọc ĐỘC LẬP y hệt ảnh tham chiếu
  // "Toàn bộ Cơ sở Dữ liệu từ Excel": Model/Năm/Quý/Tháng/Khách hàng dạng
  // dropdown riêng + 1 ô tìm kiếm tự do theo tên). Mặc định "Tất cả" cho mọi
  // dropdown và chuỗi rỗng cho ô tìm kiếm — hiển thị TOÀN BỘ dữ liệu ngay khi
  // mở tab, đúng tinh thần "liệt kê tất cả số liệu".
  const [revTableModel, setRevTableModel] = useState<string>(ALL_MODEL_VALUE);
  const [revTableYear, setRevTableYear] = useState<string>(ALL_CUSTOM_VALUE);
  const [revTableQuarter, setRevTableQuarter] = useState<string>(ALL_CUSTOM_VALUE);
  const [revTableMonth, setRevTableMonth] = useState<string>(ALL_CUSTOM_VALUE);
  const [revTableCustomer, setRevTableCustomer] = useState<string>(ALL_CUSTOM_VALUE);
  const [revTableSearch, setRevTableSearch] = useState<string>('');
  const [startYm, setStartYm] = useState<number>(globalBounds.minYm);
  const [endYm, setEndYm] = useState<number>(globalBounds.maxYm);

  // EPCC (menu5-year-quick-filter) - giá trị HIỂN THỊ của dropdown "NĂM" suy
  // ra trực tiếp từ `startYm`/`endYm` hiện tại (không dùng state riêng, tránh
  // 2 nguồn sự thật lệch nhau khi người dùng tự sửa input Tháng bắt
  // đầu/kết thúc): nếu khoảng đang chọn khớp ĐÚNG toàn bộ biên dữ liệu →
  // "all"; nếu khớp đúng trọn 1 năm (T1 → T12 của cùng 1 năm) → năm đó; các
  // trường hợp khác (khoảng tuỳ ý do người dùng tự kéo input tháng) → không
  // khớp option nào, dropdown hiển thị rỗng (không sao, vẫn không mất dữ
  // liệu lọc thực tế đang áp dụng).
  const selectedYearFilterValue = useMemo(() => {
    if (startYm === globalBounds.minYm && endYm === globalBounds.maxYm) return 'all';
    for (const y of allDataYears) {
      if (startYm === ym(y, 1) && endYm === ym(y, 12)) return String(y);
    }
    return '';
  }, [startYm, endYm, globalBounds, allDataYears]);

  function handleYearFilterChange(value: string) {
    if (value === 'all') {
      setStartYm(globalBounds.minYm);
      setEndYm(globalBounds.maxYm);
      return;
    }
    const y = Number(value);
    if (!y) return;
    // Kẹp trong biên dữ liệu thật (globalBounds) — phòng trường hợp năm đó
    // chỉ có vài tháng dữ liệu (VD năm đầu/năm cuối), tránh chọn quá phạm vi
    // thật sự có trong file Excel.
    setStartYm(Math.max(ym(y, 1), globalBounds.minYm));
    setEndYm(Math.min(ym(y, 12), globalBounds.maxYm));
  }

  // EPCC (menu5-detail-table-sync-top-toolbar-filters) - FIX theo yêu cầu
  // người dùng "khi lựa chọn theo các năm hoặc model [ở 3 ô MODEL/CUSTOM/NĂM
  // khoanh đỏ trên toolbar chính] thì dòng dữ liệu phía dưới [bảng Toàn bộ
  // Model theo Doanh số ở tab Chi tiết & dữ liệu] thay đổi theo sự lựa chọn
  // đó": trước đây 3 ô này (`selectedModel`/`selectedCustom`/
  // `selectedYearFilterValue`, dùng chung cho toàn trang) hoàn toàn KHÔNG
  // liên quan tới bộ lọc RIÊNG của bảng (`revTableModel`/`revTableCustomer`/
  // `revTableYear`) — đổi Model/Custom/Năm ở toolbar chính không hề ảnh
  // hưởng bảng bên dưới. Đồng bộ MỘT CHIỀU (toolbar chính → bộ lọc bảng):
  // đổi ở toolbar chính sẽ tự cập nhật bộ lọc bảng theo, nhưng người dùng
  // vẫn có thể tự tay đổi RIÊNG bộ lọc của bảng (Model/Năm/Quý/Tháng/Khách
  // hàng) để tinh chỉnh sâu hơn (VD xem thêm Quý/Tháng) mà không bị toolbar
  // chính ghi đè lại — chỉ đồng bộ khi CHÍNH toolbar chính thay đổi.
  // `ALL_MODEL_VALUE`/`ALL_CUSTOM_VALUE` là 2 sentinel "Tất cả" DÙNG CHUNG
  // giữa 2 nơi nên copy thẳng giá trị, không cần map lại. Riêng NĂM:
  // `selectedYearFilterValue` trả về 'all' (map sang sentinel `ALL_CUSTOM_VALUE`
  // của bảng), 1 chuỗi năm cụ thể (dùng thẳng, cùng định dạng chuỗi với
  // `revTableYear`), hoặc '' khi người dùng tự kéo khoảng THÁNG BẮT ĐẦU/KẾT
  // THÚC tuỳ ý (không khớp trọn năm nào) — trường hợp này KHÔNG đồng bộ (giữ
  // nguyên lựa chọn Năm hiện tại của bảng) vì không có 1 năm rõ ràng nào để
  // gán.
  useEffect(() => {
    setRevTableModel(selectedModel);
  }, [selectedModel]);
  useEffect(() => {
    setRevTableCustomer(selectedCustom);
  }, [selectedCustom]);
  useEffect(() => {
    if (selectedYearFilterValue === 'all') setRevTableYear(ALL_CUSTOM_VALUE);
    else if (selectedYearFilterValue) setRevTableYear(selectedYearFilterValue);
  }, [selectedYearFilterValue]);

  // EPCC (menu5-reset-filter-on-new-source) - khi upload file mới hoặc khôi
  // phục mẫu, danh sách model/khoảng thời gian có thể khác hẳn dữ liệu cũ →
  // phải reset lại filter, tránh selectedModel trỏ tới model không tồn tại
  // trong dữ liệu mới (màn hình trắng do modelData = null).
  useEffect(() => {
    // EPCC (menu5-model-default-all) - đồng bộ với default lúc khởi tạo
    // state phía trên: khi upload file mới/khôi phục mẫu cũng reset MODEL về
    // "Tất cả" thay vì trỏ lại 1 model cụ thể.
    setSelectedModel(ALL_MODEL_VALUE);
    setSelectedCustom(ALL_CUSTOM_VALUE);
    setStartYm(globalBounds.minYm);
    setEndYm(globalBounds.maxYm);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dataSource]);

  // EPCC (menu5-model-all-option / menu5-model-all-chart1-empty) - FIX theo
  // yêu cầu người dùng "chọn 'Tất cả' ở MODEL thì áp dụng toàn trang (KPI +
  // bảng chi tiết + Chart 1 & 2) = tổng hợp tất cả model": vì KPI, bảng chi
  // tiết, Chart 1 & 2 đều lấy dữ liệu qua `modelData` duy nhất, chỉ cần đổi
  // nguồn ở đây là toàn bộ các phần đó tự động đồng bộ theo. Dùng
  // `allModelsAggregate` (F-Cost% lấy nguyên dòng 'Actual', 4 item số lượng
  // còn lại cộng dồn từ real model — xem giải thích đầy đủ ở khai báo
  // `allModelsAggregate` phía trên) thay vì trỏ thẳng vào dòng 'Actual' gốc
  // (dòng đó KHÔNG có dữ liệu Shipment/Prod'n/Sales/Scrap, từng khiến Chart 1
  // trống trơn).
  const modelData = useMemo(() => {
    if (selectedModel === ALL_MODEL_VALUE) {
      return allModelsAggregate;
    }
    return allModels.find(m => m.model === selectedModel) ?? null;
  }, [allModels, selectedModel, allModelsAggregate]);








  // ── Năm đang truy vấn (dùng chung cho 4 thẻ KPI + Chart 1-4) ────────────
  // EPCC (menu5-kpi-link-year-query) - FIX theo yêu cầu người dùng "4 Card
  // KPI thay đổi link dữ liệu: Card 1/2/3 hiển thị theo NĂM đang truy vấn,
  // Card 4 so sánh năm đang truy vấn với năm liền trước": dời khai báo
  // `lastRangeYear`/`ttlOf` (vốn nằm SAU khối KPI, chỉ dùng cho Chart 1-4)
  // lên ĐÂY để 4 thẻ KPI cũng dùng chung được — không tạo bản sao logic
  // riêng, tránh 2 nơi tính "năm đang truy vấn" theo 2 cách khác nhau rồi
  // lệch nhau sau này.
  const startYear = Math.floor(startYm / 12);
  const endYear = Math.floor(endYm / 12);
  const rangeYears = useMemo(() => {
    const arr: number[] = [];
    for (let y = startYear; y <= endYear; y++) arr.push(y);
    return arr;
  }, [startYear, endYear]);
  const lastRangeYear = endYear;

  const findSpecialModel = (name: string) => allModels.find(m => m.model === name) ?? null;
  const ttlOf = (m: ModelEntry | null, item: string, year: number): number | null =>
    m?.ttlByItemYear?.[item]?.[year] ?? null;

  // ── KPI — LINK theo năm đang truy vấn (`lastRangeYear`), dùng thẳng dòng
  // TTL gốc từ Excel (`ttlOf`) thay vì tự cộng/trung bình lại từ dữ liệu
  // tháng đang lọc theo THÁNG BẮT ĐẦU/KẾT THÚC — đúng nguyên tắc đã áp dụng
  // cho Chart 1-4 (xem comment `ttlByItemYear`): TTL trong Excel là số liệu
  // GỐC, tin cậy hơn suy luận lại từ tháng (có thể lệch nếu vài tháng dữ
  // liệu lỗi/bẩn). ────────────────────────────────────────────────────────
  // Card 1 "F-COST TB": TTL F-Cost% của model đang chọn, ĐÚNG năm đang
  // truy vấn — đổi năm trên bộ lọc thì số đổi theo, không còn phụ thuộc
  // khoảng THÁNG BẮT ĐẦU/KẾT THÚC (VD lọc Jan-Jun thì trước đây chỉ tính TB
  // 6 tháng đó, giờ luôn ra ĐÚNG TB cả năm được chọn).
  const kpi1AvgFcost = ttlOf(modelData, ITEM_FCOST, lastRangeYear) ?? 0;
  // EPCC (menu5-kpi-target-from-real-target-row) - FIX ROOT CAUSE "Card 1
  // hiện 'Mục tiêu: 5%' trong khi đường TARGET trên Chart 1 (cùng năm) chỉ
  // ~1%": '5%' lấy từ hằng số VIẾT CỨNG `FCOST_TARGET = 0.05` (giá trị
  // placeholder từ giai đoạn đầu, không hề đọc từ Excel) — không liên quan
  // gì tới dòng MODEL='Target' thật trong file. Fix: tra đúng dòng TTL của
  // model 'Target' (item F-Cost %) ĐÚNG năm đang truy vấn (`lastRangeYear`),
  // y hệt cách Chart 1 lấy đường TARGET — Card 1 và Chart 1 giờ luôn khớp
  // nhau vì cùng 1 nguồn dữ liệu.
  const targetModel = findSpecialModel('Target');
  const kpiTargetTtl = ttlOf(targetModel, ITEM_FCOST, lastRangeYear);
  // Card 2 "SCRAP AMT": SỐ TIỀN Scrap (K$) TTL cả năm đang truy vấn — EPCC
  // (menu5-kpi2-scrap-amt) đổi từ tỷ lệ SCRAP/PRODUCTION (%) sang số tiền
  // thẳng, xem chi tiết comment tại JSX Card 2 phía dưới.
  const kpiScrapTtl = ttlOf(modelData, ITEM_SCRAP, lastRangeYear);
  // Card 3 "THÁNG ĐẠT MỤC TIÊU": đếm trên TOÀN BỘ các tháng CÓ dữ liệu của
  // đúng năm đang truy vấn (không còn phụ thuộc khoảng THÁNG BẮT ĐẦU/KẾT
  // THÚC) — các tháng chưa diễn ra (VD JUL-DEC/2026 chưa có số) tự động
  // không nằm trong `fcostPtsOfYear` (đã bị loại từ bước parse), nên mẫu số
  // luôn dừng ĐÚNG ở tháng gần nhất đã có số liệu thật của năm đó.
  const fcostPtsOfYear = (modelData?.series[ITEM_FCOST] ?? []).filter(p => p.y === lastRangeYear);
  // EPCC (menu5-kpi-target-from-real-target-row) - cùng gốc lỗi với "Mục
  // tiêu" ở Card 1: ngưỡng "đạt/không đạt" mỗi tháng trước đây so với
  // `FCOST_TARGET` (5% viết cứng) — quá cao so với Target thật (~1%), khiến
  // hầu như tháng nào cũng "đạt" một cách giả tạo (không phản ánh đúng thực
  // tế). Đổi sang so với ĐÚNG giá trị Target CỦA TỪNG THÁNG (tra theo dòng
  // model 'Target' tại (năm, tháng) tương ứng, không phải TTL năm) — sát
  // với ý nghĩa "tháng đó có đạt mục tiêu tháng đó hay không". Tháng nào
  // không tra được Target tương ứng thì không tính là "đạt" (loại khỏi tử
  // số, vẫn tính vào mẫu số vì tháng đó đã có F-Cost% thật).
  const targetMonthlyMap = new Map((targetModel?.series[ITEM_FCOST] ?? []).map(p => [ym(p.y, p.m), p.v]));
  const monthsOnTarget = fcostPtsOfYear.filter(p => {
    const tgt = targetMonthlyMap.get(ym(p.y, p.m));
    return tgt != null && p.v <= tgt;
  }).length;
  const kpi3Label = `${monthsOnTarget}/${fcostPtsOfYear.length}`;
  // Card 4 "CHÊNH LỆCH SO VỚI NĂM TRƯỚC": so TTL năm đang truy vấn với TTL
  // ĐÚNG năm liền trước (lastRangeYear - 1) — chọn 2026 → so 2026/2025,
  // chọn 2025 → so 2025/2024... Trước đây so bằng "12 tháng liền trước
  // startYm/endYm" (dịch theo khoảng THÁNG BẮT ĐẦU/KẾT THÚC, không hẳn là
  // "năm trước" nếu khoảng lọc không phải trọn 1 năm).
  const prevYearFcostTtl = ttlOf(modelData, ITEM_FCOST, lastRangeYear - 1);
  const kpi4DeltaPp = prevYearFcostTtl != null ? (kpi1AvgFcost - prevYearFcostTtl) * 100 : null;

  // EPCC (menu5-dead-code-cleanup) - đã xoá `radar1`/`radar2` (tính từ
  // `modelData.radarType`/`radarCustom`) — 2 biến này được tính xong nhưng
  // KHÔNG có bất kỳ chỗ nào render (`RadarChart` gọi chúng cũng đã bị xoá ở
  // trên), là code chết còn sót từ giai đoạn phát triển trước.

  // ═══════════════════════════════════════════════════════════════════
  // == 4 BIỂU ĐỒ MỚI == (thay cho 6 biểu đồ cũ đã xoá theo yêu cầu trước)
  // EPCC (menu5-chart1-chart2-link-selected-model): quy tắc CŨ ở đây từng
  // ghi "cả 4 biểu đồ đều KHÔNG phản hồi bộ lọc MODEL" — đã bị đảo ngược
  // theo yêu cầu người dùng. Quy tắc HIỆN TẠI:
  //  - Chart 1 & 2: gắn trực tiếp với `selectedModel` (bộ lọc MODEL) — chọn
  //    model nào, 2 biểu đồ này hiển thị ĐÚNG model đó, không còn cảnh đa
  //    model. Chart 1 vẫn phản hồi khoảng thời gian (THÁNG BẮT ĐẦU/KẾT
  //    THÚC) + viewBy như cũ; Target vẫn là mốc chung công ty (model lẻ
  //    không có Target riêng trong Excel).
  //  - Chart 3 & 4: GIỮ NGUYÊN dữ liệu TỔNG HỢP TOÀN BỘ MODEL (cross-model,
  //    top-15/top-9), KHÔNG phản hồi bộ lọc MODEL — vẫn đúng vai trò "nhìn
  //    toàn cảnh nhiều model" như thiết kế gốc.
  // Dùng thẳng ttlByItemYear (giá trị TTL gốc từ Excel) thay vì suy luận lại
  // từ dữ liệu tháng — xem lý do ở comment khai báo ttlByItemYear phía trên.
  // ═══════════════════════════════════════════════════════════════════

  // ── Chart 1: F-Cost($) + F-Cost(억원) + TARGET/ACTUAL % — LINK theo viewBy ──
  // EPCC (menu5-chart1-viewby-link) - FIX theo yêu cầu người dùng "cho 4
  // chart đều link giá trị thay đổi khi thay đổi tháng/quý/năm": trước đây
  // Chart 1 LUÔN cố định trục X = danh sách năm trong khoảng lọc
  // (`rangeYears`), không phản hồi nút "XEM THEO" (viewBy) ở toolbar — y hệt
  // vấn đề Chart 2 đã gặp trước đây (xem menu5-chart2-viewby-follow). Áp
  // dụng ĐÚNG pattern đã dùng cho Chart 2: tách 3 nhánh theo `viewBy`:
  //  - 'year' (giữ nguyên logic gốc): trục X = từng năm trong `rangeYears`.
  //  - 'quarter': trục X = Q1..Q4 của năm cuối (`lastRangeYear`). Bar
  //    F-cost($)/F-Cost(억원) là SỐ TIỀN nên CỘNG DỒN theo quý (khác F-Cost%
  //    ở Chart 2 vốn phải lấy trung bình); Target/Actual là % nên vẫn lấy
  //    TRUNG BÌNH các tháng có dữ liệu trong quý, đúng nguyên tắc đã áp
  //    dụng cho Chart 2.
  //  - 'month': trục X = JAN..DEC của năm cuối, lấy thẳng giá trị từng
  //    tháng (không cộng dồn/trung bình gì thêm) cho cả 4 series.
  const chart1 = useMemo(() => {
    const targetM = findSpecialModel('Target');
    // EPCC (menu5-chart1-restore-fcost-won-bar) - FIX theo yêu cầu người
    // dùng "Chart 1 (ALL YEAR) thiếu cột F-Cost (억원), bổ sung lại thành
    // biểu đồ cột": `barsWon` trước đây LUÔN trả về mảng toàn `null` ở cả 3
    // nhánh viewBy (`qLabels.map(() => null)`/`months.map(() => null)`/
    // `rangeYears.map(() => null)`) — chỉ là code placeholder, chưa từng
    // nối vào nguồn dữ liệu thật, nên cột thứ 2 + legend "F-Cost (억원)"
    // không bao giờ hiện ra (điều kiện `chart1.barsWon.some(v => v != null)`
    // luôn false). Đối chiếu Test5.xlsx xác nhận: model đặc biệt
    // 'F- Cost (억원)' có sẵn TRONG FILE, chứa đúng giá trị số tiền (đơn vị
    // 억원) theo (năm, tháng) — y hệt cấu trúc dòng 'F-cost($)'/'Target'/
    // 'Actual', chỉ khác MODEL name. Nối `barsWon` vào model này qua
    // `findSpecialModel('F- Cost (억원)')`, dùng ĐÚNG `ITEM_FCOST` (cột duy
    // nhất mà 4 model đặc biệt này có trong Excel — xem SPECIAL_FCOST_MODELS).
    // Vì là SỐ TIỀN (không phải %), quy tắc gộp theo quý là CỘNG DỒN (giống
    // bar SHIPMENT), không lấy trung bình như Target/Actual.
    const wonM = findSpecialModel('F- Cost (억원)');
    // EPCC (menu5-chart1-chart2-link-selected-model) - xem giải thích đầy đủ
    // ở comment khai báo `chart2Models` phía dưới: Chart 1 giờ cũng gắn với
    // `modelData` (model đang chọn ở bộ lọc) thay vì 2 model đặc biệt cố
    // định "F-cost($)"/"Actual". Model thật KHÔNG có cột $ riêng trong Excel
    // (chỉ 4 model đặc biệt Target/Actual/F-cost($)/F-Cost(억원) mới có) —
    // nên cột (bars) đổi từ "F-Cost($)" sang SHIPMENT (KEA) của model đang
    // chọn (số liệu thật, có ý nghĩa "sản lượng xuất hàng" đi kèm F-Cost%).
    // Đường TARGET vẫn giữ nguyên mốc chung của công ty (model lẻ không có
    // Target riêng trong Excel), đường ACTUAL đổi sang F-Cost% CHÍNH model
    // đang chọn (thay vì "Actual" tổng hợp toàn công ty).
    if (viewBy === 'quarter') {
      const monthPtsOf = (m: ModelEntry | null, item: string) => (m?.series[item] ?? []).filter(p => p.y === lastRangeYear);
      const sumByQuarterOf = (m: ModelEntry | null, item: string) => {
        const pts = monthPtsOf(m, item);
        return [1, 2, 3, 4].map(q => {
          const vals = pts.filter(p => Math.ceil(p.m / 3) === q).map(p => p.v);
          return vals.length ? vals.reduce((a, v) => a + v, 0) : null;
        });
      };
      const avgByQuarterOf = (m: ModelEntry | null, item: string) => {
        const pts = monthPtsOf(m, item);
        return [1, 2, 3, 4].map(q => {
          const vals = pts.filter(p => Math.ceil(p.m / 3) === q).map(p => p.v);
          return vals.length ? vals.reduce((a, v) => a + v, 0) / vals.length : null;
        });
      };
      const qLabels = ['Q1', 'Q2', 'Q3', 'Q4'];
      const shipQ = sumByQuarterOf(modelData, ITEM_SHIPMENT);
      const bars = qLabels.map((label, i) => ({ label, value: shipQ[i] }));
      const barsWon = sumByQuarterOf(wonM, ITEM_FCOST);
      const target = avgByQuarterOf(targetM, ITEM_FCOST);
      const actual = avgByQuarterOf(modelData, ITEM_FCOST);
      return { bars, barsWon, target, actual };
    }
    if (viewBy === 'month') {
      const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const byMonthOf = (m: ModelEntry | null, item: string) => {
        const map = new Map((m?.series[item] ?? []).filter(p => p.y === lastRangeYear).map(p => [p.m, p.v]));
        return months.map((_lab, i) => map.get(i + 1) ?? null);
      };
      const shipVals = byMonthOf(modelData, ITEM_SHIPMENT);
      const bars = months.map((label, i) => ({ label, value: shipVals[i] }));
      const barsWon = byMonthOf(wonM, ITEM_FCOST);
      const target = byMonthOf(targetM, ITEM_FCOST);
      const actual = byMonthOf(modelData, ITEM_FCOST);
      return { bars, barsWon, target, actual };
    }
    // viewBy === 'year'.
    const bars = rangeYears.map(y => ({ label: String(y), value: ttlOf(modelData, ITEM_SHIPMENT, y) }));
    const barsWon = rangeYears.map(y => ttlOf(wonM, ITEM_FCOST, y));
    const target = rangeYears.map(y => ttlOf(targetM, ITEM_FCOST, y));
    const actual = rangeYears.map(y => ttlOf(modelData, ITEM_FCOST, y));
    return { bars, barsWon, target, actual };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [modelData, allModels, rangeYears, lastRangeYear, viewBy]);

  // ── Danh sách real model (loại 4 model đặc biệt) dùng chung cho chart 2/3/4 ──
  const realModels = useMemo(
    () => allModels.filter(m => !SPECIAL_FCOST_MODELS.includes(m.model)),
    [allModels]
  );

  // ── Top-9 model theo F-Cost% TTL của năm cuối trong khoảng đã lọc ───────
  const top9Models = useMemo(() => {
    return realModels
      .map(m => ({ model: m.model, v: ttlOf(m, ITEM_FCOST, lastRangeYear) }))
      .filter(x => x.v != null)
      .sort((a, b) => (b.v as number) - (a.v as number))
      .slice(0, 9)
      .map(x => x.model);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [realModels, lastRangeYear]);

  // EPCC (menu5-chart1-chart2-link-selected-model) - FIX theo yêu cầu người
  // dùng "nút model truy vấn liên kết với biểu đồ 1,2, khi chọn 1 model thì
  // biểu đồ hiển thị rõ ràng model đó": trước đây (comment "== 4 BIỂU ĐỒ
  // MỚI ==" phía trên) Chart 1/2 CHỦ ĐÍCH không phản hồi bộ lọc MODEL — quy
  // tắc đó đã được xác nhận LẠI với người dùng và đảo ngược: giờ Chart 1 & 2
  // LUÔN gắn với `selectedModel` (không còn chế độ "toàn cảnh nhiều model"),
  // còn Chart 3/4 vẫn giữ nguyên top-N/top-15 đa model như cũ (không đổi).
  // EPCC (menu5-model-all-chart2-show-all-models) - FIX theo yêu cầu người
  // dùng "chọn 'Tất cả' ở MODEL và NĂM thì Chart 2 phải hiện ra HẾT các
  // model, không phải 1 đường tổng hợp duy nhất": trước đây khi chọn "Tất
  // cả", `chart2Models` trả về `[modelData.model]` — do `modelData` lúc đó
  // là object tổng hợp có `model: 'Actual'` (xem `allModelsAggregate` phía
  // trên) nên Chart 2 chỉ vẽ ĐÚNG 1 đường "Tất cả" (đường màu hồng trong
  // ảnh người dùng chụp) thay vì so sánh nhiều model như tên tiêu đề "F-COST
  // THEO MODEL" ngụ ý. Sửa: khi chọn "Tất cả" → trả về TOÀN BỘ real model
  // (`realModels`, đã loại 4 model đặc biệt) để Chart 2 vẽ đủ từng model
  // riêng lẻ; khi chọn 1 model cụ thể → vẫn giữ nguyên hành vi cũ, chỉ 1
  // đường của đúng model đó. Nhờ đó nhãn legend (`label`) giờ dùng lại TRỰC
  // TIẾP tên model thật (bỏ `chart2ModelLabel` không còn cần nữa), vì mỗi
  // phần tử trong `chart2Models` luôn là 1 tên model có thật trong Excel.
  const chart2Models = useMemo(() => {
    if (selectedModel === ALL_MODEL_VALUE) {
      return realModels.map(m => m.model);
    }
    return modelData ? [modelData.model] : [];
  }, [selectedModel, realModels, modelData]);

  // ── Chart 2: F-Cost% theo TTL + Tháng/Quý/Năm (tuỳ `viewBy`), đúng model đang chọn ──
  // EPCC (menu5-chart2-viewby-follow) - FIX theo yêu cầu người dùng "lấy
  // F-Cost% theo từng model mà khi lựa chọn tháng/quý/năm ăn theo như vậy":
  // trước đây Chart 2 LUÔN cố định trục X = TTL + 12 tháng của năm cuối
  // (`lastRangeYear`), bất kể người dùng đã bấm nút "XEM THEO" (viewBy) =
  // Tháng/Quý/Năm ở toolbar — nút đó trước đây chỉ ảnh hưởng tới nhóm biểu
  // đồ Shipment/Prod'n/Sales/Scrap (qua `groupPoints()`), không đụng tới
  // Chart 2. Giờ tách 3 nhánh theo `viewBy`:
  //  - 'month' (giữ nguyên logic cũ): trục X = TTL + JAN..DEC của năm cuối
  //    trong khoảng lọc, TTL lấy thẳng từ `ttlByItemYear` (Excel).
  //  - 'quarter': trục X = TTL + Q1..Q4 của năm cuối; mỗi quý = TRUNG BÌNH
  //    các tháng có dữ liệu trong quý đó (F-Cost% là TỶ LỆ %, không được
  //    cộng dồn như Q'TY — cộng dồn sẽ ra số vô nghĩa, VD 3 tháng 5% cộng
  //    lại thành "15%" sai bản chất).
  //  - 'year': trục X = TTL + từng năm trong khoảng lọc (`rangeYears`, có
  //    thể nhiều năm); mỗi năm lấy thẳng TTL năm đó từ `ttlByItemYear`
  //    (đã đúng bản chất %, không cần tính lại), TTL chung = trung bình các
  //    năm có dữ liệu.
  // EPCC (menu5-model-all-chart2-hide-no-data-models) - FIX theo yêu cầu
  // người dùng "hiển thị quá nhiều model tại Chart 2, chỉ hiện model có giá
  // trị tại vùng truy vấn, còn lại ẩn đi": bản fix trước
  // (menu5-model-all-chart2-show-all-models) vẽ TOÀN BỘ real model khi chọn
  // "Tất cả", kể cả model KHÔNG có bất kỳ điểm F-Cost% nào trong khoảng
  // Tháng/Quý/Năm đang lọc — khiến legend dài tràn cả header và biểu đồ có
  // hàng chục đường phẳng lì/rỗng gây rối mắt (ảnh người dùng chụp). Sửa:
  // sau khi tính xong `values` cho từng model ở CẢ 3 nhánh viewBy, LỌC bỏ
  // model nào không có giá trị khác null ở BẤT KỲ cột nào TRONG VÙNG TRUY
  // VẤN (bỏ qua cột "TTL" — đó chỉ là số tổng hợp phụ, không phải 1 kỳ
  // trong vùng lọc) trước khi gán màu, rồi mới gán lại `color` theo CHỈ SỐ
  // SAU KHI LỌC (để không bị nhảy cóc màu do các model bị ẩn ở giữa danh
  // sách gốc).
  // EPCC (menu5-chart256-top10-by-fcost) - FIX theo yêu cầu người dùng "Chart
  // 2,3,4 (F-Cost%/Mass F-Cost%/SPL F-Cost% THEO MODEL) chỉ giới hạn 10 model
  // có F-Cost cao đến thấp thôi vì nhiều quá nhìn không rõ": trước đây khi
  // chọn "Tất cả" ở bộ lọc MODEL, cả 3 chart này vẽ HẾT mọi model có dữ liệu
  // trong vùng lọc (có thể hơn chục đường chồng lên nhau, không đọc được —
  // đúng như ảnh người dùng chụp). Sau khi lọc model không có dữ liệu
  // (`filter` cũ), giờ SẮP XẾP GIẢM DẦN theo giá trị TTL (`values[0]`, cột
  // tổng hợp của cả vùng lọc — đúng nghĩa "F-Cost cao đến thấp") rồi CHỈ LẤY
  // 10 model đầu (`CHART_TOP_N_MODELS`) trước khi gán màu. TTL=null bị đẩy
  // xuống cuối (coi như -Infinity) thay vì gây lỗi so sánh. Khi chỉ có 1
  // model (đã chọn cụ thể ở dropdown MODEL) thì slice(0, 10) không ảnh hưởng
  // gì — vẫn giữ nguyên hành vi cũ.
  // EPCC (menu5-chart256-recent-priority-top10) - FIX ROOT CAUSE "chọn
  // MODEL=Tất cả + NĂM=Tất cả, biểu đồ Mass F-Cost% (Chart 3, vị trí thứ 3)
  // không hiển thị dữ liệu ở 2025/2026 dù trục X vẫn hiện đủ nhãn 2 năm đó":
  // categories (trục X) được tính từ TOÀN BỘ model thật TRƯỚC KHI lọc Top
  // 10 (xem `perModelYearVals`/`yearsWithData` ở buildFcostByModelChart) —
  // nên chỉ cần 1 model bất kỳ (kể cả model sắp bị loại khỏi Top 10) có dữ
  // liệu ở 2025/2026 là nhãn năm đó vẫn hiện. Trong khi đó, hàm này (chạy
  // SAU, quyết định model nào thực sự được VẼ) trước đây xếp hạng Top 10
  // theo `values[0]` = TRUNG BÌNH F-Cost% CẢ GIAI ĐOẠN đang lọc — model có
  // Mass F-Cost% đột biến cao ở các năm CŨ (VD SO1C21 = 103% NHƯNG DUY NHẤT
  // Ở NĂM 2020, không hề có dữ liệu năm nào khác) vẫn kéo trung bình lên rất
  // cao, luôn giữ suất Top 10; ngược lại các model ĐANG hoạt động thật ở
  // 2025/2026 (Mass F-Cost% chỉ quanh 1-3%) bị đẩy ra khỏi Top 10 vì thấp
  // hơn nhiều. ĐÃ ĐỐI CHIẾU TRỰC TIẾP với Test5.xlsx gốc (pivot MODEL x YEAR
  // của dòng MONTH=TTL, ITEM='Mass F-Cost %') — xác nhận ĐÚNG 10 model xuất
  // hiện trong legend ảnh chụp người dùng gửi (SO1C21, SI2900, SO1G70/M,
  // SO1A00, SO1940, SO1E60, SO1E60H, SO1961, SO1923G, SO1G71) chính là 10
  // model có TRUNG BÌNH cao nhất theo công thức cũ — và ĐÚNG LÀ không model
  // nào trong 10 model đó có dòng dữ liệu ở 2025 hoặc 2026, trong khi Excel
  // gốc CÓ 13 model khác với dữ liệu Mass F-Cost% 2025/2026 (SO1B01,
  // SO1C2EF, SO1C2EH, SO1C2G... quanh 1-3%) nhưng bị loại hết khỏi Top 10.
  // Kết quả: trục X vẫn hiện nhãn 2025/2026 nhưng KHÔNG model nào trong danh
  // sách được vẽ có điểm dữ liệu ở đó → nhìn như "mất dữ liệu 2025/2026".
  // LẦN SỬA ĐẦU (đã revert) chỉ đổi tiêu chí sang "giá trị tại kỳ CUỐI CÙNG
  // có dữ liệu của mỗi model" — SAI, vì với SO1C21 kỳ cuối cùng có dữ liệu
  // VẪN LÀ năm 2020 (không có năm nào sau đó), giá trị tại đó vẫn là 103% →
  // vẫn thắng và vẫn đứng Top 1, không giải quyết được gì (đã tự kiểm chứng
  // lại bằng script đối chiếu Excel, thấy top-10 không đổi). Fix ĐÚNG: phải
  // ưu tiên theo **NĂM/KỲ của điểm dữ liệu cuối cùng** (recency) LÀM TIÊU CHÍ
  // CHÍNH — model nào có dữ liệu ở kỳ CÀNG GẦN kỳ cuối trong vùng lọc càng
  // được ưu tiên xếp trên, bất kể giá trị cao/thấp — chỉ dùng giá trị tại kỳ
  // đó làm tiêu chí PHỤ để phá thế hoà giữa các model có cùng kỳ gần nhất.
  // Áp dụng NHẤT QUÁN cho cả Chart 2/5/6 (dùng chung hàm này).
  function filterModelsWithDataInRange<T extends { values: (number | null)[] }>(list: T[]): (T & { color: string })[] {
    // periodVals bỏ cột TTL ở đầu (values[0]), chỉ xét từng kỳ thật theo
    // đúng thứ tự thời gian TĂNG DẦN (đã đúng ở cả 3 nhánh year/quarter/
    // month của buildFcostByModelChart) — recencyIndex = vị trí kỳ cuối
    // cùng CÓ dữ liệu (càng lớn = càng gần hiện tại), recentValue = giá trị
    // tại đúng kỳ đó.
    const recencyOf = (s: T): { recencyIndex: number; recentValue: number } => {
      const periodVals = s.values.slice(1);
      for (let i = periodVals.length - 1; i >= 0; i--) {
        const v = periodVals[i];
        if (v != null) return { recencyIndex: i, recentValue: v };
      }
      return { recencyIndex: -1, recentValue: -Infinity };
    };
    return list
      .filter(s => s.values.slice(1).some(v => v != null))
      .sort((a, b) => {
        const ra = recencyOf(a), rb = recencyOf(b);
        if (rb.recencyIndex !== ra.recencyIndex) return rb.recencyIndex - ra.recencyIndex;
        return rb.recentValue - ra.recentValue;
      })
      .slice(0, CHART_TOP_N_MODELS)
      .map((s, idx) => ({ ...s, color: CHART_PALETTE_9[idx % CHART_PALETTE_9.length] }));
  }

  // EPCC (menu5-chart5-6-mass-spl-fcost) - REFACTOR: khối tính Chart 2 trước
  // đây hardcode thẳng `ITEM_FCOST`, không thể tái dùng cho yêu cầu mới
  // "thêm 2 biểu đồ Mass F-Cost% / SPL F-Cost%, y hệt Chart 2, làm biểu đồ
  // 5,6, có link (ăn theo model + XEM THEO Tháng/Quý/Năm giống Chart 2)".
  // Tách nguyên logic 3 nhánh year/quarter/month thành hàm dùng chung
  // `buildFcostByModelChart(item)`, tham số hoá theo tên item — giữ 100%
  // hành vi cũ (cùng `chart2Models`/`lastRangeYear`/`rangeYears`/`viewBy`,
  // cùng `filterModelsWithDataInRange` để ẩn model không có dữ liệu trong
  // vùng lọc, cùng bảng màu `CHART_PALETTE_9`) — chỉ đổi item đang truy vấn.
  function buildFcostByModelChart(item: string) {
    // EPCC (menu5-chart456-auto-x-axis) - FIX theo yêu cầu người dùng "để
    // trục X auto theo số liệu khi truy vấn không để dạng cố định nữa" (áp
    // dụng cho Chart 4 hiển thị ở vị trí thứ 4 = SPL F-Cost%, đang dùng
    // chung hàm này với Chart 2/5): trước đây nhánh 'year' LUÔN đưa ĐỦ mọi
    // năm trong `rangeYears` (khoảng lọc THÁNG BẮT ĐẦU/KẾT THÚC) vào trục X,
    // và nhánh 'quarter' LUÔN cố định đủ Q1..Q4 — bất kể model có dữ liệu
    // thật ở năm/quý đó hay không cho ĐÚNG `item` đang vẽ (VD SPL F-Cost%
    // chỉ có dữ liệu 2 quý, 2 quý còn lại vẫn bị vẽ trống). Áp dụng lại
    // NGUYÊN TẮC đã dùng ở nhánh 'month' (menu5-chart2-hide-empty-months):
    // tính trước dữ liệu thô từng model, lọc categories chỉ giữ năm/quý có
    // ÍT NHẤT 1 model có giá trị khác null, "TTL" luôn giữ nguyên ở đầu. Vì
    // Chart 2 (F-Cost%)/Chart 5 (Mass)/Chart 6 (SPL) đều gọi hàm này, fix áp
    // dụng nhất quán cho cả 3, không lệch hành vi giữa các chart giống nhau.
    if (viewBy === 'year') {
      const perModelYearVals = chart2Models.map(modelName => {
        const m = allModels.find(mm => mm.model === modelName) ?? null;
        return rangeYears.map(y => ttlOf(m, item, y));
      });
      const yearsWithData = rangeYears.filter((_, yi) => perModelYearVals.some(vals => vals[yi] != null));
      const categories = ['TTL', ...yearsWithData.map(String)];
      const rawSeries = chart2Models.map((modelName, idx) => {
        const allYearVals = perModelYearVals[idx];
        const yearVals = rangeYears
          .map((y, yi) => ({ y, v: allYearVals[yi] }))
          .filter(({ y }) => yearsWithData.includes(y))
          .map(({ v }) => v);
        const validVals = yearVals.filter((v): v is number => v != null);
        const ttlVal = validVals.length ? validVals.reduce((a, v) => a + v, 0) / validVals.length : null;
        return { label: modelName, values: [ttlVal, ...yearVals] };
      });
      return { categories, series: filterModelsWithDataInRange(rawSeries) };
    }
    if (viewBy === 'quarter') {
      const monthPtsPerModel = chart2Models.map(modelName => {
        const m = allModels.find(mm => mm.model === modelName) ?? null;
        return (m?.series[item] ?? []).filter(p => p.y === lastRangeYear);
      });
      const quarters = [1, 2, 3, 4].filter(q =>
        monthPtsPerModel.some(pts => pts.some(p => Math.ceil(p.m / 3) === q)));
      const categories = ['TTL', ...quarters.map(q => `Q${q}`)];
      const rawSeries = chart2Models.map((modelName, idx) => {
        const m = allModels.find(mm => mm.model === modelName) ?? null;
        const monthPts = monthPtsPerModel[idx];
        const qValues = quarters.map(q => {
          const qVals = monthPts.filter(p => Math.ceil(p.m / 3) === q).map(p => p.v);
          return qVals.length ? qVals.reduce((a, v) => a + v, 0) / qVals.length : null;
        });
        const ttlVal = ttlOf(m, item, lastRangeYear);
        return { label: modelName, values: [ttlVal, ...qValues] };
      });
      return { categories, series: filterModelsWithDataInRange(rawSeries) };
    }
    // viewBy === 'month'.
    // EPCC (menu5-chart2-hide-empty-months) - FIX theo yêu cầu người dùng
    // "khi truy vấn tháng nếu có tháng chưa có dữ liệu hãy để ẩn": trước đây
    // trục X LUÔN cố định đủ 12 tháng JAN..DEC bất kể model nào cũng không
    // có dữ liệu ở tháng đó (VD năm chưa đi hết 12 tháng, hoặc tháng bị
    // thiếu trong Excel) — khiến các tháng trống hiện ra như đường đứt gãy/
    // khoảng trắng vô nghĩa giữa biểu đồ. Giờ tính "byMonth" cho TỪNG model
    // trước, rồi lọc `monthKeys` chỉ giữ lại tháng nào có ÍT NHẤT 1 model có
    // giá trị khác null — tháng nào KHÔNG có dữ liệu thì loại thẳng khỏi
    // mảng, không đưa vào categories/values nữa (khác với trước: giữ tháng
    // nhưng vẽ null). "TTL" luôn giữ nguyên ở đầu.
    const monthKeysAll = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    const monthLabels = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    const byMonthPerModel = chart2Models.map(modelName => {
      const m = allModels.find(mm => mm.model === modelName) ?? null;
      const monthPts = (m?.series[item] ?? []).filter(p => p.y === lastRangeYear);
      return new Map(monthPts.map(p => [p.m, p.v]));
    });
    const monthKeys = monthKeysAll.filter(mk => byMonthPerModel.some(bm => bm.get(mk) != null));
    const categories = ['TTL', ...monthKeys.map(mk => monthLabels[mk - 1])];
    const rawSeries = chart2Models.map((modelName, idx) => {
      const m = allModels.find(mm => mm.model === modelName) ?? null;
      const byMonth = byMonthPerModel[idx];
      const values = [ttlOf(m, item, lastRangeYear), ...monthKeys.map(mk => byMonth.get(mk) ?? null)];
      return { label: modelName, values };
    });
    return { categories, series: filterModelsWithDataInRange(rawSeries) };
  }

  const chart2 = useMemo(() => buildFcostByModelChart(ITEM_FCOST),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [chart2Models, allModels, lastRangeYear, rangeYears, viewBy]);

  // EPCC (menu5-chart5-6-mass-spl-fcost) - Chart 5 (Mass F-Cost%) & Chart 6
  // (SPL F-Cost%) THÊM MỚI theo yêu cầu người dùng "thêm 2 biểu đồ y hệt
  // Chart 2, link theo model/viewBy giống Chart 2, làm biểu đồ 5,6": tái
  // dùng đúng `buildFcostByModelChart()` ở trên, chỉ đổi item truy vấn —
  // cùng deps với chart2 nên tự động "link" theo đúng model đang chọn
  // (`chart2Models`) và chế độ Tháng/Quý/Năm (`viewBy`) giống Chart 2.
  const chart5 = useMemo(() => buildFcostByModelChart(ITEM_MASS_FCOST),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [chart2Models, allModels, lastRangeYear, rangeYears, viewBy]);

  const chart6 = useMemo(() => buildFcostByModelChart(ITEM_SPL_FCOST),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [chart2Models, allModels, lastRangeYear, rangeYears, viewBy]);

  // ── Chart 3: SCRAP cộng dồn cả khoảng đã lọc, cho TẤT CẢ real model ─────
  // (thứ tự đối xứng: giá trị lớn nhất ở đỉnh 12h, nửa lớn còn lại xếp theo
  // chiều kim đồng hồ giảm dần, nửa nhỏ xếp ngược chiều kim đồng hồ tăng dần
  // — đúng dáng "sao chổi" trong hình mẫu.)
  function symmetricOrder<T>(sortedDesc: T[]): T[] {
    if (sortedDesc.length === 0) return [];
    const [first, ...rest] = sortedDesc;
    const rightCount = Math.ceil(rest.length / 2);
    const right = rest.slice(0, rightCount);              // giảm dần, đặt bên phải (theo chiều kim đồng hồ)
    const left = rest.slice(rightCount).reverse();         // tăng dần khi đi ra xa đỉnh, đặt bên trái
    return [first, ...right, ...left];
  }

  // EPCC (menu5-rose-chart-limit-15) - FIX ROOT CAUSE "vòng hoa Chart 3/4 quá
  // rối, nhãn model chồng chéo không đọc được": trước đây chart3Items đưa
  // TẤT CẢ real model (kể cả hàng chục model không có dữ liệu Scrap) vào vòng
  // tròn, khiến mỗi lát quá mỏng và nhãn đè lên nhau. Giới hạn còn top-15 model
  // theo giá trị Scrap cao nhất (đủ để nhìn rõ từng lát + nhãn), bỏ hẳn các
  // model không có dữ liệu khỏi vòng này thay vì vẽ mờ "#N/A" cho toàn bộ.
  const CHART3_MAX_MODELS = 15;

  // EPCC (menu5-chart3-all-years-independent) - FIX theo yêu cầu người dùng
  // "Tại Chart 5 này (F-COST(K$) LŨY KẾ THEO NĂM) hãy để số lượng lũy kế theo
  // TẤT CẢ CÁC NĂM, không ảnh hưởng việc lựa chọn model và năm": trước đây
  // (menu5-chart3-viewby-link) chart này CỘNG DỒN theo `rangeYears`/
  // `lastRangeYear` — tức bị SIẾT LẠI mỗi khi người dùng đổi bộ lọc "TỪ NĂM/
  // ĐẾN NĂM" hoặc "XEM THEO" (Tháng/Quý/Năm) ở toolbar, đúng như hiện tượng
  // người dùng chụp ảnh phản ánh. Model thì vốn ĐÃ độc lập từ trước (dùng
  // thẳng `realModels` — TOÀN BỘ model thật, không lọc theo `selectedModel`),
  // chỉ có phần NĂM là chưa đúng. Đổi sang cộng dồn qua `allDataYears` (TOÀN
  // BỘ năm có mặt trong dữ liệu đã tải lên, xem menu5-year-quick-filter) —
  // cố định, không còn phụ thuộc `rangeYears`/`lastRangeYear`/`viewBy` nữa.
  const chart3Items = useMemo(() => {
    const withVal = realModels.map(m => {
      const hasAny = allDataYears.some(y => ttlOf(m, ITEM_SCRAP, y) != null);
      if (!hasAny) return { label: m.model, value: null as number | null };
      const sum = allDataYears.reduce((a, y) => a + (ttlOf(m, ITEM_SCRAP, y) ?? 0), 0);
      return { label: m.model, value: Math.round(sum * 10) / 10 };
    });
    const withData = withVal
      .filter(x => x.value != null)
      .sort((a, b) => (b.value as number) - (a.value as number))
      .slice(0, CHART3_MAX_MODELS);
    return symmetricOrder(withData);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [realModels, allDataYears]);

  // EPCC (menu5-rose-chart4-only-data) - FIX ROOT CAUSE "vòng Chart 4 vẫn vẽ
  // đầy #N/A dù chỉ có vài model có số liệu thật": trước đây chart4Items lấy
  // NGUYÊN danh sách 15 category của Chart 3 làm khung, model nào không thuộc
  // top9Models thì gán null → RoseChart vẫn vẽ đủ 15 lát (đa số "#N/A"), y hệt
  // vấn đề đã sửa ở Chart 3. Theo yêu cầu "chỉ hiển thị model có dữ liệu số",
  // đổi sang tính ĐỘC LẬP trực tiếp từ top9Models: chỉ giữ model có giá trị
  // Scrap thật của năm cuối (bỏ hẳn model không có dữ liệu), không còn ăn theo
  // category ring của Chart 3 nữa.
  // (Lịch sử: Chart 4 từng trải qua 1 giai đoạn "menu5-chart4-viewby-link"
  // tính lũy kế tới tháng/quý gần nhất theo viewBy — đã bị thay thế hoàn
  // toàn bởi fix "menu5-chart4-per-year-ascending" bên dưới, không còn áp
  // dụng nữa.)
  // EPCC (menu5-chart4-per-year-ascending) - FIX theo yêu cầu người dùng
  // "không muốn lấy dữ liệu gần nhất nữa mà thành dữ liệu của từng năm, số
  // lượng từng model xếp từ thấp đến cao": trước đây (menu5-chart4-viewby-
  // link + menu5-chart4-month-cumulative) Chart 4 có 2 nhánh phức tạp theo
  // viewBy — cộng dồn tới THÁNG/QUÝ GẦN NHẤT có dữ liệu trong năm cuối. Bỏ
  // hẳn khái niệm "gần nhất" đó: giờ LUÔN lấy đúng TTL Scrap của NGUYÊN NĂM
  // đang chọn (`lastRangeYear`), không còn phân biệt theo viewBy (tháng/quý/
  // năm cho cùng 1 kết quả). Đồng thời đổi chiều sắp xếp: trước đây cao→thấp
  // (giá trị lớn nhất ở đỉnh 12h theo dáng "sao chổi"), nay đổi thành
  // THẤP→CAO trước khi đưa vào symmetricOrder() — nghĩa là model có Scrap
  // ÍT NHẤT sẽ nằm ở đỉnh 12h, lớn dần khi toả ra 2 bên.
  const chart4Items = useMemo(() => {
    const withVal = top9Models
      .map(modelName => {
        const m = realModels.find(mm => mm.model === modelName) ?? null;
        const value = m ? ttlOf(m, ITEM_SCRAP, lastRangeYear) : null;
        return { label: modelName, value };
      })
      .filter((x): x is { label: string; value: number } => x.value != null)
      .sort((a, b) => a.value - b.value);
    return symmetricOrder(withVal);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [top9Models, realModels, lastRangeYear]);

  // ═══════════════════════════════════════════════════════════════════════
  // EPCC (menu5-revenue-tab) - Tab MỚI "TÌNH HÌNH DOANH THU": nhân bản bố
  // cục/màu sắc 4 khung biểu đồ + 4 thẻ KPI + bảng của SalesDashboard.tsx
  // (Mục 1), nhưng dữ liệu lấy thẳng từ Test5.xlsx (đã tải sẵn ở Mục 5, KHÔNG
  // đọc lại file) — dùng `allModelsAggregate` (SHIPMENT/PROD'N/SALES AMT cộng
  // dồn MỌI real model theo năm) + `salesByCustomer` (SALES AMT theo khách
  // hàng, mới thêm ở accumulateSheetRows) làm nguồn. Tab này CHỦ Ý không phụ
  // thuộc bộ lọc MODEL/XEM THEO ở toolbar (vốn phục vụ Chart 1/2 của tab F-
  // Cost) — đây là góc nhìn TOÀN CÔNG TY theo toàn bộ năm có dữ liệu, giống
  // đúng "TỪ NĂM 2017 ĐẾN NĂM 2026 / MODEL: Tất cả" trong ảnh tham chiếu.
  // ═══════════════════════════════════════════════════════════════════════

  // Toàn bộ năm có ít nhất 1 trong 3 chỉ số SHIPMENT/PROD'N/SALES AMT.
  const revenueYears = useMemo(() => {
    const ys = new Set<number>();
    for (const item of [ITEM_SHIPMENT, ITEM_PRODUCTION, ITEM_SALES]) {
      for (const yStr of Object.keys(allModelsAggregate?.ttlByItemYear[item] ?? {})) ys.add(Number(yStr));
    }
    return Array.from(ys).sort((a, b) => a - b);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allModelsAggregate]);


  // 4 thẻ KPI: TỔNG DOANH SỐ / TỔNG XUẤT HÀNG / TỔNG SẢN XUẤT / TĂNG TRƯỞNG
  // DOANH SỐ (YoY).
  // EPCC (menu5-revenue-kpi-link-year-filter) - FIX theo yêu cầu người dùng
  // "khi lựa chọn NĂM và XEM THEO thì 4 Card KPI cũng phải thay đổi theo điều
  // kiện": trước đây `sumAll` LUÔN cộng dồn TOÀN BỘ `revenueYears` (2017→
  // 2026) bất kể người dùng đã đổi bộ lọc "NĂM"/"THÁNG BẮT ĐẦU"/"THÁNG KẾT
  // THÚC" ở toolbar phía trên — 4 card này đứng yên một chỗ trong khi mọi
  // phần khác của trang (Chart 1-4, bảng...) đều phản hồi bộ lọc, gây cảm
  // giác "bấm nút không có tác dụng". Fix: đổi nguồn tính sang cộng dồn
  // TRỰC TIẾP từ dữ liệu THÁNG nằm trong đúng khoảng đang chọn (`startYm`→
  // `endYm` — 2 giá trị này CHÍNH LÀ nơi cả dropdown "NĂM" lẫn 2 ô "THÁNG BẮT
  // ĐẦU/KẾT THÚC" cùng ghi vào, xem `handleYearFilterChange` phía trên), thay
  // vì cộng cứng toàn bộ năm. Nhờ vậy:
  //  - Chọn NĂM = 1 năm cụ thể (VD 2026) → card chỉ còn tính riêng năm đó.
  //  - Chọn NĂM = "Tất cả" (hoặc tự kéo THÁNG BẮT ĐẦU/KẾT THÚC) → card cộng
  //    đúng khoảng đang hiển thị, không còn cố định 2017-2026 nữa.
  // "XEM THEO" (viewBy Tháng/Quý/Năm) CHỦ Ý không tham gia vào phép cộng này
  // — nút đó chỉ quyết định cách CHIA NHÓM trục X của các biểu đồ bên dưới
  // (gộp theo tháng/quý/năm), còn 4 card này luôn là 1 con số TỔNG duy nhất
  // của khoảng đang chọn nên không có khái niệm "chia theo tháng/quý" ở đây
  // — đổi "XEM THEO" vẫn tự nhiên giữ nguyên số card (đúng vì tổng của cùng
  // 1 khoảng thời gian không đổi dù nhìn theo tháng hay theo quý), chỉ có 4
  // biểu đồ chi tiết bên dưới mới đổi hình dạng theo tab này.
  const revenueKpis = useMemo(() => {
    const sumInSelectedRange = (item: string) => {
      const pts = allModelsAggregate?.series[item] ?? [];
      return pts.reduce((a, p) => (ym(p.y, p.m) >= startYm && ym(p.y, p.m) <= endYm ? a + p.v : a), 0);
    };
    const totalSales = sumInSelectedRange(ITEM_SALES);
    const totalShipment = sumInSelectedRange(ITEM_SHIPMENT);
    const totalProduction = sumInSelectedRange(ITEM_PRODUCTION);

    let yoyPct: number | null = null;
    let yoyLabel = '';
    if (selectedYearFilterValue === 'all') {
      // Giữ nguyên hành vi cũ khi NĂM = "Tất cả": so sánh 2 năm GẦN NHẤT có
      // dữ liệu (không lùi theo khoảng đang chọn, vì khoảng lúc này là toàn
      // bộ lịch sử, lùi thêm 1 năm sẽ ra ngoài biên dữ liệu).
      if (revenueYears.length >= 2) {
        const yCurr = revenueYears[revenueYears.length - 1];
        const yPrev = revenueYears[revenueYears.length - 2];
        const vCurr = allModelsAggregate?.ttlByItemYear[ITEM_SALES]?.[yCurr];
        const vPrev = allModelsAggregate?.ttlByItemYear[ITEM_SALES]?.[yPrev];
        if (vCurr != null && vPrev) {
          yoyPct = ((vCurr - vPrev) / vPrev) * 100;
          yoyLabel = `${yPrev} ${t.revYoyVs} ${yCurr}`;
        }
      }
    } else {
      // NĂM = 1 năm cụ thể (hoặc khoảng THÁNG BẮT ĐẦU/KẾT THÚC tuỳ ý): so
      // với ĐÚNG cùng khoảng đó nhưng lùi lại 12 tháng (1 năm), tổng quát
      // hoá cho mọi độ dài khoảng thay vì chỉ 1 năm trọn vẹn.
      const prevStartYm = startYm - 12;
      const prevEndYm = endYm - 12;
      if (prevStartYm >= globalBounds.minYm) {
        const pts = allModelsAggregate?.series[ITEM_SALES] ?? [];
        const vPrev = pts.reduce((a, p) => (ym(p.y, p.m) >= prevStartYm && ym(p.y, p.m) <= prevEndYm ? a + p.v : a), 0);
        if (vPrev) {
          yoyPct = ((totalSales - vPrev) / vPrev) * 100;
          const yCurrLabel = startYear === endYear ? `${endYear}` : `${startYear}-${endYear}`;
          const yPrevLabel = startYear === endYear ? `${endYear - 1}` : `${startYear - 1}-${endYear - 1}`;
          yoyLabel = `${yPrevLabel} ${t.revYoyVs} ${yCurrLabel}`;
        }
      }
    }
    return { totalSales, totalShipment, totalProduction, yoyPct, yoyLabel };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allModelsAggregate, startYm, endYm, startYear, endYear, selectedYearFilterValue, revenueYears, globalBounds, t.revYoyVs]);

  // Chart 1 (Xuất hàng & Doanh số theo Năm): bar=SHIPMENT + PROD'N, line=SALES AMT
  // (cùng trục trái với bar), line đứt=Tăng trưởng YoY doanh số (trục phải %).
  // EPCC (menu5-revchart1-add-prodn) - FIX theo yêu cầu người dùng "thêm
  // PROD'N (KEA) vào biểu đồ 1", đối chiếu ảnh tham chiếu (2 cột SHIPMENT +
  // PROD'N cạnh nhau, cùng đơn vị K, cùng trục trái với SALES AMT): thêm
  // `prodVals` lấy TTL SẢN XUẤT (`ITEM_PRODUCTION`) toàn công ty theo đúng
  // năm, y hệt cách `shipVals` đã lấy — dùng chung `allModelsAggregate`
  // (không đụng bucket/model nào khác).
  // EPCC (menu5-revchart1-link-model-year-viewby) - FIX theo yêu cầu người
  // dùng "liên kết thanh lựa chọn model/năm sẽ thay đổi theo đúng giá trị
  // của model và năm đó... khi thay đổi năm thì trục dưới sẽ thành tháng,
  // tức có thể thay đổi theo cả tính năng lựa chọn tháng/quý/năm nữa":
  // trước đây Chart 1 CHỦ Ý cộng dồn `allModelsAggregate` trên TOÀN BỘ năm
  // có dữ liệu (`revenueYears`, luôn 2017→2026), hoàn toàn không phản hồi
  // bộ lọc MODEL/NĂM/XEM THEO ở toolbar (xem comment cũ ở khối "Tab MỚI"
  // phía trên — CHỦ Ý làm vậy vì lúc đó tab này muốn là góc nhìn "toàn công
  // ty, mọi năm" cố định). Người dùng giờ yêu cầu ngược lại: Chart 1 phải
  // ăn theo ĐÚNG model đang chọn + đúng khoảng năm/tháng đang lọc, y hệt
  // cách `chart1` (2 biểu đồ Tháng bắt đầu/kết thúc F-COST phía trên) đã
  // làm. Áp dụng lại NGUYÊN VẸN pattern đó cho Chart 1 của tab DOANH THU:
  //  - Nguồn dữ liệu: `modelData` (đã tự ứng với `selectedModel` — "Tất cả"
  //    → `allModelsAggregate`, 1 model cụ thể → đúng model đó) thay cho
  //    `allModelsAggregate` cố định.
  //  - Khoảng năm: `rangeYears`/`lastRangeYear` (đã tự ứng với bộ lọc THÁNG
  //    BẮT ĐẦU/KẾT THÚC + ô chọn nhanh NĂM ở toolbar) thay cho `revenueYears`
  //    cố định toàn bộ dữ liệu.
  //  - 3 nhánh theo `viewBy` (year/quarter/month), giống hệt `chart1`: chọn
  //    NĂM cụ thể + "Xem theo Tháng/Quý" sẽ đổi hẳn trục X từ danh sách năm
  //    sang Q1..Q4 hoặc JAN..DEC của đúng năm đó.
  const revenueChart1 = useMemo(() => {
    // EPCC (menu5-revchart1-yoy-ratio) - công thức cột "Tr.trưởng YoY" là TỶ
    // LỆ đơn giản "kỳ này / kỳ trước" (không phải % tăng trưởng kiểu diff),
    // áp dụng chung cho cả 3 nhánh year/quarter/month — mỗi nhánh chỉ khác ở
    // đơn vị "kỳ" (năm/quý/tháng) đang so sánh liền kề.
    const buildRatio = (vals: (number | null)[]) => vals.map((v, i) => {
      if (i === 0) return null;
      const prev = vals[i - 1];
      return prev != null && prev > 0 && v != null ? v / prev : null;
    });
    if (viewBy === 'quarter') {
      const quarterPtsOf = (item: string) => (modelData?.series[item] ?? []).filter(p => p.y === lastRangeYear);
      const sumByQuarterOf = (item: string) => {
        const pts = quarterPtsOf(item);
        return [1, 2, 3, 4].map(q => {
          const vals = pts.filter(p => Math.ceil(p.m / 3) === q).map(p => p.v);
          return vals.length ? vals.reduce((a, v) => a + v, 0) : null;
        });
      };
      const labels = ['Q1', 'Q2', 'Q3', 'Q4'];
      const shipVals = sumByQuarterOf(ITEM_SHIPMENT);
      const prodVals = sumByQuarterOf(ITEM_PRODUCTION);
      const salesVals = sumByQuarterOf(ITEM_SALES);
      return { labels, shipVals, prodVals, salesVals, yoyVals: buildRatio(salesVals) };
    }
    if (viewBy === 'month') {
      const monthLabels = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const byMonthOf = (item: string) => {
        const map = new Map((modelData?.series[item] ?? []).filter(p => p.y === lastRangeYear).map(p => [p.m, p.v]));
        return monthLabels.map((_lab, i) => map.get(i + 1) ?? null);
      };
      const shipVals = byMonthOf(ITEM_SHIPMENT);
      const prodVals = byMonthOf(ITEM_PRODUCTION);
      const salesVals = byMonthOf(ITEM_SALES);
      return { labels: monthLabels, shipVals, prodVals, salesVals, yoyVals: buildRatio(salesVals) };
    }
    // viewBy === 'year'.
    const labels = rangeYears.map(String);
    const shipVals = rangeYears.map(y => ttlOf(modelData, ITEM_SHIPMENT, y));
    const prodVals = rangeYears.map(y => ttlOf(modelData, ITEM_PRODUCTION, y));
    const salesVals = rangeYears.map(y => ttlOf(modelData, ITEM_SALES, y));
    return { labels, shipVals, prodVals, salesVals, yoyVals: buildRatio(salesVals) };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [modelData, rangeYears, lastRangeYear, viewBy]);

  // EPCC (menu5-revchart2-viewby-link) - FIX theo yêu cầu người dùng "biểu đồ
  // 2,3 cũng cho áp dụng tương tự lựa chọn tháng/năm/quý": Chart 2 (Top 10
  // Model) có trục X là MODEL (không phải thời gian) nên không thể tách
  // thành nhiều cột theo quý/tháng như Chart 1 — thay vào đó, khi "XEM THEO"
  // đổi sang Quý/Tháng, Chart 2 xếp hạng theo SALES AMT của ĐÚNG kỳ (quý/
  // tháng) GẦN NHẤT có dữ liệu trong `lastRangeYear`, thay vì luôn lấy TTL cả
  // năm. Dùng `allModelsAggregate` (không phụ thuộc `selectedModel`) để xác
  // định kỳ gần nhất — đảm bảo mọi model được so sánh tại CÙNG 1 mốc thời
  // gian, kể cả khi đang xem 1 model cụ thể (model đó có thể không có dữ
  // liệu ở đúng tháng cuối cùng của công ty, coi như 0/ẩn thay vì lệch mốc).
  const revenuePeriodInfo = useMemo(() => {
    const pts = (allModelsAggregate?.series[ITEM_SALES] ?? []).filter(p => p.y === lastRangeYear);
    if (!pts.length) return { quarter: null as number | null, month: null as number | null };
    const lastMonth = Math.max(...pts.map(p => p.m));
    return { quarter: Math.ceil(lastMonth / 3), month: lastMonth };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allModelsAggregate, lastRangeYear]);

  // EPCC (menu5-chart4-6-viewby-link) - FIX theo yêu cầu người dùng "Chart4
  // biểu đồ tròn khi lựa chọn tháng/năm/quý thì số liệu thay đổi theo lựa
  // chọn đó, Chart5,6 cũng tương tự... không cố định như bây giờ": trước đây
  // hàm này CHỈ tính riêng SALES AMT (dùng cho Chart 2/revenueTopModels).
  // Đổi tên + thêm tham số `item` để DÙNG CHUNG được cho SHIPMENT/SCRAP (Chart
  // 6) — không đổi logic 3 nhánh year/quarter/month bên trong, chỉ tổng quát
  // hoá item đang truy vấn. `salesAtViewByPeriod` giữ lại làm alias mỏng gọi
  // `valueAtViewByPeriod(m, ITEM_SALES)` để không phải sửa nơi đang gọi nó
  // (revenueTopModels).
  function valueAtViewByPeriod(m: ModelEntry | null, item: string): number | null {
    if (viewBy === 'year') return ttlOf(m, item, lastRangeYear);
    const pts = (m?.series[item] ?? []).filter(p => p.y === lastRangeYear);
    if (viewBy === 'quarter') {
      if (revenuePeriodInfo.quarter == null) return null;
      const vals = pts.filter(p => Math.ceil(p.m / 3) === revenuePeriodInfo.quarter).map(p => p.v);
      return vals.length ? vals.reduce((a, v) => a + v, 0) : null;
    }
    // viewBy === 'month'.
    if (revenuePeriodInfo.month == null) return null;
    return pts.find(p => p.m === revenuePeriodInfo.month)?.v ?? null;
  }
  // Doanh số của 1 model tại đúng kỳ đang "XEM THEO" (năm/quý/tháng) —
  // dùng chung cho cả nhánh "Tất cả" lẫn nhánh 1 model cụ thể bên dưới.
  function salesAtViewByPeriod(m: ModelEntry | null): number | null {
    return valueAtViewByPeriod(m, ITEM_SALES);
  }

  // EPCC (menu5-chart4-6-viewby-link) - Doanh số của 1 KHÁCH HÀNG (không phải
  // model) tại đúng kỳ đang "XEM THEO" (năm/quý/tháng) — dùng cho Chart 4
  // (donut TỶ TRỌNG THEO KHÁCH HÀNG). Khác `valueAtViewByPeriod` (đọc từ
  // `ModelEntry.series`) vì dữ liệu khách hàng nằm ở bucket riêng
  // `salesByCustomer`(TTL/năm)/`salesByCustomerMonthly` (theo tháng) — tái
  // dùng ĐÚNG `revenuePeriodInfo` (quý/tháng gần nhất có dữ liệu, đã tính sẵn
  // ở trên cho Chart 2) để đảm bảo Chart 2/4/6 luôn đồng bộ CÙNG 1 mốc thời
  // gian khi "XEM THEO" đổi Quý/Tháng.
  function customerSalesAtViewByPeriod(custom: string): number | null {
    if (lastRangeYear == null) return null;
    if (viewBy === 'year') return salesByCustomer[custom]?.[lastRangeYear] ?? null;
    const byMonth = salesByCustomerMonthly[custom]?.[lastRangeYear] ?? {};
    if (viewBy === 'quarter') {
      if (revenuePeriodInfo.quarter == null) return null;
      let sum = 0, has = false;
      for (let m = (revenuePeriodInfo.quarter - 1) * 3 + 1; m <= revenuePeriodInfo.quarter * 3; m++) {
        if (byMonth[m] != null) { sum += byMonth[m]; has = true; }
      }
      return has ? sum : null;
    }
    // viewBy === 'month'.
    if (revenuePeriodInfo.month == null) return null;
    return byMonth[revenuePeriodInfo.month] ?? null;
  }

  // EPCC (menu5-revchart2-link-model-year) - FIX theo yêu cầu người dùng
  // "liên kết thanh lựa chọn model/năm sẽ thay đổi theo đúng giá trị của
  // model và năm đó": trước đây Chart 2 LUÔN lấy SALES AMT TTL của
  // `revenueLastYear` (năm cuối CỐ ĐỊNH trong toàn bộ dữ liệu, không đổi dù
  // đổi bộ lọc NĂM) và LUÔN vẽ đủ real model (không đổi dù chọn 1 model cụ
  // thể ở filter MODEL). Sửa lại 3 điểm, dùng ĐÚNG state đã có sẵn (không
  // thêm state mới):
  //  - Năm dùng để xếp hạng: đổi sang `lastRangeYear` (đã tự ứng với ô chọn
  //    nhanh NĂM + THÁNG BẮT ĐẦU/KẾT THÚC ở toolbar) thay vì `revenueLastYear`
  //    cố định — chọn năm nào, Top 10 hiển thị ĐÚNG doanh số năm đó.
  //  - Khi chọn 1 model cụ thể ở filter MODEL (khác "Tất cả"): thu gọn danh
  //    sách xuống ĐÚNG 1 phần tử của model đó (dùng `modelData`, đã tự ứng
  //    với `selectedModel`), không còn hiện Top 10 model khác — nhất quán
  //    với cách Chart 1 & Chart 2 (tab F-COST) đã link theo `selectedModel`
  //    trước đó (xem comment "menu5-chart1-chart2-link-selected-model").
  //    Giữ "Tất cả" thì y hệt hành vi cũ: Top 10 model giảm dần theo doanh số.
  //  - EPCC (menu5-revchart2-viewby-link) - giá trị xếp hạng đổi sang
  //    `salesAtViewByPeriod()` (xem khai báo ngay phía trên) thay vì luôn
  //    dùng thẳng `ttlOf(..., lastRangeYear)` — tự động ăn theo Tháng/Quý/
  //    Năm đang chọn ở toolbar.
  const revenueTopModels = useMemo(() => {
    if (lastRangeYear == null) return [];
    if (selectedModel !== ALL_MODEL_VALUE) {
      const value = salesAtViewByPeriod(modelData);
      return value != null && value > 0 ? [{ label: selectedModel, value }] : [];
    }
    return realModelsForAgg
      .map(m => ({ label: m.model, value: salesAtViewByPeriod(m) }))
      .filter((x): x is { label: string; value: number } => x.value != null && x.value > 0)
      .sort((a, b) => b.value - a.value)
      .slice(0, CHART_TOP_N_MODELS);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [realModelsForAgg, lastRangeYear, selectedModel, modelData, viewBy, revenuePeriodInfo]);

  // Chart 5 (RTV theo Tháng/Quý/Năm): bar=SHIPMENT(KEA) + RTV(KEA), line=%
  // Rate (RTV/Shipment). EPCC (menu5-revchart5-6-rtv) - THÊM MỚI theo yêu
  // cầu người dùng "tạo thêm chart 5,6 y hệt ảnh tham chiếu RTV THEO THÁNG
  // NĂM/RTV THEO MODEL... dựa vào dữ liệu đã có sẵn tại menu5db": KHÔNG có
  // cột "RTV" riêng trong Test5.xlsx, nên tái dùng ĐÚNG 2 item số lượng đã
  // đọc sẵn (`ITEM_SHIPMENT`, `ITEM_SCRAP` — vốn chỉ dùng cho tab F-COST) —
  // SCRAP (hàng lỗi/trả về) đúng bản chất khớp với khái niệm "RTV" trong ảnh
  // tham chiếu (cùng đơn vị KEA, cùng vai trò "phần bị loại khỏi SHIPMENT").
  // Nhân bản NGUYÊN VẸN 3 nhánh year/quarter/month của `revenueChart1` (cùng
  // `modelData`/`rangeYears`/`lastRangeYear`/`viewBy` — tự ứng theo model +
  // khoảng năm/tháng đang lọc ở toolbar), chỉ đổi item lấy dữ liệu.
  const revenueChart5 = useMemo(() => {
    const rateOf = (ship: number | null, rtv: number | null) =>
      ship != null && ship > 0 && rtv != null ? (rtv / ship) * 100 : null;
    if (viewBy === 'quarter') {
      const quarterPtsOf = (item: string) => (modelData?.series[item] ?? []).filter(p => p.y === lastRangeYear);
      const sumByQuarterOf = (item: string) => {
        const pts = quarterPtsOf(item);
        return [1, 2, 3, 4].map(q => {
          const vals = pts.filter(p => Math.ceil(p.m / 3) === q).map(p => p.v);
          return vals.length ? vals.reduce((a, v) => a + v, 0) : null;
        });
      };
      const labels = ['Q1', 'Q2', 'Q3', 'Q4'];
      const shipVals = sumByQuarterOf(ITEM_SHIPMENT);
      const rtvVals = sumByQuarterOf(ITEM_SCRAP);
      return { labels, shipVals, rtvVals, ratePct: shipVals.map((s, i) => rateOf(s, rtvVals[i])) };
    }
    if (viewBy === 'month') {
      const monthLabels = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const byMonthOf = (item: string) => {
        const map = new Map((modelData?.series[item] ?? []).filter(p => p.y === lastRangeYear).map(p => [p.m, p.v]));
        return monthLabels.map((_lab, i) => map.get(i + 1) ?? null);
      };
      const shipVals = byMonthOf(ITEM_SHIPMENT);
      const rtvVals = byMonthOf(ITEM_SCRAP);
      return { labels: monthLabels, shipVals, rtvVals, ratePct: shipVals.map((s, i) => rateOf(s, rtvVals[i])) };
    }
    // viewBy === 'year'.
    const labels = rangeYears.map(String);
    const shipVals = rangeYears.map(y => ttlOf(modelData, ITEM_SHIPMENT, y));
    const rtvVals = rangeYears.map(y => ttlOf(modelData, ITEM_SCRAP, y));
    return { labels, shipVals, rtvVals, ratePct: shipVals.map((s, i) => rateOf(s, rtvVals[i])) };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [modelData, rangeYears, lastRangeYear, viewBy]);

  // Chart 6 (RTV theo Model): top-N model (CHART_TOP_N_MODELS, cùng quy ước
  // với Chart 2/revenueTopModels) xếp theo RTV(=SCRAP) giảm dần tại ĐÚNG kỳ
  // đang "XEM THEO" (năm/quý/tháng) — tái dùng `realModelsForAgg` sẵn có.
  // EPCC (menu5-chart4-6-viewby-link) - FIX theo yêu cầu người dùng "Chart5,6
  // cũng ... hiển thị thành tháng/quý/năm không cố định như bây giờ": trước
  // đây `buildRow` LUÔN dùng `ttlOf(m, item, lastRangeYear)` — tức RTV/
  // SHIPMENT/%Rate của Chart 6 CHỈ đổi khi đổi bộ lọc NĂM, hoàn toàn không
  // phản hồi nút "XEM THEO" Tháng/Quý dù người dùng đã bấm đổi (khác Chart 5
  // ngay phía trên vốn đã branch đúng theo `viewBy`). Đổi `ttlOf(...)` sang
  // `valueAtViewByPeriod(m, item)` (hàm dùng chung mới tổng quát hoá từ
  // `salesAtViewByPeriod`, xem khai báo phía trên) — Quý/Tháng giờ lấy đúng
  // kỳ gần nhất có dữ liệu trong `lastRangeYear` (qua `revenuePeriodInfo`),
  // y hệt cách Chart 2 đã làm, nên thêm `viewBy`/`revenuePeriodInfo` vào deps.
  const revenueChart6 = useMemo(() => {
    if (lastRangeYear == null) return [];
    const buildRow = (m: ModelEntry | null, label: string) => {
      const shipVal = valueAtViewByPeriod(m, ITEM_SHIPMENT);
      const rtvVal = valueAtViewByPeriod(m, ITEM_SCRAP);
      const ratePct = shipVal != null && shipVal > 0 && rtvVal != null ? (rtvVal / shipVal) * 100 : null;
      return { label, shipVal, rtvVal, ratePct };
    };
    if (selectedModel !== ALL_MODEL_VALUE) {
      const row = buildRow(modelData, selectedModel);
      return row.rtvVal != null && row.rtvVal > 0 ? [row] : [];
    }
    return realModelsForAgg
      .map(m => buildRow(m, m.model))
      .filter(r => r.rtvVal != null && r.rtvVal > 0)
      .sort((a, b) => (b.rtvVal ?? 0) - (a.rtvVal ?? 0))
      .slice(0, CHART_TOP_N_MODELS);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [realModelsForAgg, lastRangeYear, selectedModel, modelData, viewBy, revenuePeriodInfo]);

  // Danh sách khách hàng thật có trong Test5.xlsx (CUSTOM), màu gán cố định
  // theo thứ tự bảng chữ cái để luôn nhất quán giữa Chart 3/Chart 4/legend.
  const revenueCustomers = useMemo(
    () => Object.keys(salesByCustomer).sort(),
    [salesByCustomer]
  );
  const revenueCustomerColor = (custom: string) => {
    const idx = revenueCustomers.indexOf(custom);
    return CHART_PALETTE_9[idx % CHART_PALETTE_9.length];
  };

  // Chart 3 (Doanh số theo Khách hàng, cột chồng theo năm).
  // EPCC (menu5-revenue-custom-filter) - chọn 1 khách hàng cụ thể ở bộ lọc
  // CUSTOM mới sẽ giới hạn `revenueCustomers` xuống đúng 1 phần tử (chart
  // vẫn vẽ đúng như cũ, chỉ còn lại đúng khách hàng đó); giữ "Tất cả"
  // (ALL_CUSTOM_VALUE) thì y hệt hành vi cũ (mọi khách hàng).
  const revenueCustomersFiltered = useMemo(
    () => selectedCustom === ALL_CUSTOM_VALUE ? revenueCustomers : revenueCustomers.filter(c => c === selectedCustom),
    [revenueCustomers, selectedCustom]
  );
  // EPCC (menu5-revchart3-viewby-link) - FIX theo yêu cầu người dùng "biểu đồ
  // 2,3 cũng cho áp dụng tương tự lựa chọn tháng/năm/quý và theo năm cũng
  // thay đổi như vậy": trước đây Chart 3 LUÔN cộng dồn `salesByCustomer`
  // (TTL/năm) trên `revenueYears` — TOÀN BỘ năm có dữ liệu, không đổi dù đổi
  // bộ lọc NĂM, và hoàn toàn không có khái niệm Quý/Tháng (dữ liệu nguồn
  // trước đây chỉ có TTL/năm, không đủ chi tiết). Sửa 2 điểm:
  //  - Khoảng năm (nhánh 'year'): đổi từ `revenueYears` sang `rangeYears`
  //    (đã tự ứng với ô chọn nhanh NĂM/THÁNG BẮT ĐẦU-KẾT THÚC ở toolbar).
  //  - Thêm 2 nhánh 'quarter'/'month' (giống hệt cấu trúc `revenueChart1`):
  //    dùng `salesByCustomerMonthly` (mới thêm — xem comment
  //    "menu5-revenue-chart3-viewby-monthly-customer" ở AggAccumulator) để
  //    lấy đúng SALES AMT theo Quý/Tháng của `lastRangeYear` cho từng khách
  //    hàng, thay vì chỉ có TTL/năm.
  const revenueChart3 = useMemo(() => {
    if (viewBy === 'quarter') {
      const labels = ['Q1', 'Q2', 'Q3', 'Q4'];
      const series = revenueCustomersFiltered.map(custom => {
        const byMonth = salesByCustomerMonthly[custom]?.[lastRangeYear] ?? {};
        const values = [1, 2, 3, 4].map(q => {
          let sum = 0;
          for (let m = (q - 1) * 3 + 1; m <= q * 3; m++) sum += byMonth[m] ?? 0;
          return sum;
        });
        return { label: custom, color: revenueCustomerColor(custom), values };
      });
      return { labels, series };
    }
    if (viewBy === 'month') {
      const labels = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const series = revenueCustomersFiltered.map(custom => {
        const byMonth = salesByCustomerMonthly[custom]?.[lastRangeYear] ?? {};
        const values = labels.map((_lab, i) => byMonth[i + 1] ?? 0);
        return { label: custom, color: revenueCustomerColor(custom), values };
      });
      return { labels, series };
    }
    // viewBy === 'year'.
    const labels = rangeYears.map(String);
    const series = revenueCustomersFiltered.map(custom => ({
      label: custom,
      color: revenueCustomerColor(custom),
      values: rangeYears.map(y => salesByCustomer[custom]?.[y] ?? 0),
    }));
    return { labels, series };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revenueCustomersFiltered, rangeYears, lastRangeYear, viewBy, salesByCustomer, salesByCustomerMonthly]);

  // Chart 4 (Tỷ trọng theo Khách hàng, donut tại ĐÚNG kỳ đang "XEM THEO").
  // EPCC (menu5-chart4-6-viewby-link) - FIX theo yêu cầu người dùng "Chart4
  // biểu đồ tròn khi lựa chọn tháng/năm/quý thì số liệu thay đổi theo lựa
  // chọn đó": trước đây donut LUÔN lấy TTL cả năm của `revenueLastYear` (năm
  // cuối CỐ ĐỊNH toàn bộ dữ liệu, không đổi dù bấm "XEM THEO" hay đổi bộ lọc
  // NĂM). Đổi sang `customerSalesAtViewByPeriod(custom)` (hàm mới ở trên,
  // tự ứng năm đang lọc `lastRangeYear` + đúng Quý/Tháng gần nhất có dữ liệu
  // qua `revenuePeriodInfo`) — donut giờ đổi tỷ trọng khách hàng ĐÚNG theo
  // kỳ Tháng/Quý/Năm đang chọn ở toolbar, đồng bộ với Chart 2/6.
  const revenueChart4 = useMemo(() => {
    if (lastRangeYear == null) return [];
    return revenueCustomersFiltered
      .map(custom => ({ label: custom, value: customerSalesAtViewByPeriod(custom) ?? 0, color: revenueCustomerColor(custom) }))
      .filter(x => x.value > 0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [revenueCustomersFiltered, lastRangeYear, viewBy, revenuePeriodInfo, salesByCustomer, salesByCustomerMonthly]);

  // EPCC (menu5-revchart4-yoy-panel) - THÊM MỚI theo yêu cầu người dùng "chart
  // 3 (donut TỶ TRỌNG THEO KHÁCH HÀNG, idx=3) lồng thêm bảng 'Tăng giảm giữa
  // các năm' như ảnh tham chiếu": chênh lệch SALES AMT (K$) giữa mỗi năm và
  // năm liền trước, dùng TOÀN BỘ `revenueYears`/`allModelsAggregate` đã có
  // sẵn (đúng nguồn `revenueKpis.yoyPct` đang dùng cho Card YoY) — không
  // thêm bucket/field Excel mới. Năm ĐẦU TIÊN không có năm liền trước nên
  // luôn `diff = null` (không hiển thị số, giống ô "2017" trống trong ảnh
  // tham chiếu). `.reverse()` để năm GẦN NHẤT lên trên cùng, khớp thứ tự
  // 2025→2017 trong ảnh mẫu.
  const revenueYoyByYear = useMemo(() => {
    return revenueYears
      .map((y, i) => {
        const curr = allModelsAggregate?.ttlByItemYear[ITEM_SALES]?.[y] ?? null;
        if (i === 0) return { year: y, diff: null as number | null };
        const prevYear = revenueYears[i - 1];
        const prev = allModelsAggregate?.ttlByItemYear[ITEM_SALES]?.[prevYear] ?? null;
        const diff = curr != null && prev != null ? curr - prev : null;
        return { year: y, diff };
      })
      .slice()
      .reverse();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [allModelsAggregate, revenueYears]);

  // Bảng "Toàn bộ Model theo Doanh số": liệt kê TOÀN BỘ số liệu có trong file
  // Excel (mọi model × mọi năm, không còn giới hạn "năm gần nhất" như trước),
  // kèm 1 thanh bộ lọc riêng phía trên card (Model/Năm/Quý/Tháng/Khách hàng +
  // ô tìm kiếm tự do) — nhân bản đúng kiểu bố cục "Toàn bộ Cơ sở Dữ liệu từ
  // Excel" trong ảnh tham chiếu người dùng cung cấp.
  // EPCC (menu5-revenue-table-full-list-filters) - FIX theo yêu cầu người
  // dùng "hãy liệt kê tất cả số liệu có từ file excel lên và trên đầu Card
  // top để dạng có lựa chọn như ảnh 1 tham chiếu": trước đây bảng CHỈ tính 1
  // dòng/model tại ĐÚNG năm gần nhất (`revenueLastYear`) — chỉ là 1 lát cắt
  // nhỏ của dữ liệu, không phải "tất cả số liệu". Đổi sang duyệt MỌI (model,
  // năm) có dữ liệu thật (dùng `ttlOf`, đúng nguồn TTL gốc từ Excel) thành
  // nhiều dòng/model — 1 dòng cho mỗi năm. Khi người dùng chọn cụ thể Quý/
  // Tháng ở bộ lọc mới, "khoan sâu" xuống đúng kỳ đó bằng cách cộng dồn trực
  // tiếp từ `m.series[item]` (dữ liệu THÁNG gốc, đã đọc sẵn khi parse Excel —
  // xem SeriesPoint{y,m,v}), y hệt cách Chart 5 (RTV THEO THÁNG NĂM) đã làm ở
  // trên (sumByQuarterOf/byMonthOf), không tạo nguồn dữ liệu mới. Cột "Tăng
  // trưởng doanh số" giờ luôn so với ĐÚNG kỳ cùng loại của năm liền trước
  // (năm→năm, quý→quý, tháng→tháng — so YoY tại kỳ đang xem, tránh lệch mùa
  // vụ nếu so quý này với quý trước liền kề).
  const revTableGranularity: 'year' | 'quarter' | 'month' =
    revTableMonth !== ALL_CUSTOM_VALUE ? 'month' : revTableQuarter !== ALL_CUSTOM_VALUE ? 'quarter' : 'year';
  // EPCC (menu5-detail-merge-revenue-table-add-fcost) - màu cột F-Cost % cho
  // bảng đã gộp, tái dùng đúng bảng NEON có sẵn (không bịa màu mới) — nhân
  // bản quy ước màu cũ của bảng "Chi tiết & dữ liệu" (Tháng=#4EB09B, Quý=
  // NEON.amber, Năm=NEON.violet), chỉ đổi từ tính theo `viewBy` (toolbar
  // chính) sang tính theo `revTableGranularity` (bộ lọc riêng của bảng này).
  const revTableFcostColor = revTableGranularity === 'quarter' ? NEON.amber : revTableGranularity === 'year' ? NEON.violet : '#4EB09B';
  // EPCC (menu5-detail-merge-revenue-table-add-fcost) - FIX theo yêu cầu
  // người dùng "gộp luôn dữ liệu tại sheet tình hình doanh thu sang sheet
  // chi tiết và dữ liệu thêm phần Scrap và F-cost% hiển thị ra nữa": bảng
  // này giờ được HIỂN THỊ Ở TAB "Chi tiết & dữ liệu" (thay cho bảng đơn giản
  // cũ chỉ có 1 model/tháng — xem JSX nhánh `tab === 'detail'` bên dưới),
  // nên cần thêm F-Cost% — CHỈ CÒN THIẾU so với 6 cột gốc của bảng cũ
  // (THÁNG/SHIPMENT/PROD'N/SALES AMT/SCRAP/F-COST%; cột SCRAP đã có sẵn ở
  // đây dưới tên "RTV(K)" — cùng 1 nguồn dữ liệu `ITEM_SCRAP`, không tạo cột
  // trùng lặp). F-Cost% là TỶ LỆ (không cộng dồn được như số lượng), nên ở
  // mức Quý phải lấy TRUNG BÌNH các tháng có dữ liệu trong quý (`avgByQuarterFcost`),
  // khác hẳn `sumByQuarter` (cộng dồn) dùng cho Sản xuất/Xuất hàng/Doanh số/
  // RTV phía trên.
  const avgByQuarterFcost = (m: ModelEntry, year: number, quarter: number): number | null => {
    const pts = (m.series[ITEM_FCOST] ?? []).filter(p => p.y === year && Math.ceil(p.m / 3) === quarter);
    return pts.length ? pts.reduce((a, p) => a + p.v, 0) / pts.length : null;
  };
  const revenueTopModelsTable = useMemo(() => {
    if (!revenueYears.length) return [];
    const topOf = <T extends { value: number }>(entries: T[]): T | null =>
      entries.length ? entries.reduce((a, b) => (b.value > a.value ? b : a)) : null;

    const sumByQuarter = (m: ModelEntry, item: string, year: number, quarter: number): number | null => {
      const pts = (m.series[item] ?? []).filter(p => p.y === year && Math.ceil(p.m / 3) === quarter);
      return pts.length ? pts.reduce((a, p) => a + p.v, 0) : null;
    };
    const sumByMonth = (m: ModelEntry, item: string, year: number, month: number): number | null => {
      const pt = (m.series[item] ?? []).find(p => p.y === year && p.m === month);
      return pt ? pt.v : null;
    };
    const valueOf = (m: ModelEntry, item: string, year: number): number | null => {
      if (revTableGranularity === 'month') return sumByMonth(m, item, year, Number(revTableMonth));
      if (revTableGranularity === 'quarter') return sumByQuarter(m, item, year, Number(revTableQuarter));
      return ttlOf(m, item, year);
    };

    const searchQ = revTableSearch.trim().toLowerCase();
    const candidateModels = realModelsForAgg.filter(m => {
      if (revTableModel !== ALL_MODEL_VALUE && m.model !== revTableModel) return false;
      const customer = topOf(m.radarCustom)?.custom ?? '—';
      if (revTableCustomer !== ALL_CUSTOM_VALUE && customer !== revTableCustomer) return false;
      if (searchQ) {
        const type = topOf(m.radarType)?.type ?? '';
        if (!m.model.toLowerCase().includes(searchQ) && !customer.toLowerCase().includes(searchQ) && !type.toLowerCase().includes(searchQ)) return false;
      }
      return true;
    });
    const years = revTableYear !== ALL_CUSTOM_VALUE ? [Number(revTableYear)] : revenueYears;

    const rows: Array<{
      key: string; model: string; period: string; year: number; customer: string; type: string;
      prod: number; ship: number; sales: number; rtv: number; rtvRatePct: number | null;
      growth: number | null; growthPct: number | null; fcostPct: number | null;
    }> = [];
    for (const m of candidateModels) {
      const customer = topOf(m.radarCustom)?.custom ?? '—';
      const type = topOf(m.radarType)?.type ?? '—';
      for (const y of years) {
        const sales = valueOf(m, ITEM_SALES, y);
        if (sales == null) continue; // không có dữ liệu ở đúng kỳ này — bỏ qua, không suy diễn.
        const prod = valueOf(m, ITEM_PRODUCTION, y) ?? 0;
        const ship = valueOf(m, ITEM_SHIPMENT, y) ?? 0;
        const rtv = valueOf(m, ITEM_SCRAP, y) ?? 0;
        const rtvRatePct = ship > 0 ? (rtv / ship) * 100 : null;
        const fcostPct = revTableGranularity === 'month' ? sumByMonth(m, ITEM_FCOST, y, Number(revTableMonth))
          : revTableGranularity === 'quarter' ? avgByQuarterFcost(m, y, Number(revTableQuarter))
          : ttlOf(m, ITEM_FCOST, y);
        const prevYear = y - 1;
        const prevSales = revTableGranularity === 'month' ? sumByMonth(m, ITEM_SALES, prevYear, Number(revTableMonth))
          : revTableGranularity === 'quarter' ? sumByQuarter(m, ITEM_SALES, prevYear, Number(revTableQuarter))
          : ttlOf(m, ITEM_SALES, prevYear);
        const growth = prevSales != null ? sales - prevSales : null;
        const growthPct = prevSales != null && prevSales > 0 ? (growth! / prevSales) * 100 : null;
        const period = revTableGranularity === 'month' ? `T${revTableMonth}/${y}`
          : revTableGranularity === 'quarter' ? `Q${revTableQuarter}/${y}` : `${y}`;
        rows.push({
          key: `${m.model}-${period}`, model: m.model, period, year: y, customer, type,
          prod, ship, sales, rtv, rtvRatePct, growth, growthPct, fcostPct,
        });
      }
    }
    return rows.sort((a, b) => a.model.localeCompare(b.model) || b.year - a.year);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [realModelsForAgg, revenueYears, revTableModel, revTableCustomer, revTableSearch, revTableYear, revTableQuarter, revTableMonth, revTableGranularity]);

  // EPCC (menu5-toolbar-import-not-export) - FIX ROOT CAUSE "nút toolbar tên
  // 'Tải Excel' lại gọi handleExport() và GHI FILE MỚI RA MÁY thay vì NHẬP
  // dữ liệu vào": đã xoá handleExport() (không còn nơi nào gọi). Toolbar bên
  // dưới giờ dùng đúng lại handleUploadFile()/fileInputRef sẵn có (vốn chỉ
  // được dùng ở màn hình rỗng ban đầu) để nút "Tải Excel" luôn là hành động
  // NHẬP file .xlsx mới, đồng bộ đúng ý nghĩa tên nút và đúng hành vi của
  // nút tương ứng trong TargetActualDashboard.tsx (t.loadExcelBtn).

  // ── Trạng thái rỗng: chưa upload Excel nào (và không có cache cũ) ────────
  // Thay cho việc luôn có sẵn dữ liệu mẫu test5Data.json trước đây — giờ
  // Mục 5 bắt đầu rỗng giống Mục 1/2/3, hiện màn hình mời tải file trực
  // tiếp thay vì cố render KPI/biểu đồ trên dữ liệu trống.
  if (allModels.length === 0) {
    return (
      <div className={`second-dashboard menu5-dashboard`} style={{ padding: '0 16px 16px 16px', height: '100%', overflowY: 'auto', boxSizing: 'border-box' }}>
        {/* EPCC (menu5-outer-wrapper-match-rty) - FIX ROOT CAUSE "2 nhánh
            return của Menu5 từng dùng class bao ngoài 'dashboard-screen'
            (hệ layout riêng, khác RTY) trong khi RTY dùng 'second-dashboard'
            — đổi cho khớp để tránh lặp lại bug lệch khoảng cách do trộn 2 hệ
            CSS khác nhau, xem chú thích chi tiết ở nhánh return có dữ liệu
            bên dưới (cùng slug menu5-outer-wrapper-match-rty). */}
        <style>{`
          .menu5-dashboard .dashboard-header-grid {
            background: #2F3A1D;
            border-radius: 14px;
            padding: 10px 16px;
            border: 1px solid rgba(0,0,0,0.18);
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
          }
          .menu5-dashboard .dashboard-header-left { color: #C0EF6A !important; }
          .menu5-dashboard .dashboard-header-title { color: #C0EF6A !important; }
          .menu5-dashboard .tab-container {
            display: flex;
            gap: 10px;
            margin-bottom: 8px;
          }
          .menu5-dashboard .tab-btn {
            padding: 8px 18px;
            border-radius: 8px;
            font-weight: 700;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.2s ease;
            border: 1px solid var(--border);
            background: rgba(30, 41, 59, 0.2);
            color: var(--text-2);
            display: inline-flex;
            align-items: center;
            gap: 6px;
            white-space: nowrap;
          }
          .menu5-dashboard .tab-btn.active {
            background: linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%);
            color: #ffffff;
            border-color: #3b82f6;
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.35);
          }
        `}</style>

        {/* ── Header (copy nguyên văn RtyDashboard.tsx, con trực tiếp của
            outer div — không còn div padding trung gian) ── */}
        <div className="dashboard-header-grid">
          <div className="dashboard-header-left" style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-2)', fontWeight: 700, whiteSpace: 'nowrap' }}>
            <span aria-hidden="true">🕐</span>
            {formattedTime}
          </div>
          <h1 className="dashboard-header-title">{t.title}</h1>
          <div className="dashboard-header-right" />
        </div>

        <main
          className="dash-body"
          style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '60vh', padding: 20 }}
        >
          <div
            style={{
              background: 'var(--surface)', padding: 40, borderRadius: 12,
              border: '1px solid var(--border-soft)', maxWidth: 480, width: '100%',
              textAlign: 'center', boxShadow: 'var(--shadow-sm)',
            }}
          >
            <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
            <h2 style={{ color: 'var(--text-0)', marginBottom: 8 }}>{t.title}</h2>
            <p style={{ color: 'var(--text-2)', marginBottom: 24 }}>{t.emptyDesc}</p>
            {uploadError && (
              <p style={{ color: NEON.pink, marginBottom: 16, fontSize: 13 }}>{uploadError}</p>
            )}
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              style={{ display: 'none' }}
              onChange={e => {
                const file = e.target.files?.[0];
                if (file) handleUploadFile(file);
              }}
            />
            <NeonButton
              className="btn btn-outline neon-violet"
              style={{ cursor: 'pointer', display: 'inline-flex', padding: '12px 24px', fontSize: 15 }}
              onClick={() => fileInputRef.current?.click()}
              disabled={isParsing}
            >
              {isParsing ? `⏳ ${t.parsing}` : `📤 ${t.uploadExcel}`}
            </NeonButton>
          </div>
        </main>
      </div>
    );
  }

  return (
    <div className={`second-dashboard menu5-dashboard`} style={{ padding: '0 16px 16px 16px', height: '100%', overflowY: 'auto', boxSizing: 'border-box' }}>
      {/* EPCC (menu5-outer-wrapper-match-rty) - FIX ROOT CAUSE "khoảng cách
          filter→KPI hiển thị khác hẳn RTY dù inline marginBottom giống hệt
          nhau": nguyên nhân THẬT không nằm ở số đo mà ở CLASS BAO NGOÀI —
          trước đây Menu5 dùng "dashboard-screen" + "<main className=dash-body>"
          (hệ layout RIÊNG của Menu5, khác hệ với RTY, có thể cộng dồn
          margin/gap ẩn từ CSS chung không thấy được qua inline style), còn
          RTY dùng "second-dashboard" và KHÔNG có lớp <main> nào cả — mọi
          thanh (header/tab/filter/KPI) là con trực tiếp của 1 div cuộn duy
          nhất. Đổi outer wrapper Menu5 sang đúng "second-dashboard" + xóa
          hẳn <main className="dash-body">, để DOM giống hệt RTY (không chỉ
          số đo giống hệt) — đây là cách duy nhất đảm bảo cascade CSS không
          lệch giữa 2 dashboard. */}
      <style>{`
        .menu5-dashboard .dashboard-header-grid {
          background: #2F3A1D;
          border-radius: 14px;
          padding: 10px 16px;
          border: 1px solid rgba(0,0,0,0.18);
          box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        .menu5-dashboard .dashboard-header-left { color: #C0EF6A !important; }
        .menu5-dashboard .dashboard-header-title { color: #C0EF6A !important; }
        .menu5-dashboard .tab-container {
          display: flex;
          gap: 10px;
          margin-bottom: 8px;
        }
        .menu5-dashboard .tab-btn {
          padding: 8px 18px;
          border-radius: 8px;
          font-weight: 700;
          font-size: 14px;
          cursor: pointer;
          transition: all 0.2s ease;
          border: 1px solid var(--border);
          background: rgba(30, 41, 59, 0.2);
          color: var(--text-2);
          display: inline-flex;
          align-items: center;
          gap: 6px;
          white-space: nowrap;
        }
        .menu5-dashboard .tab-btn.active {
          background: linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%);
          color: #ffffff;
          border-color: #3b82f6;
          box-shadow: 0 4px 12px rgba(59, 130, 246, 0.35);
        }
      `}</style>

      {/* ── Header (copy nguyên văn RtyDashboard.tsx, con trực tiếp của
          outer div — không còn div padding trung gian) ── */}
      <div className="dashboard-header-grid">
        <div className="dashboard-header-left" style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-2)', fontWeight: 700, whiteSpace: 'nowrap' }}>
          <span aria-hidden="true">🕐</span>
          {formattedTime}
        </div>
        <h1 className="dashboard-header-title">{t.title}</h1>
        <div className="dashboard-header-right" />
      </div>

      {/* ── 3 Tab (copy nguyên văn RtyDashboard.tsx, thêm tab thứ 3 "TÌNH HÌNH
          DOANH THU" — EPCC menu5-revenue-tab) ── */}
      {/* EPCC (menu5-revenue-tab-move-first) - FIX theo yêu cầu người dùng
          "đổi sheet tình hình doanh thu lên phía trước và luôn hiện phía
          trước": đổi thứ tự nút bấm — tab DOANH THU giờ đứng ĐẦU TIÊN (trước
          F-COST), và đổi luôn tab MẶC ĐỊNH khi mở trang từ 'overview'
          (F-COST) sang 'revenue' (xem khai báo `useState<...>('revenue')`
          phía trên) để tab này LUÔN hiện ra trước, không cần bấm chọn. */}
      <div className="tab-container">
        <button className={`tab-btn ${tab === 'revenue' ? 'active' : ''}`} onClick={() => setTab('revenue')}>
          💰 {t.tabRevenue}
        </button>
        <button className={`tab-btn ${tab === 'overview' ? 'active' : ''}`} onClick={() => setTab('overview')}>
          📊 {t.tabOverview}
        </button>
        <button className={`tab-btn ${tab === 'detail' ? 'active' : ''}`} onClick={() => setTab('detail')}>
          📋 {t.tabDetail}
        </button>
      </div>

      {/* ── Toolbar + KPI + phần còn lại: giờ là con trực tiếp của outer
          "second-dashboard" div, KHÔNG còn bọc trong <main className=
          "dash-body"> — loại bỏ hoàn toàn khác biệt hệ CSS với RTY. ── */}
        {/* ── Toolbar: bố cục 2 dòng (nhãn / control), nền olive đậm, đồng bộ
            NGUYÊN VĂN kích thước-màu sắc-bố cục với topbar-dash trong
            TargetActualDashboard.tsx (Mục 2), để 2 dashboard nhìn thống nhất. ── */}
        <div className="topbar-dash" style={{
          display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '10px',
          background: '#2F3A1D',
          borderRadius: '14px', padding: '10px 14px',
          border: '1px solid rgba(0,0,0,0.18)',
          boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
        }}>
          {/* Dòng 1 (nhãn) */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
            <div style={{ width: '170px', flexShrink: 0, textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>
              {t.updatedToLabel}
            </div>
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px' }}>
              <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.startMonth}</span>
              <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.endMonth}</span>
              <span style={{ width: '140px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.model}</span>
              {/* EPCC (menu5-revenue-custom-filter) - nhãn cột lọc CUSTOM
                  mới, đặt ngay sau MODEL theo đúng vị trí ảnh tham chiếu
                  người dùng cung cấp (cùng cụm bộ lọc dữ liệu, trước cụm
                  XEM THEO/NĂM). */}
              <span style={{ width: '120px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.customFilter}</span>
              <span style={{ width: '180px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.viewBy}</span>
              <span style={{ width: '90px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.yearFilterLabel}</span>
            </div>
            <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
              <div style={{ width: '120px' }}></div>
            </div>
          </div>

          {/* Dòng 2 (control) */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
            <div style={{ width: '170px', flexShrink: 0, textAlign: 'center', fontSize: '13.5px', fontWeight: 700, color: '#22c55e', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
              {lastUpdated
                ? `📅 ${new Date(lastUpdated).toLocaleDateString()}`
                : (lang === 'vi' ? 'Chưa có dữ liệu' : lang === 'ko' ? '데이터 없음' : 'No data yet')}
            </div>

            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px', alignItems: 'center' }}>
              <input type="month" value={toMonthInput(startYm)}
                min={toMonthInput(globalBounds.minYm)} max={toMonthInput(endYm)}
                onChange={e => setStartYm(fromMonthInput(e.target.value))}
                className="filter-date-input"
                style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }} />
              <input type="month" value={toMonthInput(endYm)}
                min={toMonthInput(startYm)} max={toMonthInput(globalBounds.maxYm)}
                onChange={e => setEndYm(fromMonthInput(e.target.value))}
                className="filter-date-input"
                style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }} />
              <select value={selectedModel} onChange={e => setSelectedModel(e.target.value)}
                style={{
                  width: '140px', height: '38px', boxSizing: 'border-box', borderRadius: '8px',
                  border: '1px solid rgba(0,0,0,0.18)', background: 'rgba(255,255,255,0.9)',
                  color: '#1f2937', padding: '0 8px', fontSize: '13px',
                }}>
                {/* EPCC (menu5-model-all-option) - FIX theo yêu cầu người dùng
                    "chọn 'Tất cả' ở MODEL áp dụng toàn trang": thêm option
                    "Tất cả" ở đầu danh sách (giá trị `ALL_MODEL_VALUE`).
                    Đồng thời loại 4 model đặc biệt (`SPECIAL_FCOST_MODELS`:
                    Target/Actual/F-cost($)/F-Cost(억원)) khỏi danh sách chọn
                    trực tiếp — chúng vốn chỉ là dòng tổng hợp phụ trợ trong
                    Excel, không phải model sản xuất thật; dòng "Actual" giờ
                    được dùng NGẦM đứng sau lựa chọn "Tất cả" thay vì hiện ra
                    như 1 model riêng lẻ gây trùng lặp/khó hiểu. */}
                <option value={ALL_MODEL_VALUE}>{t.modelFilterAll}</option>
                {allModels.filter(m => !SPECIAL_FCOST_MODELS.includes(m.model)).map(m => (
                  <option key={m.model} value={m.model}>{m.model}{m.model === bestModel ? ` ⭐` : ''}</option>
                ))}
              </select>
              {/* EPCC (menu5-revenue-custom-filter) - FIX theo yêu cầu người
                  dùng "bổ sung thêm mục CUSTOM lên thanh card này": select lọc
                  theo khách hàng (cột CUSTOM trong Excel), danh sách lấy từ
                  `revenueCustomers` (đã tổng hợp sẵn từ `salesByCustomer`,
                  đúng khách hàng thật có trong Test5.xlsx, không hardcode).
                  Đặt ngay sau MODEL theo đúng vị trí trong ảnh tham chiếu.
                  Chọn 1 khách hàng cụ thể sẽ lọc Chart 3 (Doanh số theo
                  Khách hàng) + Chart 4 (Tỷ trọng theo Khách hàng) ở tab TÌNH
                  HÌNH DOANH THU chỉ còn đúng khách hàng đó — 2 biểu đồ vốn dĩ
                  ĐÃ dùng đúng dữ liệu theo khách hàng (`salesByCustomer`),
                  nên đây là nơi filter CUSTOM có ý nghĩa nhất, không đụng tới
                  Chart 1/2 (theo model, không theo khách hàng). */}
              <select value={selectedCustom} onChange={e => setSelectedCustom(e.target.value)}
                style={{
                  width: '120px', height: '38px', boxSizing: 'border-box', borderRadius: '8px',
                  border: '1px solid rgba(0,0,0,0.18)', background: 'rgba(255,255,255,0.9)',
                  color: '#1f2937', padding: '0 8px', fontSize: '13px',
                }}>
                <option value={ALL_CUSTOM_VALUE}>{t.customFilterAll}</option>
                {revenueCustomers.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
              <div style={{ display: 'flex', gap: '0px', height: '38px', width: '180px', flexShrink: 0 }}>
                {(['month', 'quarter', 'year'] as ViewBy[]).map((v, i) => (
                  <button
                    key={v}
                    onClick={() => setViewBy(v)}
                    style={{
                      flex: 1, padding: '6px 0', fontSize: '13px', fontWeight: 600,
                      borderRadius: i === 0 ? '6px 0 0 6px' : i === 2 ? '0 6px 6px 0' : '0',
                      border: '1px solid rgba(0,0,0,0.18)',
                      borderRight: i !== 2 ? 'none' : '1px solid rgba(0,0,0,0.18)',
                      background: viewBy === v ? '#2e7d8c' : 'rgba(255,255,255,0.55)',
                      color: viewBy === v ? '#ffffff' : '#7A5A2E',
                      cursor: 'pointer', transition: 'all 0.15s ease', height: '100%',
                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center', whiteSpace: 'nowrap',
                    }}
                  >
                    {v === 'month' ? t.byMonth : v === 'quarter' ? t.byQuarter : t.byYear}
                  </button>
                ))}
              </div>
              {/* EPCC (menu5-year-quick-filter) - ô lọc nhanh theo năm: "Tất
                  cả" (reset về đúng biên dữ liệu thật `globalBounds`) + từng
                  năm có trong `allDataYears` (dò động từ file Excel đã tải
                  lên, KHÔNG hardcode danh sách năm — file có thêm năm mới thì
                  option tự xuất hiện theo, không cần sửa code). Chọn 1 năm sẽ
                  đặt lại `startYm`/`endYm` = T1→T12 của năm đó, dùng lại
                  ĐÚNG cơ chế lọc tháng sẵn có nên tự động link với mọi biểu
                  đồ/KPI khác trong dashboard (không cần thêm state lọc riêng). */}
              <select value={selectedYearFilterValue} onChange={e => handleYearFilterChange(e.target.value)}
                style={{
                  width: '90px', height: '38px', boxSizing: 'border-box', borderRadius: '8px',
                  border: '1px solid rgba(0,0,0,0.18)', background: 'rgba(255,255,255,0.9)',
                  color: '#1f2937', padding: '0 6px', fontSize: '13px', flexShrink: 0,
                }}>
                <option value="all">{t.yearFilterAll}</option>
                {allDataYears.map(y => (
                  <option key={y} value={String(y)}>{y}</option>
                ))}
              </select>
            </div>

            {/* Cụm phải: nút "Tải tệp lên" — NÚT NHẬP (import) file .xlsx,
                dùng lại đúng fileInputRef/handleUploadFile đã có sẵn cho màn
                hình rỗng ban đầu.
                EPCC (menu5-import-btn-upload-icon-label) - FIX theo yêu cầu
                người dùng "thay chữ đúng như ảnh 1 tham chiếu (nút 'Tải tệp
                lên') vào chỗ vị trí ảnh 2 (nút 'Tải Excel')": đổi text (xem
                t.importExcelBtn ở TEXT) TỪ "Tải Excel" (dùng icon mũi tên
                xoay tròn refresh/reload) SANG "Tải tệp lên" — icon cũng đổi
                theo, từ icon xoay tròn sang icon mũi tên hướng lên + gạch
                chân (icon "upload") đúng khớp với icon trong ảnh 1 tham
                chiếu, không còn dùng icon reload cũ (dễ gây hiểu nhầm là nút
                làm mới trang thay vì nút chọn file để tải lên). */}
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexShrink: 0 }}>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                style={{ display: 'none' }}
                onChange={e => {
                  const file = e.target.files?.[0];
                  if (file) handleUploadFile(file);
                }}
              />
              <NeonButton
                className="btn btn-outline btn-sm"
                onClick={() => fileInputRef.current?.click()}
                disabled={isParsing}
                style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px', height: '38px', width: '132px', boxSizing: 'border-box', fontSize: '13px' }}
              >
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15" style={{ flexShrink: 0 }}>
                  <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                  <path d="M17 8l-5-5-5 5" />
                  <path d="M12 3v12" />
                </svg>
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {isParsing ? t.parsing : t.importExcelBtn}
                </span>
              </NeonButton>
            </div>
          </div>
        </div>
        {uploadError && (
          <div style={{ fontSize: 12, color: NEON.pink, marginTop: '-10px', marginBottom: 12 }}>{uploadError}</div>
        )}

        {tab === 'overview' ? (
          <>
            {/* ── KPI Grid — đồng bộ NGUYÊN VĂN kích thước/màu sắc/bố cục với
                4 thẻ KPI của mục 2 (TargetActualDashboard.tsx): dùng chung
                className "kpi-grid"/"kpi-card" (token toàn app), 4 màu viền
                trái #2e7d8c/#10b981/#8b5cf6/#f59e0b + icon SVG + nền
                gradient tương ứng, thay cho component KpiCard tự chế trước
                đây (viền đơn sắc, không icon, không đồng bộ mục 2). ──
                EPCC (menu5-kpi-overview-scale-120pct) - FIX theo yêu cầu
                người dùng "tất cả các chữ và icon tại 4 card này cho tăng
                20% so với hiện tại": className "kpi-card-label/-value/
                -target" và kích thước icon SVG (16px) đều lấy từ CSS TOÀN
                APP (không định nghĩa trong file này — không thể sửa trực
                tiếp px gốc mà không ảnh hưởng các dashboard khác đang dùng
                chung class). Dùng `zoom: 1.2` (phóng to 20% + TỰ GIÃN layout
                theo, khác `transform: scale` vốn không giãn khung chứa nên
                sẽ đè lên card bên cạnh trong lưới 4 cột sát nhau) ngay trên
                CHÍNH div `kpi-grid` này — chỉ áp dụng cho ĐÚNG 4 card ở tab
                "TÌNH HÌNH F-COST" (không đụng lưới KPI khác, VD 4 card tab
                "TÌNH HÌNH DOANH THU" ở dưới vẫn giữ nguyên). */}
            <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '12px', marginBottom: '12px', width: '100%', zoom: 1.2 } as React.CSSProperties}>
              <div className="kpi-card" style={{ borderLeft: '4px solid #2e7d8c', background: 'linear-gradient(135deg, rgba(46,125,140,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#2e7d8c" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
                    <polyline points="3.27 6.96 12 12.01 20.73 6.96" />
                    <line x1="12" y1="22.08" x2="12" y2="12" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi1}</div>
                </div>
                <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px', flexWrap: 'wrap' }}>
                  {/* EPCC (menu5-kpi1-actual-label) - FIX theo yêu cầu người
                      dùng "Card 1 thêm chữ thực tế trước số liệu dòng 2":
                      thêm nhãn {t.kpiActual} (Thực tế/Actual/실제, cùng cỡ
                      chữ/kiểu với "Mục tiêu" bên cạnh) ngay trước số %, tái
                      dùng className "kpi-card-target" có sẵn (không bịa CSS
                      mới) để 2 nhãn "Thực tế"/"Mục tiêu" đồng bộ kiểu chữ. */}
                  <div className="kpi-card-target" style={{ marginBottom: 0 }}>{t.kpiActual}:</div>
                  <div className="kpi-card-value" style={{ marginBottom: 0 }}>
                    {(kpi1AvgFcost * 100).toFixed(2)}
                    <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85, marginLeft: '1px' }}>%</span>
                  </div>
                  <div className="kpi-card-target">{t.kpiTarget}: {kpiTargetTtl != null ? (kpiTargetTtl * 100).toFixed(2) : '—'}%</div>
                </div>
              </div>

              <div className="kpi-card" style={{ borderLeft: '4px solid #10b981', background: 'linear-gradient(135deg, rgba(16,185,129,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <circle cx="12" cy="12" r="10" />
                    <circle cx="12" cy="12" r="6" />
                    <circle cx="12" cy="12" r="2" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi2}</div>
                </div>
                {/* EPCC (menu5-kpi2-scrap-amt) - FIX theo yêu cầu người dùng
                    "Card 2 thay bằng số tiền Scrap lũy kế theo năm (Khi truy
                    vấn năm sẽ hiển thị)": trước đây hiện TỶ LỆ Scrap/Sản xuất
                    (`kpi2ScrapRate`, %). Đổi sang hiện thẳng SỐ TIỀN Scrap
                    (`kpiScrapTtl`, đơn vị K$ — cùng đơn vị với Chart 5/6
                    "F-COST(K$) LŨY KẾ/THEO NĂM" vốn cũng dùng ITEM_SCRAP),
                    lấy ĐÚNG dòng TTL của model đang chọn tại năm đang truy
                    vấn (`lastRangeYear`) — đổi năm ở bộ lọc thì số tự cập
                    nhật theo, giống nguyên tắc "LINK theo năm đang truy vấn"
                    đã áp dụng cho Card 1/3/4 phía trên. */}
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>
                  {kpiScrapTtl != null ? fmtNum(kpiScrapTtl, 1) : '—'}
                  <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85, marginLeft: '1px' }}>K$</span>
                </div>
              </div>

              <div className="kpi-card" style={{ borderLeft: '4px solid #8b5cf6', background: 'linear-gradient(135deg, rgba(139,92,246,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#8b5cf6" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <rect x="2" y="6" width="20" height="12" rx="2" />
                    <circle cx="12" cy="12" r="2" />
                    <line x1="6" y1="12" x2="6.01" y2="12" />
                    <line x1="18" y1="12" x2="18.01" y2="12" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi3}</div>
                </div>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>{kpi3Label}</div>
              </div>

              {/* EPCC (menu5-kpi4-match-ref-image) - FIX theo yêu cầu người
                  dùng "sửa mục Menu5 card 4 như hình" (ảnh: viền/icon/tiêu đề
                  màu xanh dương, dòng "2025: 0.83% → 2026: 0.73%" (năm trước
                  đen/trắng theo theme, năm nay cam), số chênh lệch tô đỏ khi
                  âm/xanh lá khi dương). Đổi màu chủ đạo từ cam (#f59e0b, dễ
                  trùng với số "năm nay" cam bên trong) sang xanh dương
                  #3b82f6 (đã có sẵn trong CHART_PALETTE_9, không bịa màu
                  mới) — áp dụng NHẤT QUÁN cho border-left/icon/tiêu đề để
                  khớp ảnh; số "năm trước" dùng ĐÚNG quy ước theme-aware đen/
                  trắng đã dùng ở LegendDot/ChartCard (menu5-legend-text-
                  theme-aware) để đổi màu đúng khi chuyển theme sáng/tối.
                  FOLLOW-UP (menu5-kpi4-delta-same-line) - người dùng gửi ảnh
                  tham chiếu thứ 2 yêu cầu "di chuyển số chênh lệch lên cùng
                  dòng, bỏ hẳn phía dưới": gộp số chênh lệch (`kpi4DeltaPp`)
                  vào CHUNG 1 hàng flex với "năm trước → năm nay" (ngăn cách
                  bằng dấu "|" mờ, theme-aware như phần chữ đen/trắng), xoá
                  hẳn `<div className="kpi-card-value">` dòng thứ 2 cũ. */}
              <div className="kpi-card" style={{ borderLeft: '4px solid #3b82f6', background: 'linear-gradient(135deg, rgba(59,130,246,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#3b82f6" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <polyline points="23 6 13.5 15.5 8.5 10.5 1 18" />
                    <polyline points="17 6 23 6 23 12" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0, color: '#3b82f6' }}>{t.kpi4}</div>
                </div>
                <div style={{ display: 'flex', alignItems: 'baseline', gap: 5, flexWrap: 'wrap', fontSize: 13, fontWeight: 700 }}>
                  <span style={{ color: theme === 'light' ? '#000000' : '#ffffff' }}>
                    {lastRangeYear - 1}: {prevYearFcostTtl != null ? (prevYearFcostTtl * 100).toFixed(2) : '—'}%
                  </span>
                  <span style={{ color: theme === 'light' ? '#000000' : '#ffffff' }}>→</span>
                  <span style={{ color: '#f59e0b' }}>
                    {lastRangeYear}: {(kpi1AvgFcost * 100).toFixed(2)}%
                  </span>
                  {kpi4DeltaPp != null && (
                    <>
                      <span style={{ color: theme === 'light' ? '#000000' : '#ffffff', opacity: 0.5 }}>|</span>
                      <span style={{ color: kpi4DeltaPp <= 0 ? '#ef4444' : '#10b981', fontSize: '20px' }}>
                        {kpi4DeltaPp >= 0 ? '+' : ''}{kpi4DeltaPp.toFixed(2)}
                        <span style={{ fontSize: '75%', fontWeight: 700, opacity: 0.9, marginLeft: '1px' }}>%</span>
                      </span>
                    </>
                  )}
                </div>
              </div>
            </div>

            {/* EPCC (menu5-4-new-charts) - 4 biểu đồ mới thay cho 6 biểu đồ cũ đã xoá,
                nhân bản theo hình mẫu người dùng cung cấp. Chart 1 & 2 dùng modelData
                của model đang chọn (bộ lọc MODEL); Chart 3 & 4 vẫn là dữ liệu TỔNG HỢP
                TOÀN BỘ MODEL (cross-model) — xem chi tiết mapping/quy tắc ở khối tính
                toán "== 4 BIỂU ĐỒ MỚI ==". */}
            {/* EPCC (menu5-2x2-fixed-grid) - FIX ROOT CAUSE "biểu đồ 3 bị đẩy
                lên chung hàng với 1,2 thay vì xuống hàng với 4": grid cũ dùng
                `repeat(auto-fit, minmax(420px, 1fr))` — số cột phụ thuộc độ
                rộng container, nên khi đủ chỗ cho 3 cột (>= ~1260px) trình
                duyệt tự xếp 3 khung/hàng (1,2,3 rồi 4 lẻ xuống dưới) thay vì
                2x2 như mong muốn. Đổi sang `repeat(2, minmax(0, 1fr))` — LUÔN
                đúng 2 cột bất kể độ rộng màn hình, đảm bảo 1,2 luôn hàng trên
                và 3,4 luôn hàng dưới. */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: 14 }}>
              <ChartCard
                title={t.nc1Title}
                accent={NEON.cyan}
                idx={0}
                theme={theme}
                headerExtra={chart1.bars.some(b => b.value != null) ? (
                  <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap', alignItems: 'center' }}>
                    <LegendDot color={SALES_CHART1_COLORS.bar} label={t.legendShipment} theme={theme} />
                    {chart1.barsWon.some(v => v != null) && <LegendDot color={NEON.pink} label={t.legendFcostWon} theme={theme} />}
                    <LegendDot color={SALES_CHART1_COLORS.lineDashed} label={t.legendTarget} theme={theme} />
                    <LegendDot color={SALES_CHART1_COLORS.line} label={t.legendActual} theme={theme} />
                  </div>
                ) : undefined}
              >
                {chart1.bars.some(b => b.value != null) ? (
                  <ComboBarDualLineChart
                    bars={chart1.bars} bars2={chart1.barsWon} bar2Color={NEON.pink}
                    target={chart1.target} actual={chart1.actual}
                    barColor={SALES_CHART1_COLORS.bar} targetColor={SALES_CHART1_COLORS.lineDashed} actualColor={SALES_CHART1_COLORS.line}
                    height={300}
                    theme={theme}
                  />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard
                title={t.nc2Title}
                accent={NEON.pink}
                idx={1}
                theme={theme}
                headerExtra={chart2.series.length ? (
                  // EPCC (menu5-legend-wrap-tighten-rowgap) - FIX theo yêu cầu
                  // người dùng "di chuyển dòng khoanh đỏ lên phía trên": trước
                  // đây `gap: 14` áp DÙNG CHUNG cho cả khoảng cách ngang (giữa
                  // các chip cùng dòng) LẪN dọc (giữa dòng 1 và dòng chip bị
                  // tràn xuống khi không đủ chỗ, VD "SO1C2EDM" ở ảnh người
                  // dùng khoanh đỏ) — khiến dòng tràn bị đẩy cách xa dòng 1 một
                  // khoảng lớn y hệt khoảng cách ngang giữa các chip, nhìn như
                  // 2 hàng tách biệt. Tách riêng `columnGap` (giữ 14, khoảng
                  // cách ngang giữa các chip không đổi) và `rowGap` (giảm còn
                  // 2px) — dòng chip tràn giờ nằm SÁT ngay dưới dòng 1.
                  <div style={{ display: 'flex', columnGap: 14, rowGap: 2, flexWrap: 'wrap', alignItems: 'center' }}>
                    {chart2.series.map((s, i) => (
                      <LegendDot key={i} color={s.color} label={s.label} theme={theme} />
                    ))}
                  </div>
                ) : undefined}
              >
                {chart2.series.length ? (
                  <MultiLineChart categories={chart2.categories} series={chart2.series} height={280} theme={theme} maxV={0.035} />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              {/* EPCC (menu5-swap-chart34-chart56) - FIX theo yêu cầu người
                  dùng "di chuyển đổi Chart 3,4 <--> 5,6 cho nhau": đổi thứ tự
                  HIỂN THỊ trong lưới — Chart 5,6 (Mass/SPL F-Cost%) lên vị
                  trí 3,4 ngay sau Chart 2; Chart 3,4 (Rose Scrap) xuống vị
                  trí 5,6. Không đổi tên/tiêu đề/nguồn dữ liệu của từng chart,
                  chỉ đổi vị trí khối JSX + cập nhật `idx` theo ĐÚNG vị trí
                  mới (idx quyết định màu header cycle theo CHART_HEADER_THEME,
                  gắn với vị trí trong lưới chứ không phải với tên chart). */}
              <ChartCard
                title={t.nc5Title}
                accent={NEON.cyan}
                idx={2}
                theme={theme}
                headerExtra={chart5.series.length ? (
                  <div style={{ display: 'flex', columnGap: 14, rowGap: 2, flexWrap: 'wrap', alignItems: 'center' }}>
                    {chart5.series.map((s, i) => (
                      <LegendDot key={i} color={s.color} label={s.label} theme={theme} />
                    ))}
                  </div>
                ) : undefined}
              >
                {chart5.series.length ? (
                  <MultiLineChart categories={chart5.categories} series={chart5.series} height={280} theme={theme} maxV={0.05} />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard
                title={t.nc6Title}
                accent={NEON.pink}
                idx={3}
                theme={theme}
                headerExtra={chart6.series.length ? (
                  <div style={{ display: 'flex', columnGap: 14, rowGap: 2, flexWrap: 'wrap', alignItems: 'center' }}>
                    {chart6.series.map((s, i) => (
                      <LegendDot key={i} color={s.color} label={s.label} theme={theme} />
                    ))}
                  </div>
                ) : undefined}
              >
                {chart6.series.length ? (
                  <MultiLineChart categories={chart6.categories} series={chart6.series} height={280} theme={theme} />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard title={t.nc3Title} subtitle={t.nc3Sub} accent={NEON.cyan} idx={4} theme={theme}>
                {chart3Items.some(it => it.value != null) ? (
                  // EPCC (menu5-rose-color-ref-img-v2) - FIX theo yêu cầu
                  // người dùng "thay màu cho vùng khoanh đỏ sang ảnh màu 1":
                  // đổi màu cánh hoa Chart 3 sang đúng mã màu đo từ ảnh
                  // swatch người dùng cung cấp (#01877E), thay cho màu
                  // placeholder cũ (#70C2B4). Chart 4 (amber) không liên
                  // quan, giữ nguyên.
                  // EPCC (menu5-rose-color-ref-img-v3) - FIX theo yêu cầu
                  // người dùng "Chuyển màu xanh ảnh 2 (Chart 3) bằng ảnh 1
                  // tham chiếu": ảnh tham chiếu mới cung cấp là swatch mã HEX
                  // #851639 (đỏ mận/burgundy) — đổi cánh hoa Chart 3 (cũ xanh
                  // teal #01877E) sang đúng mã màu này. Chart 4 (amber
                  // #EE6457) không nằm trong yêu cầu, giữ nguyên.
                  <RoseChart items={chart3Items} barColor="#851639" rotateLabels90 />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard title={t.nc4Title} subtitle={t.nc4Sub} accent={NEON.amber} idx={5} theme={theme}>
                {chart4Items.some(it => it.value != null) ? (
                  // EPCC (menu5-rose-color-ref-img-v3) - FIX theo yêu cầu
                  // người dùng "thay màu cho vùng khoanh đỏ sang ảnh màu 1":
                  // đổi màu cánh hoa Chart 4 sang đúng mã màu đo từ ảnh
                  // swatch người dùng cung cấp (#EE6457), thay cho màu amber
                  // (NEON.amber) cũ. "Khoảng trống trắng cố định" giữa label
                  // ring và cánh hoa đã có sẵn từ trước (chung component
                  // RoseChart với Chart 3, xem menu5-rose-white-hub-fixed-theme),
                  // không cần sửa gì thêm cho phần đó.
                  <RoseChart items={chart4Items} barColor="#EE6457" rotateLabels90 />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              {/* Chart 7 — biểu đồ NHỆN (radar/spider), NHÂN BẢN Y HỆT DỮ LIỆU
                  Chart 2 (chart2.categories/chart2.series, cùng model đang
                  chọn + chế độ Tháng/Quý/Năm), chỉ đổi kiểu vẽ từ đường sang
                  radar. Chiếm trọn 2 cột (gridColumn: '1 / -1') vì radar cần
                  không gian vuông để không bị bóp méo. */}
              <div style={{ gridColumn: '1 / -1' }}>
                <ChartCard
                  title={t.nc7Title}
                  accent={NEON.violet}
                  idx={6}
                  theme={theme}
                  headerExtra={chart2.series.length ? (
                    <div style={{ display: 'flex', columnGap: 14, rowGap: 2, flexWrap: 'wrap', alignItems: 'center' }}>
                      {chart2.series.map((s, i) => (
                        <LegendDot key={i} color={s.color} label={s.label} theme={theme} />
                      ))}
                    </div>
                  ) : undefined}
                >
                  {chart2.series.length ? (
                    <div style={{ display: 'flex', justifyContent: 'center' }}>
                      <div style={{ width: '100%', maxWidth: 520 }}>
                        <RadarChart categories={chart2.categories} series={chart2.series} size={480} theme={theme} maxV={0.035} />
                      </div>
                    </div>
                  ) : <NoData text={t.noData} />}
                </ChartCard>
              </div>
            </div>
          </>
        ) : tab === 'revenue' ? (
          <>
            {/* ═══ TÌNH HÌNH DOANH THU — EPCC (menu5-revenue-tab) ═══
                Nhân bản y hệt bố cục/màu KPI + 4-chart-grid + bảng của
                SalesDashboard.tsx (Mục 1), dữ liệu lấy từ allModelsAggregate/
                salesByCustomer (đã tính ở trên, nguồn Test5.xlsx). CHỦ Ý
                không phụ thuộc bộ lọc MODEL của toolbar (4 chart chi tiết bên
                dưới + bảng vẫn luôn là góc nhìn TOÀN CÔNG TY, không lọc theo
                model/khách hàng cụ thể).
                EPCC (menu5-revenue-kpi-link-year-filter) - RIÊNG 4 THẺ KPI
                phía trên (revenueKpis) đã được nối lại với bộ lọc "NĂM"/
                "THÁNG BẮT ĐẦU"/"THÁNG KẾT THÚC" — không còn cố định cộng dồn
                TOÀN BỘ 2017-2026 như trước nữa, xem chi tiết ở khai báo
                `revenueKpis` phía trên. */}
            <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '12px', marginBottom: '12px', width: '100%' }}>
              <div className="kpi-card" style={{ borderLeft: '4px solid #2e7d8c', background: 'linear-gradient(135deg, rgba(46,125,140,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#2e7d8c" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <polyline points="23 6 13.5 15.5 8.5 10.5 1 18" />
                    <polyline points="17 6 23 6 23 12" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.revKpiSales}</div>
                </div>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>{fmtNum(revenueKpis.totalSales, 0)} <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85 }}>K$</span></div>
              </div>

              <div className="kpi-card" style={{ borderLeft: '4px solid #10b981', background: 'linear-gradient(135deg, rgba(16,185,129,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
                    <polyline points="3.27 6.96 12 12.01 20.73 6.96" />
                    <line x1="12" y1="22.08" x2="12" y2="12" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.revKpiShipment}</div>
                </div>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>{fmtNum(revenueKpis.totalShipment, 0)} <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85 }}>Kea</span></div>
              </div>

              <div className="kpi-card" style={{ borderLeft: '4px solid #8b5cf6', background: 'linear-gradient(135deg, rgba(139,92,246,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#8b5cf6" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <path d="M2 20a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V8l-7 5V8l-7 5V4l-6 5Z" />
                    <path d="M17 18h1" /><path d="M12 18h1" /><path d="M7 18h1" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.revKpiProduction}</div>
                </div>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>{fmtNum(revenueKpis.totalProduction, 0)} <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85 }}>Kea</span></div>
              </div>

              <div className="kpi-card" style={{ borderLeft: '4px solid #f59e0b', background: 'linear-gradient(135deg, rgba(245,158,11,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.revKpiYoy}</div>
                </div>
                {/* EPCC (menu5-revkpi-yoy-label-same-line) - FIX theo yêu cầu
                    người dùng "di chuyển dòng dưới lên trên trong card này để
                    chỉ còn 1 dòng thôi": trước đây `%` (kpi-card-value) và
                    nhãn "20xx so với 20yy" (kpi-card-target) là 2 <div> khối
                    (block) TÁCH RIÊNG, tự động xuống 2 hàng khác nhau. Gộp cả
                    2 vào chung 1 <div> flex (`alignItems: 'baseline'` để số %
                    to và chữ nhãn nhỏ canh đúng theo đường chữ cái, không
                    lệch trên/dưới) — nhãn giờ nằm NGAY BÊN PHẢI số %, cùng 1
                    dòng, không còn dòng thứ 2 riêng bên dưới nữa. */}
                <div className="kpi-card-value" style={{ marginBottom: 0, display: 'flex', alignItems: 'baseline', gap: 6, flexWrap: 'wrap' }}>
                  <span>
                    {revenueKpis.yoyPct != null ? `${revenueKpis.yoyPct >= 0 ? '+' : ''}${revenueKpis.yoyPct.toFixed(1)}` : '—'}
                    {revenueKpis.yoyPct != null && <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85, marginLeft: '1px' }}>%</span>}
                  </span>
                  {revenueKpis.yoyLabel && <span className="kpi-card-target" style={{ marginBottom: 0 }}>{revenueKpis.yoyLabel}</span>}
                </div>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
              <ChartCard
                title={t.revChart1Title}
                headerExtra={
                  <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    <LegendDot color="#2d7f96" label={t.revLegendShip} theme={theme} />
                    <LegendDot color="#F17988" label={t.revLegendProd} theme={theme} />
                    <LegendDot color="#00a65a" label={t.revLegendSales} theme={theme} />
                    <LegendDot color="#f39c12" label={t.revLegendYoy} theme={theme} />
                  </div>
                }
                accent={NEON.violet} idx={0} theme={theme}
              >
                {revenueChart1.labels.length ? (
                  <RevenueTrendChart labels={revenueChart1.labels} shipVals={revenueChart1.shipVals}
                    prodVals={revenueChart1.prodVals}
                    salesVals={revenueChart1.salesVals} yoyVals={revenueChart1.yoyVals} theme={theme} />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard title={t.revChart2Title} accent={NEON.cyan} idx={1} theme={theme}>
                {revenueTopModels.length ? (
                  <RevenueTopModelBarChart items={revenueTopModels} theme={theme} />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard
                title={t.revChart3Title}
                headerExtra={
                  revenueChart3.series.length ? (
                    <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                      {/* EPCC (menu5-revchart3-legend-autohide-no-data) - FIX
                          theo yêu cầu người dùng "khi lựa chọn model, tháng,
                          năm nếu không có nhóm Custom nào thì hãy để chế độ
                          tự ẩn đi, khi nào có lại hiện ra": trước đây legend
                          LUÔN liệt kê ĐỦ mọi khách hàng có trong toàn bộ
                          Test5.xlsx (`revenueCustomersFiltered`), bất kể
                          khách hàng đó có SALES AMT > 0 ở đúng kỳ Tháng/Quý/
                          Năm đang chọn (`viewBy` + `rangeYears`/
                          `lastRangeYear`) hay không — dẫn tới các chip như
                          HUAWEI/SOMC/XIAOMI vẫn hiện dù cột biểu đồ hoàn
                          toàn không có phần màu tương ứng (khách hàng đó
                          không phát sinh doanh số trong kỳ đang xem). Thêm
                          `.filter(s => s.values.some(v => v > 0))` NGAY
                          TRƯỚC khi map ra `LegendDot` — chỉ giữ lại khách
                          hàng có ÍT NHẤT 1 giá trị > 0 trong mảng `values`
                          (đã tự ứng với kỳ đang chọn, tính lại mỗi khi đổi
                          bộ lọc qua `useMemo` của `revenueChart3`) — khách
                          hàng không có dữ liệu tự ẩn khỏi legend, và tự hiện
                          lại ngay khi đổi bộ lọc sang kỳ có dữ liệu của họ,
                          không cần đụng gì tới phần vẽ cột chồng bên dưới
                          (khách hàng 0 giá trị vốn đã không chiếm phần màu
                          nào trên cột, chỉ riêng legend là hiện thừa). */}
                      {revenueChart3.series.filter(s => s.values.some(v => v > 0)).map(s => <LegendDot key={s.label} color={s.color} label={s.label} theme={theme} />)}
                    </div>
                  ) : undefined
                }
                accent={NEON.amber} idx={2} theme={theme} compact
              >
                {revenueChart3.labels.length && revenueChart3.series.length ? (
                  <RevenueCustomerStackedChart labels={revenueChart3.labels} series={revenueChart3.series} theme={theme} />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard
                title={t.revChart4Title}
                headerExtra={
                  revenueChart4.length ? (
                    <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                      {revenueChart4.map(it => <LegendDot key={it.label} color={it.color} label={it.label} theme={theme} />)}
                    </div>
                  ) : undefined
                }
                accent="#EE6457" idx={3} theme={theme} compact
              >
                {/* EPCC (menu5-revchart4-top-align) - FIX theo yêu cầu người
                    dùng "di chuyển để top=0 ở phần vẽ": trước đây
                    `alignItems: 'center'` canh GIỮA donut lẫn panel YoY theo
                    chiều dọc trong hàng flex — do donut (320px) cao hơn hẳn
                    panel YoY (title + svg thấp hơn), việc canh giữa vô tình
                    đẩy panel YoY (và cả donut) lệch xuống dưới, để lại
                    khoảng trống phía trên như ảnh người dùng khoanh mũi tên
                    đỏ. Đổi sang `alignItems: 'flex-start'` — cả donut và
                    panel giờ neo thẳng mép TRÊN CÙNG (top=0) của hàng, không
                    còn khoảng trống thừa phía trên nữa. */}
                <div style={{ display: 'flex', gap: 8, alignItems: 'flex-start' }}>
                  <div style={{ flex: 1, minWidth: 0, display: 'flex', justifyContent: 'flex-start', alignItems: 'flex-start', padding: 0 }}>
                    {revenueChart4.length ? (
                      <RevenueDonutChart items={revenueChart4} size={320} theme={theme} />
                    ) : <NoData text={t.noData} />}
                  </div>
                  {revenueYoyByYear.some(r => r.diff != null) && (
                    <YoyChangePanel rows={revenueYoyByYear} title={t.revYoyPanelTitle} theme={theme} />
                  )}
                </div>
              </ChartCard>

              {/* EPCC (menu5-revchart5-6-rtv) - Chart 5/6 THÊM MỚI, nhân bản
                  đúng bố cục "RTV THEO THÁNG NĂM" / "RTV THEO MODEL" ở ảnh
                  tham chiếu người dùng cung cấp. `idx` tiếp tục 4,5 để màu
                  header CHART_HEADER_THEME tự cuộn mượt về lại violet/cyan
                  (idx % 4), khớp đúng chu kỳ 4 màu đã dùng cho Chart 1-4 —
                  không cần khai báo thêm màu mới, chuyển theme dark/light
                  cũng tự ăn theo `theme` như 4 chart còn lại. */}
              <ChartCard
                title={t.revChart5Title}
                headerExtra={
                  <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    <LegendDot color={SALES_CHART1_COLORS.bar} label={t.revLegendShipKea} theme={theme} />
                    <LegendDot color={SALES_CHART1_COLORS.lineDashed} label={t.revLegendRtvKea} theme={theme} />
                    <LegendDot color={SALES_CHART1_COLORS.line} label={t.revLegendRtvRate} theme={theme} />
                  </div>
                }
                accent={NEON.violet} idx={4} theme={theme} compact
              >
                {revenueChart5.labels.length ? (
                  <RevenueRtvComboChart labels={revenueChart5.labels} shipVals={revenueChart5.shipVals}
                    rtvVals={revenueChart5.rtvVals} ratePct={revenueChart5.ratePct} theme={theme} />
                ) : <NoData text={t.noData} />}
              </ChartCard>

              <ChartCard
                title={t.revChart6Title}
                headerExtra={
                  <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    <LegendDot color={SALES_CHART1_COLORS.bar} label={t.revLegendShipKea} theme={theme} />
                    <LegendDot color={SALES_CHART1_COLORS.lineDashed} label={t.revLegendRtvKea} theme={theme} />
                    <LegendDot color={SALES_CHART1_COLORS.line} label={t.revLegendRtvRate} theme={theme} />
                  </div>
                }
                accent={NEON.cyan} idx={5} theme={theme} compact
              >
                {revenueChart6.length ? (
                  <RevenueRtvComboChart
                    labels={revenueChart6.map(r => r.label)}
                    shipVals={revenueChart6.map(r => r.shipVal)}
                    rtvVals={revenueChart6.map(r => r.rtvVal)}
                    ratePct={revenueChart6.map(r => r.ratePct)}
                    theme={theme}
                  />
                ) : <NoData text={t.noData} />}
              </ChartCard>
            </div>
          </>
        ) : (
          <>
            {/* EPCC (menu5-detail-merge-revenue-table) - FIX theo yêu cầu
                người dùng "gộp luôn dữ liệu tại sheet tình hình doanh thu
                sang sheet chi tiết và dữ liệu thêm phần Scrap và F-cost%
                hiển thị ra nữa": bảng đơn giản cũ ở tab này (chỉ 1 model/
                tháng, dùng `shipmentPts`/`prodPts`/`salesPts`/`scrapPts`/
                `fcostByYm`) đã được THAY THẾ bằng bảng tổng hợp "Toàn bộ
                Model theo Doanh số" (trước đây chỉ hiện ở tab TÌNH HÌNH
                DOANH THU) — chuyển nguyên khối filter bar + bảng sang ĐÂY,
                không còn hiện ở tab Doanh thu nữa (tab đó giờ chỉ còn 6
                biểu đồ). Bảng này vốn ĐÃ có cột "RTV (K)" (chính là cột
                SCRAP trong Excel, dùng chung 1 nguồn `ITEM_SCRAP` — không
                tạo thêm cột trùng lặp cùng số liệu dưới 2 tên khác nhau),
                nay bổ sung thêm cột "F-Cost %" (tính TTL theo năm khi xem
                Năm, TRUNG BÌNH các tháng trong quý khi xem Quý — vì đây là
                TỶ LỆ, không cộng dồn được như số lượng — hoặc giá trị đúng
                tháng khi xem Tháng) để đủ 6 cột gốc của bảng cũ
                (SHIPMENT/PROD'N/SALES AMT/SCRAP/F-COST%) cộng thêm phần
                Model/Khách hàng/Type/Tăng trưởng/%Rate vốn có sẵn. */}
            <div style={{ background: 'var(--surface)', border: '1px solid var(--border-soft)', borderRadius: 12, padding: '10px 14px', display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'center', justifyContent: 'space-between' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontWeight: 700, fontSize: 13, whiteSpace: 'nowrap' }}>
                <span>{t.revTableTitle}</span>
                <span style={{ color: '#22c55e', fontWeight: 700, fontSize: 12 }}>{formattedTime}</span>
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
                {[
                  { label: t.revFilterModelLabel, value: revTableModel, onChange: setRevTableModel,
                    options: [{ value: ALL_MODEL_VALUE, label: t.modelFilterAll }, ...realModelsForAgg.map(m => ({ value: m.model, label: m.model }))] },
                  { label: t.revFilterYearLabel, value: revTableYear, onChange: setRevTableYear,
                    options: [{ value: ALL_CUSTOM_VALUE, label: t.yearFilterAll }, ...revenueYears.slice().reverse().map(y => ({ value: String(y), label: String(y) }))] },
                  { label: t.revFilterQuarterLabel, value: revTableQuarter, onChange: setRevTableQuarter,
                    options: [{ value: ALL_CUSTOM_VALUE, label: t.yearFilterAll }, ...[1, 2, 3, 4].map(q => ({ value: String(q), label: `Q${q}` }))] },
                  { label: t.revFilterMonthLabel, value: revTableMonth, onChange: setRevTableMonth,
                    options: [{ value: ALL_CUSTOM_VALUE, label: t.yearFilterAll }, ...Array.from({ length: 12 }, (_, i) => i + 1).map(mo => ({ value: String(mo), label: String(mo) }))] },
                  { label: t.revFilterCustomerLabel, value: revTableCustomer, onChange: setRevTableCustomer,
                    options: [{ value: ALL_CUSTOM_VALUE, label: t.customFilterAll }, ...revenueCustomers.map(c => ({ value: c, label: c }))] },
                ].map(f => (
                  <label key={f.label} style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, color: 'var(--text-2)', whiteSpace: 'nowrap' }}>
                    {f.label}
                    <select value={f.value} onChange={e => f.onChange(e.target.value)}
                      style={{ height: 30, borderRadius: 6, border: '1px solid var(--border-soft)', background: 'var(--surface)', color: 'var(--text-0)', padding: '0 6px', fontSize: 12 }}>
                      {f.options.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
                    </select>
                  </label>
                ))}
                <input
                  value={revTableSearch}
                  onChange={e => setRevTableSearch(e.target.value)}
                  placeholder={t.revFilterSearchPlaceholder}
                  style={{ height: 30, minWidth: 200, borderRadius: 6, border: '1px solid var(--border-soft)', background: 'var(--surface)', color: 'var(--text-0)', padding: '0 8px', fontSize: 12 }}
                />
              </div>
            </div>

            <div style={{ marginTop: 8, background: 'var(--surface)', border: '1px solid var(--border-soft)', borderRadius: 12, overflow: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <thead>
                  <tr style={{ background: 'var(--border-soft)' }}>
                    <th style={{ padding: '8px 10px', textAlign: 'left' }}>{t.revColModel}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'left' }}>{t.revColPeriod}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'left' }}>{t.revColCustomer}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'left' }}>{t.revColType}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'right' }}>{t.revColProd}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'right' }}>{t.revColShip}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'right' }}>{t.revColSales}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'right' }}>{t.revColGrowth}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'right' }}>{t.revColRtv}</th>
                    <th style={{ padding: '8px 10px', textAlign: 'right' }}>{t.revColRtvRate}</th>
                    {/* EPCC (menu5-detail-add-scrap-column-before-fcost) -
                        THÊM MỚI cột "SCRAP (K)" theo yêu cầu người dùng "còn
                        dòng Scrap chưa thấy hiển thị ra, hãy bổ sung thêm
                        phía trước bên trái cột F-cost%": dùng ĐÚNG giá trị
                        `row.rtv` (cùng nguồn `ITEM_SCRAP`, đã tính sẵn ở cột
                        RTV (K) phía trước) — không tính lại, chỉ hiển thị
                        thêm 1 lần dưới đúng tên gốc "SCRAP" trong Excel để
                        khớp đúng thuật ngữ người dùng quen dùng. */}
                    <th style={{ padding: '8px 10px', textAlign: 'right' }}>{t.revColScrap}</th>
                    {/* EPCC (menu5-detail-merge-revenue-table-add-fcost) -
                        THÊM MỚI cột F-Cost % theo yêu cầu người dùng, tô màu
                        theo `revTableGranularity` (Tháng/Quý/Năm) y hệt bảng
                        cũ (`fcostColumnColor` trước đây tính theo `viewBy`
                        của toolbar chính — nay đổi sang tính theo bộ lọc
                        RIÊNG của bảng này cho nhất quán với chính bảng). */}
                    <th style={{ padding: '8px 10px', textAlign: 'right', color: revTableFcostColor, fontWeight: 700 }}>{t.colFcost}</th>
                  </tr>
                </thead>
                <tbody>
                  {revenueTopModelsTable.length ? revenueTopModelsTable.map(row => (
                    <tr key={row.key} style={{ borderTop: '1px solid var(--border-soft)' }}>
                      <td style={{ padding: '7px 10px', fontWeight: 700 }}>{row.model}</td>
                      <td style={{ padding: '7px 10px' }}>{row.period}</td>
                      <td style={{ padding: '7px 10px' }}>{row.customer}</td>
                      <td style={{ padding: '7px 10px' }}>{row.type}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right' }}>{fmtNum(row.prod, 0)}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right' }}>{fmtNum(row.ship, 0)}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right', fontWeight: 700 }}>{fmtNum(row.sales, 0)}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right', fontWeight: 700, color: row.growth == null ? 'var(--text-2)' : row.growth >= 0 ? '#22c55e' : NEON.pink }}>
                        {row.growth != null
                          ? `${row.growth >= 0 ? '+' : ''}${fmtNum(row.growth, 0)}${row.growthPct != null ? ` (${row.growthPct >= 0 ? '+' : ''}${row.growthPct.toFixed(1)}%)` : ''}`
                          : '—'}
                      </td>
                      <td style={{ padding: '7px 10px', textAlign: 'right' }}>{fmtNum(row.rtv, 0)}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right' }}>{row.rtvRatePct != null ? `${row.rtvRatePct.toFixed(1)}%` : '—'}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right' }}>{fmtNum(row.rtv, 0)}</td>
                      <td style={{ padding: '7px 10px', textAlign: 'right', color: revTableFcostColor, fontWeight: 700 }}>{row.fcostPct != null ? `${(row.fcostPct * 100).toFixed(2)}%` : '—'}</td>
                    </tr>
                  )) : (
                    <tr><td colSpan={12} style={{ padding: 16, textAlign: 'center' }}><NoData text={t.noData} /></td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </>
        )}
    </div>
  );
}

// EPCC (menu5-chart-header-match-sales) - FIX theo yêu cầu người dùng "áp
// dụng đúng thanh khoanh đỏ ảnh 1 (SalesDashboard) cho menu5db": thêm thanh
// nền màu (panel-head) phía trên mỗi ChartCard, dùng ĐÚNG bảng màu
// CHART_HEADER_THEME (bgLight/bgDark theo idx) + viền trái 4px accent —
// y hệt hàm chartHeaderStyle() trong SalesDashboard.tsx, thay cho header
// cũ chỉ là 1 hàng flex không nền. `idx` chọn màu theo đúng thứ tự 1-4;
// `theme` quyết định dùng bgLight hay bgDark.
// EPCC (menu5-chart1-2-legend-into-header) - FIX theo yêu cầu người dùng
// "đường gạch chéo (subtitle) xóa bỏ đi và di chuyển vùng khoanh đỏ (legend)
// vào tab top để biểu đồ phóng to hết cỡ": thêm prop `headerExtra` — nội
// dung tuỳ ý (ở đây là legend) đặt vào ĐÚNG vị trí subtitle cũ (góc phải
// thanh tiêu đề), thay vì subtitle text. Chart 1 & 2 dùng `headerExtra`
// (legend), không còn dùng `subtitle`; chart 3 & 4 vẫn dùng `subtitle` như
// cũ (không có legend, không cần đổi). Nhờ đó phần thân card giải phóng hẳn
// hàng legend cũ → chart bên dưới có thêm không gian để phóng to.
function ChartCard({
  title, subtitle, headerExtra, accent: _accent, idx, theme, children, compact,
}: {
  title: string; subtitle?: string; headerExtra?: React.ReactNode; accent: string; idx: number; theme: ThemeMode; children: React.ReactNode;
  // EPCC (menu5-chart4-5-6-title-one-line-compact-header) - THÊM MỚI prop
  // `compact` (mặc định false, KHÔNG đụng tới các ChartCard khác đang không
  // truyền prop này) theo yêu cầu người dùng "Cardtop cho di chuyển chữ như
  // hình vẽ thành 1 dòng và thu gọn thanh Card để biểu đồ phóng to rộng
  // hơn" cho đúng 3 card có mũi tên đỏ trong ảnh (TỶ TRỌNG THEO KHÁCH HÀNG /
  // RTV THEO THÁNG NĂM / RTV THEO MODEL). Giờ title của 3 card này đã gộp về
  // 1 dòng (bỏ `\n`, xem TEXT.revChart4/5/6Title) nên thanh header không còn
  // cần chiều cao 2 dòng như trước — `compact=true` giảm padding dọc + margin
  // dưới của header, nhường thêm không gian dọc cho phần vẽ chart bên dưới.
  compact?: boolean;
}) {
  const hc = CHART_HEADER_THEME[idx % CHART_HEADER_THEME.length];
  const headerStyle: React.CSSProperties = {
    background: theme === 'light' ? hc.bgLight : hc.bgDark,
    borderLeft: `4px solid ${hc.accent}`,
    borderRadius: '8px 8px 0 0',
    padding: compact ? '6px 14px' : '10px 14px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'flex-start',
    // EPCC (menu5-legend-same-row-as-title) - FIX theo yêu cầu người dùng
    // "chưa di chuyển dòng chữ các model lên cùng với hàng F-COST THEO
    // MODEL": trước đây header dùng `flexWrap: 'wrap'` ở CHÍNH cấp ngoài
    // cùng (chứa cả tiêu đề lẫn khối legend) — khi tổng bề rộng tiêu đề +
    // toàn bộ legend vượt quá bề rộng card, TOÀN BỘ khối legend (headerExtra)
    // bị đẩy xuống thành 1 HÀNG MỚI riêng, chiếm hết chiều rộng card từ mép
    // trái (không còn nằm cạnh tiêu đề nữa) — đúng hiện tượng trong ảnh
    // người dùng khoanh đỏ + mũi tên "kéo lên". Đổi `flexWrap` cấp ngoài
    // sang 'nowrap' (không cho phép cả khối legend rớt dòng), tiêu đề được
    // đánh dấu `flexShrink: 0` (giữ nguyên kích thước, không bị co bóp), còn
    // khối headerExtra được bọc trong 1 div `flex: 1, minWidth: 0` riêng
    // (xem bên dưới) — nhờ vậy legend LUÔN bắt đầu ngay bên phải tiêu đề,
    // TRÊN CÙNG 1 HÀNG với nó; nếu không đủ chỗ cho hết các model, chỉ các
    // chip THỪA mới tự xuống dòng bên trong vùng riêng của legend (không còn
    // kéo theo cả khối tụt xuống dưới tiêu đề như trước).
    flexWrap: 'nowrap',
    gap: 8,
    margin: compact ? '-14px -14px 6px -14px' : '-14px -14px 10px -14px',
  };
  return (
    <div style={{ background: 'var(--surface)', border: '1px solid var(--border-soft)', borderRadius: 12, padding: 14 }}>
      <div style={headerStyle}>
        {/* EPCC (menu5-header-title-color-theme-aware) - FIX theo yêu cầu
            người dùng "màu chữ tiêu đề các thanh ALL YEAR khi chuyển đổi
            theme cũng đổi màu dễ nhìn hơn": trước đây màu chữ tiêu đề cố
            định `#000000` bất kể theme — ở theme dark, nền header
            (`hc.bgDark`) chỉ là 1 lớp phủ TRONG SUỐT rất nhẹ lên trên nền
            tối `var(--surface)` (gần như đen), khiến chữ đen gần như không
            đọc được. Đổi sang tính theo `theme`: giữ nguyên đen ở theme
            light (nền `hc.bgLight` là màu pastel sáng, chữ đen vẫn rõ), đổi
            sang trắng ở theme dark (nền tối, chữ trắng tương phản rõ hơn). */}
        {/* EPCC (menu5-title-line-break-model) - FIX theo yêu cầu người dùng
            "cho chữ THEO MODEL xuống dòng ngay dưới cho cả 3 chart": trước
            đây title LUÔN `whiteSpace: 'nowrap'` (bắt buộc 1 dòng) — Chart
            2/5/6 (F-Cost%/Mass/SPL theo model) đã được chèn sẵn ký tự xuống
            dòng `\n` ngay trong chuỗi title (TEXT.nc2Title/nc5Title/
            nc6Title, đủ 3 ngôn ngữ) để tách phần "THEO MODEL"/"BY MODEL"/
            "모델별" xuống dòng riêng. Đổi `whiteSpace` sang `'pre-line'` để
            tôn trọng `\n` đó — các tiêu đề KHÁC (Chart 1,3,4...) không chứa
            `\n` nên vẫn hiển thị nguyên 1 dòng như trước, không bị ảnh
            hưởng. */}
        <span style={{ fontSize: 14.4, fontWeight: 800, color: theme === 'light' ? '#000000' : '#ffffff', letterSpacing: 0.5, flexShrink: 0, whiteSpace: 'pre-line', lineHeight: 1.25 }}>{title}</span>
        <div style={{ flex: 1, minWidth: 0, display: 'flex', justifyContent: 'flex-end' }}>
          {headerExtra ? headerExtra : (subtitle && <span style={{ fontSize: 11, color: 'var(--text-2)' }}>{subtitle}</span>)}
        </div>
      </div>
      {children}
    </div>
  );
}

function NoData({ text }: { text: string }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: 220, color: 'var(--text-2)', fontSize: 13 }}>
      {text}
    </div>
  );
}
