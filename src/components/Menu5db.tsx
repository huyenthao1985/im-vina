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
interface Test5Data { bestModel: string; models: ModelEntry[]; }

// Trạng thái rỗng ban đầu — thay cho test5Data.json dựng sẵn trước đây.
const EMPTY_DATA: Test5Data = { bestModel: '', models: [] };


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
const IDB_KEY_MENU5_DATA = 'menu5:test5_agg_v2';
const IDB_KEY_MENU5_META = 'menu5:test5_agg_v2_meta';

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
}

function createAccumulator(): AggAccumulator {
  return {
    monthly: new Map(),
    radarType: new Map(),
    radarCustom: new Map(),
    rowCount: new Map(),
    ttl: new Map(),
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
  const { monthly, radarType, radarCustom, rowCount, ttl } = acc;
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
      const roundedQty = item === ITEM_FCOST
        ? Math.round(qty * 1000000) / 1000000
        : Math.round(qty * 1000) / 1000;
      byYear.get(yKey)!.set(monthNum, roundedQty);
    }

    // EPCC (menu5-4-new-charts-store-ttl-rows): lưu song song giá trị dòng
    // TTL nguyên văn — xem giải thích root-cause ở khai báo ttlByItemYear.
    if (isTtl && year) {
      if (!ttl.has(model)) ttl.set(model, new Map());
      const byItemTtl = ttl.get(model)!;
      if (!byItemTtl.has(item)) byItemTtl.set(item, new Map());
      byItemTtl.get(item)!.set(year, Math.round(qty * 1000000) / 1000000);
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
  }
}

// Dựng Test5Data cuối cùng từ bộ Map đã tích luỹ (có thể đến từ 1 hoặc
// NHIỀU sheet gộp lại) — logic giữ nguyên 100% so với bản gốc, chỉ đổi
// nguồn đầu vào từ "3 Map cục bộ trong 1 lần gọi" sang "acc dùng chung".
function buildTest5DataFromAccumulator(acc: AggAccumulator): Test5Data {
  const { monthly, radarType, radarCustom, rowCount, ttl } = acc;

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

  return { bestModel, models };
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
    title: 'HIỆU SUẤT SẢN XUẤT THEO MODEL',
    tabOverview: 'TÌNH HÌNH F-COST',
    tabDetail: 'Chi tiết & dữ liệu',
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
    viewBy: 'XEM THEO',
    byMonth: 'Tháng',
    byQuarter: 'Quý',
    byYear: 'Năm',
    // EPCC (menu5-toolbar-import-not-export) - đổi tên key: nút này BẤM VÀO
    // MỞ HỘP THOẠI CHỌN FILE ĐỂ TẢI LÊN (import), không phải xuất/tải xuống
    // báo cáo. Tên cũ "exportExcel" gắn với handleExport() (ghi file mới)
    // gây hiểu sai đúng như người dùng phản ánh — xem chỗ dùng ở toolbar.
    importExcelBtn: 'Tải Excel',
    updatedToLabel: 'DỮ LIỆU CẬP NHẬT ĐẾN',
    kpi1: 'F-COST TB',
    kpi2: 'TỶ LỆ SCRAP',
    kpi3: 'THÁNG ĐẠT MỤC TIÊU',
    kpi4: 'CHÊNH LỆCH SO VỚI NĂM TRƯỚC',
    kpiTarget: 'Mục tiêu',
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
    title: 'MODEL PRODUCTION PERFORMANCE',
    tabOverview: 'Overview',
    tabDetail: 'Detail & Data',
    startMonth: 'START MONTH',
    endMonth: 'END MONTH',
    yearFilterLabel: 'YEAR',
    yearFilterAll: 'All',
    modelFilterAll: 'All',
    model: 'MODEL',
    viewBy: 'VIEW BY',
    byMonth: 'Month',
    byQuarter: 'Quarter',
    byYear: 'Year',
    importExcelBtn: 'Import Excel',
    updatedToLabel: 'DATA UPDATED TO',
    kpi1: 'AVG F-COST',
    kpi2: 'SCRAP RATE',
    kpi3: 'MONTHS ON TARGET',
    kpi4: 'CHANGE VS LAST YEAR',
    kpiTarget: 'Target',
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
    title: '모델별 생산 실적',
    tabOverview: '전체 현황',
    tabDetail: '상세 데이터',
    startMonth: '시작 월',
    endMonth: '종료 월',
    yearFilterLabel: '연도',
    yearFilterAll: '전체',
    modelFilterAll: '전체',
    model: '모델',
    viewBy: '보기 기준',
    byMonth: '월',
    byQuarter: '분기',
    byYear: '년',
    importExcelBtn: '엑셀 불러오기',
    updatedToLabel: '데이터 업데이트 기준일',
    kpi1: '평균 F-Cost',
    kpi2: '스크랩 비율',
    kpi3: '목표 달성 월수',
    kpi4: '전년 대비 변화',
    kpiTarget: '목표',
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
  const yBar = (v: number) => padT + innerH - (v / maxBar) * innerH;
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
  maxV?: number;
}) {
  const axisTextColor = theme === 'light' ? '#000000' : '#e5e7eb';
  const width = 620;
  // EPCC (menu5-chart1-2-restore-axis-gutter) - giống Chart 1: chỉ cần
  // gutter TRÁI cho nhãn % (chart này không có trục phải) — padR giữ gần 0
  // vì không có nhãn nào cần chỗ bên phải, đường dữ liệu vẫn kéo gần sát mép
  // phải như yêu cầu "left/right=0" cho phần không cần nhãn.
  // EPCC (menu5-chart1-2-axis-label-clip-fix) - cùng nguyên nhân/cách fix
  // như Chart 1 (xem comment chi tiết ở ComboBarDualLineChart): tăng padL
  // 30→42 để nhãn % dài hơn (VD '11.1%') không bị cắt mất ký tự đầu.
  const padL = 42, padR = 6, padT = 16, padB = 24;
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
  const maxV = fixedMaxV ?? computeAutoMaxV(series);
  const xOf = (i: number) => padL + step * i;
  const yOf = (v: number) => Math.max(padT, padT + innerH - (v / maxV) * innerH);

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
      {series.map((s, si) => (
        <path key={si} d={buildSmoothLinePath(s.values, xOf, yOf)} fill="none" stroke={s.color} strokeWidth={1.8} strokeLinejoin="round" strokeLinecap="round" />
      ))}
      {series.map((s, si) => s.values.map((v, i) => v == null ? null : (
        <circle key={`${si}-${i}`} cx={xOf(i)} cy={yOf(v)} r={2.5} fill={s.color} />
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
        <text key={`vl${si}-${i}`} x={xOf(i)} y={yOf(v) - 6} fontSize={8.5} fontWeight={700}
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
  const [tab, setTab] = useState<'overview' | 'detail'>('overview');
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
          setDataSource(JSON.parse(cached));
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
  // EPCC (menu5-model-all-option) - FIX phòng ngừa: `bestModel` (tính trong
  // parseWorkbook, không thuộc phạm vi sửa lần này) vẫn duyệt qua TOÀN BỘ
  // `models` kể cả 4 model đặc biệt, nên về lý thuyết có thể trả về 'Target'/
  // 'Actual' nếu F-Cost% của chúng thấp nhất. Từ khi dropdown MODEL ẩn hẳn 4
  // model đặc biệt khỏi danh sách chọn (chỉ còn truy cập qua "Tất cả"), nếu
  // `selectedModel` mặc định lỡ trỏ vào 1 trong số đó sẽ không khớp option
  // nào trong <select> nữa. Tính sẵn "model thật đầu tiên" làm lưới an toàn,
  // dùng cho cả state khởi tạo lẫn effect reset khi đổi nguồn dữ liệu.
  const firstRealModel = allModels.find(m => !SPECIAL_FCOST_MODELS.includes(m.model))?.model ?? '';
  const safeBestModel = SPECIAL_FCOST_MODELS.includes(bestModel) ? '' : bestModel;

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

  const [selectedModel, setSelectedModel] = useState<string>(safeBestModel || firstRealModel || '');
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

  // EPCC (menu5-reset-filter-on-new-source) - khi upload file mới hoặc khôi
  // phục mẫu, danh sách model/khoảng thời gian có thể khác hẳn dữ liệu cũ →
  // phải reset lại filter, tránh selectedModel trỏ tới model không tồn tại
  // trong dữ liệu mới (màn hình trắng do modelData = null).
  useEffect(() => {
    setSelectedModel(safeBestModel || firstRealModel || '');
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

  function seriesInRange(item: string): { y: number; m: number; v: number }[] {
    if (!modelData) return [];
    return (modelData.series[item] ?? [])
      .filter(p => ym(p.y, p.m) >= startYm && ym(p.y, p.m) <= endYm)
      .sort((a, b) => ym(a.y, a.m) - ym(b.y, a.m));
  }

  // Gộp theo Tháng / Quý / Năm tuỳ viewBy — giữ nguyên đơn vị gốc (Q'TY),
  // chỉ đổi cách nhóm trục X.
  function groupPoints(pts: { y: number; m: number; v: number }[]): { label: string; value: number }[] {
    if (viewBy === 'month') {
      return pts.map(p => ({ label: `${t.byMonth === 'Tháng' ? 'T' : 'M'}${p.m}/${String(p.y).slice(2)}`, value: p.v }));
    }
    const bucket = new Map<string, number>();
    for (const p of pts) {
      const key = viewBy === 'quarter' ? `Q${Math.ceil(p.m / 3)}/${p.y}` : `${p.y}`;
      bucket.set(key, (bucket.get(key) ?? 0) + p.v);
    }
    return Array.from(bucket.entries()).map(([label, value]) => ({ label, value }));
  }

  const shipmentPts = seriesInRange(ITEM_SHIPMENT);
  const prodPts = seriesInRange(ITEM_PRODUCTION);
  const salesPts = seriesInRange(ITEM_SALES);
  const scrapPts = seriesInRange(ITEM_SCRAP);
  const fcostPts = seriesInRange(ITEM_FCOST);
  // EPCC (menu5-table-fcost-lookup-by-ym) - FIX ROOT CAUSE "bảng chi tiết có
  // thể hiện SAI cột F-Cost % lệch hàng sau khi fix menu5-fcost-zero-as-
  // missing": bảng cũ tra `fcostPts[i]` THEO VỊ TRÍ INDEX, ngầm giả định
  // shipmentPts/prodPts/salesPts/scrapPts/fcostPts LUÔN cùng độ dài và cùng
  // thứ tự tháng — đúng trước đây vì mọi ITEM đều đủ 12 tháng/năm. Từ khi
  // các tháng F-Cost%=0 giả bị loại khỏi `monthly` (không còn là 1 phần tử
  // trong mảng fcostPts), fcostPts có thể NGẮN HƠN shipmentPts → từ tháng bị
  // loại trở đi, `fcostPts[i]` sẽ trỏ NHẦM sang giá trị của tháng kế tiếp.
  // Fix: tra theo khoá (năm,tháng) thực sự thay vì theo index.
  const fcostByYm = new Map(fcostPts.map(p => [ym(p.y, p.m), p.v]));

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
  // Card 2 "TỶ LỆ SCRAP": tỷ lệ SCRAP/PRODUCTION dùng TTL cả năm (không
  // cộng dồn lại từ các tháng trong khoảng lọc như trước).
  const kpiScrapTtl = ttlOf(modelData, ITEM_SCRAP, lastRangeYear);
  const kpiProdTtl = ttlOf(modelData, ITEM_PRODUCTION, lastRangeYear);
  const kpi2ScrapRate = kpiProdTtl && kpiProdTtl > 0 ? ((kpiScrapTtl ?? 0) / kpiProdTtl) * 100 : 0;
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
  function filterModelsWithDataInRange<T extends { values: (number | null)[] }>(list: T[]): (T & { color: string })[] {
    return list
      .filter(s => s.values.slice(1).some(v => v != null))
      .sort((a, b) => (b.values[0] ?? -Infinity) - (a.values[0] ?? -Infinity))
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

  // EPCC (menu5-chart3-viewby-link) - FIX theo yêu cầu người dùng "cho 4
  // chart đều link giá trị thay đổi khi thay đổi tháng/quý/năm": trước đây
  // Chart 3 LUÔN cộng dồn Scrap qua TẤT CẢ năm trong khoảng lọc
  // (`rangeYears`), không phản hồi nút "XEM THEO" (viewBy). Áp dụng nguyên
  // tắc: viewBy='year' giữ đúng ý nghĩa "ALL YEAR" gốc (cộng dồn mọi năm đã
  // lọc); viewBy='quarter'/'month' thu hẹp phạm vi cộng dồn về ĐÚNG năm cuối
  // (`lastRangeYear`) — khớp đúng năm mà Chart 1/2 cũng đang hiển thị chi
  // tiết theo quý/tháng ở chế độ đó, thay vì cộng dồn nhiều năm không còn
  // khớp ngữ cảnh "đang xem theo quý/tháng".
  const chart3Items = useMemo(() => {
    const chart3Years = viewBy === 'year' ? rangeYears : [lastRangeYear];
    const withVal = realModels.map(m => {
      const hasAny = chart3Years.some(y => ttlOf(m, ITEM_SCRAP, y) != null);
      if (!hasAny) return { label: m.model, value: null as number | null };
      const sum = chart3Years.reduce((a, y) => a + (ttlOf(m, ITEM_SCRAP, y) ?? 0), 0);
      return { label: m.model, value: Math.round(sum * 10) / 10 };
    });
    const withData = withVal
      .filter(x => x.value != null)
      .sort((a, b) => (b.value as number) - (a.value as number))
      .slice(0, CHART3_MAX_MODELS);
    return symmetricOrder(withData);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [realModels, rangeYears, lastRangeYear, viewBy]);

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

      {/* ── 2 Tab (copy nguyên văn RtyDashboard.tsx) ── */}
      <div className="tab-container">
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

            {/* Cụm phải: nút "Tải Excel" — NÚT NHẬP (import) file .xlsx, dùng
                lại đúng fileInputRef/handleUploadFile đã có sẵn cho màn hình
                rỗng ban đầu. Icon mũi tên xoay tròn (refresh/reload), KHÔNG
                phải icon mũi tên xuống (download), để tránh gây hiểu nhầm
                đây là nút xuất dữ liệu. */}
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
                style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px', height: '38px', width: '120px', boxSizing: 'border-box', fontSize: '13px' }}
              >
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15" style={{ flexShrink: 0 }}>
                  <path d="M21 12a9 9 0 0 1-9 9c-2.52 0-4.93-1-6.74-2.74L3 16" />
                  <path d="M3 12a9 9 0 0 1 9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
                  <path d="M3 16v5h5" />
                  <path d="M16 3h5v5" />
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
                đây (viền đơn sắc, không icon, không đồng bộ mục 2). ── */}
            <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '12px', marginBottom: '12px', width: '100%' }}>
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
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>{kpi2ScrapRate.toFixed(2)}%</div>
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

              <div className="kpi-card" style={{ borderLeft: '4px solid #f59e0b', background: 'linear-gradient(135deg, rgba(245,158,11,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
                <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                    <polyline points="23 6 13.5 15.5 8.5 10.5 1 18" />
                    <polyline points="17 6 23 6 23 12" />
                  </svg>
                  <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi4}</div>
                </div>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>
                  {kpi4DeltaPp != null ? `${kpi4DeltaPp >= 0 ? '+' : ''}${kpi4DeltaPp.toFixed(2)}` : '—'}
                  {kpi4DeltaPp != null && <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85, marginLeft: '1px' }}>pp</span>}
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
            </div>
          </>
        ) : (
          <div style={{ background: 'var(--surface)', border: '1px solid var(--border-soft)', borderRadius: 12, overflow: 'auto' }}>
            {/* EPCC (menu5-table-fcost-column-color-by-viewby) - FIX theo yêu
                cầu người dùng "màu cột F-Cost % trong bảng cũng thay đổi theo
                chế độ xem": trước đây cột F-Cost % LUÔN cố định 1 màu #4EB09B
                bất kể `viewBy` (Tháng/Quý/Năm). Giờ tính màu theo `viewBy`,
                tái dùng đúng bảng NEON có sẵn (không bịa màu mới) — Tháng giữ
                nguyên #4EB09B (ảnh 1 người dùng cung cấp), Quý = NEON.amber
                (cam), Năm = NEON.violet (tím) — dùng chung 1 biến cho cả
                header lẫn từng ô dữ liệu để luôn đồng bộ. */}
            {(() => {
              const fcostColumnColor = viewBy === 'quarter' ? NEON.amber : viewBy === 'year' ? NEON.violet : '#4EB09B';
              return (
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ background: 'var(--border-soft)' }}>
                  {[t.colYear, t.colMonth, t.colShipment, t.colProd, t.colSales, t.colScrap].map(h => (
                    <th key={h} style={{ padding: '8px 12px', textAlign: 'left', color: 'var(--text-0)' }}>{h}</th>
                  ))}
                  <th style={{ padding: '8px 12px', textAlign: 'left', color: fcostColumnColor, fontWeight: 700 }}>{t.colFcost}</th>
                </tr>
              </thead>
              <tbody>
                {shipmentPts.map((p, i) => (
                  <tr key={`${p.y}-${p.m}`} style={{ borderTop: '1px solid var(--border-soft)' }}>
                    <td style={{ padding: '6px 12px', color: 'var(--text-2)' }}>{p.y}</td>
                    <td style={{ padding: '6px 12px', color: 'var(--text-2)' }}>{p.m}</td>
                    <td style={{ padding: '6px 12px', color: 'var(--text-0)' }}>{fmtNum(p.v)}</td>
                    <td style={{ padding: '6px 12px', color: 'var(--text-0)' }}>{fmtNum(prodPts[i]?.v ?? 0)}</td>
                    <td style={{ padding: '6px 12px', color: 'var(--text-0)' }}>{fmtNum(salesPts[i]?.v ?? 0)}</td>
                    <td style={{ padding: '6px 12px', color: 'var(--text-0)' }}>{fmtNum(scrapPts[i]?.v ?? 0)}</td>
                    <td style={{ padding: '6px 12px', color: fcostColumnColor, fontWeight: 700 }}>{fcostByYm.has(ym(p.y, p.m)) ? `${(fcostByYm.get(ym(p.y, p.m))! * 100).toFixed(2)}%` : '—'}</td>
                  </tr>
                ))}
                {!shipmentPts.length && (
                  <tr><td colSpan={7} style={{ padding: 20, textAlign: 'center', color: 'var(--text-2)' }}>{t.noData}</td></tr>
                )}
              </tbody>
            </table>
              );
            })()}
          </div>
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
  title, subtitle, headerExtra, accent, idx, theme, children,
}: {
  title: string; subtitle?: string; headerExtra?: React.ReactNode; accent: string; idx: number; theme: ThemeMode; children: React.ReactNode;
}) {
  const hc = CHART_HEADER_THEME[idx % CHART_HEADER_THEME.length];
  const headerStyle: React.CSSProperties = {
    background: theme === 'light' ? hc.bgLight : hc.bgDark,
    borderLeft: `4px solid ${hc.accent}`,
    borderRadius: '8px 8px 0 0',
    padding: '10px 14px',
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
    margin: '-14px -14px 10px -14px',
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
