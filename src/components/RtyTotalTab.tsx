import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import * as XLSX from 'xlsx';
import { NeonButton } from './NeonButton';

// ═════════════════════════════════════════════════════════════════════════
// MENU 5 — TAB MỚI "RTY TOTAL"
// ═════════════════════════════════════════════════════════════════════════
// Nguồn dữ liệu: Test6.xlsx (cột: MODEL, TYPE, ITEM, YEAR, MONTH, RTY %).
//
// QUY TẮC BẮT BUỘC (đã thống nhất với người dùng — KHÔNG được vi phạm):
//  - ITEM luôn = 'RTY' trong toàn bộ file → giá trị RTY được LINK THẲNG từ
//    cột "RTY %", TUYỆT ĐỐI KHÔNG tự suy ra/tính toán/phỏng đoán bằng công
//    thức nào khác (không dùng PROD'N/SCRAP/RTV như bản nháp Test5 trước —
//    bản đó đã bị người dùng từ chối vì cho ra số âm bất thường).
//  - TYPE = 'SUB1'      → RTY Sub1  (link thẳng)
//  - TYPE = 'SUB2'      → RTY Sub2  (link thẳng)
//  - TYPE ∈ {OIS IR, OIS ASSY, OIS LENS, VCMC} → RTY Main (link thẳng; mỗi
//    model trong file gốc chỉ dùng ĐÚNG 1 trong 4 loại này, không trộn).
//  - TYPE = 'RTY %'     → RTY TTL   (đã được hệ thống nguồn tính sẵn, LINK
//    THẲNG, không tự nhân SUB1×SUB2×MAIN lại).
//  - 4 dòng "model giả" SUB1 TARGET / SUB2 TARGET / MAIN TARGET / RTY TARGET
//    chính là mục tiêu (target) — LINK THẲNG từ file, không hardcode số.
//  - Giá trị RTY = 0 trong file gốc là Ô TRỐNG CHỜ DỮ LIỆU tương lai (chưa
//    phát sinh sản xuất tháng đó), KHÔNG PHẢI 0% hiệu suất thật → coi là
//    "không có dữ liệu" (null), không vẽ lên chart để tránh hiểu lầm.
//
// Cách vẽ (radar + 4 panel cột+đường), màu sắc, kích thước: TÁI SỬ DỤNG
// NGUYÊN VĂN từ RtyDashboard.tsx (Mục 4) theo đúng yêu cầu người dùng —
// xem plotMixedPanel()/plotSpiderPanel() bên dưới, giữ y hệt bảng màu,
// target line đỏ đứt nét, panel header colors, RtyLegendItem.
// ═════════════════════════════════════════════════════════════════════════

type Lang = 'vi' | 'en' | 'ko';
type ThemeMode = 'dark' | 'light';

// ── Kiểu dữ liệu sau khi gộp file ───────────────────────────────────────
interface StageMap { [ym: string]: number } // key "YYYY-MM" -> RTY (0..1)
interface ModelRtySeries { sub1: StageMap; sub2: StageMap; main: StageMap; ttl: StageMap }
interface RtyTargets { sub1: number | null; sub2: number | null; main: number | null; ttl: number | null }
interface RtyTotalData {
  months: string[];                         // toàn bộ mốc YYYY-MM có dữ liệu, đã sort tăng dần
  models: Record<string, ModelRtySeries>;    // model thật (không gồm 4 dòng TARGET)
  targets: RtyTargets;
}
const EMPTY_DATA: RtyTotalData = { months: [], models: {}, targets: { sub1: null, sub2: null, main: null, ttl: null } };

// ── IndexedDB cache riêng cho tab RTY Total (không đụng key của tab khác) ─
const IDB_DB_NAME = 'imvina_dashboard_cache';
const IDB_STORE = 'kv';
const IDB_KEY_RTY_DATA = 'menu5:rty_total_v1';
const IDB_KEY_RTY_META = 'menu5:rty_total_v1_meta';

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

// ── Parse workbook (gộp MỌI sheet có đủ cột bắt buộc, giống quy ước Menu5) ─
const MONTH_MAP: Record<string, number> = {
  JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12,
};
const MAIN_TYPES = new Set(['OIS IR', 'OIS ASSY', 'OIS LENS', 'VCMC']);
const TARGET_ROW_KEYS: Record<string, keyof RtyTargets> = {
  'SUB1 TARGET': 'sub1', 'SUB2 TARGET': 'sub2', 'MAIN TARGET': 'main', 'RTY TARGET': 'ttl',
};

function normalizeHeader(h: unknown): string {
  return String(h ?? '').toUpperCase().replace(/[^A-Z]/g, '');
}

class RtyParseError extends Error {}

function findHeaderRow(rows: unknown[][]): { headerRowIdx: number; colIdx: Record<string, number> } | null {
  for (let r = 0; r < Math.min(rows.length, 20); r++) {
    const row = rows[r] ?? [];
    const map: Record<string, number> = {};
    row.forEach((cell, i) => {
      const key = normalizeHeader(cell);
      if (key && !(key in map)) map[key] = i;
    });
    const hasCore = ['MODEL', 'TYPE', 'ITEM', 'YEAR', 'MONTH'].every(k => k in map);
    // Cột giá trị RTY có thể ghi là "RTY %", "RTY%", "RTY" — dò theo header chứa "RTY"
    const rtyKey = Object.keys(map).find(k => k.includes('RTY'));
    if (hasCore && rtyKey) return { headerRowIdx: r, colIdx: { ...map, RTYVAL: map[rtyKey] } };
  }
  return null;
}

function accumulateSheetRows(
  rows: unknown[][], colIdx: Record<string, number>, headerRowIdx: number,
  models: Record<string, ModelRtySeries>, targets: RtyTargets, monthsSet: Set<string>,
): void {
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r] ?? [];
    const modelRaw = row[colIdx.MODEL];
    const model = modelRaw != null ? String(modelRaw).trim() : '';
    if (!model) continue;
    const item = row[colIdx.ITEM] != null ? String(row[colIdx.ITEM]).trim().toUpperCase() : '';
    if (item !== 'RTY') continue; // EPCC: chỉ nhận đúng dòng ITEM = 'RTY', không suy diễn từ dòng khác
    const typ = row[colIdx.TYPE] != null ? String(row[colIdx.TYPE]).trim() : '';
    const yearRaw = row[colIdx.YEAR];
    const year = yearRaw != null ? Number(yearRaw) : NaN;
    const monthRaw = row[colIdx.MONTH];
    const ms = String(monthRaw ?? '').toUpperCase().trim();
    if (ms === 'TTL' || !(ms in MONTH_MAP) || !year) continue; // bỏ dòng TTL năm, chỉ lấy đúng tháng thật
    const valRaw = row[colIdx.RTYVAL];
    const val = typeof valRaw === 'number' ? valRaw : parseFloat(String(valRaw ?? '').replace(/,/g, ''));
    if (!isFinite(val) || val === 0) continue; // 0 = ô trống chờ dữ liệu, KHÔNG phải 0% thật — bỏ qua

    const ym = `${year}-${String(MONTH_MAP[ms]).padStart(2, '0')}`;

    const targetKey = TARGET_ROW_KEYS[model.toUpperCase()];
    if (targetKey) {
      // Dòng target: giá trị giống nhau mọi tháng trong file gốc — chỉ cần gán 1 lần.
      if (targets[targetKey] == null) targets[targetKey] = val;
      continue;
    }

    if (!models[model]) models[model] = { sub1: {}, sub2: {}, main: {}, ttl: {} };
    monthsSet.add(ym);
    if (typ === 'SUB1') models[model].sub1[ym] = val;
    else if (typ === 'SUB2') models[model].sub2[ym] = val;
    else if (typ === 'RTY %') models[model].ttl[ym] = val;
    else if (MAIN_TYPES.has(typ)) models[model].main[ym] = val;
  }
}

function aggregateWorkbookToRtyData(wb: XLSX.WorkBook): RtyTotalData {
  const models: Record<string, ModelRtySeries> = {};
  const targets: RtyTargets = { sub1: null, sub2: null, main: null, ttl: null };
  const monthsSet = new Set<string>();
  let matchedAnySheet = false;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][];
    const found = findHeaderRow(rows);
    if (!found) continue;
    matchedAnySheet = true;
    accumulateSheetRows(rows, found.colIdx, found.headerRowIdx, models, targets, monthsSet);
  }

  if (!matchedAnySheet) {
    throw new RtyParseError('Không tìm thấy đủ cột bắt buộc (MODEL, TYPE, ITEM, YEAR, MONTH, RTY %) trong bất kỳ sheet nào của file.');
  }
  if (Object.keys(models).length === 0) {
    throw new RtyParseError('Không đọc được dòng RTY hợp lệ nào (ITEM = "RTY", giá trị khác 0).');
  }

  return { months: Array.from(monthsSet).sort(), models, targets };
}

// ── Màu sắc / kiểu vẽ: TÁI SỬ DỤNG NGUYÊN VĂN từ RtyDashboard.tsx ─────────
const RADAR_COLORS = ['#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#ec4899', '#14b8a6', '#f43f5e', '#6366f1'];

interface RtyLegend { type: 'bar' | 'line' | 'dashed' | 'dotted'; label: string; color: string }
const RtyLegendItem: React.FC<RtyLegend> = ({ type, label, color }) => (
  <div style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', fontSize: '10.5px', color: 'inherit', margin: '0 4px' }}>
    {type === 'bar' && <span style={{ width: 10, height: 10, background: color, display: 'inline-block', borderRadius: 1.5, flexShrink: 0 }} />}
    {type === 'line' && <span style={{ width: 15, height: 2.5, background: color, display: 'inline-block', flexShrink: 0 }} />}
    {type === 'dashed' && <span style={{ width: 15, height: 0, borderTop: `2.5px dashed ${color}`, display: 'inline-block', flexShrink: 0 }} />}
    {type === 'dotted' && <span style={{ width: 15, height: 0, borderTop: `2.5px dotted ${color}`, display: 'inline-block', flexShrink: 0 }} />}
    <span style={{ whiteSpace: 'nowrap', opacity: 0.9 }}>{label}</span>
  </div>
);

declare global { interface Window { Plotly?: any } }

const TXT: Record<Lang, Record<string, string>> = {
  vi: {
    tabRtyTotal: 'RTY TOTAL', model: 'MODEL', importExcelBtn: 'Tải Excel RTY',
    parsing: 'Đang xử lý...', noDataYet: '📁 Chưa có dữ liệu, vui lòng tải Test6.xlsx lên',
    emptyDesc: 'Tải lên file Test6.xlsx (cột MODEL/TYPE/ITEM/YEAR/MONTH/RTY %) để xem báo cáo.',
    radarTTL: 'TTL RTY THEO MODEL', radarMAIN: 'MAIN RTY THEO MODEL',
    radarSUB1: 'SUB1 RTY THEO MODEL', radarSUB2: 'SUB2 RTY THEO MODEL',
    ttlByModelLine: 'TTL RTY THEO MODEL (8 THÁNG GẦN NHẤT)',
    panelTTL: 'TTL RTY', panelMAIN: 'MAIN', panelSUB1: 'SUB1', panelSUB2: 'SUB2',
  },
  en: {
    tabRtyTotal: 'RTY TOTAL', model: 'MODEL', importExcelBtn: 'Upload RTY Excel',
    parsing: 'Processing...', noDataYet: '📁 No data yet, please upload Test6.xlsx',
    emptyDesc: 'Upload Test6.xlsx (columns MODEL/TYPE/ITEM/YEAR/MONTH/RTY %) to view the report.',
    radarTTL: 'TTL RTY BY MODEL', radarMAIN: 'MAIN RTY BY MODEL',
    radarSUB1: 'SUB1 RTY BY MODEL', radarSUB2: 'SUB2 RTY BY MODEL',
    ttlByModelLine: 'TTL RTY BY MODEL (LAST 8 MONTHS)',
    panelTTL: 'TTL RTY', panelMAIN: 'MAIN', panelSUB1: 'SUB1', panelSUB2: 'SUB2',
  },
  ko: {
    tabRtyTotal: 'RTY TOTAL', model: '모델', importExcelBtn: 'RTY 엑셀 업로드',
    parsing: '처리 중...', noDataYet: '📁 데이터 없음, Test6.xlsx를 업로드하세요',
    emptyDesc: 'Test6.xlsx (열: MODEL/TYPE/ITEM/YEAR/MONTH/RTY %) 업로드.',
    radarTTL: '모델별 TTL RTY', radarMAIN: '모델별 MAIN RTY',
    radarSUB1: '모델별 SUB1 RTY', radarSUB2: '모델별 SUB2 RTY',
    ttlByModelLine: '모델별 TTL RTY (최근 8개월)',
    panelTTL: 'TTL RTY', panelMAIN: 'MAIN', panelSUB1: 'SUB1', panelSUB2: 'SUB2',
  },
};

export function RtyTotalTab({ theme, lang }: { theme: ThemeMode; lang: Lang }) {
  const t = TXT[lang];
  const isLightMode = theme === 'light';
  const tealAccent = isLightMode ? '#0f766e' : '#14b8a6';

  const [data, setData] = useState<RtyTotalData>(EMPTY_DATA);
  const [isParsing, setIsParsing] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [selectedModel, setSelectedModel] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // EPCC (rty-total-toolbar-like-mucd4): thêm "Xem theo" — CHỈ Tháng/Quý/Năm
  // (KHÔNG có Ngày/Tuần như ảnh tham chiếu Mục 4/Test4.xlsx) vì Test6.xlsx
  // chỉ có 2 cột YEAR+MONTH, không có ngày/tuần thật — nếu thêm "Ngày/Tuần"
  // sẽ phải suy diễn/generate ngày giả, đúng điều người dùng đã yêu cầu
  // KHÔNG được làm. Quý/Năm hiển thị = TRUNG BÌNH CỘNG đơn giản các tháng
  // có dữ liệu thật rơi vào quý/năm đó (không có cột sản lượng trong Test6
  // để tính trung bình gia quyền theo SL — đã ghi rõ trong comment tính toán).
  const [viewBy, setViewBy] = useState<'month' | 'quarter' | 'year'>('month');

  const [spiderLegendTTL, setSpiderLegendTTL] = useState<{ label: string; color: string }[]>([]);
  const [spiderLegendMAIN, setSpiderLegendMAIN] = useState<{ label: string; color: string }[]>([]);
  // EPCC (rty-total-chart56-to-spider): Chart 5,6 (SUB1/SUB2) chuyển từ dạng
  // cột+đường sang biểu đồ nhện (radar) giống hệt Chart 1,2 — cần thêm 2 state
  // legend riêng vì plotSpiderPanel() set legend theo từng stage độc lập.
  const [spiderLegendSUB1, setSpiderLegendSUB1] = useState<{ label: string; color: string }[]>([]);
  const [spiderLegendSUB2, setSpiderLegendSUB2] = useState<{ label: string; color: string }[]>([]);

  // ── Nạp cache khi mở tab lần đầu ──
  useEffect(() => {
    (async () => {
      const cached = await idbGetCache(IDB_KEY_RTY_DATA);
      if (cached) {
        try { setData(JSON.parse(cached) as RtyTotalData); } catch (e) { /* cache hỏng, bỏ qua */ }
      }
    })();
  }, []);

  const modelList = useMemo(() => Object.keys(data.models).sort(), [data.models]);

  // ── Model "hiệu suất tốt nhất" (⭐) — giống HỆT định nghĩa BEST_MODEL của
  //    RtyDashboard.tsx (Mục 4): actual RTY TTL trung bình CHÊNH LỆCH DƯƠNG/
  //    CAO NHẤT so với target.ttl, chỉ xét model có ≥3 điểm TTL thật (tránh
  //    1-2 điểm lẻ tẻ gây sai lệch). KHÔNG còn dùng tiêu chí "nhiều dữ liệu
  //    nhất" như bản trước — đổi đúng theo ảnh tham chiếu "Hiệu suất tốt nhất". ──
  const bestModel = useMemo(() => {
    let best = '', bestGap = -Infinity;
    for (const m of modelList) {
      const vals = Object.values(data.models[m].ttl);
      if (vals.length < 3) continue;
      const avg = vals.reduce((a, v) => a + v, 0) / vals.length;
      const gap = data.targets.ttl != null ? avg - data.targets.ttl : avg;
      if (gap > bestGap) { bestGap = gap; best = m; }
    }
    return best || modelList[0] || '';
  }, [modelList, data.models, data.targets.ttl]);

  // ── Biên độ tháng toàn dữ liệu (dùng làm min/max + giá trị mặc định cho
  //    2 ô THÁNG BẮT ĐẦU/KẾT THÚC) — cùng cách mã hoá y*12+(m-1) như
  //    Menu5ModelDashboard.tsx để tái sử dụng logic input type="month". ──
  const ymNum = (s: string) => { const [y, m] = s.split('-').map(Number); return y * 12 + (m - 1); };
  const numToYm = (n: number) => { const y = Math.floor(n / 12), m = (n % 12) + 1; return `${y}-${String(m).padStart(2, '0')}`; };

  const dataBounds = useMemo(() => {
    if (data.months.length === 0) return null;
    return { min: ymNum(data.months[0]), max: ymNum(data.months[data.months.length - 1]) };
  }, [data.months]);

  const [startYm, setStartYm] = useState<number | null>(null);
  const [endYm, setEndYm] = useState<number | null>(null);

  // Reset model + khoảng lọc mỗi khi có nguồn dữ liệu MỚI (upload file khác)
  // — tránh selectedModel/khoảng tháng trỏ vào dữ liệu không còn tồn tại.
  useEffect(() => {
    if (modelList.length === 0) { setSelectedModel(''); setStartYm(null); setEndYm(null); return; }
    setSelectedModel(prev => (prev && modelList.includes(prev)) ? prev : (bestModel || modelList[0]));
    if (dataBounds) { setStartYm(dataBounds.min); setEndYm(dataBounds.max); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data]);

  const handleUploadFile = async (file: File) => {
    setIsParsing(true); setUploadError(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const parsed = aggregateWorkbookToRtyData(wb);
      setData(parsed);
      await idbSetCache(IDB_KEY_RTY_DATA, JSON.stringify(parsed));
      await idbSetCache(IDB_KEY_RTY_META, JSON.stringify({ fileName: file.name, uploadedAt: new Date().toISOString() }));
    } catch (e) {
      setUploadError(e instanceof RtyParseError ? e.message : 'Không đọc được file Excel. Vui lòng kiểm tra lại định dạng.');
    } finally {
      setIsParsing(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // ── Gộp các tháng trong khoảng [startYm, endYm] thành "bucket" theo
  //    Tháng/Quý/Năm — mỗi bucket giữ lại DANH SÁCH tháng thật gốc để tính
  //    trung bình cộng đơn giản khi vẽ (KHÔNG suy diễn thêm tháng nào). ──
  const buckets = useMemo(() => {
    const filtered = data.months.filter(ymStr => {
      const n = ymNum(ymStr);
      return startYm == null || endYm == null || (n >= startYm && n <= endYm);
    });
    if (viewBy === 'month') return filtered.map(ymStr => ({ label: ymStr, months: [ymStr] }));
    const map = new Map<string, string[]>();
    for (const ymStr of filtered) {
      const [y, m] = ymStr.split('-').map(Number);
      const key = viewBy === 'quarter' ? `${y}-Q${Math.ceil(m / 3)}` : `${y}`;
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(ymStr);
    }
    return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([label, months]) => ({ label, months }));
  }, [data.months, startYm, endYm, viewBy]);

  // Chỉ lấy tối đa 10 mốc gần nhất trong khoảng đã lọc — giảm từ 12 xuống 10
  // theo yêu cầu người dùng để nhãn tháng/năm dưới trục X hiển thị đủ, không
  // bị chồng chéo/cắt chữ khi chart co hẹp.
  const xs = useMemo(() => buckets.slice(-10).map(b => b.label), [buckets]);
  const bucketByLabel = useMemo(() => new Map(buckets.map(b => [b.label, b.months])), [buckets]);

  const getY = (stage: keyof ModelRtySeries, label: string): number | null => {
    const m = data.models[selectedModel];
    const months = bucketByLabel.get(label);
    if (!m || !months) return null;
    const vals = months.map(ymStr => m[stage][ymStr]).filter((v): v is number => v != null);
    if (vals.length === 0) return null;
    return vals.reduce((a, v) => a + v, 0) / vals.length; // trung bình cộng đơn giản — Test6 không có cột sản lượng để gia quyền
  };

  const [plotlyReady, setPlotlyReady] = useState<boolean>(typeof window !== 'undefined' && !!window.Plotly);
  useEffect(() => {
    if (plotlyReady) return;
    const id = setInterval(() => {
      if (typeof window !== 'undefined' && window.Plotly) { setPlotlyReady(true); clearInterval(id); }
    }, 200);
    return () => clearInterval(id);
  }, [plotlyReady]);

  useEffect(() => {
    if (!plotlyReady || typeof window === 'undefined' || !window.Plotly || xs.length === 0) return;

    const fontColor = isLightMode ? '#334155' : '#e2e8f0';
    const gridColor = isLightMode ? 'rgba(0,0,0,0.08)' : 'rgba(255,255,255,0.08)';

    // ═══ plotMixedPanel — COPY NGUYÊN VĂN cách vẽ/màu từ RtyDashboard.tsx ═══
    const plotMixedPanel = (
      elId: string,
      barSeries: { key: keyof ModelRtySeries; name: string; color: string }[],
      lineSeries: { key: keyof ModelRtySeries; name: string; color: string; width?: number; dash?: string }[],
      targetFrac: number | null, targetLabel: string,
    ) => {
      const el = document.getElementById(elId);
      if (!el) return;
      if (targetFrac == null) { window.Plotly.purge(elId); return; }

      const getYs = (key: keyof ModelRtySeries): (number | null)[] =>
        xs.map(ym => { const v = getY(key, ym); return v == null ? null : parseFloat((v * 100).toFixed(2)); });

      const barTraces = barSeries.map(s => {
        const ys = getYs(s.key);
        return {
          x: xs, y: ys, name: s.name, type: 'bar' as const,
          marker: { color: s.color },
          text: ys.map(v => v == null ? '' : `${v.toFixed(1)}%`),
          textposition: 'inside' as const,
          textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
          insidetextanchor: 'middle' as const,
          hovertemplate: `%{x}<br><b>${s.name}: %{y:.2f}%</b><extra></extra>`,
        };
      });

      const lineTraces = lineSeries.map((s, idx) => {
        const ys = getYs(s.key);
        const pos = idx === 0 ? 'top center' : 'bottom center';
        return {
          x: xs, y: ys, name: s.name, type: 'scatter' as const, mode: 'lines+markers+text' as const,
          yaxis: 'y2' as const,
          line: { color: s.color, width: s.width ?? 1.6, shape: 'spline' as const, smoothing: 1, dash: s.dash as any },
          marker: { color: s.color, size: 6 },
          text: ys.map(v => v == null ? '' : `${v.toFixed(1)}%`),
          textposition: pos as 'top center' | 'bottom center',
          textfont: { size: 10, color: s.color, family: 'Arial Black, Arial, sans-serif' },
          cliponaxis: false, connectgaps: false, // KHÔNG nối qua chỗ null — dữ liệu thật thưa, không được "làm mượt" giả
          hovertemplate: `%{x}<br><b>${s.name}: %{y:.2f}%</b><extra></extra>`,
        };
      });

      const barYs = barTraces.flatMap(tr => (tr.y as (number | null)[]).filter((v): v is number => v != null));
      const lineYs = lineTraces.flatMap(tr => (tr.y as (number | null)[]).filter((v): v is number => v != null));
      const targetY = targetFrac * 100;
      lineYs.push(targetY);

      const BAR_TOP_FRAC = 0.55, LINE_BOTTOM_FRAC = 0.63, LINE_TOP_FRAC = 0.95;
      const maxBarVal = Math.max(...barYs, 0);
      const yRange = [0, maxBarVal > 0 ? maxBarVal / BAR_TOP_FRAC : 100 / BAR_TOP_FRAC];
      const lineMin = Math.min(...lineYs), lineMax = Math.max(...lineYs);
      const spanFrac = LINE_TOP_FRAC - LINE_BOTTOM_FRAC;
      const rawSpan = lineMax - lineMin;
      const D = rawSpan > 0.01 ? rawSpan / spanFrac : Math.max(lineMax * 0.05, 1);
      const y2Min = lineMin - LINE_BOTTOM_FRAC * D, y2Max = y2Min + D;

      window.Plotly.newPlot(elId, [
        ...barTraces, ...lineTraces,
        {
          x: xs, y: xs.map(() => targetY), name: `${targetLabel} ${targetY.toFixed(1)}%`,
          type: 'scatter', mode: 'lines+markers+text', yaxis: 'y2',
          line: { color: '#ef4444', width: 1.6, dash: 'dot', shape: 'spline', smoothing: 1 },
          marker: { color: '#ef4444', size: 5 },
          text: xs.map(() => `${targetY.toFixed(1)}%`), textposition: 'top center',
          textfont: { size: 9, color: '#ef4444', family: 'Arial Black, Arial, sans-serif' },
          cliponaxis: false, hovertemplate: `Target: ${targetY.toFixed(1)}%<extra></extra>`,
        },
      ], {
        barmode: 'group', paper_bgcolor: 'transparent', plot_bgcolor: 'transparent',
        font: { color: fontColor, size: 11 }, margin: { t: 15, r: 40, b: 28, l: 48 }, showlegend: false,
        xaxis: { tickfont: { size: 11, color: fontColor }, gridcolor: gridColor, tickangle: 0 },
        yaxis: { gridcolor: gridColor, tickfont: { size: 10 }, range: yRange, ticksuffix: '%' },
        yaxis2: { overlaying: 'y', side: 'right', showgrid: false, tickfont: { size: 10, color: fontColor }, range: [y2Min, y2Max], ticksuffix: '%' },
        hoverlabel: { font: { size: 10 } },
      }, { displayModeBar: false, responsive: true });
    };

    // EPCC (rty-total-unify-radar-line-data) - FIX ROOT CAUSE "Chart 1 (radar
    // TTL theo model) và Chart 6 (line TTL theo model) hiển thị số liệu khác
    // nhau, Chart 6 tụt xuống bất thường (32%/53%/62%...)": nguyên nhân là
    // Chart 6 trước đây tự lấy "tháng gần nhất của RIÊNG từng model" (độc lập
    // với Chart 1), nên nếu 1 model có dữ liệu thưa/model cũ ít cập nhật,
    // tháng gần nhất của model đó có thể rơi vào 1 tháng cũ hiệu suất thấp mà
    // Chart 1 không hề hiển thị (vì Chart 1 chỉ xét chung 8 tháng gần nhất
    // toàn bộ dữ liệu). FIX: tách 1 hàm DUY NHẤT getModelSeriesForStage() để
    // cả radar (scatterpolar) và line-theo-model (scatter) CÙNG dùng chung 1
    // khung "8 tháng gần nhất + danh sách model có dữ liệu trong khung đó" —
    // đảm bảo 2 biểu đồ luôn cùng 1 nguồn số liệu, chỉ khác cách trình bày.
    const getModelSeriesForStage = (stage: 'ttl' | 'main' | 'sub1' | 'sub2') => {
      const allModels = modelList;
      const candidateMonths = data.months.filter(ym => allModels.some(m => data.models[m][stage][ym] != null));
      const selected = candidateMonths.slice(-8);
      const modelsWithData = allModels.filter(m => selected.some(ym => data.models[m][stage][ym] != null));
      // EPCC (rty-total-swap-so1b01-so1c2ed): theo yêu cầu người dùng, đổi chỗ
      // thứ tự hiển thị giữa 'SO1B01' và 'SO1C2ED' trên trục model (chỉ đổi VỊ
      // TRÍ hiển thị, KHÔNG đổi số liệu của từng model — số liệu mỗi model vẫn
      // link thẳng đúng model đó). Áp dụng ở đây (dùng chung cho cả radar lẫn
      // line-theo-model) để 2 chart luôn khớp thứ tự nhau.
      const i1 = modelsWithData.indexOf('SO1B01');
      const i2 = modelsWithData.indexOf('SO1C2ED');
      if (i1 !== -1 && i2 !== -1) {
        [modelsWithData[i1], modelsWithData[i2]] = [modelsWithData[i2], modelsWithData[i1]];
      }
      return { selected, modelsWithData };
    };

    // Panel 1 — TTL: cột Sub1/Sub2, đường Main/TTL — giống hệt RtyDashboard
    plotMixedPanel('rtyTotalTTL',
      [{ key: 'sub1', name: 'RTY Sub1', color: '#1565C0' }, { key: 'sub2', name: 'RTY Sub2', color: '#ef4444' }],
      [{ key: 'main', name: 'RTY Main', color: tealAccent, width: 1.6 }, { key: 'ttl', name: 'RTY TTL', color: '#8b5cf6', width: 1.6, dash: 'dash' }],
      data.targets.ttl, 'Target RTY TTL');

    // Chart 6 — "TTL RTY THEO MODEL" dạng đường CONG MƯỢT (spline/smoothed):
    // dùng ĐÚNG getModelSeriesForStage('ttl') — CÙNG 8 tháng gần nhất + cùng
    // danh sách model + CÙNG bảng màu RADAR_COLORS theo tháng như Chart 1
    // (radar TTL) — chỉ khác trục toạ độ (Đề-các thay vì cực) và cách vẽ
    // đường nối (spline mượt thay vì toả tia). Legend header của panel này
    // dùng lại trực tiếp state spiderLegendTTL (do plotSpiderPanel('rtyTotalRadarTTL')
    // set) nên 2 legend luôn khớp nhau tuyệt đối.
    const plotModelLinePanel = (elId: string, stage: 'ttl' | 'main' | 'sub1' | 'sub2', targetFrac: number | null) => {
      const el = document.getElementById(elId);
      if (!el) return;
      const { selected, modelsWithData } = getModelSeriesForStage(stage);
      if (selected.length === 0 || modelsWithData.length === 0) { window.Plotly.purge(elId); return; }

      const traces: any[] = selected.map((ym, i) => {
        const y = modelsWithData.map(m => {
          const v = data.models[m][stage][ym];
          return v != null ? parseFloat((v * 100).toFixed(2)) : null;
        });
        const color = RADAR_COLORS[i % RADAR_COLORS.length];
        return {
          x: modelsWithData, y, name: ym, type: 'scatter' as const, mode: 'lines+markers+text' as const,
          line: { color, width: 2, shape: 'spline' as const, smoothing: 1 }, // spline = đường CONG mượt (smoothed)
          marker: { color, size: 6 },
          text: y.map(v => v == null ? '' : v.toFixed(1)),
          // EPCC (rty-chart6-declutter-low-values) - FIX theo yêu cầu người
          // dùng "những mục nào dưới 90% hãy di chuyển sang trục Y để trục X
          // nhìn rõ hơn": các điểm < 90% thường dồn cụm ở nhóm model đầu bên
          // trái (VD SO1C2ED/SO1B01 trong ảnh tham chiếu) — nhãn "top center"
          // mặc định chồng lên nhau và đè xuống gần nhãn model trên trục X.
          // Với riêng các điểm này, đẩy nhãn sang bên trái điểm (hướng về
          // phía trục Y) bằng 'middle left' để tách nhãn ra khỏi cụm, không
          // còn che nhãn trục X. Các điểm >= 90% giữ nguyên 'top center' như
          // cũ. `textposition` hỗ trợ mảng theo từng điểm dữ liệu (arrayOK).
          textposition: y.map(v => (v != null && v < 90 ? 'middle left' as const : 'top center' as const)),
          textfont: { color, size: 9.5, family: 'Arial Black, Arial, sans-serif' },
          connectgaps: false, cliponaxis: false,
          hovertemplate: `<b>${ym}</b><br>%{x}: %{y:.2f}%<extra></extra>`,
        };
      });

      if (targetFrac != null) {
        const targetY = targetFrac * 100;
        traces.push({
          x: modelsWithData, y: modelsWithData.map(() => targetY), name: `Target ${targetY.toFixed(1)}%`,
          type: 'scatter' as const, mode: 'lines' as const,
          line: { color: '#ef4444', width: 1.6, dash: 'dot', shape: 'linear' as const },
          hovertemplate: `Target: ${targetY.toFixed(1)}%<extra></extra>`,
        });
      }

      const allY = traces.flatMap(tr => (tr.y as (number | null)[]).filter((v): v is number => v != null));
      const minY = allY.length ? Math.min(...allY) : 0;
      const maxY = allY.length ? Math.max(...allY) : 100;
      const pad = Math.max((maxY - minY) * 0.15, 1.5);

      window.Plotly.newPlot(elId, traces, {
        paper_bgcolor: 'transparent', plot_bgcolor: 'transparent',
        // EPCC (rty-chart6-declutter-low-values): margin trái tăng 48→64 vì
        // nhãn 'middle left' của các điểm <90% (xem plotModelLinePanel phía
        // trên) vươn ra bên trái điểm dữ liệu, dễ bị cắt sát mép trái khi
        // điểm đó rơi đúng model đầu tiên trên trục X.
        font: { color: fontColor, size: 11 }, margin: { t: 20, r: 20, b: 44, l: 64 }, showlegend: false,
        xaxis: { tickfont: { size: 10.2, color: fontColor }, gridcolor: gridColor, tickangle: -15 },
        yaxis: { gridcolor: gridColor, tickfont: { size: 10 }, range: [Math.max(0, minY - pad), maxY + pad], ticksuffix: '%' },
        hoverlabel: { font: { size: 10 } },
      }, { displayModeBar: false, responsive: true });
    };
    plotModelLinePanel('rtyTotalMAIN', 'ttl', data.targets.ttl);

    // EPCC (rty-total-chart56-to-spider): Panel 3,4 (SUB1/SUB2 cột+đường) đã
    // được người dùng yêu cầu chuyển sang biểu đồ nhện (radar) giống hệt
    // Chart 1,2 — xem plotSpiderPanel('rtyTotalRadarSUB1','sub1') /
    // ('rtyTotalRadarSUB2','sub2') bên dưới. KHÔNG còn gọi plotMixedPanel cho
    // SUB1/SUB2 nữa (bản cột Sub1/Sub2 + đường Main tham chiếu đã bị thay thế).

    // ═══ Radar TTL/MAIN/SUB1/SUB2 — dùng chung getModelSeriesForStage() ═══
    const plotSpiderPanel = (elId: string, stage: 'ttl' | 'main' | 'sub1' | 'sub2') => {
      const el = document.getElementById(elId);
      if (!el) return;
      const setLegend = stage === 'ttl' ? setSpiderLegendTTL
        : stage === 'main' ? setSpiderLegendMAIN
        : stage === 'sub1' ? setSpiderLegendSUB1
        : setSpiderLegendSUB2;
      const { selected, modelsWithData } = getModelSeriesForStage(stage);
      if (selected.length === 0) { window.Plotly.purge(elId); setLegend([]); return; }
      if (modelsWithData.length === 0) { window.Plotly.purge(elId); setLegend([]); return; }
      const theta = [...modelsWithData, modelsWithData[0]];

      const traces = selected.map((ym, i) => {
        const r = modelsWithData.map(m => {
          const v = data.models[m][stage][ym];
          return v != null ? parseFloat((v * 100).toFixed(2)) : null;
        });
        const dataLabels = r.map(v => v == null ? '' : v.toFixed(1));
        r.push(r[0]); dataLabels.push('');
        const color = RADAR_COLORS[i % RADAR_COLORS.length];
        return {
          type: 'scatterpolar' as const, r, theta, fill: 'toself' as const, fillcolor: color + '1A',
          mode: 'lines+markers+text' as const, marker: { color, size: 5 }, line: { color, width: 2 },
          text: dataLabels, texttemplate: '%{text}', textposition: 'top center' as const,
          textfont: { color, size: 10.8 }, name: ym,
          hovertemplate: `<b>${ym}</b><br>%{theta}: %{r:.2f}%<extra></extra>`,
        };
      });
      setLegend(selected.map((ym, i) => ({ label: ym, color: RADAR_COLORS[i % RADAR_COLORS.length] })));

      const allR = traces.flatMap(tr => tr.r).filter((v): v is number => v != null);
      const minR = allR.length ? Math.max(0, Math.min(...allR)) : 0;
      const maxR = allR.length ? Math.min(100, Math.max(...allR)) : 100;
      const rMin = minR > 85 ? 85 : minR > 75 ? 75 : minR > 50 ? 50 : 0;
      const rMax = maxR < 99.8 ? 100 : 101;

      window.Plotly.newPlot(elId, traces, {
        paper_bgcolor: 'transparent', plot_bgcolor: 'transparent', font: { color: fontColor, size: 11.4 },
        margin: { t: 20, r: 20, b: 20, l: 20 }, showlegend: false,
        polar: {
          bgcolor: 'transparent',
          radialaxis: { visible: true, range: [rMin, rMax], tickfont: { size: 9.6, color: fontColor }, gridcolor: gridColor, angle: 90 },
          angularaxis: { tickfont: { size: 10.8, color: fontColor }, gridcolor: gridColor },
        },
      }, { displayModeBar: false, responsive: true });
    };

    plotSpiderPanel('rtyTotalRadarTTL', 'ttl');
    plotSpiderPanel('rtyTotalRadarMAIN', 'main');
    plotSpiderPanel('rtyTotalRadarSUB1', 'sub1');
    plotSpiderPanel('rtyTotalRadarSUB2', 'sub2');
  }, [plotlyReady, data, selectedModel, xs, isLightMode, tealAccent, modelList]);

  // EPCC (rty-total-toolbar-and-button-fix) - FIX ROOT CAUSE "tab RTY Total
  // trông trống trải sau khi ẩn toolbar Test5, nút 'Tải Excel RTY' hiện ra
  // như link nhỏ/gạch chân thay vì nút bấm rõ ràng": nút trước đây KHÔNG có
  // style inline (height/padding/display) như nút gốc bên Menu5ModelDashboard
  // — NeonButton không tự set kích thước tối thiểu nên co lại theo nội dung
  // chữ. Sửa: (1) thêm inline style giống HỆT nút gốc (height 38px, padding,
  // display inline-flex...) để chắc chắn hiển thị đúng dù class CSS nào có
  // sẵn hay không; (2) bọc trong thanh className="topbar-dash" (class CHUNG
  // đã có sẵn trong App.css, dùng lại y hệt Menu5/RTY) để tab này có 1 thanh
  // tiêu đề nhất quán với 2 tab còn lại — KHÔNG phải thanh filter Test5 cũ
  // (THÁNG BẮT ĐẦU/KẾT THÚC/XEM THEO) đã bị ẩn đúng theo yêu cầu trước, chỉ
  // còn đúng phần liên quan RTY: "DỮ LIỆU CẬP NHẬT ĐẾN" + Model + nút upload.
  const uploadBtnStyle: CSSProperties = {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 8,
    height: 38, minWidth: 150, boxSizing: 'border-box', padding: '0 16px',
    fontSize: 13, fontWeight: 600, borderRadius: 8,
    border: '1px solid rgba(255,255,255,0.35)', background: 'rgba(3,233,244,0.14)',
    color: '#03e9f4', cursor: isParsing ? 'wait' : 'pointer', whiteSpace: 'nowrap',
  };
  const uploadIcon = (
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width={15} height={15} style={{ flexShrink: 0 }}>
      <path d="M21 12a9 9 0 0 1-9 9c-2.52 0-4.93-1-6.74-2.74L3 16" />
      <path d="M3 12a9 9 0 0 1 9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
      <path d="M3 16v5h5" /><path d="M16 3h5v5" />
    </svg>
  );

  // ── Trạng thái rỗng: chưa upload file ──
  if (modelList.length === 0) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div className="topbar-dash" style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10,
          background: '#2F3A1D', borderRadius: 14, padding: '10px 14px',
          border: '1px solid rgba(0,0,0,0.18)', boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
        }}>
          <span style={{ fontSize: 12, fontWeight: 700, color: '#C0EF6A', textTransform: 'uppercase' }}>
            {lang === 'vi' ? 'RTY Total — chưa có dữ liệu' : lang === 'ko' ? 'RTY Total — 데이터 없음' : 'RTY Total — no data yet'}
          </span>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
            onChange={e => { const f = e.target.files?.[0]; if (f) handleUploadFile(f); }} />
          <NeonButton className="btn btn-outline btn-sm" style={uploadBtnStyle} onClick={() => fileInputRef.current?.click()} disabled={isParsing}>
            {uploadIcon}<span>{isParsing ? t.parsing : t.importExcelBtn}</span>
          </NeonButton>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 260, gap: 10 }}>
          <div style={{ fontSize: 14, color: 'var(--text-2)' }}>{t.noDataYet}</div>
          <div style={{ fontSize: 12, color: 'var(--text-3)', maxWidth: 420, textAlign: 'center' }}>{t.emptyDesc}</div>
          {uploadError && <div style={{ fontSize: 12, color: '#e8399a', maxWidth: 480, textAlign: 'center' }}>{uploadError}</div>}
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      {/* ── Toolbar 1 HÀNG NGANG — giữ ĐÚNG vị trí/thứ tự như ảnh tham chiếu
          Mục 4 (RtyDashboard): khối "Đang hiển thị / Dữ liệu cập nhật đến"
          nằm CỐ ĐỊNH BÊN TRÁI cùng hàng (không tách thành dòng riêng phía
          trên), rồi tới THÁNG BẮT ĐẦU / THÁNG KẾT THÚC / MODEL / XEM THEO
          theo đúng thứ tự trái→phải, nút Tải Excel RTY ở cuối hàng bên phải.
          Khác biệt DUY NHẤT có chủ đích so với ảnh gốc: "THÁNG" thay vì
          "NGÀY", "Tháng/Quý/Năm" thay vì "Ngày/Tuần/Tháng" — vì Test6.xlsx
          chỉ có cột YEAR+MONTH, không có ngày/tuần thật. */}
      <div className="topbar-dash" style={{
        display: 'flex', alignItems: 'center', gap: 20, flexWrap: 'wrap',
        background: '#2F3A1D', borderRadius: 14, padding: '10px 14px',
        border: '1px solid rgba(0,0,0,0.18)', boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
      }}>
        {/* Khối trái — "Đang hiển thị" + "Dữ liệu cập nhật đến", 2 dòng xếp dọc */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2, flexShrink: 0 }}>
          <span style={{ fontSize: 12.5, fontWeight: 700, color: '#C0EF6A', whiteSpace: 'nowrap' }}>
            {lang === 'vi' ? 'Đang hiển thị' : lang === 'ko' ? '표시 중' : 'Showing'}:{' '}
            <span style={{ color: '#facc15' }}>⭐ {selectedModel}</span>
            {selectedModel === bestModel && (
              <span style={{ color: 'var(--text-2)', fontWeight: 500 }}>
                {' '}({lang === 'vi' ? 'Hiệu suất tốt nhất' : lang === 'ko' ? '최고 성능' : 'Best performance'})
              </span>
            )}
          </span>
          <span style={{ fontSize: 11, color: 'var(--text-3)', whiteSpace: 'nowrap' }}>
            📅 {lang === 'vi' ? 'Dữ liệu cập nhật đến' : lang === 'ko' ? '데이터 업데이트 기준' : 'Data updated to'}:{' '}
            {data.months.length ? data.months[data.months.length - 1] : '—'}
          </span>
        </div>

        {/* Khối giữa — Tháng bắt đầu / Tháng kết thúc / Model / Xem theo — cùng hàng với khối trái */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 10.5, fontWeight: 700, color: '#C0EF6A', textTransform: 'uppercase' }}>
            {lang === 'vi' ? 'Tháng bắt đầu' : lang === 'ko' ? '시작 월' : 'Start month'}
          </span>
          <input type="month" value={startYm != null ? numToYm(startYm) : ''}
            min={dataBounds ? numToYm(dataBounds.min) : undefined}
            max={endYm != null ? numToYm(endYm) : undefined}
            onChange={e => setStartYm(ymNum(e.target.value))}
            style={{ height: 36, borderRadius: 8, border: '1px solid rgba(0,0,0,0.18)', padding: '0 8px', fontSize: 12.5 }} />
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 10.5, fontWeight: 700, color: '#C0EF6A', textTransform: 'uppercase' }}>
            {lang === 'vi' ? 'Tháng kết thúc' : lang === 'ko' ? '종료 월' : 'End month'}
          </span>
          <input type="month" value={endYm != null ? numToYm(endYm) : ''}
            min={startYm != null ? numToYm(startYm) : undefined}
            max={dataBounds ? numToYm(dataBounds.max) : undefined}
            onChange={e => setEndYm(ymNum(e.target.value))}
            style={{ height: 36, borderRadius: 8, border: '1px solid rgba(0,0,0,0.18)', padding: '0 8px', fontSize: 12.5 }} />
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 10.5, fontWeight: 700, color: '#C0EF6A', textTransform: 'uppercase' }}>{t.model}</span>
          <select value={selectedModel} onChange={e => setSelectedModel(e.target.value)}
            style={{ background: 'rgba(255,255,255,0.92)', color: '#1f2937', border: '1px solid rgba(0,0,0,0.18)', borderRadius: 8, padding: '0 8px', fontSize: 12.5, height: 36 }}>
            {modelList.map(m => <option key={m} value={m}>{m === bestModel ? `⭐ ${m}` : m}</option>)}
          </select>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span style={{ fontSize: 10.5, fontWeight: 700, color: '#C0EF6A', textTransform: 'uppercase' }}>
            {lang === 'vi' ? 'Xem theo' : lang === 'ko' ? '보기 기준' : 'View by'}
          </span>
          <div style={{ display: 'flex', height: 36 }}>
            {(['month', 'quarter', 'year'] as const).map((v, i) => (
              <button key={v} onClick={() => setViewBy(v)} style={{
                padding: '0 12px', fontSize: 12.5, fontWeight: 600,
                borderRadius: i === 0 ? '6px 0 0 6px' : i === 2 ? '0 6px 6px 0' : 0,
                border: '1px solid rgba(0,0,0,0.18)', borderRight: i !== 2 ? 'none' : '1px solid rgba(0,0,0,0.18)',
                background: viewBy === v ? '#2e7d8c' : 'rgba(255,255,255,0.55)',
                color: viewBy === v ? '#fff' : '#7A5A2E', cursor: 'pointer',
              }}>
                {v === 'month' ? (lang === 'vi' ? 'Tháng' : lang === 'ko' ? '월' : 'Month')
                  : v === 'quarter' ? (lang === 'vi' ? 'Quý' : lang === 'ko' ? '분기' : 'Quarter')
                  : (lang === 'vi' ? 'Năm' : lang === 'ko' ? '년' : 'Year')}
              </button>
            ))}
          </div>
        </div>

        {/* Khối phải — nút Tải Excel RTY, đẩy về cuối hàng */}
        <div style={{ flex: 1 }} />
        <input ref={fileInputRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
          onChange={e => { const f = e.target.files?.[0]; if (f) handleUploadFile(f); }} />
        <NeonButton className="btn btn-outline btn-sm" style={uploadBtnStyle} onClick={() => fileInputRef.current?.click()} disabled={isParsing}>
          {uploadIcon}<span>{isParsing ? t.parsing : t.importExcelBtn}</span>
        </NeonButton>
      </div>
      {uploadError && <div style={{ fontSize: 12, color: '#e8399a' }}>{uploadError}</div>}

      {/* ── 4 biểu đồ Radar (TTL / MAIN / SUB1 / SUB2) — layout 2x2 giống ảnh mẫu.
          EPCC (rty-total-chart56-to-spider): Chart 5,6 (SUB1/SUB2) trước đây là
          dạng cột+đường, nay chuyển thành radar theo model y hệt Chart 1,2 —
          tái sử dụng NGUYÊN VĂN plotSpiderPanel(), chỉ đổi stage='sub1'/'sub2'.
          EPCC (rty-total-restore-per-card-legend) - FIX ROOT CAUSE "đã tự ý
          tách legend thành 1 khối riêng bên ngoài và chỉ lấy dữ liệu legend
          của card đầu (TTL) dùng chung cho cả 4 card, sai với đúng vị trí
          người dùng khoanh đỏ (mỗi card có legend RIÊNG ngay trong header của
          chính nó)": trả lại legend riêng cho từng card — TTL dùng
          spiderLegendTTL, MAIN dùng spiderLegendMAIN, SUB1 dùng
          spiderLegendSUB1, SUB2 dùng spiderLegendSUB2 — chỉ chỉnh CSS
          (whiteSpace: 'nowrap' + overflowX: 'auto', gap/font nhỏ hơn) để 7
          mốc tháng nằm gọn 1 dòng trong header, không bị word-wrap xuống 2
          dòng như bản gốc, đúng ý "thu gọn card lại". ── */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
        {([
          // FIX (unify-colors-with-rtydashboard, EPCC): đồng bộ accent + bg
          // với đúng 6 card khoanh đỏ trong RtyDashboard.tsx (ảnh tham chiếu),
          // giữ đúng thứ tự ngữ nghĩa TTL → MAIN → SUB1 → SUB2:
          //  - TTL  ↔ rtySpiderTTL  (#8b5cf6, bg tím)
          //  - MAIN ↔ rtySpiderMAIN (#06b6d4, bg xanh ngọc)
          //  - SUB1 ↔ rtyChartSUB1  (#10b981, bg xanh lá)
          //  - SUB2 ↔ rtyChartSUB2  (#3b82f6, bg xanh dương)
          { id: 'rtyTotalRadarTTL', title: t.radarTTL, accent: '#8b5cf6', bg: isLightMode ? '#d6c6fc' : 'rgba(139,92,246,0.14)', legend: spiderLegendTTL },
          { id: 'rtyTotalRadarMAIN', title: t.radarMAIN, accent: '#06b6d4', bg: isLightMode ? '#a8e5f0' : 'rgba(6,182,212,0.14)', legend: spiderLegendMAIN },
          { id: 'rtyTotalRadarSUB1', title: t.radarSUB1, accent: '#10b981', bg: isLightMode ? '#abe7d3' : 'rgba(16,185,129,0.14)', legend: spiderLegendSUB1 },
          { id: 'rtyTotalRadarSUB2', title: t.radarSUB2, accent: '#3b82f6', bg: isLightMode ? '#bad3fc' : 'rgba(59,130,246,0.14)', legend: spiderLegendSUB2 },
        ] as const).map(panel => (
          <div key={panel.id} className="panel chart-panel">
            <div className="card-header-styled" style={{
              background: panel.bg, borderLeft: `4px solid ${panel.accent}`,
              color: isLightMode ? '#1f2937' : 'var(--text-1)', display: 'flex', alignItems: 'center', gap: 8,
              padding: '5px 10px', fontWeight: 700, fontSize: 12, textTransform: 'uppercase',
            }}>
              <span style={{ flexShrink: 0 }}>{panel.title}</span>
              {/* EPCC (rty-total-legend-wrap-like-chart6): không đủ chỗ thì
                  xuống dòng (flexWrap) — giống đúng cách panel "TTL RTY THEO
                  MODEL (8 THÁNG GẦN NHẤT)" (chart 6) đang hiển thị, không dùng
                  thanh cuộn ngang. */}
              <div style={{
                display: 'flex', gap: 2, flexWrap: 'wrap',
                textTransform: 'none', fontWeight: 500, flex: 1,
              }}>
                {panel.legend.map((lg, i) => <RtyLegendItem key={i} type="line" label={lg.label} color={lg.color} />)}
              </div>
            </div>
            <div className="chart-holder" style={{ height: '100%' }}>
              <div id={panel.id} style={{ width: '100%', height: '100%', minHeight: 300 }} />
            </div>
          </div>
        ))}
      </div>

      {/* ── Panel TTL (cột+đường theo tháng, model đang chọn) + Panel MAIN
          (nay đổi thành đường THẲNG "TTL RTY theo model", lấy tháng gần nhất
          mỗi model — SUB1/SUB2 đã chuyển thành radar ở lưới phía trên) ── */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
        {([
          {
            // TTL ↔ rtyChartTTL trong RtyDashboard (#f59e0b, bg vàng cam)
            id: 'rtyTotalTTL', rowLabel: `${t.panelTTL} (${selectedModel})`, accent: '#f59e0b',
            bg: isLightMode ? '#fcddaa' : 'rgba(255,255,255,0.05)',
            legends: [
              { type: 'bar', label: 'RTY Sub1', color: '#1565C0' },
              { type: 'bar', label: 'RTY Sub2', color: '#ef4444' },
              { type: 'line', label: 'RTY Main', color: tealAccent },
              { type: 'dashed', label: 'RTY TTL', color: '#8b5cf6' },
              { type: 'dotted', label: `Target TTL ${data.targets.ttl != null ? (data.targets.ttl * 100).toFixed(1) + '%' : '—'}`, color: '#ef4444' },
            ],
          },
          {
            // MAIN ↔ rtyChartMAIN trong RtyDashboard (#f97316, bg cam)
            id: 'rtyTotalMAIN', rowLabel: t.ttlByModelLine, accent: '#f97316',
            bg: isLightMode ? '#fdcead' : 'rgba(249,115,22,0.14)',
            // EPCC (rty-total-unify-radar-line-data): legend lấy TRỰC TIẾP từ
            // spiderLegendTTL (state do Chart 1 — plotSpiderPanel('ttl') — set)
            // thay vì khai báo tĩnh riêng, đảm bảo legend 2 chart LUÔN khớp
            // nhau 100% vì cùng 1 nguồn "8 tháng gần nhất".
            legends: [
              ...spiderLegendTTL.map(lg => ({ type: 'line' as const, label: lg.label, color: lg.color })),
              { type: 'dotted' as const, label: `Target TTL ${data.targets.ttl != null ? (data.targets.ttl * 100).toFixed(1) + '%' : '—'}`, color: '#ef4444' },
            ],
          },
        ] as const).map(panel => (
          <div key={panel.id} className="panel chart-panel">
            <div className="card-header-styled" style={{
              background: panel.bg, borderLeft: `4px solid ${panel.accent}`,
              color: isLightMode ? '#1f2937' : 'var(--text-1)', display: 'flex', alignItems: 'center', justifyContent: 'space-between',
              padding: '6px 10px', fontWeight: 700, fontSize: 12, textTransform: 'uppercase',
            }}>
              <span>{panel.rowLabel}</span>
              <div style={{ display: 'flex', alignItems: 'center', gap: 4, textTransform: 'none', fontWeight: 500, flexWrap: 'wrap', justifyContent: 'end' }}>
                {panel.legends.map((lg, idx) => <RtyLegendItem key={idx} type={lg.type as any} label={lg.label} color={lg.color} />)}
              </div>
            </div>
            <div className="chart-holder" style={{ height: '100%' }}>
              <div id={panel.id} style={{ width: '100%', height: '100%', minHeight: 320 }} />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

export default RtyTotalTab;
