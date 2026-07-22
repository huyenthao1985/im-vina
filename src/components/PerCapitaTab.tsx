/**
 * PerCapitaTab.tsx — 인당 생산수 & 생산 현황 (근무인력 기준)
 *
 * Tab 2 trong ManpowerDashboard — vẽ 3 chart theo layout ảnh tham chiếu:
 *
 *  Chart 1 — 인당생산수 (Per Capita Output)
 *    Dữ liệu: Value từ Test_3 × DAY_RATIO → DAY headcount
 *    Hiển thị: Bar DAY + Bar NIGHT (ước lượng) + line TTL + line TARGET (경영목표)
 *    Trục X: tháng (JAN–JUN) + tuần gần (W24–W26) + ngày gần (06/22–06/26)
 *
 *  Chart 2 — 인력 유실 비용 (Labor Loss Cost) — DAY/NIGHT/TTL 인력 유실 비용
 *    EPCC (chart2-loss-cost) — đọc thẳng dữ liệu thật từ Excel (Type "DAY/NIGHT/TTL
 *    인력 유실 비용"), có fallback cộng dồn theo model nếu Excel không có sẵn dòng
 *    model = "TTL" tổng hợp (giống pattern fallback của byModelLabel/근무시간/LINE Q'TY).
 *    Bar DAY + Bar NIGHT + line TTL. (Trước đây Chart 2 lần lượt là 근무시간 Work
 *    Hours rồi LINE Q'TY — nay thay hẳn bằng 인력 유실 비용 theo yêu cầu mới nhất.)
 *
 *  Chart 3 — 근무 인력 현황 (Manpower Trend)
 *    Bar DAY + Bar NIGHT (group) + Line TTL + dashed 기준인력 Standard
 *
 * (Đã gỡ panel "요구 LINE 수 현황" trước đây từng có ở đây — dữ liệu số line
 *  yêu cầu (DAY/NIGHT/TTL LINE Q'TY) nay chỉ hiển thị 1 nơi duy nhất: đường
 *  overlay "TTL LINE Q'TY" trong Chart 4, tránh trùng lặp thông tin.)
 *
 * Tất cả dùng data từ rows (Test_3 TTL ManPower AVG), không hard-code.
 * DAY_RATIO = 0.507, NIGHT_RATIO = 0.493 theo bản tham chiếu.
 * TARGET_PER_CAPITA = 160 (경영목표 기준, có thể override từ stdRow YR24).
 */

import React, { useState, useMemo, useRef, useEffect } from 'react';
import type { DataRow } from '../types';
import { parseManpowerDate } from './ManpowerDashboard';
import { CustomSelect } from './CustomSelect';

// ─── Constants ────────────────────────────────────────────────────────────────
const DAY_RATIO   = 0.507;   // ca ngày / tổng
const NIGHT_RATIO = 0.493;   // ca đêm / tổng
const DEFAULT_TARGET_PC = 160; // 경영목표 인당생산수
// const HOURS_PER_SHIFT = 8;   // 근무시간 chuẩn / ca (có thể chỉnh nếu nhà máy tính OT khác)

/* EPCC (header-legend-merge, tham chiếu đúng pattern SDLegendItem trong
   SalesDashboard.tsx) — legend trước đây do Plotly tự vẽ RỜI bên trong canvas
   (y>1 paper coord), luôn tạo khoảng trắng thừa/đè lên nhãn số phía trên biểu
   đồ dù chỉnh y/margin thế nào. Nay chuyển hẳn legend ra HTML, gộp chung 1
   hàng với tiêu đề trong panel-head (bên phải), tắt showlegend của Plotly. */
const PCLegendItem: React.FC<{ type: 'bar' | 'line' | 'dot'; color: string; label: string }> = ({ type, color, label }) => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '10.5px', fontWeight: 500, whiteSpace: 'nowrap' }}>
    {type === 'bar' && <span style={{ width: '9px', height: '9px', borderRadius: '2px', background: color, flexShrink: 0 }} />}
    {type === 'line' && <span style={{ width: '12px', height: 0, borderTop: `2px dashed ${color}`, flexShrink: 0 }} />}
    {type === 'dot' && <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: color, flexShrink: 0 }} />}
    {label}
  </span>
);

/**
 * getWorkingUnitsForLabel — số ngày công đại diện cho 1 label trên trục X.
 * Dùng để tính 근무시간 (Work Hours) = nhân lực TB x giờ/ca x số ngày công,
 * KHÔNG dùng chung 1 hằng số cho mọi giai đoạn (khác với targetPC cũ) để
 * Chart 2 phản ánh đúng số ngày làm việc thực tế thay đổi theo tháng
 * (vd tháng Tết ít ngày công hơn) — tạo khác biệt thật với Chart 1/Chart 3.
 * Giả định: nghỉ Chủ nhật, làm 6 ngày/tuần (Thứ 2–Thứ 7). Có thể tinh chỉnh
 * thêm ngày lễ VN sau nếu cần độ chính xác cao hơn.
 */
// function getWorkingUnitsForLabel(label: string): number {
//   const s = label.trim().toUpperCase();
  // Tháng: JAN..DEC → đếm số ngày không phải Chủ nhật trong tháng đó (năm 2026,
  // đồng bộ với parseManpowerDate ở ManpowerDashboard.tsx)
  // const monthsAbbr = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  // const monthIdx = monthsAbbr.indexOf(s);
  // if (monthIdx !== -1) {
  //   const daysInMonth = new Date(2026, monthIdx + 1, 0).getDate();
  //   let sundays = 0;
  //   for (let d = 1; d <= daysInMonth; d++) {
  //     if (new Date(2026, monthIdx, d).getDay() === 0) sundays++;
  //   }
  //   return daysInMonth - sundays;
  // }
  // Tuần: W22, W26... → 6 ngày công / tuần
  // if (/^W\d{1,2}$/.test(s)) return 6;
  // Ngày: 06/22, 06/26... → 1 ngày công (trừ khi rơi đúng Chủ nhật)
  // const dayMatch = s.match(/^(\d{2})[/-](\d{2})$/);
  // if (dayMatch) {
  //   const m = parseInt(dayMatch[1], 10) - 1;
  //   const d = parseInt(dayMatch[2], 10);
  //   return new Date(2026, m, d).getDay() === 0 ? 0 : 1;
  // }
  // Năm: mặc định ~ tổng ngày công cả năm
  // if (/^YR/.test(s)) return 313;
  // return 26; // fallback an toàn
// }

const MODEL_COLORS: Record<string, string> = {
  SO1B01:    '#3b82f6',
  SO1C2EF:   '#10b981',
  SO1C2G:    '#f59e0b',
  SO3560:    '#ef4444',
  SO2701:    '#06b6d4',
  TTL:       '#14b8a6',
  SO1C2EH:   '#ec4899',
  SO1C2EDL:  '#14b8a6',
  SO1C2EDM:  '#f43f5e',
  SO1C30S25: '#84cc16',
};

function getModelColor(model: string, idx: number) {
  const u = model.toUpperCase();
  if (MODEL_COLORS[u]) return MODEL_COLORS[u];
  const p = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#ec4899','#14b8a6'];
  return p[idx % p.length];
}

/* EPCC (panel-header-bg-from-template) — NGUỒN MÀU DUY NHẤT cho nền 6 khung
   panel-head bên dưới, lấy NGUYÊN theo đúng thứ tự 1→6 từ THEME.chartPanels
   trong DashboardTemplate.tsx (khung màu đỏ khoanh trong ảnh mẫu). Mọi
   panel-head trong file này chỉ được đọc màu từ mảng này — KHÔNG viết mã hex
   nền rời rạc lần 2 ở bất kỳ panel nào khác, tránh lệch màu khi sửa sau này. */
const PANEL_THEME = [
  { accent: '#8b5cf6', bgLight: '#d6c6fc', bgDark: 'rgba(139,92,246,0.14)' }, // Panel 1 — tím
  { accent: '#06b6d4', bgLight: '#a8e5f0', bgDark: 'rgba(6,182,212,0.14)' },  // Panel 2 — cyan
  { accent: '#f59e0b', bgLight: '#fcddaa', bgDark: 'rgba(255,255,255,0.05)' }, // Panel 3 — cam
  { accent: '#f97316', bgLight: '#fdcead', bgDark: 'rgba(249,115,22,0.14)' }, // Panel 4 — cam-đỏ
  { accent: '#10b981', bgLight: '#abe7d3', bgDark: 'rgba(16,185,129,0.14)' }, // Panel 5 — xanh lá
  { accent: '#3b82f6', bgLight: '#bad3fc', bgDark: 'rgba(59,130,246,0.14)' }, // Panel 6 — xanh dương
] as const;

/* EPCC (panel-flush-header-match-salesdashboard) — đồng bộ KÍCH THƯỚC phần
   panel-head vừa tô màu ở trên với đúng chuẩn kích thước đã dùng trong
   SalesDashboard.tsx (chartHeaderStyle): viền trái 4px, bo góc 8px 8px 0 0,
   padding 10px 14px, margin 0 để thanh màu tiêu đề dán sát 3 cạnh trên/trái/
   phải của khung .panel — đúng như 4 mũi tên đỏ khoanh góc trong ảnh mẫu. */
function panelHeadStyle(i: number, isDark: boolean): React.CSSProperties {
  const c = PANEL_THEME[i];
  return {
    background: isDark ? c.bgDark : c.bgLight,
    borderLeft: `4px solid ${c.accent}`,
    borderRadius: '8px 8px 0 0',
    padding: '10px 14px',
    margin: 0,
  };
}

// ─── i18n ─────────────────────────────────────────────────────────────────────
const T = {
  // FIX (EPCC-title-i18n, áp dụng đồng bộ với ManpowerDashboard.tsx): field
  // `vi` trước đây bị dán NHẦM nguyên văn tiếng Hàn (không có phần tiếng Việt
  // nào) — cùng lỗi đã sửa ở tiêu đề ManpowerDashboard. Sửa lại đúng tiếng
  // Việt, giữ nguyên `en`/`ko` (2 field đó vốn đã đúng ngôn ngữ).
  title:        { vi: 'SẢN LƯỢNG ĐẦU NGƯỜI & TÌNH HÌNH SẢN XUẤT (THEO NHÂN LỰC)', en: 'Per Capita Output & Production (Headcount Based)', ko: '인당 생산수 & 생산 현황 (근무인력 기준)' },
  subtitle:     { vi: 'Ước lượng từ dữ liệu nhân lực (DAY 50.7% / NIGHT 49.3%)', en: 'Estimated from manpower data (DAY 50.7% / NIGHT 49.3%)', ko: '근무인원 기준 추정 (DAY 50.7% / NIGHT 49.3%)' },
  chart2Title:  { vi: 'Chi phí hao hụt nhân lực', en: 'Labor Loss Cost', ko: '인력 유실 비용' },
  dayLossCost:   { vi: 'DAY 인력 유실 비용', en: 'DAY Labor Loss Cost', ko: 'DAY 인력 유실 비용' },
  nightLossCost: { vi: 'NIGHT 인력 유실 비용', en: 'NIGHT Labor Loss Cost', ko: 'NIGHT 인력 유실 비용' },
  ttlLossCost:   { vi: 'TTL 인력 유실 비용', en: 'TTL Labor Loss Cost', ko: 'TTL 인력 유실 비용' },
  chart3Title:  { vi: 'Tình hình nhân lực đi làm', en: 'Manpower Trend', ko: '근무 인력 현황' },
  dayPc:        { vi: 'DAY 인당생산수', en: 'DAY Per Capita', ko: 'DAY 인당생산수' },
  nightPc:      { vi: 'NIGHT 인당생산수', en: 'NIGHT Per Capita', ko: 'NIGHT 인당생산수' },
  ttlPc:        { vi: 'TTL 인당생산수', en: 'TTL Per Capita', ko: 'TTL 인당생산수' },
  target:       { vi: 'TTL TARGET (인당생산수)', en: 'TTL TARGET', ko: 'TTL TARGET (인당생산수)' },
  mgmtTarget:   { vi: '경영목표', en: 'Mgmt Target', ko: '경영목표' },
  dayPro:       { vi: 'DAY 근무시간 Work Hours', en: 'DAY Work Hours', ko: 'DAY 근무시간' },
  nightPro:     { vi: 'NIGHT 근무시간 Work Hours', en: 'NIGHT Work Hours', ko: 'NIGHT 근무시간' },
  ttlPro:       { vi: 'TTL 근무시간 Work Hours', en: 'TTL Work Hours', ko: 'TTL 근무시간' },
  targetPlan:   { vi: 'TTL 기준 근무시간 Standard', en: 'TTL Standard Hours', ko: 'TTL 기준 근무시간' },
  dayMP:        { vi: 'DAY ManPower AVG', en: 'DAY ManPower AVG', ko: 'DAY ManPower AVG' },
  nightMP:      { vi: 'NIGHT ManPower AVG', en: 'NIGHT ManPower AVG', ko: 'NIGHT ManPower AVG' },
  ttlMP:        { vi: 'TTL ManPower AVG', en: 'TTL ManPower AVG', ko: 'TTL ManPower AVG' },
  standard:     { vi: 'TTL 기준인력 Standard AVG', en: 'TTL Standard AVG', ko: 'TTL 기준인력 Standard AVG' },
  noData:       { vi: 'Không có dữ liệu nhân lực', en: 'No manpower data', ko: '인력 데이터 없음' },
  estimatedBadge: { vi: '⚠ Ước lượng', en: '⚠ Estimated', ko: '⚠ 추정값' },
  persons:      { vi: 'người', en: 'prs', ko: '명' },
  unit:         { vi: 'sp/người', en: 'pcs/cap', ko: '개/인' },
  hoursUnit:    { vi: 'giờ', en: 'hrs', ko: '시간' },
  kpiDayPC:     { vi: 'DAY 인당생산수 TB - Sản lượng đầu người ca ngày TB', en: 'AVG DAY Per Capita', ko: 'DAY 인당생산수 평균' },
  kpiNightPC:   { vi: 'NIGHT 인당생산수 TB - Sản lượng đầu người ca đêm TB', en: 'AVG NIGHT Per Capita', ko: 'NIGHT 인당생산수 평균' },
  kpiTtlPC:     { vi: 'TTL 인당생산수 TB - Sản lượng đầu người TTL TB', en: 'AVG TTL Per Capita', ko: 'TTL 인당생산수 평균' },
  kpiTarget:    { vi: '경영목표 - Target', en: 'Mgmt. Target', ko: '경영목표' },
};
function t(key: keyof typeof T, lang: 'vi'|'en'|'ko') {
  return (T[key] as any)[lang] ?? (T[key] as any)['vi'];
}

// EPCC (chart2-loss-cost-vi-keywords) — FIX lỗi Chart 2 (인력 유실 비용 / Chi phí
// hao hụt nhân lực) trống toàn bộ: bộ nhận diện Type trước đây CHỈ khớp từ khóa
// tiếng Hàn ("유실") hoặc tiếng Anh ("loss"+"cost"), bỏ sót hoàn toàn trường hợp
// Excel/Supabase nhập Type bằng TIẾNG VIỆT (vd "DAY Chi phí hao hụt nhân lực",
// "NIGHT hao hụt nhân lực"...) — vốn là ngôn ngữ hiển thị chính của chart này
// (xem chart2Title/chart2Subtitle ở trên). Khi Type không khớp bất kỳ từ khóa
// nào, dòng dữ liệu bị loại NGAY TỪ bước lọc mpRows đầu tiên → 3 map
// dayLossCostByModelLabel/nightLossCostByModelLabel/ttlLossCostByModelLabel
// rỗng cho MỌI model → buildChart2 luôn ra toàn số 0 (đúng như ảnh lỗi: trục Y
// chỉ chạy 0-10, không có cột/đường nào). stripDiacritics giúp so khớp không
// phụ thuộc dấu tiếng Việt (hao hụt / hao hut đều nhận ra như nhau).
function stripDiacritics(s: string): string {
  return s
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/gi, 'd');
}
function isLossCostType(typeStrLower: string): boolean {
  if (typeStrLower.includes('유실') || typeStrLower.includes('loss cost') ||
      (typeStrLower.includes('loss') && typeStrLower.includes('cost'))) {
    return true;
  }
  const n = stripDiacritics(typeStrLower);
  return n.includes('hao hut') || n.includes('thieu hut') ||
         n.includes('ton that nhan luc') || n.includes('chi phi hao hut') ||
         n.includes('hut nhan luc');
}

function r1(n: number | null | undefined): number {
  if (n == null || isNaN(n as number)) return 0;
  return Math.round((n as number) * 10) / 10;
}
function fmt1(n: number) {
  return n.toLocaleString('vi-VN', { maximumFractionDigits: 1 });
}

// EPCC (percapita-dedupe-avg) — fix lỗi biểu đồ cột sai do dữ liệu bị trùng: nếu 1 label
// (model+date+type) vô tình được nạp/sync nhiều lần với ĐÚNG cùng 1 giá trị (VD do
// double-insert lên Supabase), mảng vals sẽ chứa nhiều bản sao giống nhau tuyệt đối.
// Trước đây code lấy trung bình TRỰC TIẾP trên vals gốc — nếu 1 giá trị bất thường/sai
// bị lặp lại nhiều lần hơn giá trị đúng, trung bình sẽ bị kéo lệch theo giá trị sai đó
// (đây là nguyên nhân NIGHT 인당생산수 tháng 5-7 bị đội lên gấp ~2 lần so với các tháng
// trước). Khử trùng (dedupe) theo giá trị TUYỆT ĐỐI GIỐNG NHAU trước khi tính trung bình
// để mỗi giá trị thực chỉ được tính đúng 1 lần.
function avgDedup(vals: number[]): number {
  const unique = Array.from(new Set(vals));
  return unique.reduce((a, b) => a + b, 0) / unique.length;
}

// ─── Build unified label list: months + recent weeks + recent days ────────────
/**
 * Strategy: lấy tất cả labels từ TTL rows, sort theo ngày thật.
 * Trả về mảng labels theo đúng thứ tự chronological.
 */
function buildLabelOrder(
  rows: DataRow[]
): { label: string; dateMs: number }[] {
  const seen: Record<string, number> = {};
  rows.forEach(r => {
    // FIX: Khi load từ Supabase, Excel "Date" được lưu vào cột schema `month`
    // (đồng bộ tự động qua App.tsx: month = String(rawDate)).
    // Nên phải check cả r.month / r.Month làm fallback.
    const raw = String(r.date || (r as any).Date || (r as any).month || (r as any).Month || '').trim().toUpperCase();
    if (!raw || raw === 'YR24' || raw === 'JAN-JUN') return;
    const d = parseManpowerDate(raw);
    if (!d) return;
    if (!seen[raw] || d.getTime() < seen[raw]) {
      seen[raw] = d.getTime();
    }
  });
  return Object.entries(seen)
    .map(([label, dateMs]) => ({ label, dateMs }))
    .sort((a, b) => a.dateMs - b.dateMs);
}

// ─── Data hook ────────────────────────────────────────────────────────────────
interface PCTabData {
  // model → label → avg manpower (từ TTL ManPower AVG)
  byModelLabel: Record<string, Record<string, number>>;
  pcByModelLabel: Record<string, Record<string, number>>; // per capita production data
  // Số line YÊU CẦU theo Type: 'TTL LINE Q'TY' / 'DAY LINE Q'TY' / 'NIGHT LINE Q'TY'
  lineTtlByModelLabel: Record<string, Record<string, number>>;
  lineDayByModelLabel: Record<string, Record<string, number>>;
  lineNightByModelLabel: Record<string, Record<string, number>>;
  ttlByLabel: Record<string, number>;       // TTL headcount per label (từ TTL ManPower AVG)
  // FIX: dữ liệu thật cho Chart 3 (근무 인력 현황) — trước đây bị ước lượng bằng
  // DAY_RATIO/NIGHT_RATIO thay vì đọc đúng "DAY ManPower AVG"/"NIGHT ManPower AVG".
  dayMPByLabel: Record<string, number>;     // model TTL, từ DAY ManPower AVG
  nightMPByLabel: Record<string, number>;   // model TTL, từ NIGHT ManPower AVG
  dayMPByModelLabel: Record<string, Record<string, number>>;
  nightMPByModelLabel: Record<string, Record<string, number>>;
  dayPCByModelLabel: Record<string, Record<string, number>>;
  nightPCByModelLabel: Record<string, Record<string, number>>;
  ttlPCByModelLabel: Record<string, Record<string, number>>;
  targetPCByModelLabel: Record<string, Record<string, number>>;
  // EPCC (chart1-mgmt-goal-real-data) — 경영목표 lấy ĐÚNG dòng "TTL 인당생산수 경영 계획"
  // trong database, tách riêng khỏi targetPCByModelLabel (nguồn cho line "TTL TARGET (인당생산수)")
  mgmtGoalPCByModelLabel: Record<string, Record<string, number>>;
  // FIX: Chart 2 đổi từ "생산 현황 ước lượng" (chỉ là headcount x hằng số, trùng
  // hình dạng Chart 1/3) sang "근무시간 Work Hours" — ưu tiên đọc dữ liệu thật
  // nếu Excel có Type "DAY/NIGHT/TTL 근무시간" hoặc "Work Hours", fallback tính
  // từ nhân lực thật x giờ/ca x số ngày công (khác targetPC cố định trước đây).
  dayWHByModelLabel: Record<string, Record<string, number>>;
  nightWHByModelLabel: Record<string, Record<string, number>>;
  ttlWHByModelLabel: Record<string, Record<string, number>>;
  // EPCC (chart2-loss-cost) — thay thế nguồn dữ liệu Chart 2: từ LINE Q'TY sang
  // "DAY/NIGHT/TTL 인력 유실 비용" (Chi phí hao hụt nhân lực), đọc thẳng từ Excel.
  dayLossCostByModelLabel: Record<string, Record<string, number>>;
  nightLossCostByModelLabel: Record<string, Record<string, number>>;
  ttlLossCostByModelLabel: Record<string, Record<string, number>>;
  ttlStandard: number | null;               // YR24 standard
  allLabels: { label: string; dateMs: number }[];
  activeModels: string[];                   // non-TTL models with data
  lastDataDateMs: number | null;            // mốc thời gian cuối cùng có dữ liệu thật
  lastDataLabel: string | null;             // label tương ứng (để hiển thị "Dữ liệu đến...")
}

function usePCTabData(rows: DataRow[], dateFrom: string, dateTo: string): PCTabData {
  return useMemo(() => {
    const mpRows = rows.filter(r => {
      const typeStr = String(r.type || r.Type || '').toLowerCase();
      // FIX: nhận diện thêm "라인" (tiếng Hàn) bên cạnh "line" (EN) để không bỏ sót
      // dòng dữ liệu Line Q'TY nhập bằng tiếng Hàn.
      // EPCC (chart2-loss-cost) — thêm nhận diện Type "DAY/NIGHT/TTL 인력 유실 비용"
      // (Labor Loss Cost) — nguồn dữ liệu mới thay thế cho LINE Q'TY ở Chart 2.
      // EPCC (chart2-loss-cost-vi-keywords) — dùng isLossCostType() thay vì check
      // '유실'/'loss'+'cost' trực tiếp, để nhận diện thêm Type tiếng Việt
      // ("Chi phí hao hụt nhân lực"...) — nguyên nhân chart bị trống hoàn toàn.
      return typeStr.includes('manpower') || typeStr.includes('인당생산수') || typeStr.includes('line') || typeStr.includes('라인') || typeStr.includes('근무시간') || typeStr.includes('work hour') || typeStr.includes('working hour') || isLossCostType(typeStr);
    });
    if (mpRows.length === 0) {
      return { byModelLabel: {}, pcByModelLabel: {}, lineTtlByModelLabel: {}, lineDayByModelLabel: {}, lineNightByModelLabel: {}, ttlByLabel: {}, dayMPByLabel: {}, nightMPByLabel: {}, dayMPByModelLabel: {}, nightMPByModelLabel: {}, dayPCByModelLabel: {}, nightPCByModelLabel: {}, ttlPCByModelLabel: {}, targetPCByModelLabel: {}, mgmtGoalPCByModelLabel: {}, dayWHByModelLabel: {}, nightWHByModelLabel: {}, ttlWHByModelLabel: {}, dayLossCostByModelLabel: {}, nightLossCostByModelLabel: {}, ttlLossCostByModelLabel: {}, ttlStandard: null, allLabels: [], activeModels: [], lastDataDateMs: null, lastDataLabel: null };
    }

    // Standard
    const stdRow = mpRows.find(r =>
      String(r.model || (r as any).Model || '').trim().toUpperCase() === 'TTL' &&
      // FIX: check cả r.date và r.month (khi load từ Supabase, Date được lưu vào month)
      String(r.date || (r as any).Date || (r as any).month || (r as any).Month || '').trim().toUpperCase() === 'YR24' &&
      String(r.type || (r as any).Type || '').toLowerCase().includes('manpower')
    );
    const ttlStandard = stdRow ? Number(stdRow.value ?? stdRow.Value) : null;

    // Parse all rows with dates
    const parsed = mpRows.map(r => {
      const raw = String(r.date || r.Date || r.month || r.Month || '').trim().toUpperCase();
      const d = parseManpowerDate(raw);
      return { ...r, _label: raw, _date: d };
    }).filter(r => r._date && r._label !== 'YR24' && r._label !== 'JAN-JUN');

    // Date filter
    let filtered = parsed as (typeof parsed[0] & { _date: Date })[];
    const from = dateFrom ? new Date(dateFrom) : null;
    const to   = dateTo   ? new Date(dateTo) : null;
    if (to) to.setHours(23, 59, 59, 999);
    if (from) filtered = filtered.filter(r => r._date! >= from);
    if (to)   filtered = filtered.filter(r => r._date! <= to);

    // Build label order from filtered TTL rows
    const allLabelsRaw = buildLabelOrder(filtered as any).filter(x =>
      !x.label.startsWith('JAN-') && x.label !== 'YR24'
    );

    // Build byModelLabel, pcByModelLabel và 3 nhóm số line YÊU CẦU (TTL/DAY/NIGHT LINE Q'TY)
    // FIX: Test_3.xlsx có 6 biến thể Type song song cho manpower:
    //   DAY ManPower / DAY ManPower AVG / NIGHT ManPower / NIGHT ManPower AVG / TTL ManPower / TTL ManPower AVG
    // Trước đây MỌI type chứa "manpower" bị gom chung vào 1 mảng rồi lấy trung bình
    // → trộn lẫn DAY+NIGHT+TTL và raw+AVG với nhau, ra số sai. Nay tách riêng 3 nhóm
    // và CHỈ nhận đúng bản "... AVG" (bỏ qua bản raw không có AVG) cho từng ca.
    const groups: Record<string, Record<string, number[]>> = {};        // TTL ManPower AVG
    const dayMPGroups: Record<string, Record<string, number[]>> = {};   // DAY ManPower AVG
    const nightMPGroups: Record<string, Record<string, number[]>> = {}; // NIGHT ManPower AVG
    const pcGroups: Record<string, Record<string, number[]>> = {};
    const dayPCGroups: Record<string, Record<string, number[]>> = {};
    const nightPCGroups: Record<string, Record<string, number[]>> = {};
    const ttlPCGroups: Record<string, Record<string, number[]>> = {};
    const targetPCGroups: Record<string, Record<string, number[]>> = {};
    const mgmtGoalPCGroups: Record<string, Record<string, number[]>> = {};
    const dayWHGroups: Record<string, Record<string, number[]>> = {};
    const nightWHGroups: Record<string, Record<string, number[]>> = {};
    const ttlWHGroups: Record<string, Record<string, number[]>> = {};
    const lineTtlGroups: Record<string, Record<string, number[]>> = {};
    const lineDayGroups: Record<string, Record<string, number[]>> = {};
    const lineNightGroups: Record<string, Record<string, number[]>> = {};
    // EPCC (chart2-loss-cost) — 3 nhóm gom dữ liệu "DAY/NIGHT/TTL 인력 유실 비용"
    const dayLossCostGroups: Record<string, Record<string, number[]>> = {};
    const nightLossCostGroups: Record<string, Record<string, number[]>> = {};
    const ttlLossCostGroups: Record<string, Record<string, number[]>> = {};
    (filtered as any[]).forEach(r => {
      const model = String(r.model || r.Model || '').trim().toUpperCase();
      const label = r._label;
      const val   = Number(r.value ?? r.Value);
      const typeStr = String(r.type || r.Type || '').toLowerCase();
      if (!model || isNaN(val) || val == null) return;

      if (typeStr.includes('근무시간') || typeStr.includes('work hour') || typeStr.includes('working hour')) {
        // Type mẫu: "DAY 근무시간", "NIGHT 근무시간", "TTL 근무시간" (hoặc "DAY Work Hours"...)
        let target = ttlWHGroups;
        if (typeStr.includes('day') && !typeStr.includes('holiday')) {
          target = dayWHGroups;
        } else if (typeStr.includes('night')) {
          target = nightWHGroups;
        }
        if (!target[model]) target[model] = {};
        if (!target[model][label]) target[model][label] = [];
        target[model][label].push(val);
      } else if (isLossCostType(typeStr)) {
        // EPCC (chart2-loss-cost) — Type mẫu: "DAY 인력 유실 비용", "NIGHT 인력 유실 비용",
        // "TTL 인력 유실 비용" (Chi phí hao hụt nhân lực theo Ca). Cùng quy tắc phân loại
        // DAY/NIGHT/TTL như 근무시간 / LINE Q'TY ở trên.
        // EPCC (chart2-loss-cost-vi-keywords) — isLossCostType() nay cũng nhận
        // diện Type thuần tiếng Việt ("DAY Chi phí hao hụt nhân lực"...).
        let target = ttlLossCostGroups;
        if (typeStr.includes('day') && !typeStr.includes('holiday')) {
          target = dayLossCostGroups;
        } else if (typeStr.includes('night')) {
          target = nightLossCostGroups;
        }
        if (!target[model]) target[model] = {};
        if (!target[model][label]) target[model][label] = [];
        target[model][label].push(val);
      } else if (typeStr.includes('line') || typeStr.includes('라인')) {
        // Type mẫu: "TTL LINE Q'TY", "DAY LINE Q'TY", "NIGHT LINE Q'TY"
        // Kiểm tra DAY/NIGHT trước vì "TTL" có thể không xuất hiện tường minh
        // trên 1 số dòng — mặc định còn lại (không phải DAY/NIGHT) là TTL.
        let target = lineTtlGroups;
        if (typeStr.includes('day') && !typeStr.includes('holiday')) {
          target = lineDayGroups;
        } else if (typeStr.includes('night')) {
          target = lineNightGroups;
        }
        if (!target[model]) target[model] = {};
        if (!target[model][label]) target[model][label] = [];
        target[model][label].push(val);
      } else if (typeStr.includes('manpower') && typeStr.includes('avg')) {
        // FIX: chỉ nhận bản "... ManPower AVG" — bỏ qua "DAY/NIGHT/TTL ManPower" (raw, không AVG)
        let target = groups; // mặc định TTL (type không chứa rõ "day"/"night")
        if (typeStr.includes('day') && !typeStr.includes('holiday')) {
          target = dayMPGroups;
        } else if (typeStr.includes('night')) {
          target = nightMPGroups;
        }
        if (!target[model]) target[model] = {};
        if (!target[model][label]) target[model][label] = [];
        target[model][label].push(val);
      } else if (typeStr === 'day 인당생산수') {
        if (!dayPCGroups[model]) dayPCGroups[model] = {};
        if (!dayPCGroups[model][label]) dayPCGroups[model][label] = [];
        dayPCGroups[model][label].push(val);
      } else if (typeStr === 'night 인당생산수') {
        if (!nightPCGroups[model]) nightPCGroups[model] = {};
        if (!nightPCGroups[model][label]) nightPCGroups[model][label] = [];
        nightPCGroups[model][label].push(val);
      } else if (typeStr === 'ttl 인당생산수') {
        if (!ttlPCGroups[model]) ttlPCGroups[model] = {};
        if (!ttlPCGroups[model][label]) ttlPCGroups[model][label] = [];
        ttlPCGroups[model][label].push(val);
      } else if (typeStr.includes('target (인당생산수)') || typeStr.includes('target(인당생산수)')) {
        if (typeStr.startsWith('ttl')) {
          if (!targetPCGroups[model]) targetPCGroups[model] = {};
          if (!targetPCGroups[model][label]) targetPCGroups[model][label] = [];
          targetPCGroups[model][label].push(val);
        }
      } else if (typeStr.includes('경영 계획') || typeStr.includes('경영계획')) {
        // EPCC (chart1-mgmt-goal-real-data) — dòng "TTL 인당생산수 경영 계획": nguồn dữ liệu
        // THẬT cho line 경영목표, tách riêng khỏi "TARGET (인당생산수)" ở trên
        if (typeStr.startsWith('ttl')) {
          if (!mgmtGoalPCGroups[model]) mgmtGoalPCGroups[model] = {};
          if (!mgmtGoalPCGroups[model][label]) mgmtGoalPCGroups[model][label] = [];
          mgmtGoalPCGroups[model][label].push(val);
        }
      } else if (typeStr.includes('인당생산수')) {
        if (!pcGroups[model]) pcGroups[model] = {};
        if (!pcGroups[model][label]) pcGroups[model][label] = [];
        pcGroups[model][label].push(val);
      }
    });

    const byModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(groups).forEach(([model, lblMap]) => {
      byModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        byModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    // FIX: aggregate riêng cho DAY ManPower AVG / NIGHT ManPower AVG (dữ liệu thật,
    // dùng cho Chart 3 thay vì công thức ước lượng DAY_RATIO/NIGHT_RATIO cũ)
    const dayMPByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(dayMPGroups).forEach(([model, lblMap]) => {
      dayMPByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        dayMPByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const nightMPByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(nightMPGroups).forEach(([model, lblMap]) => {
      nightMPByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        nightMPByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const pcByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(pcGroups).forEach(([model, lblMap]) => {
      pcByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        pcByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const dayPCByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(dayPCGroups).forEach(([model, lblMap]) => {
      dayPCByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        dayPCByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const nightPCByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(nightPCGroups).forEach(([model, lblMap]) => {
      nightPCByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        nightPCByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const ttlPCByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(ttlPCGroups).forEach(([model, lblMap]) => {
      ttlPCByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        ttlPCByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const targetPCByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(targetPCGroups).forEach(([model, lblMap]) => {
      targetPCByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        targetPCByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const mgmtGoalPCByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(mgmtGoalPCGroups).forEach(([model, lblMap]) => {
      mgmtGoalPCByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        mgmtGoalPCByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const dayWHByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(dayWHGroups).forEach(([model, lblMap]) => {
      dayWHByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        dayWHByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const nightWHByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(nightWHGroups).forEach(([model, lblMap]) => {
      nightWHByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        nightWHByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const ttlWHByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(ttlWHGroups).forEach(([model, lblMap]) => {
      ttlWHByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        ttlWHByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    // EPCC (chart2-loss-cost) — tổng hợp avgDedup cho 3 nhóm 인력 유실 비용
    const dayLossCostByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(dayLossCostGroups).forEach(([model, lblMap]) => {
      dayLossCostByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        dayLossCostByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const nightLossCostByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(nightLossCostGroups).forEach(([model, lblMap]) => {
      nightLossCostByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        nightLossCostByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const ttlLossCostByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(ttlLossCostGroups).forEach(([model, lblMap]) => {
      ttlLossCostByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        ttlLossCostByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const lineTtlByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(lineTtlGroups).forEach(([model, lblMap]) => {
      lineTtlByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        lineTtlByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const lineDayByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(lineDayGroups).forEach(([model, lblMap]) => {
      lineDayByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        lineDayByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    const lineNightByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(lineNightGroups).forEach(([model, lblMap]) => {
      lineNightByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        lineNightByModelLabel[model][lbl] = avgDedup(vals);
      });
    });

    // Nếu không có model TTL trong dữ liệu, tính tổng động từ các model khác
    if (!byModelLabel['TTL']) {
      byModelLabel['TTL'] = {};
      allLabelsRaw.forEach(x => {
        let sum = 0;
        let hasValue = false;
        Object.keys(byModelLabel).forEach(m => {
          if (m !== 'TTL') {
            const v = byModelLabel[m][x.label];
            if (v != null && v > 0) {
              sum += v;
              hasValue = true;
            }
          }
        });
        if (hasValue) {
          byModelLabel['TTL'][x.label] = sum;
        }
      });
    }

    const ttlByLabel: Record<string, number> = byModelLabel['TTL'] ?? {};

    // Nếu không có model TTL cho DAY/NIGHT ManPower AVG, tính tổng động từ các model khác
    // (tương tự fallback của byModelLabel['TTL'] ở trên)
    if (!dayMPByModelLabel['TTL']) {
      dayMPByModelLabel['TTL'] = {};
      allLabelsRaw.forEach(x => {
        let sum = 0; let hasValue = false;
        Object.keys(dayMPByModelLabel).forEach(m => {
          if (m !== 'TTL') {
            const v = dayMPByModelLabel[m][x.label];
            if (v != null && v > 0) { sum += v; hasValue = true; }
          }
        });
        if (hasValue) dayMPByModelLabel['TTL'][x.label] = sum;
      });
    }
    if (!nightMPByModelLabel['TTL']) {
      nightMPByModelLabel['TTL'] = {};
      allLabelsRaw.forEach(x => {
        let sum = 0; let hasValue = false;
        Object.keys(nightMPByModelLabel).forEach(m => {
          if (m !== 'TTL') {
            const v = nightMPByModelLabel[m][x.label];
            if (v != null && v > 0) { sum += v; hasValue = true; }
          }
        });
        if (hasValue) nightMPByModelLabel['TTL'][x.label] = sum;
      });
    }
    const dayMPByLabel: Record<string, number> = dayMPByModelLabel['TTL'] ?? {};
    const nightMPByLabel: Record<string, number> = nightMPByModelLabel['TTL'] ?? {};

    // EPCC (chart2-line-qty-ttl-fallback) — FIX lỗi Chart 2 (LINE Q'TY) không hiện
    // dữ liệu: Excel thường nhập "DAY/NIGHT/TTL LINE Q'TY" theo TỪNG MODEL (SO1B01,
    // SO1C2EF...) chứ không có sẵn dòng model = "TTL" tổng hợp. buildChart2 lại chỉ
    // đọc đúng modelKey "TTL" nên luôn ra toàn số 0/rỗng — dù Chart 4 vẫn hiển thị
    // đúng vì nó quét qua TẤT CẢ model. Áp dụng lại đúng pattern fallback đã dùng
    // cho byModelLabel/dayMPByModelLabel/nightMPByModelLabel ở trên: nếu chưa có
    // model "TTL" tường minh, cộng dồn từ các model khác cho từng label.
    // EPCC (chart2-line-qty-ttl-fallback) — FIX lỗi Chart 2 (LINE Q'TY) không hiện
    // dữ liệu: Excel thường nhập "DAY/NIGHT/TTL LINE Q'TY" theo TỪNG MODEL (SO1B01,
    // SO1C2EF...) chứ không có sẵn dòng model = "TTL" tổng hợp. buildChart2 lại chỉ
    // đọc đúng modelKey "TTL" nên luôn ra toàn số 0/rỗng — dù Chart 4 vẫn hiển thị
    // đúng vì nó quét qua TẤT CẢ model. Áp dụng lại đúng pattern fallback đã dùng
    // cho byModelLabel/dayMPByModelLabel/nightMPByModelLabel ở trên: nếu chưa có
    // model "TTL" tường minh, cộng dồn từ các model khác cho từng label.
    // EPCC (chart2-negative-values) — thêm tham số allowNegative: LINE Q'TY/
    // headcount không bao giờ âm nên giữ nguyên lọc `v > 0` (bỏ qua rác/placeholder
    // âm nếu có); nhưng Chi phí hao hụt nhân lực (LossCost) có thể âm hợp lệ
    // (tháng tiết kiệm được chi phí so với kế hoạch) — nếu chỉ cộng số dương sẽ
    // làm mất hẳn phần đóng góp âm, ra tổng TTL sai lệch hẳn so với thực tế.
    const sumFallbackTTL = (
      target: Record<string, Record<string, number>>,
      allowNegative: boolean = false
    ) => {
      if (target['TTL']) return; // đã có dữ liệu TTL thật, không ghi đè
      target['TTL'] = {};
      allLabelsRaw.forEach(x => {
        let sum = 0; let hasValue = false;
        Object.keys(target).forEach(m => {
          if (m !== 'TTL') {
            const v = target[m][x.label];
            if (v != null && (allowNegative || v > 0)) { sum += v; hasValue = true; }
          }
        });
        if (hasValue) target['TTL'][x.label] = sum;
      });
    };
    sumFallbackTTL(lineTtlByModelLabel);
    sumFallbackTTL(lineDayByModelLabel);
    sumFallbackTTL(lineNightByModelLabel);
    sumFallbackTTL(dayLossCostByModelLabel, true);
    sumFallbackTTL(nightLossCostByModelLabel, true);
    sumFallbackTTL(ttlLossCostByModelLabel, true);

    // ─── Trim các label "kế tiếp" chưa có dữ liệu thật ─────────────────────
    // File Excel nguồn (Test_3) dựng khung sẵn cho cả năm → các ngày/tuần/
    // tháng chưa tới vẫn xuất hiện thành label nhưng TTL = 0/blank.
    // Ở đây ta tìm mốc thời gian (dateMs) MỚI NHẤT trong các label có dữ liệu
    // thật (TTL > 0), rồi bỏ hết các label nằm SAU mốc đó khỏi trục X.
    // Lưu ý: các khoảng trống NẰM TRONG vùng dữ liệu (ví dụ nghỉ Tết) vẫn giữ
    // nguyên — chỉ phần đuôi sau cùng (chưa phát sinh dữ liệu) mới bị ẩn.
    const labelsWithData = allLabelsRaw.filter(x => (ttlByLabel[x.label] ?? 0) > 0);
    const lastDataDateMs = labelsWithData.length > 0
      ? Math.max(...labelsWithData.map(x => x.dateMs))
      : null;
    const allLabels = lastDataDateMs != null
      ? allLabelsRaw.filter(x => x.dateMs <= lastDataDateMs)
      : allLabelsRaw;
    const lastDataLabel = lastDataDateMs != null
      ? labelsWithData.reduce((latest, x) => (x.dateMs > latest.dateMs ? x : latest)).label
      : null;

    const activeModels = Object.keys(byModelLabel)
      .filter(m => m !== 'TTL')
      .sort();

    return { byModelLabel, pcByModelLabel, lineTtlByModelLabel, lineDayByModelLabel, lineNightByModelLabel, ttlByLabel, dayMPByLabel, nightMPByLabel, dayMPByModelLabel, nightMPByModelLabel, dayPCByModelLabel, nightPCByModelLabel, ttlPCByModelLabel, targetPCByModelLabel, mgmtGoalPCByModelLabel, dayWHByModelLabel, nightWHByModelLabel, ttlWHByModelLabel, dayLossCostByModelLabel, nightLossCostByModelLabel, ttlLossCostByModelLabel, ttlStandard, allLabels, activeModels, lastDataDateMs, lastDataLabel };
  }, [rows, dateFrom, dateTo]);
}

// ─── Chart helpers ────────────────────────────────────────────────────────────

/**
 * Chart 1: 인당생산수
 * X = allLabels; Y = headcount × DAY_RATIO hoặc × NIGHT_RATIO → ước lượng "인당생산수"
 * Hiển thị: grouped bar (DAY + NIGHT) + line TTL + line 경영목표
 */
function buildChart1(
  data: PCTabData,
  labels: string[],
  textColor: string,
  gridColor: string,
  lang: 'vi'|'en'|'ko',
  targetPC: number,
  isDark: boolean,
  modelFilter: string = 'all'
) {
  // Màu accent đậm hơn ở light mode để chữ/đường không bị mờ trên nền sáng
  const tealAccent = isDark ? '#14b8a6' : '#0f766e';
  const roseAccent    = isDark ? '#f43f5e' : '#be123c';

  const modelKey = modelFilter === 'all' ? 'TTL' : modelFilter.trim().toUpperCase();
  const modelMP = data.byModelLabel[modelKey] ?? {};
  const dayPCMap   = data.dayPCByModelLabel[modelKey] ?? {};
  const nightPCMap = data.nightPCByModelLabel[modelKey] ?? {};
  const ttlPCMap   = data.ttlPCByModelLabel[modelKey] ?? {};

  // Sử dụng dữ liệu thực tế từ Excel (DAY/NIGHT/TTL 인당생산수) nếu có, fallback về công thức tính từ nhân lực
  const dayVals    = labels.map(l => {
    const val = dayPCMap[l];
    return r1(val != null && val > 0 ? val : (modelMP[l] ?? 0) * DAY_RATIO);
  });
  const nightVals  = labels.map(l => {
    const val = nightPCMap[l];
    return r1(val != null && val > 0 ? val : (modelMP[l] ?? 0) * NIGHT_RATIO);
  });
  const ttlVals    = labels.map(l => {
    const val = ttlPCMap[l];
    return r1(val != null && val > 0 ? val : (modelMP[l] ?? 0));
  });
  const targetVals = labels.map(l => {
    const customTarget = data.targetPCByModelLabel[modelKey]?.[l];
    return customTarget != null && customTarget > 0 ? r1(customTarget) : targetPC;
  });
  // EPCC (chart1-mgmt-goal-real-data) — 경영목표 lấy đúng dòng "TTL 인당생산수 경영 계획"
  // trong database; chỉ fallback về targetPC (hằng số) khi giai đoạn đó chưa có dữ liệu.
  const mgmtGoalVals = labels.map(l => {
    const customGoal = data.mgmtGoalPCByModelLabel[modelKey]?.[l];
    return customGoal != null && customGoal > 0 ? r1(customGoal) : targetPC;
  });

  const traces: any[] = [
    {
      x: labels, y: dayVals, name: t('dayPc', lang),
      type: 'bar', marker: { color: '#1565C0' },
      text: dayVals.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      x: labels, y: nightVals, name: t('nightPc', lang),
      type: 'bar', marker: { color: '#ef4444' },
      text: nightVals.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      x: labels, y: ttlVals, name: t('ttlPc', lang),
      type: 'scatter', mode: 'lines+markers+text',
      yaxis: 'y2',
      line: { color: tealAccent, width: 1.6, shape: 'spline', smoothing: 1 },
      marker: { color: tealAccent, size: 6 },
      text: ttlVals.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'top center',
      textfont: { size: 10, color: tealAccent, family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      hovertemplate: `<b>${t('ttlPc', lang)}</b>: %{y}<extra></extra>`,
    },
    {
      x: labels, y: targetVals, name: t('target', lang),
      type: 'scatter', mode: 'lines+markers+text',
      yaxis: 'y2',
      line: { color: '#f59e0b', width: 1.6, dash: 'dot', shape: 'spline', smoothing: 1 },
      marker: { color: '#f59e0b', size: 5 },
      text: targetVals.map(v => v > 0 ? fmt1(v) : ''),
      // EPCC (chart1-target-label-below) — đường TTL 인당생산수 (teal) và đường
      // TTL TARGET (vàng, nét chấm) trước đây CÙNG để nhãn số 'top center' nên
      // đè số lên nhau khi 2 đường ở gần nhau (đúng vùng khoanh đỏ người dùng
      // gửi). Chuyển nhãn đường vàng xuống 'bottom center' — đường teal vẫn
      // hiện số ở trên, đường vàng hiện số ở dưới, không còn chồng chữ.
      textposition: 'bottom center',
      textfont: { size: 9, color: '#f59e0b', family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      hovertemplate: `<b>${t('target', lang)}</b>: %{y}<extra></extra>`,
    },
    {
      x: labels, y: mgmtGoalVals, name: t('mgmtTarget', lang),
      type: 'scatter', mode: 'lines+markers+text',
      yaxis: 'y2',
      line: { color: roseAccent, width: 1.6, dash: 'dash', shape: 'spline', smoothing: 1 },
      marker: { color: roseAccent, size: 5 },
      text: mgmtGoalVals.map(v => v > 0 ? fmt1(v) : ''),
      // EPCC (chart1-mgmtgoal-label-above) — theo yêu cầu người dùng: nhãn số
      // của đường 경영목표 (đỏ, nét đứt) chuyển từ 'bottom center' lên 'top center'.
      textposition: 'top center',
      textfont: { size: 9, color: roseAccent, family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      hovertemplate: `<b>${t('mgmtTarget', lang)}</b>: %{y}<extra></extra>`,
    },
  ];

  // EPCC (chart1-lines-fully-separated) — trước đây dùng 1 công thức range đơn giản
  // (lineMin*0.55 → lineMax*1.2) nên vùng đường vẫn còn chồng lên đỉnh cột ở vài
  // điểm (VD tháng có cột cao gần bằng đường). Nay tính toán CHÍNH XÁC 2 vùng
  // KHÔNG GIAO NHAU theo tỉ lệ chiều cao khung chart:
  //   - Cột (yaxis)  chiếm 0%  → 55% chiều cao
  //   - 3 line (yaxis2) chiếm hẳn 63% → 95% chiều cao (cách cột 8% để không dính)
  const BAR_TOP_FRAC = 0.55;
  const LINE_BOTTOM_FRAC = 0.63;
  const LINE_TOP_FRAC = 0.95;

  const lineValsAll = [...ttlVals, ...targetVals, ...mgmtGoalVals].filter(v => v > 0);
  const lineMin = lineValsAll.length ? Math.min(...lineValsAll) : 0;
  const lineMax = lineValsAll.length ? Math.max(...lineValsAll) : 10;
  // D = độ rộng khoảng range y2 cần thiết để lineMin/lineMax rơi đúng vào
  // 2 mốc tỉ lệ LINE_BOTTOM_FRAC/LINE_TOP_FRAC; fallback khi 3 line gần như
  // bằng nhau (lineMax≈lineMin) để tránh range suy biến (rộng = 0).
  const spanFrac = LINE_TOP_FRAC - LINE_BOTTOM_FRAC;
  const rawSpan = lineMax - lineMin;
  const D = rawSpan > 0.01 ? rawSpan / spanFrac : Math.max(lineMax * 0.3, 10);
  const y2Min = lineMin - LINE_BOTTOM_FRAC * D;
  const y2Max = y2Min + D;
  const y2Range = [y2Min, y2Max];

  const maxBarVal = Math.max(...dayVals, ...nightVals, 0);
  const yRange = [0, maxBarVal > 0 ? maxBarVal / BAR_TOP_FRAC : 10];

  const layout = {
    barmode: 'group',
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor, size: 11 },
    margin: { t: 12, r: 40, b: 28, l: 48 },
    showlegend: false,
    xaxis: { tickfont: { size: 11, color: textColor }, gridcolor: gridColor },
    yaxis: { gridcolor: gridColor, tickfont: { size: 10 }, range: yRange },
    yaxis2: {
      overlaying: 'y', side: 'right', showgrid: false,
      tickfont: { size: 10, color: textColor }, range: y2Range,
    },
    hoverlabel: { font: { size: 10 } },
  };
  return { traces, layout };
}

/**
 * Chart 2: 인력 유실 비용 (Labor Loss Cost)
 * EPCC (chart2-loss-cost) — thay hẳn nguồn dữ liệu từ LINE Q'TY sang đúng 3 dòng
 * Type thật trong database: "DAY 인력 유실 비용" / "NIGHT 인력 유실 비용" /
 * "TTL 인력 유실 비용" (đã được parse sẵn thành dayLossCostByModelLabel /
 * nightLossCostByModelLabel / ttlLossCostByModelLabel ở usePCTabData, có fallback
 * cộng dồn theo model nếu Excel không có sẵn dòng model = "TTL").
 * Không dùng công thức ước lượng/fallback từ nhân lực vì chi phí hao hụt là số
 * liệu nhập thẳng từ Excel, không suy ra được từ headcount.
 */
function buildChart2(
  data: PCTabData,
  labels: string[],
  textColor: string,
  gridColor: string,
  lang: 'vi'|'en'|'ko',
  _targetPC: number,
  isDark: boolean,
  modelFilter: string = 'all'
) {
  const tealAccent = isDark ? '#14b8a6' : '#0f766e';

  const modelKey = modelFilter === 'all' ? 'TTL' : modelFilter.trim().toUpperCase();
  const dayCostMap   = data.dayLossCostByModelLabel[modelKey] ?? {};
  const nightCostMap = data.nightLossCostByModelLabel[modelKey] ?? {};
  const ttlCostMap   = data.ttlLossCostByModelLabel[modelKey] ?? {};

  const dayCost = labels.map(l => r1(dayCostMap[l] ?? 0));
  const nightCost = labels.map(l => r1(nightCostMap[l] ?? 0));
  // EPCC (chart2-negative-values) — FIX: Chi phí hao hụt nhân lực (인력 유실 비용)
  // là số liệu THỰC TẾ có thể ÂM (tháng nào hao hụt ít/âm nghĩa là tiết kiệm được
  // chi phí so với kế hoạch) — KHÔNG phải lỗi hay thiếu dữ liệu. Trước đây điều
  // kiện `real > 0` coi TTL âm là "không có dữ liệu thật" rồi tự ý thay bằng
  // dayCost+nightCost (thường ra kết quả khác hẳn, sai lệch số liệu thật) — chỉ
  // nên fallback khi ô đó THỰC SỰ không có dữ liệu (null/undefined), không phải
  // khi có dữ liệu thật nhưng là số âm.
  const ttlCost = labels.map((l, i) => {
    const real = ttlCostMap[l];
    return r1(real != null ? real : dayCost[i] + nightCost[i]);
  });

  const traces: any[] = [
    {
      x: labels, y: dayCost, name: t('dayLossCost', lang),
      type: 'bar', marker: { color: '#1565C0' },
      text: dayCost.map(v => v !== 0 ? fmt1(v) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
      hovertemplate: `<b>${t('dayLossCost', lang)}</b>: %{y}<extra></extra>`,
    },
    {
      x: labels, y: nightCost, name: t('nightLossCost', lang),
      type: 'bar', marker: { color: '#ef4444' },
      text: nightCost.map(v => v !== 0 ? fmt1(v) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
      hovertemplate: `<b>${t('nightLossCost', lang)}</b>: %{y}<extra></extra>`,
    },
    {
      x: labels, y: ttlCost, name: t('ttlLossCost', lang),
      type: 'scatter', mode: 'lines+markers+text',
      line: { color: tealAccent, width: 1.6, shape: 'spline', smoothing: 1 },
      marker: { color: tealAccent, size: 6 },
      text: ttlCost.map(v => v !== 0 ? fmt1(v) : ''),
      textposition: 'top center',
      // EPCC (chart2-negative-red-label) — số dương giữ nguyên màu teal hiện tại;
      // số âm chuyển sang đỏ để phân biệt nhanh tháng lỗ/tiết kiệm mà không cần
      // nhìn dấu trừ (theo yêu cầu người dùng, khoanh đỏ các điểm âm trên chart).
      textfont: {
        size: 10,
        color: ttlCost.map(v => v < 0 ? '#ef4444' : tealAccent),
        family: 'Arial Black, Arial, sans-serif',
      },
      cliponaxis: false,
      hovertemplate: `<b>${t('ttlLossCost', lang)}</b>: %{y}<extra></extra>`,
    },
  ];

  // EPCC (chart2-negative-values) — FIX: trục Y trước đây LUÔN neo cứng ở 0
  // (`range: [0, ...]`) — hợp lý cho số liệu chỉ dương (headcount, sản lượng…)
  // nhưng với Chi phí hao hụt nhân lực (có thể âm/dương xen kẽ theo tháng),
  // ép sàn = 0 khiến MỌI cột/đường âm bị đẩy ra ngoài vùng nhìn thấy của chart
  // → nhìn như "mất dữ liệu" dù số liệu vẫn có, chỉ là âm. Nay tính range theo
  // CẢ min lẫn max thực tế của 3 chuỗi dữ liệu, có đệm 15% mỗi đầu, và chỉ giữ
  // sàn = 0 khi toàn bộ dữ liệu không âm (giữ nguyên hành vi cũ cho các chart
  // number luôn dương khác nếu tái sử dụng hàm này).
  const allVals = [...dayCost, ...nightCost, ...ttlCost];
  const maxVal = Math.max(...allVals, 0);
  const minVal = Math.min(...allVals, 0);
  const pad = Math.max(Math.abs(maxVal), Math.abs(minVal)) * 0.15 || 10;
  const yRange = [minVal < 0 ? minVal - pad : 0, maxVal > 0 ? maxVal + pad : 10];

  const layout = {
    barmode: 'group',
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor, size: 11 },
    margin: { t: 12, r: 12, b: 28, l: 56 },
    showlegend: false,
    xaxis: { tickfont: { size: 11, color: textColor }, gridcolor: gridColor },
    yaxis: { gridcolor: gridColor, tickfont: { size: 10 }, range: yRange, zeroline: true, zerolinecolor: gridColor, zerolinewidth: 1.5 },
    hoverlabel: { font: { size: 10 } },
  };
  return { traces, layout };
}

/**
 * Chart 3: 근무 인력 현황
 * Bar DAY + Bar NIGHT (group) + Line TTL + dashed Standard
 */
function buildChart3(
  data: PCTabData,
  labels: string[],
  textColor: string,
  gridColor: string,
  lang: 'vi'|'en'|'ko',
  isDark: boolean,
  modelFilter: string = 'all'
) {
  const tealAccent = isDark ? '#14b8a6' : '#0f766e';
  const roseAccent    = isDark ? '#f43f5e' : '#be123c';

  const modelKey = modelFilter === 'all' ? 'TTL' : modelFilter.trim().toUpperCase();
  const modelMP = data.byModelLabel[modelKey] ?? {};

  // FIX: dùng đúng dữ liệu thật "DAY ManPower AVG"/"NIGHT ManPower AVG" từ Excel
  // (qua data.dayMPByLabel/nightMPByLabel) thay vì công thức ước lượng DAY_RATIO/
  // NIGHT_RATIO cũ. Chỉ fallback về công thức ước lượng nếu period đó thực sự
  // không có dữ liệu DAY/NIGHT ManPower AVG (an toàn, tránh chart bị trống).
  const dayMP = labels.map(l => {
    const real = data.dayMPByModelLabel[modelKey]?.[l];
    return r1(real != null && real > 0 ? real : (modelMP[l] ?? 0) * DAY_RATIO);
  });
  const nightMP = labels.map(l => {
    const real = data.nightMPByModelLabel[modelKey]?.[l];
    return r1(real != null && real > 0 ? real : (modelMP[l] ?? 0) * NIGHT_RATIO);
  });
  const ttlMP   = labels.map(l => r1(modelMP[l] ?? 0));
  const stdVal  = data.ttlStandard ?? 0;
  const stdLine = labels.map(() => r1(stdVal));

  const traces: any[] = [
    {
      // FIX: đổi DAY ManPower AVG từ line → bar (column), đồng bộ với Chart1/Chart2
      x: labels, y: dayMP, name: t('dayMP', lang),
      type: 'bar', marker: { color: '#1565C0' },
      text: dayMP.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'inside',
      textfont: { size: 13, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      // FIX: đổi NIGHT ManPower AVG từ line → bar (column)
      x: labels, y: nightMP, name: t('nightMP', lang),
      type: 'bar', marker: { color: '#ef4444' },
      text: nightMP.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'inside',
      textfont: { size: 13, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      // TTL ManPower AVG: giữ nguyên dạng line để vẫn thấy xu hướng tổng
      x: labels, y: ttlMP, name: t('ttlMP', lang),
      type: 'scatter', mode: 'lines+markers+text',
      line: { color: tealAccent, width: 2.5, shape: 'spline', smoothing: 1 },
      marker: { color: tealAccent, size: 8, symbol: 'circle' },
      text: ttlMP.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'top center',
      textfont: { size: 13, color: tealAccent, family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      hovertemplate: `<b>${t('ttlMP', lang)}</b>: %{y}<extra></extra>`,
    },
  ];

  if (stdVal > 0) {
    traces.push({
      x: labels, y: stdLine, name: t('standard', lang),
      type: 'scatter', mode: 'lines',
      line: { color: roseAccent, width: 2, dash: 'dashdot', shape: 'spline', smoothing: 1 },
    });
  }

  // FIX: trục Y trước đây tính maxVal = max(ttlMP, stdVal, dayMP, nightMP) — nếu
  // stdVal (chuẩn nhân lực YR24) bị lệch/sai định dạng và lớn hơn nhiều lần dữ
  // liệu thực (vd 49.894 trong khi dữ liệu thực chỉ quanh 2.900–7.400), nó kéo
  // giãn toàn bộ trục Y lên rất cao → cột DAY/NIGHT và line TTL thật bị dồn sát
  // đáy biểu đồ, trông như "gạch ngang" phẳng lì phía trên (đúng lỗi trong ảnh).
  // Cách sửa: trục Y chỉ auto-scale theo DỮ LIỆU THỰC (dayMP/nightMP/ttlMP).
  // stdVal chỉ được tính vào nếu nó cùng bậc độ lớn với dữ liệu thực (≤ 3 lần
  // giá trị lớn nhất) — nếu lệch quá xa thì bỏ qua khi tính range (đường standard
  // vẫn được vẽ, chỉ có thể bị cắt ở mép trên thay vì phá vỡ cả biểu đồ).
  const maxDataVal = Math.max(...ttlMP, ...dayMP, ...nightMP, 0);
  const stdInRange = stdVal > 0 && stdVal <= maxDataVal * 3 ? stdVal : 0;
  const maxVal = Math.max(maxDataVal, stdInRange, 0);
  const yRange = [0, maxVal > 0 ? maxVal * 1.15 : 10];

  const layout = {
    barmode: 'group', // FIX: cho DAY/NIGHT đứng cạnh nhau dạng cột thay vì đè line
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor, size: 14 },
    margin: { t: 18, r: 15, b: 35, l: 65 },
    showlegend: false,
    xaxis: { tickfont: { size: 14, color: textColor }, gridcolor: gridColor },
    yaxis: { gridcolor: gridColor, tickfont: { size: 13 }, range: yRange },
    hoverlabel: { font: { size: 13 } },
  };
  return { traces, layout };
}

/**
 * resolveChart4Info — NGUỒN SỰ THẬT DUY NHẤT cho Chart 4.
 * Trước đây badge "● Kèm số line yêu cầu" (JSX) và phần vẽ line thật trong
 * buildChart4 dùng 2 điều kiện khác nhau (badge check toàn cục theo mọi model/
 * giai đoạn, còn line trace chỉ check theo modelsWithData ở 1 giai đoạn cụ thể)
 * → có thể lệch pha: badge nói "có" nhưng chart không vẽ, hoặc ngược lại.
 * Hàm này gộp lại 1 chỗ, dùng chung cho cả buildChart4 lẫn phần badge/hint trong JSX.
 */
function resolveChart4Info(data: PCTabData, displayLabels: string[]) {
  const models = data.activeModels;

  // Tìm label mới nhất có ít nhất 1 model có dữ liệu nhân lực
  let latestLabel = '';
  for (let i = displayLabels.length - 1; i >= 0; i--) {
    const lbl = displayLabels[i];
    const hasAny = models.some(m => (data.byModelLabel[m]?.[lbl] ?? 0) > 0);
    if (hasAny) { latestLabel = lbl; break; }
  }
  if (!latestLabel && displayLabels.length > 0) latestLabel = displayLabels[displayLabels.length - 1];

  // Model có nhân lực > 0 tại latestLabel (nền tảng của trục X / cột)
  // FIX (chart4-max-10-models, EPCC): sắp theo nhân lực GIẢM DẦN trước khi giới
  // hạn, để nếu quá 10 model thì giữ lại 10 model có nhân lực LỚN NHẤT (thông
  // tin quan trọng nhất của chart), thay vì cắt theo thứ tự activeModels tùy ý.
  const headcountModels = models
    .filter(m => (data.byModelLabel[m]?.[latestLabel] ?? 0) > 0)
    .sort((a, b) => (data.byModelLabel[b]?.[latestLabel] ?? 0) - (data.byModelLabel[a]?.[latestLabel] ?? 0));

  // Số line YÊU CẦU — quét TẤT CẢ label (mọi tầng: ngày/tuần/tháng/năm, giảm dần
  // theo thời gian) để tìm giai đoạn gần nhất có dữ liệu Line TTL, độc lập với
  // granularity nhân lực đang chọn (Line Q'TY thường chỉ nhập theo Tháng).
  // FIX: quét theo TẤT CẢ activeModels (không giới hạn trong headcountModels) để
  // không bỏ sót model có Line Q'TY nhưng lại thiếu/0 nhân lực ở latestLabel.
  const allLabelsDesc = [...data.allLabels].sort((a, b) => b.dateMs - a.dateMs);
  let lineLabel = '';
  for (const { label } of allLabelsDesc) {
    const hasAny = models.some(m => (data.lineTtlByModelLabel[m]?.[label] ?? 0) > 0);
    if (hasAny) { lineLabel = label; break; }
  }

  // Model nào có Line Q'TY ở lineLabel
  const lineModels = lineLabel
    ? models.filter(m => (data.lineTtlByModelLabel[m]?.[lineLabel] ?? 0) > 0)
    : [];

  // FIX (chart4-max-10-models, EPCC): giới hạn tối đa 10 model hiển thị trên
  // trục X, giống giới hạn 10 giai đoạn của Chart 5 (Spider). Ưu tiên giữ
  // headcountModels (đã sort giảm dần theo nhân lực) trước — model đông nhân
  // lực nhất không bao giờ bị cắt; model chỉ-có-Line-không-nhân-lực lấp các chỗ
  // trống còn lại nếu có. Nếu tổng số model < 10 thì mảng tự ngắn hơn — bar/label
  // dư KHÔNG được vẽ (ẩn hẳn), không có cột trống hay placeholder.
  const modelsWithData = Array.from(new Set([...headcountModels, ...lineModels]))
    .sort((a, b) => {
      const ia = headcountModels.indexOf(a); const ib = headcountModels.indexOf(b);
      const ra = ia === -1 ? Infinity : ia; const rb = ib === -1 ? Infinity : ib;
      return ra - rb || a.localeCompare(b);
    })
    .slice(0, 10);

  const hasLineData = lineLabel !== '' && modelsWithData.some(m => (data.lineTtlByModelLabel[m]?.[lineLabel] ?? 0) > 0);

  return { latestLabel, modelsWithData, lineLabel, hasLineData };
}

/**
 * Chart 4: Phân bố nhân lực theo Model — Giai đoạn gần nhất
 * Bar chart: X = activeModels, Y = headcount tại label mới nhất có dữ liệu
 * + Line chart (trục y phụ): số line YÊU CẦU theo Model tại cùng giai đoạn (Type = 'TTL LINE Q'TY')
 * Tối đa 10 model trên trục X (ưu tiên model nhân lực lớn nhất — xem
 * resolveChart4Info); nếu ít hơn 10 model có dữ liệu thì tự ẩn, không có cột
 * trống/placeholder.
 */
function buildChart4(
  data: PCTabData,
  displayLabels: string[],
  textColor: string,
  gridColor: string,
  lang: 'vi'|'en'|'ko',
  isDark: boolean
) {
  const { latestLabel, modelsWithData, lineLabel, hasLineData } = resolveChart4Info(data, displayLabels);

  // FIX (chart4-line-color-theme, EPCC): màu #facc15 (vàng nhạt) trước đây cố
  // định cho mọi theme — trên nền sáng (light mode) độ tương phản quá thấp,
  // khó đọc điểm/nhãn số. Đổi sang màu vàng cam đậm hơn cho light mode
  // (#b45309, cùng tông nhưng tối hơn nhiều để đủ tương phản trên nền trắng),
  // giữ nguyên #facc15 sáng cho dark mode (đã rõ trên nền tối).
  const lineAccent = isDark ? '#facc15' : '#b45309';

  const yVals = modelsWithData.map(m => r1(data.byModelLabel[m]?.[latestLabel] ?? 0));
  const colors = modelsWithData.map((m, i) => getModelColor(m, i));

  const lineTtlVals   = modelsWithData.map(m => data.lineTtlByModelLabel[m]?.[lineLabel] ?? null);
  const lineDayVals   = modelsWithData.map(m => data.lineDayByModelLabel[m]?.[lineLabel] ?? null);
  const lineNightVals = modelsWithData.map(m => data.lineNightByModelLabel[m]?.[lineLabel] ?? null);

  const traces: any[] = [{
    x: modelsWithData,
    y: yVals,
    type: 'bar',
    name: lang === 'vi' ? 'Nhân lực' : lang === 'ko' ? '인원' : 'Headcount',
    marker: { color: colors },
    text: yVals.map(v => v > 0 ? String(v) : ''),
    textposition: 'outside',
    textfont: { size: 13, color: textColor },
    hovertemplate: '%{x}: %{y}<extra></extra>',
  }];

  if (hasLineData) {
    const customdata = modelsWithData.map((_, i) => [
      lineDayVals[i] != null ? r1(lineDayVals[i] as number) : null,
      lineNightVals[i] != null ? r1(lineNightVals[i] as number) : null,
    ]);
    const dayLbl   = lang === 'vi' ? 'DAY' : lang === 'ko' ? 'DAY' : 'DAY';
    const nightLbl = lang === 'vi' ? 'NIGHT' : lang === 'ko' ? 'NIGHT' : 'NIGHT';
    // Nếu giai đoạn có dữ liệu Line khác giai đoạn đang hiển thị của cột nhân lực
    // (VD: nhân lực xem theo Ngày "06/25" nhưng Line chỉ có theo Tháng "JUN"),
    // ghi rõ giai đoạn đó ngay trong tên trace để không gây hiểu lầm 2 giá trị
    // cùng 1 thời điểm.
    const linePeriodSuffix = lineLabel && lineLabel !== latestLabel ? ` [${lineLabel}]` : '';
    traces.push({
      x: modelsWithData,
      y: lineTtlVals.map(v => v != null ? r1(v) : null),
      type: 'scatter',
      mode: 'lines+markers+text',
      name: (lang === 'vi' ? "Số line yêu cầu (TTL LINE Q'TY)" : lang === 'ko' ? "요구 라인 수 (TTL LINE Q'TY)" : "Required Lines (TTL LINE Q'TY)") + linePeriodSuffix,
      yaxis: 'y2',
      line: { color: lineAccent, width: 2, shape: 'spline', smoothing: 1 },
      marker: { color: lineAccent, size: 9, symbol: 'diamond' },
      text: lineTtlVals.map(v => v != null && v > 0 ? fmt1(r1(v)) : ''),
      textposition: 'top center',
      textfont: { size: 12, color: lineAccent, family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      customdata,
      hovertemplate:
        `${lang === 'vi' ? 'Số line yêu cầu (TTL)' : lang === 'ko' ? '요구 라인 수 (TTL)' : 'Required Lines (TTL)'}: %{y}` +
        `<br>${dayLbl}: %{customdata[0]}` +
        `<br>${nightLbl}: %{customdata[1]}<extra></extra>`,
    });
  }

  const lineMax = hasLineData ? Math.max(...lineTtlVals.filter((v): v is number => v != null)) : 0;

  const layout: any = {
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor, size: 13 },
    margin: { t: 40, r: hasLineData ? 60 : 20, b: 70, l: 60 },
    xaxis: { tickfont: { size: 13, color: textColor }, gridcolor: gridColor, type: 'category' },
    yaxis: { gridcolor: gridColor, tickfont: { size: 12 }, title: { text: lang === 'vi' ? 'Nhân lực (người)' : lang === 'ko' ? '인원 (명)' : 'Headcount (prs)', font: { size: 13 } } },
    showlegend: hasLineData,
    legend: { orientation: 'h', y: -0.18, font: { size: 11, color: textColor } },
    annotations: latestLabel ? [{
      text: `${lang === 'vi' ? 'Giai đoạn' : lang === 'ko' ? '기간' : 'Period'}: ${latestLabel}`,
      xref: 'paper', yref: 'paper',
      x: 1, y: 1.05, xanchor: 'right', yanchor: 'bottom',
      showarrow: false,
      font: { size: 12, color: textColor },
    }] : [],
  };

  if (hasLineData) {
    layout.yaxis2 = {
      overlaying: 'y',
      side: 'right',
      showgrid: false,
      tickfont: { size: 12, color: lineAccent },
      title: { text: lang === 'vi' ? 'Số line yêu cầu' : lang === 'ko' ? '요구 라인 수' : 'Required Lines', font: { size: 13, color: lineAccent } },
      range: [0, lineMax > 0 ? lineMax * 1.4 : 10],
    };
  }

  return { traces, layout, hasLineData, lineLabel, latestLabel };
}

/**
 * Chart 5: Phân bố Model theo Giai đoạn (Spider / Radar)
 * Mỗi trace = 1 time label (giai đoạn), mỗi axis = 1 model
 * Chọn tối đa 10 giai đoạn gần nhất có dữ liệu để chart không quá rối
 * Nguồn dữ liệu (EPCC spider-data-source): ttlPCByModelLabel (đúng, số hàng
 * trăm) → fallback pcByModelLabel → fallback byModelLabel (manpower)
 */
function buildChart5(
  data: PCTabData,
  displayLabels: string[],
  textColor: string,
  _gridColor: string,
  _lang: 'vi'|'en'|'ko'
) {
  const models = data.activeModels;
  if (models.length === 0 || displayLabels.length === 0) {
    return { traces: [], layout: {}, legend: [] as { label: string; color: string }[] };
  }

  // FIX (spider-data-source, EPCC): trước đây luôn ưu tiên pcByModelLabel — map
  // "gom nhặt" (fallback) MỌI type chỉ cần chứa chữ 인당생산수, kể cả các dòng
  // KHÔNG PHẢI giá trị sản lượng thật (dòng target/tỉ lệ...) → ra số lẻ nhỏ
  // 0.x-2.x sai. Nguồn ĐÚNG là ttlPCByModelLabel (lấy riêng từ type
  // 'ttl 인당생산수', cùng nguồn với KPI "avgTtl"), luôn ra số tự nhiên hàng
  // trăm, cùng bậc với DEFAULT_TARGET_PC = 160. Thứ tự ưu tiên mới:
  // ttlPCByModelLabel → pcByModelLabel → byModelLabel (manpower, fallback cuối).
  const hasTtlPcData = Object.keys(data.ttlPCByModelLabel || {}).length > 0;
  const hasPcData = Object.keys(data.pcByModelLabel || {}).length > 0;
  const activeDataMap = hasTtlPcData
    ? data.ttlPCByModelLabel
    : hasPcData
    ? data.pcByModelLabel
    : data.byModelLabel;

  // Lọc các label có ít nhất 1 model có dữ liệu, lấy tối đa 10 label cuối
  // (nâng từ 6 → 10 theo yêu cầu; model/trục không bị giới hạn số lượng)
  const labelsWithData = displayLabels.filter(lbl =>
    models.some(m => (activeDataMap[m]?.[lbl] ?? 0) > 0)
  );
  const selectedLabels = labelsWithData.slice(-10);

  // Mở rộng bảng màu từ 6 → 10 màu để không lặp màu khi hiển thị đủ 10 giai đoạn
  const radarColors = [
    '#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4',
    '#ec4899','#84cc16','#14b8a6','#f43f5e',
  ];

  // Chỉ giữ models có ít nhất 1 giá trị > 0 trong các selectedLabels
  const modelsWithData = models.filter(m =>
    selectedLabels.some(lbl => (activeDataMap[m]?.[lbl] ?? 0) > 0)
  );
  if (modelsWithData.length === 0) return { traces: [], layout: {}, legend: [] as { label: string; color: string }[] };

  // Đóng vòng radar: thêm model đầu vào cuối
  const modelsClosed = [...modelsWithData, modelsWithData[0]];

  const traces: any[] = selectedLabels.map((lbl, i) => {
    // FIX (spider-integer-format, EPCC): dữ liệu ttlPCByModelLabel là số tự
    // nhiên hàng trăm (sp/người) — làm tròn nguyên (Math.round) thay vì giữ 1
    // số thập phân (r1/fmt1, vốn hợp lý hơn cho map pcByModelLabel số lẻ cũ).
    const rVals = modelsClosed.map(m => Math.round(activeDataMap[m]?.[lbl] ?? 0));
    // FIX(spider-labels): hiển thị giá trị từng điểm theo đúng màu của trace để
    // đọc số trực tiếp trên biểu đồ, không cần hover. Điểm cuối bị lặp lại
    // (đóng vòng radar) nên bỏ trống label để tránh chồng số lên chính nó.
    const textVals = rVals.map((v, idx) =>
      idx === rVals.length - 1 ? '' : (v > 0 ? v.toLocaleString('vi-VN') : '')
    );
    return {
      type: 'scatterpolar',
      r: rVals,
      theta: modelsClosed,
      fill: 'toself',
      // EPCC (percapita-radar-unify-rtytotal) - FIX theo yêu cầu người dùng
      // "lấy cách làm của RTY Total... đồng nhất từ vòng ring trong lẫn phủ
      // màu mờ bên trong": đổi alpha fillcolor từ '22' (~13%) sang '1A'
      // (~10%), ĐÚNG giá trị `plotSpiderPanel()` dùng trong RtyTotalTab.tsx
      // (`fillcolor: color + '1A'`) — cùng độ mờ lớp phủ, không còn lệch
      // giữa 2 tab dùng chung "ngôn ngữ" biểu đồ nhện.
      fillcolor: radarColors[i % radarColors.length] + '1A',
      line: { color: radarColors[i % radarColors.length], width: 2 },
      mode: 'lines+markers+text',
      marker: { color: radarColors[i % radarColors.length], size: 5 },
      text: textVals,
      texttemplate: '%{text}',
      textposition: 'top center',
      // EPCC (percapita-radar-unify-rtytotal): cỡ chữ nhãn số liệu 10→10.8,
      // ĐÚNG `textfont.size` của plotSpiderPanel() trong RtyTotalTab.tsx.
      textfont: { color: radarColors[i % radarColors.length], size: 10.8, family: 'inherit' },
      name: lbl,
      hovertemplate: `${lbl}<br>%{theta}: %{r}<extra></extra>`,
    };
  });

  // EPCC (percapita-radar-unify-rtytotal) - THÊM MỚI: tự động "zoom" trục
  // bán kính theo dữ liệu thực, ĐÚNG nguyên lý rMin/rMax của
  // plotSpiderPanel() trong RtyTotalTab.tsx (radar RTY dùng ngưỡng cố định
  // 0/50/75/85 vì luôn là thang % 0-100 — Chart 5 ở đây là số lượng
  // sản phẩm/người, không có trần cố định, nên tổng quát hoá bằng cách chọn
  // "bước tròn đẹp" (niceStep) theo độ lớn dữ liệu thay vì hard-code %).
  // Mục đích giống hệt RTY: tránh toàn bộ đa giác dồn cụm sát viền ngoài khi
  // các giá trị đều cao & gần nhau, giúp đọc rõ khác biệt giữa các model hơn
  // (đúng vấn đề "vòng ring trong" người dùng phản ánh ở ảnh 1 — trục luôn cố
  // định 0→250 dù dữ liệu thật dồn hết ở nửa ngoài).
  const allR = traces.flatMap(tr => (tr.r as number[])).filter(v => v > 0);
  const minR = allR.length ? Math.min(...allR) : 0;
  const maxR = allR.length ? Math.max(...allR) : 100;
  const niceStep = maxR > 500 ? 100 : maxR > 200 ? 50 : maxR > 80 ? 20 : maxR > 20 ? 10 : 5;
  const rMin = Math.max(0, Math.floor((minR * 0.9) / niceStep) * niceStep);
  const rMax = Math.ceil((maxR * 1.08) / niceStep) * niceStep;

  const legend = selectedLabels.map((lbl, i) => ({ label: lbl, color: radarColors[i % radarColors.length] }));

  const layout = {
    paper_bgcolor: 'transparent',
    polar: {
      bgcolor: 'transparent',
      radialaxis: {
        visible: true,
        range: [rMin, rMax],
        color: textColor,
        gridcolor: textColor + '30',
        // EPCC (percapita-radar-unify-rtytotal): cỡ chữ trục bán kính
        // 12→9.6 + thêm `angle: 90`, ĐÚNG `radialaxis` của plotSpiderPanel()
        // trong RtyTotalTab.tsx.
        tickfont: { size: 9.6, color: textColor },
        angle: 90,
      },
      angularaxis: {
        // EPCC (percapita-radar-unify-rtytotal): cỡ chữ trục góc (tên model)
        // 13→10.8, ĐÚNG `angularaxis.tickfont.size` của plotSpiderPanel().
        tickfont: { size: 10.8, color: textColor },
        gridcolor: textColor + '20',
      },
    },
    font: { color: textColor, size: 11.4 },
    // EPCC (percapita-radar-unify-rtytotal): margin thu gọn về
    // {t:20,r:20,b:20,l:20} + tắt legend nội bộ Plotly (showlegend:false),
    // ĐÚNG layout của plotSpiderPanel() — legend giờ hiển thị ở header panel
    // (component ngoài Plotly) bằng PCLegendItem, cùng kiểu RtyLegendItem
    // RTY Total đang dùng, thay vì legend dọc bên phải chiếm nhiều chỗ như
    // trước (margin.r từng phải để 85).
    margin: { t: 20, r: 20, b: 20, l: 20 },
    showlegend: false,
  };
  return { traces, layout, legend };
}

// ─── Production (DAY / NIGHT / TTL) data & charts ──────────────────────────
// FIX(prod-shift-not-division): Dữ liệu thật trong Excel KHÔNG có cột Division
// = SUB1/SUB2/MAIN cho sản xuất. Thay vào đó, cột Type chứa trực tiếp các giá
// trị theo CA (shift) + Model, ví dụ:
//   "DAY PRON'D PLAN" / "DAY Pro Actual"
//   "NIGHT PRON'D PLAN" / "NIGHT Pro Actual"
//   "TTL PRON'D PLAN" / "TTL Pro Actual"
// cho từng ngày (Date) và từng Model (cột Model). Điều kiện lọc theo Division
// SUB1/SUB2/MAIN trước đây là SAI (không khớp cấu trúc file thật) khiến 3 chart
// luôn trống. Sửa: lọc theo Type bắt đầu bằng day/night/ttl + chứa plan/actual,
// nhóm theo shift (DAY/NIGHT/TTL) × label (ngày), có áp dụng modelFilter để
// xem theo từng Model hoặc cộng dồn tất cả Model khi chọn "Tất cả".

interface ProdLineData {
  planByLabel:   Record<string, number>;
  actualByLabel: Record<string, number>;
  hasData: boolean;
}

interface AllProdData {
  day:       ProdLineData;
  night:     ProdLineData;
  ttl:       ProdLineData;
  allLabels: { label: string; dateMs: number }[];
}

type Shift = 'day' | 'night' | 'ttl';

/**
 * Trích xuất dữ liệu sản xuất Plan / Actual theo ca DAY / NIGHT / TTL.
 * Tìm rows có Type bắt đầu bằng "day"/"night"/"ttl" VÀ chứa "plan"/"actual"
 * (vd "DAY PRON'D PLAN", "TTL Pro Actual"...), lọc thêm theo Model nếu có chọn.
 * KHÔNG ước tính — nếu không tìm được dữ liệu thật thì hasData = false.
 */
function useProdData(
  rows: DataRow[],
  dateFrom: string,
  dateTo: string,
  modelFilter: string = 'all'
): AllProdData {
  return useMemo(() => {
    const empty: ProdLineData = { planByLabel: {}, actualByLabel: {}, hasData: false };

    const prodRows = rows.filter(r => {
      const ts = String(r.type || (r as any).Type || '').trim().toLowerCase();
      const isPlan = ts.includes('plan') || ts.includes('target');
      const isActual = ts.includes('actual') || (ts.includes('인당생산수') && !ts.includes('target') && !ts.includes('avg') && !ts.includes('달성율'));
      if (!isPlan && !isActual) return false;

      const hasShift = ts.startsWith('day') || ts.startsWith('night') || ts.startsWith('ttl');
      if (!hasShift) return false;

      const m = String(r.model || (r as any).Model || '').trim().toUpperCase();
      if (modelFilter !== 'all') {
        if (m !== modelFilter.trim().toUpperCase()) return false;
      } else {
        // Khi chọn "Tất cả" (all), loại bỏ model TTL để tránh bị cộng đúp
        // giữa tổng tự tính với tổng TTL có sẵn trong Excel
        if (m === 'TTL') return false;
      }
      return true;
    });

    if (prodRows.length === 0) {
      return { day: empty, night: empty, ttl: empty, allLabels: [] };
    }

    const parsed = prodRows.map(r => {
      const raw = String(r.date || (r as any).Date || (r as any).month || (r as any).Month || '').trim().toUpperCase();
      const d = parseManpowerDate(raw);
      return { ...r, _label: raw, _date: d };
    }).filter(r => r._date && r._label !== 'YR24' && r._label !== 'JAN-JUN');

    let filtered = parsed as (typeof parsed[0] & { _date: Date })[];
    const from = dateFrom ? new Date(dateFrom) : null;
    const to   = dateTo   ? new Date(dateTo)   : null;
    if (to) to.setHours(23, 59, 59, 999);
    if (from) filtered = filtered.filter(r => r._date! >= from);
    if (to)   filtered = filtered.filter(r => r._date! <= to);

    const allLabels = buildLabelOrder(filtered as any).filter(x =>
      !x.label.startsWith('JAN-') && x.label !== 'YR24'
    );

    // Nhóm dữ liệu: shift × kind × label × values[] (cộng dồn nếu >1 Model khớp)
    const acc: Record<Shift, { plan: Record<string, number[]>; actual: Record<string, number[]> }> = {
      day:   { plan: {}, actual: {} },
      night: { plan: {}, actual: {} },
      ttl:   { plan: {}, actual: {} },
    };

    (filtered as any[]).forEach(r => {
      const ts  = String(r.type || (r as any).Type || '').trim().toLowerCase();
      const lbl = (r as any)._label as string;
      const val = Number((r as any).value ?? (r as any).Value);
      if (isNaN(val)) return;

      const shift: Shift | null = ts.startsWith('day') ? 'day' : ts.startsWith('night') ? 'night' : ts.startsWith('ttl') ? 'ttl' : null;
      if (!shift) return;

      let kind: 'plan' | 'actual' | null = null;
      if (ts.includes('plan') || ts.includes('target')) {
        kind = 'plan';
      } else if (ts.includes('actual') || (ts.includes('인당생산수') && !ts.includes('target') && !ts.includes('avg') && !ts.includes('달성율'))) {
        kind = 'actual';
      }
      if (!kind) return;

      if (!acc[shift][kind][lbl]) acc[shift][kind][lbl] = [];
      acc[shift][kind][lbl].push(val);
    });

    const processShift = (shift: Shift): ProdLineData => {
      const planByLabel: Record<string, number> = {};
      const actualByLabel: Record<string, number> = {};

      // Nhiều Model cùng khớp (vd modelFilter='all') → cộng dồn theo label
      Object.entries(acc[shift].plan).forEach(([lbl, vals]) => {
        planByLabel[lbl] = vals.reduce((a, b) => a + b, 0);
      });
      Object.entries(acc[shift].actual).forEach(([lbl, vals]) => {
        actualByLabel[lbl] = vals.reduce((a, b) => a + b, 0);
      });

      return {
        planByLabel,
        actualByLabel,
        hasData: Object.keys(planByLabel).length > 0 || Object.keys(actualByLabel).length > 0,
      };
    };

    return {
      day:   processShift('day'),
      night: processShift('night'),
      ttl:   processShift('ttl'),
      allLabels,
    };
  }, [rows, dateFrom, dateTo, modelFilter]);
}

/**
 * Chart sản xuất cho 1 ca (DAY / NIGHT / TTL)
 * Bar Plan (orange) + Bar Actual (teal) + Line Yield% (navy dashed, trục Y phụ)
 * KHÔNG ước tính — nếu không có dữ liệu thật, component hiện thông báo.
 */
function buildProductionChart(
  lineData: ProdLineData,
  labels: string[],
  textColor: string,
  gridColor: string,
  isDark: boolean
) {
  const planColor   = isDark ? '#f97316' : '#ea580c';
  const actualColor = '#14b8a6';
  const yieldColor  = '#1565C0';

  const planVals   = labels.map(l => r1(lineData.planByLabel[l]   ?? 0));
  const actualVals = labels.map(l => r1(lineData.actualByLabel[l] ?? 0));
  const yieldVals: (number | null)[] = labels.map((_, i) => {
    const p = planVals[i];
    const a = actualVals[i];
    if (!p || p === 0) return null;
    return Math.round((a / p) * 100);
  });

  const fmtQty = (v: number) =>
    v.toLocaleString('vi-VN', { maximumFractionDigits: 0 });

  const traces: any[] = [
    {
      x: labels, y: planVals, name: 'Plan',
      type: 'bar', marker: { color: planColor },
      text: planVals.map(v => v > 0 ? fmtQty(v) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      x: labels, y: actualVals, name: 'Actual',
      type: 'bar', marker: { color: actualColor },
      text: actualVals.map(v => v > 0 ? fmtQty(v) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      x: labels, y: yieldVals, name: 'Yield %',
      type: 'scatter', mode: 'lines+markers+text',
      yaxis: 'y2',
      line: { color: yieldColor, width: 2, dash: 'dot' },
      marker: { color: yieldColor, size: 6, symbol: 'circle' },
      text: yieldVals.map(v => v !== null && v > 0 ? `${v}%` : ''),
      textposition: 'top center',
      textfont: { size: 10, color: yieldColor, family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      connectgaps: false,
    },
  ];

  const maxQty      = Math.max(...planVals, ...actualVals, 0);
  const validYields = yieldVals.filter((v): v is number => v !== null);
  const maxYield    = validYields.length > 0 ? Math.max(...validYields) : 0;

  const layout: any = {
    barmode: 'group',
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor, size: 10 },
    margin: { t: 12, r: 48, b: 28, l: 48 },
    showlegend: false,
    xaxis: { tickfont: { size: 10, color: textColor }, gridcolor: gridColor },
    yaxis: {
      gridcolor: gridColor,
      tickfont: { size: 9 },
      range: [0, maxQty > 0 ? maxQty * 1.25 : 100],
    },
    yaxis2: {
      overlaying: 'y', side: 'right',
      showgrid: false,
      tickfont: { size: 9, color: yieldColor },
      ticksuffix: '%',
      range: [-20, Math.max(maxYield > 0 ? maxYield * 1.3 : 0, 120)],
    },
    hoverlabel: { font: { size: 10 } },
  };

  return { traces, layout };
}

// ─── Component ────────────────────────────────────────────────────────────────
interface PerCapitaTabProps {
  rows: DataRow[];
  theme: 'light' | 'dark';
  lang: 'vi' | 'en' | 'ko';
  onToggleTheme: () => void;
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
  onFileSelected: (file: File, wb: any) => void;
  /** Khi true = tab đang được hiển thị → trigger Plotly.react để re-draw */
  isVisible?: boolean;
}

export const PerCapitaTab: React.FC<PerCapitaTabProps> = ({
  // FIX(upload-1-lan): PerCapitaTab không còn nút Tải Excel riêng — dữ liệu
  // luôn đến từ prop `rows` dùng chung với tab Nhân lực (ManpowerDashboard),
  // vốn đã được App.tsx nạp 1 lần cho cả 2 tab. Giữ prop trong interface để
  // không phá signature/API của component, chỉ đánh dấu unused ở đây.
  rows, theme, lang, onToggleTheme: _onToggleTheme, setLang: _setLang, onFileSelected: _onFileSelected, isVisible = true,
}) => {
  const isDark = theme === 'dark';
  const textColor = isDark ? '#e2e8f0' : '#1e293b';
  const gridColor = isDark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.06)';
  // FIX(toolbar-chuan-hoa): đồng bộ y hệt màu label với TargetActualDashboard (ảnh 3)
  // FIX (EPCC-vanilla-toolbar): thanh filter đổi nền sang "Vanilla"
  // (#FFF4D6, cố định, không đổi theo theme) theo ảnh tham chiếu — nên màu
  // chữ nhãn cũng cố định luôn (cam đậm) thay vì đổi theo theme như trước.
  const filterLabelColor = '#C0EF6A';

  // ── Fix: "biểu đồ hiện khung trống, không có hình/số" khi mới mở trang ──
  // Root cause: script Plotly load từ CDN (bất đồng bộ) có thể CHƯA sẵn sàng
  // tại thời điểm effect vẽ chart chạy lần đầu. Guard cũ chỉ kiểm tra
  // `typeof window.Plotly === 'undefined'` rồi bail ra — và vì không có gì
  // trigger effect chạy lại sau khi Plotly tải xong, chart bị TRỐNG VĨNH VIỄN
  // dù data đã có sẵn. Fix: poll tới khi Plotly sẵn sàng rồi set state để
  // effect vẽ chart tự chạy lại.
  const [plotlyReady, setPlotlyReady] = useState<boolean>(
    typeof window !== 'undefined' && !!(window as any).Plotly
  );
  useEffect(() => {
    if (plotlyReady) return;
    const id = setInterval(() => {
      if (typeof window !== 'undefined' && (window as any).Plotly) {
        setPlotlyReady(true);
        clearInterval(id);
      }
    }, 150);
    return () => clearInterval(id);
  }, [plotlyReady]);

  // ── Toolbar states (đồng nhất với ManpowerDashboard) ──────────────────────
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo,   setDateTo]   = useState('');
  // FIX (EPCC-default-view-by-day): mặc định khi mở dashboard phải là "Ngày"
  // (day) theo yêu cầu, đồng bộ với ManpowerDashboard (sheet còn lại), không
  // còn mặc định "Tháng" (month) như trước.
  const [granularity, setGranularity] = useState<'day'|'week'|'month'|'year'>('day');
  const [modelFilter, setModelFilter] = useState<string>('all');
  const [dateError, setDateError] = useState<string | null>(null);
  // EPCC (percapita-radar-unify-rtytotal) - THÊM MỚI: legend động của Chart 5
  // (Spider), set bên trong effect vẽ Plotly sau khi buildChart5() tính ra
  // selectedLabels/màu — ĐÚNG pattern `spiderLegendTTL` (state do
  // plotSpiderPanel() set) đã dùng ở RtyTotalTab.tsx, để header panel hiển
  // thị legend bằng component ngoài (PCLegendItem) thay vì legend nội bộ
  // của Plotly.
  const [chart5Legend, setChart5Legend] = useState<{ label: string; color: string }[]>([]);

  // FIX: đồng hồ realtime đã bị bỏ khỏi Mục 3 sheet 2 theo yêu cầu — không
  // còn cần state/interval/formatter riêng cho clock ở đây nữa.

  // ── Date handlers (giống ManpowerDashboard) ────────────────────────────────
  const handleDateFromChange = (val: string) => {
    if (dateTo && val > dateTo) {
      setDateError(lang === 'vi' ? 'Từ ngày phải nhỏ hơn hoặc bằng Đến ngày' : lang === 'ko' ? '시작일은 종료일보다 이전이어야 합니다' : 'From Date must be <= To Date');
    } else {
      setDateFrom(val);
      setDateError(null);
    }
  };
  const handleDateToChange = (val: string) => {
    if (dateFrom && val < dateFrom) {
      setDateError(lang === 'vi' ? 'Từ ngày phải nhỏ hơn hoặc bằng Đến ngày' : lang === 'ko' ? '시작일은 종료일보다 이전이어야 합니다' : 'From Date must be <= To Date');
    } else {
      setDateTo(val);
      setDateError(null);
    }
  };

  // FIX(upload-1-lan): đã bỏ handleFileUpload riêng của tab này — upload chỉ
  // còn 1 điểm duy nhất ở tab Nhân lực (ManpowerDashboard), dữ liệu tự share
  // sang đây qua prop `rows`.

  // ── Data ──────────────────────────────────────────────────────────────────
  const data = usePCTabData(rows, dateFrom, dateTo);
  const prodData = useProdData(rows, dateFrom, dateTo, modelFilter);

  const hasData = data.allLabels.length > 0 && Object.keys(data.ttlByLabel).length > 0;

  // EPCC (chart2-loss-cost-vi-keywords) — cờ báo Chart 2 có dữ liệu thật hay
  // không, để hiện gợi ý ngay dưới chart khi Type trong Excel không khớp bất
  // kỳ từ khóa nào (tránh lặp lại tình trạng chart trống mà không rõ lý do).
  const hasChart2Data =
    Object.keys(data.dayLossCostByModelLabel).length > 0 ||
    Object.keys(data.nightLossCostByModelLabel).length > 0 ||
    Object.keys(data.ttlLossCostByModelLabel).length > 0;

  // FIX: cùng nguyên nhân với lỗi trục Y ở Chart 3 — nếu data.ttlStandard bị
  // lệch/sai định dạng (vd quá lớn bất thường), targetPC suy ra từ nó cũng sẽ
  // bị thổi phồng theo, kéo hỏng scale của Chart 1 & Chart 2 (dùng chung targetPC).
  // Sanity check: targetPC hợp lệ chỉ nên cùng bậc độ lớn với DEFAULT_TARGET_PC
  // (đơn vị sp/người, thường vài trăm) — nếu vượt quá 5 lần mức mặc định thì coi
  // là dữ liệu nguồn có vấn đề và fallback về DEFAULT_TARGET_PC thay vì dùng số sai.
  const computedTargetPC = data.ttlStandard && data.ttlStandard > 0
    ? r1(data.ttlStandard * DAY_RATIO)
    : 0;
  const targetPC = computedTargetPC > 0 && computedTargetPC <= DEFAULT_TARGET_PC * 5
    ? computedTargetPC
    : DEFAULT_TARGET_PC;

  const labelsAll = data.allLabels.map(x => x.label);

  const filteredLabels = useMemo(() => {
    if (granularity === 'year')  return labelsAll.filter(l => /^YR/i.test(l));
    if (granularity === 'month') return labelsAll.filter(l => /^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)$/i.test(l));
    if (granularity === 'week')  return labelsAll.filter(l => /^W\d+$/i.test(l));
    // day
    return labelsAll.filter(l => /^\d{2}\/\d{2}$/.test(l));
  }, [labelsAll, granularity]);

  const displayLabels = filteredLabels.length > 0 ? filteredLabels : labelsAll;

  const modelKey = modelFilter === 'all' ? 'TTL' : modelFilter.trim().toUpperCase();
  const modelMPMap = data.byModelLabel[modelKey] ?? {};
  const dayPCMap   = data.dayPCByModelLabel[modelKey] ?? {};
  const nightPCMap = data.nightPCByModelLabel[modelKey] ?? {};
  const ttlPCMap   = data.ttlPCByModelLabel[modelKey] ?? {};

  // KPI averages using actual Per Capita values if available, fallback to headcount estimation
  const avgDay = displayLabels.length > 0
    ? r1(displayLabels.reduce((s, l) => {
        const val = dayPCMap[l];
        return s + (val != null && val > 0 ? val : (modelMPMap[l] ?? 0) * DAY_RATIO);
      }, 0) / displayLabels.length)
    : 0;

  const avgNight = displayLabels.length > 0
    ? r1(displayLabels.reduce((s, l) => {
        const val = nightPCMap[l];
        return s + (val != null && val > 0 ? val : (modelMPMap[l] ?? 0) * NIGHT_RATIO);
      }, 0) / displayLabels.length)
    : 0;

  const avgTtl = displayLabels.length > 0
    ? r1(displayLabels.reduce((s, l) => {
        const val = ttlPCMap[l];
        return s + (val != null && val > 0 ? val : (modelMPMap[l] ?? 0));
      }, 0) / displayLabels.length)
    : 0;

  // Chart refs
  const ids = useRef({
    c1: 'pctab-chart1', c2: 'pctab-chart2', c3: 'pctab-chart3',
    cDay: 'pctab-prod-day', cNight: 'pctab-prod-night', cTtl: 'pctab-prod-ttl',
    c4: 'pctab-chart4', c5: 'pctab-chart5',
  });

  // FIX: badge "● Kèm số line yêu cầu" của Chart 4 giờ dùng CHUNG 1 hàm
  // (resolveChart4Info) với phần vẽ line thật trong buildChart4, thay vì check
  // rời rạc theo toàn bộ data — tránh tình trạng badge nói "có" nhưng chart
  // không vẽ được đường (hoặc ngược lại).
  const chart4Info = useMemo(
    () => resolveChart4Info(data, displayLabels),
    [data, displayLabels]
  );

  useEffect(() => {
    if (!hasData || !plotlyReady) return;
    // Only draw when this tab is visible; charts are behind display:none otherwise
    // and Plotly will throw a "Script error" trying to measure a zero-size container.
    if (!isVisible) return;

    const draw = () => {
      // EPCC (chart123-max-15-points) — trước đây Chart 1/2/3 vẽ TOÀN BỘ
      // displayLabels (có thể lên tới hàng trăm điểm khi lọc theo ngày/tuần
      // trên khoảng thời gian rộng) khiến cột/nhãn trục X bị dồn nén, chữ đè
      // lên nhau, không đọc được (xem ảnh lỗi). Giới hạn lại tối đa 15 giai
      // đoạn GẦN NHẤT — đủ để thấy xu hướng mà vẫn rõ ràng, dễ đọc.
      const MAX_CHART_POINTS = 15;
      const labels = displayLabels.slice(-MAX_CHART_POINTS);
      // FIX(prod-last-7): 3 chart sản xuất DAY/NIGHT/TTL chỉ nên hiện tối đa
      // 7 giai đoạn GẦN NHẤT để tránh chart bị dồn quá nhiều cột khi khoảng
      // ngày lọc rộng — tách riêng khỏi `labels` dùng cho Chart 1-3 (Chart 1-3
      // vẫn hiển thị đầy đủ toàn bộ khoảng đã lọc).
      const prodLabels = labels.slice(-7);
      try {
        const ch1 = buildChart1(data, labels, textColor, gridColor, lang, targetPC, isDark, modelFilter);
        window.Plotly.react(ids.current.c1, ch1.traces, ch1.layout, { displayModeBar: false, responsive: true });

        const ch2 = buildChart2(data, labels, textColor, gridColor, lang, targetPC, isDark, modelFilter);
        window.Plotly.react(ids.current.c2, ch2.traces, ch2.layout, { displayModeBar: false, responsive: true });

        const ch3 = buildChart3(data, labels, textColor, gridColor, lang, isDark, modelFilter);
        window.Plotly.react(ids.current.c3, ch3.traces, ch3.layout, { displayModeBar: false, responsive: true });

        // Production Charts: DAY / NIGHT / TTL (chỉ vẽ khi có dữ liệu thật, tối đa 7 giai đoạn gần nhất)
        if (prodData.day.hasData) {
          const cDay = buildProductionChart(prodData.day, prodLabels, textColor, gridColor, isDark);
          window.Plotly.react(ids.current.cDay, cDay.traces, cDay.layout, { displayModeBar: false, responsive: true });
        }
        if (prodData.night.hasData) {
          const cNight = buildProductionChart(prodData.night, prodLabels, textColor, gridColor, isDark);
          window.Plotly.react(ids.current.cNight, cNight.traces, cNight.layout, { displayModeBar: false, responsive: true });
        }
        if (prodData.ttl.hasData) {
          const cTtl = buildProductionChart(prodData.ttl, prodLabels, textColor, gridColor, isDark);
          window.Plotly.react(ids.current.cTtl, cTtl.traces, cTtl.layout, { displayModeBar: false, responsive: true });
        }

        // Chart 7 + Chart 8 (cũ: Chart 4 + Chart 5)
        const ch4 = buildChart4(data, labels, textColor, gridColor, lang, isDark);
        window.Plotly.react(ids.current.c4, ch4.traces, ch4.layout, { displayModeBar: false, responsive: true });

        const ch5 = buildChart5(data, labels, textColor, gridColor, lang);
        window.Plotly.react(ids.current.c5, ch5.traces, ch5.layout, { displayModeBar: false, responsive: true });
        setChart5Legend(ch5.legend);
      } catch (err) {
        console.warn('[PerCapitaTab] Plotly render error (ignored):', err);
      }
    };

    // Use a 50ms delay when the tab just became visible so the browser has time
    // to repaint display:none → block before Plotly measures container dimensions.
    // This prevents "Script error / no stack" on lang/theme change while hidden.
    const timerId = setTimeout(draw, 50);
    return () => clearTimeout(timerId);
  }, [data, prodData, textColor, gridColor, lang, hasData, targetPC, displayLabels, isVisible, isDark, plotlyReady]);

  // ── CSS vars cho dark/light ────────────────────────────────────────────────
  // Tất cả màu dùng var(--...) để tự thích nghi theo theme; chỉ explicit color
  // cho các element active/highlight dùng màu cố định (#2e7d8c, #fff, v.v.)

  // FIX (EPCC-unify-kpi-icon-size-muc3-theo-muc2): mục 2 (TargetActualDashboard,
  // thẻ chuẩn) dùng SVG line-icon phẳng 16x16px, KHÔNG có khung tròn nền màu.
  // className="kpi-card-icon" (badge tròn ~26x26px chứa emoji) trước đây dùng ở
  // đây khiến hàng header (icon+label) cao hơn chuẩn ~8-10px, kéo theo tổng
  // chiều cao card bị lệch so với mục 2. Thay bằng SVG line-icon 16x16 giống
  // hệt mục 2 để chiều cao khớp hoàn toàn.
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
      case 'sun':
        return (
          <svg {...common}>
            <circle cx="12" cy="12" r="4" />
            <line x1="12" y1="2" x2="12" y2="4" />
            <line x1="12" y1="20" x2="12" y2="22" />
            <line x1="4.22" y1="4.22" x2="5.64" y2="5.64" />
            <line x1="18.36" y1="18.36" x2="19.78" y2="19.78" />
            <line x1="2" y1="12" x2="4" y2="12" />
            <line x1="20" y1="12" x2="22" y2="12" />
            <line x1="4.22" y1="19.78" x2="5.64" y2="18.36" />
            <line x1="18.36" y1="5.64" x2="19.78" y2="4.22" />
          </svg>
        );
      case 'moon':
        return (
          <svg {...common}>
            <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
          </svg>
        );
      case 'bar-chart':
        return (
          <svg {...common}>
            <line x1="18" y1="20" x2="18" y2="10" />
            <line x1="12" y1="20" x2="12" y2="4" />
            <line x1="6" y1="20" x2="6" y2="14" />
          </svg>
        );
      case 'target':
        return (
          <svg {...common}>
            <circle cx="12" cy="12" r="10" />
            <circle cx="12" cy="12" r="6" />
            <circle cx="12" cy="12" r="2" />
          </svg>
        );
      default:
        return null;
    }
  };

  return (
    <div className="per-capita-tab" style={{ padding: '0 0 24px' }}>

      {/* ══════════════════════════════════════════════════════════════════════
          TOOLBAR — FIX(toolbar-chuan-hoa): đồng bộ y hệt cấu trúc/kích thước/
          font-size với toolbar chuẩn tham chiếu ở TargetActualDashboard.tsx
          (ảnh 3) — layout 2 dòng (dòng nhãn + dòng control), đồng hồ dạng text
          thường, input ngày cao 38px, CustomSelect cho Model.
         ══════════════════════════════════════════════════════════════════════ */}
      <div className="topbar-dash" style={{
        display: 'flex', flexDirection: 'row', alignItems: 'stretch', gap: '0px', marginBottom: '10px',
        background: '#2F3A1D',
        borderRadius: '14px', padding: '10px 14px',
        border: '1px solid rgba(0,0,0,0.18)',
        boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
      }}>
        {/* FIX (EPCC-fit-mgmt-target-row1): 경영목표 vẫn bị rớt xuống dòng
            riêng (dòng 2) do 300px chưa đủ chứa DAY+NIGHT+경영목표 trên
            cùng 1 dòng. Nới thêm cột trái lên 340px, giảm font 12px +
            columnGap 6px để kéo 경영목표 lên nằm chung dòng 1 với DAY/NIGHT
            như đúng khoanh đỏ yêu cầu (dòng 2 giờ chỉ còn "Dữ liệu cập nhật
            đến"). */}
        <div style={{
          width: '340px', flexShrink: 0,
          display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: '4px',
          fontSize: '12px', color: '#D7E4B8', lineHeight: 1.3,
        }}>
          <div style={{ display: 'flex', flexWrap: 'wrap', columnGap: '6px', rowGap: '2px' }}>
            <span>DAY: <strong style={{ color: '#ffffff' }}>{(DAY_RATIO * 100).toFixed(1)}%</strong></span>
            <span>NIGHT: <strong style={{ color: '#ffffff' }}>{(NIGHT_RATIO * 100).toFixed(1)}%</strong></span>
            <span>
              {lang === 'vi' ? '경영목표' : lang === 'ko' ? '경영목표' : 'Mgmt Target'}:
              <strong style={{ color: '#6fd3e0', marginLeft: '4px' }}>{targetPC} {lang === 'vi' ? 'cái/người' : lang === 'ko' ? '개/인' : 'pcs/cap'}</strong>
            </span>
            {data.ttlStandard && (
              <span>
                {lang === 'vi' ? 'Chuẩn YR24' : lang === 'ko' ? '기준인원(YR24)' : 'Standard YR24'}:
                <strong style={{ color: '#ff8a8a', marginLeft: '4px' }}>{fmt1(r1(data.ttlStandard))} {lang === 'vi' ? 'người' : lang === 'ko' ? '명' : 'prs'}</strong>
              </span>
            )}
          </div>
          {data.lastDataLabel && (
            <div>
              <span>
                {lang === 'vi' ? 'Dữ liệu cập nhật đến' : lang === 'ko' ? '데이터 기준' : 'Data as of'}:
                <strong style={{ color: '#7be6ab', marginLeft: '4px' }}>{data.lastDataLabel}</strong>
              </span>
            </div>
          )}
        </div>

        {/* Cột giữa: 2 hàng label + control như cũ, không còn spacer trái */}
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '6px', minWidth: 0 }}>
        {/* Dòng 1 (labels) */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          {/* Cụm giữa dòng 1: Labels */}
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 8px' }}>
            <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>
              {lang === 'vi' ? 'NGÀY BẮT ĐẦU' : lang === 'ko' ? '시작일' : 'START DATE'}
            </span>
            <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>
              {lang === 'vi' ? 'NGÀY KẾT THÚC' : lang === 'ko' ? '종료일' : 'END DATE'}
            </span>
            <span style={{ width: '140px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>
              {lang === 'vi' ? 'MODEL' : lang === 'ko' ? '모델' : 'MODEL'}
            </span>
            <span style={{ width: '220px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>
              {lang === 'vi' ? 'XEM THEO' : lang === 'ko' ? '보기 방식' : 'VIEW BY'}
            </span>
          </div>
          {/* Cụm phải dòng 1: Spacer matching Dòng 2 (badge dùng chung) */}
          <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
            <div style={{ width: '220px' }}></div>
          </div>
        </div>

        {/* Dòng 2 (values/controls) */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          {/* Cụm giữa dòng 2: Controls */}
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 8px', alignItems: 'center' }}>
            <input
              type="date" value={dateFrom}
              onChange={e => handleDateFromChange(e.target.value)}
              className="filter-date-input"
              style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
            />
            <input
              type="date" value={dateTo}
              onChange={e => handleDateToChange(e.target.value)}
              className="filter-date-input"
              style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
            />
            <CustomSelect
              value={modelFilter}
              onChange={setModelFilter}
              options={[
                { value: 'all', label: lang === 'vi' ? 'Tất cả' : lang === 'ko' ? '전체' : 'All' },
                ...(Object.keys(data.byModelLabel).includes('TTL') ? [{ value: 'TTL', label: 'TTL' }] : []),
                ...data.activeModels.map(m => ({ value: m, label: m })),
              ]}
              style={{ width: '140px', height: '38px' }}
            />
            <div style={{ display: 'flex', gap: '0px', height: '38px', width: '220px', flexShrink: 0 }}>
              {(['day', 'week', 'month', 'year'] as const).map(mode => {
                const labels: Record<string, Record<string, string>> = {
                  day:   { vi: 'Ngày',  en: 'Day',   ko: '일별' },
                  week:  { vi: 'Tuần',  en: 'Week',  ko: '주별' },
                  month: { vi: 'Tháng', en: 'Month', ko: '월별' },
                  year:  { vi: 'Năm',   en: 'Year',  ko: '연별' },
                };
                const isActive = granularity === mode;
                return (
                  <button
                    key={mode}
                    onClick={() => setGranularity(mode)}
                    style={{
                      flex: 1,
                      padding: '6px 0', fontSize: '13px', fontWeight: 600,
                      borderRadius: mode === 'day' ? '6px 0 0 6px' : mode === 'year' ? '0 6px 6px 0' : '0',
                      // FIX (EPCC-vanilla-toolbar-contrast): màu cố định,
                      // không phụ thuộc theme — tránh chữ mờ/biến mất khi
                      // đổi theme trên nền Vanilla.
                      border: '1px solid rgba(0,0,0,0.18)',
                      borderRight: mode !== 'year' ? 'none' : '1px solid rgba(0,0,0,0.18)',
                      background: isActive ? '#2e7d8c' : 'rgba(255,255,255,0.55)',
                      color: isActive ? '#ffffff' : '#7A5A2E',
                      cursor: 'pointer',
                      transition: 'all 0.15s ease',
                      height: '100%',
                      display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {labels[mode][lang]}
                  </button>
                );
              })}
            </div>
            {dateError && (
              <span style={{ color: 'var(--rose)', fontSize: '11px', fontWeight: 500, whiteSpace: 'nowrap' }}>
                ⚠️ {dateError}
              </span>
            )}
          </div>

          {/* Cụm phải dòng 2: badge "dữ liệu dùng chung" (FIX upload-1-lan) */}
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexShrink: 0 }}>
            <span
              title={lang === 'vi' ? 'Dữ liệu được nạp chung từ tab Nhân lực — chỉ cần tải Excel 1 lần' : lang === 'ko' ? '근무 인력 탭과 데이터를 공유합니다 — 엑셀은 한 번만 업로드하면 됩니다' : 'Data is shared from the Manpower tab — upload Excel only once'}
              style={{
                display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px',
                height: '38px', width: '220px', padding: '0 12px', boxSizing: 'border-box',
                fontSize: '12px', fontWeight: 600, color: 'var(--text-2)',
                background: 'var(--surface-2)', border: '1px dashed var(--border)',
                borderRadius: '8px', whiteSpace: 'nowrap',
              }}
            >
              🔗 {lang === 'vi' ? 'Dữ liệu chung với tab Nhân lực' : lang === 'ko' ? '근무 인력 탭과 데이터 공유' : 'Shared with Manpower tab'}
            </span>
          </div>
        </div>
        </div>
      </div>

      {/* ── KPI Row ── */}
      {hasData && (
        <div
          className="kpi-grid"
          style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '12px', marginBottom: '12px', width: '100%' }}
        >
          {[
            {
              icon: 'sun',
              label: lang === 'vi' ? 'Sản lượng đầu người ca ngày TB' : lang === 'ko' ? 'DAY 인당생산수 평균' : 'AVG DAY Per Capita',
              value: fmt1(avgDay),
              unit: lang === 'vi' ? 'sp/người' : lang === 'ko' ? '개/인' : 'pcs/cap',
              bg: 'var(--blue-soft)', color: '#1565C0',
              // FIX (EPCC-copy-kpi-colors-from-muc2): copy đúng bộ màu KPI card
              // từ mục 2 (TargetActualDashboard, tab "Tình hình doanh số") sang
              // đây, đúng thứ tự trái→phải: teal #2e7d8c (card 1) → green #10b981
              // (card 2) → purple #8b5cf6 (card 3) → orange #f59e0b (card 4).
              // rgbaBg dùng đúng định dạng rgba(...,0.1) như mục 2 (không dùng
              // suffix hex-alpha để tránh sai khác/tương thích trình duyệt).
              borderColor: '#2e7d8c', rgbaBg: 'rgba(46,125,140,0.1)',
            },
            {
              icon: 'moon',
              label: lang === 'vi' ? 'Sản lượng đầu người ca đêm TB' : lang === 'ko' ? 'NIGHT 인당생산수 평균' : 'AVG NIGHT Per Capita',
              value: fmt1(avgNight),
              unit: lang === 'vi' ? 'sp/người' : lang === 'ko' ? '개/인' : 'pcs/cap',
              bg: 'var(--rose-soft)', color: '#ef4444',
              borderColor: '#10b981', rgbaBg: 'rgba(16,185,129,0.1)',
            },
            {
              icon: 'bar-chart',
              label: lang === 'vi' ? 'Sản lượng đầu người tổng TB' : lang === 'ko' ? 'TTL 인당생산수 평균' : 'AVG TTL Per Capita',
              value: fmt1(avgTtl),
              unit: lang === 'vi' ? 'sp/người' : lang === 'ko' ? '개/인' : 'pcs/cap',
              bg: 'var(--cyan-soft)', color: '#14b8a6',
              borderColor: '#8b5cf6', rgbaBg: 'rgba(139,92,246,0.1)',
            },
            {
              icon: 'target',
              label: lang === 'vi' ? 'Mục tiêu' : lang === 'ko' ? '경영목표' : 'Mgmt. Target',
              value: String(targetPC),
              unit: lang === 'vi' ? 'sp/người' : lang === 'ko' ? '개/인' : 'pcs/cap',
              bg: 'var(--rose-soft)', color: '#f43f5e',
              borderColor: '#f59e0b', rgbaBg: 'rgba(245,158,11,0.1)',
            },
          ].map((kpi, i) => (
            <div
              key={i}
              className="kpi-card"
              style={{
                borderLeft: `4px solid ${kpi.borderColor}`,
                background: `linear-gradient(135deg, ${kpi.rgbaBg} 0%, rgba(30,41,59,0.4) 100%)`,
              }}
            >
              {/* FIX (EPCC-unify-kpi-typography-all-sections): copy đúng cỡ
              chữ/spacing tường minh từ mục 2 (TargetActualDashboard —
              .kpi-card-header: gap 6px + margin-bottom 8px; .kpi-card-label:
              13.5px/700/var(--text-0); .kpi-card-value: weight 800/
              var(--text-0); dòng đơn vị kiểu .kpi-card-target: 13.5px/700/
              var(--text-1)) sang toàn bộ 4 card ở đây để 3 mục có tỷ lệ
              dài/rộng/cao và cỡ chữ trên-dưới giống hệt nhau.
              FIX (EPCC-remove-kpi-value-fontsize-override): mục 2 KHÔNG tự set
              fontSize riêng cho .kpi-card-value (chỉ set marginBottom:0), số to
              (867/30.1%/...) hoàn toàn dùng cỡ chữ mặc định của class. Trước đây
              ở đây tự đặt fontSize:'26px' khiến số to hơn chuẩn mục 2. Bỏ hẳn
              fontSize override để value dùng đúng cỡ chữ mặc định giống hệt mục 2. */}
          <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
                {renderKpiIcon(kpi.icon, kpi.color)}
                <div className="kpi-card-label" style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-0)' }}>{kpi.label}</div>
              </div>
              <div className="kpi-card-value" style={{ fontWeight: 800, color: 'var(--text-0)' }}>
                {kpi.value}
                <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-1)', marginLeft: '4px' }}>{kpi.unit}</span>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* ── Charts ── */}
      {!hasData ? (
        <div style={{ padding: '60px', textAlign: 'center', color: 'var(--text-2)' }}>
          <div style={{ fontSize: '40px' }}>🏭</div>
          <p style={{ marginTop: '12px' }}>
            {lang === 'vi' ? 'Không có dữ liệu nhân lực' : lang === 'ko' ? '인력 데이터 없음' : 'No manpower data'}
          </p>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>

          {/* Chart 1: 인당생산수 */}
          <div className="panel">
            <div className="panel-head" style={{ ...panelHeadStyle(0, isDark), display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '10px', flexWrap: 'wrap' }}>
              <h3 style={{ margin: 0 }}>{lang === 'vi' ? 'Sản lượng theo đầu người' : lang === 'ko' ? '인당생산수' : 'Per Capita Output'}</h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <PCLegendItem type="bar" color="#1565C0" label={t('dayPc', lang)} />
                <PCLegendItem type="bar" color="#ef4444" label={t('nightPc', lang)} />
                <PCLegendItem type="line" color={isDark ? '#14b8a6' : '#0f766e'} label={t('ttlPc', lang)} />
                <PCLegendItem type="line" color="#f59e0b" label={t('target', lang)} />
                <PCLegendItem type="line" color={isDark ? '#f43f5e' : '#be123c'} label={t('mgmtTarget', lang)} />
              </div>
            </div>
            <div className="chart-holder">
              <div id={ids.current.c1} style={{ height: '270px' }} />
            </div>
          </div>

          {/* ════ Chart: Sản xuất theo ca DAY / NIGHT / TTL ════════════════════════ */}
          <div className="panel">
            <div className="panel-head" style={{ ...panelHeadStyle(1, isDark), marginBottom: '12px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '10px', flexWrap: 'wrap' }}>
              <h3 style={{ margin: 0 }}>
                {lang === 'vi'
                  ? `📊 Tình hình sản xuất — Plan / Actual / Yield (%)${modelFilter !== 'all' ? ` — Model: ${modelFilter}` : ' — Tất cả Model'}`
                  : lang === 'ko'
                  ? `📊 생산 현황 — 계획 / 실적 / 달성률 (%)${modelFilter !== 'all' ? ` — 모델: ${modelFilter}` : ' — 전체 모델'}`
                  : `📊 Production Status — Plan / Actual / Yield (%)${modelFilter !== 'all' ? ` — Model: ${modelFilter}` : ' — All Models'}`}
              </h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <PCLegendItem type="bar" color={isDark ? '#f97316' : '#ea580c'} label="Plan" />
                <PCLegendItem type="bar" color="#14b8a6" label="Actual" />
                <PCLegendItem type="line" color="#1565C0" label="Yield %" />
              </div>
              {!prodData.day.hasData && !prodData.night.hasData && !prodData.ttl.hasData && (
                <span style={{
                  display: 'inline-flex', alignItems: 'center', gap: '6px',
                  fontSize: '12px', color: 'var(--text-3)', fontStyle: 'italic',
                  padding: '5px 12px', borderRadius: '6px', marginTop: '8px',
                  background: 'var(--surface-2)', border: '1px dashed var(--border)',
                }}>
                  ⚠️ {lang === 'vi'
                    ? 'Chưa có dữ liệu sản xuất DAY/NIGHT/TTL trong file. Thêm dòng có Type chứa "DAY PRON\'D PLAN", "DAY Pro Actual", "NIGHT PRON\'D PLAN", "TTL Pro Actual"... vào Excel rồi tải lại.'
                    : lang === 'ko'
                    ? 'DAY/NIGHT/TTL 생산 데이터 없음. Type에 "DAY PRON\'D PLAN", "DAY Pro Actual", "TTL Pro Actual" 등 포함된 행을 Excel에 추가 후 다시 업로드하세요.'
                    : 'No DAY/NIGHT/TTL production data found. Add rows with Type containing "DAY PRON\'D PLAN", "DAY Pro Actual", "TTL Pro Actual", etc. to Excel, then re-upload.'}
                </span>
              )}
            </div>
            <div className="chart-holder" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '12px' }}>

              {/* Chart: DAY */}
              <div>
                <div style={{
                  fontSize: '12px', fontWeight: 700, color: textColor,
                  textAlign: 'center', marginBottom: '4px',
                  textTransform: 'uppercase', letterSpacing: '0.06em',
                }}>
                  DAY
                </div>
                {prodData.day.hasData ? (
                  <div id={ids.current.cDay} style={{ height: '220px' }} />
                ) : (
                  <div style={{
                    height: '220px', display: 'flex', alignItems: 'center',
                    justifyContent: 'center', color: 'var(--text-3)', fontSize: '12px',
                    border: '1px dashed var(--border)', borderRadius: '8px',
                  }}>
                    {lang === 'vi' ? 'Chưa có dữ liệu' : lang === 'ko' ? '데이터 없음' : 'No data'}
                  </div>
                )}
              </div>

              {/* Chart: NIGHT */}
              <div>
                <div style={{
                  fontSize: '12px', fontWeight: 700, color: textColor,
                  textAlign: 'center', marginBottom: '4px',
                  textTransform: 'uppercase', letterSpacing: '0.06em',
                }}>
                  NIGHT
                </div>
                {prodData.night.hasData ? (
                  <div id={ids.current.cNight} style={{ height: '220px' }} />
                ) : (
                  <div style={{
                    height: '220px', display: 'flex', alignItems: 'center',
                    justifyContent: 'center', color: 'var(--text-3)', fontSize: '12px',
                    border: '1px dashed var(--border)', borderRadius: '8px',
                  }}>
                    {lang === 'vi' ? 'Chưa có dữ liệu' : lang === 'ko' ? '데이터 없음' : 'No data'}
                  </div>
                )}
              </div>

              {/* Chart: TTL */}
              <div>
                <div style={{
                  fontSize: '12px', fontWeight: 700, color: textColor,
                  textAlign: 'center', marginBottom: '4px',
                  textTransform: 'uppercase', letterSpacing: '0.06em',
                }}>
                  TTL
                </div>
                {prodData.ttl.hasData ? (
                  <div id={ids.current.cTtl} style={{ height: '220px' }} />
                ) : (
                  <div style={{
                    height: '220px', display: 'flex', alignItems: 'center',
                    justifyContent: 'center', color: 'var(--text-3)', fontSize: '12px',
                    border: '1px dashed var(--border)', borderRadius: '8px',
                  }}>
                    {lang === 'vi' ? 'Chưa có dữ liệu' : lang === 'ko' ? '데이터 없음' : 'No data'}
                  </div>
                )}
              </div>

            </div>
          </div>

          {/* Chart 7 + Chart 8 — 2 cột (cũ: Chart 4 + Chart 5) */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>

            {/* Chart 4: Phân bố nhân lực theo Model — Giai đoạn gần nhất */}
            <div className="panel">
              <div className="panel-head" style={{ ...panelHeadStyle(2, isDark), display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <h3 style={{ margin: 0 }}>
                  {lang === 'vi'
                    ? 'Phân bố nhân lực theo Model — Giai đoạn gần nhất'
                    : lang === 'ko'
                    ? '모델별 인원 분포 — 최근 기간'
                    : 'Headcount by Model — Latest Period'}
                </h3>
                {chart4Info.hasLineData ? (
                  <span style={{ fontSize: '12px', color: isDark ? '#facc15' : '#b45309', display: 'inline-flex', alignItems: 'center' }}>
                    {lang === 'vi' ? "● Kèm số line yêu cầu (TTL LINE Q'TY)" : lang === 'ko' ? "● 요구 라인 수 포함 (TTL LINE Q'TY)" : "● Incl. required lines (TTL LINE Q'TY)"}
                    {chart4Info.lineLabel && chart4Info.lineLabel !== chart4Info.latestLabel
                      ? ` [${chart4Info.lineLabel}]`
                      : ''}
                  </span>
                ) : (
                  <span style={{ fontSize: '12px', color: 'var(--text-3)', display: 'inline-flex', alignItems: 'center', fontStyle: 'italic' }}>
                    {lang === 'vi'
                      ? "○ Chưa có dữ liệu Line Q'TY — thêm dòng Type \"…LINE Q'TY\" trong Excel để hiển thị đường"
                      : lang === 'ko'
                      ? "○ Line Q'TY 데이터 없음 — Excel에 Type \"…LINE Q'TY\" 행을 추가하면 라인이 표시됩니다"
                      : "○ No Line Q'TY data yet — add a Type \"…LINE Q'TY\" row in Excel to show the line"}
                  </span>
                )}
              </div>
              <div className="chart-holder">
                <div id={ids.current.c4} style={{ minHeight: '450px' }} />
              </div>
            </div>

            {/* Chart 5: Phân bố Model theo Giai đoạn (Spider) */}
            <div className="panel">
              <div className="panel-head" style={{ ...panelHeadStyle(3, isDark), display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <h3 style={{ margin: 0 }}>
                  {lang === 'vi'
                    ? 'Phân bố Model theo Giai đoạn (Spider)'
                    : lang === 'ko'
                    ? '기간별 모델 분포 (Spider)'
                    : 'Model Distribution by Period (Spider)'}
                </h3>
                <span style={{ fontSize: '11px', color: 'var(--text-3)', display: 'inline-flex', alignItems: 'center' }}>
                  {lang === 'vi' ? 'Tối đa 10 giai đoạn gần nhất' : lang === 'ko' ? '최근 10기간' : 'Up to last 10 periods'}
                </span>
                {/* EPCC (percapita-radar-unify-rtytotal) - THÊM MỚI: legend
                    động ngoài Plotly (đúng "cách vẽ" plotSpiderPanel() của
                    RtyTotalTab.tsx — showlegend:false trong layout, legend
                    thật hiển thị ở header panel bằng component riêng, ở đây
                    tái sử dụng PCLegendItem sẵn có thay vì tạo mới). */}
                {chart5Legend.length > 0 && (
                  <div style={{ display: 'flex', gap: '10px', rowGap: '4px', flexWrap: 'wrap', alignItems: 'center' }}>
                    {chart5Legend.map((lg, i) => (
                      <PCLegendItem key={i} type="line" color={lg.color} label={lg.label} />
                    ))}
                  </div>
                )}
              </div>
              <div className="chart-holder">
                <div id={ids.current.c5} style={{ minHeight: '450px' }} />
              </div>
            </div>

          </div>

          {/* Chart 2: 인력 유실 비용 (Labor Loss Cost) — EPCC (chart2-loss-cost) */}
          <div className="panel">
            <div className="panel-head" style={{ ...panelHeadStyle(4, isDark), display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '10px', flexWrap: 'wrap' }}>
              <h3 style={{ margin: 0 }}>{t('chart2Title', lang)}</h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <PCLegendItem type="bar" color="#1565C0" label={t('dayLossCost', lang)} />
                <PCLegendItem type="bar" color="#ef4444" label={t('nightLossCost', lang)} />
                <PCLegendItem type="line" color={isDark ? '#14b8a6' : '#0f766e'} label={t('ttlLossCost', lang)} />
              </div>
              {!hasChart2Data && (
                <span style={{
                  display: 'inline-flex', alignItems: 'center', gap: '6px',
                  fontSize: '12px', color: 'var(--text-3)', fontStyle: 'italic',
                  padding: '5px 12px', borderRadius: '6px',
                  background: 'var(--surface-2)', border: '1px dashed var(--border)',
                }}>
                  ⚠️ {lang === 'vi'
                    ? 'Chưa nhận diện được dữ liệu Chi phí hao hụt nhân lực. Kiểm tra cột Type trong Excel có chứa "DAY/NIGHT/TTL 인력 유실 비용" hoặc "hao hụt nhân lực" không, rồi tải lại.'
                    : lang === 'ko'
                    ? '인력 유실 비용 데이터를 찾을 수 없습니다. Excel의 Type 열에 "DAY/NIGHT/TTL 인력 유실 비용"이 포함되어 있는지 확인 후 다시 업로드하세요.'
                    : 'No labor loss cost data recognized. Check that the Type column in Excel contains "DAY/NIGHT/TTL Labor Loss Cost", then re-upload.'}
                </span>
              )}
            </div>
            <div className="chart-holder">
              <div id={ids.current.c2} style={{ height: '270px' }} />
            </div>
          </div>

          {/* Chart 3: 근무 인력 현황 */}
          <div className="panel">
            <div className="panel-head" style={{ ...panelHeadStyle(5, isDark), display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '10px', flexWrap: 'wrap' }}>
              <h3 style={{ margin: 0 }}>{t('chart3Title', lang)}</h3>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <PCLegendItem type="bar" color="#1565C0" label={t('dayMP', lang)} />
                <PCLegendItem type="bar" color="#ef4444" label={t('nightMP', lang)} />
                <PCLegendItem type="line" color={isDark ? '#14b8a6' : '#0f766e'} label={t('ttlMP', lang)} />
                {!!data.ttlStandard && (
                  <PCLegendItem type="line" color={isDark ? '#f43f5e' : '#be123c'} label={t('standard', lang)} />
                )}
              </div>
              {data.ttlStandard && (
                <span style={{
                  fontSize: '13px', padding: '2px 8px', borderRadius: '4px',
                  background: 'var(--rose-soft)', color: 'var(--rose)',
                  border: '1px dashed var(--rose)',
                  display: 'inline-flex', alignItems: 'center',
                }}>
                  {lang === 'vi' ? 'TTL 기준인력 Standard AVG' : lang === 'ko' ? 'TTL 기준인력 Standard AVG' : 'TTL Standard AVG'}: {fmt1(r1(data.ttlStandard))} {lang === 'vi' ? 'người' : lang === 'ko' ? '명' : 'prs'}
                </span>
              )}
            </div>
            <div className="chart-holder">
              <div id={ids.current.c3} style={{ height: '270px' }} />
            </div>
          </div>

        </div>
      )}

      {/* EPCC (panel-flush-header-match-salesdashboard) — .panel là class CSS
          global dùng chung nhiều Mục khác nên KHÔNG sửa trực tiếp; chỉ ghi đè
          CỤC BỘ trong phạm vi .per-capita-tab bằng specificity cao (lặp lại
          class 3 lần), đúng pattern đã áp dụng ở SalesDashboard.tsx, để thanh
          màu panel-head dán sát 3 cạnh trên/trái/phải khung .panel (đúng như
          4 mũi tên đỏ khoanh góc trong ảnh mẫu). Phần padding bỏ khỏi .panel
          được dồn sang .chart-holder để nội dung biểu đồ bên trong không bị
          dính sát mép. */}
      <style>{`
        .per-capita-tab.per-capita-tab.per-capita-tab .panel {
          padding: 0 !important;
          overflow: hidden;
        }
        .per-capita-tab.per-capita-tab.per-capita-tab .panel-head {
          margin: 0 !important;
        }
        .per-capita-tab.per-capita-tab.per-capita-tab .chart-holder {
          box-sizing: border-box;
          padding: 14px 16px 16px 16px !important;
        }
      `}</style>
    </div>
  );
};

export default PerCapitaTab;

