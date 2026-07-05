import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import type { DataRow } from '../types';
import { parseToDate } from '../utils';
import { PerCapitaTab } from './PerCapitaTab';
import { GlobalHeaderControls } from './GlobalHeaderControls';
import { CustomSelect } from './CustomSelect';

interface ManpowerDashboardProps {
  rows: DataRow[];
  theme: 'light' | 'dark';
  lang: 'vi' | 'en' | 'ko';
  onToggleTheme: () => void;
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
  onFileSelected: (file: File, wb: any) => void;
}

// ─── Constants ────────────────────────────────────────────────────────────────
const TTL_MODEL = 'TTL';

const MODEL_COLORS: Record<string, string> = {
  SO1B01:  '#3b82f6',
  SO1C2EF: '#10b981',
  SO1C2G:  '#f59e0b',
  SO3560:  '#ef4444',
  TTL:     '#14b8a6',
};

// ─── i18n ─────────────────────────────────────────────────────────────────────
const T = {
  title:      { vi: '근무 인력 현황 (Manpower Dashboard)', en: 'Manpower Status Dashboard', ko: '근무 인력 현황' },
  subtitle:   { vi: 'Phân tích nhân lực theo Model & Thời gian', en: 'Manpower by Model & Period', ko: '모델별 인력 현황 분석' },
  kpiTtlAvg: { vi: 'Tổng nhân lực TB', en: 'Total Manpower Avg', ko: '총 인원 평균' },
  kpiPeak:   { vi: 'Giai đoạn cao điểm', en: 'Peak Period', ko: '최고 기간' },
  kpiModels: { vi: 'Models đang hoạt động', en: 'Active Models', ko: '가동 모델 수' },
  kpiLatest: { vi: 'Thời gian mới nhất', en: 'Latest Period', ko: '최근 기간' },
  // FIX(chart2-upph): chart 2 trước đây là "Nhân lực TB theo Tuần" — trùng ý nghĩa
  // với chart 1 (đều là stacked bar + TTL line theo Model). Đổi chart 2 thành
  // biểu đồ UPPH (Target/Actual/Đạt tỷ lệ theo DAY/NIGHT/TTL) để bổ sung góc nhìn
  // năng suất, không lặp lại chart 1.
  chartWeek: { vi: 'UPPH theo Ca (Target vs Actual & Đạt tỷ lệ)', en: 'UPPH by Shift (Target vs Actual & Achievement Rate)', ko: '교대별 UPPH (목표 대비 실적 & 달성율)' },
  upphSubtitle: { vi: 'Tối đa 7 giai đoạn gần nhất', en: 'Last 7 periods max', ko: '최근 7개 구간까지' },
  chartModel:{ vi: 'Phân bổ nhân lực theo Model — Giai đoạn gần nhất', en: 'Manpower by Model — Latest Period', ko: '모델별 인원 현황 (최근 기간)' },
  chartDay:  { vi: 'Nhân lực hàng ngày (TTL)', en: 'Daily Manpower (TTL)', ko: '일별 인원 현황 (TTL)' },
  chartRadar:{ vi: 'Phân bổ Model theo Giai đoạn (Spider)', en: 'Model Distribution by Period (Radar)', ko: '모델 인원 분포 (레이더)' },
  noData:    { vi: 'Không có dữ liệu nhân lực', en: 'No manpower data found', ko: '인력 데이터가 없습니다' },
  standard:  { vi: 'Tiêu chuẩn YR24', en: 'YR24 Standard', ko: 'YR24 기준' },
  persons:   { vi: 'người', en: 'prs', ko: '명' },
  
  // New toolbar labels
  from: { vi: 'Từ', en: 'From', ko: '시작일' },
  to: { vi: 'Đến', en: 'To', ko: '종료일' },
  modelFilter: { vi: 'Model', en: 'Model', ko: '모델' },
  allModels: { vi: 'Tất cả Model', en: 'All Models', ko: '전체 모델' },
  uploadBtn: { vi: 'Tải tệp', en: 'Upload File', ko: '파일 업로드' },
  exportBtn: { vi: 'Tải Excel', en: 'Export Excel', ko: '엑셀 다운로드' },
  dateError: { vi: 'Từ ngày phải nhỏ hơn hoặc bằng Đến ngày', en: 'From Date must be <= To Date', ko: '시작일은 종료일보다 이전이어야 합니다' },
  viewBy: { vi: 'XEM THEO', en: 'VIEW BY', ko: '보기 방식' },
  day: { vi: 'Ngày', en: 'Day', ko: '일별' },
  week: { vi: 'Tuần', en: 'Week', ko: '주별' },
  month: { vi: 'Tháng', en: 'Month', ko: '월별' },
  year: { vi: 'Năm', en: 'Year', ko: '연별' },

  // UPPH chart labels (chart 2)
  upphUnit:    { vi: 'UPPH', en: 'UPPH', ko: 'UPPH' },
  upphRate:    { vi: 'Đạt tỷ lệ (%)', en: 'Achievement Rate (%)', ko: '달성율 (%)' },
  upphTarget:  { vi: 'Target', en: 'Target', ko: 'Target' },
  upphActual:  { vi: 'Actual', en: 'Actual', ko: 'Actual' },
  shiftDay:    { vi: 'DAY', en: 'DAY', ko: 'DAY' },
  shiftNight:  { vi: 'NIGHT', en: 'NIGHT', ko: 'NIGHT' },
  shiftTtl:    { vi: 'TTL', en: 'TTL', ko: 'TTL' },
  noUpphData:  { vi: 'Chưa có dữ liệu UPPH (DAY/NIGHT/TTL Target-Actual)', en: 'No UPPH data (DAY/NIGHT/TTL Target-Actual)', ko: 'UPPH 데이터 없음 (DAY/NIGHT/TTL Target-Actual)' },
};

function t(key: keyof typeof T, lang: 'vi'|'en'|'ko'): string {
  return (T[key] as any)[lang] ?? (T[key] as any)['vi'];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function round1(n: number | null | undefined): number {
  if (n == null || isNaN(n as number)) return 0;
  return Math.round((n as number) * 10) / 10;
}

function fmt1(n: number): string {
  return n.toLocaleString('vi-VN', { maximumFractionDigits: 1 });
}

export function parseManpowerDate(raw: any): Date | null {
  if (!raw) return null;
  if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;

  const s = String(raw).trim().toUpperCase();

  // 1. Handle Month Strings like JAN, FEB, MAR, APR, MAY, JUN...
  const monthsAbbr = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const monthIdx = monthsAbbr.indexOf(s);
  if (monthIdx !== -1) {
    return new Date(2026, monthIdx, 1);
  }

  // 2. Handle Week Strings like W22, W26, etc.
  const weekMatch = s.match(/^W(\d{1,2})$/);
  if (weekMatch) {
    const weekNum = parseInt(weekMatch[1], 10);
    // Base date for ISO week N in 2026: W01 starts on Mon, Dec 29, 2025
    const base = new Date(Date.UTC(2025, 11, 29));
    const time = base.getTime() + (weekNum - 1) * 7 * 24 * 60 * 60 * 1000;
    return new Date(time);
  }

  // 3. Handle Day Strings like 06/22, 06/26, etc.
  const dayMatch = s.match(/^(\d{2})[/-](\d{2})$/);
  if (dayMatch) {
    const month = parseInt(dayMatch[1], 10) - 1;
    const day = parseInt(dayMatch[2], 10);
    return new Date(2026, month, day);
  }

  // 4. Default fallback to parseToDate
  return parseToDate(raw);
}

function getISOWeek(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

function getModelColor(model: string, index: number): string {
  const upper = model.toUpperCase();
  if (MODEL_COLORS[upper]) return MODEL_COLORS[upper];
  const palette = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#14b8a6', '#f43f5e'];
  return palette[index % palette.length];
}

// ─── Shared date/period helpers (dùng chung cho useManpowerData & useUpphData) ─
// FIX(chart2-upph): trước đây các hàm này được khai báo LOCAL bên trong
// useManpowerData nên useUpphData (hook mới cho chart 2) không tái sử dụng được,
// phải chép lại code → dễ lệch logic phân loại tầng dữ liệu (day/week/month).
// Đưa lên module scope để 2 hook luôn phân loại period giống hệt nhau.
const MONTH_ABBRS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
type RowDateType = 'month' | 'week' | 'day' | 'other';

function classifyRowDateType(dateVal: any): RowDateType {
  // Excel date cell thật (cellDates:true) → coi như dòng daily
  if (dateVal instanceof Date) return 'day';
  const rawStr = String(dateVal).trim().toUpperCase();
  if (MONTH_ABBRS.includes(rawStr)) return 'month';
  if (/^W\d{1,2}$/.test(rawStr)) return 'week';
  if (/^\d{1,2}[/-]\d{1,2}$/.test(rawStr)) return 'day';
  return 'other';
}

function formatDayLabel(d: Date): string {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  return `${mm}/${dd}`;
}

function getPeriodLabel(d: Date, granularity: 'day' | 'week' | 'month' | 'year'): string {
  switch (granularity) {
    case 'day': return formatDayLabel(d);
    case 'week': return 'W' + getISOWeek(d);
    case 'month': return MONTH_ABBRS[d.getMonth()];
    case 'year': return String(d.getFullYear());
    default: return MONTH_ABBRS[d.getMonth()];
  }
}

// ─── Data hook ────────────────────────────────────────────────────────────────
interface ManpowerData {
  byModelPeriod: Record<string, Record<string, number | null>>;  // model → period → avg
  ttlStandard: number | null;
  activeModels: string[];
  labels: string[];
  activeLabels: string[];
}

function useManpowerData(
  rows: DataRow[],
  dateFrom: string,
  dateTo: string,
  granularity: 'day' | 'week' | 'month' | 'year'
): ManpowerData {
  return useMemo(() => {
    // FIX: Test_3.xlsx có 6 biến thể Type song song cho mỗi model:
    //   DAY ManPower / DAY ManPower AVG / NIGHT ManPower / NIGHT ManPower AVG / TTL ManPower / TTL ManPower AVG
    // Filter cũ `.includes('manpower')` khớp CẢ 6 loại → bị gom chung vào 1 mảng
    // rồi lấy trung bình (xem dưới), trộn lẫn DAY+NIGHT+TTL và raw+AVG với nhau
    // → ra số hoàn toàn sai cho mỗi model. Dashboard này hiển thị tổng nhân lực
    // theo model (không tách ca ngày/đêm) nên chỉ lấy đúng "TTL ManPower AVG".
    const mpRows = rows.filter(r => {
      const typeStr = String(r.type || r.Type || '').trim().toLowerCase();
      return typeStr === 'ttl manpower avg';
    });

    if (mpRows.length === 0) {
      // FIX: thiếu field `activeLabels` ở nhánh early-return này gây lỗi
      // "Cannot read properties of undefined (reading 'length')" tại mọi nơi
      // dùng data.activeLabels.length (KPI cards, charts...).
      return { byModelPeriod: {}, ttlStandard: null, activeModels: [], labels: [], activeLabels: [] };
    }

    // Retrieve Standard YR24 (independent of filters)
    const stdRow = mpRows.find(r => {
      const m = String(r.model || r.Model || '').trim().toUpperCase();
      const d = String(r.date || r.Date || '').trim().toUpperCase();
      return m === 'TTL' && d === 'YR24';
    });
    const ttlStandard = stdRow && stdRow.value != null && !isNaN(Number(stdRow.value)) ? Number(stdRow.value) : null;

    // ── FIX: phân loại dòng theo "tầng" dữ liệu gốc (month/week/day) ───────
    // Test_3.xlsx có 3 tầng dữ liệu cùng tồn tại trong 1 sheet:
    //   - Monthly summary rows : Date = "JAN".."DEC"  (đã pre-aggregated)
    //   - Weekly summary rows  : Date = "W01".."W26"  (đã pre-aggregated)
    //   - Daily raw rows       : Date = "01/02", "06/26"... (raw, chưa aggregate)
    // Trước đây mọi loại dòng đều bị quy đổi qua getPeriodLabel() về CÙNG 1
    // label khi granularity=month (vd cả "JUN" và toàn bộ "06/xx" đều ra "JUN")
    // → 2 tầng dữ liệu bị trộn + average sai, đồng thời latestPeriod có thể bị
    // lệch sang 1 label rỗng theo model → Model Bar & Spider render trống.
    // Cách sửa: chỉ giữ đúng tầng dữ liệu khớp granularity đang chọn.
    // (MONTH_ABBRS / classifyRowDateType nay dùng bản module-scope ở trên)

    // Parse dates and validate
    type ParsedManpowerRow = DataRow & { _parsedDate: Date | null; _dateType: RowDateType };
    const parsedRows = (mpRows.map(r => {
      const dateVal = (r as any).date || (r as any).Date || (r as any).month || (r as any).Month || '';
      const parsedDate = parseManpowerDate(dateVal);
      const dateType = classifyRowDateType(dateVal);
      return { ...r, _parsedDate: parsedDate, _dateType: dateType };
    }) as ParsedManpowerRow[]).filter(r => {
      const dateVal = (r as any).date || (r as any).Date || (r as any).month || (r as any).Month || '';
      return r._parsedDate !== null && String(dateVal).trim().toUpperCase() !== 'YR24';
    }) as (DataRow & { _parsedDate: Date; _dateType: RowDateType })[];

    // Chỉ giữ tầng dữ liệu khớp granularity hiện tại (year → dùng tạm tầng month)
    const wantedType: RowDateType = granularity === 'year' ? 'month' : granularity;
    let typedRows = parsedRows.filter(r => r._dateType === wantedType);
    // Fallback an toàn: nếu tầng mong muốn chưa có dòng nào (vd chưa nhập daily)
    // thì dùng tạm toàn bộ để dashboard không bị trống hẳn.
    if (typedRows.length === 0) typedRows = parsedRows;

    // Apply date range filters
    let filtered = typedRows;
    const fromDate = dateFrom ? new Date(dateFrom) : null;
    const toDate = dateTo ? new Date(dateTo) : null;
    if (toDate) {
      toDate.setHours(23, 59, 59, 999);
    }
    if (fromDate) {
      filtered = filtered.filter(r => r._parsedDate >= fromDate);
    }
    if (toDate) {
      filtered = filtered.filter(r => r._parsedDate <= toDate);
    }

    // Dynamic Active Models: only models with at least one value > 0 in current filter
    const activeModels = [...new Set(
      filtered
        .filter(r => { const v = Number(r.value || r.Value); return !isNaN(v) && v > 0; })
        .map(r => String(r.model || r.Model || '').trim())
        .filter(m => m && m.toUpperCase() !== 'TTL')
    )].sort();

    // Extract sorted unique labels chronologically
    // (formatDayLabel / getPeriodLabel nay dùng bản module-scope ở trên)
    const labelToMinDate: Record<string, number> = {};
    filtered.forEach(r => {
      const label = getPeriodLabel(r._parsedDate, granularity);
      const time = r._parsedDate.getTime();
      if (!labelToMinDate[label] || time < labelToMinDate[label]) {
        labelToMinDate[label] = time;
      }
    });
    const labels = Object.keys(labelToMinDate).sort((a, b) => labelToMinDate[a] - labelToMinDate[b]);

    // Group and average values
    const groups: Record<string, Record<string, number[]>> = {};
    filtered.forEach(r => {
      const model = String(r.model || r.Model || '').trim().toUpperCase() || 'UNKNOWN';
      const label = getPeriodLabel(r._parsedDate, granularity);
      const val = Number(r.value || r.Value);
      if (isNaN(val)) return;

      if (!groups[model]) groups[model] = {};
      if (!groups[model][label]) groups[model][label] = [];
      groups[model][label].push(val);
    });

    const byModelPeriod: Record<string, Record<string, number | null>> = {};
    Object.keys(groups).forEach(model => {
      byModelPeriod[model] = {};
      Object.keys(groups[model]).forEach(label => {
        const vals = groups[model][label];
        const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
        byModelPeriod[model][label] = avg;
      });
    });

    // Nếu không có model TTL trong dữ liệu, tính tổng động từ các model khác
    if (!byModelPeriod[TTL_MODEL]) {
      byModelPeriod[TTL_MODEL] = {};
      labels.forEach(l => {
        let sum = 0;
        let hasValue = false;
        Object.keys(byModelPeriod).forEach(m => {
          if (m !== TTL_MODEL) {
            const v = byModelPeriod[m][l];
            if (v != null && v > 0) {
              sum += v;
              hasValue = true;
            }
          }
        });
        byModelPeriod[TTL_MODEL][l] = hasValue ? sum : null;
      });
    }

    // Giới hạn 10 giai đoạn gần nhất cho Chart 2 (Nhân lực hàng ngày) để dễ đọc khi xem theo Ngày/Tuần
    const activeLabels = labels.filter(l => {
      const ttlVal = byModelPeriod[TTL_MODEL]?.[l];
      return ttlVal != null && ttlVal > 0;
    }).slice(-10);

    return { byModelPeriod, ttlStandard, activeModels, labels, activeLabels };
  }, [rows, dateFrom, dateTo, granularity]);
}

// ─── UPPH Data hook (Chart 2) ──────────────────────────────────────────────────
// EPCC — Explore: chart 1 (Tháng) và chart 2 (Tuần) trước đây dùng chung
// buildMonthlyChart/buildWeeklyChart với CÙNG dữ liệu byModelPeriod (TTL ManPower
// AVG) → 2 biểu đồ trình bày giống hệt nhau về ý nghĩa, chỉ khác trục thời gian.
// Plan: thay chart 2 bằng một góc nhìn NĂNG SUẤT khác hẳn — UPPH (Units Per
// Person-Hour) theo 3 loại dòng dữ liệu trong Excel:
//   DAY  TARGET (UPPH) / DAY  UPPH / DAY  UPPH 달성율(%)
//   NIGHT TARGET (UPPH) / NIGHT UPPH / NIGHT UPPH 달성율(%)
//   TTL  TARGET (UPPH) / TTL  UPPH / TTL  UPPH 달성율(%)
// Code: hook riêng useUpphData, tái dùng đúng logic phân loại tầng dữ liệu
// (classifyRowDateType/getPeriodLabel) như useManpowerData để 2 chart luôn khớp
// trục X (cùng granularity/date range) nhưng khác hẳn nội dung.
type UpphShift = 'DAY' | 'NIGHT' | 'TTL';
type UpphMetric = 'target' | 'actual' | 'rate';

function classifyUpphRow(typeStr: string): { shift: UpphShift; metric: UpphMetric } | null {
  const norm = String(typeStr || '').trim().toUpperCase();
  if (!norm.includes('UPPH')) return null;

  let shift: UpphShift | null = null;
  if (norm.startsWith('DAY')) shift = 'DAY';
  else if (norm.startsWith('NIGHT')) shift = 'NIGHT';
  else if (norm.startsWith('TTL')) shift = 'TTL';
  if (!shift) return null;

  let metric: UpphMetric;
  if (norm.includes('TARGET')) metric = 'target';
  else if (norm.includes('달성율') || norm.includes('RATE') || norm.includes('%')) metric = 'rate';
  else metric = 'actual';

  return { shift, metric };
}

interface UpphData {
  labels: string[]; // period labels theo đúng thứ tự thời gian, khớp granularity hiện tại
  byShift: Record<UpphShift, Record<string, { target: number | null; actual: number | null; rate: number | null }>>;
  hasData: boolean;
}

const EMPTY_UPPH: UpphData = {
  labels: [],
  byShift: { DAY: {}, NIGHT: {}, TTL: {} },
  hasData: false,
};

function useUpphData(
  rows: DataRow[],
  dateFrom: string,
  dateTo: string,
  granularity: 'day' | 'week' | 'month' | 'year'
): UpphData {
  return useMemo(() => {
    // EPCC (chart2-upph-v2) — Explore: bản cũ lấy UPPH từ TẤT CẢ Model (từng
    // model đơn lẻ SO1B01, SO1C2G... LẪN dòng Model=TTL) rồi gộp trung bình
    // chung 1 label → trộn lẫn 2 thang đo khác hẳn nhau (UPPH riêng từng model
    // vs UPPH tổng hợp "Tất cả Model"), ra số sai và không khớp hình tham chiếu
    // (Plan/Actual/Yield — Tất cả Model). Excel đã có sẵn dòng Model=TTL cho
    // đúng 9 loại DAY/NIGHT/TTL x TARGET(UPPH)/UPPH/UPPH 달성율 — đây chính là
    // dòng "Tất cả Model" cần dùng.
    // Plan: chỉ lấy dòng Model === 'TTL' làm nguồn cho chart 2.
    let upphRows = rows.filter(r => {
      const model = String((r as any).model || (r as any).Model || '').trim().toUpperCase();
      return model === TTL_MODEL && classifyUpphRow(String(r.type || r.Type || '')) !== null;
    });
    if (upphRows.length === 0) {
      // Fallback: Nếu file upload mới không có dòng Model = 'TTL' cho UPPH, lấy tất cả các model khác
      upphRows = rows.filter(r => {
        return classifyUpphRow(String(r.type || r.Type || '')) !== null;
      });
    }
    if (upphRows.length === 0) return EMPTY_UPPH;

    const parsedRows = upphRows.map(r => {
      const dateVal = (r as any).date || (r as any).Date || (r as any).month || (r as any).Month || '';
      return { ...r, _parsedDate: parseManpowerDate(dateVal), _dateType: classifyRowDateType(dateVal) };
    }).filter(r => {
      const dateVal = (r as any).date || (r as any).Date || (r as any).month || (r as any).Month || '';
      return r._parsedDate !== null && String(dateVal).trim().toUpperCase() !== 'YR24';
    }) as (DataRow & { _parsedDate: Date; _dateType: RowDateType })[];

    const wantedType: RowDateType = granularity === 'year' ? 'month' : granularity;
    let typedRows = parsedRows.filter(r => r._dateType === wantedType);
    if (typedRows.length === 0) typedRows = parsedRows;

    let filtered = typedRows;
    const fromDate = dateFrom ? new Date(dateFrom) : null;
    const toDate = dateTo ? new Date(dateTo) : null;
    if (toDate) toDate.setHours(23, 59, 59, 999);
    if (fromDate) filtered = filtered.filter(r => r._parsedDate >= fromDate);
    if (toDate) filtered = filtered.filter(r => r._parsedDate <= toDate);

    const labelToMinDate: Record<string, number> = {};
    filtered.forEach(r => {
      const label = getPeriodLabel(r._parsedDate, granularity);
      const time = r._parsedDate.getTime();
      if (!labelToMinDate[label] || time < labelToMinDate[label]) labelToMinDate[label] = time;
    });
    const sortedUpphLabels = Object.keys(labelToMinDate).sort((a, b) => labelToMinDate[a] - labelToMinDate[b]);
    // EPCC (upph-match-reference) — hình tham chiếu giới hạn "Tối đa 7 giai
    // đoạn gần nhất" cho mỗi biểu đồ nhỏ. Cắt về 7 label gần nhất theo thời
    // gian (giữ nguyên thứ tự tăng dần) để 3 khối DAY/NIGHT/TTL luôn hiển thị
    // tối đa 7 cột, kể cả khi khoảng lọc ngày rộng hơn.
    const labels = sortedUpphLabels.slice(-7); // EPCC: tối đa 7 giai đoạn gần nhất cho UPPH (Chart 1)

    // Gom nhóm theo shift/metric/label, giá trị trùng label lấy trung bình
    // (giống hành vi của useManpowerData) để ổn định khi có nhiều dòng raw/tuần.
    const groups: Record<UpphShift, Record<UpphMetric, Record<string, number[]>>> = {
      DAY: { target: {}, actual: {}, rate: {} },
      NIGHT: { target: {}, actual: {}, rate: {} },
      TTL: { target: {}, actual: {}, rate: {} },
    };
    filtered.forEach(r => {
      const cls = classifyUpphRow(String((r as any).type || (r as any).Type || ''));
      if (!cls) return;
      const val = Number((r as any).value ?? (r as any).Value);
      if (isNaN(val)) return;
      const label = getPeriodLabel(r._parsedDate, granularity);
      const bucket = groups[cls.shift][cls.metric];
      if (!bucket[label]) bucket[label] = [];
      bucket[label].push(val);
    });

    const avg = (arr: number[] | undefined): number | null =>
      arr && arr.length > 0 ? arr.reduce((a, b) => a + b, 0) / arr.length : null;

    const byShift: UpphData['byShift'] = { DAY: {}, NIGHT: {}, TTL: {} };
    (['DAY', 'NIGHT', 'TTL'] as UpphShift[]).forEach(shift => {
      byShift[shift] = {};
      labels.forEach(label => {
        const target = avg(groups[shift].target[label]);
        const actual = avg(groups[shift].actual[label]);
        // FIX(chart2-upph-v2): cột "UPPH 달성율" trong Excel lưu dạng PHÂN SỐ
        // (vd 0.88, 1.09) chứ không phải phần trăm — bản cũ dùng thẳng giá trị
        // này làm rate nên hiển thị sai (vd "1.0%" thay vì "109%"). Để luôn
        // đồng nhất với hình tham chiếu (Yield % dạng 91%, 109%, 121%...),
        // TỰ TÍNH rate = Actual/Target*100 mỗi khi có đủ target & actual;
        // chỉ dùng cột 달성율 gốc (đã nhân 100) làm phương án dự phòng khi
        // thiếu target hoặc actual.
        let rate: number | null = null;
        if (target != null && target > 0 && actual != null) {
          rate = (actual / target) * 100;
        } else {
          const rawRate = avg(groups[shift].rate[label]);
          rate = rawRate != null ? rawRate * 100 : null;
        }
        byShift[shift][label] = { target, actual, rate };
      });
    });

    const hasData = labels.some(l => byShift.TTL[l]?.actual != null || byShift.DAY[l]?.actual != null || byShift.NIGHT[l]?.actual != null);

    return { labels, byShift, hasData };
  }, [rows, dateFrom, dateTo, granularity]);
}

// ─── Chart builders ───────────────────────────────────────────────────────────
// EPCC (remove-chart1-duplicate) — Explore: "Nhân lực TB theo Tháng" (chart 1,
// stacked bar theo Model + line TTL theo THÁNG) và "Nhân lực hàng ngày" (chart 3,
// line theo Model + TTL theo NGÀY) cùng vẽ lại đúng 1 nguồn dữ liệu
// (useManpowerData.byModelPeriod) chỉ khác granularity hiển thị → trùng lặp ý
// nghĩa, chiếm chỗ mà không thêm thông tin mới.
// Plan: xoá hẳn buildMonthlyChart + panel/route render của nó, đồng thời xoá
// theo các phần thừa liên quan (chart id 'monthly', lệnh vẽ trong effect) để
// dồn không gian cho biểu đồ UPPH (chart 2) hiển thị rõ hơn.
// Code: xem phần Component bên dưới — panel chart 1 đã được gỡ khỏi JSX, biểu
// đồ UPPH chuyển sang chiếm trọn 1 hàng riêng (full width).

// EPCC (chart2-upph-v2) — Commit: buildUpphChart viết lại theo đúng phong
// cách của hình tham chiếu "Tình hình sản xuất — Plan/Actual/Yield(%) — Tất
// cả Model": 3 KHỐI CẠNH NHAU (DAY | NIGHT | TTL), mỗi khối là 1 combo chart
// riêng gồm bar nhóm Plan (Target) + Actual (cùng 1 cặp màu cố định cho cả 3
// khối, không đổi màu theo ca — giống hệt hình tham chiếu) và line nét chấm
// Yield(%) trên trục phụ riêng của khối đó. Mỗi khối có domain trục X/Y độc
// lập vì thang giá trị DAY/NIGHT/TTL khác nhau rất nhiều (TTL ~ DAY+NIGHT).
const UPPH_PLAN_COLOR = '#f97316';   // cam — Plan/Target, đồng bộ hình tham chiếu
const UPPH_ACTUAL_COLOR = '#2A788E'; // xanh ngọc đậm — Actual, màu theo yêu cầu người dùng
const UPPH_RATE_COLOR = '#60a5fa';   // xanh dương nhạt — Yield (%), nổi hơn trên nền tối

// 3 domain cột bằng nhau, có khoảng hở ở giữa để tách biệt 3 khối
const UPPH_COL_DOMAINS: [number, number][] = [
  [0, 0.30],
  [0.36, 0.66],
  [0.72, 1.0],
];

function buildUpphChart(
  data: UpphData,
  chartTextColor: string,
  chartGridColor: string,
  lang: 'vi'|'en'|'ko'
) {
  const traces: any[] = [];
  const labels = data.labels;
  const shifts: UpphShift[] = ['DAY', 'NIGHT', 'TTL'];
  const shiftLabel: Record<UpphShift, string> = {
    DAY: t('shiftDay', lang),
    NIGHT: t('shiftNight', lang),
    TTL: t('shiftTtl', lang),
  };

  // Chỉ vẽ ca nào thực sự có ít nhất 1 giá trị Actual/Target > 0
  const activeShifts = shifts.filter(s =>
    labels.some(l => (data.byShift[s][l]?.actual ?? 0) > 0 || (data.byShift[s][l]?.target ?? 0) > 0)
  );

  const layout: any = {
    barmode: 'group',
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { family: 'Inter, sans-serif', color: chartTextColor, size: 11 },
    // EPCC (upph-match-reference) — tăng margin.t để có chỗ cho tiêu đề +
    // legend riêng của từng khối (annotation), tắt hẳn legend chung ở cuối
    // chart (showlegend:false + không set 'legend') vì hình tham chiếu KHÔNG
    // có 1 legend dùng chung — mỗi khối tự có legend nhỏ ngay dưới tiêu đề.
    margin: { l: 15, r: 15, t: 55, b: 35 },
    showlegend: false,
    hovermode: 'x unified',
    annotations: [] as any[],
  };

  activeShifts.forEach((shift, i) => {
    // suffix trục theo Plotly convention: khối đầu dùng x/y, khối 2 dùng x2/y2 (bar)
    // + y3 (%) v.v — mỗi khối có 1 cặp trục y riêng (bar bên trái, % bên phải).
    const xIdx = i === 0 ? '' : String(i + 1);
    const yBarIdx = i === 0 ? '' : String(i * 2 + 1);
    const yRateIdx = String(i * 2 + 2);
    const xKey = `x${xIdx}`;
    const yBarKey = `y${yBarIdx}`;
    const yRateKey = `y${yRateIdx}`;
    const [domStart, domEnd] = UPPH_COL_DOMAINS[i] || [0, 1];

    const planVals = labels.map(l => {
      const v = data.byShift[shift][l]?.target;
      return v != null ? round1(v) : null;
    });
    const actualVals = labels.map(l => {
      const v = data.byShift[shift][l]?.actual;
      return v != null ? round1(v) : null;
    });
    const rateVals = labels.map(l => {
      const v = data.byShift[shift][l]?.rate;
      return v != null ? round1(v) : null;
    });

    traces.push({
      x: labels, y: planVals, xaxis: xKey, yaxis: yBarKey,
      name: t('upphTarget', lang), showlegend: false,
      type: 'bar', marker: { color: UPPH_PLAN_COLOR },
      text: planVals.map(v => v != null ? fmt1(v) : ''),
      textposition: 'inside',
      insidetextanchor: 'middle',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      hovertemplate: `<b>${shiftLabel[shift]} ${t('upphTarget', lang)}</b><br>%{x}: %{y:.1f} ${t('upphUnit', lang)}<extra></extra>`,
    });
    traces.push({
      x: labels, y: actualVals, xaxis: xKey, yaxis: yBarKey,
      name: t('upphActual', lang), showlegend: false,
      type: 'bar', marker: { color: UPPH_ACTUAL_COLOR },
      text: actualVals.map(v => v != null ? fmt1(v) : ''),
      textposition: 'inside',
      insidetextanchor: 'middle',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      hovertemplate: `<b>${shiftLabel[shift]} ${t('upphActual', lang)}</b><br>%{x}: %{y:.1f} ${t('upphUnit', lang)}<extra></extra>`,
    });
    if (rateVals.some(v => v != null)) {
      traces.push({
        x: labels, y: rateVals, xaxis: xKey, yaxis: yRateKey,
        name: t('upphRate', lang), showlegend: false,
        type: 'scatter', mode: 'lines+markers+text', cliponaxis: false,
        line: { color: UPPH_RATE_COLOR, width: 2.5, dash: 'solid', shape: 'spline', smoothing: 1 },
        marker: { size: 6, color: UPPH_RATE_COLOR },
        text: rateVals.map(v => v != null ? `${fmt1(v)}%` : ''),
        textposition: 'top center', textfont: { size: 9, color: UPPH_RATE_COLOR },
        hovertemplate: `<b>${shiftLabel[shift]} ${t('upphRate', lang)}</b><br>%{x}: %{y:.1f}%<extra></extra>`,
      });
    }

    // Tiêu đề nhỏ (DAY / NIGHT / TTL) kết hợp với Legend màu thành 1 dòng duy nhất để thu gọn tối đa không gian đứng
    const combinedText = `<b>${shiftLabel[shift].toUpperCase()}</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ` +
      `<span style="color:${UPPH_PLAN_COLOR}">■</span> ${t('upphTarget', lang)} &nbsp;&nbsp;&nbsp; ` +
      `<span style="color:${UPPH_ACTUAL_COLOR}">■</span> ${t('upphActual', lang)} &nbsp;&nbsp;&nbsp; ` +
      `<span style="color:${UPPH_RATE_COLOR}">─◆─</span> ${t('upphRate', lang)}`;

    layout.annotations.push({
      text: combinedText,
      x: (domStart + domEnd) / 2, y: 1.15, xref: 'paper', yref: 'paper',
      xanchor: 'center', yanchor: 'middle', showarrow: false,
      font: { size: 14, color: chartTextColor },
    });

    const maxVal = Math.max(
      ...planVals.filter((v): v is number => v !== null),
      ...actualVals.filter((v): v is number => v !== null),
      0
    );
    const barMax = maxVal > 0 ? maxVal * 1.45 : 10;

    layout[xKey.replace('x', 'xaxis')] = {
      domain: [domStart, domEnd], anchor: yBarKey,
      gridcolor: chartGridColor, tickfont: { size: 9 },
    };
    layout[yBarKey.replace('y', 'yaxis')] = {
      anchor: xKey, gridcolor: chartGridColor, tickfont: { size: 8 },
      range: [0, barMax],
      title: i === 0 ? { text: t('upphUnit', lang), font: { size: 9 } } : undefined,
    };
    layout[yRateKey.replace('y', 'yaxis')] = {
      overlaying: yBarKey, anchor: xKey, side: 'right', showgrid: false, tickfont: { size: 8 },
      range: [0, 150], ticksuffix: '%',
      title: i === activeShifts.length - 1 ? { text: t('upphRate', lang), font: { size: 9 } } : undefined,
    };
  });

  return { traces, layout };
}


function buildDailyChart(
  data: ManpowerData,
  chartTextColor: string,
  chartGridColor: string,
  lang: 'vi'|'en'|'ko',
  models: string[]
) {
  const traces: any[] = [];
  const activeLabels = data.activeLabels;

  // Chỉ giữ models có ít nhất 1 label có value > 0
  const activeModels = models.filter(m =>
    activeLabels.some(l => (data.byModelPeriod[m]?.[l] ?? 0) > 0)
  );

  activeModels.forEach((model, index) => {
    const vals = activeLabels.map(l => {
      const val = data.byModelPeriod[model]?.[l];
      return val != null && val > 0 ? round1(val) : null;
    });
    traces.push({
      x: activeLabels,
      y: vals,
      name: model,
      type: 'scatter',
      mode: 'lines+markers',
      connectgaps: false,
      line: { color: getModelColor(model, index), width: 1.8, shape: 'spline', smoothing: 1 },
      marker: { size: 5, color: getModelColor(model, index) },
      hovertemplate: `<b>${model}</b><br>%{x}: %{y:.0f}<extra></extra>`,
    });
  });

  const ttlVals = activeLabels.map(l => {
    const val = data.byModelPeriod[TTL_MODEL]?.[l];
    return val != null && val > 0 ? round1(val) : null;
  });
  traces.push({
    x: activeLabels,
    y: ttlVals,
    name: 'TTL',
    type: 'scatter',
    mode: 'lines+markers+text',
    connectgaps: false,
    cliponaxis: false,
    line: { color: '#14b8a6', width: 2.5, shape: 'spline', smoothing: 1 },
    marker: { size: 7, color: '#14b8a6' },
    text: ttlVals.map(v => v != null && v > 0 ? fmt1(v) : ''),
    textposition: 'top center',
    textfont: { size: 9, color: '#14b8a6', weight: 'bold' },
    yaxis: 'y2',
    hovertemplate: `<b>TTL</b><br>%{x}: %{y:.0f}<extra></extra>`,
  });

  const layout = {
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { family: 'Inter, sans-serif', color: chartTextColor, size: 11 },
    margin: { l: 45, r: 55, t: 28, b: 36 },
    legend: { orientation: 'h', y: -0.22, font: { size: 10 } },
    xaxis: { gridcolor: chartGridColor, tickfont: { size: 10 } },
    yaxis:  { gridcolor: chartGridColor, tickfont: { size: 9 }, title: { text: t('persons', lang), font: { size: 10 } } },
    yaxis2: { overlaying: 'y', side: 'right', showgrid: false, tickfont: { size: 9 } },
    hovermode: 'x unified',
  };

  return { traces, layout };
}


// ─── Component ────────────────────────────────────────────────────────────────
export const ManpowerDashboard: React.FC<ManpowerDashboardProps> = ({
  rows,
  theme,
  lang,
  onToggleTheme,
  setLang,
  onFileSelected,
}) => {
  const isDark = theme === 'dark';
  const chartTextColor = isDark ? '#e2e8f0' : '#1e293b';
  const chartGridColor = isDark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.06)';

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

  // ── Stable rows cache ──────────────────────────────────────────────────────
  // Đối xứng với SalesDashboard: khi allRows bị ghi đè bởi Test_1/Test_2 (dữ
  // liệu Sales không có type='manpower'), ManpowerDashboard sẽ mất data.
  // Fix: cache snapshot cuối cùng có manpower rows; nếu rows mới không phải
  // dữ liệu Mục 3 thì dùng lại snapshot cũ thay vì reset charts.
  const stableRowsRef = useRef<DataRow[]>([]);

  const effectiveRows = useMemo(() => {
    if (rows.length === 0) return stableRowsRef.current;
    // Nhận dạng: dữ liệu Mục 3 có ít nhất 1 dòng type chứa 'manpower'
    const hasManpower = rows.some(r =>
      String((r as any).type || (r as any).Type || '').toLowerCase().includes('manpower')
    );
    if (hasManpower) {
      stableRowsRef.current = rows;
      return rows;
    }
    // Rows từ dashboard khác → giữ lại snapshot cũ
    return stableRowsRef.current.length > 0 ? stableRowsRef.current : rows;
  }, [rows]);
  // ──────────────────────────────────────────────────────────────────────────

  // ── Tab state: 'manpower' | 'percapita' ──────────────────────────────────
  const [activeTab, setActiveTab] = useState<'manpower' | 'percapita'>('manpower');

  // Toolbar & Filter States
  const [dateFrom, setDateFrom] = useState<string>('');
  const [dateTo, setDateTo] = useState<string>('');
  const [granularity, setGranularity] = useState<'day' | 'week' | 'month' | 'year'>('month');
  const [modelFilter, setModelFilter] = useState<string>('all');

  // Error States
  const [dateError, setDateError] = useState<string | null>(null);

  // Save to Cloud state

  // Clock state
  const [currentTime, setCurrentTime] = useState<Date>(new Date());

  // FIX(toolbar-chuan-hoa): màu label đồng bộ với toolbar chuẩn (PerCapitaTab /
  // TargetActualDashboard) — tự đổi theo Light/Dark mode.
  const filterLabelColor = theme === 'light' ? 'var(--text-0)' : 'var(--text-2)';

  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  const formatClock = (d: Date) => {
    const pad = (n: number) => String(n).padStart(2, '0');
    const hh = pad(d.getHours());
    const mm = pad(d.getMinutes());
    const ss = pad(d.getSeconds());
    const day = pad(d.getDate());
    const month = pad(d.getMonth() + 1);
    const year = d.getFullYear();
    return `${hh}:${mm}:${ss} ${day}/${month}/${year}`;
  };

  // Nguồn data: dùng effectiveRows (stable cache) thay vì rows raw
  // để tránh bị reset khi allRows bị ghi đè bởi upload ở dashboard khác
  const data = useManpowerData(effectiveRows, dateFrom, dateTo, granularity);
  // Chart 2 (trước là Weekly Manpower, nay là UPPH theo Ca — xem buildUpphChart)
  const upphData = useUpphData(effectiveRows, dateFrom, dateTo, granularity);

  const hasData = data.activeLabels.length > 0;

  // Dynamic KPI Card Values
  const latestPeriod = data.activeLabels.length > 0 ? data.activeLabels[data.activeLabels.length - 1] : '-';

  const prevPeriod = data.activeLabels.length > 1 ? data.activeLabels[data.activeLabels.length - 2] : '';

  const ttlLatest = latestPeriod !== '-' ? round1(data.byModelPeriod[TTL_MODEL]?.[latestPeriod]) : 0;
  const ttlPrev = prevPeriod ? round1(data.byModelPeriod[TTL_MODEL]?.[prevPeriod]) : 0;
  const deltaLatest = ttlPrev > 0 ? ttlLatest - ttlPrev : 0;

  // peakPeriod: highest TTL value across the currently-filtered labels
  // (data.activeLabels already reflects dateFrom/dateTo/granularity filter from the hook)
  const peakPeriod = useMemo(() => {
    const ttlData = data.byModelPeriod[TTL_MODEL] || {};
    return data.activeLabels.reduce<{ period: string; val: number }>(
      (best, p) => {
        const v = round1(ttlData[p]);
        return v > best.val ? { period: p, val: v } : best;
      },
      { period: '-', val: 0 }
    );
  }, [data]);

  // Filter models for trace plotting
  const modelsToPlot = modelFilter === 'all' ? data.activeModels : [modelFilter];

  // Active models count: models in current filter that have at least 1 value > 0
  const activeModelsCount = useMemo(() => {
    const candidates = modelFilter === 'all' ? data.activeModels : [modelFilter];
    return candidates.filter(m =>
      data.activeLabels.some(l => (data.byModelPeriod[m]?.[l] ?? 0) > 0)
    ).length;
  }, [data, modelFilter]);

  // Chart rendering — chart 1 (Monthly) đã bị xoá vì trùng dữ liệu với chart 3
  // (Daily) — xem EPCC (remove-chart1-duplicate) ở phần buildDailyChart.
  const chartIds = useRef({
    upph:  'mp-chart-upph',
    daily: 'mp-chart-daily',
  });

  useEffect(() => {
    // Only draw when the Manpower tab is active — chart holder divs are not in the
    // DOM when activeTab !== 'manpower', so Plotly would throw a "Script error".
    if (activeTab !== 'manpower') return;
    if (!hasData || !plotlyReady) return;

    const draw = () => {
      const ids = chartIds.current;
      try {
        if (upphData.hasData) {
          const upph = buildUpphChart(upphData, chartTextColor, chartGridColor, lang);
          window.Plotly.react(ids.upph, upph.traces as any, upph.layout as any, { displayModeBar: false, responsive: true });
        }

        const daily = buildDailyChart(data, chartTextColor, chartGridColor, lang, modelsToPlot);
        window.Plotly.react(ids.daily, daily.traces as any, daily.layout as any, { displayModeBar: false, responsive: true });
      } catch (err) {
        console.warn('[ManpowerDashboard] Plotly render error (ignored):', err);
      }
    };

    const timerId = setTimeout(draw, 0);
    return () => clearTimeout(timerId);
  }, [data, upphData, chartTextColor, chartGridColor, lang, hasData, modelFilter, modelsToPlot, activeTab, plotlyReady]);

  // Toolbar Handlers
  const handleDateFromChange = (val: string) => {
    if (dateTo && val > dateTo) {
      setDateError(t('dateError', lang));
    } else {
      setDateFrom(val);
      setDateError(null);
    }
  };

  const handleDateToChange = (val: string) => {
    if (dateFrom && val < dateFrom) {
      setDateError(t('dateError', lang));
    } else {
      setDateTo(val);
      setDateError(null);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const fileData = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(fileData, { type: 'array', cellDates: true });
        onFileSelected(file, workbook);
      } catch (err) {
        alert(lang === 'vi' ? 'Lỗi đọc tệp Excel!' : lang === 'ko' ? '엑셀 파일 읽기 오류!' : 'Error reading Excel file!');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  return (
    <div className="sales-dashboard" style={{ padding: '0 24px 24px', boxSizing: 'border-box' }}>
      
      {/* Header ngang hàng với GlobalHeaderControls */}
      <div className="dashboard-header-grid stretched">
        <div className="dashboard-header-left" />
        <div className="dashboard-header-title" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span style={{ fontSize: '20px' }}>👷</span>
          <h1 className="dashboard-header-title" style={{ border: 'none', padding: 0 }}>
            {t('title', lang)}
          </h1>
        </div>
        <div className="dashboard-header-right">
          <GlobalHeaderControls 
            lang={lang} 
            setLang={setLang} 
            isDark={theme === 'dark'} 
            onToggleTheme={onToggleTheme} 
          />
        </div>
      </div>


        {/* ── Tab Navigation ── */}
        <div style={{ display: 'flex', gap: '0px', marginTop: '12px', borderBottom: '2px solid var(--border-soft)' }}>
          {([
            { id: 'manpower', label: lang === 'vi' ? '👷 근무 인력 현황 - Tình hình nhân lực' : lang === 'ko' ? '👷 근무 인력 현황' : '👷 Manpower Status' },
            { id: 'percapita', label: lang === 'vi' ? '📐 인당 생산수 현황 - Sản lượng theo đầu người' : lang === 'ko' ? '📐 인당 생산수 현황' : '📐 Per Capita Output' },
          ] as const).map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              style={{
                padding: '8px 20px',
                fontSize: '13px',
                fontWeight: activeTab === tab.id ? 700 : 500,
                border: 'none',
                borderBottom: activeTab === tab.id ? '2px solid #2e7d8c' : '2px solid transparent',
                background: 'transparent',
                color: activeTab === tab.id ? '#2e7d8c' : 'var(--text-3)',
                cursor: 'pointer',
                marginBottom: '-2px',
                transition: 'all 0.15s ease',
              }}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* ── Toolbar — chỉ hiển thị khi ở tab Manpower ──────────────────────
            FIX(toolbar-chuan-hoa): đồng bộ y hệt cấu trúc/kích thước/font-size
            với toolbar chuẩn tham chiếu ở TargetActualDashboard.tsx (ảnh 2) —
            layout 2 dòng (dòng nhãn + dòng control), đồng hồ dạng text thường
            (không bo viền/pill), input ngày cao 38px, CustomSelect cho Model,
            control cao 38px đồng nhất. Áp dụng cùng chuẩn đã dùng ở PerCapitaTab. */}
        {activeTab === 'manpower' && (
        <div className="topbar-dash" style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '12px', marginBottom: '16px' }}>
          {/* Dòng 1 (labels) */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
            {/* Cụm trái dòng 1: Đồng hồ */}
            <div style={{ width: '170px', display: 'flex', alignItems: 'center', flexShrink: 0, fontSize: '13px', color: filterLabelColor, fontWeight: '700', whiteSpace: 'nowrap' }}>
              {formatClock(currentTime)}
            </div>
            {/* Cụm giữa dòng 1: Labels */}
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px' }}>
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
            {/* Cụm phải dòng 1: Spacer matching Dòng 2 (Tải Excel) */}
            <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
              <div style={{ width: '120px' }}></div>
            </div>
          </div>

          {/* Dòng 2 (values/controls) */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
            {/* Cụm trái dòng 2: Trống để giữ căn lề */}
            <div style={{ width: '170px', flexShrink: 0 }}></div>
            {/* Cụm giữa dòng 2: Controls */}
            <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px', alignItems: 'center' }}>
              <input
                type="date"
                value={dateFrom}
                onChange={e => handleDateFromChange(e.target.value)}
                className="filter-date-input"
                style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
              />
              <input
                type="date"
                value={dateTo}
                onChange={e => handleDateToChange(e.target.value)}
                className="filter-date-input"
                style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
              />
              <CustomSelect
                value={modelFilter}
                onChange={setModelFilter}
                options={[
                  { value: 'all', label: t('allModels', lang) },
                  ...data.activeModels.map(m => ({ value: m, label: m })),
                ]}
                style={{ width: '140px', height: '38px' }}
              />
              <div style={{ display: 'flex', gap: '0px', height: '38px', width: '220px', flexShrink: 0 }}>
                {(['day', 'week', 'month', 'year'] as const).map(mode => (
                  <button
                    key={mode}
                    onClick={() => setGranularity(mode)}
                    style={{
                      flex: 1,
                      padding: '6px 0',
                      fontSize: '13px',
                      fontWeight: 600,
                      borderRadius: mode === 'day' ? '6px 0 0 6px' : mode === 'year' ? '0 6px 6px 0' : '0',
                      border: '1px solid var(--border)',
                      borderRight: mode !== 'year' ? 'none' : '1px solid var(--border)',
                      background: granularity === mode ? '#2e7d8c' : 'var(--surface-2)',
                      color: granularity === mode ? '#ffffff' : 'var(--text-2)',
                      cursor: 'pointer',
                      transition: 'all 0.15s ease',
                      height: '100%',
                      display: 'inline-flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {t(mode, lang)}
                  </button>
                ))}
              </div>
              {dateError && (
                <span style={{ color: 'var(--rose)', fontSize: '11px', fontWeight: 500, whiteSpace: 'nowrap' }}>
                  ⚠️ {dateError}
                </span>
              )}
            </div>

            {/* Cụm phải dòng 2: Tải Excel (upload) + Lên mây (Save to Cloud) */}
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexShrink: 0 }}>
              <label className="btn-outline" style={{ cursor: 'pointer', margin: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px', color: 'var(--text-1)', background: 'transparent', border: '1px solid var(--border)', height: '38px', width: '120px', boxSizing: 'border-box', fontSize: '13px' }}>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  style={{ display: 'none' }}
                  onChange={handleFileUpload}
                />
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15" style={{ flexShrink: 0 }}>
                  <path d="M21 12a9 9 0 0 1-9 9c-2.52 0-4.93-1-6.74-2.74L3 16" />
                  <path d="M3 12a9 9 0 0 1 9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
                  <path d="M3 16v5h5" />
                  <path d="M16 3h5v5" />
                </svg>
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{t('exportBtn', lang)}</span>
              </label>
            </div>
          </div>
        </div>
        )} {/* end activeTab === 'manpower' toolbar */}
      
      {/* ── Tab 2: Per Capita — luôn render (display:none khi inactive) để tránh
          unmount/remount mất Plotly chart khi đổi ngôn ngữ / dark-light mode ── */}
      <div style={{ display: activeTab === 'percapita' ? 'block' : 'none' }}>
        <PerCapitaTab
          rows={effectiveRows}
          theme={theme}
          lang={lang}
          onToggleTheme={onToggleTheme}
          setLang={setLang}
          onFileSelected={onFileSelected}
          isVisible={activeTab === 'percapita'}
        />
      </div>

      {/* ── Tab 1: Manpower data guard ── */}
      {activeTab === 'manpower' && (!hasData ? (
        <div className="panel-empty" style={{
          padding: '60px', textAlign: 'center', background: 'var(--surface)',
          border: '1px solid var(--border-soft)', borderRadius: '12px',
        }}>
          <div style={{ fontSize: '48px', marginBottom: '12px' }}>👷</div>
          <p style={{ color: 'var(--text-2)' }}>{t('noData', lang)}</p>
          <p style={{ color: 'var(--text-3)', fontSize: '13px', marginTop: '8px' }}>
            {lang === 'vi'
              ? 'Dữ liệu cần có cột: Model, Type (TTL ManPower AVG), Date, Value'
              : lang === 'en'
              ? 'Data needs columns: Model, Type (TTL ManPower AVG), Date, Value'
              : '데이터 컬럼 필요: Model, Type (TTL ManPower AVG), Date, Value'}
          </p>
        </div>
      ) : (
        <div className="dash-container" style={{ animation: 'fadeUp 0.5s ease both' }}>

          {/* ── KPI Row ── */}
          <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '14px', marginBottom: '20px' }}>

            <div className="kpi-card" style={{ '--kpi-bg': 'var(--purple-soft)' } as any}>
              <div className="kpi-card-header">
                <div className="kpi-card-icon" style={{ background: 'var(--purple-soft)', color: 'var(--purple)' }}>👥</div>
                <div className="kpi-card-label">{t('kpiTtlAvg', lang)} ({latestPeriod})</div>
              </div>
              <div className="kpi-card-value" style={{ fontSize: '22px' }}>
                {fmt1(ttlLatest)}
                <span style={{ fontSize: '13px', color: 'var(--text-3)', marginLeft: '4px' }}>{t('persons', lang)}</span>
              </div>
            </div>

            <div className="kpi-card" style={{ '--kpi-bg': 'var(--amber-soft)' } as any}>
              <div className="kpi-card-header">
                <div className="kpi-card-icon" style={{ background: 'var(--amber-soft)', color: 'var(--amber)' }}>📈</div>
                <div className="kpi-card-label">{t('kpiPeak', lang)}</div>
              </div>
              <div className="kpi-card-value" style={{ fontSize: '22px' }}>
                {peakPeriod.period}
                <span style={{ fontSize: '12px', color: 'var(--text-3)', marginLeft: '6px' }}>({fmt1(peakPeriod.val)})</span>
              </div>
            </div>

            <div className="kpi-card" style={{ '--kpi-bg': 'var(--cyan-soft)' } as any}>
              <div className="kpi-card-header">
                <div className="kpi-card-icon" style={{ background: 'var(--cyan-soft)', color: 'var(--cyan)' }}>🏭</div>
                <div className="kpi-card-label">{t('kpiModels', lang)}</div>
              </div>
              <div className="kpi-card-value" style={{ fontSize: '22px' }}>
                {activeModelsCount}
              </div>
            </div>

            <div className="kpi-card" style={{ '--kpi-bg': deltaLatest >= 0 ? 'var(--green-soft)' : 'var(--rose-soft)' } as any}>
              <div className="kpi-card-header">
                <div className="kpi-card-icon" style={{
                  background: deltaLatest >= 0 ? 'var(--green-soft)' : 'var(--rose-soft)',
                  color: deltaLatest >= 0 ? 'var(--green)' : 'var(--rose)'
                }}>
                  {deltaLatest >= 0 ? '↑' : '↓'}
                </div>
                <div className="kpi-card-label">{t('kpiLatest', lang)} ({latestPeriod})</div>
              </div>
              <div className="kpi-card-value" style={{ fontSize: '22px' }}>
                {fmt1(ttlLatest)}
                {prevPeriod && (
                  <span style={{
                    fontSize: '12px', marginLeft: '6px',
                    color: deltaLatest >= 0 ? 'var(--green)' : 'var(--rose)',
                  }}>
                    {deltaLatest >= 0 ? '+' : ''}{fmt1(deltaLatest)}
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* ── Row 1: UPPH (chart 2) — full width ──
              EPCC (remove-chart1-duplicate): chart 1 (Nhân lực TB theo Tháng)
              đã bị xoá vì trùng dữ liệu/ý nghĩa với chart 3 (Nhân lực hàng
              ngày) bên dưới — cả 2 đều vẽ byModelPeriod theo Model, chỉ khác
              granularity. Nhường toàn bộ chiều rộng hàng này cho UPPH để 3
              khối DAY/NIGHT/TTL hiển thị rõ hơn, đỡ bị bóp nhỏ. */}
          <div style={{ marginBottom: '16px' }}>
            <div className="panel">
              <div className="panel-head" style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap', marginBottom: '8px' }}>
                <h3 style={{ margin: 0 }}>{t('chartWeek', lang)}</h3>
                <span style={{
                  fontSize: '11px', color: 'var(--text-3)',
                  background: 'var(--surface-2)', padding: '2px 8px', borderRadius: '4px',
                  border: '1px solid var(--border)'
                }}>
                  {t('upphSubtitle', lang)}
                </span>
              </div>
              {upphData.hasData ? (
                <div className="chart-holder" id={chartIds.current.upph} style={{ minHeight: '300px' }} />
              ) : (
                <div style={{
                  minHeight: '300px', display: 'flex', alignItems: 'center', justifyContent: 'center',
                  textAlign: 'center', color: 'var(--text-3)', fontSize: '13px', padding: '0 16px',
                }}>
                  {t('noUpphData', lang)}
                </div>
              )}
            </div>
          </div>

          {/* ── Row 2: Daily Chart & Summary Table ── */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
            <div className="panel">
              <div className="panel-head">
                <h3>{t('chartDay', lang)}</h3>
              </div>
              <div className="chart-holder" id={chartIds.current.daily} style={{ minHeight: '300px' }} />
            </div>

            <div className="panel" style={{ overflowX: 'auto', position: 'relative' }}>
              <div className="panel-head">
                <h3>
                  {lang === 'vi' ? 'Bảng tổng hợp nhân lực' : lang === 'en' ? 'Manpower Summary Table' : '인력 현황 종합표'}
                </h3>
              </div>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '14px' }}>
                <thead>
                  <tr style={{ background: 'var(--surface-2)' }}>
                    <th style={{ ...thStyle, position: 'sticky', left: 0, zIndex: 3, background: 'var(--surface-2)' }}>Model</th>
                    {data.activeLabels.map((lbl, idx) => {
                      const isLatest = idx === data.activeLabels.length - 1;
                      return (
                        <th 
                          key={lbl} 
                          style={{ 
                            ...thStyle,
                            ...(isLatest ? { background: 'var(--purple-soft)', color: 'var(--purple)' } : {})
                          }}
                        >
                          {lbl}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {[
                    // Chỉ hiện models có ít nhất 1 giá trị > 0 trong data.activeLabels
                    ...modelsToPlot.filter(m =>
                      data.activeLabels.some(l => (data.byModelPeriod[m]?.[l] ?? 0) > 0)
                    ),
                    TTL_MODEL,
                  ].map((model, ri) => {
                    const isTtl = model === TTL_MODEL;
                    return (
                      <tr key={model} style={{ background: ri % 2 === 0 ? 'transparent' : 'var(--surface-2)' }}>
                        <td style={{
                          ...tdStyle,
                          fontWeight: isTtl ? 700 : 500,
                          color: getModelColor(model, ri),
                          position: 'sticky', left: 0, zIndex: 2,
                          background: ri % 2 === 0 ? 'var(--surface-1, #1a1d2e)' : 'var(--surface-2)',
                          boxShadow: '2px 0 4px rgba(0,0,0,0.25)',
                        }}>
                          {model}
                        </td>
                        {data.activeLabels.map((lbl, idx) => {
                          const v = data.byModelPeriod[model]?.[lbl];
                          const isLatest = idx === data.activeLabels.length - 1;
                          return (
                            <td 
                              key={lbl} 
                              style={{ 
                                ...tdStyle,
                                ...(isLatest ? { fontWeight: 600, color: 'var(--purple)' } : {})
                              }}
                            >
                              {v != null && v > 0 ? fmt1(round1(v)) : <span style={{ color: 'var(--text-3)' }}>—</span>}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>

              {data.ttlStandard != null && data.ttlStandard > 0 && (
                <div style={{
                  marginTop: '12px', padding: '8px 12px', borderRadius: '6px',
                  background: 'var(--rose-soft)', border: '1px dashed var(--rose)',
                  fontSize: '12px', color: 'var(--rose)',
                }}>
                  ── {t('standard', lang)}: TTL = <strong>{fmt1(round1(data.ttlStandard))}</strong> {t('persons', lang)}
                </div>
              )}
            </div>
          </div>

        </div>
      ))} {/* end activeTab === 'manpower' && (!hasData ? ... : ...) */}
    </div>
  );
};

const thStyle: React.CSSProperties = {
  padding: '7px 8px',
  textAlign: 'center',
  fontWeight: 600,
  fontSize: '13px',
  color: 'var(--text-2)',
  borderBottom: '1px solid var(--border)',
  whiteSpace: 'nowrap',
};

const tdStyle: React.CSSProperties = {
  padding: '6px 8px',
  textAlign: 'center',
  borderBottom: '1px solid var(--border-soft)',
  color: 'var(--text-1)',
  fontSize: '14px',
};
