import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import type { DataRow } from '../types';
import { parseToDate } from '../utils';
import { PerCapitaTab } from './PerCapitaTab';

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
  TTL:     '#a855f7',
};

// ─── i18n ─────────────────────────────────────────────────────────────────────
const T = {
  title:      { vi: '근무 인력 현황 (Manpower Dashboard)', en: 'Manpower Status Dashboard', ko: '근무 인력 현황' },
  subtitle:   { vi: 'Phân tích nhân lực theo Model & Thời gian', en: 'Manpower by Model & Period', ko: '모델별 인력 현황 분석' },
  kpiTtlAvg: { vi: 'Tổng nhân lực TB', en: 'Total Manpower Avg', ko: '총 인원 평균' },
  kpiPeak:   { vi: 'Giai đoạn cao điểm', en: 'Peak Period', ko: '최고 기간' },
  kpiModels: { vi: 'Models đang hoạt động', en: 'Active Models', ko: '가동 모델 수' },
  kpiLatest: { vi: 'Thời gian mới nhất', en: 'Latest Period', ko: '최근 기간' },
  chartMonth:{ vi: 'Nhân lực TB theo Tháng (TTL ManPower AVG)', en: 'Monthly Avg Manpower (TTL)', ko: '월별 평균 인원 (TTL)' },
  chartWeek: { vi: 'Nhân lực TB theo Tuần (Recent Weeks)', en: 'Weekly Avg Manpower (Recent)', ko: '주별 평균 인원 (최근)' },
  chartModel:{ vi: 'Phân bổ nhân lực theo Model — Giai đoạn gần nhất', en: 'Manpower by Model — Latest Period', ko: '모델별 인원 현황 (최근 기간)' },
  chartDay:  { vi: 'Nhân lực hàng ngày (TTL)', en: 'Daily Manpower (TTL)', ko: '일별 인원 현황 (TTL)' },
  chartRadar:{ vi: 'Phân bổ Model theo Giai đoạn (Spider)', en: 'Model Distribution by Period (Radar)', ko: '모델 인원 분포 (레이더)' },
  noData:    { vi: 'Không có dữ liệu nhân lực (Test_3)', en: 'No manpower data found (Test_3)', ko: '인력 데이터가 없습니다 (Test_3)' },
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
    // Test_3.xlsx: Type = "TTL ManPower AVG" — lọc tất cả dòng có type chứa "manpower"
    const mpRows = rows.filter(r => String(r.type || r.Type || '').toLowerCase().includes('manpower'));

    if (mpRows.length === 0) {
      return { byModelPeriod: {}, ttlStandard: null, activeModels: [], labels: [] };
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

    // Parse dates and validate
    const parsedRows = mpRows.map(r => {
      const dateVal = r.date || r.Date || '';
      const parsedDate = parseManpowerDate(dateVal);
      const dateType = classifyRowDateType(dateVal);
      return { ...r, _parsedDate: parsedDate, _dateType: dateType };
    }).filter(r =>
      r._parsedDate !== null && String(r.date || r.Date || '').trim().toUpperCase() !== 'YR24'
    ) as (DataRow & { _parsedDate: Date; _dateType: RowDateType })[];

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

    // Helper functions for granularity formatting
    const formatDayLabel = (d: Date) => {
      const dd = String(d.getDate()).padStart(2, '0');
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      return `${mm}/${dd}`;
    };
    const monthsAbbr = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

    const getPeriodLabel = (d: Date): string => {
      switch (granularity) {
        case 'day': return formatDayLabel(d);
        case 'week': return 'W' + getISOWeek(d);
        case 'month': return monthsAbbr[d.getMonth()];
        case 'year': return String(d.getFullYear());
        default: return monthsAbbr[d.getMonth()];
      }
    };

    // Extract sorted unique labels chronologically
    const labelToMinDate: Record<string, number> = {};
    filtered.forEach(r => {
      const label = getPeriodLabel(r._parsedDate);
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
      const label = getPeriodLabel(r._parsedDate);
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

    const activeLabels = labels.filter(l => {
      const ttlVal = byModelPeriod[TTL_MODEL]?.[l];
      return ttlVal != null && ttlVal > 0;
    });

    return { byModelPeriod, ttlStandard, activeModels, labels, activeLabels };
  }, [rows, dateFrom, dateTo, granularity]);
}

// ─── Chart builders ───────────────────────────────────────────────────────────
function buildMonthlyChart(
  data: ManpowerData,
  chartTextColor: string,
  chartGridColor: string,
  lang: 'vi'|'en'|'ko',
  models: string[]
) {
  const traces: any[] = [];
  const activeLabels = data.activeLabels;

  // Chỉ giữ models có ít nhất 1 label có value > 0 trong khoảng thời gian hiện tại
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
      type: 'bar',
      marker: { color: getModelColor(model, index) },
      hovertemplate: `<b>${model}</b><br>%{x}: %{y:.1f}<extra></extra>`,
    });
  });

  const ttlMonthly = activeLabels.map(l => {
    const val = data.byModelPeriod[TTL_MODEL]?.[l];
    return val != null && val > 0 ? round1(val) : null;
  });
  traces.push({
    x: activeLabels,
    y: ttlMonthly,
    name: 'TTL',
    type: 'scatter',
    mode: 'lines+markers+text',
    cliponaxis: false,
    line: { color: '#a855f7', width: 2.5, dash: 'dot' },
    marker: { size: 7, color: '#a855f7' },
    text: ttlMonthly.map(v => v != null && v > 0 ? fmt1(v) : ''),
    textposition: 'top center',
    textfont: { size: 10, color: '#a855f7', weight: 'bold' },
    yaxis: 'y2',
    hovertemplate: `<b>TTL</b><br>%{x}: %{y:.1f}<extra></extra>`,
  });

  const stdVal = round1(data.ttlStandard);
  if (stdVal > 0) {
    traces.push({
      x: activeLabels,
      y: Array(activeLabels.length).fill(stdVal),
      name: t('standard', lang),
      type: 'scatter',
      mode: 'lines',
      line: { color: '#f43f5e', width: 1.5, dash: 'dash' },
      yaxis: 'y2',
      hoverinfo: 'skip',
    });
  }

  const layout = {
    barmode: 'stack',
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { family: 'Inter, sans-serif', color: chartTextColor, size: 11 },
    margin: { l: 45, r: 55, t: 28, b: 36 },
    legend: { orientation: 'h', y: -0.18, font: { size: 10 } },
    xaxis: { gridcolor: chartGridColor, tickfont: { size: 10 } },
    yaxis:  { gridcolor: chartGridColor, tickfont: { size: 9 }, title: { text: t('persons', lang), font: { size: 10 } } },
    yaxis2: { overlaying: 'y', side: 'right', showgrid: false, tickfont: { size: 9 } },
    hovermode: 'x unified',
  };

  return { traces, layout };
}

function buildWeeklyChart(
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
      type: 'bar',
      marker: { color: getModelColor(model, index) },
      hovertemplate: `<b>${model}</b><br>%{x}: %{y:.1f}<extra></extra>`,
    });
  });

  const ttlWeekly = activeLabels.map(l => {
    const val = data.byModelPeriod[TTL_MODEL]?.[l];
    return val != null && val > 0 ? round1(val) : null;
  });
  traces.push({
    x: activeLabels,
    y: ttlWeekly,
    name: 'TTL',
    type: 'scatter',
    mode: 'lines+markers+text',
    cliponaxis: false,
    line: { color: '#a855f7', width: 2.5, dash: 'dot' },
    marker: { size: 7, color: '#a855f7' },
    text: ttlWeekly.map(v => v != null && v > 0 ? fmt1(v) : ''),
    textposition: 'top center',
    textfont: { size: 10, color: '#a855f7', weight: 'bold' },
    yaxis: 'y2',
    hovertemplate: `<b>TTL</b><br>%{x}: %{y:.1f}<extra></extra>`,
  });

  const stdVal = round1(data.ttlStandard);
  if (stdVal > 0) {
    traces.push({
      x: activeLabels,
      y: Array(activeLabels.length).fill(stdVal),
      name: t('standard', lang),
      type: 'scatter',
      mode: 'lines',
      line: { color: '#f43f5e', width: 1.5, dash: 'dash' },
      yaxis: 'y2',
      hoverinfo: 'skip',
    });
  }

  const layout = {
    barmode: 'stack',
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { family: 'Inter, sans-serif', color: chartTextColor, size: 11 },
    margin: { l: 45, r: 55, t: 28, b: 36 },
    legend: { orientation: 'h', y: -0.18, font: { size: 10 } },
    xaxis: { gridcolor: chartGridColor, tickfont: { size: 10 } },
    yaxis:  { gridcolor: chartGridColor, tickfont: { size: 9 }, title: { text: t('persons', lang), font: { size: 10 } } },
    yaxis2: { overlaying: 'y', side: 'right', showgrid: false, tickfont: { size: 9 } },
    hovermode: 'x unified',
  };

  return { traces, layout };
}

function buildModelBarChart(
  data: ManpowerData,
  chartTextColor: string,
  chartGridColor: string,
  models: string[]
): { traces: any[]; layout: any; isEmpty: boolean } {
  const latestPeriod = data.activeLabels.length > 0 ? data.activeLabels[data.activeLabels.length - 1] : '';

  // Chỉ giữ models (non-TTL) có value > 0 tại latestPeriod
  const displayModels = models.filter(m => (data.byModelPeriod[m]?.[latestPeriod] ?? 0) > 0);

  // Nếu không có non-TTL data → thử dùng TTL và các models khác trong byModelPeriod
  const fallbackModels = displayModels.length === 0
    ? Object.keys(data.byModelPeriod).filter(m => (data.byModelPeriod[m]?.[latestPeriod] ?? 0) > 0)
    : displayModels;

  // Không có data gì cả → trả isEmpty
  if (fallbackModels.length === 0) {
    return { traces: [], layout: {}, isEmpty: true };
  }

  const vals = fallbackModels.map(m => round1(data.byModelPeriod[m]?.[latestPeriod] ?? 0));
  const target = 90;

  const traces: any[] = [
    {
      x: fallbackModels,
      y: vals,
      type: 'bar',
      marker: { color: fallbackModels.map((m, idx) => getModelColor(m, idx)) },
      text: vals.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'outside',
      textfont: { size: 11, color: chartTextColor, weight: 'bold' },
      hovertemplate: `<b>%{x}</b><br>${latestPeriod} avg: %{y:.1f}<extra></extra>`,
      name: `${latestPeriod} avg`,
    },
    {
      x: fallbackModels,
      y: Array(fallbackModels.length).fill(target),
      type: 'scatter',
      mode: 'lines',
      line: { color: '#f59e0b', width: 2, dash: 'dash' },
      name: 'Target 90',
      hoverinfo: 'skip',
    },
  ];

  const maxVal = vals.length > 0 ? Math.max(...vals, 120) : 120;
  const layout = {
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { family: 'Inter, sans-serif', color: chartTextColor, size: 11 },
    margin: { l: 45, r: 20, t: 12, b: 80 },
    legend: { orientation: 'h', y: -0.25, font: { size: 10 } },
    xaxis: { gridcolor: chartGridColor, tickfont: { size: 10 }, tickangle: -30 },
    yaxis: { gridcolor: chartGridColor, tickfont: { size: 9 }, range: [0, maxVal * 1.3] },
    hovermode: 'x unified',
  };

  return { traces, layout, isEmpty: false };
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
      line: { color: getModelColor(model, index), width: 1.8 },
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
    line: { color: '#a855f7', width: 2.5 },
    marker: { size: 7, color: '#a855f7' },
    text: ttlVals.map(v => v != null && v > 0 ? fmt1(v) : ''),
    textposition: 'top center',
    textfont: { size: 9, color: '#a855f7', weight: 'bold' },
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

function buildRadarChart(
  data: ManpowerData,
  chartTextColor: string,
  models: string[]
): { traces: any[]; layout: any; isEmpty: boolean } {
  // Dùng 5 period gần nhất làm trục radar
  const radarLabels = data.activeLabels.slice(-5);
  if (radarLabels.length === 0) {
    return { traces: [], layout: {}, isEmpty: true };
  }

  // Chỉ giữ models có ít nhất 1 giá trị > 0 trong radarLabels
  const modelsForRadar = models.filter(m =>
    radarLabels.some(l => (data.byModelPeriod[m]?.[l] ?? 0) > 0)
  );

  // Fallback: nếu không có non-TTL data → dùng tất cả models trong byModelPeriod có data
  const finalModels = modelsForRadar.length > 0
    ? modelsForRadar
    : Object.keys(data.byModelPeriod).filter(m =>
        radarLabels.some(l => (data.byModelPeriod[m]?.[l] ?? 0) > 0)
      );

  if (finalModels.length === 0) {
    return { traces: [], layout: {}, isEmpty: true };
  }

  const traces: any[] = finalModels.map((model, index) => {
    const vals = [...radarLabels, radarLabels[0]].map(l =>
      round1(data.byModelPeriod[model]?.[l] ?? 0)
    );
    return {
      type: 'scatterpolar',
      r: vals,
      theta: [...radarLabels, radarLabels[0]],
      fill: 'toself',
      name: model,
      line: { color: getModelColor(model, index), width: 2 },
      marker: { size: 5 },
      fillcolor: getModelColor(model, index) + '22',
      hovertemplate: `<b>${model}</b><br>%{theta}: %{r:.1f}<extra></extra>`,
    };
  });

  // Tính max để set radialaxis range tự động
  const allVals = finalModels.flatMap(m =>
    radarLabels.map(l => data.byModelPeriod[m]?.[l] ?? 0)
  );
  const maxVal = allVals.length > 0 ? Math.max(...allVals) : 100;

  const layout = {
    paper_bgcolor: 'rgba(0,0,0,0)',
    polar: {
      bgcolor: 'rgba(0,0,0,0)',
      radialaxis: {
        visible: true,
        range: [0, maxVal * 1.15],
        gridcolor: chartTextColor + '30',
        tickfont: { size: 9, color: chartTextColor },
      },
      angularaxis: { tickfont: { size: 10, color: chartTextColor } },
    },
    font: { family: 'Inter, sans-serif', color: chartTextColor, size: 11 },
    margin: { l: 60, r: 60, t: 30, b: 30 },
    legend: { orientation: 'h', y: -0.15, font: { size: 10 } },
    showlegend: true,
  };

  return { traces, layout, isEmpty: false };
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

  // ── Tab state: 'manpower' | 'percapita' ──────────────────────────────────
  const [activeTab, setActiveTab] = useState<'manpower' | 'percapita'>('manpower');

  // Toolbar & Filter States
  const [dateFrom, setDateFrom] = useState<string>('');
  const [dateTo, setDateTo] = useState<string>('');
  const [granularity, setGranularity] = useState<'day' | 'week' | 'month' | 'year'>('month');
  const [modelFilter, setModelFilter] = useState<string>('all');

  // Error States
  const [dateError, setDateError] = useState<string | null>(null);

  // Clock state
  const [currentTime, setCurrentTime] = useState<Date>(new Date());

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

  // Nguồn data: dùng trực tiếp rows global từ App.tsx (đồng bộ với Production Per Capita Dashboard)
  const data = useManpowerData(rows, dateFrom, dateTo, granularity);

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

  // Chart rendering — 4 charts only
  const chartIds = useRef({
    monthly: 'mp-chart-monthly',
    weekly:  'mp-chart-weekly',
    daily:   'mp-chart-daily',
  });

  useEffect(() => {
    // Only draw when the Manpower tab is active — chart holder divs are not in the
    // DOM when activeTab !== 'manpower', so Plotly would throw a "Script error".
    if (activeTab !== 'manpower') return;
    if (!hasData || typeof window.Plotly === 'undefined') return;

    const draw = () => {
      const ids = chartIds.current;
      try {
        const monthly = buildMonthlyChart(data, chartTextColor, chartGridColor, lang, modelsToPlot);
        window.Plotly.react(ids.monthly, monthly.traces as any, monthly.layout as any, { displayModeBar: false, responsive: true });

        const weekly = buildWeeklyChart(data, chartTextColor, chartGridColor, lang, modelsToPlot);
        window.Plotly.react(ids.weekly, weekly.traces as any, weekly.layout as any, { displayModeBar: false, responsive: true });

        const daily = buildDailyChart(data, chartTextColor, chartGridColor, lang, modelsToPlot);
        window.Plotly.react(ids.daily, daily.traces as any, daily.layout as any, { displayModeBar: false, responsive: true });
      } catch (err) {
        console.warn('[ManpowerDashboard] Plotly render error (ignored):', err);
      }
    };

    const timerId = setTimeout(draw, 0);
    return () => clearTimeout(timerId);
  }, [data, chartTextColor, chartGridColor, lang, hasData, modelFilter, modelsToPlot, activeTab]);

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

  const handleExportExcel = () => {
    try {
      const sourceRows = rows;
      const mpRows = sourceRows.filter(r => String(r.type || r.Type || '').toLowerCase().includes('manpower'));

      const parsedRows = mpRows.map(r => {
        const dateVal = r.date || r.Date || '';
        const parsedDate = parseManpowerDate(dateVal);
        return { ...r, _parsedDate: parsedDate };
      }).filter(r => r._parsedDate !== null) as (DataRow & { _parsedDate: Date })[];

      let filtered = parsedRows;
      if (dateFrom) {
        filtered = filtered.filter(r => r._parsedDate >= new Date(dateFrom));
      }
      if (dateTo) {
        const toLimit = new Date(dateTo);
        toLimit.setHours(23, 59, 59, 999);
        filtered = filtered.filter(r => r._parsedDate <= toLimit);
      }
      if (modelFilter !== 'all') {
        filtered = filtered.filter(r => String(r.model || r.Model || '').trim().toUpperCase() === modelFilter.toUpperCase());
      }

      const exportData = filtered.map(r => ({
        [lang === 'vi' ? 'Model' : lang === 'ko' ? '모델' : 'Model']: r.model || r.Model || '',
        [lang === 'vi' ? 'Phân loại' : lang === 'ko' ? '구분' : 'Type']: r.type || r.Type || '',
        [lang === 'vi' ? 'Ngày/Kỳ' : lang === 'ko' ? '날짜/기간' : 'Date']: r.date || r.Date || '',
        [lang === 'vi' ? 'Giá trị' : lang === 'ko' ? '값' : 'Value']: r.value || r.Value || 0,
      }));

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'ManpowerData');

      const fName = `Manpower_${dateFrom || 'start'}_${dateTo || 'end'}.xlsx`;
      XLSX.writeFile(workbook, fName);
    } catch (err) {
      console.error("Failed to export manpower excel:", err);
    }
  };

  return (
    <div className="sales-dashboard" style={{ padding: '24px', boxSizing: 'border-box' }}>
      
      {/* ── Header ── */}
      <div className="hero" style={{ position: 'relative', padding: '16px 20px 4px', margin: '0 0 16px' }}>
        <div style={{ display: 'flex', alignItems: 'baseline', gap: '10px', justifyContent: 'center' }}>
          <span style={{ fontSize: '22px' }}>👷</span>
          <h1 style={{ fontSize: '26px', margin: 0, fontWeight: 700 }}>{t('title', lang)}</h1>
        </div>


        {/* ── Tab Navigation ── */}
        <div style={{ display: 'flex', gap: '0px', marginTop: '12px', borderBottom: '2px solid var(--border-soft)' }}>
          {([
            { id: 'manpower', label: lang === 'vi' ? '👷 근무 인력 현황' : lang === 'ko' ? '👷 근무 인력 현황' : '👷 Manpower Status' },
            { id: 'percapita', label: lang === 'vi' ? '📐 인당 생산수 현황' : lang === 'ko' ? '📐 인당 생산수 현황' : '📐 Per Capita Output' },
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

        {/* ── Toolbar Row — chỉ hiển thị khi ở tab Manpower ── */}
        {activeTab === 'manpower' && (
        <div className="filter-bar" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap', marginTop: '12px', padding: '10px 14px' }}>
          {/* Left: Clock & Upload */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <span className="pill-rows" style={{ margin: 0, fontSize: '12px', padding: '6px 10px', height: '34px', display: 'inline-flex', alignItems: 'center', fontWeight: '600' }}>
              ⏰ {formatClock(currentTime)}
            </span>
            
            <label className="btn btn-outline" style={{ cursor: 'pointer', margin: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px', fontSize: '13px', padding: '6px 12px', height: '34px', boxSizing: 'border-box' }}>
              <input 
                type="file" 
                accept=".xlsx, .xls" 
                style={{ display: 'none' }} 
                onChange={handleFileUpload} 
              />
              📤 {t('uploadBtn', lang)}
            </label>
          </div>

          {/* Center: Date Filters + Model Filter + Granularity — with column labels (matching PerCapitaTab) */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>

            {/* Date range with labels above */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>
              <div style={{ display: 'flex', gap: '6px' }}>
                <span style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text-2)', textTransform: 'uppercase', letterSpacing: '0.04em', width: '130px', textAlign: 'center' }}>
                  {lang === 'vi' ? 'Ngày bắt đầu' : lang === 'ko' ? '시작일' : 'Start Date'}
                </span>
                <span style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text-2)', textTransform: 'uppercase', letterSpacing: '0.04em', width: '130px', textAlign: 'center' }}>
                  {lang === 'vi' ? 'Ngày kết thúc' : lang === 'ko' ? '종료일' : 'End Date'}
                </span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <input
                  type="date"
                  value={dateFrom}
                  onChange={e => handleDateFromChange(e.target.value)}
                  style={{ height: '34px', padding: '4px 8px', fontSize: '13px', borderRadius: '4px', border: '1px solid var(--border)', background: 'var(--surface-2)', color: 'var(--text-0)', width: '130px' }}
                />
                <input
                  type="date"
                  value={dateTo}
                  onChange={e => handleDateToChange(e.target.value)}
                  style={{ height: '34px', padding: '4px 8px', fontSize: '13px', borderRadius: '4px', border: '1px solid var(--border)', background: 'var(--surface-2)', color: 'var(--text-0)', width: '130px' }}
                />
                {dateError && (
                  <span style={{ color: 'var(--rose)', fontSize: '11px', fontWeight: '500' }}>
                    ⚠️ {dateError}
                  </span>
                )}
              </div>
            </div>

            {/* Model dropdown with label above */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>
              <span style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text-2)', textTransform: 'uppercase', letterSpacing: '0.04em', textAlign: 'center' }}>
                {lang === 'vi' ? 'Model' : lang === 'ko' ? '모델' : 'Model'}
              </span>
              <select
                value={modelFilter}
                onChange={e => setModelFilter(e.target.value)}
                style={{ height: '34px', padding: '4px 8px', fontSize: '13px', borderRadius: '4px', border: '1px solid var(--border)', background: 'var(--surface-2)', color: 'var(--text-0)', minWidth: '110px' }}
              >
                <option value="all">{t('allModels', lang)}</option>
                {data.activeModels.map(m => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </select>
            </div>

            {/* Granularity pills with label above */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '3px' }}>
              <span style={{ fontSize: '10px', fontWeight: 700, color: 'var(--text-2)', textTransform: 'uppercase', letterSpacing: '0.04em', textAlign: 'center' }}>
                {lang === 'vi' ? 'Xem theo' : lang === 'ko' ? '보기 방식' : 'View By'}
              </span>
              <div style={{ display: 'flex', gap: '0px', height: '34px' }}>
                {(['day', 'week', 'month', 'year'] as const).map(mode => (
                  <button
                    key={mode}
                    onClick={() => setGranularity(mode)}
                    style={{
                      padding: '0 10px',
                      fontSize: '12px',
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
            </div>

          </div>

          {/* Right: empty — Export Excel đã xóa theo yêu cầu */}
          <div />

        </div>
        )} {/* end activeTab === 'manpower' toolbar */}
      </div>

      {/* ── Tab 2: Per Capita — luôn render (display:none khi inactive) để tránh
          unmount/remount mất Plotly chart khi đổi ngôn ngữ / dark-light mode ── */}
      <div style={{ display: activeTab === 'percapita' ? 'block' : 'none' }}>
        <PerCapitaTab
          rows={rows}
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

          {/* ── Row 1: Monthly & Weekly ── */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px', marginBottom: '16px' }}>
            <div className="panel">
              <div className="panel-head">
                <h3>{t('chartMonth', lang)}</h3>
              </div>
              <div className="chart-holder" id={chartIds.current.monthly} style={{ minHeight: '280px' }} />
            </div>
            <div className="panel">
              <div className="panel-head">
                <h3>{t('chartWeek', lang)}</h3>
              </div>
              <div className="chart-holder" id={chartIds.current.weekly} style={{ minHeight: '280px' }} />
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

            <div className="panel" style={{ overflow: 'auto' }}>
              <div className="panel-head">
                <h3>
                  {lang === 'vi' ? 'Bảng tổng hợp nhân lực' : lang === 'en' ? 'Manpower Summary Table' : '인력 현황 종합표'}
                </h3>
              </div>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                <thead>
                  <tr style={{ background: 'var(--surface-2)' }}>
                    <th style={thStyle}>Model</th>
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
  fontSize: '11px',
  color: 'var(--text-2)',
  borderBottom: '1px solid var(--border)',
  whiteSpace: 'nowrap',
};

const tdStyle: React.CSSProperties = {
  padding: '6px 8px',
  textAlign: 'center',
  borderBottom: '1px solid var(--border-soft)',
  color: 'var(--text-1)',
  fontSize: '12px',
};
