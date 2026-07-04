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
 *  Chart 2 — 생산 현황 (Production Status) — ước lượng từ headcount × target_unit
 *    Bar DAY + Bar NIGHT + line TTL + line TARGET
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

// ─── i18n ─────────────────────────────────────────────────────────────────────
const T = {
  title:        { vi: '인당 생산수 & 생산 현황 (근무인력 기준)', en: 'Per Capita Output & Production (Headcount Based)', ko: '인당 생산수 & 생산 현황 (근무인력 기준)' },
  subtitle:     { vi: 'Ước lượng từ dữ liệu nhân lực (DAY 50.7% / NIGHT 49.3%)', en: 'Estimated from manpower data (DAY 50.7% / NIGHT 49.3%)', ko: '근무인원 기준 추정 (DAY 50.7% / NIGHT 49.3%)' },
  chart1Title:  { vi: '인당생산수 - Sản lượng theo đầu người', en: 'Per Capita Output', ko: '인당생산수' },
  chart2Title:  { vi: '생산 현황 - Tình hình sản xuất', en: 'Production Status (estimated)', ko: '생산 현황' },
  chart3Title:  { vi: '근무 인력 현황 - Tình hình nhân lực đi làm', en: 'Manpower Trend', ko: '근무 인력 현황' },
  dayPc:        { vi: 'DAY 인당생산수', en: 'DAY Per Capita', ko: 'DAY 인당생산수' },
  nightPc:      { vi: 'NIGHT 인당생산수', en: 'NIGHT Per Capita', ko: 'NIGHT 인당생산수' },
  ttlPc:        { vi: 'TTL 인당생산수', en: 'TTL Per Capita', ko: 'TTL 인당생산수' },
  target:       { vi: 'TTL TARGET (인당생산수)', en: 'TTL TARGET', ko: 'TTL TARGET (인당생산수)' },
  mgmtTarget:   { vi: '경영목표', en: 'Mgmt Target', ko: '경영목표' },
  dayPro:       { vi: 'DAY Pro Actual', en: 'DAY Pro Actual', ko: 'DAY Pro Actual' },
  nightPro:     { vi: 'NIGHT Pro Actual', en: 'NIGHT Pro Actual', ko: 'NIGHT Pro Actual' },
  ttlPro:       { vi: 'TTL Pro Actual', en: 'TTL Pro Actual', ko: 'TTL Pro Actual' },
  targetPlan:   { vi: "TTL PRON'D PLAN", en: "TTL PRON'D PLAN", ko: "TTL PRON'D PLAN" },
  dayMP:        { vi: 'DAY ManPower AVG', en: 'DAY ManPower AVG', ko: 'DAY ManPower AVG' },
  nightMP:      { vi: 'NIGHT ManPower AVG', en: 'NIGHT ManPower AVG', ko: 'NIGHT ManPower AVG' },
  ttlMP:        { vi: 'TTL ManPower AVG', en: 'TTL ManPower AVG', ko: 'TTL ManPower AVG' },
  standard:     { vi: 'TTL 기준인력 Standard AVG', en: 'TTL Standard AVG', ko: 'TTL 기준인력 Standard AVG' },
  noData:       { vi: 'Không có dữ liệu nhân lực', en: 'No manpower data', ko: '인력 데이터 없음' },
  estimatedBadge: { vi: '⚠ Ước lượng', en: '⚠ Estimated', ko: '⚠ 추정값' },
  persons:      { vi: 'người', en: 'prs', ko: '명' },
  unit:         { vi: 'sp/người', en: 'pcs/cap', ko: '개/인' },
  kpiDayPC:     { vi: 'DAY 인당생산수 TB - Sản lượng đầu người ca ngày TB', en: 'AVG DAY Per Capita', ko: 'DAY 인당생산수 평균' },
  kpiNightPC:   { vi: 'NIGHT 인당생산수 TB - Sản lượng đầu người ca đêm TB', en: 'AVG NIGHT Per Capita', ko: 'NIGHT 인당생산수 평균' },
  kpiTtlPC:     { vi: 'TTL 인당생산수 TB - Sản lượng đầu người TTL TB', en: 'AVG TTL Per Capita', ko: 'TTL 인당생산수 평균' },
  kpiTarget:    { vi: '경영목표 - Target', en: 'Mgmt. Target', ko: '경영목표' },
};
function t(key: keyof typeof T, lang: 'vi'|'en'|'ko') {
  return (T[key] as any)[lang] ?? (T[key] as any)['vi'];
}

function r1(n: number | null | undefined): number {
  if (n == null || isNaN(n as number)) return 0;
  return Math.round((n as number) * 10) / 10;
}
function fmt1(n: number) {
  return n.toLocaleString('vi-VN', { maximumFractionDigits: 1 });
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
  ttlByLabel: Record<string, number>;       // TTL headcount per label
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
      return typeStr.includes('manpower') || typeStr.includes('인당생산수') || typeStr.includes('line') || typeStr.includes('라인');
    });
    if (mpRows.length === 0) {
      return { byModelLabel: {}, pcByModelLabel: {}, lineTtlByModelLabel: {}, lineDayByModelLabel: {}, lineNightByModelLabel: {}, ttlByLabel: {}, ttlStandard: null, allLabels: [], activeModels: [], lastDataDateMs: null, lastDataLabel: null };
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
    const groups: Record<string, Record<string, number[]>> = {};
    const pcGroups: Record<string, Record<string, number[]>> = {};
    const lineTtlGroups: Record<string, Record<string, number[]>> = {};
    const lineDayGroups: Record<string, Record<string, number[]>> = {};
    const lineNightGroups: Record<string, Record<string, number[]>> = {};
    (filtered as any[]).forEach(r => {
      const model = String(r.model || r.Model || '').trim().toUpperCase();
      const label = r._label;
      const val   = Number(r.value ?? r.Value);
      const typeStr = String(r.type || r.Type || '').toLowerCase();
      if (!model || isNaN(val) || val == null) return;

      if (typeStr.includes('line') || typeStr.includes('라인')) {
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
      } else if (typeStr.includes('manpower')) {
        if (!groups[model]) groups[model] = {};
        if (!groups[model][label]) groups[model][label] = [];
        groups[model][label].push(val);
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
        byModelLabel[model][lbl] = vals.reduce((a, b) => a + b, 0) / vals.length;
      });
    });

    const pcByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(pcGroups).forEach(([model, lblMap]) => {
      pcByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        pcByModelLabel[model][lbl] = vals.reduce((a, b) => a + b, 0) / vals.length;
      });
    });

    const lineTtlByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(lineTtlGroups).forEach(([model, lblMap]) => {
      lineTtlByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        lineTtlByModelLabel[model][lbl] = vals.reduce((a, b) => a + b, 0) / vals.length;
      });
    });

    const lineDayByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(lineDayGroups).forEach(([model, lblMap]) => {
      lineDayByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        lineDayByModelLabel[model][lbl] = vals.reduce((a, b) => a + b, 0) / vals.length;
      });
    });

    const lineNightByModelLabel: Record<string, Record<string, number>> = {};
    Object.entries(lineNightGroups).forEach(([model, lblMap]) => {
      lineNightByModelLabel[model] = {};
      Object.entries(lblMap).forEach(([lbl, vals]) => {
        lineNightByModelLabel[model][lbl] = vals.reduce((a, b) => a + b, 0) / vals.length;
      });
    });

    const ttlByLabel: Record<string, number> = byModelLabel['TTL'] ?? {};

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

    return { byModelLabel, pcByModelLabel, lineTtlByModelLabel, lineDayByModelLabel, lineNightByModelLabel, ttlByLabel, ttlStandard, allLabels, activeModels, lastDataDateMs, lastDataLabel };
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
  isDark: boolean
) {
  // Màu accent đậm hơn ở light mode để chữ/đường không bị mờ trên nền sáng
  const tealAccent = isDark ? '#14b8a6' : '#0f766e';
  const roseAccent    = isDark ? '#f43f5e' : '#be123c';

  const dayVals    = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * DAY_RATIO));
  const nightVals  = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * NIGHT_RATIO));
  const ttlVals    = labels.map(l => r1((data.ttlByLabel[l] ?? 0)));
  const targetVals = labels.map(() => targetPC);

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
      line: { color: tealAccent, width: 1.6, shape: 'spline', smoothing: 1 },
      marker: { color: tealAccent, size: 6 },
      text: ttlVals.map(v => v > 0 ? fmt1(v) : ''),
      textposition: 'top center',
      textfont: { size: 10, color: tealAccent, family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      hovertemplate: `<b>${t('ttlPc', lang)}</b>: %{y}<extra></extra>`,
    },
    {
      x: labels, y: targetVals, name: t('mgmtTarget', lang),
      type: 'scatter', mode: 'lines',
      line: { color: roseAccent, width: 1.6, dash: 'dash', shape: 'spline', smoothing: 1 },
    },
  ];

  const maxVal = Math.max(...ttlVals, ...dayVals, ...nightVals, targetPC, 0);
  const yRange = [0, maxVal > 0 ? maxVal * 1.15 : 10];

  const layout = {
    barmode: 'group',
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor, size: 11 },
    margin: { t: 44, r: 12, b: 28, l: 48 },
    legend: { orientation: 'h', x: 0, y: 1.05, xanchor: 'left', yanchor: 'bottom', font: { size: 10 } },
    xaxis: { tickfont: { size: 11, color: textColor }, gridcolor: gridColor },
    yaxis: { gridcolor: gridColor, tickfont: { size: 10 }, range: yRange },
    hoverlabel: { font: { size: 10 } },
  };
  return { traces, layout };
}

/**
 * Chart 2: 생산 현황
 * X = allLabels; Y = headcount × ratio × targetPC ≈ sản lượng ước lượng
 */
function buildChart2(
  data: PCTabData,
  labels: string[],
  textColor: string,
  gridColor: string,
  lang: 'vi'|'en'|'ko',
  targetPC: number,
  isDark: boolean
) {
  const tealAccent = isDark ? '#14b8a6' : '#0f766e';
  const amberAccent   = isDark ? '#f59e0b' : '#b45309';

  const dayPro   = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * DAY_RATIO   * targetPC));
  const nightPro = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * NIGHT_RATIO * targetPC));
  const ttlPro   = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * targetPC));
  const planLine = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * targetPC * 1.05)); // +5% plan

  const traces: any[] = [
    {
      x: labels, y: dayPro, name: t('dayPro', lang),
      type: 'bar', marker: { color: '#1565C0' },
      text: dayPro.map(v => v > 0 ? (v >= 1000 ? Math.round(v/1000)+'k' : fmt1(v)) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      x: labels, y: nightPro, name: t('nightPro', lang),
      type: 'bar', marker: { color: '#ef4444' },
      text: nightPro.map(v => v > 0 ? (v >= 1000 ? Math.round(v/1000)+'k' : fmt1(v)) : ''),
      textposition: 'inside',
      textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
      insidetextanchor: 'middle',
    },
    {
      x: labels, y: ttlPro, name: t('ttlPro', lang),
      type: 'scatter', mode: 'lines+markers+text',
      line: { color: tealAccent, width: 1.6, shape: 'spline', smoothing: 1 },
      marker: { color: tealAccent, size: 6 },
      text: ttlPro.map(v => v > 0 ? (v >= 1000 ? Math.round(v/1000)+'k' : fmt1(v)) : ''),
      textposition: 'top center',
      textfont: { size: 10, color: tealAccent, family: 'Arial Black, Arial, sans-serif' },
      cliponaxis: false,
      hovertemplate: `<b>${t('ttlPro', lang)}</b>: %{y:,.0f}<extra></extra>`,
    },
    {
      x: labels, y: planLine, name: t('targetPlan', lang),
      type: 'scatter', mode: 'lines',
      line: { color: amberAccent, width: 1.2, dash: 'dash', shape: 'spline', smoothing: 1 },
    },
  ];

  const maxVal = Math.max(...ttlPro, ...planLine, ...dayPro, ...nightPro, 0);
  const yRange = [0, maxVal > 0 ? maxVal * 1.15 : 10000];

  const layout = {
    barmode: 'group',
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor, size: 11 },
    margin: { t: 44, r: 12, b: 28, l: 56 },
    legend: { orientation: 'h', x: 0, y: 1.05, xanchor: 'left', yanchor: 'bottom', font: { size: 10 } },
    xaxis: { tickfont: { size: 11, color: textColor }, gridcolor: gridColor },
    yaxis: { gridcolor: gridColor, tickfont: { size: 10 }, range: yRange },
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
  isDark: boolean
) {
  const tealAccent = isDark ? '#14b8a6' : '#0f766e';
  const roseAccent    = isDark ? '#f43f5e' : '#be123c';

  const dayMP   = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * DAY_RATIO));
  const nightMP = labels.map(l => r1((data.ttlByLabel[l] ?? 0) * NIGHT_RATIO));
  const ttlMP   = labels.map(l => r1(data.ttlByLabel[l] ?? 0));
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
    margin: { t: 55, r: 15, b: 35, l: 65 },
    legend: { orientation: 'h', x: 0, y: 1.05, xanchor: 'left', yanchor: 'bottom', font: { size: 13 } },
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
  const headcountModels = models.filter(m => (data.byModelLabel[m]?.[latestLabel] ?? 0) > 0);

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

  // FIX: union — model có Line nhưng thiếu nhân lực vẫn phải xuất hiện trên trục X
  // (cột nhân lực = 0) để đường Line không bị "mất điểm" một cách vô lý.
  // Giữ nguyên thứ tự headcountModels trước, model chỉ-có-Line (không có nhân lực)
  // xếp sau cùng theo alphabet. indexOf trả -1 cho model không có nhân lực nên
  // phải map về Infinity, tránh bị "-1" đẩy nhầm lên đầu danh sách.
  const modelsWithData = Array.from(new Set([...headcountModels, ...lineModels]))
    .sort((a, b) => {
      const ia = headcountModels.indexOf(a); const ib = headcountModels.indexOf(b);
      const ra = ia === -1 ? Infinity : ia; const rb = ib === -1 ? Infinity : ib;
      return ra - rb || a.localeCompare(b);
    });

  const hasLineData = lineLabel !== '' && modelsWithData.some(m => (data.lineTtlByModelLabel[m]?.[lineLabel] ?? 0) > 0);

  return { latestLabel, modelsWithData, lineLabel, hasLineData };
}

/**
 * Chart 4: Phân bố nhân lực theo Model — Giai đoạn gần nhất
 * Bar chart: X = activeModels, Y = headcount tại label mới nhất có dữ liệu
 * + Line chart (trục y phụ): số line YÊU CẦU theo Model tại cùng giai đoạn (Type = 'TTL LINE Q'TY')
 */
function buildChart4(
  data: PCTabData,
  displayLabels: string[],
  textColor: string,
  gridColor: string,
  lang: 'vi'|'en'|'ko'
) {
  const { latestLabel, modelsWithData, lineLabel, hasLineData } = resolveChart4Info(data, displayLabels);

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
      line: { color: '#facc15', width: 2, shape: 'spline', smoothing: 1 },
      marker: { color: '#facc15', size: 9, symbol: 'diamond' },
      text: lineTtlVals.map(v => v != null && v > 0 ? fmt1(r1(v)) : ''),
      textposition: 'top center',
      textfont: { size: 12, color: '#facc15', family: 'Arial Black, Arial, sans-serif' },
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
      tickfont: { size: 12, color: '#facc15' },
      title: { text: lang === 'vi' ? 'Số line yêu cầu' : lang === 'ko' ? '요구 라인 수' : 'Required Lines', font: { size: 13, color: '#facc15' } },
      range: [0, lineMax > 0 ? lineMax * 1.4 : 10],
    };
  }

  return { traces, layout, hasLineData, lineLabel, latestLabel };
}

/**
 * Chart 5: Phân bố Model theo Giai đoạn (Spider / Radar)
 * Mỗi trace = 1 time label (giai đoạn), mỗi axis = 1 model
 * Chọn tối đa 6 giai đoạn gần nhất có dữ liệu để chart không quá rối
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
    return { traces: [], layout: {} };
  }

  // Use pcByModelLabel if available, otherwise fallback to byModelLabel (manpower)
  const hasPcData = Object.keys(data.pcByModelLabel || {}).length > 0;
  const activeDataMap = hasPcData ? data.pcByModelLabel : data.byModelLabel;

  // Lọc các label có ít nhất 1 model có dữ liệu, lấy tối đa 6 label cuối
  const labelsWithData = displayLabels.filter(lbl =>
    models.some(m => (activeDataMap[m]?.[lbl] ?? 0) > 0)
  );
  const selectedLabels = labelsWithData.slice(-6);

  const radarColors = [
    '#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4',
  ];

  // Chỉ giữ models có ít nhất 1 giá trị > 0 trong các selectedLabels
  const modelsWithData = models.filter(m =>
    selectedLabels.some(lbl => (activeDataMap[m]?.[lbl] ?? 0) > 0)
  );
  if (modelsWithData.length === 0) return { traces: [], layout: {} };

  // Đóng vòng radar: thêm model đầu vào cuối
  const modelsClosed = [...modelsWithData, modelsWithData[0]];

  const traces: any[] = selectedLabels.map((lbl, i) => {
    const rVals = modelsClosed.map(m => r1(activeDataMap[m]?.[lbl] ?? 0));
    // FIX(spider-labels): hiển thị giá trị từng điểm theo đúng màu của trace để
    // đọc số trực tiếp trên biểu đồ, không cần hover. Điểm cuối bị lặp lại
    // (đóng vòng radar) nên bỏ trống label để tránh chồng số lên chính nó.
    const textVals = rVals.map((v, idx) =>
      idx === rVals.length - 1 ? '' : (v > 0 ? fmt1(v) : '')
    );
    return {
      type: 'scatterpolar',
      r: rVals,
      theta: modelsClosed,
      fill: 'toself',
      fillcolor: radarColors[i % radarColors.length] + '22', // 13% opacity
      line: { color: radarColors[i % radarColors.length], width: 2 },
      mode: 'lines+markers+text',
      marker: { color: radarColors[i % radarColors.length], size: 5 },
      text: textVals,
      texttemplate: '%{text}',
      textposition: 'top center',
      textfont: { color: radarColors[i % radarColors.length], size: 10, family: 'inherit' },
      name: lbl,
      hovertemplate: `${lbl}<br>%{theta}: %{r}<extra></extra>`,
    };
  });

  const layout = {
    paper_bgcolor: 'transparent',
    polar: {
      bgcolor: 'transparent',
      radialaxis: {
        visible: true,
        color: textColor,
        gridcolor: textColor + '30',
        tickfont: { size: 12, color: textColor },
      },
      angularaxis: {
        tickfont: { size: 13, color: textColor },
        gridcolor: textColor + '20',
      },
    },
    font: { color: textColor, size: 13 },
    margin: { t: 25, r: 85, b: 25, l: 45 },
    legend: { orientation: 'v', x: 1.05, y: 0.5, xanchor: 'left', yanchor: 'middle', font: { size: 13, color: textColor } },
    showlegend: true,
  };
  return { traces, layout };
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
      const hasKind  = ts.includes('plan') || ts.includes('actual');
      const hasShift = ts.startsWith('day') || ts.startsWith('night') || ts.startsWith('ttl');
      if (!hasKind || !hasShift) return false;

      if (modelFilter !== 'all') {
        const m = String(r.model || (r as any).Model || '').trim().toUpperCase();
        if (m !== modelFilter.trim().toUpperCase()) return false;
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
      if (ts.includes('plan'))        kind = 'plan';
      else if (ts.includes('actual')) kind = 'actual';
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
    margin: { t: 36, r: 48, b: 28, l: 48 },
    legend: { orientation: 'h', x: 0, y: 1.06, xanchor: 'left', yanchor: 'bottom', font: { size: 10 } },
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
  const filterLabelColor = theme === 'light' ? 'var(--text-0)' : 'var(--text-2)';

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
  const [granularity, setGranularity] = useState<'day'|'week'|'month'|'year'>('month');
  const [modelFilter, setModelFilter] = useState<string>('all');
  const [dateError, setDateError] = useState<string | null>(null);

  // Clock (giống ManpowerDashboard)
  const [currentTime, setCurrentTime] = useState<Date>(new Date());
  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);
  const formatClock = (d: Date) => {
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())} ${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()}`;
  };

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

  // Lọc theo modelFilter (nếu không phải TTL-only)
  const filteredLabels = useMemo(() => {
    if (granularity === 'year')  return labelsAll.filter(l => /^YR/i.test(l));
    if (granularity === 'month') return labelsAll.filter(l => /^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)$/i.test(l));
    if (granularity === 'week')  return labelsAll.filter(l => /^W\d+$/i.test(l));
    // day
    return labelsAll.filter(l => /^\d{2}\/\d{2}$/.test(l));
  }, [labelsAll, granularity]);

  const displayLabels = filteredLabels.length > 0 ? filteredLabels : labelsAll;

  // KPI averages
  const avgTtl = displayLabels.length > 0
    ? r1(displayLabels.reduce((s, l) => s + (data.ttlByLabel[l] ?? 0), 0) / displayLabels.length)
    : 0;
  const avgDay   = r1(avgTtl * DAY_RATIO);
  const avgNight = r1(avgTtl * NIGHT_RATIO);

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
      const labels = displayLabels;
      // FIX(prod-last-7): 3 chart sản xuất DAY/NIGHT/TTL chỉ nên hiện tối đa
      // 7 giai đoạn GẦN NHẤT để tránh chart bị dồn quá nhiều cột khi khoảng
      // ngày lọc rộng — tách riêng khỏi `labels` dùng cho Chart 1-3 (Chart 1-3
      // vẫn hiển thị đầy đủ toàn bộ khoảng đã lọc).
      const prodLabels = labels.slice(-7);
      try {
        const ch1 = buildChart1(data, labels, textColor, gridColor, lang, targetPC, isDark);
        window.Plotly.react(ids.current.c1, ch1.traces, ch1.layout, { displayModeBar: false, responsive: true });

        const ch2 = buildChart2(data, labels, textColor, gridColor, lang, targetPC, isDark);
        window.Plotly.react(ids.current.c2, ch2.traces, ch2.layout, { displayModeBar: false, responsive: true });

        const ch3 = buildChart3(data, labels, textColor, gridColor, lang, isDark);
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
        const ch4 = buildChart4(data, labels, textColor, gridColor, lang);
        window.Plotly.react(ids.current.c4, ch4.traces, ch4.layout, { displayModeBar: false, responsive: true });

        const ch5 = buildChart5(data, labels, textColor, gridColor, lang);
        window.Plotly.react(ids.current.c5, ch5.traces, ch5.layout, { displayModeBar: false, responsive: true });
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

  return (
    <div style={{ padding: '0 0 24px' }}>

      {/* ══════════════════════════════════════════════════════════════════════
          TOOLBAR — FIX(toolbar-chuan-hoa): đồng bộ y hệt cấu trúc/kích thước/
          font-size với toolbar chuẩn tham chiếu ở TargetActualDashboard.tsx
          (ảnh 3) — layout 2 dòng (dòng nhãn + dòng control), đồng hồ dạng text
          thường, input ngày cao 38px, CustomSelect cho Model.
         ══════════════════════════════════════════════════════════════════════ */}
      <div className="topbar-dash" style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '16px' }}>
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
          {/* Cụm phải dòng 1: Spacer matching Dòng 2 (badge dùng chung) */}
          <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
            <div style={{ width: '220px' }}></div>
          </div>
        </div>

        {/* Dòng 2 (values/controls) */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          {/* Cụm trái dòng 2: Trống để giữ căn lề */}
          <div style={{ width: '170px', flexShrink: 0 }}></div>
          {/* Cụm giữa dòng 2: Controls */}
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px', alignItems: 'center' }}>
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
                      border: '1px solid var(--border)',
                      borderRight: mode !== 'year' ? 'none' : '1px solid var(--border)',
                      background: isActive ? '#2e7d8c' : 'var(--surface-2)',
                      color: isActive ? '#ffffff' : 'var(--text-2)',
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

      {/* Ratio info row */}
      <div style={{
        fontSize: '13px', color: 'var(--text-3)',
        padding: '4px 14px 10px',
        display: 'flex', gap: '16px', flexWrap: 'wrap', alignItems: 'center',
      }}>
        <span>DAY: <strong style={{ color: 'var(--text-1)' }}>{(DAY_RATIO * 100).toFixed(1)}%</strong></span>
        <span>NIGHT: <strong style={{ color: 'var(--text-1)' }}>{(NIGHT_RATIO * 100).toFixed(1)}%</strong></span>
        <span>
          {lang === 'vi' ? '경영목표' : lang === 'ko' ? '경영목표' : 'Mgmt Target'}:
          <strong style={{ color: '#2e7d8c', marginLeft: '4px' }}>{targetPC} {lang === 'vi' ? 'cái/người' : lang === 'ko' ? '개/인' : 'pcs/cap'}</strong>
        </span>
        {data.ttlStandard && (
          <span>
            {lang === 'vi' ? 'Chuẩn YR24' : lang === 'ko' ? '기준인원(YR24)' : 'Standard YR24'}:
            <strong style={{ color: '#ef4444', marginLeft: '4px' }}>{fmt1(r1(data.ttlStandard))} {lang === 'vi' ? 'người' : lang === 'ko' ? '명' : 'prs'}</strong>
          </span>
        )}
        {data.lastDataLabel && (
          <span>
            {lang === 'vi' ? 'Dữ liệu cập nhật đến' : lang === 'ko' ? '데이터 기준' : 'Data as of'}:
            <strong style={{ color: '#10b981', marginLeft: '4px' }}>{data.lastDataLabel}</strong>
          </span>
        )}
      </div>

      {/* ── KPI Row ── */}
      {hasData && (
        <div
          className="kpi-grid"
          style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '12px', marginBottom: '18px' }}
        >
          {[
            {
              icon: '☀️',
              label: lang === 'vi' ? 'Sản lượng đầu người ca ngày TB' : lang === 'ko' ? 'DAY 인당생산수 평균' : 'AVG DAY Per Capita',
              value: fmt1(avgDay),
              unit: lang === 'vi' ? 'sp/người' : lang === 'ko' ? '개/인' : 'pcs/cap',
              bg: 'var(--blue-soft)', color: '#1565C0',
            },
            {
              icon: '🌙',
              label: lang === 'vi' ? 'Sản lượng đầu người ca đêm TB' : lang === 'ko' ? 'NIGHT 인당생산수 평균' : 'AVG NIGHT Per Capita',
              value: fmt1(avgNight),
              unit: lang === 'vi' ? 'sp/người' : lang === 'ko' ? '개/인' : 'pcs/cap',
              bg: 'var(--rose-soft)', color: '#ef4444',
            },
            {
              icon: '📊',
              label: lang === 'vi' ? 'Sản lượng đầu người tổng TB' : lang === 'ko' ? 'TTL 인당생산수 평균' : 'AVG TTL Per Capita',
              value: fmt1(avgTtl),
              unit: lang === 'vi' ? 'người' : lang === 'ko' ? '명' : 'prs',
              bg: 'var(--cyan-soft)', color: '#14b8a6',
            },
            {
              icon: '🎯',
              label: lang === 'vi' ? 'Mục tiêu' : lang === 'ko' ? '경영목표' : 'Mgmt. Target',
              value: String(targetPC),
              unit: lang === 'vi' ? 'sp/người' : lang === 'ko' ? '개/인' : 'pcs/cap',
              bg: 'var(--rose-soft)', color: '#f43f5e',
            },
          ].map((kpi, i) => (
            <div key={i} className="kpi-card">
              <div className="kpi-card-header">
                <div className="kpi-card-icon" style={{ background: kpi.bg, color: kpi.color }}>{kpi.icon}</div>
                <div className="kpi-card-label" style={{ fontSize: '12px', fontWeight: 700 }}>{kpi.label}</div>
              </div>
              <div className="kpi-card-value" style={{ fontSize: '30px' }}>
                {kpi.value}
                <span style={{ fontSize: '16px', color: 'var(--text-3)', marginLeft: '4px' }}>{kpi.unit}</span>
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
            <div className="panel-head" style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
              <h3 style={{ margin: 0 }}>{lang === 'vi' ? '인당생산수 - Sản lượng theo đầu người' : lang === 'ko' ? '인당생산수' : 'Per Capita Output'}</h3>
              <span style={{ fontSize: '13px', color: 'var(--text-3)', display: 'inline-flex', alignItems: 'center' }}>
                {lang === 'vi' ? 'TTL nhân lực x tỉ lệ ca (DAY/NIGHT)' : lang === 'ko' ? '근무인원 x 근무비율 (추정)' : 'Headcount x shift ratio (estimated)'}
              </span>
            </div>
            <div id={ids.current.c1} style={{ height: '270px' }} />
          </div>

          {/* Chart 2: 생산 현황 */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
              <h3 style={{ margin: 0 }}>{lang === 'vi' ? '생산 현황 - Tình hình sản xuất' : lang === 'ko' ? '생산 현황' : 'Production Status (estimated)'}</h3>
              <span style={{ fontSize: '13px', color: 'var(--text-3)', display: 'inline-flex', alignItems: 'center' }}>
                {lang === 'vi' ? `Nhân lực x tỉ lệ ca x ${targetPC} sp/người` : lang === 'ko' ? `근무인원 x 근무비율 x ${targetPC}개/인 (추정)` : `Headcount x ratio x ${targetPC} pcs/cap`}
              </span>
            </div>
            <div id={ids.current.c2} style={{ height: '270px' }} />
          </div>

          {/* Chart 3: 근무 인력 현황 */}
          <div className="panel">
            <div className="panel-head" style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
              <h3 style={{ margin: 0 }}>{lang === 'vi' ? '근무 인력 현황 - Tình hình nhân lực đi làm' : lang === 'ko' ? '근무 인력 현황' : 'Manpower Trend'}</h3>
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
            <div id={ids.current.c3} style={{ height: '270px' }} />
          </div>

          {/* ════ Chart: Sản xuất theo ca DAY / NIGHT / TTL ════════════════════════ */}
          <div className="panel">
            <div className="panel-head" style={{ marginBottom: '12px' }}>
              <h3 style={{ margin: 0 }}>
                {lang === 'vi'
                  ? `📊 Tình hình sản xuất — Plan / Actual / Yield (%)${modelFilter !== 'all' ? ` — Model: ${modelFilter}` : ' — Tất cả Model'}`
                  : lang === 'ko'
                  ? `📊 생산 현황 — 계획 / 실적 / 달성률 (%)${modelFilter !== 'all' ? ` — 모델: ${modelFilter}` : ' — 전체 모델'}`
                  : `📊 Production Status — Plan / Actual / Yield (%)${modelFilter !== 'all' ? ` — Model: ${modelFilter}` : ' — All Models'}`}
              </h3>
              <span style={{ fontSize: '11px', color: 'var(--text-3)', display: 'block', marginTop: '2px' }}>
                {lang === 'vi' ? 'Tối đa 7 giai đoạn gần nhất' : lang === 'ko' ? '최근 7기간' : 'Up to last 7 periods'}
              </span>
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
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '12px' }}>

              {/* Chart: DAY */}
              <div>
                <div style={{
                  fontSize: '12px', fontWeight: 700, color: textColor,
                  textAlign: 'center', marginBottom: '4px',
                  textTransform: 'uppercase', letterSpacing: '0.06em',
                }}>
                  DAY — Plan / Actual / Yield
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
                  NIGHT — Plan / Actual / Yield
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
                  TTL — Plan / Actual / Yield
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
              <div className="panel-head" style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <h3 style={{ margin: 0 }}>
                  {lang === 'vi'
                    ? 'Phân bố nhân lực theo Model — Giai đoạn gần nhất'
                    : lang === 'ko'
                    ? '모델별 인원 분포 — 최근 기간'
                    : 'Headcount by Model — Latest Period'}
                </h3>
                {chart4Info.hasLineData ? (
                  <span style={{ fontSize: '12px', color: '#facc15', display: 'inline-flex', alignItems: 'center' }}>
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
              <div id={ids.current.c4} style={{ minHeight: '450px' }} />
            </div>

            {/* Chart 5: Phân bố Model theo Giai đoạn (Spider) */}
            <div className="panel">
              <div className="panel-head" style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <h3 style={{ margin: 0 }}>
                  {lang === 'vi'
                    ? 'Phân bố Model theo Giai đoạn (Spider)'
                    : lang === 'ko'
                    ? '기간별 모델 분포 (Spider)'
                    : 'Model Distribution by Period (Spider)'}
                </h3>
                <span style={{ fontSize: '11px', color: 'var(--text-3)', display: 'inline-flex', alignItems: 'center' }}>
                  {lang === 'vi' ? 'Tối đa 6 giai đoạn gần nhất' : lang === 'ko' ? '최근 6기간' : 'Up to last 6 periods'}
                </span>
              </div>
              <div id={ids.current.c5} style={{ minHeight: '450px' }} />
            </div>

          </div>

        </div>
      )}
    </div>
  );
};

export default PerCapitaTab;

