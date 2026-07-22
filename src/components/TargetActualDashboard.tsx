import React, { useState, useEffect, useMemo, useRef } from 'react';
import type { DataRow } from '../types';
import { translations } from '../translations';
import { CustomSelect } from './CustomSelect';
import { usePagination } from '../hooks/usePagination';
import { chartTheme, getChartLayout } from './chartTheme';
import { NeonButton } from './NeonButton';

/* ═══════════════════════════════════════════════════════════════════════
 * FIX (header-legend-merge, EPCC) — chuyển legend "Mục tiêu(Target)/Thực
 * tế(Actual)/% Rate" và legend donut "By Customer" từ 1 HÀNG TRẮNG RIÊNG
 * (do Plotly tự vẽ bên trong chart-holder, `showlegend:true`) LÊN GỘP
 * CHUNG vào đúng 1 dòng với thanh tiêu đề màu (card-header-styled) —
 * đúng theo ảnh tham chiếu (6 khung biểu đồ demo) và spec màu/cỡ chữ của
 * DashboardTemplate.tsx (label bên phải header: cỡ chữ ~10.5-11px, không
 * in hoa, opacity nhẹ). Đồng thời BỎ dòng subtitle thừa ("Target vs
 * Actual" / "QTY / AMT" / "% Performance" / "By Customer") vì thông tin
 * này giờ đã được thể hiện trực quan hơn qua chính các item legend màu.
 *
 * EXPLORE: 6 khung ở Mục 2 (TargetActualDashboard) dùng chung class
 *   ".card-header-styled" (2 children: <span> tiêu đề trái, <span>
 *   subtitle phải — CSS global đã có sẵn display:flex;
 *   justify-content:space-between nên chỉ cần đổi NỘI DUNG bên phải,
 *   không cần đổi cấu trúc/class ngoài).
 * PLAN: viết 1 component <TALegendItem> tái sử dụng cho cả 6 khung, thay
 *   <span>subtitle</span> bằng 1 <div> chứa nhiều <TALegendItem>, tắt
 *   `showlegend` mặc định của Plotly (đang tự vẽ legend rời) cho các
 *   chart có legend, và mở rộng lại domain donut về full-width vì không
 *   còn cần chừa chỗ cho legend Plotly bên cạnh.
 * ═══════════════════════════════════════════════════════════════════════ */
const TALegendItem: React.FC<{ type: 'bar' | 'line' | 'dot'; color: string; label: string }> = ({ type, color, label }) => (
  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '10.5px', fontWeight: 500, whiteSpace: 'nowrap', textTransform: 'none' }}>
    {type === 'bar' && <span style={{ width: '9px', height: '9px', borderRadius: '2px', background: color, flexShrink: 0 }} />}
    {type === 'line' && <span style={{ width: '12px', height: 0, borderTop: `2px dashed ${color}`, flexShrink: 0 }} />}
    {type === 'dot' && <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: color, flexShrink: 0 }} />}
    {label}
  </span>
);

/** Bảng màu donut "By Customer" — PHẢI khớp nguyên văn mảng `marker.colors`
 *  dùng cho traceDonutQty/traceDonutAmt bên dưới, để legend header và màu
 *  lát cắt donut luôn đồng bộ dù danh sách khách hàng thay đổi theo dữ liệu. */
const DONUT_COLORS = ['#ea580c', '#0ea5e9', '#8b5cf6', '#b45309', '#0d9488', '#ec4899'];

/** Bảng màu 5 vòng cung "% Rate model" — PHẢI khớp nguyên văn với mảng dùng
 *  trong SVG gauge (Chart 4) để legend header và màu vòng cung luôn đồng bộ. */
const RING_COLORS = ['#2563eb', '#15803d', '#b91c1c', '#7c3aed', '#c2410c'];

/** Bảng màu 6 khung header "chart-grid" — lấy ĐÚNG NGUYÊN VĂN từ 6 màu tham
 *  chiếu trong DashboardTemplate.tsx (chartPanels[]), map theo ĐÚNG THỨ TỰ
 *  1→6 người dùng khoanh đỏ trong ảnh yêu cầu:
 *    1. Sản lượng Bán hàng Q'TY   → tím    #8b5cf6
 *    2. Doanh số Bán hàng AMT     → cyan   #06b6d4
 *    3. Doanh số bán hàng tháng   → cam    #f59e0b
 *    4. % Rate model              → cam-đỏ #f97316
 *    5. Khách hàng (Qty)          → xanh lá #10b981
 *    6. Khách hàng (AMT)          → xanh dương #3b82f6 */
const CHART_HEADER_THEME = [
  { accent: '#8b5cf6', bgLight: '#d6c6fc', bgDark: 'rgba(139,92,246,0.14)' },
  { accent: '#06b6d4', bgLight: '#a8e5f0', bgDark: 'rgba(6,182,212,0.14)' },
  { accent: '#f59e0b', bgLight: '#fcddaa', bgDark: 'rgba(255,255,255,0.05)' },
  { accent: '#f97316', bgLight: '#fdcead', bgDark: 'rgba(249,115,22,0.14)' },
  { accent: '#10b981', bgLight: '#abe7d3', bgDark: 'rgba(16,185,129,0.14)' },
  { accent: '#3b82f6', bgLight: '#bad3fc', bgDark: 'rgba(59,130,246,0.14)' },
] as const;

const PAGE_SIZES = [10, 25, 31, 50, 100];

/**
 * Returns the ISO week-of-month label ('W1'–'W4') for a given date.
 * W1: days 1–7, W2: days 8–14, W3: days 15–21, W4: day 22 – last day of month.
 * The last day is computed dynamically (new Date(year, month, 0).getDate())
 * so Feb 2026 (28 days) and Feb 2028 (29 days) are handled correctly.
 */
const getWeekOfMonth = (_year: number, _month: number, day: number): 'W1' | 'W2' | 'W3' | 'W4' | 'W5' => {
  const idx = Math.ceil(day / 7);
  return `W${idx}` as any;
};

/**
 * Build a stable, sortable week key: "YYYY-MM-WN"
 * e.g. "2026-06-W1", "2026-07-W4"
 * This prevents cross-month / cross-year collisions when grouping.
 */
const buildWeekKey = (year: number, monthNum: number, week: string): string =>
  `${year}-${String(monthNum).padStart(2, '0')}-${week}`;

/**
 * Normalizes an OIS division Type string for QTY Target / Actual matching.
 *
 * Problem: models use inconsistent Type naming for the same concept:
 *   • "Plan"                → QTY Target (single plan)
 *   • "Plan 1", "Plan 2"   → QTY Target (two sub-plans, must be summed)
 *   • "Plan S1", "Plan S2" → QTY Target (shift-based sub-plans)
 *   • "Final Sales"         → QTY Actual
 *
 * Rules applied (case-insensitive):
 *  1. Any string whose lowercased value starts with "acc" is EXCLUDED (accumulative / cumulative rows).
 *  2. Strip trailing sub-plan suffixes: " 1", " 2", " S1", " S2" (and no-space variants "s1","s2","1","2").
 *  3. Match the cleaned base against "plan" → returns 'plan'.
 *  4. Match the cleaned base against "final sales" variants → returns 'final_sales'.
 *  5. Anything else → returns null (caller should skip the row).
 */
const normalizeOisType = (typeStr: string): 'plan' | 'final_sales' | 'actual' | null => {
  const s = typeStr.trim().toLowerCase();

  // Guard: exclude cumulative rows (e.g. "Acc Plan", "ACC.Actual", "Acc Actual S1")
  if (s.startsWith('acc')) return null;

  // Strip trailing sub-plan variant suffixes (space-separated or run-on):
  //   " s1" | " s2" | " 1" | " 2"  — case-insensitive, optional spaces
  const base = s.replace(/\s+(?:s\d+|\d+)$/i, '').trim();

  if (base === 'plan') return 'plan';
  if (base === 'final sales' || base === 'final_sales' || base === 'finalsales') return 'final_sales';
  if (base === 'actual') return 'actual';

  return null; // unrecognised → caller skips the row
};

/**
 * Safe local timezone YYYY-MM-DD string parser
 */
const parseYYYYMMDD = (str: string): Date | null => {
  if (!str) return null;
  const parts = str.split('-');
  if (parts.length === 3) {
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const d = parseInt(parts[2], 10);
    if (!isNaN(y) && !isNaN(m) && !isNaN(d)) {
      return new Date(y, m - 1, d);
    }
  }
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
};

/**
 * Returns the first day of the given month/year
 */
const getDefaultStartDate = (viewMonth: number, viewYear: number): Date => {
  return new Date(viewYear, viewMonth - 1, 1);
};

/**
 * Formats a Date object to YYYY-MM-DD string
 */
const formatDateToYYYYMMDD = (date: Date): string => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

/**
 * Formats the report title dynamically based on the date range
 */
const formatReportTitle = (prefix: string, startDate: Date | null, endDate: Date | null, lang: 'vi' | 'en' | 'ko'): string => {
  if (!startDate || !endDate) {
    return prefix;
  }

  const startYear = startDate.getFullYear();
  const endYear = endDate.getFullYear();
  const startMonth = startDate.getMonth() + 1;
  const endMonth = endDate.getMonth() + 1;

  const formatDateStr = (date: Date, l: string) => {
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const y = date.getFullYear();
    if (l === 'ko') {
      return `${y}/${m}/${d}`;
    }
    return `${d}/${m}/${y}`;
  };

  if (startYear !== endYear) {
    return `${prefix} (${formatDateStr(startDate, lang)} - ${formatDateStr(endDate, lang)})`;
  }

  // Same year
  // FIX (duplicate-thang-word, EPCC): nếu `prefix` (tiêu đề gốc, VD "DOANH
  // SỐ BÁN HÀNG THÁNG") đã sẵn có chữ "Tháng"/"THÁNG" ở cuối, hàm không
  // được nối thêm chữ "Tháng" lần 2 nữa — chỉ nối số tháng/năm, tránh lặp
  // "...THÁNG Tháng 7/2026" như ảnh phản hồi.
  const prefixEndsWithThang = /tháng\s*$/i.test(prefix.trim());
  if (startMonth === endMonth) {
    if (lang === 'vi') {
      return prefixEndsWithThang
        ? `${prefix} ${startMonth}/${startYear}`
        : `${prefix} Tháng ${startMonth}/${startYear}`;
    } else if (lang === 'ko') {
      return `${prefix} ${startYear}년 ${startMonth}월`;
    } else {
      const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      return `${prefix} - ${monthNames[startMonth - 1]} ${startYear}`;
    }
  } else {
    if (lang === 'vi') {
      return prefixEndsWithThang
        ? `${prefix} ${startMonth}-${endMonth}/${startYear}`
        : `${prefix} Tháng ${startMonth}-${endMonth}/${startYear}`;
    } else if (lang === 'ko') {
      return `${prefix} ${startYear}년 ${startMonth}-${endMonth}월`;
    } else {
      return `${prefix} - Months ${startMonth}-${endMonth}/${startYear}`;
    }
  }
};



const parseToDateObject = (row: DataRow, dateColName: string): Date | null => {
  const v = row[dateColName];
  if (v instanceof Date) return v;
  if (v === null || v === undefined || v === '') return null;

  if (typeof v === 'number') {
    if (v >= 30000 && v <= 60000) {
      const utc_days = Math.floor(v - 25569);
      const utc_value = utc_days * 86400;
      return new Date(utc_value * 1000);
    }
    if (v >= 1970 && v <= 2100) {
      return new Date(v, 0, 1);
    }
  }

  const s = String(v).trim();
  if (s === 'TTL') return null;

  if (s.includes('/')) {
    const parts = s.split('/');
    if (parts.length === 2) {
      const mNum = parseInt(parts[0], 10);
      const dNum = parseInt(parts[1], 10);
      let year = 2026;
      const yearKey = Object.keys(row).find(k => k.toLowerCase().includes('year') || k.toLowerCase().includes('nam'));
      if (yearKey && row[yearKey]) {
        year = parseInt(String(row[yearKey]), 10) || 2026;
      }
      if (mNum >= 1 && mNum <= 12 && dNum >= 1 && dNum <= 31) {
        return new Date(year, mNum - 1, dNum);
      }
    } else if (parts.length === 3) {
      const p0 = parseInt(parts[0], 10);
      const p1 = parseInt(parts[1], 10);
      const p2 = parseInt(parts[2], 10);
      if (p0 > 1000) {
        return new Date(p0, p1 - 1, p2);
      } else if (p2 > 1000) {
        return new Date(p2, p1 - 1, p0);
      }
    }
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }

  const monthsShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const monthsLong = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
  
  let monthIdx = -1;
  const sUpper = s.toUpperCase();
  
  const shortIdx = monthsShort.indexOf(sUpper);
  if (shortIdx !== -1) monthIdx = shortIdx;
  
  const longIdx = monthsLong.indexOf(sUpper);
  if (longIdx !== -1) monthIdx = longIdx;
  
  const parsedInt = parseInt(s, 10);
  if (!isNaN(parsedInt) && parsedInt >= 1 && parsedInt <= 12) {
    monthIdx = parsedInt - 1;
  }

  if (monthIdx !== -1) {
    let year = 2026;
    const yearKey = Object.keys(row).find(k => k.toLowerCase().includes('year') || k.toLowerCase().includes('nam'));
    if (yearKey && row[yearKey]) {
      year = parseInt(String(row[yearKey]), 10) || 2026;
    }
    return new Date(year, monthIdx, 1);
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
};

interface TargetActualDashboardProps {
  rows: DataRow[];
  theme: 'dark' | 'light';
  onToggleTheme: () => void;
  lang: 'vi' | 'en' | 'ko';
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
  onFileSelected: (file: File, wb: any) => void;
}

export const TargetActualDashboard: React.FC<TargetActualDashboardProps> = ({
  rows,
  theme,
  onToggleTheme: _onToggleTheme,
  lang,
  setLang: _setLang,
  onFileSelected
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const t = translations[lang];
  const isLightMode = theme === 'light';

  // FIX (chart-header-color-match-template, EPCC): style header dùng
  // chung cho 6 khung "chart-grid" — lấy nền sáng (bgLight/bgDark) + màu
  // chữ tối tương phản từ CHART_HEADER_THEME (đúng màu tham chiếu
  // DashboardTemplate), thay cho nền gradient đậm + chữ trắng cũ
  // (green-style/purple-style/blue-style). idx = vị trí 0-5 tương ứng
  // đúng thứ tự 1-6 người dùng khoanh đỏ.
  const chartHeaderStyle = (idx: number): React.CSSProperties => {
    const c = CHART_HEADER_THEME[idx % CHART_HEADER_THEME.length];
    return {
      background: isLightMode ? c.bgLight : c.bgDark,
      color: isLightMode ? '#1f2937' : 'var(--text-1)',
      borderLeft: `4px solid ${c.accent}`,
      borderBottom: 'none',
    };
  };

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

  // Mapping States
  const [mapModel, setMapModel] = useState<string>('');
  const [mapCustomer, setMapCustomer] = useState<string>('');
  const [mapMonth, setMapMonth] = useState<string>('');
  const [mapValue, setMapValue] = useState<string>('');
  const [mapType, setMapType] = useState<string>(''); // For Target/Actual type (e.g. Type1)
  const [targetMatch] = useState<string>('Plan');
  const [actualMatch] = useState<string>('Actual');

  // For Qty/Amt distinction
  const [mapMetricMode] = useState<'both' | 'qty_only' | 'amt_only'>('both');
  const [mapMetricCol, setMapMetricCol] = useState<string>('');
  const [qtyMatch] = useState<string>('SUB1');
  const [amtMatch] = useState<string>('SUB2');

  // Auto-detect columns on mount/data load
  useEffect(() => {
    if (rows.length > 0) {
      const keys = Object.keys(rows[0]);
      const findKey = (names: string[]) => {
        for (const n of names) {
          const found = keys.find(k => k.toLowerCase().replace(/[\s\-_'`]/g, '') === n.toLowerCase().replace(/[\s\-_'`]/g, ''));
          if (found) return found;
        }
        return '';
      };

      setMapModel(findKey(['model', 'so', '모델']) || keys[0]);
      setMapCustomer(findKey(['customer', 'custom', 'type2', 'khachhang', '고객']) || keys[1] || keys[0]);
      setMapMonth(findKey(['month', 'date', 'thang', 'ngay', '월']) || keys[2] || keys[0]);
      setMapValue(findKey(['value', 'val', 'soluong', 'giatri', '값']) || keys[keys.length - 1]);
      
      const foundType = keys.find(k => k.toLowerCase().includes('type1') || k.toLowerCase().includes('type') || k.toLowerCase().includes('phanloai') || k.toLowerCase().includes('구분'));
      setMapType(foundType || keys[0]);
      
      const foundDiv = keys.find(k => k.toLowerCase().includes('division') || k.toLowerCase().includes('bophan') || k.toLowerCase().includes('loai'));
      setMapMetricCol(foundDiv || '');
    }
  }, [rows]);

  // 1. Process and normalize uploaded rows
  const normalizedRows = useMemo(() => {
    if (rows.length === 0) return [];

    const firstRow = rows[0];
    const keys = Object.keys(firstRow);
    const keysLower = keys.map(k => k.toLowerCase().trim());

    // 1.1 Detect if it's the Target_Actual_Sample.xlsx format
    const hasQtyTarget = keysLower.includes('qty target') || keysLower.includes('qty_target');
    const hasQtyActual = keysLower.includes('qty actual') || keysLower.includes('qty_actual');
    const hasAmtTarget = keysLower.includes('amt target') || keysLower.includes('amt_target');
    const hasAmtActual = keysLower.includes('amt actual') || keysLower.includes('amt_actual');

    if (hasQtyTarget && hasQtyActual && hasAmtTarget && hasAmtActual) {
      const keyModel = keys.find(k => k.toLowerCase().replace(/[\s\-_'`]/g, '') === 'model') || keys[0];
      const keyCustomer = keys.find(k => k.toLowerCase().replace(/[\s\-_'`]/g, '') === 'customer') || keys[1];
      const keyMonth = keys.find(k => k.toLowerCase().replace(/[\s\-_'`]/g, '') === 'month') || keys[2];
      
      const keyQtyTarget = keys.find(k => k.toLowerCase().trim() === 'qty target' || k.toLowerCase().trim() === 'qty_target') || 'Qty Target';
      const keyQtyActual = keys.find(k => k.toLowerCase().trim() === 'qty actual' || k.toLowerCase().trim() === 'qty_actual') || 'Qty Actual';
      const keyAmtTarget = keys.find(k => k.toLowerCase().trim() === 'amt target' || k.toLowerCase().trim() === 'amt_target') || 'Amt Target';
      const keyAmtActual = keys.find(k => k.toLowerCase().trim() === 'amt actual' || k.toLowerCase().trim() === 'amt_actual') || 'Amt Actual';

      return rows.map(r => {
        const model = String(r[keyModel] || '').trim();
        const customer = String(r[keyCustomer] || '').trim();
        const parsedDate = parseToDateObject(r, keyMonth);
        if (!parsedDate) return null;
        
        const monthsShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
        const month = monthsShort[parsedDate.getMonth()];

        const year = parsedDate.getFullYear();
        const monthNum = parsedDate.getMonth() + 1;
        const dayOfMonth = parsedDate.getDate();
        const week = getWeekOfMonth(year, monthNum, dayOfMonth);

        return {
          model,
          customer,
          date: parsedDate,
          month,
          qtyTarget: Number(r[keyQtyTarget]) || 0,
          qtyActual: Number(r[keyQtyActual]) || 0,
          amtTarget: Number(r[keyAmtTarget]) || 0,
          amtActual: Number(r[keyAmtActual]) || 0,
          monthNum,
          year,
          week
        };
      }).filter((r): r is NonNullable<typeof r> => r !== null && r.model !== '');
    }

    // Pre-scan: Identify which months/years/models have daily detail rows (e.g., date strings containing a slash '/')
    const targetMonthsWithDaily = new Set<string>();
    const actualMonthsWithDaily = new Set<string>();
    const amtTargetMonthsWithDaily = new Set<string>();
    const amtActualMonthsWithDaily = new Set<string>();
    const monthsShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    
    rows.forEach(r => {
      const dateKey = Object.keys(r).find(k => {
        const kl = k.trim().toLowerCase();
        return kl === 'date' || kl === 'month';
      });
      const dateVal = r.date ?? (r as any).Date ?? r.month ?? (r as any).Month ?? (dateKey ? r[dateKey] : undefined);
      if (!dateVal) return;
      const dateStr = String(dateVal).trim();
      if (dateStr.includes('/')) {
        const parts = dateStr.split('/');
        if (parts.length >= 2) {
          const mNum = parseInt(parts[0], 10);
          if (mNum >= 1 && mNum <= 12) {
            const monthAbbr = monthsShort[mNum - 1];
            let year = 2026;
            const yearKey = Object.keys(r).find(k => k.toLowerCase().includes('year') || k.toLowerCase().includes('nam'));
            if (yearKey && (r as any)[yearKey]) {
              year = parseInt(String((r as any)[yearKey]), 10) || 2026;
            } else if (parts.length === 3) {
              const p2 = parseInt(parts[2], 10);
              if (p2 > 1000) year = p2;
              else {
                const p0 = parseInt(parts[0], 10);
                if (p0 > 1000) year = p0;
              }
            }
            
            const modelKey = Object.keys(r).find(k => k.trim().toLowerCase() === 'model');
            const model = String(modelKey ? r[modelKey] : (r.model ?? (r as any).Model ?? '')).trim();
            
            const divisionKey = Object.keys(r).find(k => k.trim().toLowerCase() === 'division');
            const division = String(divisionKey ? r[divisionKey] : (r.division ?? (r as any).Division ?? '')).trim().toUpperCase();
            
            const typeKey = Object.keys(r).find(k => {
              const kl = k.trim().toLowerCase();
              return kl === 'type' || kl === 'type1' || kl === 'phanloai' || kl === '구분';
            });
            const typeStr = String(typeKey ? r[typeKey] : (r.type ?? (r as any).Type ?? r.type1 ?? (r as any).Type1 ?? '')).trim().toLowerCase();
            const normType = normalizeOisType(typeStr);
            const isAmtDivision = division.startsWith('AMT');
            
            const isTarget = (division === 'SHIPMENT' || division === 'SUB1' || division === 'SUB2' || division === 'OIS') && normType === 'plan';
            const isActual = (division === 'OIS' && normType === 'final_sales') || 
                             ((division === 'SHIPMENT' || division === 'SUB1' || division === 'SUB2') && normType === 'actual');
            // FIX: nhánh riêng cho AMT — Plan = target, Actual = actual (AMT không phân biệt OIS/Shipment)
            const isAmtTarget = isAmtDivision && normType === 'plan';
            const isAmtActual = isAmtDivision && normType === 'actual';
            
            const key = `${monthAbbr}|${year}|${model}`;
            if (isTarget) {
              targetMonthsWithDaily.add(key);
            }
            if (isActual) {
              actualMonthsWithDaily.add(key);
            }
            if (isAmtTarget) {
              amtTargetMonthsWithDaily.add(key);
            }
            if (isAmtActual) {
              amtActualMonthsWithDaily.add(key);
            }
          }
        }
      }
    });




    // 1.2 Detect if it's the Test 1.xlsx format (having MODEL, DIVISION, Type1/Type, Type2/Custom, DATE/Date, Value)
    const keysLower2 = keys.map(k => k.toLowerCase().trim());
    const hasModelKey = keysLower2.includes('model');
    const hasDivisionKey = keysLower2.includes('division');
    const hasTypeCol = keysLower2.includes('type1') || keysLower2.includes('type');
    const hasCustCol = keysLower2.includes('type2') || keysLower2.includes('custom');
    const hasDateKey = keysLower2.includes('date');
    const hasValueKey = keysLower2.includes('value');

    if (hasModelKey && hasDivisionKey && hasTypeCol && hasCustCol && hasDateKey && hasValueKey) {
      const modelCol = keys.find(k => k.toLowerCase().trim() === 'model')!;
      const divisionCol = keys.find(k => k.toLowerCase().trim() === 'division')!;
      const typeCol = keys.find(k => k.toLowerCase().trim() === 'type1' || k.toLowerCase().trim() === 'type')!;
      const custCol = keys.find(k => k.toLowerCase().trim() === 'type2' || k.toLowerCase().trim() === 'custom')!;
      const dateCol = keys.find(k => k.toLowerCase().trim() === 'date')!;
      const valueCol = keys.find(k => k.toLowerCase().trim() === 'value')!;

      const REAL_CUSTOMERS = new Set(['GAOXIN', 'SEMV', 'SUNNY', 'LGIT', 'Q-TECH']);

      return rows.map(r => {
        const dateRawStr = String(r[dateCol] || '').trim().toUpperCase();
        if (dateRawStr === 'TTL') return null;

        const model = String(r[modelCol] || '').trim();
        const parsedDate = parseToDateObject(r, dateCol);
        if (!parsedDate) return null;
        
        const month = monthsShort[parsedDate.getMonth()];
        const year = parsedDate.getFullYear();

        const div = String(r[divisionCol] || '').trim();
        const divUpper = div.toUpperCase();
        const type1 = String(r[typeCol] || '').trim();
        const type1Lower = type1.toLowerCase();
        const type2 = String(r[custCol] || '').trim();
        const type2Upper = type2.toUpperCase();
        const val = Number(r[valueCol]) || 0;

        // Skip accumulated rows
        if (type1Lower.includes('acc') || type2.toLowerCase().includes('acc') || type1Lower === 'acc.') return null;

        const shipNorm = normalizeOisType(type1);

        // Skip this monthly summary row if daily rows are already present for the same month/year/model
        const MONTHS_SET = new Set(monthsShort);
        if (MONTHS_SET.has(dateRawStr)) {
          const targetKey = `${dateRawStr}|${year}|${model}`;
          const actualKey = `${dateRawStr}|${year}|${model}`;
          
          const isTargetRow = (divUpper === 'SHIPMENT' || divUpper === 'SUB1' || divUpper === 'SUB2' || divUpper === 'OIS') && shipNorm === 'plan';
          const isActualRow = (divUpper === 'OIS' && shipNorm === 'final_sales') ||
                              ((divUpper === 'SHIPMENT' || divUpper === 'SUB1' || divUpper === 'SUB2') && shipNorm === 'actual');
          // FIX: check riêng cho AMT — trước đây monthly-summary của AMT không bao giờ bị loại
          // vì isTargetRow/isActualRow chỉ tính QTY division, khiến AMT luôn bị cộng đôi khi có breakdown ngày.
          const isAmtTargetRow = divUpper.startsWith('AMT') && shipNorm === 'plan';
          const isAmtActualRow = divUpper.startsWith('AMT') && shipNorm === 'actual';



          if (isTargetRow && targetMonthsWithDaily.has(targetKey)) {
            return null;
          }
          if (isActualRow && actualMonthsWithDaily.has(actualKey)) {
            return null;
          }
          if (isAmtTargetRow && amtTargetMonthsWithDaily.has(targetKey)) {
            return null;
          }
          if (isAmtActualRow && amtActualMonthsWithDaily.has(actualKey)) {
            return null;
          }
        }

        let customer = 'Unknown';
        let qtyTarget = 0;
        let qtyActual = 0;
        let amtTarget = 0;
        let amtActual = 0;

        const isShift = /\s+S\d+/i.test(model);

        if (divUpper === 'SHIPMENT' || divUpper === 'SUB1' || divUpper === 'SUB2') {
          // QTY TARGET from Shipment (covers all models except SO2701)
          if (!REAL_CUSTOMERS.has(type2Upper)) return null;
          customer = type2;
          if (shipNorm === 'plan') {
            if (model === 'SO2701') return null; // SO2701 uses OIS Plan 1 & Plan 2 instead
            qtyTarget = val;
          } else if (shipNorm === 'actual' && isShift) {
            qtyActual = val;
          } else {
            // Non-plan/non-shift actual Shipment rows → skip
            return null;
          }

        } else if (divUpper === 'OIS') {
          // QTY ACTUAL from OIS
          if (!REAL_CUSTOMERS.has(type2Upper)) return null;
          customer = type2;
          if (shipNorm === 'final_sales') {
            if (isShift) return null; // Shift models get Actual from Shipment division
            qtyActual = val;
          } else if (shipNorm === 'plan' && model === 'SO2701') {
            qtyTarget = val;
          } else {
            return null;
          }

        } else if (divUpper === 'AMT K$' || divUpper === 'AMT_K$' || divUpper === 'AMTK$'
                   || divUpper.startsWith('AMT')) {
          // AMT TARGET and AMT ACTUAL
          if (!REAL_CUSTOMERS.has(type2Upper)) return null;
          customer = type2;
          if (shipNorm === 'plan') {
            amtTarget = val;
          } else if (shipNorm === 'actual') {
            amtActual = val;
          } else {
            return null;
          }

        } else {
          return null; // unrecognised division → skip
        }

        if (qtyTarget === 0 && qtyActual === 0 && amtTarget === 0 && amtActual === 0) {
          return null;
        }

        const monthNum = parsedDate.getMonth() + 1;
        const dayOfMonth = parsedDate.getDate();
        const week = getWeekOfMonth(year, monthNum, dayOfMonth);

        return {
          model,
          customer,
          date: parsedDate,
          month,
          qtyTarget,
          qtyActual,
          amtTarget,
          amtActual,
          monthNum,
          year,
          week
        };
      }).filter((r): r is NonNullable<typeof r> => r !== null && r.model !== '');
    }

    // 1.3 Supabase TargetActual format (được reconstruct từ nhiều dòng phân mảnh lưu trên cloud)
    const isSupabaseTargetActual = rows.some(r => r.source_tag === 'TargetActual' || r.origin === 'TargetActual');



    if (isSupabaseTargetActual) {
      const grouped: Record<string, any> = {};
      const MONTHS_SET = new Set(monthsShort);
      const REAL_CUSTOMERS = new Set(['GAOXIN', 'SEMV', 'SUNNY', 'LGIT', 'Q-TECH']);
      
      rows.forEach(r => {
        const model = String(r.model || r.Model || '').trim();
        if (!model) return;
        
        const customer = String(r.customer || r.Customer || 'Unknown').trim();
        const customerUpper = customer.toUpperCase();
        if (!REAL_CUSTOMERS.has(customerUpper)) return; // ignore non-real customers
        
        const monthRaw = String(r.month || r.Month || 'JAN').trim();
        const monthUpper = monthRaw.toUpperCase();
        const year = Number(r.year || r.Year) || 2026;
        
        const div = String(r.division || r.Division || '').toUpperCase();
        const typeStr = String(r.type || r.Type || '').toLowerCase();
        const normType = normalizeOisType(typeStr);

        // Skip monthly summary row if we already have daily breakdown for that month/year/model
        if (MONTHS_SET.has(monthUpper)) {
          const targetKey = `${monthUpper}|${year}|${model}`;
          const actualKey = `${monthUpper}|${year}|${model}`;
          
          const isTargetRow = (div === 'SHIPMENT' || div === 'SUB1' || div === 'SUB2' || div === 'OIS') && normType === 'plan';
          const isActualRow = (div === 'OIS' && normType === 'final_sales') ||
                              ((div === 'SHIPMENT' || div === 'SUB1' || div === 'SUB2') && normType === 'actual');
          // FIX: check riêng cho AMT (xem giải thích ở nhánh Test1.xlsx phía trên)
          const isAmtTargetRow = div.startsWith('AMT') && normType === 'plan';
          const isAmtActualRow = div.startsWith('AMT') && normType === 'actual';



          if (isTargetRow && targetMonthsWithDaily.has(targetKey)) {
            return;
          }
          if (isActualRow && actualMonthsWithDaily.has(actualKey)) {
            return;
          }
          if (isAmtTargetRow && amtTargetMonthsWithDaily.has(targetKey)) {
            return;
          }
          if (isAmtActualRow && amtActualMonthsWithDaily.has(actualKey)) {
            return;
          }
        }
        
        const parsedDate = parseToDateObject(r, 'month');
        if (!parsedDate) return;
        
        const monthNum = parsedDate.getMonth() + 1;
        const dayOfMonth = parsedDate.getDate();
        const week = getWeekOfMonth(year, monthNum, dayOfMonth);
        
        const key = `${model}|${customer}|${monthRaw}|${year}`;
        if (!grouped[key]) {
          grouped[key] = {
            model,
            customer,
            date: parsedDate,
            month: monthsShort[parsedDate.getMonth()],
            qtyTarget: 0,
            qtyActual: 0,
            amtTarget: 0,
            amtActual: 0,
            monthNum,
            year,
            week
          };
        }
        
        const val = Number(r.value || r.Value) || 0;
        
        const isShift = /\s+S\d+/i.test(model);
        
        if (div === 'SHIPMENT' || div === 'SUB1' || div === 'SUB2') {
          if (normType === 'plan') {
            if (model === 'SO2701') return; // SO2701 uses OIS Plan 1 & Plan 2 instead
            grouped[key].qtyTarget += val;
          } else if (normType === 'actual' && isShift) {
            grouped[key].qtyActual += val;
          }
        } else if (div === 'OIS') {
          if (normType === 'final_sales') {
            if (isShift) return; // Shift models get Actual from Shipment division
            grouped[key].qtyActual += val;
          } else if (normType === 'plan' && model === 'SO2701') {
            grouped[key].qtyTarget += val;
          }
        } else if (div.startsWith('AMT')) {
          if (normType === 'plan') {
            grouped[key].amtTarget += val;
          } else if (normType === 'actual') {
            grouped[key].amtActual += val;
          }
        }
      });
      
      return Object.values(grouped);
    }

    // 1.4 Fallback to default user-defined mapping
    if (!mapModel || !mapValue) return [];

    return rows.map(r => {
      const model = String(r[mapModel] || '').trim();
      const customer = mapCustomer ? String(r[mapCustomer] || '').trim() : 'Unknown';
      
      const parsedDate = mapMonth ? parseToDateObject(r, mapMonth) : null;
      if (!parsedDate) return null;

      const monthsShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const month = monthsShort[parsedDate.getMonth()];

      const rawVal = Number(r[mapValue]) || 0;
      
      // Determine if Target or Actual
      const typeVal = mapType ? String(r[mapType] || '').trim().toLowerCase() : '';
      const isTarget = typeVal.includes(targetMatch.toLowerCase());
      const isActual = actualMatch ? typeVal.includes(actualMatch.toLowerCase()) : !isTarget;

      // Determine metric type (Qty vs Amt)
      let isQty = true;
      let isAmt = true;

      if (mapMetricMode === 'qty_only') {
        isAmt = false;
      } else if (mapMetricMode === 'amt_only') {
        isQty = false;
      } else if (mapMetricCol) {
        const metricVal = String(r[mapMetricCol] || '').trim().toLowerCase();
        isQty = metricVal.includes(qtyMatch.toLowerCase());
        isAmt = metricVal.includes(amtMatch.toLowerCase());
      }

      const qtyTarget = (isQty && isTarget) ? rawVal : 0;
      const qtyActual = (isQty && isActual) ? rawVal : 0;
      
      let amtTarget = (isAmt && isTarget) ? rawVal : 0;
      let amtActual = (isAmt && isActual) ? rawVal : 0;

      // Fallback mirror in case of qty_only or amt_only mode to fill both graphs
      if (mapMetricMode === 'qty_only') {
        amtTarget = qtyTarget * 1.8;
        amtActual = qtyActual * 1.8;
      }

      const year = parsedDate.getFullYear();
      const monthNum = parsedDate.getMonth() + 1;
      const dayOfMonth = parsedDate.getDate();
      const week = getWeekOfMonth(year, monthNum, dayOfMonth);

      // Skip this monthly summary row if daily rows are already present for the same month/year/model
      const MONTHS_SET = new Set(monthsShort);
      const dateRawVal = mapMonth ? r[mapMonth] : undefined;
      const dateRawStr = String(dateRawVal ?? '').trim().toUpperCase();
      if (MONTHS_SET.has(dateRawStr)) {
        const targetKey = `${dateRawStr}|${year}|${model}`;
        const actualKey = `${dateRawStr}|${year}|${model}`;
        if (isTarget && (targetMonthsWithDaily.has(targetKey) || amtTargetMonthsWithDaily.has(targetKey))) {
          return null;
        }
        if (isActual && (actualMonthsWithDaily.has(actualKey) || amtActualMonthsWithDaily.has(actualKey))) {
          return null;
        }
      }

      return {
        model,
        customer,
        date: parsedDate,
        month,
        qtyTarget,
        qtyActual,
        amtTarget,
        amtActual,
        monthNum,
        year,
        week
      };
    }).filter((r): r is NonNullable<typeof r> => r !== null && r.model !== '');
  }, [rows, mapModel, mapCustomer, mapMonth, mapValue, mapType, targetMatch, actualMatch, mapMetricMode, mapMetricCol, qtyMatch, amtMatch]);

  /**
   * "Dữ liệu cập nhật đến" — thời điểm dữ liệu MỚI NHẤT thực sự có trong
   * database (normalizedRows), KHÔNG phải giờ hệ thống/đồng hồ máy người
   * dùng. Ưu tiên lấy ngày lớn nhất trong các dòng có phát sinh Actual
   * (qtyActual > 0 hoặc amtActual > 0) vì đó là mốc dữ liệu thực tế đã
   * được ghi nhận lên hệ thống; nếu chưa có dòng Actual nào thì fallback
   * lấy ngày lớn nhất trong toàn bộ dữ liệu hiện có.
   */
  const lastDataUpdateDate = useMemo(() => {
    if (normalizedRows.length === 0) return null;

    let maxActualDate: Date | null = null;
    let maxAnyDate: Date | null = null;

    for (const r of normalizedRows) {
      if (!r.date) continue;
      if (!maxAnyDate || r.date.getTime() > maxAnyDate.getTime()) maxAnyDate = r.date;
      if ((r.qtyActual > 0 || r.amtActual > 0) && (!maxActualDate || r.date.getTime() > maxActualDate.getTime())) {
        maxActualDate = r.date;
      }
    }

    return maxActualDate ?? maxAnyDate;
  }, [normalizedRows]);

  const formattedLastDataUpdate = lastDataUpdateDate
    ? lastDataUpdateDate.toLocaleDateString(lang === 'vi' ? 'vi-VN' : lang === 'ko' ? 'ko-KR' : 'en-US', {
        day: '2-digit', month: '2-digit', year: 'numeric'
      })
    : (lang === 'vi' ? 'Chưa có dữ liệu' : lang === 'ko' ? '데이터 없음' : 'No data');

  // 2. State values
  const [activeTab, setActiveTab] = useState<'summary' | 'merged'>('summary');
  // FIX (EPCC bottom=0): thay vì maxHeight cố định '62vh' (luôn để hở một
  // khoảng trắng cố định phía dưới bảng, bất kể cửa sổ cao/thấp), đo TRỰC
  // TIẾP vị trí thật của khung cuộn so với đáy viewport bằng
  // getBoundingClientRect(), rồi set maxHeight = phần còn lại tới đáy màn
  // hình (trừ 1 khoảng đệm nhỏ để không dính sát viền dưới cùng của trình
  // duyệt). Cách này KHÔNG phụ thuộc cấu trúc/height của DOM cha (App.tsx,
  // Sidebar...) — chỉ dựa vào vị trí thực tế của chính element, nên không
  // lặp lại lỗi "khoảng trắng thừa" từng gặp khi ép flex:1/100dvh trước đây.
  const tableScrollRef = useRef<HTMLDivElement | null>(null);
  const paginationBarRef = useRef<HTMLDivElement | null>(null);
  const [tableScrollMaxHeight, setTableScrollMaxHeight] = useState<number>(
    typeof window !== 'undefined' ? Math.max(240, window.innerHeight * 0.62) : 480
  );
  const [selectedDateStart, setSelectedDateStart] = useState<string>('');
  const [selectedDateEnd, setSelectedDateEnd] = useState<string>('');
  const [selectedCustomer, setSelectedCustomer] = useState<string>('');
  const [selectedModel, setSelectedModel] = useState<string>('');
  const [selectedMonth, setSelectedMonth] = useState<string>('');
  const [qtyTargetFilter, setQtyTargetFilter] = useState<string>('all');
  const [qtyActualFilter, setQtyActualFilter] = useState<string>('all');
  const [amtTargetFilter, setAmtTargetFilter] = useState<string>('all');
  const [amtActualFilter, setAmtActualFilter] = useState<string>('all');
  const [sortField, setSortField] = useState<string>('');
  const [sortAsc, setSortAsc] = useState<boolean>(true);
  const [currentTime, setCurrentTime] = useState<Date>(new Date());
  const [gaugeData, setGaugeData] = useState<Array<{ name: string; rate: number }>>([]);
  // FIX (header-legend-merge, EPCC): legend donut "Khách hàng (Qty/AMT)"
  // chuyển vào header màu — cần lưu vào state vì header (JSX) render TÁCH
  // RỜI khỏi useEffect xây chart (nơi duy nhất biết danh sách khách hàng
  // thật + màu tương ứng theo DONUT_COLORS).
  const [customQtyLegend, setCustomQtyLegend] = useState<Array<{ label: string; color: string }>>([]);
  const [customAmtLegend, setCustomAmtLegend] = useState<Array<{ label: string; color: string }>>([]);
  const [gaugeOverallRate, setGaugeOverallRate] = useState(0);
  const [viewMode, setViewMode] = useState<'day' | 'week' | 'month'>('day');
  const [selectedWeek, setSelectedWeek] = useState<string>('');

  // Auto-set initial date bounds based on available data
  const dateBounds = useMemo(() => {
    const dates = normalizedRows.map(r => r.date.getTime());
    if (dates.length === 0) return { min: null, max: null };
    return {
      min: new Date(Math.min(...dates)),
      max: new Date(Math.max(...dates))
    };
  }, [normalizedRows]);

  const defaultStartDateStr = useMemo(() => {
    if (!dateBounds.max) return '';
    const now = new Date();
    // Ưu tiên "hôm nay" làm mốc mặc định nếu nó nằm trong phạm vi dữ liệu hiện có.
    // Nếu dữ liệu không phủ tới ngày hiện tại (VD chỉ có data quá khứ hoặc chỉ có
    // kế hoạch tương lai) → fallback về tháng của ngày mới nhất trong data như cũ.
    const useToday = !!dateBounds.min && now >= dateBounds.min && now <= dateBounds.max;
    const refDate = useToday ? now : dateBounds.max;
    const y = refDate.getFullYear();
    const m = refDate.getMonth() + 1;
    return formatDateToYYYYMMDD(getDefaultStartDate(m, y));
  }, [dateBounds.max, dateBounds.min]);

  const defaultEndDateStr = useMemo(() => {
    if (!dateBounds.max) return '';
    const now = new Date();
    const useToday = !!dateBounds.min && now >= dateBounds.min && now <= dateBounds.max;
    const refDate = useToday ? now : dateBounds.max;
    const y = refDate.getFullYear();
    const m = refDate.getMonth() + 1;
    const lastDay = new Date(y, m, 0).getDate();
    return formatDateToYYYYMMDD(new Date(y, m - 1, lastDay));
  }, [dateBounds.max, dateBounds.min]);

  const selectedMonthDays = useMemo(() => {
    const dStr = selectedDateEnd || defaultEndDateStr;
    if (!dStr) return 30;
    const d = parseYYYYMMDD(dStr);
    if (!d) return 30;
    return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  }, [selectedDateEnd, defaultEndDateStr]);

  const hasW5 = selectedMonthDays >= 29;

  const showReset = !!(selectedCustomer || selectedModel || selectedWeek || selectedMonth ||
    (defaultStartDateStr && selectedDateStart !== defaultStartDateStr) ||
    (defaultEndDateStr && selectedDateEnd !== defaultEndDateStr));

  // Set default dates on load or when new data (Excel) is uploaded
  useEffect(() => {
    if (defaultStartDateStr && defaultEndDateStr) {
      setSelectedDateStart(defaultStartDateStr);
      setSelectedDateEnd(defaultEndDateStr);
    }
  }, [defaultStartDateStr, defaultEndDateStr]);


  /**
   * Snaps a given date to the first day of its month.
   */
  const snapToMonthStart = (d: Date): Date => new Date(d.getFullYear(), d.getMonth(), 1);

  /**
   * Snaps a given date to the last day of its month.
   */
  const snapToMonthEnd = (d: Date): Date => {
    const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
    return new Date(d.getFullYear(), d.getMonth(), lastDay);
  };

  /**
   * Whenever the user picks any date in either the Start or End date picker,
   * the whole selection is snapped to cover the FULL month of that date
   * (Start → day 1, End → last day of month).
   *
   * Why: Target/Plan rows only exist at monthly granularity (they're parsed
   * as day 1 of the month), while Actual rows exist per-day. If a partial
   * month range were allowed, the whole month's Target would still land on
   * day 1 and get counted in full, while Actual would only reflect the days
   * in range — making %Rate collapse artificially low. Forcing the filter to
   * always span a complete month keeps Target vs Actual comparable, and the
   * Day/Week view modes remain available to drill into that month's data.
   */
  const handleStartDateChange = (value: string) => {
    const d = parseYYYYMMDD(value);
    if (!d) { setSelectedDateStart(value); return; }
    setSelectedDateStart(formatDateToYYYYMMDD(snapToMonthStart(d)));
    setSelectedDateEnd(formatDateToYYYYMMDD(snapToMonthEnd(d)));
  };

  const handleEndDateChange = (value: string) => {
    const d = parseYYYYMMDD(value);
    if (!d) { setSelectedDateEnd(value); return; }
    setSelectedDateStart(formatDateToYYYYMMDD(snapToMonthStart(d)));
    setSelectedDateEnd(formatDateToYYYYMMDD(snapToMonthEnd(d)));
  };

  // Handle selectedMonth dropdown change (e.g. from bottom table filter)
  useEffect(() => {
    if (selectedMonth) {
      const matchingRow = normalizedRows.find(r => r.month === selectedMonth);
      const year = matchingRow ? matchingRow.year : (dateBounds.max ? dateBounds.max.getFullYear() : 2026);
      const monthsShort = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const monthIdx = monthsShort.indexOf(selectedMonth.toUpperCase());
      if (monthIdx !== -1) {
        const monthNum = monthIdx + 1;
        const defaultStart = getDefaultStartDate(monthNum, year);
        setSelectedDateStart(formatDateToYYYYMMDD(defaultStart));
        
        const lastDay = new Date(year, monthNum, 0).getDate();
        const endOfSelectedMonth = new Date(year, monthNum - 1, lastDay);
        setSelectedDateEnd(formatDateToYYYYMMDD(endOfSelectedMonth));
      }
    }
  }, [selectedMonth, normalizedRows, dateBounds.max]);

  const customers = useMemo(() => {
    return [...new Set(normalizedRows.map(r => r.customer).filter(Boolean))].sort();
  }, [normalizedRows]);

  // Models are ordered by total shipped quantity (QTY ACTUAL) descending, so the
  // most-shipped model appears first in the dropdown. Models that have never had
  // any plan (QTY TARGET always 0) are hidden from the list since there's nothing
  // to track for them.
  const models = useMemo(() => {
    const statsByModel: Record<string, { qtyActual: number; qtyTarget: number }> = {};
    normalizedRows.forEach(r => {
      if (!r.model) return;
      if (!statsByModel[r.model]) statsByModel[r.model] = { qtyActual: 0, qtyTarget: 0 };
      statsByModel[r.model].qtyActual += r.qtyActual;
      statsByModel[r.model].qtyTarget += r.qtyTarget;
    });

    return Object.keys(statsByModel)
      .filter(m => statsByModel[m].qtyTarget > 0) // hide models with no plan at all
      .sort((a, b) => statsByModel[b].qtyActual - statsByModel[a].qtyActual);
  }, [normalizedRows]);

  const monthsList = useMemo(() => {
    return [...new Set(normalizedRows.map(r => r.month).filter(Boolean))].sort();
  }, [normalizedRows]);

  // Clock sync
  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  const formattedTime = currentTime.toLocaleString(lang === 'vi' ? 'vi-VN' : lang === 'ko' ? 'ko-KR' : 'en-US', {
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  });

  // Filtered rows for current selection
  const filteredRecords = useMemo(() => {
    const start = selectedDateStart ? new Date(selectedDateStart) : null;
    if (start) start.setHours(0, 0, 0, 0);
    const end = selectedDateEnd ? new Date(selectedDateEnd) : null;
    if (end) end.setHours(23, 59, 59, 999);

    // Derive the "context year" from the selected date range for M/D rows
    // whose year might have been set to a fallback (2026) during parsing.
    // When the filter dates share a single year, use that year.
    // When they span two years, let each row's actual date.getFullYear() take precedence.
    const filterYear = start ? start.getFullYear() : (end ? end.getFullYear() : new Date().getFullYear());

    let result = normalizedRows.filter(r => {
      const rowTime = r.date.getTime();
      const inDateRange = (!start || rowTime >= start.getTime()) && (!end || rowTime <= end.getTime());
      
      return inDateRange &&
        (!selectedCustomer || r.customer === selectedCustomer) &&
        (!selectedModel || r.model === selectedModel);
    });

    // Re-compute year/monthNum/week using the live filter year, so that when the
    // user changes the date range to a different year the week keys update correctly.
    // For M/D-only date strings, parseToDateObject uses 2026 as the fallback year.
    // Here we override that with the actual year from the filter selection.
    const endYear = end ? end.getFullYear() : filterYear;
    result = result.map(r => {
      const mNum = r.date.getMonth() + 1;
      const day  = r.date.getDate();
      // Choose year: if the row's month falls in the start half, use filterYear;
      // if it falls in the end half (when range crosses a year boundary), use endYear.
      const resolvedYear = (filterYear !== endYear && mNum < (start?.getMonth() ?? 0) + 1)
        ? endYear    // month is "ahead" in the next year
        : filterYear;
      const week = getWeekOfMonth(resolvedYear, mNum, day);
      return { ...r, year: resolvedYear, monthNum: mNum, week };
    });

    if (selectedMonth) {
      result = result.filter(r => r.month === selectedMonth);
    }

    if (selectedWeek) {
      result = result.filter(r => r.week === selectedWeek);
    }

    if (qtyTargetFilter === 'greater_than_zero') {
      result = result.filter(r => r.qtyTarget > 0);
    } else if (qtyTargetFilter === 'equal_to_zero') {
      result = result.filter(r => r.qtyTarget === 0);
    }

    if (qtyActualFilter === 'greater_than_zero') {
      result = result.filter(r => r.qtyActual > 0);
    } else if (qtyActualFilter === 'equal_to_zero') {
      result = result.filter(r => r.qtyActual === 0);
    }

    if (amtTargetFilter === 'greater_than_zero') {
      result = result.filter(r => r.amtTarget > 0);
    } else if (amtTargetFilter === 'equal_to_zero') {
      result = result.filter(r => r.amtTarget === 0);
    }

    if (amtActualFilter === 'greater_than_zero') {
      result = result.filter(r => r.amtActual > 0);
    } else if (amtActualFilter === 'equal_to_zero') {
      result = result.filter(r => r.amtActual === 0);
    }

    if (sortField) {
      result.sort((a, b) => {
        let valA: any = a[sortField as keyof typeof a];
        let valB: any = b[sortField as keyof typeof b];
        
        if (valA instanceof Date) valA = valA.getTime();
        if (valB instanceof Date) valB = valB.getTime();

        if (typeof valA === 'string') {
          return sortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        } else {
          return sortAsc ? valA - valB : valB - valA;
        }
      });
    }

    return result;
  }, [normalizedRows, selectedDateStart, selectedDateEnd, selectedCustomer, selectedModel, selectedMonth, selectedWeek, qtyTargetFilter, qtyActualFilter, amtTargetFilter, amtActualFilter, sortField, sortAsc]);



  const formatWeekLabel = (year: number, monthNum: number, weekStr: string): string => {
    const weekIndex = parseInt(weekStr.replace('W', ''), 10) || 1;
    const mStr = String(monthNum).padStart(2, '0');
    const lastDay = new Date(year, monthNum, 0).getDate();
    const start = (weekIndex - 1) * 7 + 1;
    const end = Math.min(weekIndex * 7, lastDay);
    const startStr = String(start).padStart(2, '0');
    const endStr = String(end).padStart(2, '0');

    if (lang === 'ko') {
      return `${weekIndex}주차 (${startStr}-${endStr}/${mStr})`;
    }
    if (lang === 'en') {
      return `Week ${weekIndex} (${startStr}-${endStr}/${mStr})`;
    }
    return `Tuần ${weekIndex} (${startStr}-${endStr}/${mStr})`;
  };

  const bottomTableRows = useMemo(() => {
    const computeRow = (model: string, customer: string, period: string, sortKeyDate: number | string, qtyTarget: number, qtyActual: number, amtTarget: number, amtActual: number) => ({
      model,
      customer,
      period,
      sortKeyDate,
      qtyTarget,
      qtyActual,
      qtyRatio: qtyTarget > 0 ? (qtyActual / qtyTarget) * 100 : 0,
      amtTarget,
      amtActual,
      amtRatio: amtTarget > 0 ? (amtActual / amtTarget) * 100 : 0,
    });

    let rows: ReturnType<typeof computeRow>[] = [];

    if (viewMode === 'day') {
      const groups: Record<string, { model: string; customer: string; period: string; qtyTarget: number; qtyActual: number; amtTarget: number; amtActual: number; sortKeyDate: number }> = {};
      filteredRecords.forEach(r => {
        const dateStr = r.date ? r.date.toLocaleDateString('vi-VN') : r.month;
        const sortKey = r.date ? r.date.getTime() : 0;
        const key = `${r.model}||${r.customer}||${dateStr}`;
        if (!groups[key]) {
          groups[key] = {
            model: r.model,
            customer: r.customer,
            period: dateStr,
            qtyTarget: 0,
            qtyActual: 0,
            amtTarget: 0,
            amtActual: 0,
            sortKeyDate: sortKey
          };
        }
        groups[key].qtyTarget += r.qtyTarget;
        groups[key].qtyActual += r.qtyActual;
        groups[key].amtTarget += r.amtTarget;
        groups[key].amtActual += r.amtActual;
      });
      rows = Object.values(groups).map(g =>
        computeRow(
          g.model,
          g.customer,
          g.period,
          g.sortKeyDate,
          g.qtyTarget,
          g.qtyActual,
          g.amtTarget,
          g.amtActual
        )
      );
    } else if (viewMode === 'week') {
      const groups: Record<string, { model: string; customer: string; weekLabel: string; qtyTarget: number; qtyActual: number; amtTarget: number; amtActual: number; sortKey: string }> = {};
      filteredRecords.forEach(r => {
        const yr = r.year ?? r.date.getFullYear();
        const mNum = r.monthNum ?? (r.date.getMonth() + 1);
        const wName = r.week ?? getWeekOfMonth(yr, mNum, r.date.getDate());
        const weekLabel = formatWeekLabel(yr, mNum, wName);
        const key = `${r.model}||${r.customer}||${wName}||${mNum}||${yr}`;
        if (!groups[key]) groups[key] = { model: r.model, customer: r.customer, weekLabel, qtyTarget: 0, qtyActual: 0, amtTarget: 0, amtActual: 0, sortKey: `${yr}-${String(mNum).padStart(2, '0')}-${wName}` };
        groups[key].qtyTarget += r.qtyTarget;
        groups[key].qtyActual += r.qtyActual;
        groups[key].amtTarget += r.amtTarget;
        groups[key].amtActual += r.amtActual;
      });
      rows = Object.values(groups).map(g => computeRow(g.model, g.customer, g.weekLabel, g.sortKey, g.qtyTarget, g.qtyActual, g.amtTarget, g.amtActual));
    } else if (viewMode === 'month') {
      const groups: Record<string, { model: string; customer: string; monthLabel: string; qtyTarget: number; qtyActual: number; amtTarget: number; amtActual: number; sortKey: string }> = {};
      filteredRecords.forEach(r => {
        const yr = r.year ?? r.date.getFullYear();
        const mNum = r.monthNum ?? (r.date.getMonth() + 1);
        const monthLabel = lang === 'vi' ? `Tháng ${mNum}/${yr}` : lang === 'ko' ? `${yr}년 ${mNum}월` : `Month ${mNum}/${yr}`;
        const key = `${r.model}||${r.customer}||${mNum}||${yr}`;
        if (!groups[key]) groups[key] = { model: r.model, customer: r.customer, monthLabel, qtyTarget: 0, qtyActual: 0, amtTarget: 0, amtActual: 0, sortKey: `${yr}-${String(mNum).padStart(2, '0')}` };
        groups[key].qtyTarget += r.qtyTarget;
        groups[key].qtyActual += r.qtyActual;
        groups[key].amtTarget += r.amtTarget;
        groups[key].amtActual += r.amtActual;
      });
      rows = Object.values(groups).map(g => computeRow(g.model, g.customer, g.monthLabel, g.sortKey, g.qtyTarget, g.qtyActual, g.amtTarget, g.amtActual));
    }

    // Sort
    if (sortField) {
      rows.sort((a, b) => {
        let valA: any = sortField === 'date' ? a.sortKeyDate : (a as any)[sortField];
        let valB: any = sortField === 'date' ? b.sortKeyDate : (b as any)[sortField];
        if (typeof valA === 'string' && typeof valB === 'string') {
          return sortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        }
        return sortAsc ? (valA ?? 0) - (valB ?? 0) : (valB ?? 0) - (valA ?? 0);
      });
    } else {
      // Default order (no column header sort active): rank models by their total
      // shipped quantity (QTY ACTUAL) across the whole filtered range, highest first,
      // so models with the most output float to the top on every page, and models
      // with zero output sink to the bottom.
      const totalQtyActualByModel: Record<string, number> = {};
      rows.forEach(r => {
        totalQtyActualByModel[r.model] = (totalQtyActualByModel[r.model] || 0) + r.qtyActual;
      });

      rows.sort((a, b) => {
        const totalA = totalQtyActualByModel[a.model] || 0;
        const totalB = totalQtyActualByModel[b.model] || 0;
        if (totalA !== totalB) return totalB - totalA; // highest shipped-qty model first
        if (a.model !== b.model) return a.model.localeCompare(b.model);
        if (a.customer !== b.customer) return a.customer.localeCompare(b.customer);
        return String(a.sortKeyDate).localeCompare(String(b.sortKeyDate));
      });
    }

    return rows;
  }, [filteredRecords, viewMode, lang, sortField, sortAsc]);

  // FIX (EPCC bottom=0): thay vì maxHeight cố định '62vh' (luôn để hở một
  // khoảng trắng cố định phía dưới bảng, bất kể cửa sổ cao/thấp), đo TRỰC
  // TIẾP vị trí thật của khung cuộn so với đáy viewport bằng
  // getBoundingClientRect(), rồi set maxHeight = phần còn lại tới đáy màn
  // hình (trừ chiều cao thanh phân trang + 1 khoảng đệm nhỏ để không dính
  // sát viền dưới cùng của trình duyệt). Cách này KHÔNG phụ thuộc cấu
  // trúc/height của DOM cha (App.tsx, Sidebar...) — chỉ dựa vào vị trí
  // thực tế của chính element, nên không lặp lại lỗi "khoảng trắng thừa"
  // từng gặp khi ép flex:1/100dvh trước đây.
  useEffect(() => {
    const recalcTableScrollHeight = () => {
      const el = tableScrollRef.current;
      if (!el) return;
      const top = el.getBoundingClientRect().top;
      // Trừ luôn chiều cao thật của thanh phân trang (nằm NGAY DƯỚI khung
      // cuộn, cùng bên trong panel) — nếu không trừ, panel (bảng+phân
      // trang) sẽ tràn xuống dưới đáy màn hình thêm ~chiều cao thanh đó.
      const paginationHeight = paginationBarRef.current?.getBoundingClientRect().height ?? 44;
      const BOTTOM_GAP = 12; // đệm nhỏ để không dính sát mép dưới cửa sổ
      const available = window.innerHeight - top - paginationHeight - BOTTOM_GAP;
      setTableScrollMaxHeight(Math.max(200, available));
    };
    recalcTableScrollHeight();
    window.addEventListener('resize', recalcTableScrollHeight);
    // Tính lại khi chuyển sang tab 'merged' (lúc element vừa mount/hiện ra)
    const raf = requestAnimationFrame(recalcTableScrollHeight);
    return () => {
      window.removeEventListener('resize', recalcTableScrollHeight);
      cancelAnimationFrame(raf);
    };
  }, [activeTab, bottomTableRows.length]);

  // FIX (EPCC-fit-31-days-per-page): tính lại font-size/padding dọc của
  // TỪNG Ô sao cho đúng (pageSize + 1 dòng TỔNG) luôn vừa khít chiều cao
  // khung cuộn (tableScrollMaxHeight) — KHÔNG cần cuộn thêm bên trong, dù
  // đang ở trang 1, 2, 3... nào. Đo theadHeight THẬT (ref) thay vì đoán
  // cứng, nên vẫn đúng cả khi nội dung header (select lọc, sort icon...)
  // thay đổi chiều cao.
  const theadRef = useRef<HTMLTableSectionElement | null>(null);
  const [rowFontSize, setRowFontSize] = useState<number>(14);
  const [rowPaddingY, setRowPaddingY] = useState<number>(10);

  const processedRows = bottomTableRows;

  const {
    page: safePage,
    pageSize,
    setPage,
    setPageSize,
    totalPages,
    pagedData: pagedRows,
    pageNums,
  } = usePagination(processedRows, 31);

  // FIX (EPCC-fit-31-days-per-page): đo chiều cao THẬT của thead (2 dòng
  // tiêu đề sticky, kể cả select lọc bên trong) rồi chia phần chiều cao
  // còn lại của khung cuộn (tableScrollMaxHeight) cho đúng (pageSize + 1)
  // dòng — dùng pageSize CỐ ĐỊNH (không phải pagedRows.length thực tế của
  // trang hiện tại) làm mẫu số, để chiều cao mỗi dòng GIỐNG HỆT nhau dù
  // đang ở trang 1, 2, 3... (kể cả trang cuối có ít dòng hơn 31 vẫn không
  // bị kéo giãn cao hơn các trang khác). fontSize/padding dọc được scale
  // theo tỉ lệ so với kích thước gốc (14px/10px, ứng với hàng ~39-41px)
  // và clamp trong khoảng đọc được — nếu màn hình quá thấp khiến giá trị
  // chạm mức tối thiểu, overflow-y: auto của table-scroll vẫn đảm nhận
  // việc cuộn bên trong (không vỡ layout).
  useEffect(() => {
    const BASE_FONT = 14, BASE_PADDING = 10;
    const BASE_ROW_HEIGHT = BASE_FONT * 1.3 + BASE_PADDING * 2 + 1; // ≈ 39.2px
    const MIN_FONT = 10, MAX_FONT = 16;
    const MIN_PAD = 3, MAX_PAD = 14;
    // FIX (EPCC-increase-table-font-27pct): chữ trong bảng cần lớn hơn ~25-30%
    // so với trước. Vì thuật toán fit-vừa-khung tự bù trừ theo tỉ lệ
    // BASE_FONT/BASE_ROW_HEIGHT (đổi BASE_FONT không đổi kết quả cuối), ta
    // nhân thêm FONT_BOOST sau khi tính "scale" như cũ, và nới trần
    // MIN/MAX_FONT, MIN/MAX_PAD theo đúng hệ số này để mức tăng không bị
    // clamp che mất. Nếu khung quá thấp, table-scroll vẫn overflow-y: auto.
    const FONT_BOOST = 1.28; // +28%, nằm giữa khoảng 25-30% yêu cầu

    const recalcRowSize = () => {
      const theadEl = theadRef.current;
      if (!theadEl) return;
      const theadHeight = theadEl.getBoundingClientRect().height;
      const targetRowCount = pageSize + 1; // +1 cho dòng TỔNG
      const availableForRows = tableScrollMaxHeight - theadHeight;
      const rawRowHeight = availableForRows / targetRowCount;
      const scale = rawRowHeight / BASE_ROW_HEIGHT;

      const nextFont = Math.min(MAX_FONT * FONT_BOOST, Math.max(MIN_FONT * FONT_BOOST, BASE_FONT * scale * FONT_BOOST));
      const nextPad = Math.min(MAX_PAD * FONT_BOOST, Math.max(MIN_PAD * FONT_BOOST, BASE_PADDING * scale * FONT_BOOST));
      setRowFontSize(nextFont);
      setRowPaddingY(nextPad);
    };

    recalcRowSize();
    // thead vừa mount/đổi trang/đổi tab cần 1 frame để có kích thước thật
    const raf = requestAnimationFrame(recalcRowSize);
    return () => cancelAnimationFrame(raf);
  }, [pageSize, tableScrollMaxHeight, activeTab, safePage]);

  // Reset page to 1 whenever filters change
  useEffect(() => {
    setPage(1);
  }, [
    selectedCustomer, selectedModel, qtyTargetFilter, qtyActualFilter,
    amtTargetFilter, amtActualFilter, selectedWeek, selectedMonth,
    selectedDateStart, selectedDateEnd, viewMode, setPage
  ]);

  // Totals row computed from current bottomTableRows
  const tableTotals = useMemo(() => {
    const qtyTarget = bottomTableRows.reduce((s, r) => s + r.qtyTarget, 0);
    const qtyActual = bottomTableRows.reduce((s, r) => s + r.qtyActual, 0);
    const amtTarget = bottomTableRows.reduce((s, r) => s + r.amtTarget, 0);
    const amtActual = bottomTableRows.reduce((s, r) => s + r.amtActual, 0);
    return {
      qtyTarget, qtyActual, amtTarget, amtActual,
      qtyRatio: qtyTarget > 0 ? (qtyActual / qtyTarget) * 100 : 0,
      amtRatio: amtTarget > 0 ? (amtActual / amtTarget) * 100 : 0,
    };
  }, [bottomTableRows]);

  // Re-draw Plotly charts on filter change
  useEffect(() => {
    if (activeTab !== 'summary') return;
    console.log('[Chart Debug]', {
      activeTab,
      salesQtyChartDomExists: !!document.getElementById('salesQtyChart'),
      salesQtyChartWidth: document.getElementById('salesQtyChart')?.clientWidth,
      salesQtyChartHeight: document.getElementById('salesQtyChart')?.clientHeight,
    });
    if (filteredRecords.length === 0 || !plotlyReady) return;

    const isDark = theme === 'dark';
    const axisTextColor = typeof window !== 'undefined'
      ? getComputedStyle(document.documentElement).getPropertyValue('--axis-text-color').trim() || (isDark ? '#E8E8F0' : '#1A1A2E')
      : (isDark ? '#E8E8F0' : '#1A1A2E');

    /* FIX (EPCC-rate-line-theme-color):
     * EXPLORE — line "% Rate" (traceQtyRate/traceAmtRate/traceSumRate) đang
     *   hard-code màu cố định '#eab308' (vàng) cho line và '#10b981' (xanh
     *   lá) cho viền marker, KHÔNG phụ thuộc `theme` → khi đổi Light/Dark,
     *   3 line/marker này không đổi màu như phần còn lại của chart (bar,
     *   text label đều đã dùng axisTextColor theo theme), gây tương phản
     *   kém / khó nhìn ở 1 trong 2 theme.
     * PLAN — khai báo 2 biến màu theo `isDark` ngay tại đây, dùng chung cho
     *   cả 3 line trace bên dưới, để đồng bộ và dễ bảo trì (đổi 1 chỗ áp
     *   dụng cho cả 3 chart).
     * CODE — xem rateLineColor/rateMarkerColor áp dụng tại traceQtyRate,
     *   traceAmtRate, traceSumRate. */
    const rateLineColor = isDark ? '#eab308' : '#b45309';
    const rateMarkerColor = isDark ? '#10b981' : '#047857';

    let xQty: string[] = [];
    let qtyTargets: number[] = [];
    let qtyActuals: number[] = [];
    let qtyRates: number[] = [];

    let xAmt: string[] = [];
    let amtTargets: number[] = [];
    let amtActuals: number[] = [];
    let amtRates: number[] = [];

    if (viewMode === 'week') {
      // Group by fully-qualified YYYY-MM-WN key → prevents cross-month / cross-year collision
      type WeekEntry = { qtyTarget: number; qtyActual: number; amtTarget: number; amtActual: number; label: string };
      const weekMap: Record<string, WeekEntry> = {};

      filteredRecords.forEach(r => {
        const yr   = r.year      ?? r.date.getFullYear();
        const mNum = r.monthNum  ?? (r.date.getMonth() + 1);
        const wName = r.week     ?? getWeekOfMonth(yr, mNum, r.date.getDate());
        const key = buildWeekKey(yr, mNum, wName); // e.g. "2026-06-W1"

        if (!weekMap[key]) {
          // Build human-readable label from the stable key
          const labelW = lang === 'vi'
            ? `Tháng ${mNum}/${yr} - ${wName}`
            : lang === 'ko'
              ? `${yr}년 ${mNum}월 - ${wName}`
              : `${yr}-${String(mNum).padStart(2, '0')} ${wName}`;
          weekMap[key] = { qtyTarget: 0, qtyActual: 0, amtTarget: 0, amtActual: 0, label: labelW };
        }
        weekMap[key].qtyTarget += r.qtyTarget;
        weekMap[key].qtyActual += r.qtyActual;
        weekMap[key].amtTarget += r.amtTarget;
        weekMap[key].amtActual += r.amtActual;
      });

      // Lexicographic sort on YYYY-MM-WN keys is chronologically correct
      const activeWeeks = Object.keys(weekMap).sort();

      xQty = activeWeeks.map(k => weekMap[k].label);
      qtyTargets = activeWeeks.map(k => weekMap[k].qtyTarget);
      qtyActuals = activeWeeks.map(k => weekMap[k].qtyActual);
      qtyRates   = activeWeeks.map(k => weekMap[k].qtyTarget > 0 ? (weekMap[k].qtyActual / weekMap[k].qtyTarget) * 100 : 0);

      xAmt = [...xQty];
      amtTargets = activeWeeks.map(k => weekMap[k].amtTarget);
      amtActuals = activeWeeks.map(k => weekMap[k].amtActual);
      amtRates   = activeWeeks.map(k => weekMap[k].amtTarget > 0 ? (weekMap[k].amtActual / weekMap[k].amtTarget) * 100 : 0);

    } else if (viewMode === 'month') {
      // Use "YYYY-MM" as key to prevent cross-year collisions
      const monthMap: Record<string, { qtyTarget: number; qtyActual: number; amtTarget: number; amtActual: number; label: string }> = {};
      filteredRecords.forEach(r => {
        const yr   = r.year     ?? r.date.getFullYear();
        const mNum = r.monthNum ?? (r.date.getMonth() + 1);
        const key  = `${yr}-${String(mNum).padStart(2, '0')}`; // e.g. "2026-06"
        if (!monthMap[key]) {
          const lbl = lang === 'vi'
            ? `Tháng ${mNum}/${yr}`
            : lang === 'ko'
              ? `${yr}년 ${mNum}월`
              : `${yr}-${String(mNum).padStart(2, '0')}`;
          monthMap[key] = { qtyTarget: 0, qtyActual: 0, amtTarget: 0, amtActual: 0, label: lbl };
        }
        monthMap[key].qtyTarget += r.qtyTarget;
        monthMap[key].qtyActual += r.qtyActual;
        monthMap[key].amtTarget += r.amtTarget;
        monthMap[key].amtActual += r.amtActual;
      });

      // Lexicographic sort on YYYY-MM keys is chronologically correct
      const activeMonths = Object.keys(monthMap).sort();

      xQty = activeMonths.map(k => monthMap[k].label);
      qtyTargets = activeMonths.map(k => monthMap[k].qtyTarget);
      qtyActuals = activeMonths.map(k => monthMap[k].qtyActual);
      qtyRates   = activeMonths.map(k => monthMap[k].qtyTarget > 0 ? (monthMap[k].qtyActual / monthMap[k].qtyTarget) * 100 : 0);

      xAmt = [...xQty];
      amtTargets = activeMonths.map(k => monthMap[k].amtTarget);
      amtActuals = activeMonths.map(k => monthMap[k].amtActual);
      amtRates   = activeMonths.map(k => monthMap[k].amtTarget > 0 ? (monthMap[k].amtActual / monthMap[k].amtTarget) * 100 : 0);

    } else {
      const modelMap: Record<string, { qtyTarget: number; qtyActual: number; amtTarget: number; amtActual: number }> = {};
      filteredRecords.forEach(r => {
        if (!modelMap[r.model]) {
          modelMap[r.model] = { qtyTarget: 0, qtyActual: 0, amtTarget: 0, amtActual: 0 };
        }
        modelMap[r.model].qtyTarget += r.qtyTarget;
        modelMap[r.model].qtyActual += r.qtyActual;
        modelMap[r.model].amtTarget += r.amtTarget;
        modelMap[r.model].amtActual += r.amtActual;
      });

      // FIX "model giá trị 0 vẫn hiện trên biểu đồ": model không có dữ liệu
      // thật (cả Target VÀ Actual đều bằng 0) chỉ làm rối biểu đồ, không
      // mang thông tin gì — lọc bỏ trước khi build trục X. Model chỉ thiếu
      // 1 trong 2 (có Target chưa đạt, hoặc có Actual dù không đặt Target)
      // vẫn giữ lại vì vẫn là thông tin có ý nghĩa.
      const activeModelsQty = Object.keys(modelMap)
        .filter(m => modelMap[m].qtyTarget !== 0 || modelMap[m].qtyActual !== 0)
        .sort((a, b) => {
        if (modelMap[b].qtyActual !== modelMap[a].qtyActual) {
          return modelMap[b].qtyActual - modelMap[a].qtyActual;
        }
        if (modelMap[b].qtyTarget !== modelMap[a].qtyTarget) {
          return modelMap[b].qtyTarget - modelMap[a].qtyTarget;
        }
        return a.localeCompare(b);
      });
      xQty = activeModelsQty;
      qtyTargets = activeModelsQty.map(m => modelMap[m].qtyTarget);
      qtyActuals = activeModelsQty.map(m => modelMap[m].qtyActual);
      qtyRates = activeModelsQty.map(m => modelMap[m].qtyTarget > 0 ? (modelMap[m].qtyActual / modelMap[m].qtyTarget) * 100 : 0);

      const activeModelsAmt = Object.keys(modelMap)
        .filter(m => modelMap[m].amtTarget !== 0 || modelMap[m].amtActual !== 0)
        .sort((a, b) => {
        if (modelMap[b].amtActual !== modelMap[a].amtActual) {
          return modelMap[b].amtActual - modelMap[a].amtActual;
        }
        if (modelMap[b].amtTarget !== modelMap[a].amtTarget) {
          return modelMap[b].amtTarget - modelMap[a].amtTarget;
        }
        return a.localeCompare(b);
      });
      xAmt = activeModelsAmt;
      amtTargets = activeModelsAmt.map(m => modelMap[m].amtTarget);
      amtActuals = activeModelsAmt.map(m => modelMap[m].amtActual);
      amtRates = activeModelsAmt.map(m => modelMap[m].amtTarget > 0 ? (modelMap[m].amtActual / modelMap[m].amtTarget) * 100 : 0);
    }

    // ----------------------------------------------------
    // CHART 1: 판매 Q'TY
    // ----------------------------------------------------
    const maxQtyVal = Math.max(...qtyTargets, ...qtyActuals, 10);

    const formatNumber = (val: number, isAmt: boolean): string => {
      const rounded = Math.round(val);
      const formatted = rounded.toLocaleString('vi-VN');
      return isAmt ? `$${formatted}` : formatted;
    };

    const traceQtyTarget = {
      x: xQty,
      y: qtyTargets,
      type: 'bar',
      name: lang === 'vi' ? 'Mục tiêu (Target)' : lang === 'ko' ? '목표 (Target)' : 'Target',
      marker: {
        color: '#E8836B',
        line: { color: '#d16b53', width: 1.5 }
      },
      yaxis: 'y',
      text: qtyTargets.map(v => v > 0 ? formatNumber(v, false) : ''),
      textposition: 'auto',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      insidetextfont: {
        color: '#ffffff',
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      outsidetextfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      cliponaxis: false
    };

    const traceQtyActual = {
      x: xQty,
      y: qtyActuals,
      type: 'bar',
      name: lang === 'vi' ? 'Thực tế (Actual)' : lang === 'ko' ? '실적 (Actual)' : 'Actual',
      marker: {
        color: '#2e7d8c',
        line: { color: '#1f5a66', width: 1 }
      },
      yaxis: 'y',
      text: qtyActuals.map(v => v > 0 ? formatNumber(v, false) : ''),
      textposition: 'auto',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      insidetextfont: {
        color: '#ffffff',
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      outsidetextfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      cliponaxis: false
    };

    const traceQtyRate = {
      x: xQty,
      y: qtyRates,
      type: 'scatter',
      mode: 'lines+markers+text',
      name: '% Rate',
      line: { color: rateLineColor, width: 3, dash: 'dash' },
      marker: { size: 10, color: rateMarkerColor, line: { color: rateLineColor, width: 2 } },
      yaxis: 'y2',
      text: qtyRates.map(v => v > 0 ? Math.round(v) + '%\n' : ''),
      textposition: 'top center',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      }
    };

    const layoutQty = getChartLayout('bar', theme, {
      categoryarray: xQty,
      yaxisRange: [0, maxQtyVal * 1.15],
      yaxis2Range: [-20, 140]
    });
    // FIX (header-legend-merge v2, EPCC): getChartLayout() tính margin.t
    // (khoảng trống phía trên biểu đồ) DỰA TRÊN showlegend MẶC ĐỊNH (true)
    // ngay TẠI THỜI ĐIỂM gọi hàm — set showlegend=false SAU đó không làm
    // margin.t tự co lại theo, nên khoảng trắng dành cho legend cũ vẫn còn
    // nguyên (đúng như ảnh phản hồi: khoảng cách header↔biểu đồ quá rộng).
    // Ép thẳng margin.t nhỏ lại, giữ nguyên l/r/b đang có.
    (layoutQty as any).showlegend = false;
    (layoutQty as any).margin = { ...(((layoutQty as any).margin) || {}), t: 20 };

    /* FIX (EPCC-xaxis-model-labels-cut-off):
     * EXPLORE — khi số lượng model nhiều (>8), tên model ở trục X ("SO2701
     *   S1", "SO1C00 S2"...) bị chồng/lem lên nhau hoặc bị CẮT CỤT phía
     *   dưới (nhìn ảnh phản hồi: chữ mờ, thiếu nét) vì (a) `tickangle` mặc
     *   định trong chartTheme.ts không đủ dốc để tách các nhãn dài ra xa
     *   nhau, và (b) margin-bottom mặc định không đủ chỗ chứa hết chiều
     *   cao của nhãn xoay nghiêng → Plotly clip bớt phần dưới.
     * PLAN — ép `tickangle: -45`, bật `automargin: true` (Plotly tự đo và
     *   chừa đủ khoảng trống cho nhãn dài thay vì dùng margin cố định),
     *   đồng thời tăng margin.b tối thiểu lên 95px làm nền, để nhãn luôn
     *   hiển thị TRỌN VẸN dù có bao nhiêu model.
     * CODE — override xaxis + margin.b ngay sau khi lấy layout từ
     *   getChartLayout(), áp dụng cho cả layoutQty và layoutAmt (2 chart
     *   có trục X theo model). */
    (layoutQty as any).xaxis = {
      ...(((layoutQty as any).xaxis) || {}),
      tickangle: -45,
      automargin: true,
      tickfont: { ...((((layoutQty as any).xaxis) || {}).tickfont || {}), size: 11 }
    };
    (layoutQty as any).margin = { ...(((layoutQty as any).margin) || {}), t: 20, b: Math.max(95, ((layoutQty as any).margin?.b) || 0) };

    window.Plotly.newPlot('salesQtyChart', [traceQtyTarget, traceQtyActual, traceQtyRate] as any, layoutQty as any, { displayModeBar: false, responsive: true });

    // ----------------------------------------------------
    // CHART 2: 판매 AMT
    // ----------------------------------------------------
    const maxAmtVal = Math.max(...amtTargets, ...amtActuals, 10);

    const traceAmtTarget = {
      x: xAmt,
      y: amtTargets,
      type: 'bar',
      name: lang === 'vi' ? 'Mục tiêu (Target)' : lang === 'ko' ? '목표 (Target)' : 'Target',
      marker: {
        color: '#E8836B',
        line: { color: '#d16b53', width: 1.5 }
      },
      yaxis: 'y',
      text: amtTargets.map(v => v > 0 ? formatNumber(v, true) : ''),
      textposition: 'auto',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      insidetextfont: {
        color: '#ffffff',
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      outsidetextfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      cliponaxis: false
    };

    const traceAmtActual = {
      x: xAmt,
      y: amtActuals,
      type: 'bar',
      name: lang === 'vi' ? 'Thực tế (Actual)' : lang === 'ko' ? '실적 (Actual)' : 'Actual',
      marker: {
        color: '#2e7d8c',
        line: { color: '#1f5a66', width: 1 }
      },
      yaxis: 'y',
      text: amtActuals.map(v => v > 0 ? formatNumber(v, true) : ''),
      textposition: 'auto',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      insidetextfont: {
        color: '#ffffff',
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      outsidetextfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      cliponaxis: false
    };

    const traceAmtRate = {
      x: xAmt,
      y: amtRates,
      type: 'scatter',
      mode: 'lines+markers+text',
      name: '% Rate',
      line: { color: rateLineColor, width: 3, dash: 'dash' },
      marker: { size: 10, color: rateMarkerColor, line: { color: rateLineColor, width: 2 } },
      yaxis: 'y2',
      text: amtRates.map(v => v > 0 ? Math.round(v) + '%\n' : ''),
      textposition: 'top center',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      }
    };

    const layoutAmt = getChartLayout('bar', theme, {
      categoryarray: xAmt,
      yaxisRange: [0, maxAmtVal * 1.15],
      yaxis2Range: [-20, 140]
    });
    (layoutAmt as any).showlegend = false; // FIX (header-legend-merge, EPCC) — legend đã chuyển lên header
    (layoutAmt as any).margin = { ...(((layoutAmt as any).margin) || {}), t: 20 }; // FIX v2 — margin.t vẫn tính theo showlegend cũ, ép nhỏ lại

    // FIX (EPCC-xaxis-model-labels-cut-off) — xem giải thích đầy đủ ở layoutQty
    // phía trên; áp dụng tương tự cho chart AMT vì trục X cũng theo model.
    (layoutAmt as any).xaxis = {
      ...(((layoutAmt as any).xaxis) || {}),
      tickangle: -45,
      automargin: true,
      tickfont: { ...((((layoutAmt as any).xaxis) || {}).tickfont || {}), size: 11 }
    };
    (layoutAmt as any).margin = { ...(((layoutAmt as any).margin) || {}), t: 20, b: Math.max(95, ((layoutAmt as any).margin?.b) || 0) };

    window.Plotly.newPlot('salesAmtChart', [traceAmtTarget, traceAmtActual, traceAmtRate] as any, layoutAmt as any, { displayModeBar: false, responsive: true });

    // ----------------------------------------------------
    // CHART 3: SALES AMT OF SELECTED MONTH
    // ----------------------------------------------------
    const sumQtyTarget = filteredRecords.reduce((s, r) => s + r.qtyTarget, 0);
    const sumQtyActual = filteredRecords.reduce((s, r) => s + r.qtyActual, 0);
    const sumAmtTarget = filteredRecords.reduce((s, r) => s + r.amtTarget, 0);
    const sumAmtActual = filteredRecords.reduce((s, r) => s + r.amtActual, 0);

    const totalQtyRate = sumQtyTarget > 0 ? (sumQtyActual / sumQtyTarget) * 100 : 0;
    const totalAmtRate = sumAmtTarget > 0 ? (sumAmtActual / sumAmtTarget) * 100 : 0;

    const maxSumVal = Math.max(sumQtyTarget, sumQtyActual, sumAmtTarget, sumAmtActual, 10);

    const traceSumTarget = {
      x: ['QTY', 'AMT'],
      y: [sumQtyTarget, sumAmtTarget],
      type: 'bar',
      name: lang === 'vi' ? 'Target' : 'Target',
      marker: {
        color: '#E8836B',
        line: { color: '#d16b53', width: 1.5 }
      },
      text: [sumQtyTarget, sumAmtTarget].map((v, i) => formatNumber(v, i === 1)),
      textposition: 'auto',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      insidetextfont: {
        color: '#ffffff',
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      outsidetextfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      cliponaxis: false
    };

    const traceSumActual = {
      x: ['QTY', 'AMT'],
      y: [sumQtyActual, sumAmtActual],
      type: 'bar',
      name: lang === 'vi' ? 'Actual' : 'Actual',
      marker: {
        color: '#2e7d8c',
        line: { color: '#1f5a66', width: 1 }
      },
      text: [sumQtyActual, sumAmtActual].map((v, i) => formatNumber(v, i === 1)),
      textposition: 'auto',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      insidetextfont: {
        color: '#ffffff',
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      outsidetextfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      },
      cliponaxis: false
    };

    const traceSumRate = {
      x: ['QTY', 'AMT'],
      y: [totalQtyRate, totalAmtRate],
      type: 'scatter',
      mode: 'lines+markers+text',
      name: '% Rate',
      line: { color: rateLineColor, width: 3, dash: 'dash' },
      marker: { size: 10, color: rateMarkerColor },
      yaxis: 'y2',
      text: [totalQtyRate, totalAmtRate].map(v => Math.round(v) + '%\n'),
      textposition: 'top center',
      textfont: {
        color: axisTextColor,
        size: chartTheme.common.dataLabelFontSize,
        weight: 'bold'
      }
    };

    const layoutSum = getChartLayout('summary', theme, {
      categoryarray: ['QTY', 'AMT'],
      yaxisRange: [0, maxSumVal * 1.15],
      yaxis2Range: [Math.min(totalQtyRate, totalAmtRate) - 10, Math.max(totalQtyRate, totalAmtRate) + 10]
    });
    (layoutSum as any).showlegend = false; // FIX (header-legend-merge, EPCC) — legend đã chuyển lên header
    (layoutSum as any).margin = { ...(((layoutSum as any).margin) || {}), t: 20 }; // FIX v2 — margin.t vẫn tính theo showlegend cũ, ép nhỏ lại

    window.Plotly.newPlot('salesSummaryChart', [traceSumTarget, traceSumActual, traceSumRate] as any, layoutSum as any, { displayModeBar: false, responsive: true });

    // ----------------------------------------------------
    // CHART 4: % Rate Concentric Semi-Donut Gauge by MODEL — rendered as SVG React component
    // ----------------------------------------------------
    const modelMapGauge: Record<string, { target: number; actual: number }> = {};
    filteredRecords.forEach(r => {
      if (!modelMapGauge[r.model]) {
        modelMapGauge[r.model] = { target: 0, actual: 0 };
      }
      modelMapGauge[r.model].target += r.amtTarget;
      modelMapGauge[r.model].actual += r.amtActual;
    });

    const gaugeModels = Object.entries(modelMapGauge)
      // FIX "model giá trị 0 vẫn hiện trên biểu đồ": bỏ model hoàn toàn
      // không có dữ liệu (target=0 VÀ actual=0) trước khi tính rate/xếp hạng,
      // tránh chúng chiếm chỗ trong top 5 hoặc hiện rate=0 gây hiểu nhầm.
      .filter(([, val]) => val.target !== 0 || val.actual !== 0)
      .map(([name, val]) => {
        const rate = val.target > 0 ? Math.min((val.actual / val.target) * 100, 100) : 0;
        return { name, rate };
      })
      .sort((a, b) => b.rate - a.rate)
      .slice(0, 5);

    const totTG = Object.values(modelMapGauge).reduce((s, v) => s + v.target, 0);
    const totAG = Object.values(modelMapGauge).reduce((s, v) => s + v.actual, 0);
    setGaugeData(gaugeModels);
    setGaugeOverallRate(totTG > 0 ? Math.round((totAG / totTG) * 100) : 0);

    // ── DEBUG: đối chiếu amtTarget/amtActual THẬT (không clamp 100%) từng
    // model trong khoảng filter hiện tại — dùng để xác minh 145.8%/146% là
    // số liệu thật (Actual > Target thật) hay do DB thiếu dòng Target (Plan/
    // Plan S1/Plan S2 của AMT). So với file Excel gốc cùng khoảng thời gian
    // để kết luận. Chỉ log ra console, không ảnh hưởng UI/hiệu năng.
    if (typeof console !== 'undefined' && console.table) {
      const debugRows = Object.entries(modelMapGauge)
        .map(([name, val]) => ({
          model: name,
          amtTarget: Math.round(val.target),
          amtActual: Math.round(val.actual),
          rate_thuc_khong_clamp: val.target > 0 ? Math.round((val.actual / val.target) * 1000) / 10 + '%' : 'N/A (target=0)',
        }))
        .sort((a, b) => b.amtActual - a.amtActual);
      console.log('[DEBUG %Rate AMT] So sánh Target vs Actual THẬT theo model (khoảng filter hiện tại):');
      console.table(debugRows);
    }

    // ----------------------------------------------------
    // CHART 5: Custom (Qty) Donut
    // ----------------------------------------------------
    const customerQtyMap: Record<string, number> = {};
    filteredRecords.forEach(r => {
      customerQtyMap[r.customer] = (customerQtyMap[r.customer] || 0) + r.qtyActual;
    });

    const customQtyLabels = Object.keys(customerQtyMap);
    const customQtyValues = Object.values(customerQtyMap);

    const customQtyTexts = customQtyLabels.map((lbl, idx) => {
      const val = customQtyValues[idx];
      const total = customQtyValues.reduce((a, b) => a + b, 0);
      const pct = total > 0 ? ((val / total) * 100).toFixed(1) : '0';
      return `${lbl}<br>${formatNumber(val, false)}<br>${pct}%`;
    });

    const traceDonutQty = {
      labels: customQtyLabels,
      values: customQtyValues,
      type: 'pie',
      hole: 0.5,
      text: customQtyTexts,
      textinfo: 'text',
      textposition: 'inside',
      insidetextorientation: 'horizontal',
      marker: {
        colors: DONUT_COLORS
      },
      // FIX (header-legend-merge, EPCC): legend donut chuyển lên header màu
      // (xem customQtyLegend/<TALegendItem> trong JSX) — tắt legend Plotly
      // và mở lại domain full-width (trước đây chừa 24% bên phải cho
      // legend Plotly, giờ không cần nữa nên donut được vẽ to hơn).
      showlegend: false,
      domain: { x: [0, 1] }
    };

    const layoutDonutQty = getChartLayout('donut', theme);
    // FIX (donut-fill-empty-space, EPCC): cùng nguyên nhân với Chart 1/2/3 —
    // getChartLayout('donut', theme) tính margin dựa trên trạng thái
    // showlegend MẶC ĐỊNH (true, dành chỗ cho legend Plotly cũ bên phải)
    // TẠI THỜI ĐIỂM gọi hàm, không tự co lại khi trace override
    // showlegend:false phía trên. Ép nhỏ margin 4 phía + domain y full-height
    // để donut phóng to lấp hết khung trắng (đúng yêu cầu ảnh mẫu).
    (layoutDonutQty as any).margin = { l: 10, r: 10, t: 10, b: 10 };
    (layoutDonutQty as any).height = undefined; // để tự giãn theo chart-holder thay vì chiều cao cứng cũ

    window.Plotly.newPlot('customQtyChart', [traceDonutQty] as any, layoutDonutQty as any, { displayModeBar: false, responsive: true });

    // FIX (header-legend-merge, EPCC): dựng danh sách legend (tên khách +
    // màu khớp DONUT_COLORS theo đúng thứ tự) để header phía JSX render.
    setCustomQtyLegend(customQtyLabels.map((lbl, idx) => ({ label: lbl, color: DONUT_COLORS[idx % DONUT_COLORS.length] })));

    // ----------------------------------------------------
    // CHART 6: Custom (AMT) Donut
    // ----------------------------------------------------
    const customerAmtMap: Record<string, number> = {};
    filteredRecords.forEach(r => {
      customerAmtMap[r.customer] = (customerAmtMap[r.customer] || 0) + r.amtActual;
    });

    const customAmtLabels = Object.keys(customerAmtMap);
    const customAmtValues = Object.values(customerAmtMap);

    const customAmtTexts = customAmtLabels.map((lbl, idx) => {
      const val = customAmtValues[idx];
      const total = customAmtValues.reduce((a, b) => a + b, 0);
      const pct = total > 0 ? ((val / total) * 100).toFixed(1) : '0';
      return `${lbl}<br>${formatNumber(val, true)}<br>${pct}%`;
    });

    const traceDonutAmt = {
      labels: customAmtLabels,
      values: customAmtValues,
      type: 'pie',
      hole: 0.5,
      text: customAmtTexts,
      textinfo: 'text',
      textposition: 'inside',
      insidetextorientation: 'horizontal',
      marker: {
        colors: DONUT_COLORS
      },
      showlegend: false, // FIX (header-legend-merge, EPCC) — legend đã chuyển lên header
      domain: { x: [0, 1] }
    };

    const layoutDonutAmt = {
      ...layoutDonutQty
    };

    window.Plotly.newPlot('customAmtChart', [traceDonutAmt] as any, layoutDonutAmt as any, { displayModeBar: false, responsive: true });

    setCustomAmtLegend(customAmtLabels.map((lbl, idx) => ({ label: lbl, color: DONUT_COLORS[idx % DONUT_COLORS.length] })));

  }, [filteredRecords, lang, theme, selectedDateStart, selectedDateEnd, viewMode, activeTab, plotlyReady]);

  // Aggregate values for quick statistics
  const summaryStats = useMemo(() => {
    const qtyTarget = filteredRecords.reduce((s, r) => s + r.qtyTarget, 0);
    const qtyActual = filteredRecords.reduce((s, r) => s + r.qtyActual, 0);
    const amtTarget = filteredRecords.reduce((s, r) => s + r.amtTarget, 0);
    const amtActual = filteredRecords.reduce((s, r) => s + r.amtActual, 0);

    const qtyRate = qtyTarget > 0 ? (qtyActual / qtyTarget) * 100 : 0;
    const amtRate = amtTarget > 0 ? (amtActual / amtTarget) * 100 : 0;

    return { qtyTarget, qtyActual, qtyRate, amtTarget, amtActual, amtRate };
  }, [filteredRecords]);

  // FIX (EPCC-vanilla-toolbar): thanh filter đổi nền sang "Vanilla"
  // (#FFF4D6, cố định, không đổi theo theme) theo ảnh tham chiếu — nên màu
  // chữ nhãn cũng cố định luôn (cam đậm) thay vì đổi theo theme như trước,
  // để luôn đủ tương phản trên nền vàng nhạt ở CẢ light lẫn dark theme.
  const filterLabelColor = '#C0EF6A';

  // FIX (EPCC-no-scroll-page): trang không cuộn được để xem phần dưới
  // (Khách hàng Qty/Amt, ...). Root div này trước đây chỉ có position:relative,
  // không tự khai báo height/overflow — nếu App.tsx bọc dashboard trong 1
  // container display:flex có chiều cao cố định + overflow:hidden (layout
  // sidebar cố định), flex-item con mặc định min-height:auto nên phình to
  // theo nội dung, vượt khung cha, và phần dư bị cha overflow:hidden cắt mất
  // — không nơi nào có scrollbar để cuộn.
  // Cách sửa AN TOÀN, không phụ thuộc tên class/cấu trúc DOM cha (tránh lặp
  // lại lỗi "dễ vỡ" đã gặp với .app-content/.app-layout trước đây): để chính
  // component này tự quản lý cuộn của nó — minHeight:0 cho phép nó co lại
  // đúng theo chiều cao cha cấp (dù cha là flex hay block), height:100% để
  // lấp đầy đúng khung cha (không hơn không kém), overflowY:auto để bản thân
  // nó có scrollbar riêng khi nội dung dài hơn khung nhìn.
  return (
    <div 
      className="sales-dashboard second-dashboard" 
      style={{ 
        position: 'relative', 
        zIndex: 1,
        height: '100%',
        minHeight: 0,
        overflowY: 'auto',
      }}
    >
      <style>{`
        /* ═══════════════════════════════════════════════════════════════
           FIX (EPCC-merged-table-blank-space v2): bảng "DOANH SỐ & SẢN
           LƯỢNG" từng bị khoảng trắng lớn ở dưới do ép panel/bảng
           flex:1 fill hết 100dvh dựa vào giả định cấu trúc DOM cha
           (.app-content/.app-layout) — cách này DỄ VỠ vì phụ thuộc tên
           class của cha (có thể đổi bất cứ lúc nào khi sửa App.tsx/
           Sidebar, như đã xảy ra) và trình duyệt phải hỗ trợ :has().

           THAM CHIẾU: bảng "Toàn bộ Cơ sở Dữ liệu từ Excel" (ảnh mẫu)
           không dùng chiêu ép 100dvh này — nó chỉ cho khu vực cuộn của
           bảng một CHIỀU CAO CỐ ĐỊNH/hợp lý (vd theo vh) với overflow-y
           riêng, độc lập hoàn toàn với DOM cha. Panel/table-container
           vẫn có height TỰ NHIÊN theo nội dung — không còn khoảng trắng
           thừa vì không còn ép chiều cao vượt quá nội dung thật.
           ═══════════════════════════════════════════════════════════════ */
        .second-dashboard .hero {
          padding: 16px 20px 4px !important;
        }
        .second-dashboard .hero h1 {
          margin: 0 0 6px !important;
          font-size: 26px !important;
        }
        .second-dashboard .header-line {
          margin-bottom: 12px !important;
        }
        .second-dashboard .tab-container {
          display: flex;
          gap: 10px;
          margin-bottom: 8px;
        }
        .second-dashboard .tab-btn {
          padding: 8px 18px;
          border-radius: 8px;
          font-weight: 700;
          font-size: 14px;
          cursor: pointer;
          transition: all 0.2s ease;
          border: 1px solid var(--border);
          background: rgba(30, 41, 59, 0.2);
          color: var(--text-2);
        }
        .second-dashboard .tab-btn.active {
          background: linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%);
          color: #ffffff;
          border-color: #3b82f6;
          box-shadow: 0 4px 12px rgba(59, 130, 246, 0.35);
        }
        .second-dashboard .chart-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          grid-auto-rows: 1fr;
          gap: 20px;
          margin-bottom: 24px;
        }
        /* Panel bảng "DOANH SỐ & SẢN LƯỢNG": height TỰ NHIÊN theo nội
           dung (giống mọi panel bình thường khác trong app, kể cả bảng
           tham chiếu "Toàn bộ Cơ sở Dữ liệu") — không ép flex:1/100dvh. */
        .second-dashboard.second-dashboard.second-dashboard .merged-fill-panel,
        div.panel.merged-fill-panel {
          height: auto !important;
          max-height: none !important;
        }
        .second-dashboard .chart-panel {
          padding: 0 !important;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          height: 338px;
        }
        .second-dashboard .chart-holder {
          padding: 0 16px 12px 16px;
          box-sizing: border-box;
          flex: 1;
          height: 100%;
        }
        @media (max-width: 1024px) {
          .second-dashboard .chart-grid {
            grid-template-columns: 1fr;
          }
        }
        .second-dashboard .card-header-styled {
          background: linear-gradient(90deg, #1e293b 0%, #0f172a 100%);
          color: #ffffff;
          border-radius: 8px 8px 0 0;
          padding: 12px 16px;
          font-weight: 800;
          font-size: 15px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-bottom: 1.5px solid #22d3ee;
        }
        .second-dashboard .card-header-styled.green-style {
          background: linear-gradient(90deg, #15803d 0%, #166534 100%);
          border-bottom: 1.5px solid #a3e635;
        }
        .second-dashboard .card-header-styled.blue-style {
          background: linear-gradient(90deg, #1d4ed8 0%, #1e40af 100%);
          border-bottom: 1.5px solid #38bdf8;
        }
        .second-dashboard .card-header-styled.purple-style {
          background: linear-gradient(90deg, #6d28d9 0%, #5b21b6 100%);
          border-bottom: 1.5px solid #c084fc;
        }
        .second-dashboard .stat-table {
          width: 100%;
          border-collapse: collapse;
          margin-top: 15px;
        }
        .second-dashboard .stat-table th, .second-dashboard .stat-table td {
          border: 1px solid var(--border);
          padding: 10px 14px;
          text-align: right;
          font-size: 13px;
        }
        .second-dashboard .stat-table th {
          background: rgba(30, 41, 59, 0.4);
          color: var(--text-1);
          font-weight: 700;
          text-align: left;
        }
        [data-theme="light"] .second-dashboard .stat-table th {
          background: rgba(100, 116, 139, 0.12);
          color: var(--tbl-head-color);
        }
        .second-dashboard .stat-table td.text-left {
          text-align: left;
          font-weight: 600;
        }
        .second-dashboard .stat-table tr.total-row {
          font-weight: 800;
          background: rgba(46, 125, 140, 0.1);
        }
        /* ── Table typography CSS variables ─────────────────── */
        .second-dashboard {
          --tbl-head-color:  #1a1a2e;
          --tbl-cell-color:  #1a1a2e;
          --tbl-model-color: #1e40af;
          --tbl-date-color:  #1e3a5f;
          --tbl-num-color:   #1a1a2e;
          --tbl-nil-color:   #6b7280;
          --tbl-head-bg:     rgba(59, 130, 246, 0.12);
        }
        [data-theme="dark"] .second-dashboard {
          --tbl-head-color:  #e8e8f0;
          --tbl-cell-color:  #e8e8f0;
          --tbl-model-color: #93c5fd;
          --tbl-date-color:  #bfdbfe;
          --tbl-num-color:   #f1f5f9;
          --tbl-nil-color:   #94a3b8;
          --tbl-head-bg:     rgba(59, 130, 246, 0.18);
        }
        /* ── Filter-select in table headers ──────────────────── */
        .second-dashboard .header-filter-select {
          width: 100%;
          background: var(--surface-2);
          border: 1px solid var(--border);
          color: var(--tbl-head-color);
          border-radius: var(--radius-xs);
          font-size: 16.5px; /* EPCC-increase-table-font-27pct: 13px -> 16.5px (+27%) */
          padding: 4px 6px;
          outline: none;
          cursor: pointer;
          font-family: var(--font-body);
          font-weight: 500;
          margin-top: 4px;
        }
        .second-dashboard .header-filter-select:hover {
          border-color: var(--border-hover);
        }
        /* ── Sort-able column header labels ──────────────────── */
        .second-dashboard .sort-header {
          cursor: pointer;
          user-select: none;
          font-size: 16.5px; /* EPCC-increase-table-font-27pct: 13px -> 16.5px (+27%) */
          font-weight: 600;
          color: var(--tbl-head-color);
          letter-spacing: 0.01em;
        }
        .second-dashboard .sort-header:hover {
          color: var(--purple-light);
        }
        .second-dashboard .filter-field {
          display: inline-flex;
          flex-direction: column;
          align-items: center;
          text-align: center;
        }
        .second-dashboard .filter-field label {
          text-align: center;
          width: 100%;
          display: block;
        }
        .second-dashboard .filter-field input,
        .second-dashboard .filter-field select {
          text-align: center;
          text-align-last: center;
        }
        /* FIX (EPCC-unify-kpi-box-size-muc2-theo-muc3): mục 2 (dashboard này)
           đang tự đặt padding riêng '12px 14px' cho .kpi-card, khiến kích
           thước (dài/rộng/cao) của ô KPI to hơn và lệch chuẩn so với ô KPI
           chuẩn ở mục 3 (PerCapitaTab / ManpowerDashboard) và mục 1
           (SalesDashboard) — cả 2 mục kia đều dùng đúng padding 7px 10px
           (className="kpi-card" mặc định theo index.css, border-radius
           var(--radius-md) = 12px không đổi). Đồng bộ lại padding về đúng
           7px 10px để 3 mục có ô KPI kích thước giống hệt nhau. */
        .second-dashboard .kpi-card {
          padding: 7px 10px;
        }
        .second-dashboard .kpi-card-header {
          margin-bottom: 8px;
        }
        .second-dashboard .kpi-card-label {
          font-size: 13.5px;
          font-weight: 700;
          color: var(--text-0);
        }
        .second-dashboard .kpi-card-value {
          font-weight: 800;
          color: var(--text-0);
        }
        .second-dashboard .kpi-card-target {
          font-size: 13.5px;
          font-weight: 700;
          color: var(--text-1);
        }
        /* Theme-aware and synchronized chart data labels style */
        .second-dashboard .textlayer g.trace:nth-of-type(1) text {
          fill: var(--label-color-target) !important;
          font-size: 13.5px !important;
          font-weight: 700 !important;
        }
        .second-dashboard .textlayer g.trace:nth-of-type(2) text {
          fill: var(--label-color-actual) !important;
          font-size: 13.5px !important;
          font-weight: 700 !important;
        }
      `}</style>

      {/* Header ngang hàng — Lang+Theme đã chuyển vào Sidebar (dùng chung
          cho Mục 1-4), không lặp lại riêng ở đây nữa. */}
      <div className="dashboard-header-grid">
        <div className="dashboard-header-left" style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-2)', fontWeight: 700, whiteSpace: 'nowrap' }}>
          <span aria-hidden="true">🕐</span>
          {formattedTime}
        </div>
        <h1 className="dashboard-header-title">
          {formatReportTitle(
            t.dash2Title,
            selectedDateStart ? parseYYYYMMDD(selectedDateStart) : dateBounds.min,
            selectedDateEnd ? parseYYYYMMDD(selectedDateEnd) : dateBounds.max,
            lang
          ).toUpperCase()}
        </h1>
        <div className="dashboard-header-right" />
      </div>

      {/* Tabs */}
      <div className="tab-container">
        <button className={`tab-btn ${activeTab === 'summary' ? 'active' : ''}`} onClick={() => setActiveTab('summary')}>
          📊 {t.tabMonthlyStatus}
        </button>
        <button className={`tab-btn ${activeTab === 'merged' ? 'active' : ''}`} onClick={() => setActiveTab('merged')}>
          💼 {lang === 'vi' ? 'DOANH SỐ & SẢN LƯỢNG' : lang === 'ko' ? '매출 & 출하량' : 'SALES & VOLUME'}
        </button>
      </div>

      {/* Filter and Cloud bar
          FIX (EPCC-vanilla-toolbar): nền đổi sang "Vanilla" (#FFF4D6) theo
          ảnh tham chiếu — nút ĐANG ĐƯỢC CHỌN (Ngày/Tuần/Tháng...) vẫn giữ
          nguyên màu xanh/teal để phân biệt rõ lựa chọn hiện tại. */}
      <div className="topbar-dash" style={{
        display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '10px',
        background: '#2F3A1D',
        borderRadius: '14px', padding: '10px 14px',
        border: '1px solid rgba(0,0,0,0.18)',
        boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
      }}>
        {/* Dòng 1 (labels) */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          {/* Cụm trái dòng 1: nhãn "Dữ liệu cập nhật đến" — thay cho spacer
              trống trước đây (vùng người dùng khoanh đỏ). Giữ nguyên bề
              rộng 170px để không lệch layout các cột nhãn bên phải. */}
          <div style={{ width: '170px', flexShrink: 0, textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>
            {lang === 'vi' ? 'Dữ liệu cập nhật đến' : lang === 'ko' ? '데이터 업데이트 기준일' : 'Data updated to'}
          </div>
          {/* Cụm giữa dòng 1: Labels */}
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px' }}>
            <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.startDate}</span>
            <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.endDate}</span>
            <span style={{ width: '140px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.customer}</span>
            <span style={{ width: '140px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.colModel}</span>
            <span style={{ width: '180px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{lang === 'vi' ? 'XEM THEO' : lang === 'ko' ? '보기 방식' : 'VIEW BY'}</span>
            {viewMode === 'week' && <span style={{ width: '160px', textAlign: 'center', fontSize: '12px', fontWeight: '700', color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{lang === 'vi' ? 'TUẦN' : lang === 'ko' ? '주차' : 'WEEK'}</span>}
            {showReset && <span style={{ width: '60px', flexShrink: 0 }}></span>}
          </div>
          {/* Cụm phải dòng 1: Spacer khớp ĐÚNG với Dòng 2 (chỉ có 1 nút Tải
              Excel rộng 120px) — trước đây có 3 spacer (120+120+38) trong
              khi Dòng 2 chỉ có 1 nút, khiến cụm giữa (label) và cụm giữa
              (control) co giãn khác nhau → 2 dòng lệch nhau, không đối xứng
              như toolbar chuẩn ở Mục 1 (SalesDashboard). */}
          <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
            <div style={{ width: '120px' }}></div>
          </div>
        </div>

        {/* Dòng 2 (values/controls) */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          {/* Cụm trái dòng 2: giá trị ngày dữ liệu cập nhật gần nhất — lấy
              TRỰC TIẾP từ dữ liệu thật trong database (lastDataUpdateDate),
              không phải giờ đồng hồ hệ thống. */}
          <div style={{ width: '170px', flexShrink: 0, textAlign: 'center', fontSize: '13.5px', fontWeight: 700, color: '#22c55e', whiteSpace: 'nowrap' }}>
            📅 {formattedLastDataUpdate}
          </div>
          {/* Cụm giữa dòng 2: Controls */}
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px', alignItems: 'center' }}>
            <input 
              type="date" 
              value={selectedDateStart} 
              onChange={e => handleStartDateChange(e.target.value)}
              className="filter-date-input"
              style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
            />
            <input 
              type="date" 
              value={selectedDateEnd} 
              onChange={e => handleEndDateChange(e.target.value)}
              className="filter-date-input"
              style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
            />
            <CustomSelect
              value={selectedCustomer}
              onChange={setSelectedCustomer}
              options={[
                { value: '', label: t.allOption },
                ...customers.map(c => ({ value: c, label: c }))
              ]}
              style={{ width: '140px', height: '38px' }}
            />
            <CustomSelect
              value={selectedModel}
              onChange={setSelectedModel}
              options={[
                { value: '', label: t.allOption },
                ...models.map(m => ({ value: m, label: m }))
              ]}
              style={{ width: '140px', height: '38px' }}
            />
            <div style={{ display: 'flex', gap: '0px', height: '38px', width: '180px', flexShrink: 0 }}>
              {(['day', 'week', 'month'] as const).map(mode => (
                <button
                  key={mode}
                  onClick={() => { setViewMode(mode); setSelectedWeek(''); }}
                  style={{
                    flex: 1,
                    padding: '6px 0',
                    fontSize: '13px',
                    fontWeight: 600,
                    borderRadius: mode === 'day' ? '6px 0 0 6px' : mode === 'month' ? '0 6px 6px 0' : '0',
                    // FIX (EPCC-vanilla-toolbar-contrast): màu cố định,
                    // không phụ thuộc theme — tránh chữ mờ/biến mất khi đổi
                    // theme trên nền Vanilla (xem SalesDashboard.tsx cùng fix).
                    border: '1px solid rgba(0,0,0,0.18)',
                    borderRight: mode !== 'month' ? 'none' : '1px solid rgba(0,0,0,0.18)',
                    background: viewMode === mode ? '#2e7d8c' : 'rgba(255,255,255,0.55)',
                    color: viewMode === mode ? '#ffffff' : '#7A5A2E',
                    cursor: 'pointer',
                    transition: 'all 0.15s ease',
                    height: '100%',
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    whiteSpace: 'nowrap'
                  }}
                >
                  {mode === 'day'
                    ? (lang === 'vi' ? 'Ngày' : lang === 'ko' ? '일별' : 'Day')
                    : mode === 'week'
                      ? (lang === 'vi' ? 'Tuần' : lang === 'ko' ? '주별' : 'Week')
                      : (lang === 'vi' ? 'Tháng' : lang === 'ko' ? '월별' : 'Month')
                  }
                </button>
              ))}
            </div>

            {viewMode === 'week' && (
              <CustomSelect
                value={selectedWeek}
                onChange={setSelectedWeek}
                options={[
                  { value: '', label: lang === 'vi' ? 'Tất cả tuần' : lang === 'ko' ? '전체 주' : 'All weeks' },
                  { value: 'W1', label: lang === 'vi' ? 'Tuần 1 (Ngày 1–7)' : lang === 'ko' ? '1주차 (1–7일)' : 'Week 1 (Day 1–7)' },
                  { value: 'W2', label: lang === 'vi' ? 'Tuần 2 (Ngày 8–14)' : lang === 'ko' ? '2주차 (8–14일)' : 'Week 2 (Day 8–14)' },
                  { value: 'W3', label: lang === 'vi' ? 'Tuần 3 (Ngày 15–21)' : lang === 'ko' ? '3주차 (15–21일)' : 'Week 3 (Day 15–21)' },
                  { value: 'W4', label: lang === 'vi' ? 'Tuần 4 (Ngày 22–28)' : lang === 'ko' ? '4주차 (22–28일)' : 'Week 4 (Day 22–28)' },
                  ...(hasW5 ? [{ value: 'W5', label: lang === 'vi' ? 'Tuần 5 (Ngày 29–cuối)' : lang === 'ko' ? '5주차 (29일–말일)' : 'Week 5 (Day 29–end)' }] : [])
                ]}
                style={{ width: '160px', height: '38px' }}
              />
            )}

            {showReset && (
              <NeonButton
                className="btn btn-ghost btn-sm"
                onClick={() => {
                  setSelectedCustomer('');
                  setSelectedModel('');
                  setSelectedWeek('');
                  setSelectedMonth('');
                  if (defaultStartDateStr && defaultEndDateStr) {
                    setSelectedDateStart(defaultStartDateStr);
                    setSelectedDateEnd(defaultEndDateStr);
                  }
                }}
                style={{ height: '38px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}
              >
                {t.resetBtn}
              </NeonButton>
            )}
          </div>
          {/* Cụm phải dòng 2: Buttons */}
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexShrink: 0 }}>
            <input 
              ref={fileInputRef}
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
                    const XLSX = await import('xlsx');
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                    onFileSelected(file, workbook);
                  } catch (err) {
                    alert('Error reading Excel file');
                  }
                };
                reader.readAsArrayBuffer(file);
                e.target.value = '';
              }} 
            />
            <NeonButton
              className="btn btn-outline btn-sm"
              onClick={() => fileInputRef.current?.click()}
              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px', height: '38px', padding: '0 12px', minWidth: '120px', boxSizing: 'border-box', fontSize: '13px' }}
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15" style={{ flexShrink: 0 }}>
                <path d="M21 12a9 9 0 0 1-9 9c-2.52 0-4.93-1-6.74-2.74L3 16" />
                <path d="M3 12a9 9 0 0 1 9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
                <path d="M3 16v5h5" />
                <path d="M16 3h5v5" />
              </svg>
              <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{t.loadExcelBtn}</span>
            </NeonButton>

            {/* Lang select + Theme toggle đã di chuyển lên GlobalHeaderControls
                ở góc trên-phải toàn trang trong App.tsx */}
          </div>
        </div>
      </div>

      {activeTab === 'summary' && (
        <>
          {/* Top 4 statistics cards */}
          {/* FIX (EPCC-kpi-grid-fill-width-muc2): trước đây div này chỉ dùng
              className="kpi-grid" mặc định, không set gridTemplateColumns
              riêng — CSS .kpi-grid mặc định của toàn app tính chỗ cho tối đa
              5 cột (dùng chung với KPIGrid.tsx), nên khi chỉ có 4 card thực
              tế, cột thứ 5 vẫn được grid chừa chỗ nhưng bỏ trống, tạo khoảng
              trắng bên phải hàng KPI. Set cứng gridTemplateColumns:
              'repeat(4,1fr)' (đồng bộ đúng cách PerCapitaTab.tsx /
              ManpowerDashboard.tsx đang làm ở mục 3) để 4 card luôn chia đều
              lấp đầy 100% chiều rộng, không còn ô trống phía phải. */}
          <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '12px', marginBottom: '12px', width: '100%' }}>
            <div className="kpi-card" style={{ borderLeft: '4px solid #2e7d8c', background: 'linear-gradient(135deg, rgba(46,125,140,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#2e7d8c" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                  <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
                  <polyline points="3.27 6.96 12 12.01 20.73 6.96" />
                  <line x1="12" y1="22.08" x2="12" y2="12" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{lang === 'vi' ? 'Sản lượng Thực tế' : 'Actual QTY'}</div>
              </div>
              <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px', flexWrap: 'wrap' }}>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>
                  {Math.round(summaryStats.qtyActual).toLocaleString('vi-VN')}
                  <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85, marginLeft: '1px' }}>K ea</span>
                </div>
                <div className="kpi-card-target">
                  {lang === 'vi' ? `Mục tiêu: ${Math.round(summaryStats.qtyTarget).toLocaleString('vi-VN')}K ea` : `Target: ${Math.round(summaryStats.qtyTarget).toLocaleString('en-US')}K ea`}
                </div>
              </div>
            </div>
            
            <div className="kpi-card" style={{ borderLeft: '4px solid #10b981', background: 'linear-gradient(135deg, rgba(16,185,129,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                  <circle cx="12" cy="12" r="10" />
                  <circle cx="12" cy="12" r="6" />
                  <circle cx="12" cy="12" r="2" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{lang === 'vi' ? 'Tỷ lệ sản lượng' : 'QTY Rate'}</div>
              </div>
              <div className="kpi-card-value" style={{ marginBottom: 0 }}>{summaryStats.qtyRate.toFixed(1)}%</div>
            </div>

            <div className="kpi-card" style={{ borderLeft: '4px solid #8b5cf6', background: 'linear-gradient(135deg, rgba(139,92,246,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#8b5cf6" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                  <rect x="2" y="6" width="20" height="12" rx="2" />
                  <circle cx="12" cy="12" r="2" />
                  <line x1="6" y1="12" x2="6.01" y2="12" />
                  <line x1="18" y1="12" x2="18.01" y2="12" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{lang === 'vi' ? 'Doanh số Thực tế' : 'Actual AMT'}</div>
              </div>
              <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px', flexWrap: 'wrap' }}>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>
                  {Math.round(summaryStats.amtActual).toLocaleString('vi-VN')}
                  <span style={{ fontSize: '60%', fontWeight: 700, opacity: 0.85, marginLeft: '1px' }}>K$</span>
                </div>
                <div className="kpi-card-target">
                  {lang === 'vi' ? `Mục tiêu: ${Math.round(summaryStats.amtTarget).toLocaleString('vi-VN')}K$` : `Target: ${Math.round(summaryStats.amtTarget).toLocaleString('en-US')}K$`}
                </div>
              </div>
            </div>

            <div className="kpi-card" style={{ borderLeft: '4px solid #f59e0b', background: 'linear-gradient(135deg, rgba(245,158,11,0.1) 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: '16px', height: '16px', flexShrink: 0 }}>
                  <polyline points="23 6 13.5 15.5 8.5 10.5 1 18" />
                  <polyline points="17 6 23 6 23 12" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{lang === 'vi' ? 'Tỷ lệ Doanh số' : 'AMT Rate'}</div>
              </div>
              <div className="kpi-card-value" style={{ marginBottom: 0 }}>{summaryStats.amtRate.toFixed(1)}%</div>
            </div>
          </div>

          {/* 6 Charts grid layout */}
          <div className="chart-grid">
            {/* Chart 1: 판매 Q'TY */}
            <div className="panel chart-panel">
              <div className="card-header-styled" style={chartHeaderStyle(0)}>
                <span>{t.salesQtyChartTitle}</span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  <TALegendItem type="bar" color="#E8836B" label={lang === 'vi' ? 'Mục tiêu (Target)' : lang === 'ko' ? '목표 (Target)' : 'Target'} />
                  <TALegendItem type="bar" color="#2e7d8c" label={lang === 'vi' ? 'Thực tế (Actual)' : lang === 'ko' ? '실적 (Actual)' : 'Actual'} />
                  <TALegendItem type="line" color={isLightMode ? '#b45309' : '#eab308'} label="% Rate" />
                </div>
              </div>
              <div className="chart-holder" id="salesQtyChart" style={{ height: '100%' }}></div>
            </div>

            {/* Chart 2: 판매 AMT */}
            <div className="panel chart-panel">
              <div className="card-header-styled" style={chartHeaderStyle(1)}>
                <span>{t.salesAmtChartTitle}</span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  <TALegendItem type="bar" color="#E8836B" label={lang === 'vi' ? 'Mục tiêu (Target)' : lang === 'ko' ? '목표 (Target)' : 'Target'} />
                  <TALegendItem type="bar" color="#2e7d8c" label={lang === 'vi' ? 'Thực tế (Actual)' : lang === 'ko' ? '실적 (Actual)' : 'Actual'} />
                  <TALegendItem type="line" color={isLightMode ? '#b45309' : '#eab308'} label="% Rate" />
                </div>
              </div>
              <div className="chart-holder" id="salesAmtChart" style={{ height: '100%' }}></div>
            </div>

            {/* Chart 3: SALES AMT OF SELECTED MONTH */}
            <div className="panel chart-panel">
              <div className="card-header-styled" style={chartHeaderStyle(2)}>
                <span>
                  {formatReportTitle(
                    t.salesSummaryChartTitle,
                    selectedDateStart ? parseYYYYMMDD(selectedDateStart) : dateBounds.min,
                    selectedDateEnd ? parseYYYYMMDD(selectedDateEnd) : dateBounds.max,
                    lang
                  )}
                </span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  <TALegendItem type="bar" color="#E8836B" label="Target" />
                  <TALegendItem type="bar" color="#2e7d8c" label="Actual" />
                  <TALegendItem type="line" color={isLightMode ? '#b45309' : '#eab308'} label="% Rate" />
                </div>
              </div>
              <div className="chart-holder" id="salesSummaryChart" style={{ height: '100%' }}></div>
            </div>

            {/* Chart 4: % Rate by Model — SVG concentric semi-donut gauge */}
            <div className="panel chart-panel">
              <div className="card-header-styled" style={chartHeaderStyle(3)}>
                <span>% Rate model</span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                  {gaugeData.map((item, idx) => (
                    <TALegendItem
                      key={item.name}
                      type="dot"
                      color={RING_COLORS[idx % RING_COLORS.length]}
                      label={`${item.name} ${Math.round(item.rate)}%`}
                    />
                  ))}
                </div>
              </div>
              <div className="chart-holder" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }}>
                {gaugeData.length > 0 ? (() => {
                  const GREY = theme === 'dark' ? 'rgba(30,41,59,0.5)' : 'rgba(200,200,200,0.6)';
                  const textCol = theme === 'dark' ? '#f3f4f6' : '#1e293b';
                  // FIX (header-legend-merge v2, EPCC): legend model (SO2701 S1
                  // 46%...) đã CHUYỂN HẲN lên header màu (xem TALegendItem ở
                  // JSX phía trên) — bỏ khối "Legend on right" vẽ trong SVG và
                  // mở rộng viewBox full-width + đưa tâm gauge (cx) về giữa,
                  // tận dụng hết phần trống bên phải trước đây dành cho legend.
                  const W = 320, H = 200;
                  const cx = W / 2, cy = 185;
                  const maxR = 170, minR = 65;
                  const n = gaugeData.length;
                  const ringW = n > 1 ? (maxR - minR) / n : 40;
                  const gap = 2;

                  const arcPath = (cx: number, cy: number, outerR: number, innerR: number, rateFrac: number) => {
                    // startAngle = π (9 o'clock left), sweep clockwise by rateFrac * π
                    const startAngle = Math.PI;
                    const endAngle = Math.PI + rateFrac * Math.PI;
                    const fullEndAngle = Math.PI * 2; // half-circle background end = 0 (3 o'clock)

                    const polarToXY = (angle: number, r: number) => ({
                      x: cx + r * Math.cos(angle),
                      y: cy + r * Math.sin(angle)
                    });

                    // Background arc (full semicircle)
                    const bgO1 = polarToXY(startAngle, outerR);
                    const bgO2 = polarToXY(fullEndAngle, outerR);
                    const bgI2 = polarToXY(fullEndAngle, innerR);
                    const bgI1 = polarToXY(startAngle, innerR);
                    const bgPath = [
                      `M ${bgO1.x} ${bgO1.y}`,
                      `A ${outerR} ${outerR} 0 0 1 ${bgO2.x} ${bgO2.y}`,
                      `L ${bgI2.x} ${bgI2.y}`,
                      `A ${innerR} ${innerR} 0 0 0 ${bgI1.x} ${bgI1.y}`,
                      'Z'
                    ].join(' ');

                    if (rateFrac <= 0.001) return { bgPath, fillPath: '' };

                    // Filled arc
                    const fO1 = polarToXY(startAngle, outerR);
                    const fO2 = polarToXY(endAngle, outerR);
                    const fI2 = polarToXY(endAngle, innerR);
                    const fI1 = polarToXY(startAngle, innerR);
                    const largeArc = 0;
                    const fillPath = [
                      `M ${fO1.x} ${fO1.y}`,
                      `A ${outerR} ${outerR} 0 ${largeArc} 1 ${fO2.x} ${fO2.y}`,
                      `L ${fI2.x} ${fI2.y}`,
                      `A ${innerR} ${innerR} 0 ${largeArc} 0 ${fI1.x} ${fI1.y}`,
                      'Z'
                    ].join(' ');

                    return { bgPath, fillPath };
                  };

                  // Label position — midpoint of filled arc
                  const labelPos = (cx: number, cy: number, outerR: number, innerR: number, rateFrac: number) => {
                    const mid = Math.PI + (rateFrac / 2) * Math.PI;
                    const r = (outerR + innerR) / 2;
                    return { x: cx + r * Math.cos(mid), y: cy + r * Math.sin(mid) };
                  };

                  return (
                    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: '100%', height: '100%', overflow: 'visible' }}>
                      {gaugeData.map((item, idx) => {
                        const outerR = maxR - idx * ringW;
                        const innerR = outerR - ringW + gap;
                        const rateFrac = Math.min(Math.max(item.rate / 100, 0), 1);
                        const color = RING_COLORS[idx % RING_COLORS.length];
                        const { bgPath, fillPath } = arcPath(cx, cy, outerR, Math.max(innerR, 10), rateFrac);
                        const lPos = rateFrac >= 0.12 ? labelPos(cx, cy, outerR, Math.max(innerR, 10), rateFrac) : null;
                        return (
                          <g key={item.name}>
                            <path d={bgPath} fill={GREY} />
                            {fillPath && <path d={fillPath} fill={color} />}
                            {lPos && (
                              <text
                                x={lPos.x} y={lPos.y}
                                textAnchor="middle" dominantBaseline="middle"
                                fontSize="10" fontWeight="bold" fill="#fff"
                                style={{ pointerEvents: 'none' }}
                              >
                                {Math.round(item.rate)}%
                              </text>
                            )}
                          </g>
                        );
                      })}
                      {/* Center overall rate label */}
                      <text x={cx} y={cy - 10} textAnchor="middle" fontSize="24" fontWeight="bold" fill={textCol}>
                        {gaugeOverallRate}%
                      </text>
                      <text x={cx} y={cy + 8} textAnchor="middle" fontSize="9.5" fill={textCol} opacity={0.65}>
                        Overall Rate
                      </text>
                    </svg>
                  );
                })() : (
                  <div style={{ color: theme === 'dark' ? '#94a3b8' : '#64748b', fontSize: '13px' }}>Chưa có dữ liệu</div>
                )}
              </div>
            </div>

            {/* Chart 5: Custom (Qty) */}
            <div className="panel chart-panel">
              <div className="card-header-styled" style={chartHeaderStyle(4)}>
                <span>{t.customQtyChartTitle}</span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                  {customQtyLegend.map(item => (
                    <TALegendItem key={item.label} type="dot" color={item.color} label={item.label} />
                  ))}
                </div>
              </div>
              <div className="chart-holder" id="customQtyChart" style={{ height: '100%' }}></div>
            </div>

            {/* Chart 6: Custom (AMT) */}
            <div className="panel chart-panel">
              <div className="card-header-styled" style={chartHeaderStyle(5)}>
                <span>{t.customAmtChartTitle}</span>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                  {customAmtLegend.map(item => (
                    <TALegendItem key={item.label} type="dot" color={item.color} label={item.label} />
                  ))}
                </div>
              </div>
              <div className="chart-holder" id="customAmtChart" style={{ height: '100%' }}></div>
            </div>
          </div>
        </>
      )}

      {/* Table details for active tabs
          FIX (EPCC-merged-table-blank-space v2): panel/table-container
          KHÔNG còn dùng flex:1 để "xin" chiều cao còn lại của cha (cách
          này gây khoảng trắng thừa khi tổng chiều cao cha không đúng như
          giả định). Panel giờ có height TỰ NHIÊN theo nội dung, chỉ khu
          vực table-scroll bên trong có max-height cố định (giống bảng
          tham chiếu "Toàn bộ Cơ sở Dữ liệu từ Excel") để tự cuộn khi
          nhiều dòng, không phụ thuộc DOM cha. */}
      {activeTab === 'merged' && (
        <div className="panel merged-fill-panel" style={{ display: 'flex', flexDirection: 'column' }}>
          {/* FIX: bỏ thanh tiêu đề xanh "Chi tiết Sản lượng & Doanh số bán
              hàng..." theo yêu cầu — kéo bảng lên sát ngay dưới toolbar,
              không còn dòng thừa chiếm chỗ phía trên bảng. */}

          {bottomTableRows.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '40px', color: 'var(--text-3)', fontSize: '14px', fontWeight: '500' }}>
              {lang === 'vi' ? 'Không có dữ liệu phù hợp.' : 'No matching data.'}
            </div>
          ) : (
            <div className="table-container merged-fill-table-container" style={{ display: 'flex', flexDirection: 'column', paddingTop: '0px' }}>
              {/* FIX (EPCC top=0/bottom=0): bỏ hẳn đệm trên (trước là 12px)
                  và đệm dưới (trước là 15px) quanh vùng bảng để bảng "kéo
                  sát" mép trên (ngay dưới toolbar) và mép dưới (ngay trên
                  thanh phân trang) — tăng tối đa số dòng hiển thị được
                  trong cùng maxHeight: 62vh. Padding trái/phải (15px) vẫn
                  giữ nguyên để nội dung không dính sát viền ngang. thead
                  vẫn sticky top:0 nên không có khe hở lộ dòng cũ khi cuộn. */}
              <div ref={tableScrollRef} className="table-scroll" style={{ overflowY: 'auto', overflowX: 'auto', maxHeight: `${tableScrollMaxHeight}px`, padding: '0 15px 0 15px' }}>
                <table className="stat-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, marginTop: 0 }}>
                  <thead ref={theadRef} style={{ position: 'sticky', top: 0, zIndex: 20 }}>
                    {/* Group header row */}
                    <tr style={{ background: 'var(--surface)', position: 'sticky', top: 0, zIndex: 20 }}>
                      <th colSpan={3} style={{ borderBottom: 'none', background: 'var(--surface)', position: 'sticky', top: 0, zIndex: 20 }}></th>
                      <th colSpan={3} style={{ textAlign: 'center', background: isLightMode ? 'rgba(46, 125, 140, 0.95)' : 'rgba(46, 125, 140, 0.85)', color: '#ffffff', fontWeight: '800', fontSize: '16.5px', letterSpacing: '0.05em', textTransform: 'uppercase', padding: '10px 14px', textShadow: isLightMode ? '0 1px 2px rgba(0,0,0,0.4)' : 'none', position: 'sticky', top: 0, zIndex: 20 }}>
                        {lang === 'vi' ? 'SẢN LƯỢNG (QTY)' : lang === 'ko' ? '출하량 (QTY)' : 'VOLUME (QTY)'}
                      </th>
                      <th colSpan={3} style={{ textAlign: 'center', background: isLightMode ? 'rgba(29, 78, 216, 0.95)' : 'rgba(29, 78, 216, 0.85)', color: '#ffffff', fontWeight: '800', fontSize: '16.5px', letterSpacing: '0.05em', textTransform: 'uppercase', padding: '10px 14px', textShadow: isLightMode ? '0 1px 2px rgba(0,0,0,0.4)' : 'none', position: 'sticky', top: 0, zIndex: 20 }}>
                        {lang === 'vi' ? 'DOANH SỐ (AMT $)' : lang === 'ko' ? '매출액 (AMT $)' : 'SALES (AMT $)'}
                      </th>
                    </tr>
                    {/* Column header row with sort + filter */}
                    <tr style={{ background: 'var(--surface)', position: 'sticky', top: '43px', zIndex: 20 }}>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'left', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                          <span className="sort-header" onClick={() => { setSortField('model'); setSortAsc(sortField === 'model' ? !sortAsc : true); }}>
                            MODEL {sortField === 'model' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                          </span>
                          <select value={selectedModel} onChange={e => setSelectedModel(e.target.value)} className="header-filter-select" style={{ fontSize: '16.5px' }}>
                            <option value="">{lang === 'vi' ? 'Tất cả Model' : 'All Models'}</option>
                            {models.map(m => <option key={m} value={m}>{m}</option>)}
                          </select>
                        </div>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'left', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                          <span className="sort-header" onClick={() => { setSortField('customer'); setSortAsc(sortField === 'customer' ? !sortAsc : true); }}>
                            {lang === 'vi' ? 'KHÁCH HÀNG' : 'CUSTOMER'} {sortField === 'customer' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                          </span>
                          <select value={selectedCustomer} onChange={e => setSelectedCustomer(e.target.value)} className="header-filter-select" style={{ fontSize: '16.5px' }}>
                            <option value="">{lang === 'vi' ? 'Tất cả khách' : 'All Customers'}</option>
                            {customers.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'center', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'center' }}>
                          <span className="sort-header" onClick={() => { setSortField('date'); setSortAsc(sortField === 'date' ? !sortAsc : true); }}>
                            {viewMode === 'week' ? (lang === 'vi' ? 'TUẦN' : 'WEEK') : viewMode === 'month' ? (lang === 'vi' ? 'THÁNG' : 'MONTH') : (lang === 'vi' ? 'NGÀY' : 'DATE')} {sortField === 'date' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                          </span>
                          {viewMode === 'week' ? (
                            <select value={selectedWeek} onChange={e => setSelectedWeek(e.target.value)} className="header-filter-select" style={{ width: '110px', fontSize: '16.5px' }}>
                              <option value="">{lang === 'vi' ? 'Tất cả tuần' : 'All Weeks'}</option>
                              <option value="W1">{lang === 'vi' ? 'Tuần 1' : 'Week 1'}</option>
                              <option value="W2">{lang === 'vi' ? 'Tuần 2' : 'Week 2'}</option>
                              <option value="W3">{lang === 'vi' ? 'Tuần 3' : 'Week 3'}</option>
                              <option value="W4">{lang === 'vi' ? 'Tuần 4' : 'Week 4'}</option>
                              {hasW5 && <option value="W5">{lang === 'vi' ? 'Tuần 5' : 'Week 5'}</option>}
                            </select>
                          ) : (
                            <select value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)} className="header-filter-select" style={{ width: '110px', fontSize: '16.5px' }}>
                              <option value="">{lang === 'vi' ? 'Tất cả tháng' : 'All Months'}</option>
                              {monthsList.map(m => <option key={m} value={m}>{m}</option>)}
                            </select>
                          )}
                        </div>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'right', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                          <span className="sort-header" onClick={() => { setSortField('qtyTarget'); setSortAsc(sortField === 'qtyTarget' ? !sortAsc : true); }}>
                            QTY TARGET {sortField === 'qtyTarget' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                          </span>
                          <select value={qtyTargetFilter} onChange={e => setQtyTargetFilter(e.target.value)} className="header-filter-select" style={{ width: '100px', fontSize: '16.5px' }}>
                            <option value="all">{lang === 'vi' ? 'Tất cả' : 'All'}</option>
                            <option value="greater_than_zero">{lang === 'vi' ? 'Có (> 0)' : 'Has (> 0)'}</option>
                            <option value="equal_to_zero">{lang === 'vi' ? 'Bằng 0' : 'Equals 0'}</option>
                          </select>
                        </div>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'right', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                          <span className="sort-header" onClick={() => { setSortField('qtyActual'); setSortAsc(sortField === 'qtyActual' ? !sortAsc : true); }}>
                            QTY ACTUAL {sortField === 'qtyActual' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                          </span>
                          <select value={qtyActualFilter} onChange={e => setQtyActualFilter(e.target.value)} className="header-filter-select" style={{ width: '100px', fontSize: '16.5px' }}>
                            <option value="all">{lang === 'vi' ? 'Tất cả' : 'All'}</option>
                            <option value="greater_than_zero">{lang === 'vi' ? 'Có (> 0)' : 'Has (> 0)'}</option>
                            <option value="equal_to_zero">{lang === 'vi' ? 'Bằng 0' : 'Equals 0'}</option>
                          </select>
                        </div>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'right', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <span className="sort-header" onClick={() => { setSortField('qtyRatio'); setSortAsc(sortField === 'qtyRatio' ? !sortAsc : true); }}>
                          {lang === 'vi' ? 'TỈ LỆ QTY' : 'QTY %'} {sortField === 'qtyRatio' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                        </span>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'right', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                          <span className="sort-header" onClick={() => { setSortField('amtTarget'); setSortAsc(sortField === 'amtTarget' ? !sortAsc : true); }}>
                            AMT TARGET {sortField === 'amtTarget' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                          </span>
                          <select value={amtTargetFilter} onChange={e => setAmtTargetFilter(e.target.value)} className="header-filter-select" style={{ width: '100px', fontSize: '16.5px' }}>
                            <option value="all">{lang === 'vi' ? 'Tất cả' : 'All'}</option>
                            <option value="greater_than_zero">{lang === 'vi' ? 'Có (> 0)' : 'Has (> 0)'}</option>
                            <option value="equal_to_zero">{lang === 'vi' ? 'Bằng 0' : 'Equals 0'}</option>
                          </select>
                        </div>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'right', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', alignItems: 'flex-end' }}>
                          <span className="sort-header" onClick={() => { setSortField('amtActual'); setSortAsc(sortField === 'amtActual' ? !sortAsc : true); }}>
                            AMT ACTUAL {sortField === 'amtActual' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                          </span>
                          <select value={amtActualFilter} onChange={e => setAmtActualFilter(e.target.value)} className="header-filter-select" style={{ width: '100px', fontSize: '16.5px' }}>
                            <option value="all">{lang === 'vi' ? 'Tất cả' : 'All'}</option>
                            <option value="greater_than_zero">{lang === 'vi' ? 'Có (> 0)' : 'Has (> 0)'}</option>
                            <option value="equal_to_zero">{lang === 'vi' ? 'Bằng 0' : 'Equals 0'}</option>
                          </select>
                        </div>
                      </th>
                      <th style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: 'right', color: 'var(--tbl-head-color)', position: 'sticky', top: '43px', zIndex: 20, background: 'var(--surface)' }}>
                        <span className="sort-header" onClick={() => { setSortField('amtRatio'); setSortAsc(sortField === 'amtRatio' ? !sortAsc : true); }}>
                          {lang === 'vi' ? 'TỈ LỆ AMT' : 'AMT %'} {sortField === 'amtRatio' ? (sortAsc ? '🔼' : '🔽') : '↕️'}
                        </span>
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {pagedRows.map((row, index) => (
                      <tr key={index} style={{ borderBottom: '1px solid var(--border)' }}>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, color: 'var(--tbl-model-color)', fontWeight: 600, textAlign: 'left' }}>{row.model}</td>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, color: 'var(--tbl-cell-color)', fontWeight: 600, textAlign: 'left' }}>{row.customer}</td>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, color: 'var(--tbl-date-color)', fontWeight: 500, textAlign: 'center' }}>{row.period}</td>
                        {/* FIX (EPCC-merged-table-blank-space v2): tham chiếu
                            bảng "Toàn bộ Cơ sở Dữ liệu từ Excel" — bảng đó
                            LUÔN hiển thị giá trị thật (kể cả 0, vd "0 K")
                            thay vì dấu "-". Áp dụng cùng cách cho bảng này:
                            bỏ điều kiện ẩn giá trị 0 thành "-", chỉ còn dùng
                            màu nhạt (--tbl-nil-color) để phân biệt trực quan
                            khi giá trị = 0, không còn ô trống gây hiểu nhầm
                            là thiếu dữ liệu. */}
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, color: row.qtyTarget > 0 ? 'var(--tbl-num-color)' : 'var(--tbl-nil-color)', textAlign: 'right' }}>
                          {Math.round(row.qtyTarget).toLocaleString('vi-VN')}
                        </td>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, color: row.qtyActual > 0 ? 'var(--green)' : 'var(--tbl-nil-color)', fontWeight: row.qtyActual > 0 ? 600 : 400, textAlign: 'right' }}>
                          {Math.round(row.qtyActual).toLocaleString('vi-VN')}
                        </td>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right', fontWeight: 700, color: row.qtyTarget > 0 ? (row.qtyRatio < 100 ? 'red' : 'green') : 'var(--tbl-nil-color)' }}>
                          {row.qtyTarget > 0 ? `${row.qtyRatio.toFixed(1)}%` : '0.0%'}
                        </td>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, color: row.amtTarget > 0 ? 'var(--tbl-num-color)' : 'var(--tbl-nil-color)', textAlign: 'right' }}>
                          {'$' + Math.round(row.amtTarget).toLocaleString('vi-VN')}
                        </td>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, color: row.amtActual > 0 ? 'var(--purple-light)' : 'var(--tbl-nil-color)', fontWeight: row.amtActual > 0 ? 600 : 400, textAlign: 'right' }}>
                          {'$' + Math.round(row.amtActual).toLocaleString('vi-VN')}
                        </td>
                        <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right', fontWeight: 700, color: row.amtTarget > 0 ? (row.amtRatio < 100 ? 'red' : 'green') : 'var(--tbl-nil-color)' }}>
                          {row.amtTarget > 0 ? `${row.amtRatio.toFixed(1)}%` : '0.0%'}
                        </td>
                      </tr>
                    ))}
                    <tr className="total-row">
                      <td className="text-left" colSpan={3} style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, fontWeight: 800 }}>
                        {lang === 'vi' ? 'TỔNG' : lang === 'ko' ? '합계' : 'TOTAL'}
                      </td>
                      <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right' }}>{Math.round(tableTotals.qtyTarget).toLocaleString('vi-VN')}</td>
                      <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right' }}>{Math.round(tableTotals.qtyActual).toLocaleString('vi-VN')}</td>
                      <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right', color: '#3b82f6', fontWeight: 800 }}>{tableTotals.qtyRatio.toFixed(1)}%</td>
                      <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right' }}>${Math.round(tableTotals.amtTarget).toLocaleString('vi-VN')}</td>
                      <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right' }}>${Math.round(tableTotals.amtActual).toLocaleString('vi-VN')}</td>
                      <td style={{ fontSize: `${rowFontSize}px`, padding: `${rowPaddingY}px 12px`, textAlign: 'right', color: '#3b82f6', fontWeight: 800 }}>{tableTotals.amtRatio.toFixed(1)}%</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Pagination bar */}
              <div ref={paginationBarRef} className="pagination-bar" style={{ padding: '10px 15px', background: 'var(--surface)', borderTop: '1px solid var(--border-soft)', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap', flexShrink: 0 }}>
                <div className="pagination-info">
                  {lang === 'vi' 
                    ? `Hiển thị ${processedRows.length > 0 ? (safePage - 1) * pageSize + 1 : 0}-${Math.min(safePage * pageSize, processedRows.length)} / ${processedRows.length} dòng`
                    : lang === 'ko'
                      ? `표시 ${processedRows.length > 0 ? (safePage - 1) * pageSize + 1 : 0}-${Math.min(safePage * pageSize, processedRows.length)} / ${processedRows.length} 행`
                      : `Showing ${processedRows.length > 0 ? (safePage - 1) * pageSize + 1 : 0}-${Math.min(safePage * pageSize, processedRows.length)} / ${processedRows.length} rows`
                  }
                </div>
                <div className="pagination-controls">
                  <button className="page-btn" onClick={() => setPage(1)} disabled={safePage === 1}>«</button>
                  <button className="page-btn" onClick={() => setPage(p => Math.max(1, p - 1))} disabled={safePage === 1}>‹</button>
                  {pageNums.map((n, i) => n === '...'
                    ? <span key={`ellipsis-${i}`} style={{ padding: '0 4px', color: 'var(--text-3)' }}>…</span>
                    : <button key={n} className={`page-btn ${safePage === n ? 'active' : ''}`} onClick={() => setPage(n as number)}>{n}</button>
                  )}
                  <button className="page-btn" onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={safePage === totalPages}>›</button>
                  <button className="page-btn" onClick={() => setPage(totalPages)} disabled={safePage === totalPages}>»</button>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ fontSize: 11.5, color: 'var(--text-2)' }}>
                    {lang === 'vi' ? 'Dòng/trang:' : lang === 'ko' ? '페이지당 행 수:' : 'Rows/page:'}
                  </span>
                  <select
                    className="page-size-select"
                    value={pageSize}
                    onChange={e => setPageSize(Number(e.target.value))}
                  >
                    {PAGE_SIZES.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

    </div>
  );
};
