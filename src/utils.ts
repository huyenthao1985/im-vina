import type { DataRow, ColumnType } from './types';

export function parseToDate(v: any): Date | null {
  if (v instanceof Date) return v;
  if (v === null || v === undefined || v === '') return null;
  
  if (typeof v === 'number') {
    if (v >= 1970 && v <= 2100) {
      return new Date(v, 0, 1);
    }
    if (v >= 30000 && v <= 60000) {
      const utc_days = Math.floor(v - 25569);
      const utc_value = utc_days * 86400;
      return new Date(utc_value * 1000);
    }
    return null;
  }
  
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  
  const ddmmyyyy = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (ddmmyyyy) {
    const day = parseInt(ddmmyyyy[1], 10);
    const month = parseInt(ddmmyyyy[2], 10) - 1;
    const year = parseInt(ddmmyyyy[3], 10);
    const d = new Date(year, month, day);
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

const MONTH_MAP: Record<string, number> = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
};

export function parseRowToDate(row: DataRow, dateColName: string): Date | null {
  const v = row[dateColName];
  if (v instanceof Date) return v;
  if (v === null || v === undefined || v === '') return null;

  const hasYearInHeader = dateColName.toLowerCase().includes('year');
  if (hasYearInHeader) {
    const yearNum = parseInt(v, 10);
    if (!isNaN(yearNum) && yearNum >= 1970 && yearNum <= 2100) {
      const monthKey = Object.keys(row).find(k => k.toLowerCase().includes('month') || k.toLowerCase().includes('tháng') || k.toLowerCase().includes('thang'));
      if (monthKey) {
        const monthVal = String(row[monthKey]).toLowerCase().trim();
        const m = MONTH_MAP[monthVal] !== undefined ? MONTH_MAP[monthVal] : 0;
        return new Date(yearNum, m, 1);
      }
      return new Date(yearNum, 0, 1);
    }
  }

  return parseToDate(v);
}

export function detectColumnType(rows: DataRow[], header: string): 'text' | 'numeric' | 'date' {
  const vals = rows.map(r => r[header]).filter(v => v !== null && v !== undefined && v !== '');
  if (vals.length === 0) return 'text';
  let dateCount = 0, numCount = 0;
  vals.forEach(v => {
    if (v instanceof Date) dateCount++;
    else if (typeof v === 'number') numCount++;
    else if (typeof v === 'string' && v.trim() !== '' && !isNaN(parseFloat(v)) && isFinite(Number(v))) numCount++;
  });
  if (dateCount / vals.length > 0.7) return 'date';
  if (numCount / vals.length > 0.7) return 'numeric';
  return 'text';
}

export function detectColumns(rows: DataRow[], headers: string[]): ColumnType[] {
  return headers.map(h => ({ name: h, type: detectColumnType(rows, h) }));
}

export function toNumber(v: any): number {
  if (typeof v === 'number') return v;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

export function fmtVND(n: number, currency: string = 'VND'): string {
  const abs = Math.abs(n);
  const isNeg = n < 0;
  if (currency === 'USD') {
    const prefix = isNeg ? '-$' : '$';
    if (abs >= 1e9) return prefix + (abs / 1e9).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + ' B';
    if (abs >= 1e6) return prefix + (abs / 1e6).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + ' M';
    return prefix + abs.toLocaleString('vi-VN', { maximumFractionDigits: 0 });
  } else {
    // VND
    const suffix = ' VNĐ';
    let valStr = '';
    if (abs >= 1e9) valStr = (abs / 1e9).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + ' tỷ';
    else if (abs >= 1e6) valStr = (abs / 1e6).toLocaleString('vi-VN', { maximumFractionDigits: 1 }) + ' tr';
    else valStr = abs.toLocaleString('vi-VN', { maximumFractionDigits: 0 });
    return (isNeg ? '-' : '') + valStr + suffix;
  }
}

export function fmtNum(n: number): string {
  return n.toLocaleString('vi-VN', { maximumFractionDigits: 2 });
}

export function fmtPct(n: number): string {
  return (n >= 0 ? '+' : '') + n.toFixed(1) + '%';
}

export function formatCellValue(v: any): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toLocaleDateString('vi-VN');
  
  const num = Number(v);
  if (typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && !isNaN(num))) {
    if (num >= 1900 && num <= 2100 && Number.isInteger(num)) {
      return String(num);
    }
  }
  
  if (typeof v === 'number') {
    return v.toLocaleString('vi-VN', { maximumFractionDigits: 2 });
  }
  return String(v);
}

export function escapeHtml(str: string): string {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Get date boundaries for filtered rows
export function getDateBounds(rows: DataRow[], dateColName: string): { min: Date | null; max: Date | null } {
  const dates = rows
    .map(r => parseRowToDate(r, dateColName))
    .filter((d): d is Date => d !== null);
  if (!dates.length) return { min: null, max: null };
  return {
    min: new Date(Math.min(...dates.map(d => d.getTime()))),
    max: new Date(Math.max(...dates.map(d => d.getTime()))),
  };
}

export function toDateInputValue(d: Date): string {
  return d.toISOString().slice(0, 10);
}

// Channel colors
const PALETTE_LIST = [
  { bg: 'rgba(139, 92, 246, 0.15)', color: '#a78bfa' },
  { bg: 'rgba(34, 211, 238, 0.15)', color: '#22d3ee' },
  { bg: 'rgba(16, 185, 129, 0.15)', color: '#10b981' },
  { bg: 'rgba(245, 158, 11, 0.15)', color: '#f59e0b' },
  { bg: 'rgba(244, 63, 94, 0.15)', color: '#f43f5e' },
  { bg: 'rgba(129, 140, 248, 0.15)', color: '#818cf8' },
  { bg: 'rgba(251, 146, 60, 0.15)', color: '#fb923c' },
  { bg: 'rgba(52, 211, 153, 0.15)', color: '#34d399' },
];

const channelColorMap = new Map<string, { bg: string; color: string }>();
let paletteIndex = 0;

export function getChannelColor(name: string): { bg: string; color: string } {
  if (channelColorMap.has(name)) return channelColorMap.get(name)!;
  const c = PALETTE_LIST[paletteIndex % PALETTE_LIST.length];
  paletteIndex++;
  channelColorMap.set(name, c);
  return c;
}

export function resetChannelColors(): void {
  channelColorMap.clear();
  paletteIndex = 0;
}

export const CHART_COLORS = [
  '#8b5cf6', '#22d3ee', '#10b981', '#f59e0b',
  '#f43f5e', '#818cf8', '#fb923c', '#34d399',
  '#a78bfa', '#67e8f9',
];

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export function exportToCSV(rows: DataRow[], headers: string[], filename: string): void {
  const csvContent = [
    headers.join(','),
    ...rows.map(r => headers.map(h => {
      const v = formatCellValue(r[h]);
      return `"${v.replace(/"/g, '""')}"`;
    }).join(',')),
  ].join('\n');
  const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
  downloadBlob(blob, filename);
}
