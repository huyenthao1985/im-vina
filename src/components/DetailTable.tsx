import { useState, useMemo, useRef, useEffect } from 'react';
import type { DataRow, SortState, ColumnMapping } from '../types';
import { formatCellValue, exportToCSV, toNumber, fmtVND } from '../utils';
import * as XLSX from 'xlsx';

interface DetailTableProps {
  rows: DataRow[];
  headers: string[];
  mapping: ColumnMapping;
}

const PAGE_SIZES = [10, 25, 50, 100];
const NONE = '__none__';

export const DetailTable: React.FC<DetailTableProps> = ({ rows, headers, mapping }) => {
  const [search, setSearch] = useState('');
  const [sort, setSort] = useState<SortState>({ column: null, direction: null });
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);
  const [showExport, setShowExport] = useState(false);
  const exportRef = useRef<HTMLDivElement>(null);
  const [columnFilters, setColumnFilters] = useState<Record<string, string>>({});

  // Reset column filters when sheet headers change
  useEffect(() => {
    setColumnFilters({});
  }, [headers]);

  const getUniqueValues = (col: string) => {
    const vals = rows
      .map(r => r[col])
      .filter(v => v !== null && v !== undefined && v !== '')
      .map(v => String(v));
    return Array.from(new Set(vals)).sort();
  };

  const handleColumnFilterChange = (col: string, val: string) => {
    setColumnFilters(prev => ({ ...prev, [col]: val }));
    setPage(1);
  };

  // Close export dropdown on outside click
  const handleExportToggle = () => setShowExport(v => !v);

  // Filter
  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    let result = rows;
    if (q) {
      result = result.filter(r => headers.some(h => formatCellValue(r[h]).toLowerCase().includes(q)));
    }
    
    // Apply column filters
    Object.entries(columnFilters).forEach(([col, filterVal]) => {
      if (!filterVal) return;
      const fv = filterVal.toLowerCase();
      if (col === '__profit__') {
        result = result.filter(r => {
          const profit = toNumber(r[mapping.revenueCol]) - toNumber(r[mapping.costCol]);
          return String(profit).toLowerCase().includes(fv);
        });
      } else {
        result = result.filter(r => {
          const val = formatCellValue(r[col]).toLowerCase();
          return val.includes(fv);
        });
      }
    });
    
    return result;
  }, [rows, headers, search, columnFilters, mapping]);

  // Sort
  const sorted = useMemo(() => {
    if (!sort.column || !sort.direction) return filtered;
    return [...filtered].sort((a, b) => {
      const va = a[sort.column!];
      const vb = b[sort.column!];
      
      const isNumA = va !== null && va !== undefined && va !== '' && !isNaN(parseFloat(va)) && isFinite(Number(va));
      const isNumB = vb !== null && vb !== undefined && vb !== '' && !isNaN(parseFloat(vb)) && isFinite(Number(vb));
      
      if (isNumA && isNumB) {
        const numA = toNumber(va);
        const numB = toNumber(vb);
        return sort.direction === 'asc' ? numA - numB : numB - numA;
      }
      
      const sa = formatCellValue(va).toLowerCase();
      const sb = formatCellValue(vb).toLowerCase();
      return sort.direction === 'asc' ? sa.localeCompare(sb) : sb.localeCompare(sa);
    });
  }, [filtered, sort]);

  // Paginate
  const totalPages = Math.max(1, Math.ceil(sorted.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const paged = useMemo(() => {
    const start = (safePage - 1) * pageSize;
    return sorted.slice(start, start + pageSize);
  }, [sorted, safePage, pageSize]);

  const handleSort = (col: string) => {
    setSort(prev => {
      if (prev.column !== col) return { column: col, direction: 'asc' };
      if (prev.direction === 'asc') return { column: col, direction: 'desc' };
      return { column: null, direction: null };
    });
    setPage(1);
  };

  const handleSearch = (val: string) => { setSearch(val); setPage(1); };
  const handlePageSize = (val: number) => { setPageSize(val); setPage(1); };

  // Export
  const handleExportCSV = () => {
    exportToCSV(filtered, headers, `marketing_data_${Date.now()}.csv`);
    setShowExport(false);
  };

  const handleExportExcel = () => {
    const ws = XLSX.utils.json_to_sheet(
      filtered.map(r => {
        const out: any = {};
        headers.forEach(h => { out[h] = r[h] instanceof Date ? r[h].toLocaleDateString('vi-VN') : r[h]; });
        return out;
      })
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    XLSX.writeFile(wb, `marketing_data_${Date.now()}.xlsx`);
    setShowExport(false);
  };

  // Profit/ROI column
  const hasRevenue = mapping.revenueCol !== NONE;
  const hasCost = mapping.costCol !== NONE;

  const isNumeric = (v: any) => {
    if (typeof v === 'number') return true;
    if (typeof v === 'string' && v.trim() !== '' && !isNaN(parseFloat(v)) && isFinite(Number(v))) return true;
    return false;
  };

  const isYearValue = (v: any) => {
    const n = toNumber(v);
    return n >= 1900 && n <= 2100 && Number.isInteger(n);
  };

  const getCellStyle = (header: string, value: any) => {
    const isYear = ['year', 'năm', 'nam'].some(k => header.toLowerCase().includes(k)) || isYearValue(value);
    const isNum = isNumeric(value);
    if (header === mapping.revenueCol && isNum && !isYear) {
      return { color: 'var(--cyan)', fontVariantNumeric: 'tabular-nums', textAlign: 'right' as const };
    }
    if (header === mapping.costCol && isNum && !isYear) {
      return { color: 'var(--amber)', fontVariantNumeric: 'tabular-nums', textAlign: 'right' as const };
    }
    if (isNum && !isYear) {
      return { fontVariantNumeric: 'tabular-nums', textAlign: 'right' as const };
    }
    return {};
  };

  const getSortIcon = (col: string) => {
    if (sort.column !== col) return <span className="sort-icon">⇅</span>;
    return <span className="sort-icon">{sort.direction === 'asc' ? '↑' : '↓'}</span>;
  };

  // Page numbers to show
  const pageNums = useMemo(() => {
    const nums: (number | '...')[] = [];
    if (totalPages <= 7) {
      for (let i = 1; i <= totalPages; i++) nums.push(i);
    } else {
      nums.push(1);
      if (safePage > 3) nums.push('...');
      for (let i = Math.max(2, safePage - 1); i <= Math.min(totalPages - 1, safePage + 1); i++) nums.push(i);
      if (safePage < totalPages - 2) nums.push('...');
      nums.push(totalPages);
    }
    return nums;
  }, [totalPages, safePage]);

  return (
    <div className="table-panel">
      <div className="table-toolbar">
        <div className="table-toolbar-left">
          <span className="table-title">📋 Dữ liệu chi tiết</span>
          <span className="table-count-chip">{filtered.length.toLocaleString('vi-VN')} dòng</span>
        </div>
        <div className="table-toolbar-right">
          <input
            id="searchBox"
            className="search-box"
            type="text"
            placeholder="🔍 Tìm kiếm..."
            value={search}
            onChange={e => handleSearch(e.target.value)}
          />

          <div className="export-menu" ref={exportRef}>
            <button className="btn btn-secondary btn-sm" onClick={handleExportToggle}>
              ⬇️ Xuất
            </button>
            {showExport && (
              <div className="export-dropdown">
                <button className="export-dropdown-item" onClick={handleExportExcel}>
                  📊 Excel (.xlsx)
                </button>
                <button className="export-dropdown-item" onClick={handleExportCSV}>
                  📄 CSV (.csv)
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="table-scroll">
        <table>
          <thead>
            <tr>
              {headers.map(h => (
                <th
                  key={h}
                  className={`sortable ${sort.column === h ? 'sorted' : ''}`}
                  onClick={() => handleSort(h)}
                >
                  {h} {getSortIcon(h)}
                </th>
              ))}
              {hasRevenue && hasCost && (
                <th style={{ textAlign: 'right' }}>Lợi nhuận {getSortIcon('__profit__')}</th>
              )}
            </tr>
            <tr className="filter-row">
              {headers.map(h => {
                const uniqueVals = getUniqueValues(h);
                const isCategorical = uniqueVals.length > 0 && uniqueVals.length <= 25;
                return (
                  <th key={`filter-${h}`} style={{ padding: '6px 8px' }}>
                    {isCategorical ? (
                      <select
                        className="column-filter-select"
                        value={columnFilters[h] || ''}
                        onChange={e => handleColumnFilterChange(h, e.target.value)}
                        onClick={e => e.stopPropagation()}
                      >
                        <option value="">Tất cả ({uniqueVals.length})</option>
                        {uniqueVals.map(v => (
                          <option key={v} value={v}>{v}</option>
                        ))}
                      </select>
                    ) : (
                      <input
                        type="text"
                        className="column-filter-input"
                        placeholder={`Lọc ${h}...`}
                        value={columnFilters[h] || ''}
                        onChange={e => handleColumnFilterChange(h, e.target.value)}
                        onClick={e => e.stopPropagation()}
                      />
                    )}
                  </th>
                );
              })}
              {hasRevenue && hasCost && (
                <th key="filter-profit" style={{ padding: '6px 8px' }}>
                  <input
                    type="text"
                    className="column-filter-input"
                    placeholder="Lọc..."
                    value={columnFilters['__profit__'] || ''}
                    onChange={e => handleColumnFilterChange('__profit__', e.target.value)}
                    onClick={e => e.stopPropagation()}
                  />
                </th>
              )}
            </tr>
          </thead>
          <tbody>
            {paged.length === 0 ? (
              <tr>
                <td colSpan={headers.length + (hasRevenue && hasCost ? 1 : 0)} style={{ textAlign: 'center', padding: '40px', color: 'var(--text-3)' }}>
                  Không tìm thấy dữ liệu phù hợp
                </td>
              </tr>
            ) : paged.map((row, ri) => {
              const profit = hasRevenue && hasCost
                ? toNumber(row[mapping.revenueCol]) - toNumber(row[mapping.costCol])
                : null;
              return (
                <tr key={ri}>
                  {headers.map(h => (
                    <td key={h} style={getCellStyle(h, row[h])}>
                      {(h === mapping.revenueCol || h === mapping.costCol) && 
                       isNumeric(row[h]) &&
                       !['year', 'năm', 'nam'].some(k => h.toLowerCase().includes(k)) &&
                       !isYearValue(row[h])
                        ? fmtVND(toNumber(row[h]), mapping.currency)
                        : formatCellValue(row[h])}
                    </td>
                  ))}
                  {profit !== null && (
                    <td style={{
                      textAlign: 'right',
                      fontVariantNumeric: 'tabular-nums',
                      color: profit >= 0 ? 'var(--emerald)' : 'var(--rose)',
                      fontWeight: 600,
                    }}>
                      {profit >= 0 ? '+' : ''}{fmtVND(profit, mapping.currency)}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="pagination-bar">
        <div className="pagination-info">
          Trang {safePage} / {totalPages} · Hiển thị {paged.length} / {filtered.length} dòng
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
          <span style={{ fontSize: 11.5, color: 'var(--text-2)' }}>Dòng/trang:</span>
          <select
            className="page-size-select"
            value={pageSize}
            onChange={e => handlePageSize(Number(e.target.value))}
          >
            {PAGE_SIZES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
      </div>
    </div>
  );
};
