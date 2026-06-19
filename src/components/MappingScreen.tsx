import React, { useState, useEffect, useMemo } from 'react';
import type { ColumnMapping, ColumnType, DataRow, FilterState } from '../types';
import { parseRowToDate } from '../utils';

interface MappingScreenProps {
  filename: string;
  fileSize: number;
  sheetNames: string[];
  currentSheet: string;
  headers: string[];
  columnTypes: ColumnType[];
  mapping: ColumnMapping;
  onSheetChange: (sheet: string) => void;
  onMappingChange: (mapping: ColumnMapping) => void;
  onConfirm: (dateStart?: string, dateEnd?: string, selectedCategories?: string[]) => void;
  onBack: () => void;
  theme: string;
  onToggleTheme: () => void;
  rows: DataRow[];
  initialFilters: FilterState;
}

const NONE_VALUE = '__none__';

export const MappingScreen: React.FC<MappingScreenProps> = ({
  filename,
  fileSize,
  sheetNames,
  currentSheet,
  headers,
  // columnTypes removed from destructuring
  mapping,
  onSheetChange,
  onMappingChange,
  onConfirm,
  onBack,
  theme,
  onToggleTheme,
  rows,
  initialFilters,
}) => {


  // Compute date bounds from rows for the selected dateCol using parseRowToDate
  const dateBounds = useMemo(() => {
    if (mapping.dateCol === NONE_VALUE || !rows.length) return { min: '', max: '' };
    const dates = rows
      .map(r => parseRowToDate(r, mapping.dateCol))
      .filter((d): d is Date => d !== null);
    if (!dates.length) return { min: '', max: '' };
    const minD = new Date(Math.min(...dates.map(d => d.getTime())));
    const maxD = new Date(Math.max(...dates.map(d => d.getTime())));
    return {
      min: minD.toISOString().slice(0, 10),
      max: maxD.toISOString().slice(0, 10)
    };
  }, [rows, mapping.dateCol]);

  // Compute distinct categories from rows for the selected categoryCol
  const distinctCategories = useMemo(() => {
    if (mapping.categoryCol === NONE_VALUE || !rows.length) return [];
    return Array.from(
      new Set(
        rows
          .map(r => r[mapping.categoryCol])
          .filter(v => v !== null && v !== undefined && v !== '')
          .map(String)
      )
    ).sort();
  }, [rows, mapping.categoryCol]);

  // Local state for filters
  const [selectedDateStart, setSelectedDateStart] = useState('');
  const [selectedDateEnd, setSelectedDateEnd] = useState('');
  const [selectedCats, setSelectedCats] = useState<string[]>([]);

  // Reset/Initialize dates
  useEffect(() => {
    if (mapping.dateCol === NONE_VALUE) {
      setSelectedDateStart('');
      setSelectedDateEnd('');
    } else if (dateBounds.min && dateBounds.max) {
      if (initialFilters.dateStart && initialFilters.dateStart >= dateBounds.min && initialFilters.dateStart <= dateBounds.max) {
        setSelectedDateStart(initialFilters.dateStart);
      } else {
        setSelectedDateStart(dateBounds.min);
      }

      if (initialFilters.dateEnd && initialFilters.dateEnd >= dateBounds.min && initialFilters.dateEnd <= dateBounds.max) {
        setSelectedDateEnd(initialFilters.dateEnd);
      } else {
        setSelectedDateEnd(dateBounds.max);
      }
    } else {
      setSelectedDateStart(initialFilters.dateStart || '');
      setSelectedDateEnd(initialFilters.dateEnd || '');
    }
  }, [mapping.dateCol, dateBounds, initialFilters.dateStart, initialFilters.dateEnd]);

  // Reset/Initialize categories
  useEffect(() => {
    if (mapping.categoryCol === NONE_VALUE) {
      setSelectedCats([]);
    } else {
      if (initialFilters.categories.length > 0 && initialFilters.categories.every(c => distinctCategories.includes(c))) {
        setSelectedCats(initialFilters.categories);
      } else {
        setSelectedCats(distinctCategories);
      }
    }
  }, [mapping.categoryCol, distinctCategories, initialFilters.categories]);

  // Extract years dynamically from rows for the selected dateCol using parseRowToDate
  const availableYears = useMemo(() => {
    if (!rows.length || mapping.dateCol === NONE_VALUE) {
      return Array.from({ length: 15 }, (_, i) => String(2015 + i));
    }
    const years = rows
      .map(r => {
        const d = parseRowToDate(r, mapping.dateCol);
        return d ? d.getFullYear() : null;
      })
      .filter((y): y is number => y !== null);
    const unique = Array.from(new Set(years)).sort((a, b) => a - b);
    if (unique.length > 0) return unique.map(String);
    return Array.from({ length: 15 }, (_, i) => String(2015 + i));
  }, [rows, mapping.dateCol]);

  const monthOptions = Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0'));
  const dayOptions = Array.from({ length: 31 }, (_, i) => String(i + 1).padStart(2, '0'));

  // Split selected dates
  const { startYear, startMonth, startDay } = useMemo(() => {
    if (!selectedDateStart) return { startYear: '', startMonth: '', startDay: '' };
    const parts = selectedDateStart.split('-');
    return {
      startYear: parts[0] || '',
      startMonth: parts[1] || '',
      startDay: parts[2] || ''
    };
  }, [selectedDateStart]);

  const { endYear, endMonth, endDay } = useMemo(() => {
    if (!selectedDateEnd) return { endYear: '', endMonth: '', endDay: '' };
    const parts = selectedDateEnd.split('-');
    return {
      endYear: parts[0] || '',
      endMonth: parts[1] || '',
      endDay: parts[2] || ''
    };
  }, [selectedDateEnd]);

  // Date change handlers
  const handleStartYearChange = (val: string) => {
    setSelectedDateStart(`${val}-${startMonth || '01'}-${startDay || '01'}`);
  };
  const handleStartMonthChange = (val: string) => {
    setSelectedDateStart(`${startYear || '2026'}-${val}-${startDay || '01'}`);
  };
  const handleStartDayChange = (val: string) => {
    setSelectedDateStart(`${startYear || '2026'}-${startMonth || '01'}-${val}`);
  };

  const handleEndYearChange = (val: string) => {
    setSelectedDateEnd(`${val}-${endMonth || '12'}-${endDay || '31'}`);
  };
  const handleEndMonthChange = (val: string) => {
    setSelectedDateEnd(`${endYear || '2026'}-${val}-${endDay || '31'}`);
  };
  const handleEndDayChange = (val: string) => {
    setSelectedDateEnd(`${endYear || '2026'}-${endMonth || '12'}-${val}`);
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  const canConfirm = mapping.dateCol !== NONE_VALUE || mapping.categoryCol !== NONE_VALUE || mapping.revenueCol !== NONE_VALUE;

  return (
    <>
      <header className="topbar">
        <div className="topbar-left">
          <div className="brandmark">
            <div className="brandmark-dot">📊</div>
            <span className="brandmark-text">Marketing Insights</span>
          </div>
        </div>
        <div className="topbar-right">
          <button className="btn btn-ghost btn-sm" onClick={onBack}>← Quay lại</button>
          <button className="theme-toggle" onClick={onToggleTheme}>
            {theme === 'dark' ? '☀️' : '🌙'}
          </button>
        </div>
      </header>

      <main className="mapping-screen">
        <div className="mapping-card">
          <div className="mapping-header">
            <div>
              <h2>Cấu hình Dữ liệu</h2>
              <p>Chọn cột tương ứng cho từng chiều phân tích</p>
            </div>
            <div className="file-info-chip">
              <span className="icon">📊</span>
              <div>
                <div className="name">{filename}</div>
                <div className="meta">{formatFileSize(fileSize)} · {sheetNames.length} sheet</div>
              </div>
            </div>
          </div>

          {sheetNames.length > 1 && (
            <div className="sheet-select-row">
              <label>📋 Sheet đang dùng:</label>
              <select
                className="select-field"
                value={currentSheet}
                onChange={e => onSheetChange(e.target.value)}
              >
                {sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          )}

          <div className="mapping-grid">
            <div className="mapping-field field-date">
              <label>
                <span>📅</span>
                Cột Thời gian
              </label>
              <select
                className="select-field"
                value={mapping.dateCol}
                onChange={e => onMappingChange({ ...mapping, dateCol: e.target.value })}
              >
                <option value={NONE_VALUE}>— Không có —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>

            <div className="mapping-field field-category">
              <label>
                <span>🏷️</span>
                Cột Kênh / Danh mục
              </label>
              <select
                className="select-field"
                value={mapping.categoryCol}
                onChange={e => onMappingChange({ ...mapping, categoryCol: e.target.value })}
              >
                <option value={NONE_VALUE}>— Không có —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>

            <div className="mapping-field field-revenue">
              <label>
                <span>💰</span>
                Cột Doanh thu
              </label>
              <select
                className="select-field"
                value={mapping.revenueCol}
                onChange={e => onMappingChange({ ...mapping, revenueCol: e.target.value })}
              >
                <option value={NONE_VALUE}>— Không có —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>

            <div className="mapping-field field-cost">
              <label>
                <span>📉</span>
                Cột Chi phí
              </label>
              <select
                className="select-field"
                value={mapping.costCol}
                onChange={e => onMappingChange({ ...mapping, costCol: e.target.value })}
              >
                <option value={NONE_VALUE}>— Không có —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>

            <div className="mapping-field field-currency">
              <label>
                <span>💱</span>
                Đơn vị Tiền tệ
              </label>
              <select
                className="select-field"
                value={mapping.currency || 'VND'}
                onChange={e => onMappingChange({ ...mapping, currency: e.target.value })}
              >
                <option value="VND">VNĐ (đ)</option>
                <option value="USD">USD ($)</option>
              </select>
            </div>
          </div>

          {/* Dynamic Configuration Section */}
          {(mapping.dateCol !== NONE_VALUE || mapping.categoryCol !== NONE_VALUE) && (
            <div className="mapping-advanced-filters">
              <h3 className="section-title">⚙️ Thiết lập Lọc dữ liệu ban đầu</h3>
              
              {mapping.dateCol !== NONE_VALUE && (
                <div className="advanced-field-row date-selection-row">
                  <div className="field-title">📅 Khoảng thời gian phân tích:</div>
                  <div className="date-picker-inputs" style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <div className="date-input-container" style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                      <span className="label" style={{ minWidth: '80px', fontWeight: '500' }}>Từ ngày:</span>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <select className="select-field date-select-part" value={startYear} onChange={e => handleStartYearChange(e.target.value)} style={{ width: '90px' }}>
                          {availableYears.map(y => <option key={y} value={y}>{y}</option>)}
                        </select>
                        <select className="select-field date-select-part" value={startMonth} onChange={e => handleStartMonthChange(e.target.value)} style={{ width: '100px' }}>
                          {monthOptions.map(m => <option key={m} value={m}>Tháng {m}</option>)}
                        </select>
                        <select className="select-field date-select-part" value={startDay} onChange={e => handleStartDayChange(e.target.value)} style={{ width: '90px' }}>
                          {dayOptions.map(d => <option key={d} value={d}>Ngày {d}</option>)}
                        </select>
                      </div>
                    </div>
                    
                    <div className="date-input-container" style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                      <span className="label" style={{ minWidth: '80px', fontWeight: '500' }}>Đến ngày:</span>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <select className="select-field date-select-part" value={endYear} onChange={e => handleEndYearChange(e.target.value)} style={{ width: '90px' }}>
                          {availableYears.map(y => <option key={y} value={y}>{y}</option>)}
                        </select>
                        <select className="select-field date-select-part" value={endMonth} onChange={e => handleEndMonthChange(e.target.value)} style={{ width: '100px' }}>
                          {monthOptions.map(m => <option key={m} value={m}>Tháng {m}</option>)}
                        </select>
                        <select className="select-field date-select-part" value={endDay} onChange={e => handleEndDayChange(e.target.value)} style={{ width: '90px' }}>
                          {dayOptions.map(d => <option key={d} value={d}>Ngày {d}</option>)}
                        </select>
                      </div>
                    </div>
                  </div>
                  {dateBounds.min && dateBounds.max && (
                    <div className="date-bounds-hint" style={{ marginTop: '8px', fontSize: '11px', color: 'var(--text-3)' }}>
                      Khoảng ngày phát hiện được: {new Date(dateBounds.min).toLocaleDateString('vi-VN')} - {new Date(dateBounds.max).toLocaleDateString('vi-VN')}
                    </div>
                  )}
                </div>
              )}

            </div>
          )}
          {/* Category selection and Debug info removed at user's request */}

          <div className="mapping-footer">
            <div className="mapping-note">
              ℹ️ Chỉ cần map ít nhất 1 cột để bắt đầu
            </div>
            <div style={{ display: 'flex', gap: 10 }}>
              <button className="btn btn-secondary" onClick={onBack}>← Quay lại</button>
              <button
                className="btn btn-primary"
                onClick={() => onConfirm(selectedDateStart, selectedDateEnd, selectedCats)}
                disabled={!canConfirm}
                style={{ opacity: canConfirm ? 1 : 0.5 }}
              >
                Bắt đầu Phân tích →
              </button>
            </div>
          </div>
        </div>
      </main>
    </>
  );
};
