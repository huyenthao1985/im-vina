import type { FilterState, ColumnMapping } from '../types';
import { toDateInputValue } from '../utils';
import { NeonButton } from './NeonButton';

interface FilterBarProps {
  filters: FilterState;
  onFilterChange: (f: FilterState) => void;
  categories: string[];
  minDate: Date | null;
  maxDate: Date | null;
  mapping: ColumnMapping;
  activeCount: number;
  onReset: () => void;
}

const NONE = '__none__';

const QUICK_PRESETS = [
  { label: '7 ngày', days: 7 },
  { label: '30 ngày', days: 30 },
  { label: '90 ngày', days: 90 },
];

export const FilterBar: React.FC<FilterBarProps> = ({
  filters,
  onFilterChange,
  categories,
  minDate,
  maxDate,
  mapping,
  activeCount,
  onReset,
}) => {
  const hasDate = mapping.dateCol !== NONE;
  const hasCategory = mapping.categoryCol !== NONE;

  const applyQuickPreset = (days: number) => {
    if (!maxDate) return;
    const end = new Date(maxDate);
    const start = new Date(end);
    start.setDate(start.getDate() - days + 1);
    onFilterChange({
      ...filters,
      dateStart: toDateInputValue(start),
      dateEnd: toDateInputValue(end),
    });
  };

  const toggleCategory = (cat: string) => {
    const selected = filters.categories;
    if (selected.includes(cat)) {
      onFilterChange({ ...filters, categories: selected.filter(c => c !== cat) });
    } else {
      onFilterChange({ ...filters, categories: [...selected, cat] });
    }
  };

  return (
    <div className="filter-bar">
      <span className="filter-label">🔍 Lọc</span>

      {hasDate && (
        <>
          <input
            type="date"
            className="date-input"
            value={filters.dateStart}
            min={minDate ? toDateInputValue(minDate) : undefined}
            max={filters.dateEnd || (maxDate ? toDateInputValue(maxDate) : undefined)}
            onChange={e => onFilterChange({ ...filters, dateStart: e.target.value })}
          />
          <span className="date-sep">→</span>
          <input
            type="date"
            className="date-input"
            value={filters.dateEnd}
            min={filters.dateStart || (minDate ? toDateInputValue(minDate) : undefined)}
            max={maxDate ? toDateInputValue(maxDate) : undefined}
            onChange={e => onFilterChange({ ...filters, dateEnd: e.target.value })}
          />

          <div className="filter-sep" />
          <div className="quick-filters">
            {QUICK_PRESETS.map(p => {
              const isActive = (() => {
                if (!maxDate || !filters.dateStart || !filters.dateEnd) return false;
                const start = new Date(maxDate);
                start.setDate(start.getDate() - p.days + 1);
                return filters.dateStart === toDateInputValue(start) && filters.dateEnd === toDateInputValue(maxDate);
              })();
              return (
                <button
                  key={p.label}
                  className={`quick-filter-btn ${isActive ? 'active' : ''}`}
                  onClick={() => applyQuickPreset(p.days)}
                >
                  {p.label}
                </button>
              );
            })}
          </div>
        </>
      )}

      {hasDate && hasCategory && <div className="filter-sep" />}

      {hasCategory && categories.length > 0 && (
        <div className="quick-filters">
          {categories.slice(0, 8).map(cat => (
            <button
              key={cat}
              className={`quick-filter-btn ${filters.categories.includes(cat) ? 'active' : ''}`}
              onClick={() => toggleCategory(cat)}
            >
              {cat}
            </button>
          ))}
          {categories.length > 8 && (
            <span style={{ fontSize: 11, color: 'var(--text-3)', paddingLeft: 4 }}>
              +{categories.length - 8} khác
            </span>
          )}
        </div>
      )}

      <div className="filter-right">
        {activeCount > 0 && (
          <>
            <span className="active-filter-count">{activeCount}</span>
            <NeonButton className="btn btn-ghost btn-sm" onClick={onReset}>✕ Xóa lọc</NeonButton>
          </>
        )}
      </div>
    </div>
  );
};
