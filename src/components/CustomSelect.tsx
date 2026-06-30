/**
 * CustomSelect — theme-aware dropdown that replaces native <select>
 * so we can fully control the option-list colors in both Dark & Light mode.
 *
 * Usage:
 *   <CustomSelect
 *     value={lang}
 *     onChange={v => setLang(v as any)}
 *     options={[
 *       { value: 'vi', label: 'Tiếng Việt' },
 *       { value: 'en', label: 'English' },
 *       { value: 'ko', label: '한국어' },
 *     ]}
 *     className="lang-select"   // optional extra class
 *     style={{ minWidth: 130 }} // optional extra style
 *   />
 */

import React, { useState, useRef, useEffect, useCallback } from 'react';

export interface SelectOption {
  value: string;
  label: string;
}

interface CustomSelectProps {
  value: string;
  onChange: (value: string) => void;
  options: SelectOption[];
  className?: string;
  style?: React.CSSProperties;
  /** If true the trigger button shows a caret. Default true. */
  showCaret?: boolean;
  placeholder?: string;
}

export const CustomSelect: React.FC<CustomSelectProps> = ({
  value,
  onChange,
  options,
  className = '',
  style,
  showCaret = true,
  placeholder,
}) => {
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  const selectedLabel =
    options.find(o => o.value === value)?.label ?? placeholder ?? value;

  const handleSelect = useCallback(
    (v: string) => {
      onChange(v);
      setOpen(false);
    },
    [onChange],
  );

  // Close when clicking outside
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [open]);

  // Close on Escape
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [open]);

  return (
    <div
      ref={containerRef}
      className={`csel-wrapper ${className}`}
      style={{ position: 'relative', display: 'inline-block', ...style }}
    >
      {/* Trigger */}
      <button
        type="button"
        className={`csel-trigger ${open ? 'csel-trigger--open' : ''}`}
        onClick={() => setOpen(o => !o)}
        aria-haspopup="listbox"
        aria-expanded={open}
      >
        <span className="csel-trigger__label">{selectedLabel}</span>
        {showCaret && (
          <span className={`csel-caret ${open ? 'csel-caret--up' : ''}`} aria-hidden>
            ▾
          </span>
        )}
      </button>

      {/* Dropdown list */}
      {open && (
        <ul
          className="csel-menu"
          role="listbox"
          aria-label="options"
        >
          {options.map(opt => (
            <li
              key={opt.value}
              role="option"
              aria-selected={opt.value === value}
              className={`csel-option ${opt.value === value ? 'csel-option--selected' : ''}`}
              onMouseDown={() => handleSelect(opt.value)}
            >
              {opt.value === value && <span className="csel-option__check">✓</span>}
              {opt.label}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};
