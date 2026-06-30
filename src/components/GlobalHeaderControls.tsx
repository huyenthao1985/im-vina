/**
 * GlobalHeaderControls.tsx
 *
 * Cụm control TOÀN CỤC: dropdown ngôn ngữ + nút Dark/Light.
 * Mount DUY NHẤT 1 lần ở App.tsx trong một sticky header bar —
 * KHÔNG dùng position:fixed để tránh đè lên nội dung khi scroll.
 *
 * KHÔNG đổi state: lang/theme vẫn sống ở App.tsx như cũ, component này
 * chỉ là phần UI render — các tab con vẫn nhận `lang`/`theme` qua props.
 */

import React from 'react';

interface GlobalHeaderControlsProps {
  lang: 'vi' | 'en' | 'ko';
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
  isDark: boolean;
  onToggleTheme: () => void;
}

export const GlobalHeaderControls: React.FC<GlobalHeaderControlsProps> = ({
  lang, setLang, isDark, onToggleTheme,
}) => {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
      {/* Lang select */}
      <select
        value={lang}
        onChange={e => setLang(e.target.value as any)}
        className="lang-select"
        style={{ height: '34px' }}
      >
        <option value="vi">Tiếng Việt</option>
        <option value="en">English</option>
        <option value="ko">한국어</option>
      </select>

      {/* Theme toggle */}
      <button
        className="theme-toggle"
        onClick={onToggleTheme}
        style={{
          position: 'static', height: '34px', width: '34px',
          padding: 0, display: 'inline-flex',
          alignItems: 'center', justifyContent: 'center',
        }}
        title={isDark ? 'Chuyển sang Sáng' : 'Chuyển sang Tối'}
      >
        {isDark ? '☀️' : '🌙'}
      </button>
    </div>
  );
};

export default GlobalHeaderControls;
