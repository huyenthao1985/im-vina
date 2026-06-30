import React from 'react';

interface SidebarProps {
  activeViewId: string;
  onSelectView: (id: string) => void;
  collapsed: boolean;
  onToggleCollapse: () => void;
  lang: 'vi' | 'en' | 'ko';
}

const ITEMS = [
  {
    id: 'overview',
    index: '1',
    name: { vi: 'Tổng quan SX - Xuất hàng - Doanh số', en: 'Production - Shipment - Sales', ko: '생산-출하-매출 개요' },
  },
  {
    id: 'target_actual',
    index: '2',
    name: { vi: 'Báo cáo doanh số (Target/Actual)', en: 'Sales Target vs Actual', ko: '매출 목표 대비 실적' },
  },
  {
    // Mục 3 = Manpower + ProductionPerCapita gộp lại, dùng tab bên trong
    id: 'manpower',
    index: '3',
    name: { vi: '인당 생산수 & 근무 인력 현황', en: 'Production Per Capita & Manpower', ko: '인당 생산수 & 근무 인력 현황' },
  },
  {
    id: 'placeholder',
    index: '4',
    name: { vi: 'Cấu hình & Khác', en: 'Config & Others', ko: '설정 및 기타' },
  },
];

export const Sidebar: React.FC<SidebarProps> = ({
  activeViewId,
  onSelectView,
  collapsed,
  onToggleCollapse,
  lang,
}) => {
  return (
    <aside className={`sidebar ${collapsed ? 'collapsed' : ''}`}>
      <div className="sidebar-header">
        <div className="sidebar-logo">
          <span style={{ fontSize: '18px' }}>📊</span>
          <span style={{ fontWeight: 800, letterSpacing: '0.5px' }}>IM-VINA MENU</span>
        </div>
        <button
          className="sidebar-toggle-btn"
          onClick={onToggleCollapse}
          title={collapsed ? (lang === 'vi' ? 'Mở rộng' : 'Expand') : (lang === 'vi' ? 'Thu gọn' : 'Collapse')}
        >
          {collapsed ? '➡️' : '⬅️'}
        </button>
      </div>

      <ul className="sidebar-menu">
        {ITEMS.map((item) => {
          const isActive = activeViewId === item.id || (item.id === 'manpower' && activeViewId === 'production_per_capita');
          const label = item.name[lang];
          return (
            <li
              key={item.id}
              className={`sidebar-item ${isActive ? 'active' : ''}`}
              onClick={() => onSelectView(item.id)}
              title={collapsed ? `${item.index}. ${label}` : undefined}
            >
              <div className="sidebar-item-index">{item.index}</div>
              <span className="sidebar-item-label">{label}</span>
              <div className="sidebar-item-dot" />
            </li>
          );
        })}
      </ul>
    </aside>
  );
};
