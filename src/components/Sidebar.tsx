import React from 'react';

interface SidebarProfile {
  full_name: string;
  email: string;
  role: 'user' | 'editor' | 'admin' | null;
}

interface SidebarProps {
  activeViewId: string;
  onSelectView: (id: string) => void;
  collapsed: boolean;
  onToggleCollapse: () => void;
  lang: 'vi' | 'en' | 'ko';
  // FIX (user-info-in-sidebar): thông tin người đăng nhập + hành động, hiển
  // thị cố định ở chân Sidebar thay vì nổi rời trong App.tsx.
  profile: SidebarProfile;
  onSignOut: () => void;
  onOpenAdmin: () => void;
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
    name: { vi: 'Sản lượng đầu người & Nhân lực', en: 'Production Per Capita & Manpower', ko: '인당 생산수 & 근무 인력 현황' },
  },
  {
    id: 'placeholder',
    index: '4',
    name: { vi: 'Cấu hình & Khác', en: 'Config & Others', ko: '설정 및 기타' },
  },
];

// Icon SVG kiểu Lucide (stroke="currentColor" — tự ăn theo màu chữ của thẻ
// cha, luôn đúng theme sáng/tối, không cần set màu riêng). Icon "log-out" là
// icon chuẩn dev hay dùng cho hành động đăng xuất (mũi tên thoát khỏi khung
// cửa), "settings" (bánh răng) chuẩn cho hành động quản trị/cấu hình.
function LogOutIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
      <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
      <polyline points="16 17 21 12 16 7" />
      <line x1="21" y1="12" x2="9" y2="12" />
    </svg>
  );
}

function SettingsIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
      <circle cx="12" cy="12" r="3" />
      <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
    </svg>
  );
}

export const Sidebar: React.FC<SidebarProps> = ({
  activeViewId,
  onSelectView,
  collapsed,
  onToggleCollapse,
  lang,
  profile,
  onSignOut,
  onOpenAdmin,
}) => {
  return (
    <aside
      className={`sidebar ${collapsed ? 'collapsed' : ''}`}
      style={{ display: 'flex', flexDirection: 'column', height: '100%' }}
    >
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

      <ul className="sidebar-menu" style={{ flex: 1, overflowY: 'auto' }}>
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

      {/* FIX (user-info-in-sidebar): chân Sidebar — thông tin người đăng nhập
          + nút Quản trị (chỉ admin) + Đăng xuất. Thay thế cho overlay nổi cũ
          trong App.tsx. Ở trạng thái collapsed chỉ hiện avatar viết tắt +
          icon đăng xuất để không phá layout thu gọn. Dùng biến CSS theme-aware
          (--text-0, --text-2, --border-soft) để tự đổi màu đúng theo
          theme sáng/tối, tránh chữ trắng biến mất trên nền sáng. */}
      <div style={{
        borderTop: '1px solid var(--border-soft, rgba(255,255,255,0.08))',
        padding: collapsed ? '10px 8px' : '12px',
        display: 'flex', flexDirection: 'column', gap: '8px', flexShrink: 0,
      }}>
        <div style={{
          display: 'flex', alignItems: 'center', gap: '8px',
          padding: collapsed ? '5px' : '7px 8px',
          borderRadius: '10px',
          border: '1px solid rgba(120,120,130,0.35)',
          background: 'rgba(120,120,130,0.16)',
          boxShadow: '0 1px 2px rgba(0,0,0,0.06)',
        }}>
          <div style={{
            width: '28px', height: '28px', borderRadius: '50%', flexShrink: 0,
            background: 'var(--primary, #6366f1)', color: '#fff',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            fontWeight: 700, fontSize: '12px', textTransform: 'uppercase',
          }}>
            {(profile.full_name || profile.email || '?').charAt(0)}
          </div>
          {!collapsed && (
            <div style={{ display: 'flex', flexDirection: 'column', lineHeight: 1.25, minWidth: 0 }}>
              <span style={{
                fontSize: '13px', fontWeight: 700, color: 'var(--text-0, #111827)',
                whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
              }}>
                {profile.full_name || profile.email}
              </span>
              <span style={{
                fontSize: '11px', color: 'var(--text-2, #6b7280)',
                whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
              }}>
                {profile.email}
              </span>
            </div>
          )}
        </div>

        {profile.role === 'admin' && (
          <button
            onClick={onOpenAdmin}
            style={{
              padding: '8px 10px', borderRadius: '8px',
              border: '1px solid rgba(120,120,130,0.35)',
              background: 'rgba(120,120,130,0.16)',
              boxShadow: '0 1px 2px rgba(0,0,0,0.06)',
              color: 'var(--text-0, #111827)', fontWeight: 600,
              fontSize: '13px', cursor: 'pointer',
              display: 'flex', alignItems: 'center',
              justifyContent: collapsed ? 'center' : 'flex-start',
              gap: '8px',
            }}
            title={lang === 'vi' ? 'Quản trị' : lang === 'en' ? 'Admin' : '관리자'}
          >
            <SettingsIcon />
            {!collapsed && (lang === 'vi' ? 'Quản trị' : lang === 'en' ? 'Admin' : '관리자')}
          </button>
        )}
        <button
          onClick={onSignOut}
          style={{
            padding: '8px 10px', borderRadius: '8px',
            border: '1px solid rgba(120,120,130,0.35)',
            background: 'rgba(120,120,130,0.16)',
            boxShadow: '0 1px 2px rgba(0,0,0,0.06)',
            color: 'var(--text-0, #111827)', fontWeight: 600,
            fontSize: '13px', cursor: 'pointer',
            display: 'flex', alignItems: 'center',
            justifyContent: collapsed ? 'center' : 'flex-start',
            gap: '8px',
          }}
          title={lang === 'vi' ? 'Đăng xuất' : lang === 'en' ? 'Sign out' : '로그아웃'}
        >
          <LogOutIcon />
          {!collapsed && (lang === 'vi' ? 'Đăng xuất' : lang === 'en' ? 'Sign out' : '로그아웃')}
        </button>
      </div>
    </aside>
  );
};
