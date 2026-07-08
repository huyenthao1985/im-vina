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
          icon đăng xuất để không phá layout thu gọn. */}
      <div style={{
        borderTop: '1px solid rgba(255,255,255,0.08)',
        padding: collapsed ? '10px 8px' : '12px',
        display: 'flex', flexDirection: 'column', gap: '8px', flexShrink: 0,
      }}>
        <div style={{
          display: 'flex', alignItems: 'center', gap: '8px',
          padding: collapsed ? '4px' : '6px 8px',
          borderRadius: '10px', background: 'rgba(255,255,255,0.06)',
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
                fontSize: '13px', fontWeight: 700, color: '#fff',
                whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
              }}>
                {profile.full_name || profile.email}
              </span>
              <span style={{
                fontSize: '11px', color: 'rgba(255,255,255,0.6)',
                whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis',
              }}>
                {profile.email} · {
                  profile.role === 'admin' ? (lang === 'vi' ? 'Quản trị viên' : lang === 'en' ? 'Admin' : '관리자')
                  : profile.role === 'editor' ? (lang === 'vi' ? 'Biên tập viên' : lang === 'en' ? 'Editor' : '편집자')
                  : (lang === 'vi' ? 'Người dùng' : lang === 'en' ? 'User' : '사용자')
                }
              </span>
            </div>
          )}
        </div>

        {profile.role === 'admin' && (
          <button
            onClick={onOpenAdmin}
            style={{
              padding: '7px 10px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.12)',
              background: 'rgba(255,255,255,0.06)', color: '#fff', fontWeight: 600,
              fontSize: '13px', cursor: 'pointer', textAlign: collapsed ? 'center' : 'left',
            }}
            title={lang === 'vi' ? 'Quản trị' : lang === 'en' ? 'Admin' : '관리자'}
          >
            {collapsed ? '⚙️' : (lang === 'vi' ? '⚙️ Quản trị' : lang === 'en' ? '⚙️ Admin' : '⚙️ 관리자')}
          </button>
        )}
        <button
          onClick={onSignOut}
          style={{
            padding: '7px 10px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.12)',
            background: 'rgba(255,255,255,0.06)', color: '#fff', fontWeight: 600,
            fontSize: '13px', cursor: 'pointer', textAlign: collapsed ? 'center' : 'left',
          }}
          title={lang === 'vi' ? 'Đăng xuất' : lang === 'en' ? 'Sign out' : '로그아웃'}
        >
          {collapsed ? '🚪' : (lang === 'vi' ? 'Đăng xuất' : lang === 'en' ? 'Sign out' : '로그아웃')}
        </button>
      </div>
    </aside>
  );
};
