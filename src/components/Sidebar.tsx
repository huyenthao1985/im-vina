import React from 'react';
// FIX (sidebar-bg-camera): ảnh nền camera exploded-view cho Sidebar. Đặt file
// ảnh thật tại src/assets/sidebar-bg-camera.png (copy từ file đã upload) rồi
// import như dưới đây. Nếu bạn để ảnh ở thư mục public/ thay vì import, thay
// biến BG_IMAGE_URL bằng chuỗi đường dẫn tĩnh, ví dụ: '/sidebar-bg-camera.png'.
import sidebarBgImage from '../assets/sidebar-bg-camera.png';

const BG_IMAGE_URL = sidebarBgImage;

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
  setLang: (lang: 'vi' | 'en' | 'ko') => void;
  // FIX (sidebar-utility-widgets): đồng hồ realtime + chọn ngôn ngữ/theme
  // trước đây lặp lại ở header riêng của TỪNG Dashboard con (Mục 1/2/3) qua
  // <GlobalHeaderControls>. Giờ gộp về hiển thị DUY NHẤT 1 lần ở đầu Sidebar
  // (phía trên Mục 1), áp dụng chung cho toàn bộ Mục 1-4 — cần thêm theme +
  // onToggleTheme để Sidebar tự vẽ nút Sáng/Tối tại đây.
  theme: 'dark' | 'light';
  onToggleTheme: () => void;
  // FIX (user-info-in-sidebar): thông tin người đăng nhập + hành động, hiển
  // thị cố định ở chân Sidebar thay vì nổi rời trong App.tsx.
  profile: SidebarProfile;
  onSignOut: () => void;
  onOpenAdmin: () => void;
}

// EPCC (remove-menu1-overview) - FIX theo yêu cầu người dùng "xóa Mục 1 vì đã
// có trong Mục 5": xóa hẳn item 'overview' (trước là index '1') — nội dung
// "Tổng quan xuất hàng - doanh số" nay đã có trong tab "TÌNH HÌNH DOANH THU"
// của Mục 5 (Menu5db.tsx). Đánh số lại 4 mục còn lại từ 1-4, không còn nhảy
// cóc số. Không còn nơi nào khác trong Sidebar.tsx tham chiếu id 'overview'.
const ITEMS = [
  {
    id: 'target_actual',
    index: '1',
    name: { vi: 'DOANH SỐ THEO THÁNG', en: 'Sales Target vs Actual', ko: '매출 목표 대비 실적' },
  },
  {
    // Mục 3 = Manpower + ProductionPerCapita gộp lại, dùng tab bên trong
    id: 'manpower',
    index: '2',
    name: { vi: 'SẢN LƯỢNG ĐẦU NGƯỜI & NHÂN LỰC', en: 'Production Per Capita & Manpower', ko: '인당 생산수 & 근무 인력 현황' },
  },
  {
    id: 'placeholder',
    index: '3',
    name: { vi: 'HIỆU SUẤT RTY', en: 'RTY Performance', ko: 'RTY 성과' },
  },
  {
    // FIX (add-template-dashboard-menu): mục mới trỏ tới DashboardTemplate.tsx
    // (file mẫu EPCC) — id thật khớp CHÍNH XÁC với điều kiện
    // `activeViewId === '...'` đang thêm trong App.tsx.
    // EPCC (menu5-real-id-fcost) - FIX theo yêu cầu người dùng "đổi luôn
    // trong app": dashboard này đã thực sự là Menu5db (F-Cost), không còn là
    // trang mẫu DashboardTemplate nữa — đổi id 'my_new_dashboard' → 'fcost'
    // (id thật), khớp CHÍNH XÁC với điều kiện `activeViewId === 'fcost'` mới
    // sửa trong App.tsx. Không có nơi nào khác trong Sidebar.tsx/App.tsx
    // tham chiếu tới id cũ ngoài 2 chỗ này (đã rà soát trước khi đổi).
    id: 'fcost',
    index: '4',
    name: { vi: 'DOANH SỐ - FCOST', en: 'Sales - F-Cost', ko: '매출 - F-Cost' },
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

// FIX (toggle-btn-3d): mũi tên thu gọn/mở rộng — trước dùng emoji ➡️/⬅️ (dẹt,
// không đồng bộ style). Đổi sang icon SVG chevron mảnh kiểu Lucide, xoay
// 180° bằng CSS transform tùy trạng thái collapsed thay vì 2 icon riêng.
function ChevronIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, display: 'block' }}>
      <polyline points="15 18 9 12 15 6" />
    </svg>
  );
}

export const Sidebar: React.FC<SidebarProps> = ({
  activeViewId,
  onSelectView,
  collapsed,
  onToggleCollapse,
  lang,
  setLang,
  theme,
  onToggleTheme,
  profile,
  onSignOut,
  onOpenAdmin,
}) => {
  return (
    <aside
      className={`sidebar ${collapsed ? 'collapsed' : ''}`}
      style={{ display: 'flex', flexDirection: 'column', height: '100%' }}
    >
      {/* FIX (sidebar-bg-camera v2): TRƯỚC ĐÂY set position:'relative' +
          overflow:'hidden' NGAY TRÊN <aside> đã vô tình GHI ĐÈ position
          (vd position:'fixed'/height:100vh) mà class CSS ".sidebar" gốc
          đang dùng để giữ Sidebar cao hết màn hình → Sidebar bị co lại
          theo nội dung, đẩy khối Quản trị/Đăng xuất lên giữa thay vì dính
          đáy. Fix: KHÔNG đụng style của <aside>, dời position:'relative'
          sang div bọc NGAY BÊN TRONG này — vừa làm containing-block cho
          ảnh nền tuyệt đối, vừa giữ nguyên hành vi cao 100% gốc của
          <aside> (flex column height 100% do chính <aside> quy định). */}
      <div style={{ position: 'relative', flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column' }}>
        {/* FIX (sidebar-bg-camera): lớp ảnh nền camera exploded-view, phủ
            toàn bộ Sidebar, nằm dưới cùng (zIndex 0). */}
        <div
          aria-hidden="true"
          style={{
            position: 'absolute', inset: 0, zIndex: 0,
            backgroundImage: `url(${BG_IMAGE_URL})`,
            backgroundSize: 'cover',
            backgroundPosition: 'center 20%',
            backgroundRepeat: 'no-repeat',
          }}
        />
        {/* FIX (sidebar-bg-camera v2): GIẢM MẠNH độ mờ overlay (từ ~0.90-0.97
            xuống ~0.42-0.6) — bản trước phủ gần kín 90%+ nên ảnh gần như
            không thấy được như user phản hồi. Vẫn đủ tương phản cho chữ nhờ
            kết hợp thêm lớp glass mờ (backdrop-blur) ở từng item bên dưới. */}
        <div className="sidebar-bg-overlay" aria-hidden="true" style={{ position: 'absolute', inset: 0, zIndex: 1 }} />
        <style>{`
          [data-theme="dark"] .sidebar .sidebar-bg-overlay {
            background:
              linear-gradient(180deg, rgba(8,12,18,0.5) 0%, rgba(8,12,18,0.6) 55%, rgba(8,12,18,0.72) 100%);
          }
          [data-theme="light"] .sidebar .sidebar-bg-overlay,
          :root:not([data-theme]) .sidebar .sidebar-bg-overlay {
            background:
              linear-gradient(180deg, rgba(255,255,255,0.55) 0%, rgba(255,255,255,0.62) 55%, rgba(255,255,255,0.72) 100%);
          }
          /* FIX (theme-switch-clarity): header trước để background:'transparent'
             hoàn toàn — chữ/icon dễ bị chìm vào ảnh máy ảnh phía sau khi đổi
             theme (đặc biệt vùng ống kính tối màu). Giờ header có lớp nền mờ
             (blur) + màu riêng theo từng theme để LUÔN đủ tương phản, không
             phụ thuộc vùng ảnh đang hiện phía sau nó. */
          /* FIX (header-brand-colors): header trước đổi màu theo từng theme
             (đen mờ / trắng mờ) — giờ theo yêu cầu, header luôn dùng gradient
             thương hiệu #026466 → #cfdc00 CỐ ĐỊNH, không đổi theo theme, để
             đúng bộ nhận diện "IM-VINA PPC TEAM". */
          [data-theme="dark"] .sidebar .sidebar-header,
          [data-theme="light"] .sidebar .sidebar-header,
          :root:not([data-theme]) .sidebar .sidebar-header {
            background: linear-gradient(90deg, #026466 0%, #026466 62%, #cfdc00 100%) !important;
            backdrop-filter: none;
            -webkit-backdrop-filter: none;
            border-bottom: 1px solid rgba(0,0,0,0.15);
            position: relative;
          }
          /* FIX (theme-switch-clarity): tất cả khung menu/profile/nút giờ
             cùng 1 màu lightseagreen đặc — thêm viền/bóng trắng mờ riêng cho
             mục đang active để vẫn phân biệt được đang chọn mục nào, không
             bị "hòa lẫn" với các mục còn lại. */
          [data-theme="dark"] .sidebar .sidebar-glass-card,
          [data-theme="dark"] .sidebar .sidebar-glass-btn,
          [data-theme="light"] .sidebar .sidebar-glass-card,
          [data-theme="light"] .sidebar .sidebar-glass-btn,
          :root:not([data-theme]) .sidebar .sidebar-glass-card,
          :root:not([data-theme]) .sidebar .sidebar-glass-btn {
            border: 1px solid rgb(32,178,170) !important;
            background: rgb(32,178,170) !important;
            backdrop-filter: none;
            -webkit-backdrop-filter: none;
          }
          .sidebar .sidebar-item.active.sidebar-glass-card {
            border: 1px solid #ffffff !important;
            box-shadow: 0 0 0 2px rgba(32,178,170,0.55), 0 2px 6px rgba(0,0,0,0.25) !important;
          }
          [data-theme="dark"] .sidebar .sidebar-item-label,
          [data-theme="dark"] .sidebar .sidebar-glass-btn,
          [data-theme="light"] .sidebar .sidebar-item-label,
          [data-theme="light"] .sidebar .sidebar-glass-btn,
          :root:not([data-theme]) .sidebar .sidebar-item-label,
          :root:not([data-theme]) .sidebar .sidebar-glass-btn {
            color: #ffffff !important;
          }
          .sidebar .sidebar-item-index {
            background: rgb(220,216,0) !important;
            border: 1px solid rgb(220,216,0);
            color: #1a1a1a !important;
          }
          /* FIX (EPCC-sidebar-item-teal-shades): mỗi Mục 1-4 dùng 1 sắc độ
             teal riêng theo bảng tham chiếu, thay vì cùng 1 màu teal như
             trước. Chỉ đổi background — border/box-shadow/chữ giữ nguyên
             theo các rule .sidebar-glass-card ở trên (không khai báo lại). */
          .sidebar .sidebar-item[data-idx="1"] { background: #008489 !important; }
          .sidebar .sidebar-item[data-idx="2"] { background: #009298 !important; }
          .sidebar .sidebar-item[data-idx="3"] { background: #00A6AD !important; }
          .sidebar .sidebar-item[data-idx="4"] { background: #00B2BF !important; }
          /* FIX (toggle-btn-3d): nút mũi tên giờ là hình tròn nổi khối (gradient
             xanh dương + inset highlight/shadow để tạo cảm giác 3D), đặt
             position:absolute lệch sang phải (-16px) để "tràn" ra khỏi mép
             phải của thanh header/Sidebar cho đẹp, giống nút nổi (floating
             action button). overflow:visible trên header để không bị cắt. */
          .sidebar .sidebar-header {
            position: relative;
            overflow: visible !important;
          }
          .sidebar .sidebar-toggle-btn-3d {
            position: absolute;
            top: 50%;
            right: -16px;
            transform: translateY(-50%);
            width: 34px;
            height: 34px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            background: linear-gradient(145deg, #6d94ff 0%, #3f66e0 55%, #2c4fc4 100%);
            box-shadow: 0 3px 8px rgba(20,30,80,0.45), inset 0 1px 1px rgba(255,255,255,0.55), inset 0 -2px 4px rgba(10,15,50,0.4);
            border: 1px solid rgba(255,255,255,0.35);
            cursor: pointer;
            color: #ffffff;
            z-index: 5;
            padding: 0;
            transition: filter 0.15s ease, transform 0.15s ease;
          }
          .sidebar .sidebar-toggle-btn-3d:hover {
            filter: brightness(1.1);
          }
          .sidebar .sidebar-toggle-btn-3d:active {
            transform: translateY(-50%) scale(0.92);
          }
          .sidebar .sidebar-toggle-btn-3d svg {
            transition: transform 0.25s ease;
          }
          .sidebar.collapsed .sidebar-toggle-btn-3d svg {
            transform: rotate(180deg);
          }
        `}</style>

        {/* FIX (sidebar-bg-camera): nội dung thật của Sidebar, zIndex cao
            hơn overlay để luôn hiển thị phía trên ảnh nền + gradient. Dùng
            flex:1 + minHeight:0 (thay vì height:100% cứng) để item menu
            (flex:1 bên trong) vẫn co giãn đúng và đẩy khối Quản trị/Đăng
            xuất dính sát đáy Sidebar. */}
        <div style={{ position: 'relative', zIndex: 2, display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      <div className="sidebar-header">
        <div
          className="sidebar-logo"
          style={{
            display: 'flex', flexDirection: 'column', lineHeight: 1.15,
            gap: '1px', minWidth: 0, maxWidth: collapsed ? '48px' : 'calc(100% - 36px)',
            overflow: 'hidden',
          }}
        >
          {/* FIX (header-text-responsive): trước để fontSize cố định 15px —
              vừa nhỏ khi Sidebar mở rộng (user muốn to hơn nữa), vừa bị TRÀN/
              CẮT CHỮ khi Sidebar thu gọn (khung hẹp lại nhưng cỡ chữ không
              đổi). Giờ cỡ chữ đọc trực tiếp theo prop `collapsed`: to hẳn lên
              (20px) khi mở rộng, và thu nhỏ hẳn (7px) + bỏ letter-spacing khi
              thu gọn để chữ "IM VINA"/"PPC TEAM" vẫn nằm gọn, nhìn được hết,
              không bị cắt "...". */}
          <span style={{
            fontWeight: 800,
            fontSize: collapsed ? '7px' : '20px',
            letterSpacing: collapsed ? '0' : '0.4px',
            color: '#ffffff',
            whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'clip',
          }}>IM VINA</span>
          <span style={{
            fontWeight: 800,
            fontSize: collapsed ? '7px' : '20px',
            letterSpacing: collapsed ? '0' : '0.4px',
            color: '#cfdc00',
            whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'clip',
          }}>PPC TEAM</span>
        </div>
        {/* FIX (toggle-btn-3d): mũi tên chevron mảnh trong nút tròn nổi khối,
            tự xoay 180° qua CSS (`.sidebar.collapsed .sidebar-toggle-btn-3d
            svg`) khi Sidebar thu gọn, không cần 2 icon/emoji riêng nữa. */}
        <button
          className="sidebar-toggle-btn sidebar-toggle-btn-3d"
          onClick={onToggleCollapse}
          title={collapsed ? (lang === 'vi' ? 'Mở rộng' : 'Expand') : (lang === 'vi' ? 'Thu gọn' : 'Collapse')}
        >
          <ChevronIcon />
        </button>
      </div>

      <ul className="sidebar-menu" style={{ flex: 1, overflowY: 'auto' }}>
        {ITEMS.map((item) => {
          const isActive = activeViewId === item.id || (item.id === 'manpower' && activeViewId === 'production_per_capita');
          const label = item.name[lang];
          return (
            <li
              key={item.id}
              data-idx={item.index}
              className={`sidebar-item sidebar-glass-card ${isActive ? 'active' : ''}`}
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
        {/* FIX (sidebar-utility-widgets v3): đồng hồ realtime đã CHUYỂN RA
            lại từng Dashboard con (SalesDashboard/TargetActualDashboard/
            ManpowerDashboard — hiển thị trong ô "dashboard-header-left" cạnh
            tiêu đề, đúng vị trí ô trống được khoanh đỏ trong ảnh yêu cầu).
            Sidebar chỉ còn giữ khung Ngôn ngữ/Theme — vẫn đặt ở chân Sidebar,
            ngay trên khối tài khoản, dùng "sidebar-glass-card" để giữ màu
            khung đồng bộ. */}
        <div
          className="sidebar-glass-card"
          style={{
            display: 'flex', alignItems: 'center',
            justifyContent: collapsed ? 'center' : 'space-between',
            gap: '6px', padding: collapsed ? '8px' : '6px 8px',
            borderRadius: '10px',
          }}
        >
          {!collapsed && (
            <select
              value={lang}
              onChange={e => setLang(e.target.value as any)}
              className="lang-select"
              style={{ height: '30px', flex: 1, minWidth: 0, fontSize: '12px' }}
            >
              <option value="vi">Tiếng Việt</option>
              <option value="en">English</option>
              <option value="ko">한국어</option>
            </select>
          )}
          <button
            onClick={onToggleTheme}
            style={{
              height: '30px', width: collapsed ? '30px' : '32px', flexShrink: 0,
              border: 'none', background: 'rgba(255,255,255,0.16)', borderRadius: '7px',
              cursor: 'pointer', display: 'flex', alignItems: 'center',
              justifyContent: 'center', fontSize: '15px',
            }}
            title={theme === 'dark'
              ? (lang === 'vi' ? 'Chuyển sang Sáng' : lang === 'ko' ? '라이트 모드' : 'Switch to Light')
              : (lang === 'vi' ? 'Chuyển sang Tối' : lang === 'ko' ? '다크 모드' : 'Switch to Dark')}
          >
            {theme === 'dark' ? '☀️' : '🌙'}
          </button>
        </div>

        <div style={{
          display: 'flex', alignItems: 'center', gap: '8px',
          padding: collapsed ? '5px' : '7px 8px',
          borderRadius: '10px',
          border: '1px solid rgba(120,120,130,0.35)',
          background: 'rgba(120,120,130,0.16)',
          boxShadow: '0 1px 2px rgba(0,0,0,0.06)',
        }} className="sidebar-glass-card">
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
            className="sidebar-glass-btn"
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
          className="sidebar-glass-btn"
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
      </div>
      </div>
    </aside>
  );
};
