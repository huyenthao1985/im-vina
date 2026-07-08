interface AccessDeniedProps {
  lang: 'vi' | 'en' | 'ko';
}

// Hiển thị khi user đã đăng nhập và có role, nhưng role đó không nằm trong
// TAB_ACCESS cho tab đang chọn (xem canAccessTab trong lib/auth.ts).
export function AccessDenied({ lang }: AccessDeniedProps) {
  const t = {
    vi: {
      title: 'Không có quyền truy cập',
      body: 'Tài khoản của bạn không có quyền xem mục này. Vui lòng liên hệ quản trị viên nếu bạn cần được cấp thêm quyền.',
    },
    en: {
      title: 'Access denied',
      body: 'Your account does not have permission to view this section. Please contact an admin if you need additional access.',
    },
    ko: {
      title: '접근 권한 없음',
      body: '이 메뉴를 볼 수 있는 권한이 없습니다. 추가 권한이 필요하면 관리자에게 문의하세요.',
    },
  }[lang];

  return (
    <div style={{
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      minHeight: '60vh', padding: '20px',
    }}>
      <div style={{
        background: 'var(--surface, #fff)', borderRadius: '16px', padding: '36px',
        maxWidth: '420px', width: '100%', textAlign: 'center',
        border: '1px solid var(--border-soft, #e5e7eb)', boxShadow: 'var(--shadow-sm, 0 10px 30px rgba(0,0,0,0.1))',
      }}>
        <div style={{ fontSize: '40px', marginBottom: '12px' }}>🔒</div>
        <h2 style={{ fontSize: '19px', fontWeight: 700, color: 'var(--text-0, #111827)', marginBottom: '8px' }}>{t.title}</h2>
        <p style={{ fontSize: '13px', color: 'var(--text-2, #6b7280)' }}>{t.body}</p>
      </div>
    </div>
  );
}
