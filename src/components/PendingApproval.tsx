import type { Profile } from '../lib/auth';

interface PendingApprovalProps {
  profile: Profile | null;
  onSignOut: () => void;
  lang: 'vi' | 'en' | 'ko';
}

// Hiển thị khi user đã đăng nhập thành công (có session) nhưng profile.role
// vẫn là null — nghĩa là admin CHƯA duyệt/phân quyền cho tài khoản này.
export function PendingApproval({ profile, onSignOut, lang }: PendingApprovalProps) {
  const t = {
    vi: { title: 'Đang chờ phân quyền', body: 'Tài khoản của bạn đã đăng ký thành công. Vui lòng chờ quản trị viên phân quyền để truy cập hệ thống.', signOut: 'Đăng xuất' },
    en: { title: 'Pending approval', body: 'Your account has been registered. Please wait for an admin to assign your role.', signOut: 'Sign out' },
    ko: { title: '승인 대기 중', body: '계정이 등록되었습니다. 관리자의 권한 부여를 기다려주세요.', signOut: '로그아웃' },
  }[lang];

  return (
    <div style={{
      minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px',
    }}>
      <div style={{
        background: 'var(--surface, #fff)', borderRadius: '16px', padding: '36px',
        maxWidth: '420px', width: '100%', textAlign: 'center',
        border: '1px solid var(--border-soft, #e5e7eb)', boxShadow: 'var(--shadow-sm, 0 10px 30px rgba(0,0,0,0.1))',
      }}>
        <div style={{ fontSize: '40px', marginBottom: '12px' }}>⏳</div>
        <h2 style={{ fontSize: '19px', fontWeight: 700, color: 'var(--text-0, #111827)', marginBottom: '8px' }}>{t.title}</h2>
        <p style={{ fontSize: '14px', color: 'var(--text-2, #6b7280)', marginBottom: '6px' }}>
          {profile?.full_name || profile?.email}
        </p>
        <p style={{ fontSize: '13px', color: 'var(--text-2, #6b7280)', marginBottom: '20px' }}>{t.body}</p>
        <button
          onClick={onSignOut}
          style={{
            padding: '10px 20px', borderRadius: '8px', border: '1px solid var(--border-soft, #e5e7eb)',
            background: 'transparent', color: 'var(--text-0, #111827)', cursor: 'pointer', fontWeight: 600,
          }}
        >{t.signOut}</button>
      </div>
    </div>
  );
}
