import { useEffect, useState } from 'react';
import { supabase } from '../lib/supabase';
import type { Profile, UserRole } from '../lib/auth';

interface AdminUsersPanelProps {
  onClose: () => void;
  lang: 'vi' | 'en' | 'ko';
}

// Overlay modal cho admin duyệt/phân quyền người dùng mới đăng ký.
// Gọi qua RPC `admin_assign_role` (xem supabase/auth_schema.sql) — hàm này tự
// kiểm tra quyền admin ngay trong Postgres (security definer), không cho phép
// client UPDATE trực tiếp cột role qua REST API.
export function AdminUsersPanel({ onClose, lang }: AdminUsersPanelProps) {
  const t = {
    vi: {
      title: 'Bảng quản trị', sub: 'Phân quyền cho người dùng đăng ký',
      pending: 'Đang chờ phân quyền', approved: 'Đã phân quyền',
      noPending: 'Không có yêu cầu nào', noApproved: 'Chưa có ai được phân quyền',
      approve: 'Duyệt', close: 'Đóng',
      roleUser: 'Người dùng', roleEditor: 'Biên tập viên', roleAdmin: 'Quản trị viên',
    },
    en: {
      title: 'Admin panel', sub: 'Assign roles to registered users',
      pending: 'Pending approval', approved: 'Approved',
      noPending: 'No pending requests', noApproved: 'No users approved yet',
      approve: 'Approve', close: 'Close',
      roleUser: 'User', roleEditor: 'Editor', roleAdmin: 'Admin',
    },
    ko: {
      title: '관리자 패널', sub: '가입한 사용자에게 권한을 부여하세요',
      pending: '승인 대기 중', approved: '승인됨',
      noPending: '대기 중인 요청이 없습니다', noApproved: '아직 승인된 사용자가 없습니다',
      approve: '승인', close: '닫기',
      roleUser: '사용자', roleEditor: '편집자', roleAdmin: '관리자',
    },
  }[lang];

  const [profiles, setProfiles] = useState<Profile[]>([]);
  const [roleChoice, setRoleChoice] = useState<Record<string, UserRole>>({});
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState('');

  async function load() {
    if (!supabase) {
      setErr('Supabase is not initialized');
      setLoading(false);
      return;
    }
    setLoading(true);
    const { data, error } = await supabase.from('profiles').select('*').order('email', { ascending: true });
    if (error) { setErr(error.message); setLoading(false); return; }
    setProfiles((data as Profile[]) || []);
    setLoading(false);
  }

  useEffect(() => { load(); }, []);

  async function approve(id: string) {
    if (!supabase) {
      alert('Supabase is not initialized');
      return;
    }
    const role = roleChoice[id] || 'user';
    const { error } = await supabase.rpc('admin_assign_role', { target_id: id, new_role: role });
    if (error) { alert(error.message); return; }
    load();
  }

  function roleLabel(r: UserRole) {
    return r === 'admin' ? t.roleAdmin : r === 'editor' ? t.roleEditor : t.roleUser;
  }

  const pending = profiles.filter(p => !p.role);
  const approved = profiles.filter(p => p.role);

  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 100000,
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px',
    }}>
      <div style={{
        background: 'var(--surface, #fff)', borderRadius: '16px', padding: '28px',
        maxWidth: '640px', width: '100%', maxHeight: '85vh', overflowY: 'auto',
        border: '1px solid var(--border-soft, #e5e7eb)',
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '18px' }}>
          <div>
            <h2 style={{ fontSize: '19px', fontWeight: 700, color: 'var(--text-0, #111827)' }}>{t.title}</h2>
            <p style={{ fontSize: '13px', color: 'var(--text-2, #6b7280)' }}>{t.sub}</p>
          </div>
          <button onClick={onClose} style={{ border: 'none', background: 'transparent', fontSize: '20px', cursor: 'pointer', color: 'var(--text-2, #6b7280)' }}>✕</button>
        </div>

        {err && <div style={{ color: '#b91c1c', fontSize: '13px', marginBottom: '12px' }}>{err}</div>}
        {loading ? (
          <div style={{ color: 'var(--text-2, #6b7280)', fontSize: '13px' }}>…</div>
        ) : (
          <>
            <h3 style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-2, #6b7280)', marginBottom: '8px', textTransform: 'uppercase' }}>{t.pending}</h3>
            {pending.length === 0 ? (
              <div style={{ fontSize: '13px', color: 'var(--text-2, #6b7280)', marginBottom: '20px' }}>{t.noPending}</div>
            ) : (
              <div style={{ marginBottom: '20px' }}>
                {pending.map(p => (
                  <div key={p.id} style={{ display: 'flex', gap: '8px', alignItems: 'center', padding: '10px 0', borderBottom: '1px solid var(--border-soft, #eee)' }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: '14px', color: 'var(--text-0, #111827)' }}>{p.full_name}</div>
                      <div style={{ fontSize: '12px', color: 'var(--text-2, #6b7280)' }}>{p.email}</div>
                    </div>
                    <select
                      value={roleChoice[p.id] || 'user'}
                      onChange={e => setRoleChoice(r => ({ ...r, [p.id]: e.target.value as UserRole }))}
                      style={{ padding: '6px 8px', borderRadius: '6px', border: '1px solid var(--border-soft, #e5e7eb)' }}
                    >
                      <option value="user">{t.roleUser}</option>
                      <option value="editor">{t.roleEditor}</option>
                      <option value="admin">{t.roleAdmin}</option>
                    </select>
                    <button
                      onClick={() => approve(p.id)}
                      style={{ padding: '6px 14px', borderRadius: '6px', border: 'none', background: 'var(--primary, #6366f1)', color: '#fff', fontWeight: 600, cursor: 'pointer' }}
                    >{t.approve}</button>
                  </div>
                ))}
              </div>
            )}

            <h3 style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-2, #6b7280)', marginBottom: '8px', textTransform: 'uppercase' }}>{t.approved}</h3>
            {approved.length === 0 ? (
              <div style={{ fontSize: '13px', color: 'var(--text-2, #6b7280)' }}>{t.noApproved}</div>
            ) : (
              approved.map(p => (
                <div key={p.id} style={{ display: 'flex', gap: '8px', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid var(--border-soft, #eee)' }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: '14px', color: 'var(--text-0, #111827)' }}>{p.full_name}</div>
                    <div style={{ fontSize: '12px', color: 'var(--text-2, #6b7280)' }}>{p.email}</div>
                  </div>
                  <span style={{ fontSize: '12px', fontWeight: 700, padding: '4px 10px', borderRadius: '999px', background: 'var(--bg, #f3f4f6)', color: 'var(--text-0, #111827)' }}>
                    {roleLabel(p.role as UserRole)}
                  </span>
                </div>
              ))
            )}
          </>
        )}
      </div>
    </div>
  );
}
