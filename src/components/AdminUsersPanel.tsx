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
      resetPassword: 'Đặt lại mật khẩu', newPasswordPlaceholder: 'Nhập mật khẩu mới',
      save: 'Lưu', cancel: 'Hủy',
      passwordTooShort: 'Mật khẩu phải có ít nhất 6 ký tự.',
      resetSuccess: 'Đã đặt lại mật khẩu.',
    },
    en: {
      title: 'Admin panel', sub: 'Assign roles to registered users',
      pending: 'Pending approval', approved: 'Approved',
      noPending: 'No pending requests', noApproved: 'No users approved yet',
      approve: 'Approve', close: 'Close',
      roleUser: 'User', roleEditor: 'Editor', roleAdmin: 'Admin',
      resetPassword: 'Reset password', newPasswordPlaceholder: 'Enter new password',
      save: 'Save', cancel: 'Cancel',
      passwordTooShort: 'Password must be at least 6 characters.',
      resetSuccess: 'Password has been reset.',
    },
    ko: {
      title: '관리자 패널', sub: '가입한 사용자에게 권한을 부여하세요',
      pending: '승인 대기 중', approved: '승인됨',
      noPending: '대기 중인 요청이 없습니다', noApproved: '아직 승인된 사용자가 없습니다',
      approve: '승인', close: '닫기',
      roleUser: '사용자', roleEditor: '편집자', roleAdmin: '관리자',
      resetPassword: '비밀번호 재설정', newPasswordPlaceholder: '새 비밀번호 입력',
      save: '저장', cancel: '취소',
      passwordTooShort: '비밀번호는 최소 6자 이상이어야 합니다.',
      resetSuccess: '비밀번호가 재설정되었습니다.',
    },
  }[lang];

  const [profiles, setProfiles] = useState<Profile[]>([]);
  const [roleChoice, setRoleChoice] = useState<Record<string, UserRole>>({});
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState('');

  // ── Đặt lại mật khẩu cho user (khi họ quên mật khẩu) ───────────────────
  // KHÔNG có cách nào "xem lại" mật khẩu cũ của user — Supabase Auth chỉ lưu
  // bcrypt hash (một chiều), không lưu plaintext ở bất kỳ đâu. Vì vậy admin
  // chỉ có thể ĐẶT mật khẩu MỚI, không thể xem mật khẩu hiện tại. Icon con
  // mắt dưới đây chỉ để xem lại mật khẩu MỚI vừa gõ trước khi lưu (tránh gõ
  // nhầm), không liên quan tới mật khẩu cũ của user.
  // Việc lưu đi qua RPC `admin_assign_role`-style: `admin_reset_password`
  // (security definer trong Postgres, tự kiểm tra quyền admin) — xem SQL
  // đính kèm (supabase/admin_reset_password.sql) để thêm vào schema.
  const [resetTarget, setResetTarget] = useState<string | null>(null);
  const [newPassword, setNewPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [resetBusy, setResetBusy] = useState(false);
  const [resetMsg, setResetMsg] = useState<{ id: string; ok: boolean; text: string } | null>(null);

  function openReset(id: string) {
    setResetTarget(id); setNewPassword(''); setShowPassword(false); setResetMsg(null);
  }
  function closeReset() {
    setResetTarget(null); setNewPassword(''); setShowPassword(false); setResetMsg(null);
  }
  async function resetPassword(id: string) {
    if (!supabase) { alert('Supabase is not initialized'); return; }
    if (newPassword.length < 6) { setResetMsg({ id, ok: false, text: t.passwordTooShort }); return; }
    setResetBusy(true);
    const { error } = await supabase.rpc('admin_reset_password', { target_id: id, new_password: newPassword });
    setResetBusy(false);
    if (error) { setResetMsg({ id, ok: false, text: error.message }); return; }
    setResetMsg({ id, ok: true, text: t.resetSuccess });
    setTimeout(() => closeReset(), 1200);
  }

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
                  <div key={p.id} style={{ padding: '10px 0', borderBottom: '1px solid var(--border-soft, #eee)' }}>
                    <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
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
                      <button
                        onClick={() => (resetTarget === p.id ? closeReset() : openReset(p.id))}
                        title={t.resetPassword}
                        style={{ padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--border-soft, #e5e7eb)', background: 'transparent', cursor: 'pointer', fontSize: '13px', flexShrink: 0 }}
                      >🔑</button>
                    </div>
                    {resetTarget === p.id && <ResetPasswordRow
                      id={p.id} t={t} newPassword={newPassword} setNewPassword={setNewPassword}
                      showPassword={showPassword} setShowPassword={setShowPassword}
                      resetBusy={resetBusy} resetMsg={resetMsg}
                      onSave={() => resetPassword(p.id)} onCancel={closeReset}
                    />}
                  </div>
                ))}
              </div>
            )}

            <h3 style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-2, #6b7280)', marginBottom: '8px', textTransform: 'uppercase' }}>{t.approved}</h3>
            {approved.length === 0 ? (
              <div style={{ fontSize: '13px', color: 'var(--text-2, #6b7280)' }}>{t.noApproved}</div>
            ) : (
              approved.map(p => (
                <div key={p.id} style={{ padding: '8px 0', borderBottom: '1px solid var(--border-soft, #eee)' }}>
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: '14px', color: 'var(--text-0, #111827)' }}>{p.full_name}</div>
                      <div style={{ fontSize: '12px', color: 'var(--text-2, #6b7280)' }}>{p.email}</div>
                    </div>
                    <span style={{ fontSize: '12px', fontWeight: 700, padding: '4px 10px', borderRadius: '999px', background: 'var(--bg, #f3f4f6)', color: 'var(--text-0, #111827)' }}>
                      {roleLabel(p.role as UserRole)}
                    </span>
                    <button
                      onClick={() => (resetTarget === p.id ? closeReset() : openReset(p.id))}
                      title={t.resetPassword}
                      style={{ padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--border-soft, #e5e7eb)', background: 'transparent', cursor: 'pointer', fontSize: '13px', flexShrink: 0 }}
                    >🔑</button>
                  </div>
                  {resetTarget === p.id && <ResetPasswordRow
                    id={p.id} t={t} newPassword={newPassword} setNewPassword={setNewPassword}
                    showPassword={showPassword} setShowPassword={setShowPassword}
                    resetBusy={resetBusy} resetMsg={resetMsg}
                    onSave={() => resetPassword(p.id)} onCancel={closeReset}
                  />}
                </div>
              ))
            )}
          </>
        )}
      </div>
    </div>
  );
}

// Ô nhập mật khẩu mới + icon con mắt để xem lại trước khi lưu (chỉ xem lại
// giá trị vừa gõ, KHÔNG phải mật khẩu cũ — xem ghi chú ở đầu component).
function ResetPasswordRow({
  id, t, newPassword, setNewPassword, showPassword, setShowPassword, resetBusy, resetMsg, onSave, onCancel,
}: {
  id: string;
  t: { newPasswordPlaceholder: string; save: string; cancel: string };
  newPassword: string;
  setNewPassword: (v: string) => void;
  showPassword: boolean;
  setShowPassword: (v: boolean) => void;
  resetBusy: boolean;
  resetMsg: { id: string; ok: boolean; text: string } | null;
  onSave: () => void;
  onCancel: () => void;
}) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', marginTop: '8px', padding: '10px', borderRadius: '8px', background: 'var(--bg, #f9fafb)' }}>
      <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
        <div style={{ position: 'relative', flex: 1 }}>
          <input
            type={showPassword ? 'text' : 'password'}
            value={newPassword}
            onChange={e => setNewPassword(e.target.value)}
            placeholder={t.newPasswordPlaceholder}
            autoComplete="new-password"
            style={{ width: '100%', boxSizing: 'border-box', padding: '6px 34px 6px 8px', borderRadius: '6px', border: '1px solid var(--border-soft, #e5e7eb)', fontSize: '13px' }}
          />
          <button
            type="button"
            onClick={() => setShowPassword(!showPassword)}
            aria-label={showPassword ? 'Hide password' : 'Show password'}
            style={{ position: 'absolute', right: '4px', top: '50%', transform: 'translateY(-50%)', border: 'none', background: 'transparent', cursor: 'pointer', fontSize: '15px', padding: '2px 6px', color: 'var(--text-2, #6b7280)' }}
          >{showPassword ? '🙈' : '👁️'}</button>
        </div>
        <button
          onClick={onSave}
          disabled={resetBusy}
          style={{ padding: '6px 14px', borderRadius: '6px', border: 'none', background: 'var(--primary, #6366f1)', color: '#fff', fontWeight: 600, cursor: resetBusy ? 'default' : 'pointer', opacity: resetBusy ? 0.6 : 1, flexShrink: 0 }}
        >{t.save}</button>
        <button
          onClick={onCancel}
          style={{ padding: '6px 12px', borderRadius: '6px', border: '1px solid var(--border-soft, #e5e7eb)', background: 'transparent', cursor: 'pointer', flexShrink: 0 }}
        >{t.cancel}</button>
      </div>
      {resetMsg && resetMsg.id === id && (
        <div style={{ fontSize: '12px', color: resetMsg.ok ? '#15803d' : '#b91c1c' }}>{resetMsg.text}</div>
      )}
    </div>
  );
}
