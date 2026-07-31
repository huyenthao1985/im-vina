import { useEffect, useState, useCallback } from 'react';
import type { Session } from '@supabase/supabase-js';
import { supabase } from './supabase';

export type UserRole = 'user' | 'editor' | 'admin';

export interface Profile {
  id: string;
  full_name: string;
  email: string;
  role: UserRole | null;
}

// FIX (tab-role-gate): bảng phân quyền theo tab — chỉnh trực tiếp mảng role
// được phép cho từng activeViewId tại đây nếu muốn đổi ai xem được tab nào.
// Hiện tại: mọi role đã đăng nhập (user/editor/admin) đều xem được tất cả
// các tab, kể cả manpower/target_actual. 'placeholder' vẫn giữ admin-only vì
// đây là khu vực đang phát triển dở, chưa sẵn sàng cho người dùng thường.
export const TAB_ACCESS: Record<string, UserRole[]> = {
  overview: ['user', 'editor', 'admin'],
  target_actual: ['user', 'editor', 'admin'],
  manpower: ['user', 'editor', 'admin'],
  placeholder: ['admin'],
};

// Tab chưa khai báo trong TAB_ACCESS mặc định cho phép role nào xem.
const DEFAULT_TAB_ROLES: UserRole[] = ['user', 'editor', 'admin'];

export function canAccessTab(viewId: string, role: UserRole | null | undefined): boolean {
  if (!role) return false;
  const allowed = TAB_ACCESS[viewId] ?? DEFAULT_TAB_ROLES;
  return allowed.includes(role);
}

// useAuthGate — hook DUY NHẤT quản lý toàn bộ trạng thái đăng nhập cho App.tsx:
//   - session: null khi chưa đăng nhập
//   - profile: null khi chưa có hồ sơ / role: null khi đang chờ admin phân quyền
// Dùng CHUNG client `supabase` đã có sẵn ở ./supabase (client đang lưu dữ liệu
// sales/manpower) — KHÔNG tạo thêm client thứ 2 để tránh xung đột phiên/cache.
// EPCC (login-hang-no-timeout) - FIX ROOT CAUSE "đăng nhập vào trang tải quá
// lâu, không hiện màn hình": trước đây getSession()/loadProfile() KHÔNG có
// timeout nào — nếu mạng chậm hoặc query 'profiles' bị treo (cold-start
// Supabase, RLS chậm...), authLoading kẹt `true` vô thời hạn, người dùng chỉ
// thấy màn hình trắng + spinner, không có cách nào tự thoát ngoài refresh
// tay. Thêm timeout 15s (giống pattern withTimeoutMs đã dùng cho phần tải
// dữ liệu dashboard trong App.tsx) — hết 15s mà chưa xong thì dừng spinner,
// chuyển sang trạng thái lỗi có nút "Thử lại" thay vì treo mãi.
const AUTH_TIMEOUT_MS = 15_000;
function withAuthTimeout<T>(promise: Promise<T>): Promise<T> {
  return Promise.race([
    promise,
    new Promise<T>((_resolve, reject) =>
      setTimeout(() => reject(new Error('auth-timeout')), AUTH_TIMEOUT_MS)
    ),
  ]);
}

export function useAuthGate() {
  const [loading, setLoading] = useState(true);
  const [session, setSession] = useState<Session | null>(null);
  const [profile, setProfile] = useState<Profile | null>(null);
  // EPCC (login-hang-no-timeout): true khi getSession/loadProfile hết 15s mà
  // chưa xong — App.tsx dùng cờ này để hiện màn hình lỗi + nút "Thử lại"
  // thay vì spinner treo vô thời hạn.
  const [authTimedOut, setAuthTimedOut] = useState(false);

  const loadProfile = useCallback(async (userId: string, email: string, fallbackName: string) => {
    if (!supabase) return;
    let { data: prof } = await supabase.from('profiles').select('*').eq('id', userId).single();
    if (!prof) {
      // Chưa có hồ sơ (VD lần đầu xác nhận email) -> tự tạo, role = null
      // (nghĩa là "đang chờ admin phân quyền").
      const { data: created } = await supabase
        .from('profiles')
        .insert({ id: userId, full_name: fallbackName, email, role: null })
        .select()
        .single();
      prof = created;
    }
    setProfile((prof as Profile) ?? null);
  }, []);

  const refreshProfile = useCallback(async () => {
    if (!supabase) return;
    const { data: { user } } = await supabase.auth.getUser();
    if (user) {
      await loadProfile(user.id, user.email || '', user.user_metadata?.full_name || user.email || '');
    }
  }, [loadProfile]);

  useEffect(() => {
    if (!supabase) {
      setLoading(false);
      return;
    }
    let mounted = true;

    withAuthTimeout(supabase.auth.getSession())
      .then(async ({ data: { session } }) => {
        if (!mounted) return;
        setSession(session);
        if (session?.user) {
          await withAuthTimeout(
            loadProfile(
              session.user.id,
              session.user.email || '',
              session.user.user_metadata?.full_name || session.user.email || ''
            )
          );
        }
        if (mounted) setLoading(false);
      })
      .catch((err: Error) => {
        // EPCC (login-hang-no-timeout): getSession/loadProfile quá 15s —
        // dừng spinner, báo lỗi thay vì treo trắng vô thời hạn.
        console.warn('useAuthGate: hết thời gian chờ đăng nhập —', err.message);
        if (mounted) {
          setAuthTimedOut(true);
          setLoading(false);
        }
      });

    const { data: sub } = supabase.auth.onAuthStateChange(async (_event, newSession) => {
      setSession(newSession);
      if (newSession?.user) {
        await loadProfile(
          newSession.user.id,
          newSession.user.email || '',
          newSession.user.user_metadata?.full_name || newSession.user.email || ''
        );
      } else {
        setProfile(null);
      }
    });

    return () => {
      mounted = false;
      sub.subscription.unsubscribe();
    };
  }, [loadProfile]);

  const signOut = useCallback(async () => {
    if (!supabase) return;
    await supabase.auth.signOut();
    setSession(null);
    setProfile(null);
  }, []);

  // EPCC (login-hang-no-timeout): App.tsx dùng authTimedOut để hiện màn hình
  // lỗi + nút "Thử lại" (reload trang) thay vì spinner treo mãi.
  return { loading, session, profile, signOut, refreshProfile, authTimedOut };
}
