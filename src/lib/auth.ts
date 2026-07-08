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
export function useAuthGate() {
  const [loading, setLoading] = useState(true);
  const [session, setSession] = useState<Session | null>(null);
  const [profile, setProfile] = useState<Profile | null>(null);

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

    supabase.auth.getSession().then(async ({ data: { session } }) => {
      if (!mounted) return;
      setSession(session);
      if (session?.user) {
        await loadProfile(
          session.user.id,
          session.user.email || '',
          session.user.user_metadata?.full_name || session.user.email || ''
        );
      }
      if (mounted) setLoading(false);
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

  return { loading, session, profile, signOut, refreshProfile };
}
