import { useEffect, useState } from 'react';
import { supabase } from '../lib/supabase';
import loginBg from '../assets/login-bg.png';

interface LoginGateProps {
  lang: 'vi' | 'en' | 'ko';
  setLang: (l: 'vi' | 'en' | 'ko') => void;
  theme: 'light' | 'dark';
}

// LoginGate — tái tạo lại đúng mẫu "đèn bàn 3D kéo dây" từ bản demo tĩnh
// (index.html/app.js gốc) làm component React, dùng chung Supabase client
// hiện có (./lib/supabase — KHÔNG tạo client thứ 2).
//
// Ảnh nền: import trực tiếp từ src/assets/login-bg.png (ảnh linh kiện camera
// module IM Vina), gán qua style inline của .imv-bg-photo, phủ 1 lớp
// gradient tối để chữ + hiệu ứng đèn vẫn đọc rõ. Đổi ảnh khác chỉ cần thay
// file cùng tên trong src/assets/, không cần sửa code.
//
// Toàn bộ class được đặt tiền tố "imv-" để KHÔNG đụng các class chung của
// dashboard (.card, .field, .btn... đã dùng ở nơi khác trong app).
const TXT = {
  vi: {
    eyebrow: 'IM VINA',
    headline: 'Nỗ lực làm hết mình',
    hint: 'Kéo dây để bật / tắt đèn',
    cardTitle: 'Chào mừng', cardSub: 'PPC Team — IM VINA',
    labelEmail: 'Email', placeholderEmail: 'Nhập email',
    labelPass: 'Mật khẩu', placeholderPass: 'Nhập mật khẩu',
    btnSignIn: 'Đăng nhập',
    footNote: 'Chưa có tài khoản?', signup: 'Đăng ký',
    registerTitle: 'Tạo tài khoản mới', registerSub: 'Đăng ký để chờ admin phân quyền',
    labelFullname: 'Họ và tên', placeholderFullname: 'Nhập họ và tên',
    labelConfirmPass: 'Xác nhận mật khẩu', placeholderConfirmPass: 'Nhập lại mật khẩu',
    btnRegister: 'Đăng ký',
    haveAccount: 'Đã có tài khoản?', backToLogin: 'Đăng nhập',
    errMismatch: 'Mật khẩu xác nhận không khớp.',
    errGeneric: 'Đã có lỗi xảy ra, vui lòng thử lại.',
    checkEmail: 'Đăng ký thành công! Vui lòng kiểm tra email để xác nhận trước khi đăng nhập.',
    autoIn: 'Đăng ký thành công! Đang đưa bạn vào hệ thống…',
    processing: 'Đang xử lý…',
    rememberMe: 'Nhớ tôi',
    forgotPassword: 'Quên mật khẩu?',
    forgotNotice: 'Vui lòng liên hệ quản trị viên để được đặt lại mật khẩu.',
  },
  en: {
    eyebrow: 'IM VINA',
    headline: 'We will do our best',
    hint: 'Pull the cord to switch the lamp on / off',
    cardTitle: 'Welcome', cardSub: 'PPC Team — IM VINA',
    labelEmail: 'Email', placeholderEmail: 'Enter email',
    labelPass: 'Password', placeholderPass: 'Enter password',
    btnSignIn: 'Sign In',
    footNote: "Don't have an account?", signup: 'Sign up',
    registerTitle: 'Create new account', registerSub: 'Sign up and wait for admin approval',
    labelFullname: 'Full name', placeholderFullname: 'Enter your full name',
    labelConfirmPass: 'Confirm password', placeholderConfirmPass: 'Re-enter password',
    btnRegister: 'Sign Up',
    haveAccount: 'Already have an account?', backToLogin: 'Sign in',
    errMismatch: 'Passwords do not match.',
    errGeneric: 'Something went wrong, please try again.',
    checkEmail: 'Registration successful! Please check your email to confirm before signing in.',
    autoIn: 'Registration successful! Signing you in…',
    processing: 'Processing…',
    rememberMe: 'Remember me',
    forgotPassword: 'Forgot password?',
    forgotNotice: 'Please contact an admin to reset your password.',
  },
  ko: {
    eyebrow: 'IM VINA',
    headline: '최선을 다 하겠습니다',
    hint: '줄을 당겨 전등을 켜거나 끄세요',
    cardTitle: '환영합니다', cardSub: 'PPC Team — IM VINA',
    labelEmail: '이메일', placeholderEmail: '이메일 입력',
    labelPass: '비밀번호', placeholderPass: '비밀번호 입력',
    btnSignIn: '로그인',
    footNote: '계정이 없으신가요?', signup: '회원가입',
    registerTitle: '새 계정 만들기', registerSub: '가입 후 관리자의 승인을 기다려주세요',
    labelFullname: '이름', placeholderFullname: '이름을 입력하세요',
    labelConfirmPass: '비밀번호 확인', placeholderConfirmPass: '비밀번호를 다시 입력하세요',
    btnRegister: '가입하기',
    haveAccount: '이미 계정이 있으신가요?', backToLogin: '로그인',
    errMismatch: '비밀번호가 일치하지 않습니다.',
    errGeneric: '오류가 발생했습니다. 다시 시도해주세요.',
    checkEmail: '가입이 완료되었습니다! 로그인 전에 이메일을 확인해주세요.',
    autoIn: '가입이 완료되었습니다! 로그인 중입니다…',
    processing: '처리 중…',
    rememberMe: '아이디 저장',
    forgotPassword: '비밀번호를 잊으셨나요?',
    forgotNotice: '비밀번호 재설정은 관리자에게 문의해주세요.',
  },
} as const;

// Nạp Google Fonts đúng bộ font của bản gốc (chỉ 1 lần / toàn app).
function useImVinaFonts() {
  useEffect(() => {
    if (document.getElementById('imv-fonts')) return;
    const link = document.createElement('link');
    link.id = 'imv-fonts';
    link.rel = 'stylesheet';
    link.href = 'https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&family=Poppins:wght@700;800;900&family=Pixelify+Sans:wght@600;700&family=Fredoka:wght@600;700&display=swap';
    document.head.appendChild(link);
  }, []);
}

// Icon SVG kiểu Lucide, dùng currentColor để tự ăn theo màu chữ input.
function MailIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="2" y="4" width="20" height="16" rx="2" />
      <path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7" />
    </svg>
  );
}
function LockIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="11" width="18" height="11" rx="2" ry="2" />
      <path d="M7 11V7a5 5 0 0 1 10 0v4" />
    </svg>
  );
}
function EyeIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8Z" />
      <circle cx="12" cy="12" r="3" />
    </svg>
  );
}
function EyeOffIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M9.88 9.88a3 3 0 1 0 4.24 4.24" />
      <path d="M10.73 5.08A10.43 10.43 0 0 1 12 5c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-3.36 2.4A10.6 10.6 0 0 1 12 19c-7 0-11-8-11-8a18.5 18.5 0 0 1 4.22-5.94" />
      <line x1="1" y1="1" x2="23" y2="23" />
    </svg>
  );
}

export function LoginGate({ lang, setLang }: LoginGateProps) {
  useImVinaFonts();
  const t = TXT[lang];

  const [lampOn, setLampOn] = useState(false);
  const [pulled, setPulled] = useState(false);
  const [mode, setMode] = useState<'login' | 'register'>('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirm, setConfirm] = useState('');
  const [fullName, setFullName] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);
  const [rememberMe, setRememberMe] = useState(false);

  // FIX (remember-me): chỉ lưu EMAIL (không lưu mật khẩu vì lý do bảo mật)
  // vào localStorage nếu người dùng tick "Nhớ tôi", để lần sau mở lại tự
  // điền sẵn email, đỡ phải gõ lại.
  useEffect(() => {
    const saved = localStorage.getItem('imv_remember_email');
    if (saved) {
      setEmail(saved);
      setRememberMe(true);
    }
  }, []);

  function toggleLamp() {
    setPulled(true);
    setTimeout(() => setPulled(false), 250);
    setLampOn(v => !v);
  }

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    if (!supabase) { setError('Supabase is not initialized'); return; }
    setError(''); setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
    setLoading(false);
    if (error) { setError(error.message || t.errGeneric); return; }
    if (rememberMe) {
      localStorage.setItem('imv_remember_email', email.trim());
    } else {
      localStorage.removeItem('imv_remember_email');
    }
    // useAuthGate (App.tsx) tự cập nhật session/profile qua onAuthStateChange.
  }

  async function handleRegister(e: React.FormEvent) {
    e.preventDefault();
    if (!supabase) { setError('Supabase is not initialized'); return; }
    setError(''); setNotice('');
    if (password !== confirm) { setError(t.errMismatch); return; }
    setLoading(true);

    const { data, error } = await supabase.auth.signUp({
      email: email.trim(),
      password,
      options: { data: { full_name: fullName.trim() } },
    });

    if (error) { setLoading(false); setError(error.message || t.errGeneric); return; }

    if (data.user) {
      await supabase.from('profiles').insert({
        id: data.user.id, full_name: fullName.trim(), email: email.trim(), role: null,
      });
    }

    setLoading(false);
    if (data.session) {
      setNotice(t.autoIn);
    } else {
      setNotice(t.checkEmail);
      setTimeout(() => { setNotice(''); setMode('login'); }, 3500);
    }
  }

  return (
    <div className={`imv-body ${lampOn ? 'imv-is-on' : ''}`}>
      <style>{IMV_CSS}</style>

      <div
        className="imv-bg-photo"
        style={{
          backgroundImage: lampOn
            ? `linear-gradient(rgba(40,22,4,0.14), rgba(20,10,2,0.18)), url(${loginBg})`
            : `linear-gradient(rgba(8,9,14,0.24), rgba(8,9,14,0.3)), url(${loginBg})`,
          backgroundSize: 'cover',
          backgroundPosition: 'center 40%',
          backgroundRepeat: 'no-repeat',
        }}
      />
      <div className="imv-bg-overlay" />
      <div className="imv-spotlight" />

      <div className="imv-lang-switch">
        {(['vi', 'en', 'ko'] as const).map(l => (
          <button key={l} className={lang === l ? 'imv-active' : ''} onClick={() => setLang(l)}>
            {l.toUpperCase()}
          </button>
        ))}
      </div>

      <div className="imv-scene">
        <div className="imv-eyebrow">
          <span className="imv-eyebrow-brand">PPC Team</span>
          <span className="imv-eyebrow-sub">{t.eyebrow}</span>
        </div>
        <div className="imv-headline">{t.headline}</div>

        <div className="imv-lamp-wrap">
          <svg className="imv-lamp-svg" viewBox="0 0 200 300" xmlns="http://www.w3.org/2000/svg">
            <defs>
              <filter id="imvDropShadow" x="-60%" y="-60%" width="220%" height="220%">
                <feDropShadow dx="0" dy="10" stdDeviation="7" floodColor="#4c1d95" floodOpacity="0.35" />
              </filter>
              <filter id="imvSoftBlur" x="-80%" y="-80%" width="260%" height="260%">
                <feGaussianBlur stdDeviation="6" />
              </filter>
              <radialGradient id="imvGlowGrad" cx="50%" cy="50%" r="50%">
                <stop offset="0%" stopColor="#fff0d6" stopOpacity="0.95" />
                <stop offset="45%" stopColor="#ffb56b" stopOpacity="0.55" />
                <stop offset="100%" stopColor="#ff6fa8" stopOpacity="0" />
              </radialGradient>
              <linearGradient id="imvShadeGrad" x1="10%" y1="0%" x2="90%" y2="100%">
                <stop offset="0%" stopColor="#ffd166" />
                <stop offset="32%" stopColor="#ff8a5c" />
                <stop offset="64%" stopColor="#ff5da2" />
                <stop offset="100%" stopColor="#8b5cf6" />
              </linearGradient>
              <linearGradient id="imvShadeGloss" x1="0%" y1="0%" x2="0%" y2="100%">
                <stop offset="0%" stopColor="#ffffff" stopOpacity="0.65" />
                <stop offset="100%" stopColor="#ffffff" stopOpacity="0" />
              </linearGradient>
              <linearGradient id="imvPoleGrad" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stopColor="#ffd166" />
                <stop offset="50%" stopColor="#ff5da2" />
                <stop offset="100%" stopColor="#8b5cf6" />
              </linearGradient>
              <radialGradient id="imvBaseGrad" cx="42%" cy="25%" r="75%">
                <stop offset="0%" stopColor="#c4b5fd" />
                <stop offset="55%" stopColor="#a855f7" />
                <stop offset="100%" stopColor="#6d28d9" />
              </radialGradient>
            </defs>

            <ellipse cx="100" cy="270" rx="56" ry="11" fill="#4c1d95" opacity="0.3" filter="url(#imvSoftBlur)" />

            <g filter="url(#imvDropShadow)">
              <circle className="imv-bulb-glow" cx="100" cy="139" r="72" fill="url(#imvGlowGrad)" />
              <path d="M55,112 Q100,66 145,112 L163,150 Q100,171 37,150 Z" fill="url(#imvShadeGrad)" stroke="#7c3aed" strokeWidth="1" strokeOpacity="0.25" />
              <path d="M60,107 Q100,78 128,102 Q104,93 60,107 Z" fill="url(#imvShadeGloss)" />
              <circle className="imv-bulb" cx="100" cy="139" r="10" />
              <circle cx="97" cy="136" r="3" fill="#ffffff" opacity="0.85" />
              <rect x="96" y="150" width="8" height="98" rx="4" fill="url(#imvPoleGrad)" />
              <rect x="96.5" y="150" width="2.5" height="98" rx="1.2" fill="#ffffff" opacity="0.3" />
              <ellipse cx="100" cy="252" rx="48" ry="10" fill="url(#imvBaseGrad)" />
              <ellipse cx="100" cy="249" rx="48" ry="7" fill="#f3e8ff" opacity="0.45" />
            </g>

            <g className={`imv-cord-group ${pulled ? 'imv-pulled' : ''}`} onClick={toggleLamp} style={{ cursor: 'pointer' }}>
              <line className="imv-cord-line" x1="140" y1="150" x2="140" y2="182" stroke="#c084fc" strokeWidth="1.5" />
              <circle className="imv-knob" cx="140" cy="188" r="6" />
            </g>
          </svg>
          <div className="imv-hint">{t.hint}</div>
        </div>

        <div className="imv-card">
          <div className="imv-aperture" />
          <h2>{mode === 'login' ? t.cardTitle : t.registerTitle}</h2>
          <div className="imv-sub">{mode === 'login' ? t.cardSub : t.registerSub}</div>

          <form onSubmit={mode === 'login' ? handleLogin : handleRegister}>
            {mode === 'register' && (
              <div className="imv-field">
                <label>{t.labelFullname}</label>
                <input required placeholder={t.placeholderFullname} value={fullName} onChange={e => setFullName(e.target.value)} />
              </div>
            )}
            <div className="imv-field">
              <label>{t.labelEmail}</label>
              <div className="imv-input-wrap">
                <span className="imv-input-icon"><MailIcon /></span>
                <input
                  className="imv-input-has-icon"
                  required type="email" autoComplete="email"
                  placeholder={t.placeholderEmail} value={email}
                  onChange={e => setEmail(e.target.value)}
                />
              </div>
            </div>
            <div className="imv-field">
              <label>{t.labelPass}</label>
              <div className="imv-input-wrap">
                <span className="imv-input-icon"><LockIcon /></span>
                <input
                  className="imv-input-has-icon imv-input-has-toggle"
                  required type={showPassword ? 'text' : 'password'}
                  minLength={mode === 'register' ? 6 : undefined}
                  autoComplete={mode === 'login' ? 'current-password' : 'new-password'}
                  placeholder={t.placeholderPass} value={password}
                  onChange={e => setPassword(e.target.value)}
                />
                <button
                  type="button" className="imv-input-toggle"
                  onClick={() => setShowPassword(v => !v)}
                  tabIndex={-1}
                  aria-label={showPassword ? 'Hide password' : 'Show password'}
                >
                  {showPassword ? <EyeOffIcon /> : <EyeIcon />}
                </button>
              </div>
            </div>
            {mode === 'register' && (
              <div className="imv-field">
                <label>{t.labelConfirmPass}</label>
                <div className="imv-input-wrap">
                  <span className="imv-input-icon"><LockIcon /></span>
                  <input
                    className="imv-input-has-icon imv-input-has-toggle"
                    required type={showConfirm ? 'text' : 'password'}
                    minLength={6}
                    placeholder={t.placeholderConfirmPass} value={confirm}
                    onChange={e => setConfirm(e.target.value)}
                  />
                  <button
                    type="button" className="imv-input-toggle"
                    onClick={() => setShowConfirm(v => !v)}
                    tabIndex={-1}
                    aria-label={showConfirm ? 'Hide password' : 'Show password'}
                  >
                    {showConfirm ? <EyeOffIcon /> : <EyeIcon />}
                  </button>
                </div>
              </div>
            )}

            {mode === 'login' && (
              <div className="imv-row-between">
                <label className="imv-remember">
                  <input
                    type="checkbox" checked={rememberMe}
                    onChange={e => setRememberMe(e.target.checked)}
                  />
                  <span>{t.rememberMe}</span>
                </label>
                <button
                  type="button" className="imv-forgot-link"
                  onClick={() => setNotice(t.forgotNotice)}
                >
                  {t.forgotPassword}
                </button>
              </div>
            )}

            <button className="imv-signin-btn" type="submit" disabled={loading}>
              {loading ? t.processing : (mode === 'login' ? t.btnSignIn : t.btnRegister)}
            </button>
          </form>

          {error && <div className="imv-field-error">{error}</div>}
          {notice && <div className="imv-notice">{notice}</div>}

          <div className="imv-foot-note">
            {mode === 'login' ? (
              <>{t.footNote} <a href="#" onClick={e => { e.preventDefault(); setError(''); setNotice(''); setMode('register'); }}>{t.signup}</a></>
            ) : (
              <>{t.haveAccount} <a href="#" onClick={e => { e.preventDefault(); setError(''); setNotice(''); setMode('login'); }}>{t.backToLogin}</a></>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// CSS chuyển thể trực tiếp từ index.html gốc — mọi class thêm tiền tố "imv-"
// để không xung đột với .card/.field/.btn... của dashboard chính. Bỏ ảnh nền
// base64 gốc, thay bằng gradient tối cùng tông (xem .imv-bg-photo).
const IMV_CSS = `
.imv-body{
  --imv-steel:#8b93a1; --imv-steel-dark:#4b5361;
  --imv-brass:#c9a24a; --imv-brass-light:#e9cf8e;
  --imv-glow-warm:#ffdd9c; --imv-ink:#1c1f26; --imv-ink-soft:#454b57;
  --imv-glass:rgba(255,255,255,0.5); --imv-glass-border:rgba(255,255,255,0.65);
  font-family:'Inter', sans-serif;
  min-height:100vh; width:100%;
  display:flex; align-items:center; justify-content:center;
  overflow-x:hidden; position:relative; background:#eceef1;
}
.imv-bg-photo{
  position:fixed; inset:0;
  background:
    linear-gradient(rgba(8,9,14,0.24), rgba(8,9,14,0.3)),
    radial-gradient(circle at 30% 20%, #3a3550 0%, #14141c 55%, #0b0b10 100%);
  background-size: cover;
  filter: saturate(1) brightness(1.25);
  transition: filter 1.1s cubic-bezier(.4,0,.2,1);
  z-index:0;
}
.imv-is-on .imv-bg-photo{ filter: saturate(1.4) brightness(1.42) contrast(1.05); }
.imv-bg-overlay{
  position:fixed; inset:0; background: rgba(10,12,18,0.16);
  transition: background 1.1s cubic-bezier(.4,0,.2,1); z-index:1;
}
.imv-is-on .imv-bg-overlay{ background: rgba(255,178,90,0.04); }
.imv-spotlight{
  position:fixed; inset:0; z-index:2; pointer-events:none;
  background: radial-gradient(circle at 50% 38%, rgba(255,221,156,0) 0%, rgba(255,221,156,0) 100%);
  transition: background 1.1s cubic-bezier(.4,0,.2,1);
}
.imv-is-on .imv-spotlight{
  background: radial-gradient(circle at 50% 34%, rgba(255,221,156,0.65) 0%, rgba(255,221,156,0.32) 22%, rgba(255,221,156,0.0) 62%);
}
.imv-scene{
  position:relative; z-index:3; display:flex; flex-direction:column;
  align-items:center; gap:0; padding: 40px 20px 60px; width:100%; max-width:920px;
}
.imv-eyebrow{
  display:flex; align-items:baseline; justify-content:center; gap:10px;
  margin-bottom:6px; transition: color 1.1s ease;
}
.imv-eyebrow-brand{
  font-family:'Fredoka', sans-serif; font-weight:700;
  font-size:24px; letter-spacing:0.02em; text-transform:uppercase;
  color:#e1ff51;
  text-shadow: 0 2px 6px rgba(0,0,0,0.4);
}
.imv-eyebrow-sub{
  font-family:'Fredoka', sans-serif; font-size:24px; font-weight:700;
  letter-spacing:0.02em; text-transform:uppercase;
  color:#e1ff51;
  text-shadow: 0 2px 6px rgba(0,0,0,0.4);
  transition: color 1.1s ease;
}
.imv-is-on .imv-eyebrow-brand, .imv-is-on .imv-eyebrow-sub{ color:#7a9c0f; text-shadow: 0 1px 4px rgba(0,0,0,0.15); }
.imv-headline{
  font-family:'Fredoka', sans-serif; font-weight:700; text-transform:uppercase;
  font-size: clamp(34px, 5vw, 52px); text-align:center; margin-bottom:36px;
  letter-spacing:0.01em;
  color:#e1ff51;
  text-shadow: 0 3px 10px rgba(0,0,0,0.4);
  transition: color 1.1s ease, text-shadow 1.1s ease;
}
.imv-is-on .imv-headline{
  color:#7a9c0f;
  text-shadow: 0 2px 8px rgba(0,0,0,0.15);
}
.imv-lamp-wrap{ display:flex; flex-direction:column; align-items:center; margin-bottom:-18px; z-index:5; }
.imv-lamp-svg{ width:150px; height:auto; display:block; }
.imv-bulb-glow{ transition: opacity 1s cubic-bezier(.4,0,.2,1); opacity:0; }
.imv-is-on .imv-bulb-glow{ opacity:1; }
.imv-bulb{ fill:#b9a0c9; transition: fill 0.6s ease; }
.imv-is-on .imv-bulb{ fill:#ffe08a; }
.imv-cord-group:hover .imv-knob{ fill:#ffa8cd; }
.imv-knob{ fill:#ff6fa8; transition: fill 0.3s ease, transform 0.25s ease; transform-origin:100px 158px; }
.imv-cord-group.imv-pulled .imv-knob{ transform: translateY(9px); }
.imv-cord-group.imv-pulled .imv-cord-line{ transform: scaleY(1.12); }
.imv-cord-line{ transform-origin:140px 150px; transition: transform 0.25s ease; }
.imv-hint{
  font-family:'JetBrains Mono', monospace; font-size:11px; letter-spacing:0.06em;
  color: rgba(255,255,255,0.65); margin-top:6px; margin-bottom:28px;
  transition: color 1.1s ease, opacity 0.4s ease; text-align:center;
}
.imv-is-on .imv-hint{ color: rgba(60,45,20,0.55); }
.imv-card{
  position:relative; width:100%; max-width:380px; padding:38px 34px 32px;
  border-radius:22px; background: var(--imv-glass); border:1px solid var(--imv-glass-border);
  backdrop-filter: blur(18px) saturate(1.1); -webkit-backdrop-filter: blur(18px) saturate(1.1);
  box-shadow: 0 20px 60px rgba(0,0,0,0.25), inset 0 1px 0 rgba(255,255,255,0.6);
  transition: box-shadow 1.1s cubic-bezier(.4,0,.2,1), background 1.1s ease;
}
.imv-is-on .imv-card{
  box-shadow: 0 10px 50px rgba(0,0,0,0.12), 0 0 90px rgba(255,221,156,0.55), inset 0 1px 0 rgba(255,255,255,0.8);
  background: rgba(255,255,255,0.68);
}
.imv-aperture{
  position:absolute; top:-16px; right:-16px; width:56px; height:56px; border-radius:50%;
  border:1.5px solid rgba(255,255,255,0.5); display:flex; align-items:center; justify-content:center;
  pointer-events:none;
}
.imv-aperture::before{ content:''; width:32px; height:32px; border-radius:50%; border:1.5px solid rgba(255,255,255,0.35); }
.imv-card h2{ font-family:'Space Grotesk', sans-serif; font-size:22px; font-weight:600; color: var(--imv-ink); margin:0 0 4px; text-align:center; }
.imv-card .imv-sub{ font-size:13px; color: var(--imv-ink-soft); margin-bottom:26px; text-align:center; }
.imv-field{ margin-bottom:18px; }
.imv-field label{
  display:block; font-family:'JetBrains Mono', monospace; font-size:10.5px; letter-spacing:0.08em;
  text-transform:uppercase; color: var(--imv-ink-soft); margin-bottom:7px;
}
.imv-field input{
  width:100%; padding:12px 14px; border-radius:10px; border:1px solid rgba(75,83,97,0.25);
  background: rgba(255,255,255,0.55); font-family:'Inter', sans-serif; font-size:14px;
  color: var(--imv-ink); outline:none; box-sizing:border-box;
  transition: border-color 0.2s ease, box-shadow 0.2s ease, background 0.2s ease;
}
.imv-field input::placeholder{ color: rgba(69,75,87,0.45); }
.imv-field input:focus{
  border-color: var(--imv-brass); background: rgba(255,255,255,0.85);
  box-shadow: 0 0 0 3px rgba(201,162,74,0.18);
}
.imv-input-wrap{ position:relative; display:flex; align-items:center; }
.imv-input-icon{
  position:absolute; left:14px; display:flex; align-items:center; justify-content:center;
  color: var(--imv-ink-soft); pointer-events:none;
}
.imv-input-has-icon{ padding-left:38px !important; }
.imv-input-has-toggle{ padding-right:40px !important; }
.imv-input-toggle{
  position:absolute; right:10px; display:flex; align-items:center; justify-content:center;
  width:28px; height:28px; border:none; background:transparent; border-radius:6px;
  color: var(--imv-ink-soft); cursor:pointer; padding:0;
  transition: color 0.2s ease, background 0.2s ease;
}
.imv-input-toggle:hover{ color: var(--imv-ink); background: rgba(75,83,97,0.08); }
.imv-row-between{
  display:flex; align-items:center; justify-content:space-between;
  margin:-6px 0 18px; font-family:'Inter', sans-serif; font-size:12.5px;
}
.imv-remember{
  display:flex; align-items:center; gap:6px; color: var(--imv-ink-soft); cursor:pointer; user-select:none;
}
.imv-remember input[type="checkbox"]{
  width:14px; height:14px; accent-color: var(--imv-brass); cursor:pointer; margin:0;
}
.imv-forgot-link{
  border:none; background:transparent; padding:0; cursor:pointer;
  font-family:'Inter', sans-serif; font-size:12.5px; font-weight:600;
  color: var(--imv-brass); text-decoration:underline; text-underline-offset:2px;
}
.imv-forgot-link:hover{ color: var(--imv-brass-light); }
.imv-signin-btn{
  width:100%; margin-top:8px; padding:13px 14px; border:none; border-radius:10px;
  font-family:'Space Grotesk', sans-serif; font-weight:600; font-size:14.5px; letter-spacing:0.02em;
  color:#2a2110; cursor:pointer;
  background: linear-gradient(135deg, var(--imv-brass-light), var(--imv-brass) 60%, #a9843a);
  box-shadow: 0 6px 20px rgba(180,140,60,0.35);
  transition: transform 0.15s ease, box-shadow 0.2s ease, filter 0.2s ease;
}
.imv-signin-btn:hover{ filter:brightness(1.06); transform: translateY(-1px); }
.imv-signin-btn:active{ transform: translateY(0); }
.imv-signin-btn:disabled{ opacity:0.65; cursor:default; }
.imv-foot-note{ margin-top:18px; text-align:center; font-size:12px; color: var(--imv-ink-soft); }
.imv-foot-note a{ color: var(--imv-ink); font-weight:600; text-decoration:none; border-bottom:1px solid rgba(28,31,38,0.3); cursor:pointer; }
.imv-field-error{
  margin-top:10px; padding:10px 12px; border-radius:10px; background: rgba(239,68,68,0.1);
  border:1px solid rgba(239,68,68,0.3); font-size:12.5px; color:#b91c1c;
}
.imv-notice{
  margin-top:14px; padding:12px 14px; border-radius:10px; background: rgba(168,85,247,0.12);
  border:1px solid rgba(168,85,247,0.3); font-size:13px; color:#5b21b6; line-height:1.5;
}
.imv-lang-switch{
  position:fixed; top:18px; right:18px; z-index:20; display:flex; gap:4px; padding:4px;
  border-radius:999px; background: rgba(255,255,255,0.28); border:1px solid rgba(255,255,255,0.4);
  backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px);
  transition: background 1.1s ease, border-color 1.1s ease;
}
.imv-is-on .imv-lang-switch{ background: rgba(255,255,255,0.55); border-color: rgba(255,255,255,0.7); }
.imv-lang-switch button{
  border:none; background:transparent; font-family:'JetBrains Mono', monospace; font-size:10.5px;
  letter-spacing:0.05em; padding:6px 10px; border-radius:999px; color: rgba(255,255,255,0.75);
  cursor:pointer; transition: background 0.25s ease, color 0.25s ease;
}
.imv-is-on .imv-lang-switch button{ color: rgba(60,45,20,0.6); }
.imv-lang-switch button.imv-active{
  background: linear-gradient(135deg, var(--imv-brass-light), var(--imv-brass) 70%); color:#2a2110;
}
@media (max-width:480px){
  .imv-card{ padding:30px 24px 26px; }
  .imv-lamp-svg{ width:120px; }
}
@media (prefers-reduced-motion: reduce){
  .imv-body *{ transition:none !important; }
}
`;
