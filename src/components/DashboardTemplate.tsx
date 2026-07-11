import React, { useEffect, useState } from 'react';

/* ═══════════════════════════════════════════════════════════════════════
 * DASHBOARD TEMPLATE — FORM MẪU CHUẨN (viết theo quy trình EPCC)
 * ═══════════════════════════════════════════════════════════════════════
 *
 * EXPLORE (khảo sát hiện trạng):
 *   Rà lại 2 dashboard đang chạy thật — TargetActualDashboard.tsx (Mục 2,
 *   "BÁO CÁO DOANH SỐ") và RtyDashboard.tsx (Mục 4, "HIỆU SUẤT RTY") — phát
 *   hiện 2 file này VỐN PHẢI giống hệt nhau về khung/màu sắc nhưng bị lệch
 *   nhiều chỗ (label filter đổi màu theo theme khác nhau, tab không-active
 *   nền khác nhau, tiêu đề header ép màu cứng khác nhau...) vì mỗi file tự
 *   viết lại toàn bộ style riêng, không dùng chung 1 nguồn màu → sửa 1 nơi
 *   không tự động khớp nơi khác, dễ lệch khi nhân bản trang mới.
 *
 * PLAN (kế hoạch):
 *   1. Gom TOÀN BỘ màu sắc/kích thước đang dùng thật (đã kiểm chứng đúng
 *      theo ảnh mẫu do người dùng cung cấp) vào DUY NHẤT 1 object hằng số
 *      `THEME` ở đầu file — mọi thành phần bên dưới chỉ được PHÉP đọc màu
 *      từ `THEME`, không được viết số hex rời rạc lần 2 ở bất kỳ đâu khác.
 *   2. Xây lại cấu trúc 1 trang dashboard đầy đủ từ trên xuống dưới, đúng
 *      thứ tự các khung như ảnh mẫu: Header → 2 Tab → Filter bar (kèm badge
 *      "Đang hiển thị") → 4 thẻ KPI → Lưới 6 khung biểu đồ (radar/line/bar)
 *      → Bảng chi tiết (tab 2) + phân trang.
 *   3. KHÔNG gọi API/Supabase/database ở đây — toàn bộ dữ liệu là hằng số
 *      demo tĩnh (DEMO_*), có ghi chú rõ "// ← THAY DỮ LIỆU THẬT Ở ĐÂY" tại
 *      đúng vị trí cần nối dữ liệu thật khi nhân bản sang trang sản phẩm.
 *   4. Mỗi khung biểu đồ dùng 1 khu vực placeholder (SVG tĩnh, không phụ
 *      thuộc thư viện chart ngoài) để việc copy file sang dự án khác không
 *      bị vỡ do thiếu dependency — khi áp dụng thật chỉ cần thay phần
 *      placeholder bằng Plotly/Recharts thật, khung + màu viền giữ nguyên.
 *
 * CODE (triển khai): xem toàn bộ phần bên dưới.
 *
 * CHECK (đối chiếu lại):
 *   - Header/Filter bar: nền #2F3A1D, chữ #C0EF6A — khớp ảnh mẫu (đồng hồ +
 *     tiêu đề + toàn bộ nhãn NGÀY BẮT ĐẦU/NGÀY KẾT THÚC/MODEL/XEM THEO).
 *   - Tab active: gradient xanh dương #1d4ed8→#3b82f6; tab không active:
 *     nền mờ rgba(30,41,59,0.2) — ĐÚNG theo TargetActualDashboard (không
 *     còn kiểu nền xám sáng #e2e5ea như bản RTY cũ).
 *   - Badge "Đang hiển thị" nằm TRONG khung filter (không tách rời ra
 *     ngoài) — đúng theo yêu cầu chỉnh sửa gần nhất.
 *   - 4 thẻ KPI: xanh dương #0ea5e9 / xanh lá #10b981 / tím #8b5cf6 / cam
 *     #f59e0b — đúng thứ tự trái→phải như ảnh mẫu.
 *   - 6 khung biểu đồ: mỗi khung có borderLeft 4px + nền header nhạt cùng
 *     tông accent — màu lấy nguyên từ RtyDashboard.tsx (đã kiểm chứng đúng
 *     ảnh: tím #8b5cf6, cyan #06b6d4, cam #f59e0b, cam-đỏ #f97316, xanh lá
 *     #10b981, xanh dương #3b82f6).
 *   - File này KHÔNG import bất kỳ module gọi mạng/database nào → an toàn
 *     để copy làm điểm khởi đầu cho trang mới.
 *   - RESPONSIVE (zoom / laptop ↔ desktop ↔ màn nhỏ): mọi khối bên TRÁI
 *     (badge "Đang hiển thị", đồng hồ header) dùng flex-shrink:0 → vị trí
 *     cố định; khối GIỮA/PHẢI (field lọc, lưới KPI, lưới biểu đồ) dùng
 *     flex-wrap/grid auto-fit + minmax() → tự DÃN lấp đầy phần còn lại của
 *     khung hình và tự xuống hàng khi hẹp, KHÔNG tràn ngang khi zoom hoặc
 *     đổi kích thước màn hình. Xem chi tiết FIX (responsive-fixed-left-
 *     expand-right) trong khối <style> bên dưới.
 *
 * 📋 CÁCH ĐỌC FILE NÀY (dành cho người/AI copy sang trang mới):
 *   Mỗi khung lớn (Header, Tab, Filter bar, KPI, 6 biểu đồ, Bảng chi tiết)
 *   đều có 1 khối comment "📋 SPEC — KHUNG n/6: <TÊN>" đặt NGAY PHÍA TRÊN
 *   đoạn JSX tương ứng, liệt kê ĐẦY ĐỦ và CHÍNH XÁC: kích thước khung
 *   (border-radius/padding/border/box-shadow), vị trí từng phần tử con
 *   (trái/giữa/phải, cố định hay co giãn), cỡ chữ (px), màu sắc (hex/biến
 *   THEME cụ thể), và nội dung hiển thị đúng theo hiện trạng đã render ra
 *   ảnh mẫu — KHÔNG cần suy đoán/ước lượng, chỉ cần đọc SPEC rồi đối chiếu
 *   với đoạn code JSX ngay bên dưới để copy chính xác.
 * ═══════════════════════════════════════════════════════════════════════ */

/* ─────────────────────────────────────────────────────────────────────
 * 0) NGUỒN MÀU DUY NHẤT — mọi nơi trong file (và mọi file nhân bản sau
 *    này) chỉ nên tham chiếu tới THEME.*, KHÔNG viết mã hex rời lần 2.
 *    Khi cần đổi màu chuẩn cho toàn bộ hệ thống dashboard, chỉ sửa ở đây.
 * ───────────────────────────────────────────────────────────────────── */
const THEME = {
  // 1. HEADER BAR — khung trên cùng (đồng hồ + tiêu đề trang)
  headerBg: '#2F3A1D',
  headerText: '#C0EF6A',

  // 2. TAB — 2 nút chuyển tab ngay dưới header
  tab: {
    inactiveBg: 'rgba(30, 41, 59, 0.2)',
    inactiveText: 'var(--text-2)',
    inactiveBorder: 'var(--border)',
    activeBg: 'linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%)',
    activeText: '#ffffff',
    activeBorder: '#3b82f6',
    activeShadow: '0 4px 12px rgba(59, 130, 246, 0.35)',
  },

  // 3. FILTER BAR — khung lọc ngày/model/xem theo (nền giống Header)
  filterBarBg: '#2F3A1D',
  filterLabelColor: '#C0EF6A',
  viewModeActiveBg: '#2e7d8c',
  viewModeInactiveBg: 'rgba(255,255,255,0.55)',
  viewModeInactiveText: '#7A5A2E',

  // 4. BADGE "Đang hiển thị: ⭐ MODEL" — đặt TRONG filter bar
  badge: {
    bestBg: 'rgba(16,185,129,0.22)',
    bestText: '#34d399',
    normalBg: 'rgba(14,165,233,0.22)',
    normalText: '#38bdf8',
  },

  // 5. 4 THẺ KPI — trái → phải
  kpi: [
    { accent: '#0ea5e9', gradientTint: '#0ea5e91a' }, // KPI 1 — xanh dương
    { accent: '#10b981', gradientTint: '#10b9811a' }, // KPI 2 — xanh lá
    { accent: '#8b5cf6', gradientTint: '#8b5cf61a' }, // KPI 3 — tím
    { accent: '#f59e0b', gradientTint: '#f59e0b1a' }, // KPI 4 — cam
  ] as const,

  // 6. 6 KHUNG BIỂU ĐỒ — accent (viền trái + đường biểu đồ chính) và nền
  //    header nhạt tương ứng, tách riêng cho light/dark theme.
  chartPanels: [
    { key: 'chart1', accent: '#8b5cf6', bgLight: '#d6c6fc', bgDark: 'rgba(139,92,246,0.14)' },
    { key: 'chart2', accent: '#06b6d4', bgLight: '#a8e5f0', bgDark: 'rgba(6,182,212,0.14)' },
    { key: 'chart3', accent: '#f59e0b', bgLight: '#fcddaa', bgDark: 'rgba(255,255,255,0.05)' },
    { key: 'chart4', accent: '#f97316', bgLight: '#fdcead', bgDark: 'rgba(249,115,22,0.14)' },
    { key: 'chart5', accent: '#10b981', bgLight: '#abe7d3', bgDark: 'rgba(16,185,129,0.14)' },
    { key: 'chart6', accent: '#3b82f6', bgLight: '#bad3fc', bgDark: 'rgba(59,130,246,0.14)' },
  ] as const,
} as const;

/* ─────────────────────────────────────────────────────────────────────
 * 1) DỮ LIỆU DEMO TĨNH — KHÔNG đọc từ database/API. Khi nhân bản trang
 *    mới, xoá khối DEMO_* này và nối vào props/hook dữ liệu thật của
 *    trang đó; toàn bộ phần render bên dưới không cần sửa cấu trúc/màu.
 * ───────────────────────────────────────────────────────────────────── */
interface KpiDemo { label: string; value: string; sub?: string; icon: 'box' | 'target' | 'check' | 'trend'; }

const DEMO_KPIS: KpiDemo[] = [ // ← THAY DỮ LIỆU THẬT Ở ĐÂY
  { label: 'CHỈ SỐ 1 (TB)', value: '96.14%', sub: 'Mục tiêu: 96.4%', icon: 'box' },
  { label: 'CHỈ SỐ 2 (TỈ LỆ ĐẠT)', value: '0.0%', icon: 'target' },
  { label: 'CHỈ SỐ 3 (ĐẠT MỤC TIÊU)', value: '0/1', icon: 'check' },
  { label: 'CHỈ SỐ 4 (CHÊNH LỆCH)', value: '-0.26pp', icon: 'trend' },
];

interface ChartPanelDemo { title: string; note: string; kind: 'radar' | 'combo'; }

const DEMO_PANELS: ChartPanelDemo[] = [ // ← THAY DỮ LIỆU/CHART THẬT Ở ĐÂY
  { title: 'BIỂU ĐỒ RADAR 1 (THEO MODEL)', note: 'Demo radar — tổng hợp theo model', kind: 'radar' },
  { title: 'BIỂU ĐỒ RADAR 2 (THEO MODEL)', note: 'Demo radar — tổng hợp theo model', kind: 'radar' },
  { title: 'BIỂU ĐỒ 3 (CHI TIẾT)', note: 'Demo bar + line kết hợp', kind: 'combo' },
  { title: 'BIỂU ĐỒ 4 (CHI TIẾT)', note: 'Demo bar + line kết hợp', kind: 'combo' },
  { title: 'BIỂU ĐỒ 5 (CHI TIẾT)', note: 'Demo bar + line kết hợp', kind: 'combo' },
  { title: 'BIỂU ĐỒ 6 (CHI TIẾT)', note: 'Demo bar + line kết hợp', kind: 'combo' },
];

interface RowDemo { col1: string; col2: string; col3: string; target: string; actual: string; gap: string; }

const DEMO_ROWS: RowDemo[] = [ // ← THAY DỮ LIỆU THẬT Ở ĐÂY (kết quả query/props)
  { col1: 'MODEL A', col2: 'Process 1', col3: '07/2026', target: '96.4%', actual: '97.7%', gap: '+1.3pp' },
  { col1: 'MODEL B', col2: 'Process 2', col3: '07/2026', target: '96.4%', actual: '94.8%', gap: '-1.6pp' },
  { col1: 'MODEL C', col2: 'Process 3', col3: '07/2026', target: '96.4%', actual: '98.0%', gap: '+1.6pp' },
];

const DEMO_MODEL_OPTIONS = ['MODEL A', 'MODEL B', 'MODEL C']; // ← THAY DỮ LIỆU THẬT Ở ĐÂY
const BEST_MODEL_DEMO = 'MODEL A';

/* ─────────────────────────────────────────────────────────────────────
 * 2) COMPONENT PHỤ — icon KPI + placeholder biểu đồ (SVG thuần, không
 *    phụ thuộc thư viện ngoài) để file này copy sang dự án khác luôn
 *    chạy được ngay, không cần cài thêm package.
 * ───────────────────────────────────────────────────────────────────── */
const KpiIcon: React.FC<{ kind: KpiDemo['icon']; color: string }> = ({ kind, color }) => {
  const common = { viewBox: '0 0 24 24', fill: 'none', stroke: color, strokeWidth: 2.5, strokeLinecap: 'round' as const, strokeLinejoin: 'round' as const, style: { width: 16, height: 16, flexShrink: 0 } };
  if (kind === 'box') return (
    <svg {...common}><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" /><polyline points="3.27 6.96 12 12.01 20.73 6.96" /><line x1="12" y1="22.08" x2="12" y2="12" /></svg>
  );
  if (kind === 'target') return (
    <svg {...common}><circle cx="12" cy="12" r="10" /><circle cx="12" cy="12" r="6" /><circle cx="12" cy="12" r="2" /></svg>
  );
  if (kind === 'check') return (
    <svg {...common}><path d="M9 11l3 3L22 4" /><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" /></svg>
  );
  return (
    <svg {...common}><polyline points="23 6 13.5 15.5 8.5 10.5 1 18" /><polyline points="17 6 23 6 23 12" /></svg>
  );
};

/** Placeholder radar demo — chỉ để giữ đúng khung/tỉ lệ, không phải chart thật. */
const RadarPlaceholder: React.FC<{ accent: string }> = ({ accent }) => (
  <svg viewBox="0 0 240 200" style={{ width: '100%', height: '100%' }}>
    {[80, 60, 40, 20].map(r => (
      <polygon key={r} points={pentagonPoints(120, 100, r)} fill="none" stroke="rgba(148,163,184,0.35)" strokeWidth={1} />
    ))}
    <polygon points={pentagonPoints(120, 100, 62)} fill={`${accent}33`} stroke={accent} strokeWidth={2} />
  </svg>
);
function pentagonPoints(cx: number, cy: number, r: number): string {
  const pts: string[] = [];
  for (let i = 0; i < 5; i++) {
    const angle = -Math.PI / 2 + (i * 2 * Math.PI) / 5;
    pts.push(`${cx + r * Math.cos(angle)},${cy + r * Math.sin(angle)}`);
  }
  return pts.join(' ');
}

/** Placeholder bar+line demo — chỉ để giữ đúng khung/tỉ lệ, không phải chart thật. */
const ComboChartPlaceholder: React.FC<{ accent: string }> = ({ accent }) => {
  const bars = [62, 74, 58, 80, 66, 90, 70];
  return (
    <svg viewBox="0 0 280 160" style={{ width: '100%', height: '100%' }} preserveAspectRatio="none">
      {bars.map((h, i) => (
        <rect key={i} x={12 + i * 38} y={150 - h} width={14} height={h} fill={`${accent}55`} rx={2} />
      ))}
      <polyline
        points={bars.map((h, i) => `${19 + i * 38},${150 - h - 14}`).join(' ')}
        fill="none" stroke={accent} strokeWidth={2.5}
      />
    </svg>
  );
};

/* ─────────────────────────────────────────────────────────────────────
 * 3) COMPONENT CHÍNH — DashboardTemplate
 *    Props tối giản (theme) — nhân bản trang mới chỉ cần import + đổi
 *    DEMO_* thành dữ liệu thật, giữ nguyên toàn bộ style/khung bên dưới.
 * ───────────────────────────────────────────────────────────────────── */
interface DashboardTemplateProps {
  /** 'light' | 'dark' — quyết định bgLight/bgDark của khung biểu đồ. */
  theme?: 'light' | 'dark';
  /** Tiêu đề hiển thị trên header — đổi theo từng trang khi nhân bản. */
  title?: string;
}

const DashboardTemplate: React.FC<DashboardTemplateProps> = ({ theme = 'dark', title = 'TIÊU ĐỀ TRANG (THAY KHI NHÂN BẢN)' }) => {
  const isLightMode = theme === 'light';

  // ── Đồng hồ realtime góc trái header — giữ đúng hành vi 2 dashboard gốc ──
  const [formattedTime, setFormattedTime] = useState('');
  useEffect(() => {
    const tick = () => {
      const now = new Date();
      const hh = String(now.getHours()).padStart(2, '0');
      const mm = String(now.getMinutes()).padStart(2, '0');
      const ss = String(now.getSeconds()).padStart(2, '0');
      const dd = String(now.getDate()).padStart(2, '0');
      const mo = String(now.getMonth() + 1).padStart(2, '0');
      const yy = now.getFullYear();
      setFormattedTime(`${hh}:${mm}:${ss} ${dd}/${mo}/${yy}`);
    };
    tick();
    const id = setInterval(tick, 1000);
    return () => clearInterval(id);
  }, []);

  // ── State filter demo (không gọi API) ──
  const [activeTab, setActiveTab] = useState<'summary' | 'detail'>('summary');
  const [startDate, setStartDate] = useState('2026-01-02');
  const [endDate, setEndDate] = useState('2026-07-09');
  const [selectedModel, setSelectedModel] = useState('');
  const [viewMode, setViewMode] = useState<'day' | 'week' | 'month'>('day');

  const showingModel = selectedModel || BEST_MODEL_DEMO;
  const isBest = showingModel === BEST_MODEL_DEMO;

  const panelBg = (p: typeof THEME.chartPanels[number]) => (isLightMode ? p.bgLight : p.bgDark);

  return (
    <div style={{ padding: '0 16px 16px 16px', height: '100%', overflowY: 'auto', boxSizing: 'border-box' }}>
      <style>{`
        .tpl-tab-btn { transition: all 0.2s ease; cursor: pointer; }
        .tpl-page-btn { padding: 4px 10px; border-radius: 6px; border: 1px solid var(--border); background: transparent; cursor: pointer; }
        .tpl-page-btn.active { background: #2563eb; color: #fff; border-color: #2563eb; }

        /* ═════════════════════════════════════════════════════════════
           FIX (responsive-fixed-left-expand-right, EPCC):
           EXPLORE — bản cũ dùng width PX CỐ ĐỊNH (170/130/160/180/120px)
           cho từng ô trong filter bar + grid-template-columns: repeat(4|2,1fr)
           cố định số cột cho KPI/biểu đồ. Khi thu nhỏ màn hình (laptop →
           màn hình nhỏ) hoặc zoom trình duyệt (Ctrl +/-), tổng bề rộng các
           ô cố định vượt quá bề rộng khung nhìn → tràn ngang (horizontal
           overflow) thay vì co giãn gọn theo màn hình.
           PLAN — 3 nguyên tắc cho MỌI khung responsive trong trang:
             1) Khối bên TRÁI (badge "Đang hiển thị", nhãn cột trái, viền
                trái accent...) luôn giữ flex-shrink:0 → vị trí/độ rộng
                CỐ ĐỊNH, không bị bóp méo khi màn hình co lại.
             2) Khối GIỮA/PHẢI (vùng input, lưới KPI, lưới biểu đồ) dùng
                flex:1/grid auto-fit với minmax(...) → tự DÃN sang phải để
                lấp đầy phần không gian còn lại của MỖI khung hình monitor,
                và tự XUỐNG HÀNG (wrap) khi không đủ chỗ thay vì tràn ngang.
             3) Không dùng flex-basis/width cố định cho toàn khối lớn — chỉ
                dùng minmax()/min-width để đặt "ngưỡng co nhỏ nhất", còn lại
                để trình duyệt tự tính toán theo kích thước khung hình thật.
           CODE — áp dụng qua các class bên dưới (dùng chung cho toàn trang,
           không lặp lại style riêng ở từng nơi để tránh lệch khi nhân bản).
           ═════════════════════════════════════════════════════════════ */
        .tpl-header-row {
          display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
        }
        .tpl-header-left { flex: 0 0 auto; } /* CỐ ĐỊNH bên trái */
        .tpl-header-title {
          flex: 1 1 auto; min-width: 160px; text-align: center;
          overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }
        .tpl-header-right { flex: 0 0 auto; width: 0; } /* spacer cân đối, tự co về 0 khi chật chỗ */

        .tpl-filterbar-row {
          display: flex; align-items: center; gap: 12px; width: 100%; flex-wrap: wrap;
        }
        .tpl-filter-left {
          flex: 0 0 auto; display: flex; align-items: center; gap: 6px; /* CỐ ĐỊNH bên trái */
        }
        .tpl-filter-fields {
          flex: 1 1 320px; min-width: 0; /* DÃN sang phải, tự wrap khi hẹp */
          display: flex; flex-wrap: wrap; gap: 12px 16px; align-items: flex-end;
        }
        .tpl-filter-field {
          display: flex; flex-direction: column; gap: 4px; min-width: 0;
        }
        .tpl-filter-right { flex: 0 0 auto; margin-left: auto; } /* luôn bám mép phải khung */

        /* Lưới KPI: auto-fit — số cột tự giảm khi khung hình nhỏ lại, luôn
           lấp đầy hết bề rộng còn lại thay vì để trống bên phải. */
        .tpl-kpi-grid {
          display: grid; gap: 12px; margin-bottom: 12px; width: 100%;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        }

        /* Lưới 6 khung biểu đồ: 2 cột trên màn rộng, tự rút còn 1 cột khi
           màn hẹp/zoom to — không cần media query nhờ auto-fit + minmax. */
        .tpl-chart-grid {
          display: grid; gap: 16px;
          grid-template-columns: repeat(auto-fit, minmax(380px, 1fr));
        }

        /* An toàn thêm cho màn rất nhỏ / zoom rất sâu — vẫn giữ khối trái
           cố định, chỉ nới lỏng khoảng cách + cỡ chữ tiêu đề cho gọn. */
        @media (max-width: 640px) {
          .tpl-header-title { font-size: 15px !important; }
          .tpl-filter-left, .tpl-filter-right { flex-basis: 100%; justify-content: flex-start; }
          .tpl-filter-right { justify-content: flex-end; }
        }
      `}</style>

      {/* ═══════════════════════════════════════════════════════════════
          📋 SPEC — KHUNG 1/6: HEADER BAR (thanh trên cùng)
          ───────────────────────────────────────────────────────────────
          KÍCH THƯỚC KHUNG (container ngoài)
            • border-radius : 14px
            • padding       : 10px 16px
            • border        : 1px solid rgba(0,0,0,0.18)
            • box-shadow    : 0 2px 8px rgba(0,0,0,0.08)
            • background    : THEME.headerBg  = '#2F3A1D'
            • margin-bottom : 10px (cách khung Tab bên dưới)
            • bố cục        : flex hàng ngang, align-items: center,
                              flex-wrap: wrap (để không tràn khi màn hẹp)

          BỐ CỤC BÊN TRONG — 3 vùng, trái → giữa → phải:

          1) [TRÁI] Đồng hồ realtime  — class .tpl-header-left
             - Vị trí   : flex: 0 0 auto (CỐ ĐỊNH, không co giãn/không dãn)
             - Hiển thị : icon "🕐" + chuỗi giờ "HH:MM:SS DD/MM/YYYY"
                          (VD: "14:05:30 11/07/2026", tự cập nhật mỗi giây)
             - Cỡ chữ   : 13px
             - Độ đậm   : font-weight 700
             - Màu chữ  : THEME.headerText = '#C0EF6A' (xanh vàng chanh)
             - Khác     : white-space: nowrap (không xuống dòng giữa icon/giờ)

          2) [GIỮA] Tiêu đề trang — class .tpl-header-title (thẻ <h1>)
             - Vị trí   : flex: 1 1 auto, text-align: center, min-width: 160px
             - Hiển thị : nội dung PROP `title` truyền vào component
                          (VD: "HIỆU SUẤT RTY", "BÁO CÁO DOANH SỐ THÁNG 7")
             - Cỡ chữ   : 18px (màn ≤640px tự giảm còn 15px, xem media query)
             - Độ đậm   : font-weight 800
             - Màu chữ  : THEME.headerText = '#C0EF6A' (GIỐNG HỆT đồng hồ)
             - Khác     : letter-spacing 0.02em; overflow ẩn + "..." nếu quá dài

          3) [PHẢI] Spacer rỗng — class .tpl-header-right
             - Vị trí   : flex: 0 0 auto, width: 0
             - Hiển thị : KHÔNG có nội dung — chỉ tồn tại để giữ tiêu đề
                          cân giữa về mặt bố cục khi khung đủ rộng
          ═══════════════════════════════════════════════════════════════ */}
      <div className="tpl-header-row" style={{
        background: THEME.headerBg, borderRadius: '14px', padding: '10px 16px',
        border: '1px solid rgba(0,0,0,0.18)', boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
        marginBottom: '10px',
      }}>
        <div className="tpl-header-left" style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: THEME.headerText, fontWeight: 700, whiteSpace: 'nowrap' }}>
          <span aria-hidden="true">🕐</span>
          {formattedTime}
        </div>
        <h1 className="tpl-header-title" style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: THEME.headerText, letterSpacing: '0.02em' }}>
          {title}
        </h1>
        <div className="tpl-header-right" />
      </div>

      {/* ═══════════════════════════════════════════════════════════════
          📋 SPEC — KHUNG 2/6: TAB BAR (2 nút chuyển trang con)
          ───────────────────────────────────────────────────────────────
          KÍCH THƯỚC KHUNG (container ngoài)
            • không có nền/viền riêng — chỉ là 1 hàng flex chứa 2 nút
            • display     : flex, gap: 10px (khoảng cách giữa 2 nút)
            • margin-bottom: 8px (cách khung Filter bar bên dưới)

          MỖI NÚT (button) — kích thước & màu GIỐNG HỆT nhau, chỉ khác
          trạng thái active/inactive:
            • padding        : 8px 18px
            • border-radius  : 8px
            • cỡ chữ         : 14px
            • độ đậm         : font-weight 700
            • Hiển thị       : icon emoji + nhãn chữ IN HOA
                               (nút 1: "📊 TÌNH HÌNH TỔNG QUAN"
                                nút 2: "💼 CHI TIẾT & DỮ LIỆU")

          TRẠNG THÁI ACTIVE (đang được chọn):
            • background : THEME.tab.activeBg
                           = linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%)
            • màu chữ    : THEME.tab.activeText = '#ffffff'
            • viền       : 1px solid THEME.tab.activeBorder = '#3b82f6'
            • box-shadow : THEME.tab.activeShadow
                           = 0 4px 12px rgba(59, 130, 246, 0.35)

          TRẠNG THÁI INACTIVE (chưa chọn):
            • background : THEME.tab.inactiveBg = 'rgba(30, 41, 59, 0.2)'
            • màu chữ    : THEME.tab.inactiveText = 'var(--text-2)'
            • viền       : 1px solid THEME.tab.inactiveBorder = 'var(--border)'
            • box-shadow : none
          ═══════════════════════════════════════════════════════════════ */}
      <div style={{ display: 'flex', gap: '10px', marginBottom: '8px' }}>
        {(['summary', 'detail'] as const).map(tabKey => {
          const active = activeTab === tabKey;
          return (
            <button
              key={tabKey}
              className="tpl-tab-btn"
              onClick={() => setActiveTab(tabKey)}
              style={{
                padding: '8px 18px', borderRadius: '8px', fontWeight: 700, fontSize: '14px',
                border: `1px solid ${active ? THEME.tab.activeBorder : THEME.tab.inactiveBorder}`,
                background: active ? THEME.tab.activeBg : THEME.tab.inactiveBg,
                color: active ? THEME.tab.activeText : THEME.tab.inactiveText,
                boxShadow: active ? THEME.tab.activeShadow : 'none',
              }}
            >
              {tabKey === 'summary' ? '📊 TÌNH HÌNH TỔNG QUAN' : '💼 CHI TIẾT & DỮ LIỆU'}
            </button>
          );
        })}
      </div>

      {/* ═══════════════════════════════════════════════════════════════
          📋 SPEC — KHUNG 3/6: FILTER BAR (thanh lọc ngày/model/xem theo)
          ───────────────────────────────────────────────────────────────
          KÍCH THƯỚC KHUNG (container ngoài)
            • border-radius : 14px
            • padding       : 10px 14px
            • border        : 1px solid rgba(0,0,0,0.18)
            • box-shadow    : 0 2px 8px rgba(0,0,0,0.08)
            • background    : THEME.filterBarBg = '#2F3A1D' (GIỐNG HỆT header)
            • margin-bottom : 10px
            • bố cục trong  : 1 hàng flex (class .tpl-filterbar-row),
                              gap: 12px, flex-wrap: wrap, align-items: center

          BỐ CỤC BÊN TRONG — 3 vùng, trái → giữa → phải:

          1) [TRÁI] Badge "Đang hiển thị" — class .tpl-filter-left
             - Vị trí   : flex: 0 0 auto (CỐ ĐỊNH, luôn ở sát mép trái)
             - Gồm 2 phần tử con, xếp ngang, gap 6px:
               a) Nhãn tĩnh "Đang hiển thị:"
                  • cỡ chữ 11px, font-weight 600
                  • màu chữ THEME.filterLabelColor = '#C0EF6A'
               b) Pill (viên thuốc) tên model đang xem
                  • padding 3px 8px, border-radius 999px (bo tròn hết)
                  • cỡ chữ 11px, font-weight 700
                  • Nếu là model TỐT NHẤT (isBest = true):
                      - nền  THEME.badge.bestBg   = 'rgba(16,185,129,0.22)'
                      - chữ  THEME.badge.bestText = '#34d399' (xanh lá)
                      - nội dung thêm tiền tố "⭐ " và hậu tố " (Tốt nhất)"
                  • Nếu KHÔNG phải tốt nhất:
                      - nền  THEME.badge.normalBg   = 'rgba(14,165,233,0.22)'
                      - chữ  THEME.badge.normalText = '#38bdf8' (xanh dương)

          2) [GIỮA] 4 field lọc — class .tpl-filter-fields
             - Vị trí   : flex: 1 1 320px, min-width: 0 (DÃN sang phải,
                          flex-wrap để tự xuống hàng khi hẹp), gap 12px 16px
             - MỖI field là 1 khối dọc (label trên, control dưới), gap 4px:
               Nhãn field: cỡ chữ 11px, font-weight 700, uppercase,
                           màu chữ THEME.filterLabelColor = '#C0EF6A'
               (a) "Ngày bắt đầu" → <input type="date">, rộng 130px, cao 38px
               (b) "Ngày kết thúc" → <input type="date">, rộng 130px, cao 38px
               (c) "Model" → <select>, rộng 160px, cao 38px, option đầu
                   "Tất cả"; model tốt nhất có tiền tố "⭐ " trong danh sách
               (d) "Xem theo" → cụm 3 nút Ngày/Tuần/Tháng dính liền nhau,
                   tổng rộng 180px, cao 38px, mỗi nút flex:1:
                     • Đang chọn : nền THEME.viewModeActiveBg = '#2e7d8c',
                                   chữ trắng '#ffffff'
                     • Không chọn: nền THEME.viewModeInactiveBg
                                   = 'rgba(255,255,255,0.55)',
                                   chữ THEME.viewModeInactiveText = '#7A5A2E'

          3) [PHẢI] Nút "Tải Excel" — class .tpl-filter-right
             - Vị trí   : flex: 0 0 auto, margin-left: auto (luôn bám mép
                          phải khung dù màn rộng/hẹp)
             - Kích thước: rộng 120px, cao 38px, border-radius 8px
             - Nền      : rgba(255,255,255,0.9), viền 1px rgba(0,0,0,0.18)
             - Cỡ chữ   : 13px, font-weight 600
             - Hiển thị : icon "⟳" + text "Tải Excel"
          ═══════════════════════════════════════════════════════════════ */}
      <div style={{
        background: THEME.filterBarBg, borderRadius: '14px', padding: '10px 14px',
        border: '1px solid rgba(0,0,0,0.18)', boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
        marginBottom: '10px',
      }}>
        <div className="tpl-filterbar-row">
          {/* Bên trái — CỐ ĐỊNH, không co giãn (flex: 0 0 auto) */}
          <div className="tpl-filter-left">
            <span style={{ fontSize: '11px', fontWeight: 600, color: THEME.filterLabelColor, whiteSpace: 'nowrap' }}>Đang hiển thị:</span>
            <span style={{
              display: 'inline-flex', alignItems: 'center', gap: '4px', padding: '3px 8px',
              borderRadius: '999px', fontWeight: 700, fontSize: '11px', whiteSpace: 'nowrap',
              background: isBest ? THEME.badge.bestBg : THEME.badge.normalBg,
              color: isBest ? THEME.badge.bestText : THEME.badge.normalText,
            }}>
              {isBest && '⭐ '}{showingModel}{isBest && ' (Tốt nhất)'}
            </span>
          </div>

          {/* Giữa — DÃN sang phải, mỗi field là 1 khối label+input tự wrap
              theo khung hình (auto-fit tự nhiên qua flex-wrap của cha) */}
          <div className="tpl-filter-fields">
            <div className="tpl-filter-field">
              <span style={{ fontSize: '11px', fontWeight: 700, color: THEME.filterLabelColor, textTransform: 'uppercase' }}>Ngày bắt đầu</span>
              <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)}
                style={{ width: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', borderRadius: '8px', border: '1px solid rgba(0,0,0,0.18)' }} />
            </div>
            <div className="tpl-filter-field">
              <span style={{ fontSize: '11px', fontWeight: 700, color: THEME.filterLabelColor, textTransform: 'uppercase' }}>Ngày kết thúc</span>
              <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)}
                style={{ width: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', borderRadius: '8px', border: '1px solid rgba(0,0,0,0.18)' }} />
            </div>
            <div className="tpl-filter-field">
              <span style={{ fontSize: '11px', fontWeight: 700, color: THEME.filterLabelColor, textTransform: 'uppercase' }}>Model</span>
              <select value={selectedModel} onChange={e => setSelectedModel(e.target.value)}
                style={{ width: '160px', height: '38px', borderRadius: '8px', border: '1px solid rgba(0,0,0,0.18)', padding: '0 8px' }}>
                <option value="">Tất cả</option>
                {DEMO_MODEL_OPTIONS.map(m => (
                  <option key={m} value={m}>{m === BEST_MODEL_DEMO ? `⭐ ${m}` : m}</option>
                ))}
              </select>
            </div>
            <div className="tpl-filter-field">
              <span style={{ fontSize: '11px', fontWeight: 700, color: THEME.filterLabelColor, textTransform: 'uppercase' }}>Xem theo</span>
              <div style={{ display: 'flex', height: '38px', width: '180px', flexShrink: 0 }}>
                {(['day', 'week', 'month'] as const).map(mode => (
                  <button key={mode} onClick={() => setViewMode(mode)}
                    style={{
                      flex: 1, fontSize: '13px', fontWeight: 600,
                      borderRadius: mode === 'day' ? '6px 0 0 6px' : mode === 'month' ? '0 6px 6px 0' : '0',
                      border: '1px solid rgba(0,0,0,0.18)',
                      borderRight: mode !== 'month' ? 'none' : '1px solid rgba(0,0,0,0.18)',
                      background: viewMode === mode ? THEME.viewModeActiveBg : THEME.viewModeInactiveBg,
                      color: viewMode === mode ? '#ffffff' : THEME.viewModeInactiveText,
                      cursor: 'pointer',
                    }}>
                    {mode === 'day' ? 'Ngày' : mode === 'week' ? 'Tuần' : 'Tháng'}
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* Bên phải — CỐ ĐỊNH, luôn bám sát mép phải khung (margin-left:auto) */}
          <div className="tpl-filter-right">
            <button style={{ height: '38px', width: '120px', borderRadius: '8px', border: '1px solid rgba(0,0,0,0.18)', background: 'rgba(255,255,255,0.9)', fontWeight: 600, fontSize: '13px', cursor: 'pointer' }}>
              ⟳ Tải Excel
            </button>
          </div>
        </div>
      </div>

      {/* ══════════════ TAB "TỔNG QUAN": 4 KPI + 6 khung biểu đồ ══════════════ */}
      {activeTab === 'summary' && (
        <>
          {/* ═══════════════════════════════════════════════════════════
              📋 SPEC — KHUNG 4/6: LƯỚI 4 THẺ KPI
              ───────────────────────────────────────────────────────────
              KÍCH THƯỚC LƯỚI (container .tpl-kpi-grid)
                • display: grid, gap: 12px, margin-bottom: 12px
                • grid-template-columns: repeat(auto-fit, minmax(200px,1fr))
                  → 4 cột trên màn rộng, tự giảm 2 hoặc 1 cột khi màn hẹp

              MỖI THẺ KPI (4 thẻ, trái → phải, DUYỆT theo THEME.kpi[i]):
                • border-left   : 4px solid <accent riêng từng thẻ>
                • background    : linear-gradient(135deg,
                                   <gradientTint riêng> 0%,
                                   rgba(30,41,59,0.4) 100%)
                • border-radius : 10px
                • padding       : 12px 14px
                • Màu accent theo thứ tự trái→phải:
                    Thẻ 1: '#0ea5e9' (xanh dương) — icon "box"
                    Thẻ 2: '#10b981' (xanh lá)    — icon "target"
                    Thẻ 3: '#8b5cf6' (tím)        — icon "check"
                    Thẻ 4: '#f59e0b' (cam)        — icon "trend"

              BỐ CỤC BÊN TRONG MỖI THẺ — 2 hàng:
                Hàng 1 (icon + nhãn), gap 6px, margin-bottom 6px:
                  - icon SVG 16×16px, màu = accent của thẻ
                  - nhãn : cỡ chữ 11.5px, font-weight 700, uppercase,
                           màu chữ 'var(--text-2)'
                Hàng 2 (giá trị + phụ chú), flex hàng ngang, gap 8px:
                  - Giá trị chính : cỡ chữ 22px, font-weight 800
                                    (VD: "96.14%", "0/1", "-0.26pp")
                  - Phụ chú (nếu có): cỡ chữ 12px, màu 'var(--text-3)'
                                    (VD: "Mục tiêu: 96.4%")
              ═══════════════════════════════════════════════════════════ */}
          <div className="tpl-kpi-grid">
            {DEMO_KPIS.map((kpi, i) => {
              const c = THEME.kpi[i];
              return (
                <div key={kpi.label} style={{
                  borderLeft: `4px solid ${c.accent}`,
                  background: `linear-gradient(135deg, ${c.gradientTint} 0%, rgba(30,41,59,0.4) 100%)`,
                  borderRadius: '10px', padding: '12px 14px',
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '6px' }}>
                    <KpiIcon kind={kpi.icon} color={c.accent} />
                    <div style={{ fontSize: '11.5px', fontWeight: 700, color: 'var(--text-2)', textTransform: 'uppercase' }}>{kpi.label}</div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px', flexWrap: 'wrap' }}>
                    <div style={{ fontSize: '22px', fontWeight: 800 }}>{kpi.value}</div>
                    {kpi.sub && <div style={{ fontSize: '12px', color: 'var(--text-3)' }}>{kpi.sub}</div>}
                  </div>
                </div>
              );
            })}
          </div>

          {/* ═══════════════════════════════════════════════════════════
              📋 SPEC — KHUNG 5/6 + 6/6: LƯỚI 6 KHUNG BIỂU ĐỒ
              ───────────────────────────────────────────────────────────
              KÍCH THƯỚC LƯỚI (container .tpl-chart-grid)
                • display: grid, gap: 16px
                • grid-template-columns: repeat(auto-fit, minmax(380px,1fr))
                  → 2 cột trên màn rộng, tự rút còn 1 cột khi màn hẹp/zoom to

              KÍCH THƯỚC MỖI KHUNG BIỂU ĐỒ (container ngoài của 1 panel)
                • border-radius : 12px, overflow: hidden (bo góc cả header
                                  lẫn phần thân bên trong)
                • border        : 1px solid 'var(--border)'

              PHẦN HEADER CỦA MỖI KHUNG (thanh tiêu đề nhỏ phía trên chart)
                • border-left   : 4px solid <accent riêng từng khung>
                • background    : panelBg = bgLight (theme light) hoặc
                                  bgDark (theme dark) — riêng từng khung
                • padding       : 8px 12px
                • cỡ chữ        : 12px, font-weight 700, uppercase
                • màu chữ       : '#1f2937' (light mode) / 'var(--text-1)' (dark)
                • Hiển thị (trái→phải trong header, justify-content:space-between):
                    - Trái : tên khung biểu đồ, VD "BIỂU ĐỒ RADAR 1 (THEO MODEL)"
                    - Phải : ghi chú nhỏ, cỡ chữ 10.5px, opacity 0.7,
                             KHÔNG in hoa (text-transform: none)

              6 KHUNG — thứ tự & màu accent (đọc theo THEME.chartPanels):
                Khung 1: '#8b5cf6' tím    | bgLight '#d6c6fc' | bgDark 'rgba(139,92,246,0.14)'
                Khung 2: '#06b6d4' cyan   | bgLight '#a8e5f0' | bgDark 'rgba(6,182,212,0.14)'
                Khung 3: '#f59e0b' cam    | bgLight '#fcddaa' | bgDark 'rgba(255,255,255,0.05)'
                Khung 4: '#f97316' cam-đỏ| bgLight '#fdcead' | bgDark 'rgba(249,115,22,0.14)'
                Khung 5: '#10b981' xanh lá| bgLight '#abe7d3' | bgDark 'rgba(16,185,129,0.14)'
                Khung 6: '#3b82f6' xanh dương | bgLight '#bad3fc' | bgDark 'rgba(59,130,246,0.14)'

              PHẦN THÂN BIỂU ĐỒ (bên dưới header)
                • height        : 220px (CỐ ĐỊNH — không đổi theo nội dung)
                • padding       : 10px
                • background    : 'var(--surface)'
                • Hiển thị      : PLACEHOLDER SVG demo (RadarPlaceholder cho
                                  khung radar, ComboChartPlaceholder cho khung
                                  còn lại) — khi nhân bản trang thật, THAY
                                  bằng chart thật (Plotly/Recharts) tại đúng
                                  vị trí này, giữ nguyên khung cha bên ngoài.
              ═══════════════════════════════════════════════════════════ */}
          <div className="tpl-chart-grid">
            {DEMO_PANELS.map((panel, i) => {
              const c = THEME.chartPanels[i];
              return (
                <div key={c.key} style={{ borderRadius: '12px', overflow: 'hidden', border: '1px solid var(--border)' }}>
                  <div style={{
                    background: panelBg(c), borderLeft: `4px solid ${c.accent}`,
                    display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                    padding: '8px 12px', fontWeight: 700, fontSize: '12px', textTransform: 'uppercase',
                    color: isLightMode ? '#1f2937' : 'var(--text-1)',
                  }}>
                    <span title={panel.note}>{panel.title}</span>
                    <span style={{ fontSize: '10.5px', opacity: 0.7, textTransform: 'none' }}>{panel.note}</span>
                  </div>
                  <div style={{ height: '220px', padding: '10px', boxSizing: 'border-box', background: 'var(--surface)' }}>
                    {panel.kind === 'radar'
                      ? <RadarPlaceholder accent={c.accent} />
                      : <ComboChartPlaceholder accent={c.accent} />}
                  </div>
                </div>
              );
            })}
          </div>
        </>
      )}

      {/* ═══════════════════════════════════════════════════════════════
          📋 SPEC — KHUNG "CHI TIẾT & DỮ LIỆU": BẢNG + PHÂN TRANG
          ───────────────────────────────────────────────────────────────
          KÍCH THƯỚC KHUNG (container ngoài)
            • border-radius : 12px, overflow: hidden
            • border        : 1px solid 'var(--border)'

          HÀNG TIÊU ĐỀ BẢNG (thead)
            • background : 'var(--surface)'
            • mỗi ô <th> : cỡ chữ 13px, font-weight 700, padding 10px 12px,
                           căn trái, màu 'var(--tbl-head-color, inherit)'
            • Hiển thị 6 cột theo thứ tự: "Cột 1", "Cột 2", "Cột 3",
              "Target", "Actual", "Chênh lệch"

          HÀNG DỮ LIỆU (tbody, mỗi <tr> có border-top 1px 'var(--border)')
            • mỗi ô <td> : padding 8px 12px
            • cột "Cột 1" : font-weight 600 (đậm hơn các cột còn lại)
            • cột "Chênh lệch" : font-weight 700, màu ĐỘNG theo dấu:
                - bắt đầu bằng "+"  → '#10b981' (xanh lá)
                - còn lại (âm)      → '#ef4444' (đỏ)

          THANH PHÂN TRANG (dưới cùng khung)
            • padding    : 10px 15px
            • background : 'var(--surface)'
            • border-top : 1px solid 'var(--border)'
            • bố cục     : flex, justify-content: space-between
            • [Trái] chữ trạng thái, cỡ 12px, màu 'var(--text-2)'
              (VD: "Hiển thị 1-3 / 3 dòng (demo)")
            • [Phải] các nút số trang (.tpl-page-btn):
                - mặc định: nền trong suốt, viền 1px 'var(--border)'
                - trang active: nền '#2563eb', chữ trắng, viền '#2563eb'
          ═══════════════════════════════════════════════════════════════ */}
      {activeTab === 'detail' && (
        <div style={{ border: '1px solid var(--border)', borderRadius: '12px', overflow: 'hidden' }}>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ background: 'var(--surface)' }}>
                  {['Cột 1', 'Cột 2', 'Cột 3', 'Target', 'Actual', 'Chênh lệch'].map(h => (
                    <th key={h} style={{ fontSize: '13px', fontWeight: 700, padding: '10px 12px', textAlign: 'left', color: 'var(--tbl-head-color, inherit)' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {DEMO_ROWS.map((r, i) => (
                  <tr key={i} style={{ borderTop: '1px solid var(--border)' }}>
                    <td style={{ padding: '8px 12px', fontWeight: 600 }}>{r.col1}</td>
                    <td style={{ padding: '8px 12px' }}>{r.col2}</td>
                    <td style={{ padding: '8px 12px' }}>{r.col3}</td>
                    <td style={{ padding: '8px 12px' }}>{r.target}</td>
                    <td style={{ padding: '8px 12px' }}>{r.actual}</td>
                    <td style={{ padding: '8px 12px', fontWeight: 700, color: r.gap.startsWith('+') ? '#10b981' : '#ef4444' }}>{r.gap}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ padding: '10px 15px', background: 'var(--surface)', borderTop: '1px solid var(--border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '12px', color: 'var(--text-2)' }}>Hiển thị 1-{DEMO_ROWS.length} / {DEMO_ROWS.length} dòng (demo)</span>
            <div style={{ display: 'flex', gap: '4px' }}>
              <button className="tpl-page-btn active">1</button>
              <button className="tpl-page-btn" disabled>2</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default DashboardTemplate;

/* ═══════════════════════════════════════════════════════════════════════
 * HƯỚNG DẪN NHÂN BẢN SANG TRANG MỚI (đọc trước khi copy)
 * ─────────────────────────────────────────────────────────────────────
 * 1. Copy nguyên file này thành file mới (vd: SalesDashboardV2.tsx).
 * 2. KHÔNG sửa object `THEME` trừ khi muốn đổi bộ màu chuẩn cho TẤT CẢ
 *    dashboard cùng lúc — nếu chỉ trang mới cần màu khác, tạo THEME riêng
 *    nhưng vẫn giữ đúng cấu trúc field (headerBg, tab.*, kpi[], chartPanels[])
 *    để không lệch khung.
 * 3. Xoá các mảng DEMO_* và thay bằng props/hook lấy dữ liệu thật (Supabase,
 *    props từ App.tsx...) — không đụng vào phần JSX/style, chỉ đổi NGUỒN
 *    dữ liệu đổ vào .map().
 * 4. Thay 2 component `RadarPlaceholder` / `ComboChartPlaceholder` bằng
 *    chart thật (Plotly/Recharts) — giữ nguyên khung `<div>` cha (header
 *    màu + border) bao quanh, chỉ thay phần bên trong khung `height: 220px`.
 * 5. Nếu trang mới cần nhiều/ít hơn 6 khung biểu đồ hoặc 4 thẻ KPI, thêm/bớt
 *    phần tử trong THEME.kpi / THEME.chartPanels TRƯỚC, rồi map dữ liệu demo
 *    theo đúng số lượng đó — tránh trường hợp DEMO_KPIS có 5 phần tử nhưng
 *    THEME.kpi chỉ có 4 màu (sẽ bị `undefined` ở phần tử thứ 5).
 * ═══════════════════════════════════════════════════════════════════════ */
