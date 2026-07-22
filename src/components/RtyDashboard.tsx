import React, { useEffect, useMemo, useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { NeonButton } from './NeonButton';
// EPCC (rty-total-move-to-muc4): chuyển tab "RTY Total" từ Mục 5 sang đây
// (Mục 4) theo yêu cầu — cùng nhóm "RTY" để dễ quan sát/quản lý, tách hẳn
// khỏi Menu5ModelDashboard.tsx (đã dọn sạch, trả về đúng 2 tab gốc).
import { RtyTotalTab } from './RtyTotalTab';

/**
 * ═══════════════════════════════════════════════════════════════════════
 * MỤC 4 — HIỆU SUẤT RTY
 * ═══════════════════════════════════════════════════════════════════════
 * Bản cập nhật: dữ liệu THẬT được trích xuất & nhúng sẵn từ file tham
 * chiếu `Test4.xlsx` (cột gốc: Model, Process, Type, Date, RTY).
 *
 * GHI CHÚ QUAN TRỌNG VỀ PHẠM VI DỮ LIỆU (đã thống nhất với người dùng):
 *  - Test4.xlsx CHỈ có tỉ lệ RTY (%) theo Model/Process/Type/Date — KHÔNG
 *    có cột số lượng (SL sản xuất / SL NG) và KHÔNG có cột Khách hàng.
 *  - Vì vậy toàn bộ ô/cột/KPI vốn cần SL hoặc Khách hàng đã được BỎ hoặc
 *    THAY bằng chỉ số suy ra được thuần từ RTY% (xem từng vị trí bên dưới
 *    có chú thích "// ĐÃ ĐỔI:").
 *  - Đây là bản NHÚNG TĨNH (demo nhanh): dữ liệu là hằng số JS lấy từ
 *    Test4.xlsx tại thời điểm dựng, KHÔNG đọc file động, KHÔNG có nút Tải
 *    Excel hoạt động (giữ nguyên trạng thái disabled như bản gốc).
 *  - Khi có pipeline dữ liệu RTY thật (bucket kiểu `targetActualRows`),
 *    chỉ cần thay khối "EMBEDDED DATA" bên dưới bằng dữ liệu từ props/
 *    Supabase — toàn bộ phần tính toán/chart phía dưới dùng chung logic.
 * ═══════════════════════════════════════════════════════════════════════
 */

type Lang = 'vi' | 'en' | 'ko';

interface RtyDashboardProps {
  theme: 'light' | 'dark';
  onToggleTheme: () => void;
  lang: Lang;
  setLang: (l: Lang) => void;
  /** Giữ tương thích chữ ký với các dashboard khác (Mục 1-3) — hiện chưa
   *  có bucket dữ liệu RTY thật (upload) nên nút "Tải Excel" vẫn ở trạng
   *  thái chờ (disabled). Dữ liệu hiển thị bên dưới là dữ liệu THAM CHIẾU
   *  nhúng sẵn từ Test4.xlsx, không phải dữ liệu do người dùng tải lên. */
  onFileSelected?: (file: File, workbook: any) => void;
  onSyncProgress?: (progress: { bucket: string; done: number; total: number } | null) => void;
}

/* ═══════════════════════════════════════════════════════════════════════
 * EMBEDDED DATA — trích xuất thật từ Test4.xlsx
 * (Model, Process, Type, Date, RTY). Xem ghi chú phạm vi dữ liệu ở trên.
 * ═══════════════════════════════════════════════════════════════════════ */

/** RTY trung bình theo tháng (Target vs Actual) cho 10 Model có đủ dữ liệu
 *  tháng trong Test4.xlsx, tách theo Sub1 / Sub2 / Main / TTL. */
interface ModelRty { target: number; actual: number }
interface ModelSummary {
  model: string;
  sub1?: ModelRty;
  sub2?: ModelRty;
  main?: ModelRty;
  ttl: ModelRty;
}

const MODEL_SUMMARY: ModelSummary[] = [
  { model: 'SO1B01', sub1: { target: 0.992, actual: 0.9929 }, sub2: { target: 0.988, actual: 0.9325 }, main: { target: 0.983, actual: 0.8497 }, ttl: { target: 0.964, actual: 0.7832 } },
  { model: 'SO1C00', sub1: { target: 0.992, actual: 0.9754 }, sub2: { target: 0.988, actual: 0.9369 }, ttl: { target: 0.964, actual: 0.9139 } },
  { model: 'SO1C2G', sub1: { target: 0.992, actual: 0.9946 }, sub2: { target: 0.988, actual: 0.9771 }, main: { target: 0.983, actual: 0.9548 }, ttl: { target: 0.964, actual: 0.9283 } },
  { model: 'SO1C2EDL', sub1: { target: 0.992, actual: 0.9933 }, sub2: { target: 0.988, actual: 0.9835 }, main: { target: 0.983, actual: 0.9311 }, ttl: { target: 0.964, actual: 0.9316 } },
  { model: 'SO1C2EF', sub1: { target: 0.992, actual: 0.9944 }, sub2: { target: 0.988, actual: 0.9830 }, main: { target: 0.983, actual: 0.9581 }, ttl: { target: 0.964, actual: 0.9468 } },
  { model: 'SO1C30 S25', sub1: { target: 0.992, actual: 0.9912 }, sub2: { target: 0.988, actual: 0.9726 }, main: { target: 0.983, actual: 0.9677 }, ttl: { target: 0.964, actual: 0.9476 } },
  { model: 'SO3560', sub1: { target: 0.992, actual: 0.9963 }, sub2: { target: 0.988, actual: 0.9779 }, main: { target: 0.983, actual: 0.9831 }, ttl: { target: 0.964, actual: 0.9579 } },
  { model: 'SO1C2EDM', sub1: { target: 0.992, actual: 0.9948 }, sub2: { target: 0.988, actual: 0.9866 }, main: { target: 0.983, actual: 0.9700 }, ttl: { target: 0.964, actual: 0.9580 } },
  { model: 'SO2701', sub1: { target: 0.992, actual: 0.9933 }, sub2: { target: 0.988, actual: 0.9790 }, main: { target: 0.983, actual: 0.9854 }, ttl: { target: 0.964, actual: 0.9582 } },
  { model: 'SO1C2EH', sub1: { target: 0.992, actual: 0.9951 }, sub2: { target: 0.988, actual: 0.9834 }, main: { target: 0.983, actual: 0.9698 }, ttl: { target: 0.964, actual: 0.9614 } },
];

/** ═══════════════════════════════════════════════════════════════════════
 * FIX (model-dropdown-sort-desc, EPCC)
 * ───────────────────────────────────────────────────────────────────────
 * EXPLORE: dropdown "Model" trước đây liệt kê theo đúng thứ tự các dòng
 * xuất hiện trong Test4.xlsx (thứ tự trích xuất gốc) — không theo hiệu
 * suất, nên model tốt nhất có thể nằm bất kỳ đâu, người dùng phải dò cả
 * danh sách mới thấy.
 * PLAN: tạo thêm 1 mảng ĐÃ SẮP XẾP riêng cho dropdown (không đụng tới thứ
 * tự gốc của MODEL_SUMMARY, vì mảng gốc còn được dùng ở nơi khác — ví dụ
 * `modelsAll` cho biểu đồ mạng nhện — nên không nên đổi thứ tự tại nguồn
 * để tránh ảnh hưởng dây chuyền ngoài ý muốn).
 * CODE: sắp xếp giảm dần theo chênh lệch Actual − Target ở mức TTL (đúng
 * định nghĩa "hiệu suất tốt nhất" đang dùng cho BEST_MODEL bên dưới) —
 * model đạt/gần đạt mục tiêu nhất lên đầu, lùi dần xuống model yếu nhất.
 * CHECK: BEST_MODEL suy ra từ phần tử đầu mảng đã sắp xếp phải khớp với
 * kết quả tính trước đây (SO1C2EH — vẫn đúng vì cùng công thức so sánh).
 * ═══════════════════════════════════════════════════════════════════════ */
const MODEL_SUMMARY_SORTED: ModelSummary[] = [...MODEL_SUMMARY].sort(
  (a, b) => (b.ttl.actual - b.ttl.target) - (a.ttl.actual - a.ttl.target)
);

/** Model dùng làm nguồn cho 4 biểu đồ TTL/MAIN/SUB1/SUB2 khi chưa chọn gì
 *  cụ thể ở dropdown. Cả 10 model trong MODEL_SUMMARY giờ đều có dữ liệu
 *  RTY THẬT ở đủ 3 mức Ngày/Tuần/Tháng, trích xuất trực tiếp từ Test4.xlsx
 *  (xem MODEL_SERIES bên dưới) — độ phủ từng model khác nhau tuỳ thực tế
 *  phát sinh dữ liệu trong file gốc, mốc nào không có số liệu thật sẽ để
 *  trống (null) thay vì suy diễn/generate. */
const HERO_MODEL = 'SO3560';

/** ═══════════════════════════════════════════════════════════════════════
 * MODEL HIỆU SUẤT TỐT NHẤT — dùng làm giá trị MẶC ĐỊNH cho 4 thẻ KPI
 * ───────────────────────────────────────────────────────────────────────
 * Trước đây 4 thẻ KPI mặc định gộp TRUNG BÌNH cả 10 model ("Tất cả") nên
 * lúc mới vào trang số liệu trông rất yếu (VD: Tỷ lệ đạt RTY 0.0%, 0/10
 * model đạt mục tiêu) dù thực tế có nhiều model đạt/gần đạt mục tiêu.
 * Theo yêu cầu: khi mới vào, 4 thẻ KPI sẽ tự chọn Model có HIỆU SUẤT TỐT
 * NHẤT (actual RTY (TTL) chênh lệch dương/cao nhất so với mục tiêu) để
 * hiển thị — hoàn toàn tính từ dữ liệu THẬT trong MODEL_SUMMARY, không
 * suy diễn. Sau đó người dùng vào dropdown "Model" ở thanh filter để chọn
 * model khác bất kỳ, 4 thẻ KPI + 4 biểu đồ + bảng chi tiết sẽ NHẢY SỐ
 * theo model vừa chọn (cơ chế lọc dùng chung `selectedModel`, xem
 * `filteredModels` / `kpi` bên dưới). Giờ lấy trực tiếp từ phần tử đầu
 * của MODEL_SUMMARY_SORTED (đã sắp xếp giảm dần) thay vì tự reduce lại,
 * tránh trùng lặp logic so sánh. */
const BEST_MODEL: string = MODEL_SUMMARY_SORTED[0].model;

/** ═══════════════════════════════════════════════════════════════════════
 * "DỮ LIỆU CẬP NHẬT ĐẾN" — mốc ngày cuối cùng THỰC SỰ có phát sinh dữ liệu
 * RTY trong file tham chiếu Test4.xlsx (không phải ngày hệ thống/đồng hồ).
 * File này là bản NHÚNG TĨNH (xem ghi chú ở đầu component) nên mốc này
 * được xác định trực tiếp từ chính dữ liệu THẬT đã trích xuất, không suy
 * diễn — khi nào có pipeline đọc file RTY động, thay hằng số này bằng giá
 * trị tính toán tương tự (ví dụ ngày lớn nhất có Actual ≠ null) như đã áp
 * dụng ở Mục 2 (Báo cáo doanh số).
 * ═══════════════════════════════════════════════════════════════════════ */
const LAST_DATA_UPDATE_LABEL = '07/08';

/** Nhãn trục X dùng chung cho toàn bộ 4 biểu đồ (đồng bộ giữa các Type
 *  trong cùng 1 chart, giống hệt cách trục X thống nhất trong ảnh mẫu). */
const MONTH_LABELS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL'];
const WEEK_LABELS = ['W01', 'W02', 'W03', 'W04', 'W05', 'W06', 'W07', 'W08', 'W09', 'W10', 'W11', 'W12', 'W13', 'W14', 'W15', 'W16', 'W17', 'W18', 'W19', 'W20', 'W21', 'W22', 'W23', 'W24', 'W25', 'W26', 'W27', 'W28'];
const DAY_LABELS = ['01/02', '01/03', '01/05', '01/06', '01/07', '01/08', '01/09', '01/10', '01/12', '01/13', '01/14', '01/15', '01/16', '01/17', '01/18', '01/19', '01/20', '01/21', '01/22', '01/23', '01/24', '01/25', '01/26', '01/27', '01/28', '01/29', '01/30', '01/31', '02/01', '02/02', '02/03', '02/04', '02/05', '02/06', '02/07', '02/08', '02/09', '02/10', '02/11', '02/12', '02/13', '02/22', '02/23', '02/24', '02/25', '02/26', '02/27', '02/28', '03/01', '03/02', '03/03', '03/04', '03/05', '03/06', '03/07', '03/08', '03/09', '03/10', '03/11', '03/12', '03/13', '03/14', '03/16', '03/17', '03/18', '03/19', '03/20', '03/21', '03/22', '03/23', '03/24', '03/25', '03/26', '03/27', '03/28', '03/29', '03/30', '03/31', '04/01', '04/02', '04/03', '04/04', '04/05', '04/06', '04/07', '04/08', '04/09', '04/10', '04/11', '04/12', '04/13', '04/14', '04/15', '04/16', '04/17', '04/18', '04/19', '04/20', '04/21', '04/22', '04/23', '04/24', '04/25', '04/26', '04/27', '04/28', '04/29', '05/02', '05/03', '05/04', '05/05', '05/06', '05/07', '05/08', '05/09', '05/10', '05/11', '05/12', '05/13', '05/14', '05/15', '05/16', '05/17', '05/18', '05/19', '05/20', '05/21', '05/22', '05/23', '05/25', '05/26', '05/27', '05/28', '05/29', '05/30', '06/01', '06/02', '06/03', '06/04', '06/05', '06/06', '06/08', '06/09', '06/10', '06/11', '06/12', '06/13', '06/15', '06/16', '06/17', '06/18', '06/19', '06/20', '06/22', '06/23', '06/24', '06/25', '06/26', '06/27', '06/29', '06/30', '07/01', '07/02', '07/03', '07/04', '07/06', '07/07', '07/08', '07/09'];

/** RTY MỤC TIÊU (%) cố định theo từng cấp — lấy từ Test4.xlsx / xác nhận
 *  trước đó với người dùng (không có dòng Target riêng trong file gốc). */
const TARGET_TTL = 0.964;
const TARGET_MAIN = 0.983;
const TARGET_SUB1 = 0.992;
const TARGET_SUB2 = 0.988;

const MODEL_SERIES: Record<string, Record<string, { month: (number | null)[]; week: (number | null)[]; day: (number | null)[] }>> = {
  'SO1B01': {
    RTY_TTL: { month: [null, null, 0.7319, 0.7958, 0.7626, 0.8331, 0.7926], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.8472, 0.6702, 0.7371, 0.7487, 0.8796, 0.7837, 0.7565, 0.6732, 0.8205, 0.8172, 0.8812, 0.9026, 0.8243, 0.7446, 0.7454, 0.8551], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 0.8446, 0.9759, null, 0.8312, 0.5911, 0.701, 0.8274, 0.7098, null, null, 0.7653, 0.8193, 0.6844, 0.84, 0.8474, 1, null, null, 0.9712, 0.9383, 0.9592, 0.7169, 0.9275, 0.9178, 0.9264, 0.8793, 0.9259, 0.8547, 0.8369, 0.8567, 0.8771, 0.8182, 0.7416, 0.8507, 0.712, 0.7939, 0.7132, 0.716, 0.712, 0.716, 0.7525, 0.8499, 0.8428, 0.7937, 0.5805, 0.4072, 0.6869, 0.6994, 0.7833, 0.7946, 0.7482, 0.8694, 0.8458, 0.8717, 0.7952, 0.7923, 0.872, 0.8139, 0.7979, 0.862, 0.6713, 0.8884, 0.8157, 0.8893, 0.8454, 0.9378, 0.894, 0.9071, 0.8846, 0.8821, 0.9053, 0.9044, 0.9214, 0.918, 0.8986, 0.8898, 0.8358, 0.8288, 0.7236, 0.7699, 0.7662, 0.7842, 0.8277, 0.7085, 0.7375, 0.6458, 0.7759, 0.773, 0.6836, 0.7303, 0.7147, 0.7943, 0.7505, 0.8307, 0.9206, 0.9201] },
    RTY_MAIN: { month: [null, null, 0.941, 0.8668, 0.779, 0.8518, 0.8099], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.9759, 0.8856, 0.8436, 0.8219, 0.9097, 0.8032, 0.7735, 0.6859, 0.8369, 0.837, 0.9001, 0.9276, 0.84, 0.7599, 0.762, 0.8748], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9759, null, null, 0.9056, 0.9349, 0.94, 0.806, null, null, null, 0.9199, 0.7272, 0.8948, 0.9101, null, null, null, null, 0.9383, null, 0.8385, null, null, 0.9557, 0.9193, 0.9692, 0.8877, 0.8589, 0.8824, 0.896, 0.8379, 0.76, 0.8764, 0.725, 0.8154, 0.7369, 0.7356, 0.7317, 0.7341, 0.7649, 0.8613, 0.854, 0.8063, 0.5919, 0.413, 0.6996, 0.7177, 0.7987, 0.7962, 0.7624, 0.8821, 0.8629, 0.889, 0.8148, 0.8092, 0.8873, 0.8288, 0.8165, 0.8872, 0.6908, 0.9135, 0.8454, 0.9081, 0.8628, 0.9519, 0.911, 0.9218, 0.9172, 0.9111, 0.9283, 0.9287, 0.94, 0.9404, 0.9132, 0.9086, 0.8506, 0.8395, 0.7385, 0.7896, 0.7775, 0.7973, 0.8471, 0.7224, 0.7547, 0.6624, 0.7941, 0.7941, 0.6997, 0.7472, 0.7286, 0.8078, 0.7724, 0.8469, 0.9406, 0.9402] },
    RTY_SUB1: { month: [null, null, 0.9813, 0.9907, 0.9971, 0.9977, 0.9975], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.9706, 0.9864, 0.9833, 0.9844, 0.9957, 0.9974, 0.9966, 0.9975, 0.9972, 0.9977, 0.9979, 0.9977, 0.9971, 0.998, 0.9977, 0.9974], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9706, null, null, 0.9829, 0.9902, 0.9921, 0.9806, null, null, null, 0.9735, null, 1, 0.991, 0.9853, null, null, null, null, null, 0.9918, 0.9961, 0.9934, 0.96, 0.9945, 0.9958, 0.9963, 0.9948, 0.9971, 0.993, 0.9985, 0.9981, 0.9988, 0.9983, 0.9966, 0.9951, 0.9945, 0.9948, 0.9976, 0.9957, 0.9971, 0.9983, 0.9983, 0.9979, 0.9974, 0.998, 0.9971, 0.9977, 0.9962, 0.9981, 0.9976, 0.997, 0.9969, 0.9966, 0.9978, 0.9976, 0.9978, 0.9973, 0.9973, 0.9979, 0.9977, 0.9985, 0.9979, 0.9984, 0.9974, 0.9985, 0.9977, 0.9976, 0.998, 0.9977, 0.997, 0.9978, 0.9984, 0.9973, 0.9982, 0.9971, 0.9971, 0.9972, 0.9956, 0.9973, 0.9984, 0.998, 0.9979, 0.9972, 0.9989, 0.9978, 0.9989, 0.9971, 0.9972, 0.9972, 0.9979, 0.9979, 0.9975, 0.9976, 0.997, 0.9976] },
    RTY_SUB2: { month: [null, null, 0.7926, 0.9266, 0.9818, 0.9803, 0.981], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.8944, 0.7671, 0.8885, 0.9253, 0.971, 0.9783, 0.9814, 0.9839, 0.9832, 0.9785, 0.9811, 0.9752, 0.9842, 0.9818, 0.9805, 0.98], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 0.8702, null, null, 0.8456, 0.6592, 0.7558, 0.8977, 0.8806, null, null, 0.7861, 0.8906, 0.9412, 0.9473, 0.945, 1, null, null, 0.9712, null, 0.9672, 0.8583, 0.9336, 0.9561, 0.9747, 0.9606, 0.9589, 0.9679, 0.9773, 0.9778, 0.9803, 0.9783, 0.977, 0.9723, 0.9854, 0.9783, 0.9731, 0.9785, 0.9754, 0.9795, 0.9867, 0.9885, 0.9885, 0.9864, 0.9833, 0.9877, 0.9847, 0.9768, 0.9844, null, 0.9838, 0.9887, 0.9832, 0.9839, 0.978, 0.9815, 0.9849, 0.9847, 0.98, 0.9736, 0.974, 0.974, 0.9668, 0.9809, 0.9824, 0.9867, 0.9835, 0.9865, 0.9663, 0.9704, 0.9781, 0.9759, 0.9818, 0.9787, 0.9858, 0.9821, 0.9854, 0.99, 0.9843, 0.9777, 0.9871, 0.9855, 0.9792, 0.9835, 0.9783, 0.9772, 0.9782, 0.9763, 0.9799, 0.9801, 0.9829, 0.9854, 0.9741, 0.9833, 0.9817, 0.981] },
    MAIN_FVI: { month: [null, null, null, 0.9896, 0.9937, 0.995, 0.9951], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.978, null, null, 0.9918, 0.9929, 0.9945, 0.9931, 0.993, 0.9939, 0.995, 0.9951, 0.9957, 0.9946, 0.9938, 0.9959], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9922, 0.9637, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9934, 0.9932, 0.9924, 0.9914, 0.991, 0.9919, 0.9896, 0.9883, 0.9942, 0.9943, 0.9935, 0.9944, 0.9954, 0.9946, 0.9942, 0.9944, 0.9942, 0.9947, 0.994, 0.9942, 0.9912, 0.9948, 0.9949, 0.992, 0.9907, 0.9936, 0.9923, 0.993, 0.9939, 0.9926, 0.9931, 0.9933, 0.9944, 0.9935, 0.9944, 0.9926, 0.9949, 0.9937, 0.9943, 0.9951, 0.9947, 0.9951, 0.9954, 0.9953, 0.9953, 0.9948, 0.9937, 0.9966, 0.9953, 0.9951, 0.9951, 0.9957, 0.9964, 0.9968, 0.996, 0.9943, 0.9947, 0.9952, 0.9956, 0.9954, 0.9944, 0.9925, 0.993, 0.993, 0.9942, 0.9921, 0.9943, 0.9963, 0.9966, 0.9963, 0.9952, 0.9953] },
    MAIN_ASSY: { month: [null, null, 0.9566, 0.9607, 0.9886, 0.9952, 0.9785], week: [null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9568, 0.8998, 0.9576, 0.9798, 0.9886, 0.988, 0.9847, 0.9864, 0.9963, 0.9976, 0.9978, 0.9988, 0.9982, 0.9743, 0.9759], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, 0.9133, 0.9508, 0.9864, 0.9767, null, null, null, 0.9199, 0.8462, 0.9006, 0.9327, null, null, null, null, 0.9383, null, 0.9769, null, null, 0.9744, 0.9832, 0.9921, 0.9441, 0.9792, 0.9934, 0.992, 0.9903, 0.9878, 0.9892, 0.986, 0.9895, 0.9916, 0.9892, 0.9864, 0.988, 0.9942, 0.9919, 0.975, 0.9819, 0.9678, 0.99, 0.9931, 0.9964, 0.9814, 0.9821, 0.9763, 0.9715, 0.9862, 0.9916, 0.9953, 0.9976, 0.9947, 0.9968, 0.9954, 0.9957, 0.9971, 0.9979, 0.9968, 0.9965, 0.9984, 0.9976, 0.9978, 0.9986, 0.999, 0.999, 0.9957, 0.9986, 0.9989, 0.9953, 0.9966, 0.9991, 0.9993, 0.9994, 0.9992, 0.9993, 0.9986, 0.9993, 0.9991, 0.9978, 0.996, 0.9983, 0.9608, 0.9608, 0.9614, 0.9852, 0.9859, 0.9919, 0.9705, 0.9808, 0.9733, 0.9787] },
    MAIN_DRIVING: { month: [null, null, 0.9837, 0.9389, 0.8907, 0.9782, 0.9781], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.9759, 0.9514, 0.9458, 0.8858, 0.9527, 0.8807, 0.8679, 0.7968, 0.9694, 0.9702, 0.981, 0.9826, 0.9782, 0.9737, 0.9725, 0.9824], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9759, null, null, 0.9915, 0.9862, 0.963, 0.8649, null, null, null, null, 0.8594, 0.9936, 0.9844, null, null, null, null, null, null, 0.8858, null, null, 0.993, 0.952, 0.9853, 0.9524, 0.9053, 0.9387, 0.9426, 0.9272, 0.88, 0.948, 0.7809, 0.8674, 0.8357, 0.8439, 0.8514, 0.8651, 0.8779, 0.8932, 0.908, 0.8808, 0.7494, 0.5137, 0.7871, 0.7805, 0.9263, 0.9396, 0.9642, 0.9694, 0.9603, 0.9743, 0.9731, 0.9752, 0.9775, 0.9779, 0.9756, 0.9637, 0.9531, 0.9736, 0.9775, 0.982, 0.9799, 0.9839, 0.9791, 0.9836, 0.9826, 0.9794, 0.9796, 0.9814, 0.9879, 0.9845, 0.977, 0.9792, 0.9821, 0.9787, 0.9767, 0.9753, 0.9739, 0.979, 0.9769, 0.9714, 0.9768, 0.9643, 0.9696, 0.9696, 0.9688, 0.9684, 0.9725, 0.9858, 0.9799, 0.981, 0.9851, 0.9837] },
    MAIN_TILT: { month: [null, null, 1, 0.971, 0.8903, 0.8794, 0.8505], week: [null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9948, 0.9913, 0.969, 0.9826, 0.9291, 0.907, 0.8804, 0.8813, 0.8712, 0.9243, 0.9508, 0.8635, 0.7861, 0.8092, 0.9163], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, 0.9972, 0.9973, 0.99, null, null, null, null, null, null, 0.9913, null, null, null, null, null, null, 0.969, null, null, 0.9943, 0.9888, 0.9991, 0.9959, 0.9777, 0.954, 0.9683, 0.9234, 0.8793, 0.9399, 0.9477, 0.9554, 0.8934, 0.886, 0.8764, 0.8638, 0.8815, 0.9774, 0.9705, 0.9377, 0.8233, 0.8164, 0.8996, 0.9303, 0.8869, 0.8684, 0.8162, 0.9433, 0.9168, 0.927, 0.8472, 0.8374, 0.9177, 0.8558, 0.8454, 0.9315, 0.7306, 0.9462, 0.8727, 0.9326, 0.8866, 0.9745, 0.9368, 0.9428, 0.9388, 0.936, 0.9578, 0.9509, 0.9571, 0.9644, 0.9425, 0.9327, 0.8699, 0.861, 0.7597, 0.8148, 0.8038, 0.8189, 0.8717, 0.7488, 0.7801, 0.6934, 0.8583, 0.8583, 0.7556, 0.7894, 0.7644, 0.8291, 0.8149, 0.8833, 0.9857, 0.9812] },
    SUB1_FPCB: { month: [null, null, 0.9953, 0.9963, 0.999, 0.9991, 0.9989], week: [null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9954, 0.9967, 0.9881, 0.999, 0.9989, 0.999, 0.9991, 0.999, 0.9992, 0.9992, 0.9991, 0.9988, 0.9992, 0.9991, 0.9988], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, 0.9915, 0.9944, 0.9984, 0.9972, null, null, null, 0.993, null, 1, 0.9972, null, null, null, null, null, null, 0.9949, 0.9973, 1, 0.96, 0.9988, 0.9988, 0.9998, 0.9982, 0.9995, 0.9983, 0.9998, 0.9992, 0.9995, 0.999, 0.9985, 0.9984, 0.9982, 0.9988, 0.9994, 0.9986, 0.999, 0.9994, 0.9994, 0.9995, 0.999, 0.9992, 0.9989, 0.9996, 0.9984, 0.9993, 0.9987, 0.9992, 0.9988, 0.9987, 0.9993, 0.9991, 0.9992, 0.9988, 0.9989, 0.9993, 0.9991, 0.9997, 0.9992, 0.9993, 0.9989, 0.9994, 0.9992, 0.9991, 0.9993, 0.9992, 0.9986, 0.9991, 0.9996, 0.999, 0.999, 0.9988, 0.9987, 0.9987, 0.9983, 0.999, 0.9993, 0.9991, 0.9991, 0.9988, 0.9996, 0.9991, 0.9995, 0.9991, 0.999, 0.9986, 0.9992, 0.9991, 0.9989, 0.999, 0.9984, 0.9991] },
    SUB1_FVI: { month: [null, null, 0.9859, 0.9944, 0.9981, 0.9986, 0.9986], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.9706, 0.991, 0.9865, 0.9963, 0.9967, 0.9985, 0.9976, 0.9984, 0.9983, 0.9986, 0.9987, 0.9986, 0.9983, 0.9988, 0.9986, 0.9986], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9706, null, null, 0.9914, 0.9958, 0.9936, 0.9833, null, null, null, 0.9804, null, null, 0.9938, 0.9853, null, null, null, null, null, 0.9968, 0.9987, 0.9934, null, 0.9958, 0.997, 0.9965, 0.9966, 0.9976, 0.9947, 0.9988, 0.9989, 0.9993, 0.9993, 0.9981, 0.9967, 0.9963, 0.996, 0.9982, 0.9971, 0.998, 0.9988, 0.9988, 0.9984, 0.9983, 0.9988, 0.9982, 0.9981, 0.9978, 0.9988, 0.9989, 0.9977, 0.998, 0.9979, 0.9985, 0.9985, 0.9986, 0.9985, 0.9984, 0.9986, 0.9985, 0.9988, 0.9987, 0.999, 0.9985, 0.9991, 0.9985, 0.9985, 0.9987, 0.9985, 0.9984, 0.9987, 0.9988, 0.9983, 0.9992, 0.9983, 0.9984, 0.9985, 0.9973, 0.9983, 0.9991, 0.9988, 0.9988, 0.9984, 0.9993, 0.9987, 0.9994, 0.9979, 0.9982, 0.9986, 0.9987, 0.9988, 0.9986, 0.9986, 0.9986, 0.9985] },
    SUB2_HOOK: { month: [null, null, 0.9854, 0.9994, 0.9998, 0.9998, 0.9999], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.9914, 0.9904, 0.9996, 0.9997, 0.9997, 0.9998, 0.9999, 0.9998, 0.9995, 1.0, 0.9998, 0.9997, 0.9998, 0.9999, 0.9999, 0.9998], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 0.9828, null, null, 0.9679, 0.9911, 0.9965, 0.9964, 1, null, null, 1, 0.998, 1, 1, null, 1, null, null, null, null, 1, 1, 0.999, 0.9999, 1, 0.9994, 0.9995, 0.9996, 0.9999, 0.9996, 1, 0.9998, 0.9997, 0.9996, 0.9999, 1, 1, 1, 0.9998, 1, 1, 0.9999, 0.9999, 0.9998, 0.9999, 0.9998, 0.9998, 0.9999, 0.9999, null, 0.9992, 0.9998, 0.9996, 0.9998, 0.9996, 0.999, 1, 1, 1, 0.9999, 0.9999, 0.9999, 0.9997, 0.9999, 0.9998, 0.9998, 1, 0.9998, 0.9993, 0.9994, 0.9999, 0.9999, 0.9999, 0.9996, 0.9998, 0.9999, 0.9999, 0.9998, 0.9998, 0.9997, 0.9999, 0.9998, 0.9999, 0.9998, 1, 1, 0.9997, 1, 0.9999, 1, 1, 1, 0.9999, 1, 0.9999, 0.9992] },
    SUB2_OVEN: { month: [null, null, 1, 1, 1, 1, 1], week: [null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 1, null, null, 1, 1, 1, 1, 1, null, null, 1, 1, 1, 1, null, 1, null, null, null, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1] },
    SUB2_INDEX: { month: [null, null, 0.9074, 0.9499, 0.9916, 0.9956, 0.9954], week: [null, null, null, null, null, null, null, null, null, null, null, null, 0.9191, 0.8909, 0.9219, 0.949, 0.9831, 0.9881, 0.9895, 0.9928, 0.9924, 0.9927, 0.9967, 0.995, 0.9952, 0.9958, 0.9958, 0.9945], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9191, null, null, 0.8807, 0.9224, 0.8189, 0.9292, 0.903, null, null, 0.7981, 0.9141, 0.9758, 0.9767, 0.945, null, null, null, null, null, 0.9768, 0.8844, 0.955, 0.98, 0.9879, 0.977, 0.9772, 0.9762, 0.985, 0.9855, 0.9928, 0.9897, 0.9876, 0.983, 0.9908, 0.9895, 0.9828, 0.9877, 0.9839, 0.9874, 0.9944, 0.9951, 0.9951, 0.9947, 0.9901, 0.9954, 0.9912, 0.9902, 0.995, null, 0.9946, 0.9969, 0.9924, 0.9915, 0.9869, 0.9922, 0.9933, 0.994, 0.9927, 0.9919, 0.99, 0.9944, 0.9961, 0.9958, 0.9967, 0.9972, 0.9959, 0.9988, 0.9979, 0.9932, 0.9928, 0.9923, 0.9971, 0.9971, 0.9966, 0.9918, 0.9969, 0.998, 0.9957, 0.9922, 0.9954, 0.9984, 0.9912, 0.9969, 0.9964, 0.9966, 0.995, 0.9946, 0.996, 0.9967, 0.9962, 0.9962, 0.9922, 0.9952, 0.9962, 0.9945] },
  },
  'SO1C00': {
    RTY_TTL: { month: [null, 0.9416, null, null, null, 0.8862, null], week: [null, null, null, null, null, 0.9497, 0.9751, 0.9717, 0.9458, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9662, 0.9172, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.996, 0.9527, 0.9916, null, 0.9751, null, null, null, null, 0.9717, null, null, null, null, 0.9458, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9896, 0.961, 0.9907, null, 0.9347, 0.8998, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_MAIN: { month: [null, null, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB1: { month: [null, 0.9756, null, null, null, 0.9752, null], week: [null, null, null, null, null, 0.9907, 0.9751, null, 0.9458, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9752, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9899, 0.9916, null, 0.9751, null, null, null, null, null, null, null, null, null, 0.9458, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9896, 0.961, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB2: { month: [null, 0.9651, null, null, null, 0.9087, null], week: [null, null, null, null, null, 0.9586, null, 0.9717, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9907, 0.9172, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.996, 0.9624, null, null, null, null, null, null, null, 0.9717, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9907, null, 0.9347, 0.8998, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_FVI: { month: [null, null, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_ASSY: { month: [null, null, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_DRIVING: { month: [null, null, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_TILT: { month: [null, null, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FPCB: { month: [null, 0.9916, null, null, null, 0.985, null], week: [null, null, null, null, null, 0.9951, 0.9822, null, 0.9941, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.985, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9944, 0.9957, null, 0.9822, null, null, null, null, null, null, null, null, null, 0.9941, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.996, 0.9739, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FVI: { month: [null, 0.9839, null, null, null, 0.9901, null], week: [null, null, null, null, null, 0.9957, 0.9928, null, 0.9514, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9901, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9954, 0.9959, null, 0.9928, null, null, null, null, null, null, null, null, null, 0.9514, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9935, 0.9867, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_HOOK: { month: [null, 1, null, null, null, 0.9977, null], week: [null, null, null, null, null, 1, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9977, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9955, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_OVEN: { month: [null, 1, null, null, null, 1, null], week: [null, null, null, null, null, 1, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_INDEX: { month: [null, 0.9951, null, null, null, 0.9476, null], week: [null, null, null, null, null, 0.9937, null, 0.9965, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9476, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9937, null, null, null, null, null, null, null, 0.9965, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9634, 0.9319, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
  },
  'SO1C2G': {
    RTY_TTL: { month: [0.8536, 0.9301, 0.9261, 0.9379, 0.9433, 0.953, 0.9538], week: [null, null, null, 0.944, 0.8541, 0.9221, 0.9297, 0.9476, 0.9314, 0.9059, 0.9497, 0.9293, 0.9385, 0.9185, 0.9356, 0.9502, 0.9449, 0.9444, 0.9412, 0.9473, 0.9429, 0.9408, 0.9504, 0.9502, 0.951, 0.9598, 0.9561, 0.9506], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9588, 0.9846, null, null, null, 0.9583, 0.8792, 0.8939, 0.9553, 0.6895, 0.9147, 0.9622, 0.8794, 0.9418, 0.943, 0.8996, 0.9425, 0.9307, 0.9217, 0.9525, 0.912, 0.909, 0.9398, 0.9351, 0.9476, 0.9542, 0.9558, 0.9294, 0.9309, 0.9394, 0.8926, 0.9602, 0.9392, 0.919, 0.9265, 0.9051, 0.8806, 0.9527, 0.8516, 0.9758, 0.975, 0.9641, 0.9808, 0.9979, 0.9517, 0.9732, 0.9638, 0.9725, null, 0.9699, 0.9477, 0.9648, 0.9481, 0.9611, 0.9206, 0.9521, 0.9503, 0.9083, null, 0.9399, 0.9303, 0.9349, 0.9471, 0.9396, 0.8656, 0.8713, 0.9004, 0.9501, 0.9378, 0.9316, 0.9333, 0.9438, 0.9509, 0.9488, 0.9503, 0.9467, 0.9553, 0.9483, 0.9522, 0.9496, 0.9333, 0.948, 0.9537, 0.9417, 0.9424, 0.9505, null, 0.9399, 0.9318, 0.9593, 0.9563, 0.9346, 0.945, 0.9274, 0.916, 0.9445, 0.9434, 0.9556, 0.9568, 0.9477, 0.9462, 0.9394, 0.9478, 0.9489, 0.9496, 0.9514, 0.9559, 0.9395, 0.9359, 0.9358, 0.9449, 0.9451, 0.9516, 0.956, 0.9535, 0.957, 0.8979, 0.9295, 0.9539, 0.9418, 0.9545, 0.9485, 0.9567, 0.9471, 0.951, 0.9413, 0.9448, 0.9534, 0.9592, 0.9518, 0.949, 0.9421, 0.9456, 0.9602, 0.9536, 0.9554, 0.9612, 0.9631, 0.9613, 0.9494, 0.96, 0.964, 0.9578, 0.9508, 0.9561, 0.964, 0.9496, 0.9586, 0.9463, 0.9568, 0.9598, 0.9397] },
    RTY_MAIN: { month: [0.9095, 0.9467, 0.9496, 0.9678, 0.9643, 0.9737, 0.9722], week: [null, null, null, 0.9846, 0.9065, 0.9389, 0.9447, 0.9662, 0.9492, 0.9214, 0.9707, 0.9675, 0.9645, 0.974, 0.9649, 0.9725, 0.9622, 0.9644, 0.9649, 0.9671, 0.9636, 0.9595, 0.9672, 0.9715, 0.9776, 0.9783, 0.975, 0.9691], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9846, null, null, null, null, null, 0.9406, 0.9916, 0.6971, 0.9323, 0.9833, 0.8919, 0.9625, 0.96, 0.913, 0.9556, 0.9488, 0.9403, 0.9686, 0.9323, 0.9258, 0.9525, 0.9444, 0.9662, 0.9664, 0.9692, 0.9498, 0.9488, 0.9551, 0.9068, 0.9886, 0.9573, 0.9345, 0.9435, 0.9215, 0.8946, 0.9527, 0.8528, 0.9915, 0.995, 0.9916, null, null, 0.9732, 0.9732, 0.973, 0.9725, null, 0.9805, 0.9911, 0.9648, 0.9727, 0.9796, 0.9536, 0.9828, 0.9707, 0.9277, null, 0.9772, 0.9604, 0.9823, 0.9772, 0.9672, 0.9763, 0.9777, 0.9787, 0.9824, 0.9631, 0.9458, 0.9523, 0.965, 0.9672, 0.9722, 0.9703, 0.9749, 0.9754, 0.9718, 0.9717, 0.9708, 0.9513, 0.9669, 0.9704, 0.957, 0.9598, 0.9676, null, 0.9625, 0.9471, 0.9745, 0.9659, 0.9723, 0.9694, 0.9563, 0.9417, 0.9653, 0.9665, 0.9797, 0.9754, 0.9669, 0.9666, 0.9607, 0.9691, 0.9691, 0.9691, 0.9682, 0.9736, 0.9585, 0.9562, 0.9643, 0.9608, 0.9681, 0.9695, 0.9701, 0.972, 0.9737, 0.9181, 0.9539, 0.9691, 0.9563, 0.9698, 0.97, 0.9696, 0.9682, 0.9715, 0.9671, 0.9618, 0.976, 0.9782, 0.9742, 0.9764, 0.9717, 0.9797, 0.9793, 0.9755, 0.9832, 0.9769, 0.9789, 0.9797, 0.9738, 0.9814, 0.9793, 0.9809, 0.9677, 0.9766, 0.9762, 0.9717, 0.9772, 0.9644, 0.9766, 0.9764, 0.9589] },
    RTY_SUB1: { month: [0.9924, 0.9966, 0.9942, 0.9948, 0.9945, 0.9955, 0.9947], week: [null, null, null, 1, 0.9908, 0.9973, 0.9962, 0.9968, 0.9966, 0.9969, 0.9938, 0.9865, 0.9944, 0.9938, 0.995, 0.9942, 0.9955, 0.9959, 0.993, 0.994, 0.9954, 0.995, 0.9965, 0.9957, 0.9949, 0.9954, 0.9945, 0.9944], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, null, 0.9941, 0.9957, 0.9933, null, 0.9812, 0.9924, 0.9972, 0.997, 0.9982, 0.9967, 0.9974, 0.9974, null, 0.9945, 0.9973, 0.9961, 0.9949, 0.9984, 0.9968, 0.9969, 0.9968, 0.9976, 0.9973, 0.997, 0.9946, 0.9962, 0.9973, 0.9954, 0.9959, 0.9971, 0.9969, null, 0.9986, 0.9936, 0.9941, 0.9932, null, 0.9979, 0.9903, null, 0.9905, null, null, null, 0.9825, null, 0.9941, 0.9976, null, 0.9912, 0.9942, 0.995, null, 0.9937, 0.9926, 0.995, 0.9961, 0.9924, 0.9922, 0.9943, 0.9953, 0.9949, 0.9947, 0.9943, 0.9945, 0.9938, 0.9978, 0.9959, 0.9954, 0.9932, 0.9932, 0.9926, 0.9962, 0.9933, 0.9954, 0.995, 0.9952, 0.9947, 0.9962, 0.9965, null, 0.9929, 0.9954, 0.997, 0.9979, 0.9965, 0.9928, 0.9919, 0.9928, 0.9931, 0.9924, 0.9956, 0.9922, 0.9955, 0.996, 0.9956, 0.9921, 0.992, 0.9923, 0.9946, 0.9939, 0.9943, 0.9958, 0.9967, 0.9962, 0.9952, 0.9948, 0.9959, 0.9948, 0.9954, 0.9974, 0.9915, 0.997, 0.997, 0.9959, 0.9961, 0.997, 0.9959, 0.9965, 0.9946, 0.9974, 0.9944, 0.9955, 0.9957, 0.9937, 0.9952, 0.9947, 0.9958, 0.9947, 0.9956, 0.9961, 0.9955, 0.9949, 0.994, 0.9963, 0.9957, 0.9955, 0.9916, 0.9942, 0.9956, 0.9952, 0.995, 0.9942, 0.9951, 0.9946, 0.9936] },
    RTY_SUB2: { month: [0.9458, 0.9859, 0.9809, 0.9741, 0.9837, 0.9832, 0.9863], week: [null, null, null, 0.9588, 0.951, 0.9848, 0.9878, 0.984, 0.9846, 0.9862, 0.9844, 0.9736, 0.9785, 0.9489, 0.9744, 0.9828, 0.9865, 0.9832, 0.9824, 0.9854, 0.9831, 0.9855, 0.9862, 0.9824, 0.9777, 0.9856, 0.986, 0.9865], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9588, null, null, null, null, 0.9583, 0.8844, 0.9545, 0.9699, 0.9891, null, 0.9861, 0.9887, 0.9814, 0.984, 0.9886, 0.9888, 0.9835, 0.9803, 0.9889, 0.9808, 0.9857, 0.9918, 0.9918, 0.984, 0.9905, 0.9892, 0.9808, 0.9838, 0.9865, 0.9897, 0.9751, 0.9837, 0.9879, 0.9861, 0.9851, 0.9875, null, 1, 0.9906, 0.9857, 0.9789, 0.9808, null, 0.9875, null, null, null, null, 0.9891, 0.9732, null, 0.9805, 0.9835, 0.9654, 0.9774, 0.9847, 0.9841, null, 0.9678, 0.9759, 0.9565, 0.973, 0.9789, 0.8935, 0.8962, 0.9243, 0.9721, 0.9789, 0.9906, 0.9855, 0.9842, 0.9854, 0.9799, 0.9839, 0.9778, 0.9861, 0.9831, 0.9837, 0.9847, 0.9856, 0.9854, 0.9875, 0.9891, 0.9856, 0.9857, null, 0.9835, 0.9885, 0.9873, 0.9921, 0.9646, 0.9819, 0.9776, 0.9798, 0.9853, 0.9836, 0.9797, 0.9887, 0.9846, 0.9828, 0.9821, 0.9858, 0.9871, 0.9875, 0.988, 0.9879, 0.9858, 0.9829, 0.9737, 0.9872, 0.981, 0.9867, 0.9895, 0.986, 0.9873, 0.9806, 0.9827, 0.9873, 0.9878, 0.9884, 0.9815, 0.9897, 0.9822, 0.9823, 0.9786, 0.985, 0.9823, 0.985, 0.9812, 0.978, 0.9742, 0.9704, 0.9847, 0.9828, 0.9761, 0.9878, 0.9884, 0.9862, 0.9808, 0.9818, 0.9886, 0.9809, 0.9909, 0.9847, 0.992, 0.982, 0.9859, 0.987, 0.9845, 0.9883, 0.9862] },
    MAIN_FVI: { month: [null, 0.9918, 0.9895, 0.989, 0.9874, 0.9897, 0.9919], week: [null, null, null, null, null, 0.9896, 0.9926, 0.9915, 0.9936, 0.9918, 0.9806, 0.9878, 0.9894, 0.9902, 0.992, 0.9889, 0.9861, 0.9869, 0.9875, 0.99, 0.9894, 0.982, 0.9862, 0.9896, 0.9904, 0.9922, 0.9916, 0.9917], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9868, 0.992, 0.9856, 0.9912, 0.9902, 0.9899, 0.9912, 0.9921, 0.9926, 0.9897, 0.9943, 0.9942, 0.9915, 0.9944, 0.9941, 0.9948, 0.9937, 0.9932, 0.9935, 0.9912, 0.9922, 0.9928, 0.9919, 0.9896, 0.9909, 0.9933, 0.9919, null, null, null, null, null, 0.9806, 0.9806, 0.9872, 0.9904, null, 0.9905, null, 0.9902, 0.9868, 0.9898, 0.9905, 0.9913, 0.9893, 0.9887, null, 0.9896, 0.9903, 0.9922, 0.9915, 0.9846, 0.9899, 0.9933, 0.9924, 0.9922, 0.9915, 0.9928, 0.9925, 0.99, 0.9927, 0.9926, 0.9919, 0.9899, 0.9875, 0.9848, 0.9878, 0.988, 0.9903, 0.9867, 0.9886, 0.9826, 0.985, 0.9836, null, 0.9842, 0.9865, 0.9871, 0.9861, 0.9905, 0.9871, 0.987, 0.9865, 0.9855, 0.9839, 0.9921, 0.9905, 0.9906, 0.9905, 0.9896, 0.9897, 0.9897, 0.9897, 0.9903, 0.9903, 0.9912, 0.9867, 0.9899, 0.9892, 0.9893, 0.9893, 0.9901, 0.9906, 0.9901, 0.9434, 0.9887, 0.9898, 0.9749, 0.9838, 0.9881, 0.9895, 0.9909, 0.9883, 0.988, 0.9904, 0.9898, 0.991, 0.9901, 0.9907, 0.9891, 0.9907, 0.9897, 0.9889, 0.9931, 0.9918, 0.9924, 0.9917, 0.9934, 0.9933, 0.9907, 0.9916, 0.9896, 0.9924, 0.9919, 0.992, 0.9921, 0.9918, 0.9916, 0.9922, 0.9914] },
    MAIN_ASSY: { month: [0.9908, 0.9938, 0.9964, 0.9974, 0.9979, 0.9983, 0.9991], week: [null, null, null, 1, 0.9843, 0.9932, 0.9939, 0.9966, 0.9959, 0.9982, 0.9954, 0.9966, 0.9956, 0.9955, 0.9973, 0.9983, 0.998, 0.9978, 0.9981, 0.9985, 0.9974, 0.9974, 0.9985, 0.9984, 0.9981, 0.9981, 0.999, 0.9989], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, null, null, null, null, null, 0.9614, 0.9983, 0.9943, null, 0.9833, 0.9949, 0.9937, 0.9908, 0.9915, 0.9939, 0.9932, 0.9943, 0.9927, 0.9926, 0.9917, 0.9948, 0.9979, 0.9966, 0.9965, 0.9958, 0.9956, 0.9967, 0.9956, 0.9935, 0.9974, 0.9965, 0.9982, 0.9978, 0.9974, 1, 0.9993, null, 0.9979, 0.995, 0.9916, null, null, 0.9971, 0.9971, 0.9909, 1, null, null, 0.9986, null, 1, 0.9965, 0.9907, 0.9972, 0.9953, 0.9938, null, 0.9961, 0.9931, 0.9983, 1, 0.9945, 0.9931, 0.9931, 0.9938, 0.9977, 0.9984, 0.9984, 0.9979, 0.9981, 0.9969, 0.9984, 0.9981, 0.9989, 0.9987, 0.9977, 0.998, 0.9985, 0.9984, 0.9981, 0.9973, 0.9986, 0.9979, 0.9978, null, 0.9986, 0.996, 0.9974, 0.9985, 0.9988, 0.9983, 0.9983, 0.9956, 0.9981, 0.9985, 0.9989, 0.999, 0.9975, 0.9982, 0.998, 0.9989, 0.9989, 0.9989, 0.9991, 0.9988, 0.9989, 0.9988, 0.9948, 0.9965, 0.9968, 0.995, 0.9966, 0.9989, 0.9987, 0.999, 0.9964, 0.9983, 0.9989, 0.9988, 0.999, 0.999, 0.997, 0.9991, 0.9981, 0.9973, 0.9985, 0.9991, 0.9981, 0.9963, 0.9988, 0.9985, 0.9991, 0.9981, 0.9979, 0.9964, 0.9984, 0.9991, 0.9982, 0.9975, 0.9992, 0.9985, 0.9984, 0.9989, 0.9994, 0.9995, 0.9994, 0.9989, 0.9988, 0.9993, 0.9988] },
    MAIN_DRIVING: { month: [0.9928, 0.9958, 0.994, 0.9915, 0.9926, 0.9946, 0.995], week: [null, null, null, 0.9885, 0.9951, 0.9946, 0.996, 0.9945, 0.9965, 0.9935, 0.9945, 0.9935, 0.9948, 0.9947, 0.9917, 0.9917, 0.9905, 0.9889, 0.9925, 0.993, 0.9928, 0.9922, 0.9944, 0.9952, 0.9948, 0.9942, 0.9945, 0.995], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9885, null, null, null, null, null, 0.9915, 0.9932, 0.9926, 0.9979, 1, 0.9959, 0.9977, 0.9944, 0.995, 0.9925, 0.9939, 0.9926, 0.9938, 0.9949, 0.9968, 0.9976, 0.9966, 0.9945, 0.9965, 0.9961, 0.9964, 0.9968, 0.997, 0.9962, null, 0.9895, 0.9895, 0.9939, 0.9959, 0.9959, 0.9963, null, 0.9936, null, null, null, null, 0.9953, 0.9953, 0.9947, 0.992, null, 0.99, 0.9924, 0.9967, 0.9952, 0.9949, 0.9923, 0.997, 0.9955, 0.9941, null, 0.9937, 0.9944, 0.9962, 0.9953, 0.995, 0.9948, 0.9935, 0.9936, 0.9934, 0.9913, 0.9917, 0.99, 0.9907, 0.9913, 0.9913, 0.9911, 0.992, 0.9917, 0.9928, 0.992, 0.9911, 0.9894, 0.9895, 0.9914, 0.9895, 0.9917, 0.9914, null, 0.9909, 0.9773, 0.9927, 0.9925, 0.9908, 0.9903, 0.9911, 0.9924, 0.9925, 0.9924, 0.9937, 0.9954, 0.9945, 0.9917, 0.9931, 0.993, 0.993, 0.993, 0.9925, 0.9934, 0.9922, 0.9925, 0.993, 0.9925, 0.993, 0.9921, 0.9915, 0.9925, 0.9919, 0.9923, 0.9928, 0.993, 0.9936, 0.9938, 0.9958, 0.9955, 0.9949, 0.996, 0.9956, 0.9945, 0.9951, 0.9947, 0.9952, 0.9956, 0.9938, 0.9953, 0.9954, 0.9936, 0.9953, 0.9945, 0.9937, 0.9945, 0.9944, 0.994, 0.9942, 0.994, 0.9931, 0.9943, 0.9947, 0.995, 0.9961, 0.9955, 0.9949, 0.9952, 0.9945] },
    MAIN_TILT: { month: [0.9247, 0.9645, 0.9689, 0.9895, 0.9859, 0.9909, 0.9861], week: [null, null, null, 0.9961, 0.9255, 0.9605, 0.9615, 0.9832, 0.9626, 0.9368, 1, 0.9891, 0.9842, 0.9934, 0.9835, 0.9932, 0.9871, 0.9904, 0.9863, 0.9853, 0.9835, 0.9873, 0.9877, 0.988, 0.9941, 0.9936, 0.9898, 0.9831], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9961, null, null, null, null, null, 0.9867, 1, 0.7064, 0.9342, 1, 0.9121, 0.9786, 0.9887, 0.9335, 0.9783, 0.9709, 0.9612, 0.9896, 0.9512, 0.9463, 0.9652, 0.9551, 0.9832, 0.9786, 0.9829, 0.9625, 0.9611, 0.9687, 0.9222, null, 0.9784, 0.953, 0.9592, 0.9375, 0.9065, 0.9633, 0.8598, null, null, null, null, null, 1, 1, null, 0.9898, null, null, null, 0.9776, 0.9904, 0.9983, 0.9792, 0.9973, 0.9903, 0.9496, null, 0.9976, 0.9821, 0.9956, 0.9901, 0.9927, 0.9983, 0.9976, 0.9988, 0.999, 0.9814, 0.9621, 0.9712, 0.9858, 0.9859, 0.9896, 0.9889, 0.9939, 0.9973, 0.9962, 0.9936, 0.993, 0.9724, 0.9922, 0.9928, 0.9857, 0.9847, 0.9946, null, 0.9883, 0.9864, 0.9971, 0.9883, 0.992, 0.9934, 0.9794, 0.9662, 0.9888, 0.9913, 0.9949, 0.9902, 0.9839, 0.9859, 0.9795, 0.9872, 0.9872, 0.9872, 0.9859, 0.9909, 0.9757, 0.9777, 0.9861, 0.9821, 0.9886, 0.9927, 0.9916, 0.9898, 0.9927, 0.9817, 0.9754, 0.9877, 0.9884, 0.993, 0.9869, 0.9853, 0.9851, 0.9879, 0.985, 0.979, 0.9924, 0.9932, 0.9906, 0.9936, 0.9899, 0.995, 0.9949, 0.9948, 0.9968, 0.994, 0.9942, 0.9942, 0.9877, 0.9965, 0.9951, 0.9967, 0.9863, 0.9908, 0.9907, 0.9852, 0.9893, 0.9778, 0.9911, 0.9896, 0.9738] },
    SUB1_FPCB: { month: [0.9975, 0.9984, 0.9973, 0.9974, 0.9974, 0.9977, 0.9974], week: [null, null, null, 1, 0.9975, 0.9985, 0.9982, 0.9986, 0.9983, 0.9982, 0.9973, 0.9949, 0.9973, 0.9971, 0.9975, 0.9973, 0.9978, 0.9979, 0.9967, 0.9971, 0.9978, 0.9977, 0.9982, 0.9977, 0.9976, 0.9977, 0.9974, 0.9972], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, null, 0.9963, 0.9975, 0.9961, null, null, 1, 0.9985, 0.9983, 0.9988, 0.9979, 0.9987, 0.9987, null, 0.9977, 0.9986, 0.998, 0.9974, 0.9993, 0.9986, 0.9984, 0.9982, 0.9988, 0.9987, 0.9983, 0.9974, 0.9981, 0.9983, 0.9975, 0.9979, 0.9984, 0.9986, null, 0.9986, 0.9975, 0.9964, 0.9972, null, 0.9987, 0.9964, null, 0.9962, null, null, null, 0.9935, null, 0.9971, 0.9985, null, 0.9968, 0.9971, 0.997, null, 0.9972, 0.9966, 0.9974, 0.9978, 0.9967, 0.9968, 0.9974, 0.9975, 0.9975, 0.9975, 0.9971, 0.9972, 0.997, 0.9988, 0.9981, 0.9975, 0.9966, 0.9966, 0.9967, 0.9981, 0.9972, 0.9977, 0.9975, 0.9976, 0.9972, 0.9981, 0.9987, null, 0.9962, 0.9976, 0.9983, 0.999, 0.9982, 0.9964, 0.9963, 0.9963, 0.9965, 0.9962, 0.9985, 0.9968, 0.9975, 0.9981, 0.9977, 0.9963, 0.9965, 0.9964, 0.9973, 0.9971, 0.9973, 0.9978, 0.9986, 0.9984, 0.9977, 0.9973, 0.998, 0.9973, 0.9979, 0.9988, 0.9965, 0.9985, 0.9983, 0.998, 0.998, 0.9986, 0.9979, 0.9982, 0.9973, 0.9984, 0.9973, 0.9976, 0.9971, 0.9973, 0.9979, 0.9973, 0.9982, 0.9972, 0.9976, 0.9979, 0.9978, 0.9975, 0.9968, 0.9984, 0.9978, 0.9976, 0.9962, 0.9972, 0.9979, 0.998, 0.9976, 0.997, 0.9976, 0.9972, 0.9969] },
    SUB1_FVI: { month: [0.9949, 0.9981, 0.9969, 0.9973, 0.9971, 0.9977, 0.9972], week: [null, null, null, 1, 0.9933, 0.9989, 0.998, 0.9982, 0.9984, 0.9986, 0.9966, 0.9916, 0.9971, 0.9967, 0.9975, 0.997, 0.9977, 0.998, 0.9962, 0.9969, 0.9975, 0.9973, 0.9982, 0.998, 0.9973, 0.9977, 0.9971, 0.9972], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, null, 0.9978, 0.9981, 0.9973, null, 0.9812, 0.9924, 0.9987, 0.9987, 0.9994, 0.9988, 0.9988, 0.9988, null, 0.9967, 0.9987, 0.9981, 0.9975, 0.9991, 0.9982, 0.9986, 0.9986, 0.9988, 0.9987, 0.9987, 0.9972, 0.998, 0.999, 0.9979, 0.998, 0.9987, 0.9983, null, 1, 0.996, 0.9977, 0.996, null, 0.9992, 0.9939, null, 0.9943, null, null, null, 0.989, null, 0.997, 0.999, null, 0.9944, 0.9971, 0.9981, null, 0.9965, 0.996, 0.9976, 0.9984, 0.9957, 0.9954, 0.9969, 0.9978, 0.9974, 0.9972, 0.9972, 0.9973, 0.9968, 0.999, 0.9978, 0.9979, 0.9965, 0.9966, 0.9959, 0.998, 0.9961, 0.9977, 0.9975, 0.9976, 0.9975, 0.9982, 0.9978, null, 0.9967, 0.9977, 0.9987, 0.9989, 0.9983, 0.9964, 0.9956, 0.9965, 0.9966, 0.9962, 0.9971, 0.9954, 0.998, 0.998, 0.9979, 0.9958, 0.9955, 0.9959, 0.9973, 0.9968, 0.997, 0.998, 0.9981, 0.9978, 0.9975, 0.9975, 0.9979, 0.9975, 0.9975, 0.9985, 0.995, 0.9984, 0.9987, 0.9979, 0.9982, 0.9983, 0.9979, 0.9982, 0.9973, 0.999, 0.9971, 0.9979, 0.9986, 0.9965, 0.9973, 0.9974, 0.9975, 0.9974, 0.9979, 0.9982, 0.9977, 0.9974, 0.9973, 0.9979, 0.9979, 0.9979, 0.9954, 0.997, 0.9977, 0.9972, 0.9974, 0.9972, 0.9975, 0.9974, 0.9967] },
    SUB2_HOOK: { month: [0.9978, 0.9997, 0.9995, 0.9996, 0.9997, 0.9998, 0.9997], week: [null, null, null, 0.9946, 0.9987, 0.9997, 0.9997, 0.9994, 0.9996, 0.9995, 0.9999, 0.9998, 0.9991, 0.9996, 0.9998, 0.9996, 0.9996, 0.9995, 0.9997, 0.9999, 0.9996, 0.9997, 0.9999, 0.9997, 0.9998, 0.9998, 0.9998, 0.9997], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9946, null, null, null, null, 1, 1, 0.993, 0.9993, 1, null, 1, 1, 1, 0.9995, 0.9994, 0.9995, 0.9998, 1, 0.9996, 1, 0.9989, 1, 0.9999, 0.9994, 0.9999, 0.9997, 0.9997, 0.9994, 0.9996, 0.9994, 0.9996, 0.9993, 0.9997, 0.9995, 0.9995, 0.9994, null, null, 0.9997, 1, 1, 1, null, 0.9999, null, null, null, null, null, 0.9998, null, 0.9995, 0.9985, 0.998, 0.9995, 0.9993, 0.9998, null, 0.9994, 0.9992, 0.9997, 0.9996, 0.9996, 0.9998, 0.9997, 0.9999, 0.9999, 0.9998, 0.9999, 0.9998, 0.9996, 0.9994, 0.9998, 0.9992, 0.9994, 0.9996, 0.9998, 0.9998, 0.9994, 0.9998, 0.9994, 0.9998, 1, 0.9995, 0.9993, null, 0.9999, 0.9992, 0.9994, 1, 0.9992, 0.9995, 0.9998, 0.9993, 1, 0.9999, 0.9997, 0.9996, 0.9999, 0.9999, 0.9999, 1, 1, 1, 0.9994, 0.9998, 0.9999, 0.9998, 0.9999, 0.9995, 0.9984, 0.9994, 0.9999, 0.9996, 1, 0.9992, 1, 1, 1, 0.9999, 1, 1, 0.9995, 0.9997, 0.9995, 1, 0.9999, 0.9995, 0.9999, 0.9999, 0.9997, 0.9999, 0.9999, 0.9998, 0.9997, 0.9997, 0.9999, 0.9995, 0.9999, 1, 0.9998, 0.9994, 1, 0.9993, 1, 1, 0.9999, 0.9998, 0.9995, 0.9999, 0.9994] },
    SUB2_OVEN: { month: [1, 1.0, 1, 1, 1, 1, 1], week: [null, null, null, 1, 1, 1.0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 0.9998, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, null, 1, 1, 1, 1, null, 1, null, null, null, null, null, 1, null, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1] },
    SUB2_INDEX: { month: [0.9835, 0.9965, 0.9949, 0.9936, 0.9948, 0.995, 0.9977], week: [null, null, null, 0.9964, 0.9839, 0.9949, 0.9961, 0.9953, 0.9984, 0.997, 0.9941, 0.9868, 0.9962, 0.9884, 0.9948, 0.9934, 0.9958, 0.9929, 0.9944, 0.9963, 0.9947, 0.9961, 0.995, 0.9929, 0.9943, 0.9971, 0.9975, 0.998], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9964, null, null, null, null, 0.9583, 0.9816, 0.9857, 0.9819, 0.9971, null, 0.9986, 0.9975, 0.9982, 0.9949, 0.9961, 0.9938, 0.9918, 0.9924, 0.9959, 0.9917, 0.9954, 0.9989, 0.9985, 0.9953, 0.9983, 0.999, 0.9982, 0.9969, 0.9987, 0.9994, 0.9984, 0.998, 0.9977, 0.9965, 0.9946, 0.9979, null, null, 0.9957, 0.9948, 0.9952, 0.9883, null, 0.9963, null, null, null, null, null, 0.9868, null, 0.998, 0.9993, 0.9975, 0.9913, 0.9948, 0.9964, null, 0.9927, 0.9883, 0.9737, 0.9842, 0.9918, 0.9928, 0.9955, 0.9951, 0.9948, 0.9931, 0.9964, 0.9947, 0.9928, 0.9965, 0.9906, 0.9951, 0.9894, 0.9937, 0.9952, 0.9945, 0.9953, 0.9966, 0.9948, 0.9953, 0.9954, 0.9954, 0.9976, null, 0.9956, 0.9972, 0.9967, 0.9973, 0.9776, 0.9926, 0.99, 0.9934, 0.994, 0.9963, 0.9961, 0.9985, 0.9944, 0.9956, 0.9956, 0.9985, 0.9976, 0.9946, 0.9976, 0.9957, 0.9955, 0.9952, 0.9929, 0.9975, 0.9911, 0.9951, 0.997, 0.9967, 0.9961, 0.995, 0.9967, 0.9955, 0.9958, 0.9951, 0.9927, 0.9967, 0.9943, 0.9922, 0.993, 0.9949, 0.9933, 0.9931, 0.9908, 0.9948, 0.9947, 0.9891, 0.9949, 0.9959, 0.9966, 0.998, 0.9976, 0.9981, 0.9951, 0.9976, 0.996, 0.9967, 0.9981, 0.9968, 0.9991, 0.9982, 0.9959, 0.9982, 0.9984, 0.9976, 0.9977] },
  },
  'SO1C2EDL': {
    RTY_TTL: { month: [0.9829, 0.9437, 0.9112, null, 0.9737, 0.8464, null], week: [null, null, null, null, 0.9829, 0.9463, 0.9395, null, null, 0.9754, 0.9039, 0.9657, null, null, null, null, null, null, null, 0.9727, 0.9661, 0.9766, 0.8464, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9968, 0.9787, 0.984, 0.9877, null, 0.9543, 0.9771, 0.9787, 0.9731, 0.9136, 0.9963, 0.9873, 0.9092, 0.9439, 0.9739, 0.9717, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9881, 0.9742, 0.9482, 0.9853, 0.9515, 0.9346, 0.914, 0.9557, 0.9657, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9687, 0.9719, 0.9751, 0.9752, 0.9661, null, null, null, null, null, null, null, 0.9787, 0.9762, 0.9834, 0.9679, 0.8791, 0.9628, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_MAIN: { month: [null, 0.9611, 0.9434, null, 0.9737, 0.8464, null], week: [null, null, null, null, null, 0.9629, 0.9588, null, null, null, 0.9442, 0.9657, null, null, null, null, null, null, null, 0.9727, 0.9661, 0.9766, 0.8464, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9706, 0.984, 0.9787, 0.9731, 0.9136, null, null, 0.9193, 0.9703, 0.9739, 0.9717, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9854, 0.9622, 0.9212, 0.9557, 0.9657, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9687, 0.9719, 0.9751, 0.9752, 0.9661, null, null, null, null, null, null, null, 0.9787, 0.9762, 0.9834, 0.9679, 0.8791, 0.9628, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB1: { month: [0.9956, 0.9914, 0.9928, null, null, null, null], week: [null, null, null, null, 0.9956, 0.9943, 0.9874, null, null, 0.9884, 0.9937, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9942, 0.9974, 0.9953, null, 0.9957, 0.993, null, null, null, null, 0.9942, 0.9962, 0.9806, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9884, 0.9924, 0.9936, 0.9947, 0.9958, 0.9921, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB2: { month: [0.9872, 0.9904, 0.9729, null, null, null, null], week: [null, null, null, null, 0.9872, 0.9885, 0.9923, null, null, 0.9869, 0.9633, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9968, 0.9845, 0.9865, 0.9923, null, 0.9875, null, null, null, null, 0.9963, 0.9931, 0.9928, 0.9919, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9881, 0.9856, 0.9554, 0.9916, 0.9707, 0.9755, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_FVI: { month: [null, 0.9952, 0.992, null, 0.9935, 0.9885, null], week: [null, null, null, null, null, 0.9948, 0.9957, null, null, null, 0.9905, 0.9949, null, null, null, null, null, null, null, 0.9925, 0.9882, 0.9959, 0.9885, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9945, 0.9922, 0.9966, 0.9957, null, null, 0.9981, 0.9937, 0.9952, 0.9958, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.992, 0.989, null, 0.9949, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9939, 0.9927, 0.9919, 0.9916, 0.9882, null, null, null, null, null, null, null, 0.9961, 0.9959, 0.9941, 0.9976, null, 0.9885, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_ASSY: { month: [null, 0.9798, 0.9704, null, 0.9933, 0.8791, null], week: [null, null, null, null, null, 0.9766, 0.9839, null, null, null, 0.9704, null, null, null, null, null, null, null, null, 0.9936, 0.9898, 0.9939, 0.8791, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9733, 0.9933, 0.9901, 0.9861, 0.9404, null, null, 0.9722, 0.9865, 0.9848, 0.9919, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9854, 0.9794, 0.9462, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9895, 0.9939, 0.9954, 0.9958, 0.9898, null, null, null, null, null, null, null, 0.9979, 0.9927, 0.9985, 0.9863, 0.8791, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_DRIVING: { month: [null, 0.9963, 0.9807, null, 0.9869, 0.974, null], week: [null, null, null, null, null, 0.9966, 0.9959, null, null, null, 0.9831, 0.9711, null, null, null, null, null, null, null, 0.9868, 0.9879, 0.9868, 0.974, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9982, 0.9979, 0.9972, 0.9949, 0.9949, null, null, 1, 0.9944, 0.9942, 0.995, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9908, 0.986, 0.9557, 0.9711, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9857, 0.9854, 0.9876, 0.9885, 0.9879, null, null, null, null, null, null, null, 0.9849, 0.9874, 0.991, 0.9837, null, 0.974, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_TILT: { month: [null, 0.9892, 0.9993, null, 0.9997, 1, null], week: [null, null, null, null, null, 0.9945, 0.9827, null, null, null, 0.9993, 0.9995, null, null, null, null, null, null, null, 0.9995, 0.9997, 0.9998, 1, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.999, 0.9983, 0.999, 0.9952, 0.9807, null, null, 0.9474, 0.9954, 0.9995, 0.9887, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9994, 0.9984, 1, 0.9995, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9993, 0.9996, 1, 0.9992, 0.9997, null, null, null, null, null, null, null, 0.9997, 1, 0.9997, 1, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FPCB: { month: [0.9977, 0.9971, 0.9963, null, null, null, null], week: [null, null, null, null, 0.9977, 0.9968, 0.9981, null, null, 0.995, 0.9966, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9971, 0.9985, 0.9973, null, 0.9975, 0.9963, null, null, null, null, 0.9965, 0.9981, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.995, 0.9958, 0.9965, 0.9975, 0.9976, 0.9956, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FVI: { month: [0.998, 0.9942, 0.9965, null, null, null, null], week: [null, null, null, null, 0.998, 0.9975, 0.9894, null, null, 0.9934, 0.9971, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.997, 0.9989, 0.998, null, 0.9982, 0.9967, null, null, null, null, 0.9976, 0.9981, 0.9806, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9934, 0.9965, 0.9972, 0.9973, 0.9982, 0.9965, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_HOOK: { month: [0.9997, 1, 0.999, null, null, null, null], week: [null, null, null, null, 0.9997, 1, 1, null, null, 0.9997, 0.9985, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9997, 0.9998, 0.9996, null, 1, null, null, null, null, null, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9996, 0.9998, 0.9956, null, 1, 0.9998, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_OVEN: { month: [0.9999, 1, 1, null, null, null, null], week: [null, null, null, null, 0.9999, 1, 1, null, null, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9999, 0.9999, null, 1, null, null, null, null, null, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, null, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_INDEX: { month: [0.9964, 0.9973, 0.9872, null, null, null, null], week: [null, null, null, null, 0.9964, 0.9967, 0.9978, null, null, 0.9953, 0.9818, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9941, 0.9963, 0.9986, null, 0.9957, null, null, null, null, null, 0.9978, 0.9972, 0.9985, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9951, 0.9954, 0.9881, null, 0.9741, 0.983, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
  },
  'SO1C2EF': {
    RTY_TTL: { month: [null, null, 0.9821, 0.9287, 0.9233, 0.9533, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9237, 0.9242, 0.9213, 0.9473, 0.9246, 0.9309, 0.8999, 0.9279, 0.9341, 1, 0.9469, 0.9813, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9841, 0.98, 0.9595, 0.9135, 0.9106, 0.9389, 0.936, 0.8936, 0.9259, 0.9215, 0.9318, 0.926, 0.9269, 0.9587, 0.9388, 0.8323, 0.9069, 0.9329, 0.9422, 0.9384, 0.9598, 0.9528, 0.935, 0.9381, 0.95, 0.9554, 0.9596, null, 0.9414, 0.9159, 0.9297, 0.8982, 0.9381, 0.9142, 0.9284, 0.9305, 0.928, 0.9406, 0.9326, 0.9421, 0.9347, 0.9332, 0.939, 0.7962, 0.9717, 0.9855, 0.9862, 0.9701, 0.9353, 0.9094, 0.9406, 0.9329, 0.9244, 0.9377, 0.9372, 0.9292, 0.9805, 0.9815, null, 1, null, null, null, null, null, 0.9671, 0.9849, null, 0.9663, 0.9738, 0.9638, 0.9813, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_MAIN: { month: [null, null, null, 0.9501, 0.9486, 0.9755, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9437, 0.9453, 0.9442, 0.9687, 0.9469, 0.9528, 0.9241, 0.9569, 0.9613, null, 0.9735, 0.9813, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9793, 0.9361, 0.9327, 0.9573, 0.9515, 0.9152, 0.9474, 0.9402, 0.9552, 0.9479, 0.9488, 0.9626, 0.965, 0.8555, 0.9273, 0.9555, 0.9641, 0.9628, 0.9813, 0.9738, 0.9577, 0.9565, 0.9779, 0.9751, 0.9711, null, 0.9612, 0.9372, 0.9467, 0.9284, 0.9612, 0.9387, 0.9508, 0.9498, 0.9482, 0.9666, 0.9527, 0.9629, 0.9521, 0.9528, 0.9626, 0.8419, null, null, null, null, 0.9578, 0.9522, 0.9658, 0.9582, 0.9504, 0.9587, 0.9657, 0.967, null, null, null, null, null, null, null, null, null, null, null, null, 0.9663, 0.9738, 0.9805, 0.9813, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB1: { month: [null, null, 0.9948, 0.9928, 0.9939, 0.9961, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9939, 0.993, 0.9921, 0.9927, 0.9909, 0.9944, 0.9951, 0.9929, 0.9957, null, 0.9961, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9942, 0.9955, 0.9951, 0.9915, 0.9922, 0.9934, 0.9957, 0.9934, 0.9927, 0.9918, 0.9933, 0.9911, 0.9927, 0.996, 0.9924, 0.9922, 0.9931, 0.9931, 0.9916, 0.9909, 0.9916, 0.9915, 0.9927, 0.9927, 0.9933, 0.9932, null, null, 0.9925, 0.9928, 0.9939, 0.9815, 0.9941, 0.9956, 0.9951, 0.9949, 0.9943, 0.9933, 0.9939, 0.9939, 0.9957, 0.9934, 0.9943, 0.996, 0.995, 0.9962, null, 0.9938, 0.9947, 0.9931, 0.9929, 0.9923, 0.9904, 0.9965, 0.9964, 0.9951, 0.9949, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9961, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB2: { month: [null, null, 0.9872, 0.9845, 0.9792, 0.9811, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9848, 0.9846, 0.9835, 0.9852, 0.9853, 0.9825, 0.9786, 0.9767, 0.976, 1, 0.9764, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9899, 0.9845, 0.9846, 0.9843, 0.9839, 0.9874, 0.988, 0.9829, 0.9845, 0.9883, 0.982, 0.9856, 0.9842, null, 0.9803, 0.9806, 0.9848, 0.9831, 0.9855, 0.9836, 0.9864, 0.9869, 0.9834, 0.988, 0.978, 0.9866, 0.9882, null, 0.9868, 0.9844, 0.988, 0.9857, 0.9818, 0.9782, 0.9812, 0.9847, 0.9844, 0.9796, 0.9849, 0.9844, 0.986, 0.9859, 0.9811, 0.9495, 0.9766, 0.9893, 0.9862, 0.9761, 0.9817, 0.9617, 0.9809, 0.9812, 0.982, 0.9815, 0.974, 0.9657, 0.9855, 0.9815, null, 1, null, null, null, null, null, 0.9671, 0.9849, null, null, null, 0.9867, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_FVI: { month: [null, null, null, 0.9942, 0.9927, 0.9954, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9942, 0.9945, 0.9939, 0.9942, 0.9941, 0.9938, 0.9904, 0.9915, 0.9944, null, 0.9954, 0.9952, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.994, 0.996, 0.9933, 0.9935, 0.9948, 0.9948, 0.994, 0.9944, 0.994, 0.9947, 0.9946, 0.9946, 0.9921, 0.9928, 0.9955, 0.9944, 0.9926, 0.9952, 0.9945, 0.9941, 0.9953, 0.9931, 0.994, 0.9944, null, 0.9943, 0.9933, 0.9952, 0.9933, 0.9945, 0.9948, 0.9925, 0.9933, 0.9941, 0.9935, 0.994, 0.9943, 0.9953, 0.9921, 0.9929, 0.9813, null, null, null, null, 0.994, 0.9843, 0.9926, 0.9939, 0.9927, 0.9935, 0.9954, 0.9944, null, null, null, null, null, null, null, null, null, null, null, null, 0.9958, 0.9949, 0.9957, 0.9952, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_ASSY: { month: [null, null, null, 0.9871, 0.9887, 0.9915, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9824, 0.9842, 0.9891, 0.9942, 0.9832, 0.9893, 0.9868, 0.9894, 0.9921, null, 0.9892, 0.9984, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9793, 0.9686, 0.9833, 0.9943, 0.9864, 0.9759, 0.9894, 0.9777, 0.9863, 0.9862, 0.9822, 0.9913, 0.9957, 0.9806, 0.9904, 0.9899, 0.9872, 0.9843, 0.9958, 0.9947, 0.9874, 0.9913, 0.999, 0.9958, 0.997, null, 0.9836, 0.9704, 0.9945, 0.979, 0.9887, 0.9849, 0.9889, 0.9912, 0.9865, 0.9947, 0.9858, 0.9932, 0.9823, 0.9863, 0.9919, null, null, null, null, null, 0.9907, 0.9931, 0.9944, 0.987, 0.9819, 0.9909, 0.9933, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9844, 0.9901, 0.9931, 0.9984, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_DRIVING: { month: [null, null, null, 0.9765, 0.9677, 0.9887, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9695, 0.9705, 0.9825, 0.9832, 0.9703, 0.9701, 0.9481, 0.976, 0.9747, null, 0.9889, 0.9879, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.973, 0.9566, 0.972, 0.9766, 0.9634, 0.9654, 0.9685, 0.9742, 0.9674, 0.9765, 0.978, 0.9771, 0.9685, 0.985, 0.9851, 0.9838, 0.9867, 0.9911, 0.9859, 0.9777, 0.9808, 0.9877, 0.9865, 0.9805, null, 0.9845, 0.9735, 0.9594, 0.9554, 0.9785, 0.9595, 0.97, 0.9658, 0.9678, 0.9796, 0.9729, 0.9751, 0.9741, 0.9744, 0.978, 0.8658, null, null, null, null, 0.9731, 0.9744, 0.9796, 0.9776, 0.9753, 0.9738, 0.9771, 0.9732, null, null, null, null, null, null, null, null, null, null, null, null, 0.9862, 0.9888, 0.9918, 0.9879, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_TILT: { month: [null, null, null, 0.9915, 0.9989, 0.9997, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9966, 0.9952, 0.9776, 0.9967, 0.9985, 0.999, 0.9973, 0.9994, 0.9996, null, 0.9997, 0.9997, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9993, 0.9956, 0.9972, 0.9942, 0.9785, 0.9971, 0.9989, 0.9997, 0.9996, 0.9945, 0.9982, 0.9973, 0.9079, 0.9575, 0.9843, 0.9983, 0.9988, 0.9991, 0.9984, 0.998, 0.9885, 0.9979, 0.9986, 0.9989, null, 0.9983, 0.9988, 0.9971, 0.9993, 0.999, 0.9986, 0.9986, 0.9988, 0.9991, 0.9985, 0.9992, 1, 0.9997, 0.9992, 0.9995, 0.9909, null, null, null, null, 0.9995, 0.9997, 0.9988, 0.9992, 0.9997, 1, 0.9995, 0.9993, null, null, null, null, null, null, null, null, null, null, null, null, 0.9996, 0.9998, 0.9998, 0.9997, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FPCB: { month: [null, null, 0.9974, 0.9967, 0.997, 0.9976, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9971, 0.9968, 0.9964, 0.9967, 0.9959, 0.9971, 0.9976, 0.9968, 0.9973, null, 0.9976, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9971, 0.9977, 0.9976, 0.9963, 0.9967, 0.9968, 0.9977, 0.9967, 0.9969, 0.9964, 0.9968, 0.9962, 0.9971, 0.9976, 0.9965, 0.9962, 0.9967, 0.9965, 0.9962, 0.996, 0.9965, 0.9965, 0.9967, 0.9965, 0.9968, 0.9967, null, null, 0.9963, 0.9962, 0.9969, 0.993, 0.997, 0.9976, 0.9974, 0.9973, 0.997, 0.9967, 0.9969, 0.9967, 0.9976, 0.997, 0.9972, 0.9983, 0.9973, 0.9982, null, 0.9971, 0.9976, 0.9968, 0.9968, 0.9967, 0.9959, 0.9976, 0.9979, 0.9971, 0.9968, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9976, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FVI: { month: [null, null, 0.9974, 0.9962, 0.9969, 0.9985, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9968, 0.9962, 0.9958, 0.996, 0.995, 0.9973, 0.9975, 0.996, 0.9984, null, 0.9985, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.997, 0.9978, 0.9975, 0.9952, 0.9956, 0.9966, 0.9981, 0.9967, 0.9958, 0.9954, 0.9965, 0.9949, 0.9956, 0.9984, 0.9959, 0.996, 0.9964, 0.9966, 0.9954, 0.9949, 0.9951, 0.995, 0.996, 0.9962, 0.9965, 0.9965, null, null, 0.9962, 0.9965, 0.997, 0.9884, 0.9971, 0.998, 0.9977, 0.9976, 0.9972, 0.9966, 0.9971, 0.9972, 0.9981, 0.9964, 0.9971, 0.9977, 0.9977, 0.9979, null, 0.9967, 0.9971, 0.9963, 0.996, 0.9955, 0.9945, 0.9989, 0.9985, 0.998, 0.9981, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9985, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_HOOK: { month: [null, null, 1, 0.9997, 0.9999, 1, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9998, 0.9997, 0.9996, 0.9999, 0.9999, 0.9999, 0.9998, 0.9999, 1.0, 1, 1, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 0.9994, 0.9999, 0.9999, 0.9997, null, 0.9999, 0.9998, 0.9998, 0.9995, 0.9998, 0.9993, null, 0.9995, 0.9994, 0.9997, 0.9995, 1, 0.9996, 0.9997, 0.9999, 0.9999, 0.9998, 0.9997, 0.9999, 1, null, 0.9998, 1, 0.9998, 0.9999, 1, 0.9999, 0.9999, 1, 1, 1, 1, 0.9994, 1, 0.9997, 0.9999, 0.9999, 1, 0.9998, 0.9996, 0.9998, 1, 1, 0.9999, 1, 1, 0.9999, 1, 1, 1, 1, null, 1, null, null, null, null, null, 1, 1, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_OVEN: { month: [null, null, 1, 1, 1, 1, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, null, null, null, null, null, 1, 1, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_INDEX: { month: [null, null, 0.996, 0.9948, 0.9927, 0.9918, null], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9956, 0.9948, 0.9948, 0.9942, 0.9944, 0.9927, 0.995, 0.9898, 0.9923, 1, 0.9891, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9959, 0.996, 0.9966, 0.9956, 0.9944, 0.9948, null, 0.9897, 0.9972, 0.9962, 0.9959, 0.9959, 0.9942, null, 0.9951, 0.9934, 0.9944, 0.9949, 0.9962, 0.9941, 0.9954, 0.9961, 0.9932, 0.9977, 0.9885, 0.9954, 0.9942, null, 0.9937, 0.9946, 0.9964, 0.9934, 0.9937, 0.9871, 0.9939, 0.9939, 0.9929, 0.9892, 0.994, 0.9982, 0.9953, 0.9944, 0.9905, 0.9956, 0.9957, 0.9976, 0.9961, 0.9826, 0.9877, 0.9912, 0.9933, 0.9914, 0.9926, 0.9941, 0.9917, 0.9909, 0.9923, 0.9922, null, 1, null, null, null, null, null, 0.977, 0.993, null, null, null, 0.9974, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
  },
  'SO1C30 S25': {
    RTY_TTL: { month: [0.9662, 0.9537, 0.9504, 0.9264, 0.941, null, null], week: [0.9685, 0.9702, 0.965, 0.9675, 0.9549, null, null, null, 0.9393, 0.9595, 0.9573, 0.956, 0.9422, 0.8697, 0.9229, 0.9199, 0.939, 0.9363, 0.9372, null, null, null, null, null, null, null, null, null], day: [0.9662, 0.9709, 0.9751, 0.9737, 0.9701, 0.9608, 0.9704, 0.9712, 0.9683, 0.9708, 0.9513, 0.969, 0.9601, 0.9703, null, 0.9643, 0.9728, 0.9728, 0.9654, 0.967, 0.963, null, 0.9624, 0.9551, 0.9887, 0.9675, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9537, 0.9511, 0.9508, 0.9463, 0.9601, 0.961, 0.9649, 0.967, 0.9669, 0.9588, 0.951, 0.9579, 0.9635, 0.9608, 0.96, 0.9486, 0.9524, 0.9556, 0.9641, 0.9548, 0.9594, 0.9572, 0.9507, 0.9576, 0.955, 0.9424, 0.9234, 0.9248, null, 0.871, null, null, null, null, null, 0.9762, 0.9969, 0.8891, 0.8851, 0.9463, 0.9426, 0.9307, 0.9486, 0.9346, 0.9327, 0.8936, 0.892, 0.9245, 0.9239, 0.9611, 0.9456, 0.961, 0.9206, 0.9587, 0.9481, 0.9552, 0.9549, 0.9434, 0.9736, 0.9785, 0.9685, 1, null, null, null, 0.9198, 0.9545, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_MAIN: { month: [0.9856, null, 0.9756, 0.9686, 0.941, null, null], week: [0.9871, 0.9878, 0.9873, 0.9835, 0.9821, null, null, null, 0.9793, 0.9843, 0.9792, 0.9774, 0.9736, 0.8894, 0.9533, 0.9706, 0.9778, 0.9738, 0.9372, null, null, null, null, null, null, null, null, null], day: [0.9859, 0.9883, 0.9905, 0.989, 0.9902, 0.9769, 0.9909, 0.9891, 0.9891, 0.9892, 0.9858, 0.9885, 0.9841, 0.9872, null, 0.9845, 0.9881, 0.9859, 0.9808, 0.9827, 0.9792, null, 0.9858, 0.9864, 0.9887, 0.9675, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9793, 0.9833, 0.9748, 0.9837, 0.9857, 0.9888, 0.9855, 0.9883, 0.9809, 0.9693, 0.9816, 0.9856, 0.9828, 0.9835, 0.9708, 0.9778, 0.9754, 0.9791, 0.9785, 0.9825, 0.9775, 0.9821, 0.9808, 0.9791, 0.9784, 0.9539, 0.9674, null, 0.8894, null, null, null, null, null, null, null, 0.9085, 0.9047, 0.9741, 0.9833, 0.9827, 0.9705, 0.9686, 0.9697, 0.9591, 0.9725, 0.9649, 0.9763, 0.9832, 0.9708, 0.9793, 0.97, 0.9788, 0.9798, 0.9872, 0.9911, 0.9811, 0.9736, 0.9785, 0.9685, 1, null, null, null, 0.9198, 0.9545, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB1: { month: [0.9936, null, 0.992, 0.9881, null, null, null], week: [0.9932, 0.9944, 0.993, 0.9942, 0.9921, null, null, null, 0.993, 0.9938, 0.9928, 0.9939, 0.9889, 0.9778, 0.9933, 0.9904, 0.987, 0.9615, null, null, null, null, null, null, null, null, null, null], day: [0.9927, 0.9937, 0.995, 0.9946, 0.9928, 0.9948, 0.9944, 0.9947, 0.9928, 0.9939, 0.9933, 0.9922, 0.9919, 0.9939, null, 0.9934, 0.9946, 0.9948, 0.9936, 0.9936, 0.9952, null, 0.9922, 0.9921, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.993, 0.9938, 0.9927, 0.9963, 0.9932, 0.9929, 0.9934, 0.9941, 0.9931, 0.9957, 0.9921, 0.9914, 0.9925, 0.9918, 0.9926, 0.9947, 0.9924, 0.9968, 0.9923, 0.9931, 0.9957, 0.9917, 0.9914, 0.9919, 0.9865, 0.9857, 0.9862, null, 0.9793, null, null, null, null, null, 0.9762, 0.9976, 0.9951, 0.9945, 0.9936, 0.9855, 0.9934, null, 0.9928, 0.9879, 0.99, 0.9928, 0.9888, 0.9898, null, 0.9962, 0.9942, 0.9706, null, null, null, null, 0.9615, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB2: { month: [0.9866, 0.9537, 0.982, 0.968, null, null, null], week: [0.9879, 0.9878, 0.9843, 0.9895, 0.98, null, null, null, 0.9659, 0.9809, 0.9847, 0.9841, 0.9787, 1, 0.9748, 0.957, 0.973, null, null, null, null, null, null, null, null, null, null, null], day: [0.9872, 0.9886, 0.9895, 0.9898, 0.9867, 0.9886, 0.9848, 0.9871, 0.986, 0.9874, 0.9716, 0.988, 0.9836, 0.989, null, 0.986, 0.9899, 0.9919, 0.9907, 0.9904, 0.9881, null, 0.984, 0.9759, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9537, 0.9781, 0.973, 0.9779, 0.9796, 0.9816, 0.9828, 0.9877, 0.9841, 0.9842, 0.9853, 0.9835, 0.986, 0.985, 0.9842, 0.9845, 0.9792, 0.9872, 0.9878, 0.9834, 0.9833, 0.9835, 0.9761, 0.9849, 0.9834, 0.9765, 0.982, 0.9693, null, null, null, null, null, null, null, 1, 0.9993, 0.9835, 0.9838, 0.9777, 0.9727, 0.9534, 0.9775, 0.9719, 0.9737, 0.9411, 0.9238, 0.9689, 0.9561, 0.9775, 0.9778, 0.9871, 0.9778, 0.9794, 0.9676, 0.9676, 0.9635, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_FVI: { month: [0.9927, null, 0.9926, 0.9935, 0.9511, null, null], week: [0.9944, 0.9945, 0.9935, 0.9906, 0.9911, null, null, null, 0.9881, 0.9942, 0.9925, 0.9913, 0.9932, 0.9913, 0.9943, 0.9929, 0.9931, 0.9905, 0.9372, null, null, null, null, null, null, null, null, null], day: [0.9947, 0.9941, 0.9942, 0.9946, 0.9951, 0.9941, 0.995, 0.9942, 0.9937, 0.9954, 0.9932, 0.9937, 0.992, 0.9933, null, 0.9926, 0.994, 0.9912, 0.9882, 0.9887, 0.9888, null, 0.992, 0.9933, 0.9918, 0.9873, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9881, 0.9927, 0.9957, 0.9946, 0.9945, 0.9936, 0.9938, 0.9947, 0.9908, 0.992, 0.993, 0.9928, 0.995, 0.9915, 0.9874, 0.9924, 0.9914, 0.9913, 0.9933, 0.9923, 0.9912, 0.9931, 0.9939, 0.9924, 0.9934, 0.9932, 0.9932, null, 0.9913, null, null, null, null, null, null, null, null, 0.9935, 0.9941, 0.9942, 0.995, 0.9946, 0.9951, 0.9935, 0.994, 0.9925, 0.992, 0.991, 0.9921, 0.9926, 0.9933, 0.9922, 0.9936, 0.9924, 0.9944, null, 0.9945, 0.9925, 0.9959, 0.9788, null, null, null, null, 0.9198, 0.9545, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_ASSY: { month: [0.9977, null, 0.9938, 0.9908, null, null, null], week: [0.9966, 0.9986, 0.9979, 0.9976, 0.9968, null, null, null, 0.9951, 0.995, 0.9913, 0.9947, 0.9967, 0.9722, 0.9819, 0.9943, 0.9943, 0.9954, null, null, null, null, null, null, null, null, null, null], day: [0.9957, 0.9975, 0.9985, 0.9988, 0.9978, 0.9983, 0.9991, 0.9988, 0.9984, 0.9978, 0.9981, 0.9984, 0.9965, 0.9984, null, 0.9959, 0.9979, 0.9984, 0.9966, 0.9983, 0.9986, null, 0.9989, 0.9985, 0.9995, 0.9902, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9951, 0.9959, 0.9898, 0.9923, 0.9941, 0.9984, 0.9959, 0.9986, 0.9945, 0.983, 0.992, 0.9951, 0.9918, null, 0.9899, 0.9928, 0.995, 0.9936, 0.995, 0.9985, 0.9982, 0.9979, 0.9988, 0.9952, 0.997, 0.9968, 0.9947, null, 0.9722, null, null, null, null, null, null, null, 0.9146, 0.99, 0.9979, 0.9966, 0.9975, 0.9946, 0.9946, 0.9961, 0.9885, 0.9987, 0.9901, 0.9954, 0.9963, 0.9907, 0.9955, 0.9893, 0.9951, 0.997, 0.9982, null, 0.9962, 0.9946, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_DRIVING: { month: [0.996, null, 0.992, 0.9849, 0.9895, null, null], week: [0.997, 0.9954, 0.9964, 0.9964, 0.9952, null, null, null, 0.9962, 0.9953, 0.9964, 0.9921, 0.9881, 0.9618, 0.9767, 0.9836, 0.9909, 0.99, null, null, null, null, null, null, null, null, null, null], day: [0.9964, 0.9976, 0.9979, 0.9971, 0.9978, 0.9856, 0.9972, 0.9967, 0.9972, 0.9964, 0.9951, 0.9969, 0.9963, 0.9964, null, 0.9967, 0.9968, 0.9968, 0.9968, 0.9967, 0.9946, null, 0.996, 0.9961, 0.9977, 0.9911, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9962, 0.9955, 0.9895, 0.9969, 0.9972, 0.9969, 0.9959, 0.9951, 0.9968, 0.9969, 0.9974, 0.9981, 0.9964, 0.9927, 0.9941, 0.9936, 0.9892, 0.994, 0.9905, 0.9926, 0.9907, 0.9925, 0.9895, 0.9929, 0.9902, 0.9813, 0.9823, null, 0.9618, null, null, null, null, null, null, null, 0.9933, 0.9198, 0.983, 0.9929, 0.9903, 0.9811, 0.9797, 0.9799, 0.9776, 0.9814, 0.9823, 0.9899, 0.9947, 0.9872, 0.9913, 0.9883, 0.9904, 0.9903, 0.9946, 0.9945, 0.9912, 0.9911, 0.9882, 0.9895, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_TILT: { month: [0.9991, null, 0.997, 0.9991, 1, null, null], week: [0.9991, 0.9992, 0.9994, 0.9988, 0.9989, null, null, null, 0.9997, 0.9997, 0.9988, 0.9991, 0.9953, 0.9596, 0.9997, 0.9996, 0.9993, 0.9977, 1, null, null, null, null, null, null, null, null, null], day: [0.9991, 0.9991, 0.9997, 0.9985, 0.9995, 0.9988, 0.9995, 0.9994, 0.9998, 0.9996, 0.9992, 0.9995, 0.9991, 0.9991, null, 0.9992, 0.9992, 0.9994, 0.9989, 0.9989, 0.9971, null, 0.999, 0.9985, 0.9996, 0.9985, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9997, 0.9991, 0.9997, 0.9998, 0.9999, 0.9998, 0.9998, 0.9998, 0.9986, 0.9971, 0.9991, 0.9996, 0.9995, 0.9992, 0.9991, 0.9989, 0.9996, 1, 0.9996, 0.9991, 0.9972, 0.9985, 0.9985, 0.9984, 0.9976, 0.9819, 0.9967, null, 0.9596, null, null, null, null, null, null, null, 1, 1, 0.999, 0.9995, 0.9997, 1, 0.9989, 1, 0.9985, 0.9996, 1, 0.9998, 1, 1, 0.9991, 1, 0.9995, 1, 1, 0.9966, 0.9991, 0.9951, 0.9942, 1, 1, null, null, null, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FPCB: { month: [0.9975, null, 0.997, 0.9966, null, null, null], week: [0.997, 0.9977, 0.9973, 0.9976, 0.9981, null, null, null, 0.9977, 0.9975, 0.9971, 0.9974, 0.9967, 0.9937, 0.997, 0.9973, 0.9949, null, null, null, null, null, null, null, null, null, null, null], day: [0.997, 0.997, 0.9981, 0.9976, 0.997, 0.998, 0.9977, 0.9979, 0.9975, 0.9976, 0.9973, 0.997, 0.9973, 0.9973, null, 0.9975, 0.998, 0.9974, 0.9975, 0.9972, 0.9977, null, 0.9963, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9977, 0.9977, 0.9971, 0.998, 0.9973, 0.9972, 0.9972, 0.9977, 0.9975, 0.9975, 0.9968, 0.9968, 0.997, 0.997, 0.9972, 0.9973, 0.997, 0.9979, 0.997, 0.9972, 0.9985, 0.9974, 0.9976, 0.997, 0.9966, 0.9958, 0.9958, null, 0.9921, null, null, null, null, null, 0.9953, 0.9984, 0.9973, 0.9971, 0.9967, 0.9957, 0.9967, null, 0.998, 0.9977, 0.9966, 0.9981, 0.9965, 0.9969, null, 0.9976, 0.9969, 0.9902, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FVI: { month: [0.9961, null, 0.995, 0.9914, null, null, null], week: [0.9962, 0.9967, 0.9956, 0.9966, 0.994, null, null, null, 0.9953, 0.9963, 0.9957, 0.9965, 0.9922, 0.9839, 0.9963, 0.993, 0.9921, 0.9615, null, null, null, null, null, null, null, null, null, null], day: [0.9957, 0.9967, 0.9969, 0.997, 0.9958, 0.9968, 0.9967, 0.9968, 0.9953, 0.9963, 0.996, 0.9952, 0.9946, 0.9965, null, 0.9959, 0.9966, 0.9973, 0.9961, 0.9963, 0.9976, null, 0.9959, 0.9921, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9953, 0.9961, 0.9956, 0.9983, 0.9959, 0.9956, 0.9962, 0.9964, 0.9956, 0.9982, 0.9953, 0.9946, 0.9954, 0.9948, 0.9955, 0.9973, 0.9954, 0.9989, 0.9953, 0.9959, 0.9972, 0.9943, 0.9938, 0.9948, 0.9898, 0.9899, 0.9903, null, 0.9871, null, null, null, null, null, 0.9808, 0.9992, 0.9978, 0.9973, 0.9969, 0.9898, 0.9967, null, 0.9948, 0.9902, 0.9934, 0.9947, 0.9923, 0.9929, null, 0.9987, 0.9973, 0.9802, null, null, null, null, 0.9615, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_HOOK: { month: [0.9996, 0.9991, 0.9994, 0.9996, null, null, null], week: [0.9997, 0.9996, 0.9995, 0.9997, 0.9996, null, null, null, 0.9993, 0.9995, 0.9993, 0.9996, 0.9993, null, 0.999, 0.9998, 0.9998, null, null, null, null, null, null, null, null, null, null, null], day: [0.9995, 0.9999, 0.9996, 0.9996, 0.9997, 0.9996, 0.9993, 0.9998, 1, 0.9984, 0.9989, 0.9997, 0.9998, 0.9999, null, 0.9998, 0.9996, 0.9998, 0.9999, 0.9997, 0.9997, null, 0.9999, 0.9994, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9991, 0.9995, 0.9995, 0.9997, 0.9993, 0.9997, 0.9993, 0.9997, 0.9993, 0.9996, 0.9992, 0.9993, 0.9993, 0.9995, 0.9992, 0.9999, 0.9995, 0.9998, 0.9998, 0.9995, 0.9993, 0.9995, 0.9989, 0.9995, 0.9996, 0.9995, 0.9994, 0.9987, null, null, null, null, null, null, null, null, null, 0.9994, 0.9998, 0.9997, 0.9993, 0.9967, 0.9994, 0.9994, 1, 0.9998, 0.9995, 0.9996, 1, 0.9999, 0.9996, 0.9998, 0.9998, 0.9997, 0.9995, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_OVEN: { month: [1, 1, 1, 1, null, null, null], week: [1, 1, 1, 1, 1, null, null, null, 1, 1, 1, 1, 1, null, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null], day: [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, null, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, null, null, null, null, null, null, null, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_INDEX: { month: [0.9969, 0.9942, 0.9952, 0.9894, null, null, null], week: [0.9969, 0.9973, 0.9952, 0.9979, 0.9981, null, null, null, 0.9943, 0.9924, 0.9962, 0.9954, 0.9974, null, 0.9928, 0.9856, 0.9904, null, null, null, null, null, null, null, null, null, null, null], day: [0.9966, 0.9972, 0.9975, 0.9982, 0.9969, 0.9972, 0.9973, 0.9968, 0.9953, 0.9948, 0.9948, 0.9972, 0.9924, 0.9967, null, 0.9957, 0.9981, 0.9988, 0.9982, 0.9983, 0.9982, null, 0.9984, 0.9978, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9942, 0.9945, 0.9883, 0.9909, 0.9944, 0.9901, 0.9924, 0.9965, 0.9941, 0.9963, 0.9949, 0.9958, 0.9969, 0.997, 0.9963, 0.9912, 0.9964, 0.9961, 0.996, 0.9964, 0.997, 0.9948, 0.9958, 0.9969, 0.9969, 0.9976, 0.9986, 0.9983, null, null, null, null, null, null, null, null, null, 0.9912, 0.994, 0.9928, 0.9924, 0.9944, 0.9917, 0.9915, 0.9919, 0.9733, 0.9865, 0.9795, 0.9819, 0.9947, 0.9953, 0.997, 0.9897, 0.9917, 0.9888, 0.9888, 0.9815, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
  },
  'SO3560': {
    RTY_TTL: { month: [0.9631, 0.9714, 0.9614, 0.9535, 0.9475, 0.9512, 0.9571], week: [0.9678, 0.9572, 0.9586, 0.9658, 0.9696, 0.9721, 0.9704, 0.9724, 0.9709, 0.9599, 0.9608, 0.9662, 0.96, 0.9557, 0.9529, 0.9509, 0.9544, 0.9541, 0.9485, 0.9403, 0.9459, 0.9536, 0.9473, 0.9496, 0.9525, 0.9581, 0.9518, 0.958], day: [0.9825, 0.9659, 0.96, 0.956, 0.9658, 0.9586, 0.952, 0.9507, 0.9338, 0.9622, 0.987, 0.9625, 0.9656, 0.9582, 0.9656, 0.968, 0.9679, 0.9665, 0.9527, 0.9678, 0.9674, 0.9703, 0.9639, 0.9693, 0.9733, 0.9705, 0.9677, 0.9723, 0.9705, 0.9716, 0.9734, 0.9732, 0.9673, 0.9743, 0.9737, 0.9713, 0.9693, 0.9701, 0.9678, 0.9752, 0.9697, 0.9724, 0.9718, 0.9708, 0.9743, 0.9742, 0.971, 0.9669, 0.9674, 0.9626, 0.9532, 0.9587, 0.9638, 0.9551, 0.9601, 0.9656, 0.9334, 0.9519, 0.9672, 0.9715, 0.9705, 0.9707, 0.9708, 0.9699, 0.9536, 0.9704, 0.9697, 0.9579, 0.9713, 0.9664, 0.9693, 0.9648, 0.9526, 0.9577, 0.9511, 0.9894, 0.9558, 0.9504, 0.9492, 0.9542, 0.9617, 0.9615, 0.9572, 0.9585, 0.9553, 0.9465, 0.9497, 0.9529, 0.9515, 0.9563, 0.9443, 0.9352, 0.9495, 0.9552, 0.9598, 0.9575, 0.9547, 0.9561, 0.9587, 0.9559, 0.954, 0.9523, 0.9497, null, 0.96, 0.9511, 0.9484, 0.9575, 0.9537, 0.945, 0.9517, 0.9591, 0.9488, 0.9459, 0.937, 0.9521, 0.9315, 0.9387, 0.9424, 0.9416, 0.9459, 0.944, 0.9383, 0.9428, 0.9484, 0.9483, 0.9426, 0.9449, 0.9482, 0.9526, 0.9537, 0.9556, 0.9547, 0.9549, 0.9499, 0.9556, 0.9461, 0.9327, 0.9446, 0.9486, 0.9559, 0.9498, 0.9326, 0.9505, 0.9537, 0.9581, 0.9528, 0.953, 0.9588, 0.9583, 0.9468, 0.9474, 0.9508, 0.956, 0.9542, 0.9586, 0.9585, 0.9583, 0.963, 0.9477, 0.9383, 0.9541, 0.9588, 0.9592, 0.9527, 0.9571, 0.9587, 0.9629, 0.9532] },
    RTY_MAIN: { month: [0.9852, 0.9875, 0.9842, 0.9803, 0.9812, 0.9815, 0.982], week: [0.987, 0.9798, 0.9841, 0.9873, 0.9883, 0.9883, 0.9878, 0.9859, 0.9863, 0.9829, 0.9814, 0.988, 0.9854, 0.9799, 0.9797, 0.9778, 0.9835, 0.9823, 0.9822, 0.9794, 0.98, 0.9831, 0.9821, 0.981, 0.9803, 0.9838, 0.9806, 0.9821], day: [null, 0.987, 0.9835, 0.9764, 0.9822, 0.9804, 0.9744, 0.9816, 0.9889, 0.9831, null, 0.9813, 0.9846, 0.981, 0.9855, 0.9873, 0.985, 0.9858, 0.9877, 0.9883, 0.9896, 0.9876, 0.9889, 0.9893, 0.9888, 0.9874, 0.9897, 0.9889, 0.9852, 0.988, 0.9889, 0.9883, 0.9871, 0.9885, 0.9906, 0.9869, 0.9874, 0.9886, 0.9879, 0.9865, 0.9888, 0.9859, 0.9867, 0.9868, 0.987, 0.9881, 0.9874, 0.9846, 0.9837, 0.9837, 0.9828, 0.9846, 0.9797, 0.9806, 0.9836, 0.985, 0.9649, 0.9732, 0.9871, 0.9879, 0.9878, 0.9878, 0.9896, 0.9873, 0.9864, 0.9873, 0.9884, 0.9881, 0.9887, 0.9863, 0.9869, 0.9872, 0.986, 0.9833, 0.9845, 0.9894, 0.9825, 0.9779, 0.9752, 0.9792, 0.9836, 0.9818, 0.979, 0.9818, 0.9821, 0.979, 0.9768, 0.9791, 0.9784, 0.9809, 0.9757, 0.9626, 0.9803, 0.9808, 0.9811, 0.9818, 0.9822, 0.9842, 0.9843, 0.9829, 0.9826, 0.983, 0.9839, null, 0.9831, 0.983, 0.9812, 0.9825, 0.9818, 0.9805, 0.9822, 0.9838, 0.986, 0.9769, 0.9822, 0.9838, 0.9792, 0.9807, 0.9789, 0.9789, 0.9789, 0.9789, 0.9806, 0.9795, 0.9761, 0.9804, 0.9797, 0.9816, 0.983, 0.9826, 0.9838, 0.9838, 0.9814, 0.9858, 0.9813, 0.9819, 0.9854, 0.985, 0.9838, 0.9777, 0.9788, 0.9812, 0.9831, 0.9779, 0.9818, 0.9817, 0.9803, 0.9803, 0.9815, 0.9831, 0.9808, 0.9754, 0.9806, 0.981, 0.9808, 0.9841, 0.986, 0.9866, 0.9847, 0.9812, 0.9753, 0.977, 0.9833, 0.9826, 0.9845, 0.9832, 0.9821, 0.9823, 0.9806] },
    RTY_SUB1: { month: [0.9955, 0.9976, 0.9965, 0.9963, 0.9961, 0.9962, 0.9958], week: [0.9944, 0.9943, 0.9954, 0.9962, 0.9964, 0.9974, 0.9973, 0.9987, 0.9978, 0.9969, 0.9964, 0.9968, 0.9958, 0.9964, 0.9959, 0.9964, 0.9963, 0.9958, 0.9968, 0.996, 0.9955, 0.9964, 0.9965, 0.9969, 0.9962, 0.9956, 0.9958, 0.9955], day: [0.9942, 0.9947, 0.9942, 0.9953, 0.9951, 0.9921, 0.994, 0.995, 0.993, null, null, 0.9958, 0.9951, 0.9965, 0.9968, 0.9962, 0.9955, 0.9956, 0.9958, 0.9953, 0.9964, 0.9984, 0.9972, 0.997, 0.997, 0.9973, 0.9922, 0.9977, 0.9965, 0.9971, 0.9975, 0.9978, 0.9979, 0.997, 0.9977, 0.9966, 0.9973, 0.9969, 0.9974, 0.9974, 0.9973, 0.9987, 0.9983, 0.9973, 0.9985, 0.9982, 0.998, 0.9976, 0.9967, 0.9972, 0.9965, 0.997, 0.997, 0.9962, 0.9974, 0.997, 0.9964, 0.9965, 0.9963, 0.9966, 0.9958, 0.997, 0.9968, 0.9965, 0.997, 0.997, 0.9965, 0.9969, 0.9972, 0.9959, 0.9965, 0.9965, 0.9956, 0.9955, 0.9947, null, 0.9969, 0.9959, 0.9968, 0.9962, 0.9958, 0.9963, 0.9966, 0.9957, 0.9966, 0.9956, 0.9943, 0.9961, 0.9962, 0.997, 0.9967, 0.9957, 0.9966, 0.9968, 0.996, 0.9964, 0.9965, 0.9967, 0.9971, 0.9972, 0.9967, 0.9962, 0.9938, null, 0.9955, 0.9971, 0.997, 0.9958, 0.9935, 0.9964, 0.995, 0.9973, 0.9973, 0.9977, 0.9972, 0.9969, 0.9967, 0.9966, 0.9962, 0.9968, 0.9949, 0.9941, 0.9966, 0.9957, 0.9949, 0.9952, 0.9953, 0.996, 0.9963, 0.9966, 0.9961, 0.9966, 0.9964, 0.9968, 0.996, 0.9959, 0.9966, 0.9969, 0.9965, 0.9961, 0.9972, 0.9971, 0.9974, 0.9961, 0.9971, 0.9971, 0.9964, 0.9968, 0.9957, 0.9964, 0.9959, 0.9963, 0.996, 0.9955, 0.9958, 0.9957, 0.9958, 0.9954, 0.9955, 0.9953, 0.9955, 0.9964, 0.996, 0.9957, 0.9959, 0.9952, 0.9952, 0.996, 0.9959] },
    RTY_SUB2: { month: [0.982, 0.9862, 0.9803, 0.9762, 0.9694, 0.9728, 0.9788], week: [0.9861, 0.9826, 0.9786, 0.9819, 0.9846, 0.9862, 0.9851, 0.9876, 0.9865, 0.9796, 0.9825, 0.9811, 0.9783, 0.9789, 0.9766, 0.976, 0.9741, 0.9754, 0.9688, 0.964, 0.9695, 0.9734, 0.9679, 0.971, 0.9754, 0.9781, 0.9747, 0.9798], day: [0.9882, 0.9839, 0.9818, 0.9838, 0.9881, 0.9855, 0.9829, 0.9733, 0.9509, 0.9788, 0.987, 0.985, 0.9855, 0.9802, 0.9829, 0.9842, 0.987, 0.9847, 0.9686, 0.9838, 0.981, 0.984, 0.9775, 0.9826, 0.9873, 0.9856, 0.9855, 0.9854, 0.9885, 0.9862, 0.9868, 0.9869, 0.982, 0.9886, 0.9852, 0.9876, 0.9843, 0.9844, 0.9822, 0.9911, 0.9834, 0.9876, 0.9865, 0.9865, 0.9886, 0.9878, 0.9853, 0.9843, 0.9866, 0.9813, 0.9733, 0.9766, 0.9867, 0.9777, 0.9786, 0.9832, 0.9708, 0.9815, 0.9834, 0.9868, 0.9866, 0.9857, 0.9842, 0.9858, 0.9697, 0.9859, 0.9846, 0.9724, 0.9852, 0.9838, 0.9855, 0.9807, 0.9704, 0.9784, 0.9712, null, 0.9759, 0.9759, 0.9764, 0.9782, 0.9819, 0.9829, 0.9811, 0.9805, 0.976, 0.9711, 0.9778, 0.977, 0.9762, 0.9778, 0.971, 0.9757, 0.9719, 0.9771, 0.9823, 0.9789, 0.9754, 0.9747, 0.9768, 0.9752, 0.9741, 0.9724, 0.9712, null, 0.9809, 0.9703, 0.9694, 0.9787, 0.9777, 0.9673, 0.9739, 0.9775, 0.9649, 0.9706, 0.9567, 0.9708, 0.9545, 0.9604, 0.9664, 0.965, 0.9712, 0.9701, 0.9601, 0.9667, 0.9766, 0.972, 0.9667, 0.9665, 0.9683, 0.9728, 0.9732, 0.9746, 0.9764, 0.9718, 0.9719, 0.9773, 0.9635, 0.9499, 0.9635, 0.9741, 0.9794, 0.9709, 0.951, 0.9758, 0.9742, 0.9787, 0.9754, 0.9752, 0.9811, 0.9783, 0.9693, 0.9749, 0.9735, 0.979, 0.9769, 0.9784, 0.9763, 0.9759, 0.9824, 0.9705, 0.9665, 0.9801, 0.9791, 0.9804, 0.9716, 0.9782, 0.9809, 0.9842, 0.9761] },
    MAIN_FVI: { month: [0.9927, 0.9928, 0.9907, 0.9915, 0.9916, 0.992, 0.9919], week: [0.992, 0.9926, 0.9937, 0.9918, 0.9929, 0.9927, 0.9926, 0.9929, 0.9928, 0.9903, 0.9888, 0.9921, 0.9912, 0.9912, 0.992, 0.991, 0.9922, 0.9913, 0.9922, 0.9901, 0.9907, 0.9935, 0.9913, 0.9917, 0.9925, 0.9928, 0.9918, 0.992], day: [null, 0.992, 0.9936, 0.9933, 0.9925, 0.9936, 0.9916, 0.9912, 0.997, 0.9941, null, 0.9884, 0.9947, 0.9933, 0.9945, 0.9926, 0.9899, 0.9924, 0.9913, 0.9918, 0.9934, 0.9913, 0.9926, 0.9937, 0.9934, 0.9925, 0.9935, 0.9933, 0.9914, 0.9926, 0.9923, 0.992, 0.9923, 0.992, 0.9945, 0.993, 0.9923, 0.9919, 0.9924, 0.9928, 0.9938, 0.9929, 0.9931, 0.9929, 0.9929, 0.9943, 0.9941, 0.9918, 0.9906, 0.9901, 0.9895, 0.9914, 0.9895, 0.9899, 0.9908, 0.991, 0.975, 0.9881, 0.993, 0.9927, 0.9921, 0.9921, 0.9934, 0.9907, 0.9902, 0.991, 0.9929, 0.9932, 0.9934, 0.9903, 0.9909, 0.9926, 0.9921, 0.9908, 0.9922, 0.9894, 0.9933, 0.9901, 0.9918, 0.9914, 0.9901, 0.9916, 0.9903, 0.9927, 0.9922, 0.9922, 0.9908, 0.9919, 0.9914, 0.9928, 0.9907, 0.9909, 0.9912, 0.9907, 0.9903, 0.9914, 0.9917, 0.9929, 0.9927, 0.9917, 0.9918, 0.9921, 0.992, null, 0.9929, 0.9912, 0.989, 0.9917, 0.9916, 0.9921, 0.993, 0.9934, 0.9949, 0.9912, 0.99, 0.9912, 0.9889, 0.9915, 0.9898, 0.9898, 0.9898, 0.9898, 0.9909, 0.9909, 0.9902, 0.9907, 0.9901, 0.9908, 0.9914, 0.992, 0.993, 0.993, 0.9927, 0.9967, 0.9934, 0.993, 0.994, 0.9928, 0.9927, 0.9864, 0.9889, 0.9888, 0.9916, 0.9911, 0.9926, 0.9921, 0.9938, 0.994, 0.9919, 0.9929, 0.9918, 0.9917, 0.9926, 0.9922, 0.9913, 0.9927, 0.9933, 0.9937, 0.9935, 0.9931, 0.99, 0.992, 0.9916, 0.9914, 0.9926, 0.9923, 0.9929, 0.9916, 0.9912] },
    MAIN_ASSY: { month: [0.996, 0.9982, 0.998, 0.9973, 0.998, 0.9982, 0.9979], week: [0.997, 0.9922, 0.994, 0.9986, 0.9987, 0.9985, 0.9979, 0.9963, 0.998, 0.9978, 0.9976, 0.9984, 0.9984, 0.9973, 0.9964, 0.9978, 0.9975, 0.9979, 0.9983, 0.9979, 0.9977, 0.9981, 0.9985, 0.9978, 0.9979, 0.9984, 0.9984, 0.9974], day: [null, 0.997, 0.9975, 0.9881, 0.9938, 0.9903, 0.9879, 0.9959, 0.9952, 0.9889, null, 0.9969, 0.9935, 0.9941, 0.9954, 0.9977, 0.9986, 0.9988, 0.9989, 0.9986, 0.999, 0.9986, 0.999, 0.9983, 0.999, 0.9989, 0.999, 0.9981, 0.9984, 0.9983, 0.9992, 0.9985, 0.997, 0.9992, 0.9991, 0.9985, 0.9982, 0.9989, 0.9979, 0.9966, 0.9982, 0.9963, 0.9977, 0.9985, 0.9986, 0.9986, 0.9981, 0.998, 0.9968, 0.9975, 0.9977, 0.9985, 0.9985, 0.9963, 0.9975, 0.9984, 0.9949, 0.9984, 0.9974, 0.9985, 0.9983, 0.9983, 0.9988, 0.9985, 0.9987, 0.9984, 0.998, 0.9975, 0.9989, 0.9988, 0.999, 0.9987, 0.9984, 0.9972, 0.9981, null, 0.9975, 0.9981, 0.9954, 0.9984, 0.9991, 0.9956, 0.9972, 0.9963, 0.9971, 0.9933, 0.9969, 0.9971, 0.996, 0.9981, 0.9977, 0.9976, 0.997, 0.9982, 0.9984, 0.9975, 0.9984, 0.9985, 0.9985, 0.9972, 0.9969, 0.9968, 0.9974, null, 0.9961, 0.9979, 0.9988, 0.9989, 0.9979, 0.998, 0.998, 0.9988, 0.9978, 0.9978, 0.9983, 0.9994, 0.9982, 0.9976, 0.9979, 0.9979, 0.9979, 0.9979, 0.9977, 0.9971, 0.9945, 0.9991, 0.9987, 0.9989, 0.9982, 0.998, 0.9984, 0.9983, 0.9976, 0.9985, 0.9978, 0.9967, 0.9985, 0.9991, 0.9988, 0.9991, 0.9986, 0.9992, 0.9993, 0.9955, 0.9977, 0.9992, 0.9962, 0.9981, 0.9985, 0.9988, 0.9985, 0.9955, 0.9979, 0.9977, 0.9979, 0.9989, 0.999, 0.999, 0.9977, 0.9985, 0.9983, 0.9991, 0.9986, 0.9979, 0.9983, 0.9975, 0.997, 0.9981, 0.9972] },
    MAIN_DRIVING: { month: [0.9974, 0.9969, 0.9961, 0.9917, 0.9918, 0.9918, 0.9921], week: [0.9982, 0.9963, 0.9977, 0.9979, 0.9971, 0.9974, 0.9978, 0.997, 0.9959, 0.995, 0.997, 0.9977, 0.9961, 0.992, 0.9915, 0.989, 0.9938, 0.9932, 0.9918, 0.9917, 0.9919, 0.9918, 0.9928, 0.992, 0.9902, 0.9931, 0.9905, 0.9926], day: [null, 0.9982, 0.9963, 0.9961, 0.9965, 0.9974, 0.9957, 0.9961, 0.9966, 1, null, 0.9967, 0.9974, 0.9977, 0.9977, 0.9978, 0.9978, 0.9978, 0.9981, 0.9984, 0.9977, 0.9978, 0.9977, 0.9978, 0.9966, 0.9962, 0.9976, 0.9979, 0.9957, 0.9975, 0.9977, 0.9981, 0.9981, 0.9975, 0.9973, 0.9959, 0.9973, 0.9985, 0.9982, 0.9977, 0.9971, 0.997, 0.9962, 0.9957, 0.9956, 0.9954, 0.9956, 0.9964, 0.9964, 0.9963, 0.9959, 0.9952, 0.9921, 0.9946, 0.9954, 0.9958, 0.9961, 0.9966, 0.997, 0.9969, 0.9976, 0.9976, 0.9976, 0.9983, 0.9976, 0.9981, 0.9979, 0.9978, 0.9966, 0.9976, 0.9972, 0.9962, 0.9956, 0.9955, 0.9944, null, 0.9926, 0.991, 0.9888, 0.9906, 0.9946, 0.9947, 0.9918, 0.993, 0.993, 0.9938, 0.9892, 0.9902, 0.991, 0.99, 0.9874, 0.9741, 0.9922, 0.9919, 0.9925, 0.993, 0.9922, 0.9929, 0.9933, 0.9942, 0.9939, 0.9943, 0.9943, null, 0.994, 0.994, 0.9935, 0.9921, 0.9923, 0.9904, 0.9913, 0.9918, 0.9934, 0.9882, 0.9941, 0.9936, 0.9922, 0.9919, 0.9913, 0.9913, 0.9913, 0.9913, 0.9927, 0.9917, 0.9917, 0.991, 0.9913, 0.9923, 0.9936, 0.993, 0.9927, 0.9929, 0.9914, 0.9909, 0.99, 0.9924, 0.9932, 0.9933, 0.993, 0.9931, 0.9917, 0.9936, 0.9923, 0.9915, 0.9922, 0.9914, 0.9913, 0.9891, 0.9913, 0.9915, 0.9907, 0.9884, 0.9902, 0.9915, 0.9922, 0.9932, 0.994, 0.9941, 0.9936, 0.9897, 0.9872, 0.9858, 0.9932, 0.9934, 0.9938, 0.9934, 0.9922, 0.9927, 0.9924] },
    MAIN_TILT: { month: [0.999, 0.9995, 0.9993, 0.9997, 0.9997, 0.9995, 0.9999], week: [0.9998, 0.9984, 0.9986, 0.999, 0.9997, 0.9996, 0.9994, 0.9997, 0.9995, 0.9997, 0.9979, 0.9997, 0.9998, 0.9992, 0.9997, 0.9998, 0.9998, 0.9998, 0.9998, 0.9996, 0.9996, 0.9996, 0.9994, 0.9994, 0.9996, 0.9995, 0.9999, 0.9999], day: [null, 0.9998, 0.996, 0.9988, 0.9994, 0.999, 0.999, 0.9983, 1, 1, null, 0.9992, 0.999, 0.9958, 0.9979, 0.9991, 0.9987, 0.9969, 0.9994, 0.9994, 0.9994, 0.9998, 0.9995, 0.9996, 0.9998, 0.9998, 0.9996, 0.9997, 0.9996, 0.9996, 0.9997, 0.9997, 0.9997, 0.9998, 0.9996, 0.9995, 0.9995, 0.9992, 0.9993, 0.9994, 0.9997, 0.9997, 0.9997, 0.9996, 0.9997, 0.9997, 0.9996, 0.9983, 0.9997, 0.9997, 0.9996, 0.9995, 0.9996, 0.9997, 0.9998, 0.9998, 0.9988, 0.9899, 0.9996, 0.9998, 0.9997, 0.9997, 0.9997, 0.9999, 0.9998, 0.9998, 0.9996, 0.9995, 0.9997, 0.9997, 0.9998, 0.9998, 0.9998, 0.9998, 0.9997, null, 0.999, 0.9984, 0.999, 0.9987, 0.9997, 0.9998, 0.9996, 0.9997, 0.9997, 0.9996, 0.9997, 0.9998, 0.9998, 0.9999, 0.9997, 0.9996, 0.9998, 0.9998, 0.9998, 0.9998, 0.9998, 0.9998, 0.9997, 0.9998, 0.9999, 0.9998, 1, null, 0.9999, 0.9999, 0.9998, 0.9997, 0.9999, 0.9999, 0.9999, 0.9998, 0.9998, 0.9996, 0.9997, 0.9997, 0.9997, 0.9997, 0.9997, 0.9997, 0.9997, 0.9997, 0.9993, 0.9996, 0.9996, 0.9995, 0.9995, 0.9995, 0.9996, 0.9996, 0.9995, 0.9995, 0.9995, 0.9995, 1, 0.9997, 0.9996, 0.9997, 0.9993, 0.9989, 0.9993, 0.9995, 0.9999, 0.9997, 0.9993, 0.999, 0.9988, 0.9989, 0.9997, 0.9998, 0.9996, 0.9997, 0.9997, 0.9995, 0.9992, 0.9993, 0.9997, 0.9998, 0.9998, 1, 0.9996, 0.9999, 0.9998, 0.9998, 0.9999, 0.9999, 0.9999, 0.9999, 0.9998] },
    SUB1_FPCB: { month: [0.9978, 0.9987, 0.9982, 0.9981, 0.9981, 0.9982, 0.9979], week: [0.9975, 0.9972, 0.9983, 0.9978, 0.9981, 0.9986, 0.9985, 0.9994, 0.9989, 0.9984, 0.9981, 0.9983, 0.9978, 0.9981, 0.9979, 0.9983, 0.9981, 0.9978, 0.9984, 0.9982, 0.9978, 0.9983, 0.9983, 0.9985, 0.9981, 0.9978, 0.9982, 0.9976], day: [0.9974, 0.9976, 0.9973, 0.9976, 0.9976, 0.9968, 0.9966, 0.9973, 1, null, null, 0.9979, 0.9973, 0.9979, 0.9982, 0.9975, 0.9975, 0.9975, 0.9977, 0.9974, 0.998, 0.999, 0.9988, 0.9985, 0.9984, 0.9987, 0.996, 0.9986, 0.998, 0.9985, 0.9988, 0.9988, 0.999, 0.9985, 0.9987, 0.9981, 0.9984, 0.9984, 0.9986, 0.9986, 0.9986, 0.9994, 0.9991, 0.9986, 0.9991, 0.999, 0.999, 0.9989, 0.9986, 0.9985, 0.9982, 0.9983, 0.9983, 0.9981, 0.9987, 0.9984, 0.9982, 0.9982, 0.9981, 0.9982, 0.9979, 0.9982, 0.9984, 0.9982, 0.9984, 0.9983, 0.9981, 0.9983, 0.9985, 0.9979, 0.9982, 0.998, 0.9977, 0.9976, 0.9974, null, 0.9982, 0.9979, 0.9983, 0.998, 0.9979, 0.9982, 0.9985, 0.9975, 0.9982, 0.9981, 0.9972, 0.9981, 0.9979, 0.9982, 0.9986, 0.9981, 0.9981, 0.9982, 0.9982, 0.9982, 0.9985, 0.9982, 0.9984, 0.9987, 0.9983, 0.9981, 0.9972, null, 0.9979, 0.9984, 0.9981, 0.9979, 0.9966, 0.9982, 0.9975, 0.9985, 0.9987, 0.9986, 0.9991, 0.998, 0.9983, 0.9984, 0.9981, 0.9987, 0.9977, 0.997, 0.9992, 0.998, 0.9973, 0.9977, 0.9979, 0.9978, 0.9981, 0.9981, 0.9981, 0.9982, 0.9984, 0.9985, 0.9986, 0.9979, 0.9981, 0.9986, 0.9984, 0.9982, 0.9986, 0.9987, 0.9988, 0.9979, 0.9984, 0.9987, 0.9982, 0.9986, 0.9979, 0.9981, 0.9977, 0.9981, 0.998, 0.9979, 0.9979, 0.998, 0.9978, 0.9979, 0.9974, 0.9978, 0.9985, 0.9982, 0.9982, 0.998, 0.9983, 0.9976, 0.9974, 0.9978, 0.9977] },
    SUB1_FVI: { month: [0.9977, 0.9989, 0.9984, 0.9981, 0.998, 0.9981, 0.9979], week: [0.9969, 0.9971, 0.9972, 0.9984, 0.9983, 0.9988, 0.9987, 0.9992, 0.9989, 0.9985, 0.9983, 0.9985, 0.998, 0.9982, 0.9981, 0.9981, 0.9982, 0.998, 0.9984, 0.9978, 0.9977, 0.9981, 0.9982, 0.9984, 0.9981, 0.9978, 0.9976, 0.9979], day: [0.9968, 0.9971, 0.9969, 0.9977, 0.9975, 0.9953, 0.9974, 0.9978, 0.993, null, null, 0.9979, 0.9978, 0.9985, 0.9986, 0.9987, 0.998, 0.9981, 0.9982, 0.9979, 0.9985, 0.9994, 0.9985, 0.9985, 0.9986, 0.9986, 0.9962, 0.9991, 0.9986, 0.9986, 0.9987, 0.999, 0.9989, 0.9986, 0.999, 0.9986, 0.9989, 0.9985, 0.9988, 0.9988, 0.9987, 0.9992, 0.9992, 0.9987, 0.9993, 0.9992, 0.999, 0.9988, 0.9981, 0.9987, 0.9983, 0.9987, 0.9987, 0.9981, 0.9987, 0.9986, 0.9982, 0.9983, 0.9982, 0.9984, 0.9979, 0.9987, 0.9985, 0.9983, 0.9985, 0.9987, 0.9984, 0.9986, 0.9987, 0.998, 0.9984, 0.9985, 0.9979, 0.9978, 0.9973, null, 0.9988, 0.9979, 0.9985, 0.9982, 0.9979, 0.9982, 0.9981, 0.9982, 0.9984, 0.9975, 0.9971, 0.998, 0.9983, 0.9988, 0.9981, 0.9976, 0.9985, 0.9985, 0.9978, 0.9981, 0.998, 0.9985, 0.9986, 0.9985, 0.9984, 0.9982, 0.9967, null, 0.9976, 0.9987, 0.9989, 0.9979, 0.9969, 0.9982, 0.9974, 0.9988, 0.9986, 0.9991, 0.998, 0.9989, 0.9983, 0.9982, 0.9981, 0.9982, 0.9972, 0.9971, 0.9974, 0.9977, 0.9976, 0.9975, 0.9973, 0.9981, 0.9981, 0.9984, 0.9979, 0.9984, 0.998, 0.9984, 0.9974, 0.9979, 0.9984, 0.9983, 0.9981, 0.9979, 0.9986, 0.9983, 0.9986, 0.9982, 0.9986, 0.9985, 0.9983, 0.9982, 0.9978, 0.9983, 0.9982, 0.9982, 0.9979, 0.9976, 0.998, 0.9977, 0.9979, 0.9975, 0.9981, 0.9974, 0.9969, 0.9981, 0.9977, 0.9977, 0.9976, 0.9975, 0.9978, 0.9982, 0.9981] },
    SUB2_HOOK: { month: [0.9997, 0.9998, 0.9997, 0.9996, 0.9997, 0.9997, 0.9999], week: [0.9998, 0.9996, 0.9997, 0.9998, 0.9997, 0.9997, 0.9998, 1, 0.9998, 0.9998, 0.9998, 0.9997, 0.9997, 0.9995, 0.9994, 0.9996, 0.9997, 0.9997, 0.9997, 0.9996, 0.9998, 0.9998, 0.9997, 0.9997, 0.9997, 0.9998, 0.9999, 0.9999], day: [1, 0.9996, 0.999, 0.9996, 1, 0.9999, 0.9993, 0.9996, 0.9996, 0.9994, 0.9996, 0.9997, 0.9997, 0.9999, 0.9998, 0.9998, 0.9999, 1, 0.9999, 0.9997, 0.9995, 0.9998, 0.9997, 0.9998, 0.9996, 0.9994, 0.9999, 0.9998, 0.9997, 0.9993, 0.9999, 0.9997, 0.9999, 0.9997, 0.9998, 0.9998, 0.9999, 0.9995, 0.9999, 0.9997, 0.9999, 1, 0.9999, 0.9999, 0.9999, 0.9999, 0.9998, 0.9996, 0.9997, 0.9997, 0.9998, 0.9997, 0.9999, 0.9997, 0.9996, 0.9999, 0.9997, 0.9996, 0.9999, 0.9999, 0.9999, 0.9996, 0.9996, 0.9999, 0.9996, 0.9999, 0.9999, 0.9994, 0.9999, 0.9995, 0.9998, 0.9999, 0.9994, 0.9997, 0.9997, null, 0.9998, 0.9995, 0.9995, 0.9994, 0.9993, 0.9995, 0.9999, 0.9995, 0.9997, 0.9994, 0.9994, 0.9992, 0.9992, 0.9993, 0.9998, 0.9995, 0.9993, 0.9996, 0.9998, 0.9997, 0.9994, 0.9998, 0.9995, 0.9998, 0.9995, 0.9999, 1, null, 0.9998, 0.9996, 0.9994, 0.9998, 0.9998, 0.9998, 0.9998, 0.9996, 0.9995, 0.9997, 0.9999, 0.9995, 0.9999, 0.9996, 0.9996, 0.9991, 0.9999, 0.9999, 0.9994, 0.9997, 0.9999, 0.9999, 1, 0.9999, 0.9996, 0.9999, 0.9998, 0.9998, 0.9998, 0.9996, 1, 0.9999, 0.9998, 0.9997, 0.9997, 0.9997, 0.9994, 0.9994, 0.9998, 0.9997, 0.9996, 0.9999, 0.9997, 0.9998, 0.9998, 0.9993, 0.9997, 0.9997, 0.9997, 0.9997, 0.9998, 0.9997, 0.9998, 0.9999, 0.9996, 0.9998, 1, 0.9999, 0.9998, 0.9997, 0.9999, 0.9999, 0.9999, 0.9999, 0.9999] },
    SUB2_OVEN: { month: [1.0, 1.0, 1, 1, 1, 1, 1], week: [1, 1, 1, 1, 0.9999, 1.0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1], day: [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0.9998, 0.9998, 1, 1, 1, 1, 1, 1, 0.9998, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1] },
    SUB2_INDEX: { month: [0.9944, 0.9968, 0.9964, 0.9938, 0.9933, 0.9947, 0.9947], week: [0.9934, 0.9941, 0.9927, 0.9956, 0.9957, 0.9971, 0.9961, 0.994, 0.9974, 0.9952, 0.9964, 0.9969, 0.9968, 0.9955, 0.994, 0.9954, 0.9922, 0.9926, 0.992, 0.9929, 0.994, 0.9942, 0.9951, 0.9946, 0.9947, 0.9946, 0.9941, 0.995], day: [0.9936, 0.9932, 0.9932, 0.9915, 0.9956, 0.9955, 0.9937, 0.9951, 0.9825, 0.9946, 0.9968, 0.9927, 0.9965, 0.9925, 0.9932, 0.9948, 0.9962, 0.9957, 0.9951, 0.9959, 0.9955, 0.996, 0.9954, 0.9924, 0.997, 0.9972, 0.9962, 0.9949, 0.9971, 0.9949, 0.9967, 0.9968, 0.9973, 0.9982, 0.9977, 0.9983, 0.9974, 0.9962, 0.9965, 0.9985, 0.9922, 0.994, 0.9981, 0.9979, 0.9979, 0.996, 0.9972, 0.9974, 0.9972, 0.9953, 0.9953, 0.9925, 0.9961, 0.9936, 0.997, 0.9969, 0.9956, 0.9961, 0.9953, 0.9962, 0.9975, 0.9978, 0.9967, 0.9969, 0.9971, 0.9969, 0.9945, 0.9969, 0.999, 0.9969, 0.9972, 0.9975, 0.9965, 0.9975, 0.9954, null, 0.9973, 0.9964, 0.9956, 0.9944, 0.9964, 0.9956, 0.9931, 0.9954, 0.9947, 0.9877, 0.9951, 0.9945, 0.9953, 0.995, 0.9965, 0.9959, 0.9914, 0.9958, 0.9968, 0.9951, 0.9965, 0.9906, 0.994, 0.9911, 0.9904, 0.9907, 0.9961, null, 0.9944, 0.9882, 0.9911, 0.9937, 0.9955, 0.9891, 0.9918, 0.9955, 0.9844, 0.9944, 0.9943, 0.9942, 0.9931, 0.989, 0.9937, 0.9936, 0.9955, 0.9942, 0.9911, 0.9956, 0.9954, 0.9956, 0.9951, 0.9934, 0.989, 0.9924, 0.9943, 0.9968, 0.9977, 0.9911, 0.9931, 0.9932, 0.9926, 0.9943, 0.9956, 0.9968, 0.9979, 0.9944, 0.993, 0.9953, 0.9955, 0.9953, 0.9941, 0.9957, 0.9961, 0.9948, 0.9929, 0.9948, 0.9942, 0.9935, 0.996, 0.9943, 0.9945, 0.9921, 0.9972, 0.9911, 0.9956, 0.9939, 0.995, 0.997, 0.992, 0.9915, 0.9959, 0.9974, 0.9951] },
  },
  'SO1C2EDM': {
    RTY_TTL: { month: [null, 0.9542, 0.9339, 0.9858, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9533, 0.941, null, null, null, 0.9858, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9759, 0.9656, 0.9799, 0.9854, 0.9857, 0.9599, 0.9472, 0.934, 0.9761, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9858, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_MAIN: { month: [null, 0.9782, 0.946, 0.9858, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9782, 0.946, null, null, null, 0.9858, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9782, null, null, null, null, 0.9666, 0.9506, 0.934, 0.9761, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9858, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB1: { month: [null, 0.9956, 0.9939, null, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9945, 0.9947, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9968, 0.9944, null, 0.9923, 0.9931, 0.9964, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    RTY_SUB2: { month: [null, 0.9799, 0.9933, null, null, null, null], week: [null, null, null, null, null, null, null, null, 0.98, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9977, 0.9688, 0.9854, 0.9854, 0.9933, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_FVI: { month: [null, null, 0.9904, null, null, null, null], week: [null, null, null, null, null, null, null, null, null, 0.9904, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9941, 0.995, 0.9964, 0.9761, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_ASSY: { month: [null, 1, 0.9936, 1, null, null, null], week: [null, null, null, null, null, null, null, null, 1, 0.9936, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, 0.9863, 0.9944, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_DRIVING: { month: [null, 0.9782, 0.9624, 0.9858, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9782, 0.9624, null, null, null, 0.9858, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9782, null, null, null, null, 0.9862, 0.9625, 0.9385, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9858, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    MAIN_TILT: { month: [null, 1, 0.9989, 1, null, null, null], week: [null, null, null, null, null, null, null, null, 1, 0.9989, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, 0.9996, 0.9983, 0.9988, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FPCB: { month: [null, 0.9976, 0.9976, null, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9977, 0.9975, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9983, 0.997, null, 0.9979, 0.9973, 0.9977, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB1_FVI: { month: [null, 0.9979, 0.9963, null, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9968, 0.9972, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9985, 0.9974, null, 0.9944, 0.9958, 0.9987, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_HOOK: { month: [null, 0.9999, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9999, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9998, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_OVEN: { month: [null, 1, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
    SUB2_INDEX: { month: [null, 0.9893, null, null, null, null, null], week: [null, null, null, null, null, null, null, null, 0.9893, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9842, 0.9934, 0.9902, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null] },
  },
  'SO2701': {
    RTY_TTL: { month: [0.9681, 0.9652, 0.9631, 0.9541, 0.9451, 0.956, 0.9561], week: [0.9728, 0.9706, 0.9681, 0.9622, 0.9706, 0.9695, 0.9732, 0.984, 0.9527, 0.9669, 0.9708, 0.9651, 0.9558, 0.9435, 0.9497, 0.9571, 0.9624, 0.9597, 0.947, 0.946, 0.9481, 0.9322, 0.9522, 0.9633, 0.9528, 0.9561, 0.9537, 0.9589], day: [0.9709, 0.9748, 0.9733, 0.974, 0.9668, 0.9634, 0.9759, 0.9712, 0.9764, 0.9728, 0.9597, 0.9707, 0.9571, 0.9718, null, 0.9717, 0.9751, 0.9671, 0.968, 0.9315, 0.9514, 0.9952, 0.9619, 0.9698, 0.9748, 0.9775, 0.9721, 0.9678, null, 0.9671, 0.9709, 0.9582, 0.9772, 0.9724, 0.9676, 0.9732, 0.9756, 0.9757, 0.9799, 0.9769, 0.9581, 0.984, 0.9608, 0.9659, 0.9516, 0.9585, 0.9603, 0.9284, 0.9732, 0.9728, 0.9888, 0.9666, 0.9765, 0.9725, 0.9646, 0.9516, 0.9709, 0.9749, 0.9722, 0.9693, 0.9728, 0.9646, 0.9647, 0.9664, 0.967, 0.9635, 0.9679, 0.9659, 0.9599, 0.9571, 0.9516, 0.9574, 0.9556, 0.9581, 0.9553, 0.9952, 0.9545, 0.9362, 0.9353, 0.9407, 0.9426, 0.9446, 0.9504, 0.958, 0.9486, 0.9542, 0.9549, 0.9424, 0.9413, 0.9844, 0.9502, 0.9609, 0.9568, 0.9588, 0.9589, 0.9595, 0.9598, 0.9595, 0.9624, 0.9623, 0.9634, 0.9638, 0.963, null, 0.9584, 0.9502, 0.9647, 0.9664, 0.9587, 0.9595, 0.9542, 0.9463, 0.9398, 0.9375, 0.9426, 0.9705, 0.9498, 0.9418, 0.9432, 0.9455, 0.9468, 0.9484, 0.989, 0.9527, 0.9609, 0.9584, 0.9484, 0.9671, 0.9536, 0.9467, 0.9243, 0.9572, 0.9198, 0.9422, 0.9475, 0.9441, 0.9542, 0.9485, 0.9566, 0.9543, 0.9558, 0.9632, 0.9647, 0.9672, 0.9572, 0.9644, 0.9632, 0.9565, 0.9608, 0.9549, 0.953, 0.9419, 0.9497, 0.9615, 0.9572, 0.9559, 0.9516, 0.9557, 0.9549, 0.9559, 0.9532, 0.9531, 0.9459, 0.9532, 0.9609, 0.9575, 0.9576, 0.9609, 0.9595] },
    RTY_MAIN: { month: [0.9833, 0.9842, 0.9859, 0.9846, 0.9853, 0.9874, 0.9875], week: [0.9875, 0.9869, 0.9833, 0.9774, 0.985, 0.9874, 0.9884, 0.9881, 0.9779, 0.987, 0.9883, 0.9866, 0.9828, 0.9827, 0.9856, 0.9854, 0.9858, 0.983, 0.9862, 0.9849, 0.9859, 0.9824, 0.9872, 0.9878, 0.9851, 0.9897, 0.9879, 0.9867], day: [0.9854, 0.9896, 0.9896, 0.9866, 0.9782, 0.9881, 0.99, 0.9893, 0.9886, 0.9887, 0.9756, 0.986, 0.9755, 0.9852, null, 0.9902, 0.9859, 0.982, 0.983, 0.9467, 0.9661, 0.9952, 0.9746, 0.9855, 0.9876, 0.9891, 0.9866, 0.9865, null, 0.9852, 0.9848, 0.9843, 0.9912, 0.9905, 0.987, 0.989, 0.9893, 0.9893, 0.9909, 0.9866, 0.986, 0.9881, 0.9742, 0.9797, 0.9664, 0.9755, 0.9868, 0.9747, 0.9882, 0.9933, null, 0.9881, 0.9897, 0.9889, 0.9831, 0.9834, 0.9896, 0.9899, 0.9883, 0.987, 0.9887, 0.9864, 0.9908, 0.986, 0.9883, 0.9865, 0.9873, 0.9863, 0.9809, 0.9857, 0.9791, 0.9833, 0.9818, 0.9846, 0.983, 0.9952, 0.9832, 0.9799, 0.9831, 0.9824, 0.9828, 0.9836, 0.9837, 0.9859, 0.9859, 0.9852, 0.988, 0.9853, 0.9842, 0.9844, 0.9851, 0.987, 0.9855, 0.9872, 0.9841, 0.9845, 0.9841, 0.9852, 0.9866, 0.9857, 0.9856, 0.9862, 0.9852, null, 0.9852, 0.9716, 0.9849, 0.9868, 0.9862, 0.9882, 0.9862, 0.9866, 0.9854, 0.987, 0.9829, 0.9872, 0.9868, 0.9862, 0.9821, 0.9856, 0.9833, 0.9851, 0.989, 0.9838, 0.9913, 0.9946, null, null, null, null, null, 0.9954, 0.9783, 0.9817, 0.986, 0.9843, 0.9856, 0.988, 0.9859, 0.9892, 0.99, 0.9889, 0.9886, 0.9887, 0.9859, 0.9885, 0.9865, 0.9835, 0.9866, 0.9879, 0.9889, 0.9768, 0.9868, 0.9913, 0.9888, 0.9876, 0.9899, 0.9894, 0.9911, 0.9906, 0.9837, 0.9882, 0.9881, 0.9867, 0.9899, 0.9884, 0.9867, 0.9855, 0.9862] },
    RTY_SUB1: { month: [0.9947, 0.9957, 0.994, 0.9905, 0.9924, 0.9929, 0.9927], week: [0.995, 0.9941, 0.9957, 0.9947, 0.9944, 0.9948, 0.996, 0.9959, 0.9962, 0.9948, 0.9947, 0.994, 0.9928, 0.9894, 0.9895, 0.9903, 0.9922, 0.9931, 0.9917, 0.9923, 0.9926, 0.9928, 0.9926, 0.9948, 0.992, 0.9917, 0.993, 0.9928], day: [0.9954, 0.9947, 0.996, 0.9979, 0.9958, 0.9836, 0.9957, 0.9956, 0.9965, 0.9965, 0.9963, 0.9956, 0.9943, 0.9948, null, 0.9944, 0.9954, 0.9949, 0.9948, 0.994, 0.9945, null, 0.9947, 0.9935, 0.9947, 0.9947, 0.9943, 0.9943, null, 0.9949, 0.9948, 0.9947, 0.995, 0.995, 0.9946, 0.9948, 0.9955, 0.9962, 0.9966, 0.9961, 0.9957, 0.9959, 0.9953, 0.9968, 0.9981, 0.9964, 0.9955, null, 0.9953, 0.9938, null, 0.9942, 0.9948, 0.9956, 0.9948, 0.9956, 0.9945, 0.9946, 0.9943, 0.994, 0.9953, 0.9954, 0.9937, 0.9941, 0.9934, 0.9946, 0.9942, 0.9956, 0.9928, 0.9932, 0.9929, 0.9925, 0.9931, 0.993, 0.9922, null, 0.9925, 0.9923, 0.9919, 0.988, 0.9858, 0.9869, 0.9884, 0.9892, 0.9909, 0.9919, 0.9898, 0.9878, 0.9878, null, 0.9895, 0.9903, 0.989, 0.9895, 0.9917, 0.9932, 0.9887, 0.9919, 0.9927, 0.9928, 0.994, 0.9927, 0.9892, null, 0.9924, 0.9927, 0.9957, 0.9929, 0.9918, 0.9908, 0.9912, 0.9907, 0.9921, 0.9915, 0.994, null, 0.9942, 0.9922, 0.9924, 0.9908, 0.9915, 0.9928, null, 0.993, 0.9935, 0.9898, 0.9919, 0.9944, 0.9929, 0.9926, 0.9933, 0.9923, 0.992, 0.9931, 0.9939, 0.9928, 0.9926, 0.9918, 0.9932, 0.9921, 0.9929, 0.995, 0.9924, 0.9957, 0.9954, 0.9951, 0.9953, 0.995, 0.9953, 0.9937, 0.9937, 0.9937, 0.9805, 0.9943, 0.993, 0.9917, 0.989, 0.9921, 0.9899, 0.9932, 0.9949, 0.9922, 0.9921, 0.9922, 0.9935, 0.9944, 0.9921, 0.9918, 0.9931] },
    RTY_SUB2: { month: [0.9898, 0.985, 0.9827, 0.9783, 0.9666, 0.9751, 0.9754], week: [0.9901, 0.9892, 0.9888, 0.9897, 0.991, 0.9869, 0.9885, null, 0.9779, 0.9848, 0.9875, 0.984, 0.9795, 0.9704, 0.9738, 0.9809, 0.984, 0.9831, 0.9682, 0.9679, 0.9689, 0.9557, 0.9718, 0.9802, 0.975, 0.9742, 0.9722, 0.9788], day: [0.9899, 0.9903, 0.9875, 0.9894, 0.9925, 0.9913, 0.99, 0.986, 0.9911, 0.9874, 0.9874, 0.9888, 0.9868, 0.9915, null, 0.9868, 0.9936, 0.9899, 0.9899, 0.9898, 0.9904, null, 0.9923, 0.9905, 0.9922, 0.9936, 0.9909, 0.9867, null, 0.9866, 0.991, 0.9787, 0.9908, 0.9866, 0.9857, 0.9892, 0.9906, 0.9901, 0.9922, 0.994, 0.9759, null, 0.9909, 0.989, 0.9866, 0.9862, 0.9776, 0.9525, 0.9894, 0.9855, 0.9888, 0.984, 0.9918, 0.9877, 0.9864, 0.972, 0.9865, 0.9902, 0.9893, 0.9879, 0.9885, 0.9825, 0.9799, 0.9859, 0.985, 0.982, 0.9861, 0.9836, 0.9857, 0.9775, 0.979, 0.9811, 0.9801, 0.9799, 0.9795, null, 0.9782, 0.9628, 0.9592, 0.9692, 0.9729, 0.9732, 0.9775, 0.9823, 0.9709, 0.9764, 0.9765, 0.9683, 0.9683, null, 0.9748, 0.983, 0.9816, 0.9815, 0.9826, 0.9812, 0.9863, 0.9818, 0.9826, 0.9833, 0.9834, 0.9844, 0.9882, null, 0.9802, 0.9852, 0.9837, 0.9863, 0.9802, 0.98, 0.976, 0.9681, 0.9613, 0.9581, 0.9648, 0.983, 0.9682, 0.9625, 0.9678, 0.9683, 0.9711, 0.9697, null, 0.9752, 0.9757, 0.9736, 0.9562, 0.9726, 0.9605, 0.9537, 0.9306, 0.9691, 0.9478, 0.9665, 0.967, 0.9661, 0.9753, 0.9679, 0.9769, 0.9723, 0.9723, 0.9789, 0.9832, 0.9825, 0.9754, 0.9804, 0.981, 0.9774, 0.9784, 0.9727, 0.9698, 0.9703, 0.9816, 0.9755, 0.9748, 0.9761, 0.9721, 0.9736, 0.9733, 0.9715, 0.9739, 0.972, 0.965, 0.9737, 0.9771, 0.9742, 0.9783, 0.9831, 0.9798] },
    MAIN_FVI: { month: [0.9961, 0.9964, 0.9962, 0.9966, 0.9962, 0.9973, 0.9974], week: [0.9961, 0.9966, 0.9963, 0.9956, 0.9957, 0.9964, 0.9965, 0.997, 0.9963, 0.9965, 0.9971, 0.9959, 0.9958, 0.9957, 0.9969, 0.9965, 0.9968, 0.997, 0.9966, 0.9956, 0.997, 0.9953, 0.9967, 0.9975, 0.9972, 0.9976, 0.9976, 0.9974], day: [0.9958, 0.9965, 0.9969, 0.996, 0.9965, 0.9971, 0.9967, 0.9966, 0.9963, 0.9966, 0.9961, 0.9962, 0.9964, 0.9963, null, 0.9969, 0.9962, 0.9956, 0.9957, 0.9944, 0.9948, null, 0.9949, 0.9958, 0.997, 0.9964, 0.9951, 0.9952, null, 0.9943, 0.9956, 0.996, 0.9975, 0.9971, 0.9969, 0.9974, 0.9971, 0.9967, 0.997, 0.997, 0.9945, 0.997, 0.9968, 0.9963, 0.997, 0.9967, 0.9969, 0.9947, 0.9961, 0.9976, null, 0.9971, 0.9974, 0.9971, 0.9984, 0.9912, 0.997, 0.9975, 0.9964, 0.997, 0.9969, 0.998, 0.998, 0.9965, 0.9958, 0.9952, 0.9957, 0.9949, 0.9954, 0.9959, 0.9964, 0.9957, 0.9948, 0.9963, 0.9961, 0.9952, 0.9956, 0.9949, 0.9954, 0.9952, 0.9959, 0.9966, 0.9963, 0.9963, 0.9967, 0.997, 0.9976, 0.9967, 0.9966, 0.9973, 0.9958, 0.9966, 0.9966, 0.9972, 0.9964, 0.9968, 0.996, 0.9975, 0.9962, 0.9972, 0.9969, 0.9964, 0.9966, null, 0.9967, 0.9965, 0.9971, 0.9968, 0.9978, 0.9969, 0.9966, 0.9964, 0.9966, 0.9964, 0.9966, 0.9968, 0.9969, 0.9966, 0.9929, 0.9959, 0.9957, 0.9957, 0.9955, 0.9966, 0.996, 0.9985, null, null, null, null, null, null, 0.9958, 0.994, 0.9961, 0.9939, 0.9966, 0.9983, 0.9962, 0.997, 0.9983, 0.9978, 0.9982, 0.9984, 0.9963, 0.9968, 0.9973, 0.9969, 0.9972, 0.9969, 0.9978, 0.9965, 0.9976, 0.998, 0.9973, 0.9957, 0.9979, 0.998, 0.9985, 0.9982, 0.9979, 0.9974, 0.9974, 0.9969, 0.9976, 0.9978, 0.9975, 0.997, 0.9974] },
    MAIN_ASSY: { month: [0.997, 0.9958, 0.9945, 0.9956, 0.9955, 0.9962, 0.9955], week: [0.9966, 0.9976, 0.9972, 0.9972, 0.9961, 0.9958, 0.9961, 0.9947, 0.9963, 0.9955, 0.9961, 0.995, 0.9919, 0.9929, 0.996, 0.9963, 0.9958, 0.996, 0.9958, 0.9961, 0.9939, 0.994, 0.9969, 0.9967, 0.994, 0.9978, 0.9959, 0.9947], day: [0.9948, 0.9984, 0.9979, 0.9973, 0.9974, 0.9966, 0.9987, 0.9978, 0.9977, 0.9975, 0.9951, 0.9979, 0.9973, 0.9979, null, 0.9984, 0.9973, 0.9982, 0.9974, 0.9962, 0.9959, null, 0.997, 0.9953, 0.9958, 0.9969, 0.9957, 0.9958, null, 0.9956, 0.9939, 0.9927, 0.9984, 0.9984, 0.9947, 0.9966, 0.996, 0.9968, 0.9981, 0.9938, 0.9957, 0.9947, 0.9979, 0.9978, 0.9974, 0.9983, 0.996, 0.9883, 0.9981, null, null, 0.9971, 0.998, 0.997, 0.9891, 0.9966, 0.9972, 0.997, 0.9962, 0.994, 0.996, 0.996, 0.9971, 0.9935, 0.9965, 0.9958, 0.9959, 0.9958, 0.9902, 0.9947, 0.9875, 0.9924, 0.9916, 0.9929, 0.9921, null, 0.9921, 0.9899, 0.993, 0.9931, 0.9934, 0.9933, 0.9954, 0.9966, 0.9963, 0.995, 0.997, 0.9955, 0.9966, 0.995, 0.9956, 0.9973, 0.9969, 0.9969, 0.9947, 0.9959, 0.9966, 0.9968, 0.9968, 0.9952, 0.9958, 0.9961, 0.9945, null, 0.9962, 0.9946, 0.9963, 0.9976, 0.9953, 0.9977, 0.9962, 0.9963, 0.9953, 0.9956, 0.9924, 0.997, 0.9966, 0.9961, 0.9966, 0.9967, 0.9947, 0.996, null, 0.9939, null, null, null, null, null, null, null, 0.9954, 0.9893, 0.9935, 0.9977, 0.9973, 0.9958, 0.9978, 0.9956, 0.9979, 0.9972, 0.9972, 0.9962, 0.9963, 0.9961, 0.998, 0.9963, 0.9938, 0.9953, 0.9977, 0.9966, 0.9854, 0.995, 0.9984, 0.9975, 0.998, 0.9979, 0.997, 0.9979, 0.9981, 0.9917, 0.9964, 0.9963, 0.9952, 0.9975, 0.9957, 0.9951, 0.9936, 0.9943] },
    MAIN_DRIVING: { month: [0.9915, 0.9926, 0.9958, 0.9934, 0.9945, 0.9949, 0.9954], week: [0.9955, 0.9954, 0.9904, 0.9858, 0.9939, 0.9961, 0.9965, 0.9969, 0.9861, 0.9956, 0.9958, 0.9963, 0.9958, 0.9948, 0.9935, 0.9936, 0.9942, 0.9913, 0.9947, 0.994, 0.9957, 0.9942, 0.9947, 0.9945, 0.9948, 0.9953, 0.9954, 0.9954], day: [0.9956, 0.9955, 0.9955, 0.9953, 0.9955, 0.9952, 0.9955, 0.9956, 0.9954, 0.9955, 0.985, 0.9926, 0.9826, 0.9916, null, 0.9957, 0.993, 0.989, 0.993, 0.9589, 0.9759, 0.9952, 0.9834, 0.9952, 0.9955, 0.9965, 0.9964, 0.9963, null, 0.9961, 0.996, 0.9964, 0.9962, 0.9958, 0.9961, 0.9959, 0.9967, 0.9965, 0.9965, 0.9962, 0.9965, 0.9969, 0.9803, 0.9866, 0.9726, 0.9813, 0.9945, 0.9926, 0.9949, 0.9957, null, 0.9947, 0.9951, 0.9955, 0.9963, 0.9962, 0.996, 0.9961, 0.9964, 0.9967, 0.9966, 0.993, 0.9963, 0.9965, 0.9965, 0.9961, 0.9963, 0.9962, 0.9959, 0.9958, 0.9957, 0.9957, 0.996, 0.9961, 0.9954, null, 0.996, 0.9956, 0.9953, 0.9947, 0.9944, 0.9945, 0.9928, 0.9939, 0.9938, 0.994, 0.9944, 0.9938, 0.9919, 0.9928, 0.9947, 0.994, 0.9927, 0.994, 0.9938, 0.993, 0.9929, 0.9925, 0.9948, 0.9942, 0.9937, 0.9946, 0.9952, null, 0.9934, 0.9817, 0.993, 0.9941, 0.9943, 0.9947, 0.9943, 0.9947, 0.9943, 0.9958, 0.9946, 0.9944, 0.9943, 0.9944, 0.9934, 0.9938, 0.9938, 0.9944, 0.9941, 0.9941, 0.996, 0.997, null, null, null, null, null, null, 0.9943, 0.9949, 0.9934, 0.9941, 0.9944, 0.9937, 0.9951, 0.9953, 0.9956, 0.9947, 0.9949, 0.9948, 0.9943, 0.9947, 0.9938, 0.9938, 0.9949, 0.9942, 0.9953, 0.9956, 0.995, 0.9958, 0.9949, 0.995, 0.9951, 0.9952, 0.9956, 0.9953, 0.9951, 0.9954, 0.9954, 0.9954, 0.9955, 0.9957, 0.9949, 0.9957, 0.9954] },
    MAIN_TILT: { month: [0.9986, 0.9992, 0.9993, 0.999, 0.999, 0.999, 0.9991], week: [0.9992, 0.9972, 0.9992, 0.9986, 0.9992, 0.9992, 0.9993, 0.9994, 0.9991, 0.9993, 0.9993, 0.9993, 0.9993, 0.9992, 0.9991, 0.999, 0.9989, 0.9986, 0.9991, 0.9991, 0.9991, 0.9988, 0.9988, 0.9991, 0.9991, 0.999, 0.999, 0.9991], day: [0.9992, 0.9992, 0.9993, 0.9979, 0.9886, 0.9991, 0.9991, 0.9993, 0.9992, 0.9991, 0.9992, 0.9992, 0.9991, 0.9994, null, 0.9993, 0.9993, 0.9992, 0.9967, 0.9967, 0.9992, 1, 0.9992, 0.9992, 0.9992, 0.9994, 0.9992, 0.9991, null, 0.9991, 0.9992, 0.9992, 0.9991, 0.9993, 0.9993, 0.9991, 0.9994, 0.9992, 0.9994, 0.9995, 0.9992, 0.9994, 0.999, 0.9989, 0.9991, 0.9991, 0.9993, 0.9989, 0.9991, 1, null, 0.9992, 0.9992, 0.9992, 0.9992, 0.9993, 0.9994, 0.9993, 0.9993, 0.9993, 0.9992, 0.9993, 0.9993, 0.9994, 0.9994, 0.9993, 0.9994, 0.9994, 0.9992, 0.9993, 0.9994, 0.9993, 0.9993, 0.9993, 0.9993, null, 0.9993, 0.9993, 0.9992, 0.9992, 0.999, 0.9992, 0.9991, 0.9991, 0.9991, 0.9991, 0.999, 0.9993, 0.999, 0.9992, 0.999, 0.9992, 0.9992, 0.9991, 0.9991, 0.9987, 0.9986, 0.9984, 0.9988, 0.9991, 0.9991, 0.999, 0.9989, null, 0.9988, 0.9986, 0.9984, 0.9983, 0.9988, 0.9989, 0.9991, 0.9991, 0.9991, 0.9992, 0.9992, 0.999, 0.9991, 0.999, 0.9991, 0.9991, 0.999, 0.9989, 0.9993, 0.9992, 0.9993, 0.999, null, null, null, null, null, null, 0.9987, 0.9991, 0.9987, 0.9989, 0.9987, 0.9981, 0.9989, 0.999, 0.999, 0.9991, 0.9993, 0.9991, 0.9991, 0.999, 0.999, 0.9989, 0.999, 0.999, 0.9992, 0.9991, 0.9991, 0.9991, 0.999, 0.9988, 0.9989, 0.999, 0.999, 0.999, 0.999, 0.999, 0.9989, 0.9991, 0.9992, 0.9991, 0.9991, 0.9992, 0.999] },
    SUB1_FPCB: { month: [0.9975, 0.9978, 0.9972, 0.997, 0.9971, 0.997, 0.9969], week: [0.9974, 0.9979, 0.9978, 0.9972, 0.997, 0.9972, 0.9981, 0.9981, 0.9981, 0.9974, 0.9974, 0.9972, 0.9968, 0.9966, 0.9964, 0.9973, 0.9976, 0.9972, 0.997, 0.9974, 0.997, 0.9968, 0.9967, 0.9975, 0.9972, 0.9965, 0.9971, 0.9969], day: [0.9976, 0.9972, 0.9979, 0.9988, 0.9977, 0.998, 0.9975, 0.9976, 0.998, 0.998, 0.998, 0.9976, 0.9974, 0.9974, null, 0.997, 0.9974, 0.9973, 0.9973, 0.9969, 0.9971, null, 0.9972, 0.9966, 0.997, 0.9973, 0.9969, 0.9968, null, 0.9971, 0.9972, 0.9971, 0.9974, 0.9972, 0.997, 0.9973, 0.9978, 0.9982, 0.9983, 0.9979, 0.998, 0.9981, 0.9979, 0.9983, 0.9988, 0.9982, 0.9978, null, 0.9979, 0.9971, null, 0.9973, 0.9975, 0.9976, 0.9973, 0.9978, 0.9972, 0.9975, 0.9972, 0.9969, 0.9975, 0.9978, 0.9972, 0.9972, 0.997, 0.9976, 0.9974, 0.9972, 0.9969, 0.9969, 0.9969, 0.9967, 0.9969, 0.997, 0.9965, null, 0.9968, 0.9965, 0.9968, 0.9964, 0.9966, 0.9963, 0.9964, 0.9965, 0.9966, 0.9963, 0.9963, 0.9964, 0.9964, null, 0.9964, 0.9969, 0.9964, 0.997, 0.9977, 0.9975, 0.9989, 0.9971, 0.9974, 0.9978, 0.9982, 0.9977, 0.9974, null, 0.9974, 0.9968, 0.9972, 0.9973, 0.9976, 0.9969, 0.9969, 0.9966, 0.9975, 0.9966, 0.9978, null, 0.9981, 0.9967, 0.9972, 0.9977, 0.9965, 0.998, null, 0.9976, 0.998, 0.9964, 0.9962, 0.997, 0.9966, 0.9964, 0.9969, 0.9964, 0.9966, 0.9972, 0.9972, 0.9971, 0.9968, 0.9965, 0.9968, 0.9964, 0.9968, 0.9978, 0.997, 0.9978, 0.9974, 0.9974, 0.9976, 0.9976, 0.9975, 0.9971, 0.9969, 0.9971, 0.9968, 0.9971, 0.9964, 0.9965, 0.9962, 0.9965, 0.9965, 0.9968, 0.9982, 0.9968, 0.9969, 0.9967, 0.9974, 0.9975, 0.9967, 0.9967, 0.9967] },
    SUB1_FVI: { month: [0.9973, 0.9979, 0.9968, 0.9935, 0.9953, 0.9958, 0.9957], week: [0.9976, 0.9962, 0.9979, 0.9975, 0.9974, 0.9976, 0.9979, 0.9978, 0.9981, 0.9974, 0.9973, 0.9968, 0.996, 0.9928, 0.9931, 0.993, 0.9946, 0.9959, 0.9947, 0.9949, 0.9956, 0.996, 0.9958, 0.9973, 0.9948, 0.9951, 0.9959, 0.9959], day: [0.9978, 0.9974, 0.998, 0.999, 0.9981, 0.9856, 0.9982, 0.998, 0.9985, 0.9985, 0.9982, 0.998, 0.9969, 0.9974, null, 0.9973, 0.998, 0.9976, 0.9975, 0.9972, 0.9974, null, 0.9974, 0.9969, 0.9977, 0.9973, 0.9974, 0.9975, null, 0.9978, 0.9976, 0.9976, 0.9976, 0.9978, 0.9975, 0.9975, 0.9976, 0.998, 0.9983, 0.9981, 0.9976, 0.9978, 0.9974, 0.9985, 0.9993, 0.9982, 0.9976, null, 0.9974, 0.9967, null, 0.9969, 0.9973, 0.998, 0.9975, 0.9978, 0.9973, 0.997, 0.9971, 0.9971, 0.9978, 0.9976, 0.9965, 0.9968, 0.9963, 0.9969, 0.9967, 0.9985, 0.9959, 0.9963, 0.996, 0.9958, 0.9962, 0.996, 0.9957, null, 0.9957, 0.9958, 0.9951, 0.9916, 0.9891, 0.9905, 0.9919, 0.9926, 0.9943, 0.9955, 0.9935, 0.9914, 0.9914, null, 0.9931, 0.9934, 0.9925, 0.9925, 0.994, 0.9957, 0.9898, 0.9947, 0.9953, 0.995, 0.9958, 0.9949, 0.9917, null, 0.995, 0.9959, 0.9986, 0.9956, 0.9942, 0.9938, 0.9944, 0.9941, 0.9947, 0.9949, 0.9962, null, 0.996, 0.9955, 0.9952, 0.9931, 0.995, 0.9948, null, 0.9954, 0.9954, 0.9934, 0.9957, 0.9973, 0.9963, 0.9961, 0.9964, 0.9959, 0.9954, 0.9958, 0.9967, 0.9957, 0.9957, 0.9954, 0.9964, 0.9957, 0.9961, 0.9972, 0.9954, 0.9978, 0.9979, 0.9977, 0.9977, 0.9974, 0.9978, 0.9966, 0.9968, 0.9966, 0.9836, 0.9973, 0.9967, 0.9952, 0.9927, 0.9957, 0.9933, 0.9964, 0.9967, 0.9954, 0.9951, 0.9955, 0.996, 0.9968, 0.9954, 0.9951, 0.9963] },
    SUB2_HOOK: { month: [0.9998, 0.9998, 0.9996, 0.9994, 0.9993, 0.9993, 0.9993], week: [0.9999, 0.9998, 0.9998, 0.9999, 0.9997, 0.9998, 0.9999, null, 0.9997, 0.9997, 0.9997, 0.9995, 0.9994, 0.9991, 0.9993, 0.9996, 0.9995, 0.9998, 0.9996, 0.9993, 0.9991, 0.9992, 0.9991, 0.9992, 0.9993, 0.9994, 0.9993, 0.9995], day: [0.9998, 1.0, 0.9998, 1, 0.9999, 1, 0.9997, 0.9992, 0.9999, 0.9997, 0.9997, 0.9998, 0.9998, 0.9998, null, 0.9997, 1, 1, 0.9999, 1, 1, null, 0.9997, 0.9996, 1, 0.9999, 0.9998, 0.9996, null, 1, 1, 0.9994, 1, 1, 0.9998, 0.9998, 1, 0.9997, 0.9997, 0.9998, 1, null, 0.9998, 0.9997, 0.9997, 0.9996, 0.9994, 1, null, 1, 1, 1, 0.9997, 0.9996, 0.9996, 0.9987, 0.9998, 0.9995, 0.9999, 0.9997, 0.9999, 0.9993, 0.9995, 0.9996, 0.9995, 0.9994, 0.9996, 0.9993, 0.9999, 0.9996, 0.9995, 0.9997, 0.9995, 0.9991, 0.9992, null, 0.9994, 0.999, 0.9991, 0.999, 0.999, 0.9992, 0.9991, 0.9995, 0.9993, 0.9992, 0.9993, 0.9992, 0.9992, null, 0.9997, 0.9995, 0.9994, 0.9995, 0.9992, 0.9998, 1, 0.9995, 0.9993, 0.9993, 0.9995, 0.9994, 1, null, 0.9995, 0.9997, 0.9997, 0.9999, 0.9999, 0.9997, 0.9997, 0.9998, 0.9996, 0.9994, 0.9994, null, 0.9996, 0.9992, 0.9991, 0.9994, 0.9993, 0.9995, null, 0.9993, 0.9996, 0.999, 0.9986, 0.9991, 0.9987, 0.998, 0.999, 0.9995, 0.9989, 0.9997, 1, 0.9993, 0.9988, 0.9996, 0.9986, 0.9996, 0.999, 0.9993, 0.9997, 0.9991, 0.9992, 0.9989, 0.9992, 0.999, 0.9992, 0.9993, 0.9992, 0.9999, 0.9992, 0.9991, 0.9995, 0.9998, 0.9986, 0.9994, 0.9998, 0.9988, 1, 0.999, 0.999, 0.9988, 0.9999, 0.9993, 0.9994, 0.9998, 0.9996] },
    SUB2_OVEN: { month: [1.0, 1.0, 1, 1, 1, 1, 1], week: [1, 1, 1, 1, 1.0, 1.0, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1], day: [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 0.9999, 1, null, 1, 1, 1, 1, 1, 0.9997, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1] },
    SUB2_INDEX: { month: [0.9962, 0.9956, 0.9944, 0.9929, 0.9905, 0.9917, 0.9917], week: [0.9964, 0.9962, 0.9955, 0.9962, 0.9968, 0.9953, 0.9968, null, 0.9949, 0.994, 0.9968, 0.9943, 0.9928, 0.9922, 0.9915, 0.9936, 0.9951, 0.9931, 0.9908, 0.9913, 0.9917, 0.987, 0.9944, 0.9936, 0.9893, 0.9906, 0.9887, 0.9946], day: [0.9969, 0.9959, 0.994, 0.9971, 0.9973, 0.9975, 0.9962, 0.9949, 0.9974, 0.9952, 0.995, 0.9959, 0.9931, 0.9963, null, 0.9971, 0.9972, 0.9973, 0.9964, 0.9945, 0.9949, null, 0.998, 0.9965, 0.9969, 0.9989, 0.9969, 0.9937, null, 0.9964, 0.9978, 0.9963, 0.9942, 0.9906, 0.9946, 0.9969, 0.9954, 0.9972, 0.9972, 0.9982, 0.9959, null, 0.9975, 0.9957, 0.9969, 0.996, 0.9986, 0.9848, null, 0.9956, 0.9936, 0.9938, 0.998, 0.9947, 0.9971, 0.9855, 0.9956, 0.9981, 0.9985, 0.998, 0.9963, 0.9941, 0.9912, 0.9967, 0.9949, 0.9925, 0.996, 0.9944, 0.9946, 0.9913, 0.9911, 0.9949, 0.9939, 0.992, 0.9933, null, 0.994, 0.995, 0.9896, 0.9918, 0.9902, 0.991, 0.9938, 0.9959, 0.9904, 0.9934, 0.992, 0.9886, 0.9886, null, 0.9908, 0.9937, 0.9953, 0.9962, 0.993, 0.992, 0.994, 0.9933, 0.995, 0.995, 0.9943, 0.9952, 0.9975, null, 0.9914, 0.9936, 0.9921, 0.9958, 0.9924, 0.9947, 0.994, 0.9911, 0.9834, 0.9898, 0.9918, null, 0.9933, 0.9926, 0.9915, 0.9901, 0.9924, 0.9881, null, 0.9916, 0.9955, 0.9942, 0.9852, 0.9895, 0.9943, 0.9899, 0.9681, 0.9949, 0.9889, 0.9894, 0.9907, 0.9876, 0.9967, 0.9937, 0.9976, 0.9964, 0.9941, 0.9966, 0.9955, 0.9949, 0.9899, 0.9914, 0.9935, 0.9918, 0.9924, 0.9883, 0.9851, 0.9851, 0.993, 0.9918, 0.9913, 0.9889, 0.991, 0.9896, 0.991, 0.991, 0.9858, 0.9905, 0.9845, 0.9888, 0.9917, 0.993, 0.9936, 0.9972, 0.9947] },
  },
  'SO1C2EH': {
    RTY_TTL: { month: [null, null, 0.9717, 0.9577, 0.9632, 0.9615, 0.9529], week: [null, null, null, null, null, null, null, null, null, null, 0.9871, null, 0.9639, null, 0.9568, 0.9619, null, null, null, 0.9848, 0.9554, null, null, 0.97, 0.9621, 0.9609, 0.9568, 0.9474], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9871, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9758, 0.9547, null, null, null, null, null, null, null, null, 0.9568, null, null, null, null, null, null, null, 0.9599, 0.9596, 0.9808, null, 1, 0.9944, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9848, null, null, null, 0.969, null, 0.9859, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 0.9525, 0.9796, 0.9842, null, 0.9779, 0.9617, 0.9696, 0.9681, 0.9598, 0.9523, 0.9598, 0.9675, 0.9653, 0.9527, 0.9651, 0.955, 0.9514, 0.9565, 0.9508, 0.9599, 0.9625, 0.9602, 0.9586, 0.9485, 0.952, 0.9309] },
    RTY_MAIN: { month: [null, null, 0.9717, 0.9577, 0.9632, 0.9819, 0.9744], week: [null, null, null, null, null, null, null, null, null, null, 0.9871, null, 0.9639, null, 0.9568, 0.9619, null, null, null, 0.9848, 0.9554, null, null, 0.9899, 0.9829, 0.9808, 0.9773, 0.9705], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9871, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9758, 0.9547, null, null, null, null, null, null, null, null, 0.9568, null, null, null, null, null, null, null, 0.9599, 0.9596, 0.9808, null, 1, 0.9944, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9848, null, null, null, 0.969, null, 0.9859, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 0.9746, null, null, null, null, 0.9836, 0.9886, 0.9887, 0.9802, 0.9735, 0.982, 0.9874, 0.987, 0.9753, 0.9791, 0.9739, 0.9732, 0.9775, 0.9773, 0.9771, 0.9812, 0.9779, 0.9803, 0.97, 0.9721, 0.9596] },
    RTY_SUB1: { month: [null, null, null, null, null, 0.9951, 0.995], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9966, 0.9943, 0.9956, 0.9953, 0.9943], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9958, null, 0.9957, null, 0.996, 0.9936, 0.995, 0.9954, 0.9934, 0.9921, 0.9956, 0.9959, 0.9968, 0.9929, 0.9963, 0.9964, 0.9954, 0.9935, 0.9953, 0.9964, 0.9952, 0.9962, 0.9949, 0.9936, 0.9942, 0.9944] },
    RTY_SUB2: { month: [null, null, null, null, null, 0.984, 0.9828], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9832, 0.9845, 0.984, 0.9837, 0.9819], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9815, 0.9796, 0.9885, null, 0.9818, 0.9839, 0.9856, 0.9837, 0.9858, 0.986, 0.9817, 0.984, 0.9812, 0.9838, 0.9894, 0.9841, 0.9821, 0.985, 0.9775, 0.986, 0.9856, 0.9857, 0.9829, 0.9841, 0.985, 0.9755] },
    MAIN_FVI: { month: [null, null, null, 0.988, 0.9859, 0.9944, 0.9934], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.988, null, null, null, null, 0.9859, null, null, 0.997, 0.9938, 0.9944, 0.9932, 0.9935], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9899, 0.9861, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9859, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 0.994, null, null, null, null, 0.9938, 0.9962, 0.9952, 0.9948, 0.9891, 0.9938, 0.9962, 0.9936, 0.9945, 0.9949, 0.9935, 0.9946, 0.9912, 0.9938, 0.9936, 0.993, 0.9928, 0.9924, 0.9945, 0.9932, 0.994] },
    MAIN_ASSY: { month: [null, null, null, 0.9946, 0.9947, 0.9951, 0.9917], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9912, 0.9957, null, null, null, 0.9977, 0.9916, null, null, 0.9947, 0.9982, 0.9945, 0.9928, 0.9892], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9912, null, null, null, null, null, null, null, 0.9931, 0.9941, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9977, null, null, null, 0.9916, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 0.984, null, null, null, null, 0.9977, 0.9993, 0.999, 0.9975, 0.9975, 0.9968, 0.9977, 0.9977, 0.9872, 0.9953, 0.9921, 0.9841, 0.9963, 0.9929, 0.9908, 0.9973, 0.9955, 0.9968, 0.986, 0.9917, 0.9822] },
    MAIN_DRIVING: { month: [null, null, 0.9741, 0.9766, 0.9822, 0.9926, 0.9894], week: [null, null, null, null, null, null, null, null, null, null, 0.9889, null, 0.9666, null, 0.9657, 0.9803, null, null, null, 0.9871, 0.9772, null, null, 0.9985, 0.9911, 0.9919, 0.9915, 0.9876], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9889, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9758, 0.9574, null, null, null, null, null, null, null, null, 0.9657, null, null, null, null, null, null, null, 0.968, 0.9756, null, null, null, 0.9972, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9871, null, null, null, 0.9772, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 0.997, null, null, null, null, 0.9925, 0.9933, 0.9947, 0.9881, 0.9867, 0.9914, 0.9937, 0.9958, 0.9935, 0.9889, 0.9881, 0.9943, 0.9904, 0.9914, 0.9928, 0.9909, 0.9895, 0.991, 0.9893, 0.987, 0.983] },
    MAIN_TILT: { month: [null, null, 0.9976, 0.9979, 1, 0.9998, 0.9999], week: [null, null, null, null, null, null, null, null, null, null, 0.9981, null, 0.9972, null, 0.9995, 0.9975, null, null, null, 1, 1, null, null, 0.9997, 0.9997, 0.9999, 0.9999, 0.9999], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9981, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9972, null, null, null, null, null, null, null, null, 0.9995, null, null, null, null, null, null, null, 0.9985, 0.9995, 0.9947, null, null, 0.9972, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, null, null, 1, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, null, 0.9995, null, null, null, null, 0.9995, 0.9997, 0.9997, 0.9997, 1, 0.9999, 0.9998, 0.9998, 0.9999, 0.9999, 1, 1, 0.9995, 0.9999, 0.9999, 1, 1, 1, 0.9999, 0.9999, 0.9999] },
    SUB1_FPCB: { month: [null, null, null, null, null, 0.9976, 0.9976], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9983, 0.9972, 0.9979, 0.9977, 0.9972], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9975, null, 0.9973, null, 0.9984, 0.9969, 0.9974, 0.9972, 0.9968, 0.9964, 0.9978, 0.9982, 0.9988, 0.9965, 0.9981, 0.9981, 0.9976, 0.997, 0.9977, 0.9982, 0.9978, 0.9982, 0.9973, 0.9968, 0.9973, 0.9972] },
    SUB1_FVI: { month: [null, null, null, null, null, 0.9975, 0.9975], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9983, 0.9971, 0.9977, 0.9976, 0.9971], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9982, null, 0.9984, null, 0.9976, 0.9967, 0.9976, 0.9982, 0.9966, 0.9957, 0.9978, 0.9977, 0.9979, 0.9964, 0.9982, 0.9983, 0.9977, 0.9966, 0.9977, 0.9982, 0.9974, 0.9979, 0.9976, 0.9968, 0.997, 0.9971] },
    SUB2_HOOK: { month: [null, null, null, null, null, 0.9997, 0.9994], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9998, 0.9995, 0.9998, 0.9998, 0.9992], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 0.9999, 0.9996, null, 0.9995, 0.9987, 0.9998, 0.9999, 0.9996, 0.9995, 0.9997, 0.9999, 0.9999, 0.9997, 0.9999, 0.9998, 0.9999, 1, 0.9998, 0.9998, 0.9998, 0.9993, 0.9997, 1, 0.9997, 0.9974] },
    SUB2_OVEN: { month: [null, null, null, null, null, 1, 1], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, 1, 1], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 1, 1, 1, null, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1] },
    SUB2_INDEX: { month: [null, null, null, null, null, 0.9949, 0.9926], week: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9925, 0.9962, 0.9947, 0.9951, 0.9899], day: [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0.9898, 0.9922, 0.9954, null, 0.9945, 0.9962, 0.9958, 0.9949, 0.9974, 0.9985, 0.9971, 0.9953, 0.991, 0.9935, 0.9971, 0.9943, 0.9939, 0.9959, 0.9937, 0.9972, 0.9953, 0.9947, 0.9917, 0.9905, 0.9918, 0.9857] },
  },
};
/** Kiểu 1 dòng bảng chi tiết RTY theo Model / Process / Type / Thời gian.
 *  Dữ liệu các dòng giờ được SINH ĐỘNG theo `viewMode` (Ngày/Tuần/Tháng) từ
 *  `MODEL_SERIES` — cùng 1 nguồn thật với 4 biểu đồ ở Tab 1 (xem
 *  `buildTableRowsForMode` / `TABLE_ROWS_BY_MODE` bên dưới) — thay vì bảng
 *  tĩnh chỉ có mức Tháng như trước. KHÔNG có cột SL/Khách hàng (xem ghi
 *  chú đầu file). */
interface RtyTableRow { model: string; process: string; type: string; period: string; target: number; actual: number }


/** ═══════════════════════════════════════════════════════════════════════
 * FIX (table-not-following-day-week-month, EPCC)
 * ───────────────────────────────────────────────────────────────────────
 * EXPLORE: bảng "RTY & Chi tiết" (tab 2) đang đọc từ mảng TĨNH `TABLE_ROWS`
 * — vốn chỉ được trích xuất sẵn ở MỘT mức duy nhất (Tháng: JAN..JUL), hoàn
 * toàn TÁCH RỜI khỏi state `viewMode`. Trong khi đó nhóm nút "Xem theo"
 * (Ngày/Tuần/Tháng) chỉ được nối dây vào 4 biểu đồ ở Tab 1 (thông qua
 * `getSeriesValueForModel`/`MODEL_SERIES`) — đó là lý do bấm Ngày/Tuần/
 * Tháng KHÔNG làm bảng ở Tab 2 thay đổi như người dùng phản ánh.
 * PLAN: bảng cần đọc từ ĐÚNG 1 nguồn dữ liệu thật duy nhất với biểu đồ —
 * `MODEL_SERIES` (đã có sẵn đủ 3 mức Ngày/Tuần/Tháng cho cả 4 chỉ số
 * SUB1/SUB2/MAIN/TTL của toàn bộ 10 model, trích xuất từ Test4.xlsx) —
 * thay vì một bảng TABLE_ROWS chép tay riêng chỉ có mức Tháng. Sinh động
 * theo `viewMode` y hệt cách biểu đồ đang làm, mốc nào không có số liệu
 * thật (null) thì bỏ qua dòng đó (KHÔNG suy diễn/generate). Giữ nguyên thứ
 * tự nhóm Model theo hiệu suất (MODEL_SUMMARY_SORTED, xem yêu cầu trước)
 * và thứ tự Process/Type bên trong mỗi model (SUB1 → SUB2 → MAIN → TTL,
 * đúng thứ tự TABLE_ROWS gốc).
 * CODE: `SERIES_DEFS` ánh xạ 4 key MODEL_SERIES ↔ (process, type, target)
 * ĐÚNG như TABLE_ROWS gốc từng dùng; `buildTableRowsForMode()` sinh danh
 * sách dòng cho 1 mức; `TABLE_ROWS_BY_MODE` tính sẵn CẢ 3 mức 1 lần lúc
 * tải module (không tính lại mỗi lần render). Component chọn đúng mảng
 * theo `viewMode` hiện tại trong `filteredRows`.
 * CHECK: đã đối chiếu bằng script — mức Tháng sinh ra ĐÚNG 180 dòng, khớp
 * 100% (0 sai khác, 0 thiếu/thừa) với dữ liệu 180 dòng TABLE_ROWS chép tay
 * trước đây; mức Tuần sinh 575 dòng, mức Ngày sinh 3017 dòng — đều là số
 * liệu thật lấy thẳng từ MODEL_SERIES, không phát sinh số giả.
 * ═══════════════════════════════════════════════════════════════════════ */
const SERIES_DEFS: { key: string; process: string; type: string; target: number }[] = [
  { key: 'RTY_SUB1', process: 'SUB1', type: 'Sub1', target: TARGET_SUB1 },
  { key: 'RTY_SUB2', process: 'SUB2', type: 'Sub2', target: TARGET_SUB2 },
  { key: 'RTY_MAIN', process: 'MAIN', type: 'Main', target: TARGET_MAIN },
  { key: 'RTY_TTL',  process: 'MAIN', type: 'TTL',  target: TARGET_TTL },
];

const buildTableRowsForMode = (mode: 'day' | 'week' | 'month'): RtyTableRow[] => {
  const labels = mode === 'day' ? DAY_LABELS : mode === 'week' ? WEEK_LABELS : MONTH_LABELS;
  const rows: RtyTableRow[] = [];
  for (const ms of MODEL_SUMMARY_SORTED) {
    const series = MODEL_SERIES[ms.model];
    if (!series) continue;
    for (const def of SERIES_DEFS) {
      const arr = series[def.key]?.[mode];
      if (!arr) continue;
      labels.forEach((label, i) => {
        const v = arr[i];
        if (v == null) return;
        rows.push({ model: ms.model, process: def.process, type: def.type, period: label, target: def.target, actual: v });
      });
    }
  }
  return rows;
};

/** Tính sẵn cả 3 mức 1 lần lúc tải module — component chỉ cần chọn đúng
 *  mảng theo `viewMode` hiện tại, không phải tính lại trên mỗi lần render. */
const TABLE_ROWS_BY_MODE: Record<'day' | 'week' | 'month', RtyTableRow[]> = {
  day: buildTableRowsForMode('day'),
  week: buildTableRowsForMode('week'),
  month: buildTableRowsForMode('month'),
};

/** Lấy giá trị RTY THẬT cho 1 model bất kỳ (trong 10 model có ở dropdown)
 *  tại 1 mốc trên trục X, dùng chung cho cả 4 chart TTL/MAIN/SUB1/SUB2.
 *  Toàn bộ 10 model giờ đều đọc từ MODEL_SERIES — trích xuất trực tiếp từ
 *  Test4.xlsx giống hệt cách làm trước đây với SO3560, chỉ khác là áp dụng
 *  cho tất cả model thay vì 1 model. Model rỗng ("Tất cả") mặc định hiển
 *  thị theo model tham chiếu SO3560. Mốc nào Test4.xlsx không có số liệu
 *  → trả về null để ẩn hẳn điểm/cột đó (KHÔNG suy diễn / generate số giả). */
/** ═══════════════════════════════════════════════════════════════════════
 * FIX (rty-chart-blank-empty-dynamic-fallback, EPCC)
 * ───────────────────────────────────────────────────────────────────────
 * EXPLORE: biểu đồ trắng (không có cột/đường, chỉ còn đường Target hằng số)
 * dù đổi Ngày/Tuần/Tháng hay đổi Model — vì đường Target được vẽ độc lập,
 * không phụ thuộc getSeriesValueForModel(), nên vẫn hiện ngay cả khi hàm
 * này trả null ở MỌI lời gọi.
 * NGUYÊN NHÂN: dòng cũ `customSeriesMap?.[m] || MODEL_SERIES[m]` chỉ
 * fallback về dữ liệu tĩnh (demo, nhúng từ Test4.xlsx) khi model KHÔNG TỒN
 * TẠI trong dữ liệu động (customSeriesMap, tức activeModelSeries lấy từ
 * cache IndexedDB `rty_summary_dynamic_v2`). Nhưng khi bucket RTY trên
 * Supabase đang có 0 dòng (xem console "RTY: 0"), cache cũ trong
 * IndexedDB vẫn còn giữ object modelSeries có KEY cho từng model — chỉ là
 * toàn bộ mảng day/week/month bên trong đều là null. Object đó vẫn
 * TRUTHY nên toán tử `||` không bao giờ rơi xuống MODEL_SERIES tĩnh nữa,
 * khiến mọi cột/đường ở cả 4 panel (TTL/MAIN/SUB1/SUB2), mọi viewMode đều
 * vẽ null.
 * SỬA: tra theo dữ liệu động trước; nếu giá trị tại đúng key/mode/index là
 * null/undefined thì thử tiếp bằng dữ liệu tĩnh MODEL_SERIES cho model đó
 * (nếu có) trước khi thật sự trả về null — fallback theo TỪNG GIÁ TRỊ thay
 * vì theo TỪNG MODEL.
 * CHECK: khi Supabase RTY có dữ liệu thật trở lại, dynamicVal sẽ luôn ưu
 * tiên trước (không đổi hành vi khi dữ liệu động hợp lệ); chỉ rơi về tĩnh
 * khi dynamicVal thật sự thiếu.
 * ═══════════════════════════════════════════════════════════════════════ */
const getSeriesValueForModel = (
  model: string,
  key: string,
  mode: 'day' | 'week' | 'month',
  rawIndex: number,
  customSeriesMap?: Record<string, Record<string, { month: (number | null)[]; week: (number | null)[]; day: (number | null)[] }>>
): number | null => {
  const m = model || HERO_MODEL;

  const readFrom = (
    series?: Record<string, { month: (number | null)[]; week: (number | null)[]; day: (number | null)[] }>
  ): number | null | undefined => {
    if (!series || !series[key]) return undefined;
    const arr = (series[key] as any)?.[mode] as (number | null)[] | undefined;
    if (!arr) return undefined;
    return arr[rawIndex];
  };

  const dynamicVal = readFrom(customSeriesMap?.[m]);
  if (dynamicVal != null) return dynamicVal;

  const staticVal = readFrom(MODEL_SERIES[m]);
  return staticVal != null ? staticVal : null;
};

/** ═══════════════════════════════════════════════════════════════════════
 * FIX: biểu đồ mạng nhện (Radar/Spider) vẽ sai khi đổi Model
 * ───────────────────────────────────────────────────────────────────────
 * NGUYÊN NHÂN: getSeriesValueForModel() trả về null khi model không có số
 * liệu THẬT đúng ở mốc Ngày/Tuần đang chọn. Plotly scatterpolar coi null
 * là "không có giá trị" và vẽ điểm đó tụt về tâm (r=0) → đa giác bị kéo
 * lệch/"gãy" trông như lỗi dữ liệu, dù dữ liệu vẫn đúng.
 *
 * SỬA: khi mốc chi tiết không có số liệu, RƠI XUỐNG MỨC THÔ HƠN
 * (Ngày → Tháng tương ứng → Tổng hợp cả giai đoạn) thay vì vẽ null. Toàn
 * bộ giá trị dùng để rơi mức đều là số liệu THẬT đã có sẵn trong
 * MODEL_SERIES / MODEL_SUMMARY (trích xuất từ Test4.xlsx) — KHÔNG suy diễn
 * hay generate số giả ở bất kỳ bước nào.
 * ═══════════════════════════════════════════════════════════════════════ */
type SpiderSource = 'exact' | 'month' | 'summary' | 'none';
interface SpiderPoint { value: number | null; source: SpiderSource }

const SPIDER_SOURCE_LABEL: Record<SpiderSource, string> = {
  exact: '',
  month: '(rơi mức: TB Tháng)',
  summary: '(rơi mức: TB Tổng hợp)',
  none: '(không có số liệu)'
};

const getSpiderValueWithFallback = (
  model: string,
  processType: 'TTL' | 'MAIN',
  mode: 'day' | 'week' | 'month',
  rawIndex: number,
  currentLabel: string,
  customSummary?: ModelSummary[],
  customSeriesMap?: Record<string, Record<string, { month: (number | null)[]; week: (number | null)[]; day: (number | null)[] }>>
): SpiderPoint => {
  const seriesKey = processType === 'TTL' ? 'RTY_TTL' : 'RTY_MAIN';

  const exact = getSeriesValueForModel(model, seriesKey, mode, rawIndex, customSeriesMap);
  if (exact != null) return { value: exact, source: 'exact' };

  if (mode === 'day') {
    const mm = parseInt(currentLabel.split('/')[0], 10);
    const monthIdx = mm - 1;
    if (Number.isFinite(monthIdx) && monthIdx >= 0) {
      const monthVal = getSeriesValueForModel(model, seriesKey, 'month', monthIdx, customSeriesMap);
      if (monthVal != null) return { value: monthVal, source: 'month' };
    }
  }

  const list = customSummary || MODEL_SUMMARY;
  const summary = list.find(s => s.model === model);
  const agg = processType === 'TTL' ? summary?.ttl : summary?.main;
  if (agg != null) return { value: agg.actual, source: 'summary' };

  return { value: null, source: 'none' };
};

const TXT: Record<Lang, Record<string, string>> = {
  vi: {
    title: 'HIỆU SUẤT RTY',
    tab1: 'TÌNH HÌNH RTY',
    tab2: 'RTY & CHI TIẾT',
    startDate: 'NGÀY BẮT ĐẦU',
    endDate: 'NGÀY KẾT THÚC',
    customer: 'KHÁCH HÀNG',
    model: 'MODEL',
    viewBy: 'XEM THEO',
    day: 'Ngày', week: 'Tuần', month: 'Tháng',
    loadExcel: 'Tải tệp lên',
    allOption: 'Tất cả',
    kpi1: 'RTY Thực tế (TB)', kpi1Target: 'Mục tiêu',
    kpi2: 'Tỷ lệ đạt RTY',
    // ĐÃ ĐỔI: bỏ "Số lượng NG" (cần cột SL) → thay bằng số Model đạt mục tiêu
    kpi3: 'Model đạt mục tiêu',
    // ĐÃ ĐỔI: bỏ "Tỷ lệ NG/Tổng SL" (cần cột SL) → thay bằng chênh lệch trung bình
    kpi4: 'Chênh lệch TB (Actual − Target)',
    chart1: 'RTY TTL theo Model', chart2: 'Chênh lệch RTY theo Model',
    chart3: 'Xu hướng RTY (Model tham chiếu)', chart4: 'Tỷ lệ đạt mục tiêu chung',
    chart5: 'RTY theo Công đoạn (SUB1/SUB2/MAIN)', chart6: 'Top công đoạn RTY thấp nhất',
    noSource: 'Dữ liệu tham chiếu từ Test4.xlsx (chưa nối nguồn RTY thật)',
    waiting: 'Đang chờ dữ liệu',
    waitingDesc: 'Kết nối / tải file RTY để hiển thị số liệu tại đây',
    tblGroup2: 'RTY (%)',
    colModel: 'MODEL', colProcess: 'CÔNG ĐOẠN', colType: 'TYPE', colPeriod: 'THỜI GIAN',
    colRtyTarget: 'RTY MỤC TIÊU', colRtyActual: 'RTY THỰC TẾ', colRtyGap: 'CHÊNH LỆCH',
    tableEmpty: 'Chưa có dữ liệu RTY.',
    showing: 'Hiển thị',
    of: 'trong tổng',
    rows: 'dòng',
    rowsPerPage: 'Dòng/trang:',
    resetBtn: 'Đặt lại bộ lọc',
    kpiShowingModel: 'Đang hiển thị',
    kpiBestLabel: 'Hiệu suất tốt nhất',
    dataScopeNote: '⚠ Test4.xlsx chỉ có tỉ lệ RTY (%) — không có SL sản xuất/SL NG hay Khách hàng, nên các ô liên quan đã được lược bỏ.',
  },
  en: {
    title: 'RTY PERFORMANCE',
    tab1: 'RTY STATUS',
    tab2: 'RTY & DETAIL',
    startDate: 'START DATE',
    endDate: 'END DATE',
    customer: 'CUSTOMER',
    model: 'MODEL',
    viewBy: 'VIEW BY',
    day: 'Day', week: 'Week', month: 'Month',
    loadExcel: 'Upload file',
    allOption: 'All',
    kpi1: 'Actual RTY (avg)', kpi1Target: 'Target',
    kpi2: 'RTY Achievement Rate',
    kpi3: 'Models Meeting Target',
    kpi4: 'Avg Gap (Actual − Target)',
    chart1: 'RTY TTL by Model', chart2: 'RTY Gap by Model',
    chart3: 'RTY Trend (Reference Model)', chart4: 'Overall Achievement Rate',
    chart5: 'RTY by Process (SUB1/SUB2/MAIN)', chart6: 'Lowest RTY Process Steps',
    noSource: 'Reference data from Test4.xlsx (no live RTY source connected)',
    waiting: 'Waiting for data',
    waitingDesc: 'Connect / upload an RTY file to show figures here',
    tblGroup2: 'RTY (%)',
    colModel: 'MODEL', colProcess: 'PROCESS', colType: 'TYPE', colPeriod: 'PERIOD',
    colRtyTarget: 'RTY TARGET', colRtyActual: 'RTY ACTUAL', colRtyGap: 'GAP',
    tableEmpty: 'No RTY data yet.',
    showing: 'Showing',
    of: 'of',
    rows: 'rows',
    rowsPerPage: 'Rows/page:',
    resetBtn: 'Reset filters',
    kpiShowingModel: 'Showing',
    kpiBestLabel: 'Best performance',
    dataScopeNote: '⚠ Test4.xlsx only has RTY (%) ratios — no production/NG quantity or Customer column, so related cells were dropped.',
  },
  ko: {
    title: 'RTY 성과',
    tab1: 'RTY 현황',
    tab2: 'RTY & 상세',
    startDate: '시작일',
    endDate: '종료일',
    customer: '고객사',
    model: '모델',
    viewBy: '보기 방식',
    day: '일별', week: '주별', month: '월별',
    loadExcel: '파일 업로드',
    allOption: '전체',
    kpi1: '실제 RTY (평균)', kpi1Target: '목표',
    kpi2: 'RTY 달성률',
    kpi3: '목표 달성 모델 수',
    kpi4: '평균 차이 (실적 − 목표)',
    chart1: '모델별 RTY TTL', chart2: '모델별 RTY 차이',
    chart3: 'RTY 추세 (참조 모델)', chart4: '전체 달성률',
    chart5: '공정별 RTY (SUB1/SUB2/MAIN)', chart6: 'RTY 최저 공정',
    noSource: 'Test4.xlsx 참조 데이터 (실시간 RTY 소스 미연결)',
    waiting: '데이터 대기 중',
    waitingDesc: 'RTY 파일을 연결/업로드하면 여기에 수치가 표시됩니다',
    tblGroup2: 'RTY (%)',
    colModel: '모델', colProcess: '공정', colType: 'TYPE', colPeriod: '기간',
    colRtyTarget: 'RTY 목표', colRtyActual: 'RTY 실적', colRtyGap: '차이',
    tableEmpty: '아직 RTY 데이터가 없습니다.',
    showing: '표시',
    of: '/',
    rows: '행',
    rowsPerPage: '페이지당 행 수:',
    resetBtn: '필터 초기화',
    kpiShowingModel: '표시 중',
    kpiBestLabel: '최고 성과',
    dataScopeNote: '⚠ Test4.xlsx에는 RTY(%) 비율만 있고 생산/NG 수량이나 고객사 컬럼이 없어 관련 항목을 제외했습니다.',
  },
};

declare global {
  interface Window { Plotly?: any; }
}

interface RtyLegend {
  type: 'bar' | 'line' | 'dashed' | 'dotted';
  label: string;
  color: string;
}

const RtyLegendItem: React.FC<RtyLegend> = ({ type, label, color }) => {
  return (
    <div style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', fontSize: '10.5px', color: 'inherit', margin: '0 4px' }}>
      {type === 'bar' && (
        <span style={{ width: '10px', height: '10px', background: color, display: 'inline-block', borderRadius: '1.5px', flexShrink: 0 }} />
      )}
      {type === 'line' && (
        <span style={{ width: '15px', height: '2.5px', background: color, display: 'inline-block', flexShrink: 0 }} />
      )}
      {type === 'dashed' && (
        <span style={{ width: '15px', height: '0px', borderTop: `2.5px dashed ${color}`, display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }} />
      )}
      {type === 'dotted' && (
        <span style={{ width: '15px', height: '0px', borderTop: `2.5px dotted ${color}`, display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }} />
      )}
      <span style={{ whiteSpace: 'nowrap', opacity: 0.9 }}>{label}</span>
    </div>
  );
};

const IDB_KEY_RTY_SUMMARY_DATA = 'rty_summary_dynamic_v2';

function idbOpenSummary(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('imvina_dashboard_db', 2);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('cache')) {
        db.createObjectStore('cache');
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbGetCacheSummary(key: string): Promise<string | null> {
  try {
    const db = await idbOpenSummary();
    return new Promise((resolve) => {
      const tx = db.transaction('cache', 'readonly');
      const store = tx.objectStore('cache');
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror = () => resolve(null);
    });
  } catch {
    return null;
  }
}

async function idbSetCacheSummary(key: string, val: string): Promise<void> {
  try {
    const db = await idbOpenSummary();
    return new Promise((resolve) => {
      const tx = db.transaction('cache', 'readwrite');
      const store = tx.objectStore('cache');
      store.put(val, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => resolve();
    });
  } catch {
    // ignore
  }
}

interface RtySummaryDynamicData {
  modelSummary: ModelSummary[];
  modelSeries: Record<string, Record<string, { month: (number | null)[]; week: (number | null)[]; day: (number | null)[] }>>;
  dayLabels: string[];
  weekLabels: string[];
  monthLabels: string[];
  lastUpdateLabel: string;
  dataMaxDate: string;
  dataMinDate: string;
}

function parseWorkbookToRtySummaryData(wb: any): RtySummaryDynamicData | null {
  try {
    // EPCC (rty-upload-xlsx-global-missing) - FIX ROOT CAUSE "bấm Tải tệp lên nhưng
    // dữ liệu không cập nhật, không có thông báo lỗi nào cả":
    // Trước đây hàm này tìm thư viện XLSX qua `window.XLSX`/`globalThis.XLSX`, nhưng
    // dự án import XLSX theo kiểu ES module (`import * as XLSX from 'xlsx'`) nên
    // không bao giờ có biến global đó -> `XLSX_lib` luôn undefined -> hàm return null
    // ngay từ dòng đầu -> onChange coi như "đọc được file nhưng không có dữ liệu hợp lệ"
    // và bỏ qua trong im lặng. Fix: dùng thẳng module XLSX đã import tĩnh ở đầu file.
    const XLSX_lib = XLSX;
    const sheetNames = wb.SheetNames || [];
    const parsedRows: Array<{
      model: string;
      processKey: string;
      isActual: boolean;
      rty: number;
      dateInfo: NonNullable<ReturnType<typeof parseDateInfo>>;
    }> = [];

    const monthNamesMap: Record<string, number> = {
      jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
      jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
      'tháng 1': 1, 'tháng 2': 2, 'tháng 3': 3, 'tháng 4': 4, 'tháng 5': 5, 'tháng 6': 6,
      'tháng 7': 7, 'tháng 8': 8, 'tháng 9': 9, 'tháng 10': 10, 'tháng 11': 11, 'tháng 12': 12,
    };

    const parseDateInfo = (val: any, rawYear?: any) => {
      if (val == null) return null;
      let d: Date | null = null;
      let yearNum = typeof rawYear === 'number' ? rawYear : 2026;
      if (typeof rawYear === 'string' && !isNaN(Number(rawYear))) yearNum = Number(rawYear);
      if (yearNum < 2000 || yearNum > 2100) yearNum = 2026;

      if (val instanceof Date) {
        d = val;
      } else if (typeof val === 'number') {
        if (val > 1000 && val < 99999) {
          d = new Date(Math.round((val - 25569) * 86400 * 1000));
        } else if (val >= 1 && val <= 12) {
          d = new Date(yearNum, val - 1, 1);
        }
      } else if (typeof val === 'string') {
        const str = val.trim();
        const lowerStr = str.toLowerCase();
        if (monthNamesMap[lowerStr]) {
          d = new Date(yearNum, monthNamesMap[lowerStr] - 1, 1);
        } else if (str.includes('-')) {
          const parts = str.split('-');
          if (parts.length === 3) d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
          else if (parts.length === 2) {
            const p1 = Number(parts[0]);
            const p2 = Number(parts[1]);
            if (!isNaN(p1) && !isNaN(p2)) {
              if (p1 > 12) d = new Date(yearNum, p2 - 1, p1);
              else d = new Date(yearNum, p1 - 1, p2);
            }
          }
        } else if (str.includes('/')) {
          const parts = str.split('/');
          if (parts.length === 2) {
            const p1 = Number(parts[0]);
            const p2 = Number(parts[1]);
            if (!isNaN(p1) && !isNaN(p2)) {
              if (p1 > 12) d = new Date(yearNum, p2 - 1, p1);
              else d = new Date(yearNum, p1 - 1, p2);
            }
          } else if (parts.length === 3) d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        }
      }
      if (!d || isNaN(d.getTime())) return null;
      const yyyy = d.getFullYear();
      const mmNum = d.getMonth() + 1;
      const ddNum = d.getDate();
      const mm = String(mmNum).padStart(2, '0');
      const dd = String(ddNum).padStart(2, '0');
      const formattedDay = `${mm}/${dd}`;
      const isoDate = `${yyyy}-${mm}-${dd}`;
      const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
      const monthLabel = monthNames[d.getMonth()] || 'JAN';
      const startOfYear = new Date(yyyy, 0, 1);
      const weekNum = Math.ceil((((d.getTime() - startOfYear.getTime()) / 86400000) + startOfYear.getDay() + 1) / 7);
      const weekLabel = `W${String(weekNum).padStart(2, '0')}`;
      return { timestamp: d.getTime(), formattedDay, isoDate, monthLabel, weekLabel };
    };

    const parseRtyVal = (val: any): number | null => {
      if (val == null) return null;
      if (typeof val === 'number') {
        if (isNaN(val)) return null;
        return val > 1 ? val / 100 : val;
      }
      if (typeof val === 'string') {
        const num = parseFloat(val.replace('%', '').trim());
        if (isNaN(num)) return null;
        return num > 1 ? num / 100 : num;
      }
      return null;
    };

    const normalizeProcess = (proc: string): string => {
      const p = proc.trim().toUpperCase();
      if (p.includes('TTL') || p.includes('TOTAL') || p === 'RTY' || p === 'RTY %') return 'RTY_TTL';
      if (p.includes('MAIN FVI')) return 'MAIN_FVI';
      if (p.includes('MAIN ASSY') || p.includes('ASSY')) return 'MAIN_ASSY';
      if (p.includes('MAIN DRIVING') || p.includes('DRIVING')) return 'MAIN_DRIVING';
      if (p.includes('MAIN TILT') || p.includes('TILT')) return 'MAIN_TILT';
      if (p.includes('MAIN')) return 'RTY_MAIN';
      if (p.includes('SUB1 FPCB') || p.includes('FPCB')) return 'SUB1_FPCB';
      if (p.includes('SUB1 FVI')) return 'SUB1_FVI';
      if (p.includes('SUB1')) return 'RTY_SUB1';
      if (p.includes('SUB2 HOOK') || p.includes('HOOK')) return 'SUB2_HOOK';
      if (p.includes('SUB2 OVEN') || p.includes('OVEN')) return 'SUB2_OVEN';
      if (p.includes('SUB2 INDEX') || p.includes('INDEX')) return 'SUB2_INDEX';
      if (p.includes('SUB2')) return 'RTY_SUB2';
      return 'RTY_TTL';
    };

    const gv = (r: any, ...keys: string[]) => {
      for (const k of keys) {
        if (r[k] !== undefined && r[k] !== null) return r[k];
        const lowerK = k.toLowerCase();
        for (const rk of Object.keys(r)) {
          if (rk.trim().toLowerCase() === lowerK && r[rk] !== undefined && r[rk] !== null) {
            return r[rk];
          }
        }
      }
      return null;
    };

    sheetNames.forEach((sName: string) => {
      const sheet = wb.Sheets[sName];
      if (!sheet) return;
      const rows = XLSX_lib.utils.sheet_to_json(sheet, { defval: null }) as any[];
      if (!rows || rows.length === 0) return;

      rows.forEach(r => {
        const rawModel = gv(r, 'model', 'Model', 'MODEL', 'so', 'SO');
        const rawProc  = gv(r, 'process', 'Process', 'PROCESS', 'item', 'Item', 'ITEM');
        const rawType  = gv(r, 'type', 'Type', 'TYPE');
        const rawYear  = gv(r, 'year', 'Year', 'YEAR');
        const rawDate  = gv(r, 'date', 'Date', 'DATE', 'period', 'Period', 'PERIOD', 'month', 'Month', 'MONTH', 'day', 'Day', 'DAY', 'time', 'Time', 'TIME');
        const rawRty   = gv(r, 'rty', 'RTY', 'RTY %', 'RTY%', 'value', 'Value', 'VALUE', 'rate', 'Rate');

        if (rawModel && rawProc) {
          const modelStr = String(rawModel).trim();
          const processKey = normalizeProcess(String(rawProc));
          const typeStr = rawType ? String(rawType).trim().toLowerCase() : 'actual';
          // EPCC (rty-isActual-detection-miss): mở rộng nhận dạng nhãn Actual:
          // - tiếng Anh: 'actual', 'act', 'real', 'result'
          // - tiếng Việt: 'thực tế', 'thực hiện', 'thực'
          // - tiếng Hàn: '실적', '실제'
          // - đặc biệt: nếu rawType không tồn tại (null) → mặc định là actual
          //   (nhiều file RTY không có cột Type riêng, chỉ có 1 loại dữ liệu)
          const isActual = !rawType // không có cột Type → mặc định là actual
            || typeStr.includes('actual') || typeStr.includes('act')
            || typeStr.includes('real') || typeStr.includes('result')
            || typeStr.includes('thực tế') || typeStr.includes('thực hiện') || typeStr.includes('thực')
            || typeStr.includes('실적') || typeStr.includes('실제');

          if (rawRty != null && rawDate != null) {
            const rtyVal = parseRtyVal(rawRty);
            const dateInfo = parseDateInfo(rawDate, rawYear);
            if (rtyVal != null && dateInfo != null) {
              parsedRows.push({ model: modelStr, processKey, isActual, rty: rtyVal, dateInfo });
            }
          }

          Object.keys(r).forEach(colKey => {
            const keyTrim = colKey.trim();
            const dateInfo = parseDateInfo(keyTrim, rawYear);
            if (dateInfo != null) {
              const rtyVal = parseRtyVal(r[colKey]);
              if (rtyVal != null) {
                parsedRows.push({ model: modelStr, processKey, isActual, rty: rtyVal, dateInfo });
              }
            }
          });
        }
      });
    });

    if (parsedRows.length === 0) return null;

    const dayMap = new Map<string, number>();
    parsedRows.forEach(r => {
      if (!dayMap.has(r.dateInfo.formattedDay) || r.dateInfo.timestamp < dayMap.get(r.dateInfo.formattedDay)!) {
        dayMap.set(r.dateInfo.formattedDay, r.dateInfo.timestamp);
      }
    });
    const dayEntries = Array.from(dayMap.entries()).sort((a, b) => a[1] - b[1]);
    const dayLabels = dayEntries.map(e => e[0]);

    const weekMap = new Map<string, number>();
    const monthMap = new Map<string, number>();
    parsedRows.forEach(r => {
      if (!weekMap.has(r.dateInfo.weekLabel) || r.dateInfo.timestamp < weekMap.get(r.dateInfo.weekLabel)!) {
        weekMap.set(r.dateInfo.weekLabel, r.dateInfo.timestamp);
      }
      if (!monthMap.has(r.dateInfo.monthLabel) || r.dateInfo.timestamp < monthMap.get(r.dateInfo.monthLabel)!) {
        monthMap.set(r.dateInfo.monthLabel, r.dateInfo.timestamp);
      }
    });
    // EPCC (rty-weeklabels-not-sorted) - FIX ROOT CAUSE "radar/spider trống
    // khi Xem theo Tuần/Tháng dù bar chart vẫn có số liệu": bản cũ đẩy
    // weekLabels/monthLabels vào mảng theo THỨ TỰ GẶP TRONG parsedRows (thứ
    // tự dòng thô trong file Excel — có thể KHÔNG tăng dần theo thời gian
    // nếu file nhóm theo Model/Process trước khi mới tới Ngày), khác hẳn
    // dayLabels (đã sort đúng theo timestamp). filterAndTakeLast8() lấy "8
    // mốc gần nhất" bằng cách .slice(-8) trên vị trí MẢNG, nên nếu thứ tự bị
    // xáo trộn, "8 mốc cuối mảng" không còn là "8 tuần/tháng gần nhất theo
    // thời gian" nữa -> vênh giữa các model khi radar duyệt qua toàn bộ
    // model, dù bar chart (chỉ đọc đúng 1 model theo cùng idxs) có thể vẫn
    // trông hợp lý một cách tình cờ. Sửa: sort weekLabels/monthLabels theo
    // timestamp giống hệt cách dayLabels đã làm.
    const weekLabels = Array.from(weekMap.entries()).sort((a, b) => a[1] - b[1]).map(e => e[0]);
    const monthLabels = Array.from(monthMap.entries()).sort((a, b) => a[1] - b[1]).map(e => e[0]);

    const lastUpdateLabel = dayLabels[dayLabels.length - 1] || '07/08';
    let dataMaxDate = '2026-07-09';
    let dataMinDate = '2026-01-02';
    parsedRows.forEach(r => {
      if (r.dateInfo.isoDate > dataMaxDate) dataMaxDate = r.dateInfo.isoDate;
      if (r.dateInfo.isoDate < dataMinDate) dataMinDate = r.dateInfo.isoDate;
    });

    // EPCC (rty-isActual-detection-miss) - FALLBACK: nếu không có dòng actual nào
    // (ví dụ: file chỉ có nhãn 'Target'/'Kế hoạch'/tiếng Hàn không nhận dạng được),
    // dùng TOÀN BỘ parsedRows như thể chúng đều là actual. Điều này đúng với phần lớn
    // file RTY thực tế chỉ chứa một loại dữ liệu (actual) dù Type có thể được gán
    // nhãn khác nhau trong hệ thống của nhà máy.
    const actualRows = parsedRows.filter(r => r.isActual);
    const rowsForModelSeries = actualRows.length > 0 ? actualRows : parsedRows;
    if (actualRows.length === 0 && parsedRows.length > 0) {
      console.warn(`[RTY Parse] Không tìm thấy dòng nào có type='actual' trong ${parsedRows.length} dòng. Áp dụng fallback: côi toàn bộ dòng là Actual.`);
    }

    const modelSeries: Record<string, Record<string, { month: (number | null)[]; week: (number | null)[]; day: (number | null)[] }>> = {};

    rowsForModelSeries.forEach(row => {
      if (!modelSeries[row.model]) modelSeries[row.model] = {};
      if (!modelSeries[row.model][row.processKey]) {
        modelSeries[row.model][row.processKey] = {
          day: new Array(dayLabels.length).fill(null),
          week: new Array(weekLabels.length).fill(null),
          month: new Array(monthLabels.length).fill(null),
        };
      }
      const dayIdx = dayLabels.indexOf(row.dateInfo.formattedDay);
      if (dayIdx >= 0) modelSeries[row.model][row.processKey].day[dayIdx] = row.rty;

      const weekIdx = weekLabels.indexOf(row.dateInfo.weekLabel);
      if (weekIdx >= 0) modelSeries[row.model][row.processKey].week[weekIdx] = row.rty;

      const monthIdx = monthLabels.indexOf(row.dateInfo.monthLabel);
      if (monthIdx >= 0) modelSeries[row.model][row.processKey].month[monthIdx] = row.rty;
    });

    // EPCC (rty-ttl-missing-compute): nếu file không có cột Process 'TTL' riêng,
    // RTY_TTL sẽ bị rỗng → line TTL trên biểu đồ không hiện.
    // Tính tự động theo công thức RTY chuẩn trong sản xuất:
    //   RTY_TTL = RTY_MAIN × RTY_SUB1 × RTY_SUB2
    // Áp dụng cho từng điểm thời gian (day/week/month) riêng biệt.
    // Nếu model đã có RTY_TTL (từ file), bỏ qua — không ghi đè.
    const computeRtyTtl = (
      m: string,
      mode: 'day' | 'week' | 'month'
    ): (number | null)[] => {
      const main  = modelSeries[m]?.['RTY_MAIN']?.[mode]  ?? [];
      const sub1  = modelSeries[m]?.['RTY_SUB1']?.[mode]  ?? [];
      const sub2  = modelSeries[m]?.['RTY_SUB2']?.[mode]  ?? [];
      const len = Math.max(main.length, sub1.length, sub2.length);
      return Array.from({ length: len }, (_, i) => {
        const values = [main[i], sub1[i], sub2[i]].filter((v): v is number => v != null && v > 0);
        if (values.length === 0) return null;
        return values.reduce((a, b) => a * b, 1);
      });
    };
    for (const m of Object.keys(modelSeries)) {
      const hasExplicitTTL = (modelSeries[m]?.['RTY_TTL']?.day ?? []).some(v => v != null);
      if (!hasExplicitTTL) {
        const ttlDay   = computeRtyTtl(m, 'day');
        const ttlWeek  = computeRtyTtl(m, 'week');
        const ttlMonth = computeRtyTtl(m, 'month');
        const hasAny = ttlDay.some(v => v != null) || ttlWeek.some(v => v != null) || ttlMonth.some(v => v != null);
        if (hasAny) {
          modelSeries[m]['RTY_TTL'] = { day: ttlDay, week: ttlWeek, month: ttlMonth };
        }
      }
    }

    const modelsAll = Object.keys(modelSeries);
    const modelSummary: ModelSummary[] = modelsAll.map(m => {
      const getAvg = (procKey: string) => {
        const arr = (modelSeries[m]?.[procKey]?.day || []).filter((v): v is number => v != null);
        if (arr.length === 0) return null;
        return arr.reduce((a, b) => a + b, 0) / arr.length;
      };
      const sub1Act = getAvg('RTY_SUB1');
      const sub2Act = getAvg('RTY_SUB2');
      const mainAct = getAvg('RTY_MAIN');
      const ttlAct  = getAvg('RTY_TTL') ?? mainAct ?? 0.95;

      return {
        model: m,
        ...(sub1Act != null ? { sub1: { target: 0.992, actual: sub1Act } } : {}),
        ...(sub2Act != null ? { sub2: { target: 0.988, actual: sub2Act } } : {}),
        ...(mainAct != null ? { main: { target: 0.983, actual: mainAct } } : {}),
        ttl: { target: 0.964, actual: ttlAct },
      };
    });

    return {
      modelSummary,
      modelSeries,
      dayLabels,
      weekLabels,
      monthLabels,
      lastUpdateLabel,
      dataMaxDate,
      dataMinDate,
    };
  } catch (err) {
    console.error('Error parsing RTY summary workbook:', err);
    return null;
  }
}

export const RtyDashboard: React.FC<RtyDashboardProps> = ({
  theme, onToggleTheme: _onToggleTheme, lang, setLang: _setLang, onFileSelected, onSyncProgress,
}) => {
  const t = TXT[lang];
  const isLightMode = theme === 'light';
  const tealAccent  = isLightMode ? '#0f766e' : '#14b8a6';

  const [activeTab, setActiveTab] = useState<'summary' | 'rtyTotal' | 'merged'>('summary');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [dynamicRtyData, setDynamicRtyData] = useState<RtySummaryDynamicData | null>(null);

  useEffect(() => {
    (async () => {
      const cached = await idbGetCacheSummary(IDB_KEY_RTY_SUMMARY_DATA);
      if (cached) {
        try {
          const parsed = JSON.parse(cached);
          // EPCC (rty-empty-array-falsy-bug): chỉ dùng cache khi modelSummary CÓ models
          // thực sự. Cache cũ (từ lần upload lỗi trước đây) có thể có modelSummary=[]
          // → gây 0/0 KPI. Bỏ qua cache rỗng, để activeModelSummary fallback về static.
          if (parsed && parsed.modelSummary && Array.isArray(parsed.modelSummary)
              && parsed.modelSummary.length > 0 && parsed.dayLabels) {
            setDynamicRtyData(parsed);
          } else if (parsed && parsed.dayLabels && (!parsed.modelSummary || parsed.modelSummary.length === 0)) {
            console.warn('[RTY IDB] Cache có dayLabels nhưng modelSummary rỗng — bỏ qua cache hỏng, dùng dữ liệu tĩnh.');
            // Xóa luôn cache hỏng để lần sau không load lại nữa.
            await idbSetCacheSummary(IDB_KEY_RTY_SUMMARY_DATA, '');
          }
        } catch { /* ignore */ }
      }
    })();
  }, []);

  // EPCC (rty-empty-array-falsy-bug): operator `||` không fallback khi giá trị
  // là mảng rỗng `[]` ([] là TRUTHY trong JS). Phải kiểm tra length ạnh hưởng đến
  // cả modelSummary và modelSeries: nếu mảng rỗng/object rỗng, fallback về data tĩnh.
  const activeModelSummary = (dynamicRtyData?.modelSummary && dynamicRtyData.modelSummary.length > 0)
    ? dynamicRtyData.modelSummary
    : MODEL_SUMMARY;
  const activeModelSeries  = (dynamicRtyData?.modelSeries && Object.keys(dynamicRtyData.modelSeries).length > 0)
    ? dynamicRtyData.modelSeries
    : MODEL_SERIES;
  const activeDayLabels    = dynamicRtyData?.dayLabels    || DAY_LABELS;
  const activeWeekLabels   = dynamicRtyData?.weekLabels   || WEEK_LABELS;
  const activeMonthLabels  = dynamicRtyData?.monthLabels  || MONTH_LABELS;
  const activeLastUpdate   = dynamicRtyData?.lastUpdateLabel || LAST_DATA_UPDATE_LABEL;
  const activeDataMaxDate  = dynamicRtyData?.dataMaxDate  || '2026-07-09';
  const activeDataMinDate  = dynamicRtyData?.dataMinDate  || '2026-01-02';

  const activeModelSummarySorted = useMemo(
    () => [...activeModelSummary].sort((a, b) => (b.ttl.actual - b.ttl.target) - (a.ttl.actual - a.ttl.target)),
    [activeModelSummary]
  );
  const activeBestModel = activeModelSummarySorted[0]?.model || BEST_MODEL;

  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  const formattedTime = useMemo(() => {
    const p = (n: number) => String(n).padStart(2, '0');
    return `${p(now.getHours())}:${p(now.getMinutes())}:${p(now.getSeconds())} ${p(now.getDate())}/${p(now.getMonth() + 1)}/${now.getFullYear()}`;
  }, [now]);

  // ── Plotly readiness (poll cho tới khi script CDN load xong) ──────────
  const [plotlyReady, setPlotlyReady] = useState<boolean>(
    typeof window !== 'undefined' && !!(window as any).Plotly
  );
  useEffect(() => {
    if (plotlyReady) return;
    const id = setInterval(() => {
      if (typeof window !== 'undefined' && (window as any).Plotly) {
        setPlotlyReady(true);
        clearInterval(id);
      }
    }, 200);
    return () => clearInterval(id);
  }, [plotlyReady]);

  const getDefaultStartDate = (): string => activeDataMinDate;
  const getDefaultEndDate = (): string => activeDataMaxDate;

  const [spiderLegendTTL, setSpiderLegendTTL] = useState<{ label: string; color: string }[]>([]);
  const [spiderLegendMAIN, setSpiderLegendMAIN] = useState<{ label: string; color: string }[]>([]);

  const [selectedModel, setSelectedModel] = useState<string>(activeBestModel);
  const [viewMode, setViewMode] = useState<'day' | 'week' | 'month'>('day');
  const [startDate, setStartDate] = useState<string>(getDefaultStartDate);
  const [endDate,   setEndDate]   = useState<string>(getDefaultEndDate);
  const filterLabelColor = '#C0EF6A';

  // EPCC (rty-date-range-stale-on-idb-load): khi dynamicRtyData load từ IDB cache,
  // startDate/endDate vẫn giữ giá trị tĩnh mặc định (e.g. '2026-07-09') vì useState
  // chỉ dùng initializer 1 lần duy nhất lúc mount. Chart bị giới hạn đến ngày tĩnh
  // dù dữ liệu đã có đến ngày mới nhất (e.g. '2026-07-21').
  // Sửa: useEffect theo dõi thay đổi của dynamicRtyData và tự cập nhật ngày
  // để chart tự mở rộng đến ngày thực tế trong file upload.
  useEffect(() => {
    if (!dynamicRtyData) return;
    setStartDate(dynamicRtyData.dataMinDate);
    setEndDate(dynamicRtyData.dataMaxDate);
  }, [dynamicRtyData?.dataMinDate, dynamicRtyData?.dataMaxDate]);

  const _resetFilters = () => {
    setSelectedModel(activeBestModel); setViewMode('day');
    setStartDate(getDefaultStartDate()); setEndDate(getDefaultEndDate());
  }; void _resetFilters;

  // ── Quy đổi nhãn Tháng/Tuần/Ngày sang mốc ngày thật (năm 2026) để lọc
  //    theo Ngày bắt đầu/kết thúc. ─────────────────────────────────────
  const MONTH_TO_NUM: Record<string, number> = { JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUL: 7 };
  const labelToDate = (label: string, mode: 'day' | 'week' | 'month'): Date => {
    if (mode === 'day') {
      const [mm, dd] = label.split('/').map(Number);
      return new Date(2026, mm - 1, dd);
    }
    if (mode === 'week') {
      const wn = parseInt(label.replace('W', ''), 10);
      const base = new Date(2026, 0, 1);
      base.setDate(base.getDate() + (wn - 1) * 7);
      return base;
    }
    const mm = MONTH_TO_NUM[label] ?? 1;
    return new Date(2026, mm - 1, 1);
  };

  // FIX (8-periods-all-modes, EPCC): đổi từ "lấy tối đa 10" sang ĐÚNG "8 giai
  // đoạn gần nhất" theo yêu cầu, áp dụng như nhau cho cả Ngày/Tuần/Tháng —
  // kết hợp với default date range đã mở rộng ở trên, giờ Tuần/Tháng cũng sẽ
  // luôn có đủ (tối đa) 8 giai đoạn gần nhất thay vì bị kẹt trong 1 tháng.
  const filterAndTakeLast8 = (labels: string[], mode: 'day' | 'week' | 'month'): number[] => {
    const start = startDate ? new Date(startDate + 'T00:00:00') : null;
    const end = endDate ? new Date(endDate + 'T23:59:59') : null;
    const idxs: number[] = [];
    labels.forEach((lb, i) => {
      const d = labelToDate(lb, mode);
      if ((!start || d >= start) && (!end || d <= end)) idxs.push(i);
    });
    return idxs.slice(-8);
  };

  // EPCC (rty-chart-static-data-bug) - FIX ROOT CAUSE "biểu đồ không cập nhật
  // sau khi tải file lên": trước đây filteredModels dùng hằng số MODEL_SUMMARY
  // tĩnh thay vì activeModelSummary (dữ liệu động từ file upload/cache IDB).
  // Khi upload file mới → setDynamicRtyData cập nhật activeModelSummary đúng,
  // nhưng filteredModels vẫn đọc từ MODEL_SUMMARY cũ → 4 thẻ KPI không đổi.
  const filteredModels = useMemo(
    () => selectedModel ? activeModelSummary.filter(m => m.model === selectedModel) : activeModelSummary,
    [selectedModel, activeModelSummary]
  );

  const kpi = useMemo(() => {
    const list = filteredModels;
    const avgActual = list.reduce((a, m) => a + m.ttl.actual, 0) / (list.length || 1);
    const target = list[0]?.ttl.target ?? 0.964;
    const meeting = list.filter(m => m.ttl.actual >= m.ttl.target).length;
    const achieveRate = (meeting / (list.length || 1)) * 100;
    const avgGapPct = (avgActual - target) * 100;
    return {
      actual: avgActual, target, achieveRate,
      meetingCount: meeting, totalCount: list.length,
      avgGapPct,
    };
  }, [filteredModels]);

  // EPCC (rty-chart-static-data-bug) - FIX ROOT CAUSE "bảng RTY & Chi tiết
  // không cập nhật sau khi tải file lên": trước đây filteredRows dùng
  // TABLE_ROWS_BY_MODE (tính từ hằng số MODEL_SERIES tĩnh). Khi upload file
  // mới → activeModelSeries thay đổi nhưng bảng vẫn dùng data cũ.
  // Sửa: xây dựng dynamic rows từ activeModelSeries + labels động.
  const filteredRows = useMemo(() => {
    // Nếu có dữ liệu động (sau upload), xây bảng từ activeModelSeries
    if (dynamicRtyData) {
      const labels = viewMode === 'day' ? activeDayLabels : viewMode === 'week' ? activeWeekLabels : activeMonthLabels;
      const rows: RtyTableRow[] = [];
      const processKeyDefs = [
        { processKey: 'RTY_TTL',    process: 'RTY Total',     type: 'Actual', target: TARGET_TTL },
        { processKey: 'RTY_MAIN',   process: 'MAIN',          type: 'Actual', target: TARGET_MAIN },
        { processKey: 'RTY_SUB1',   process: 'SUB1',          type: 'Actual', target: TARGET_SUB1 },
        { processKey: 'RTY_SUB2',   process: 'SUB2',          type: 'Actual', target: TARGET_SUB2 },
        { processKey: 'MAIN_FVI',   process: 'MAIN FVI',      type: 'Actual', target: TARGET_MAIN },
        { processKey: 'MAIN_ASSY',  process: 'MAIN ASSY',     type: 'Actual', target: TARGET_MAIN },
        { processKey: 'MAIN_DRIVING', process: 'MAIN DRIVING', type: 'Actual', target: TARGET_MAIN },
        { processKey: 'MAIN_TILT',  process: 'MAIN TILT',     type: 'Actual', target: TARGET_MAIN },
        { processKey: 'SUB1_FPCB',  process: 'SUB1 FPCB',     type: 'Actual', target: TARGET_SUB1 },
        { processKey: 'SUB1_FVI',   process: 'SUB1 FVI',      type: 'Actual', target: TARGET_SUB1 },
        { processKey: 'SUB2_HOOK',  process: 'SUB2 HOOK',     type: 'Actual', target: TARGET_SUB2 },
        { processKey: 'SUB2_OVEN',  process: 'SUB2 OVEN',     type: 'Actual', target: TARGET_SUB2 },
        { processKey: 'SUB2_INDEX', process: 'SUB2 INDEX',    type: 'Actual', target: TARGET_SUB2 },
      ];
      const modelsToShow = selectedModel
        ? activeModelSummary.filter(m => m.model === selectedModel).map(m => m.model)
        : activeModelSummary.map(m => m.model);
      modelsToShow.forEach(modelName => {
        const seriesForModel = activeModelSeries[modelName];
        if (!seriesForModel) return;
        processKeyDefs.forEach(def => {
          const arr = (seriesForModel[def.processKey] as any)?.[viewMode] as (number | null)[] | undefined;
          if (!arr) return;
          labels.forEach((label, i) => {
            const v = arr[i];
            if (v == null) return;
            rows.push({ model: modelName, process: def.process, type: def.type, period: label, target: def.target, actual: v });
          });
        });
      });
      return rows;
    }
    // Fallback: dùng bảng tĩnh đã sinh sẵn (khi chưa có data động)
    const rows = TABLE_ROWS_BY_MODE[viewMode];
    return selectedModel ? rows.filter(r => r.model === selectedModel) : rows;
  }, [viewMode, selectedModel, dynamicRtyData, activeModelSeries, activeModelSummary, activeDayLabels, activeWeekLabels, activeMonthLabels]);
  const [page, setPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(25);
  useEffect(() => { setPage(1); }, [selectedModel, viewMode, rowsPerPage]);
  const totalPages = Math.max(1, Math.ceil(filteredRows.length / rowsPerPage));
  const pageRows = useMemo(() => {
    const start = (page - 1) * rowsPerPage;
    return filteredRows.slice(start, start + rowsPerPage);
  }, [filteredRows, page, rowsPerPage]);

  const RADAR_COLORS = useMemo(() => [
    '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
    '#06b6d4', '#ec4899', '#14b8a6', '#f43f5e', '#6366f1'
  ], []);

  const { xs, idxs } = useMemo(() => {
    const labelsForMode = viewMode === 'day' ? activeDayLabels : viewMode === 'week' ? activeWeekLabels : activeMonthLabels;
    const idxsVal = filterAndTakeLast8(labelsForMode, viewMode);
    const xsVal   = idxsVal.map(i => labelsForMode[i]);
    return { xs: xsVal, idxs: idxsVal };
  }, [viewMode, startDate, endDate, activeDayLabels, activeWeekLabels, activeMonthLabels]);

  const tableScrollRef = useRef<HTMLDivElement | null>(null);
  const paginationBarRef = useRef<HTMLDivElement | null>(null);
  const [tableScrollMaxHeight, setTableScrollMaxHeight] = useState<number>(
    typeof window !== 'undefined' ? Math.max(240, window.innerHeight * 0.62) : 480
  );

  useEffect(() => {
    const recalcTableScrollHeight = () => {
      const el = tableScrollRef.current;
      if (!el) return;
      const top = el.getBoundingClientRect().top;
      const paginationHeight = paginationBarRef.current?.getBoundingClientRect().height ?? 44;
      const BOTTOM_GAP = 12; // đệm nhỏ để không dính sát mép dưới cửa sổ
      const available = window.innerHeight - top - paginationHeight - BOTTOM_GAP;
      setTableScrollMaxHeight(Math.max(200, available));
    };
    recalcTableScrollHeight();
    window.addEventListener('resize', recalcTableScrollHeight);
    const raf = requestAnimationFrame(recalcTableScrollHeight);
    return () => {
      window.removeEventListener('resize', recalcTableScrollHeight);
      cancelAnimationFrame(raf);
    };
  }, [activeTab, pageRows.length]);

  useEffect(() => {
    if (!plotlyReady || activeTab !== 'summary' || typeof window === 'undefined' || !window.Plotly) return;

    const fontColor = isLightMode ? '#334155' : '#e2e8f0';
    const gridColor = isLightMode ? 'rgba(0,0,0,0.08)' : 'rgba(255,255,255,0.08)';

    const plotMixedPanel = (
      elId: string,
      barSeries:  { key: string; name: string; color: string }[],
      lineSeries: { key: string; name: string; color: string; width?: number; dash?: string }[],
      targetFrac: number, targetLabel: string
    ) => {
      const el = document.getElementById(elId);
      if (!el) return;

      const getYs = (key: string): (number | null)[] =>
        idxs.map(i => {
          const v = getSeriesValueForModel(selectedModel, key, viewMode, i, activeModelSeries);
          return v == null ? null : parseFloat((v * 100).toFixed(2));
        });

      const barTraces = barSeries.map(s => {
        const ys = getYs(s.key);
        return {
          x: xs, y: ys, name: s.name,
          type: 'bar' as const,
          marker: { color: s.color },
          text: ys.map(v => v == null ? '' : `${v.toFixed(1)}%`),
          textposition: 'inside' as const,
          textfont: { size: 10, color: '#ffffff', family: 'Arial Black, Arial, sans-serif' },
          insidetextanchor: 'middle' as const,
          hovertemplate: `%{x}<br><b>${s.name}: %{y:.2f}%</b><extra></extra>`,
        };
      });

      const lineTraces = lineSeries.map((s, idx) => {
        const ys = getYs(s.key);
        const pos = idx === 0 ? 'top center' : 'bottom center';
        return {
          x: xs, y: ys, name: s.name,
          type: 'scatter' as const, mode: 'lines+markers+text' as const,
          yaxis: 'y2' as const,
          line: { color: s.color, width: s.width ?? 1.6, shape: 'spline' as const, smoothing: 1, dash: s.dash as any },
          marker: { color: s.color, size: 6 },
          text: ys.map(v => v == null ? '' : `${v.toFixed(1)}%`),
          textposition: pos as 'top center' | 'bottom center',  // cast to valid Plotly literal
          textfont: { size: 10, color: s.color, family: 'Arial Black, Arial, sans-serif' },
          cliponaxis: false,
          connectgaps: true,
          hovertemplate: `%{x}<br><b>${s.name}: %{y:.2f}%</b><extra></extra>`,
        };
      });

      const barYs   = barTraces.flatMap(tr => (tr.y as (number|null)[]).filter((v): v is number => v != null));
      const lineYs  = lineTraces.flatMap(tr => (tr.y as (number|null)[]).filter((v): v is number => v != null));
      const targetY = targetFrac * 100;
      lineYs.push(targetY);

      // ── Tách biệt 2 trục Y để đường bay hẳn lên trên cột ──
      const BAR_TOP_FRAC     = 0.55;
      const LINE_BOTTOM_FRAC = 0.63;
      const LINE_TOP_FRAC    = 0.95;

      const maxBarVal = Math.max(...barYs, 0);
      const yRange   = [0, maxBarVal > 0 ? maxBarVal / BAR_TOP_FRAC : 100 / BAR_TOP_FRAC];

      const lineMin  = Math.min(...lineYs);
      const lineMax  = Math.max(...lineYs);
      const spanFrac = LINE_TOP_FRAC - LINE_BOTTOM_FRAC;
      const rawSpan  = lineMax - lineMin;
      const D        = rawSpan > 0.01 ? rawSpan / spanFrac : Math.max(lineMax * 0.05, 1);
      const y2Min    = lineMin - LINE_BOTTOM_FRAC * D;
      const y2Max    = y2Min + D;

      window.Plotly.newPlot(elId, [
        ...barTraces,
        ...lineTraces,
        {
          x: xs, y: xs.map(() => targetY),
          name: `${targetLabel} ${targetY.toFixed(1)}%`,
          type: 'scatter', mode: 'lines+markers+text',
          yaxis: 'y2',
          // FIX (target-line-red-above, EPCC): đường Target trước đây màu
          // amber (#f59e0b), nhãn số nằm DƯỚI đường ('bottom center'). Đổi
          // sang màu ĐỎ (#ef4444) theo yêu cầu, nhãn số chuyển lên TRÊN
          // đường ('top center' = "Above"). Vì màu đỏ này trước đó đã được
          // dùng cho 1 cột trong mỗi biểu đồ (RTY Sub2 / Main FVI Final /
          // S1 FVI / S2 Oven Cure), 4 cột đó đã được đổi sang màu khác ở
          // lời gọi plotMixedPanel() bên dưới để không trùng màu với Target.
          line: { color: '#ef4444', width: 1.6, dash: 'dot', shape: 'spline', smoothing: 1 },
          marker: { color: '#ef4444', size: 5 },
          text: xs.map(() => `${targetY.toFixed(1)}%`),
          textposition: 'top center',
          textfont: { size: 9, color: '#ef4444', family: 'Arial Black, Arial, sans-serif' },
          cliponaxis: false,
          hovertemplate: `Target: ${targetY.toFixed(1)}%<extra></extra>`,
        },
      ], {
        barmode: 'group',
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: { color: fontColor, size: 11 },
        margin: { t: 15, r: 40, b: 28, l: 48 }, // Thu nhỏ top margin xuống 15px để nới rộng biểu đồ lên trên
        showlegend: false, // Ẩn legend của Plotly để nhúng lên thanh tiêu đề header
        xaxis: { tickfont: { size: 11, color: fontColor }, gridcolor: gridColor, tickangle: 0 },
        yaxis:  { gridcolor: gridColor, tickfont: { size: 10 }, range: yRange, ticksuffix: '%' },
        yaxis2: {
          overlaying: 'y', side: 'right', showgrid: false,
          tickfont: { size: 10, color: fontColor }, range: [y2Min, y2Max], ticksuffix: '%',
        },
        hoverlabel: { font: { size: 10 } },
      }, { displayModeBar: false, responsive: true });
    };

    // Target: đỏ đứt nét (#ef4444) — nhãn số hiển thị Above (top center).
    // FIX (revert-bar-recolor-line, EPCC): bản trước đổi màu CỘT (bar) để
    // tránh trùng đỏ với Target — theo yêu cầu mới, cột trả về ĐÚNG màu đỏ
    // gốc như ảnh tham chiếu. Thay vào đó, đổi màu ĐƯỜNG (line) nào đang có
    // tông đỏ/hồng gần giống Target: "RTY TTL" (TTL RTY chart) và "RTY Main
    // (line)" (MAIN chart) — trước đây cùng dùng rose #f43f5e, dễ lẫn với
    // đỏ Target — nay đổi sang tím/hồng magenta để phân biệt rõ.

    // Panel 1 — TTL RTY
    //   Cột: RTY Sub1 (xanh lá), RTY Sub2 (xanh lam)
    //   Đường: RTY Main (cam), RTY TTL (vàng đậm — nổi bật tổng thể)
    plotMixedPanel('rtyChartTTL',
      [
        { key: 'RTY_SUB1', name: 'RTY Sub1', color: '#1565C0' },   // xanh đậm (PerCapita DAY)
        { key: 'RTY_SUB2', name: 'RTY Sub2', color: '#ef4444' },   // đỏ (trả về màu gốc theo ảnh tham chiếu)
      ],
      [
        { key: 'RTY_MAIN', name: 'RTY Main',  color: tealAccent,  width: 1.6 },         // teal (PerCapita TTL)
        { key: 'RTY_TTL',  name: 'RTY TTL',   color: '#8b5cf6',   width: 1.6, dash: 'dash' }, // tím (đổi từ rose để tránh trùng tông đỏ với Target)
      ],
      TARGET_TTL, 'Target RTY TTL');

    // Panel 2 — MAIN
    //   Cột: RTY Main (xanh lá), Main FVI Final (xanh lam)
    //   Đường: Main Assy (cam), Main Driving (tím), Main Tilt (xám), RTY Main line (vàng)
    plotMixedPanel('rtyChartMAIN',
      [
        { key: 'RTY_MAIN', name: 'RTY Main',       color: '#1565C0' },
        { key: 'MAIN_FVI', name: 'Main FVI Final',  color: '#ef4444' }, // đỏ (trả về màu gốc theo ảnh tham chiếu)
      ],
      [
        { key: 'MAIN_ASSY',    name: 'Main Assy',         color: tealAccent, width: 1.6 },
        { key: 'MAIN_DRIVING', name: 'Main Driving test',  color: '#f59e0b', width: 1.6 },
        { key: 'MAIN_TILT',   name: 'Main Tilt test',     color: '#a78bfa', width: 1.6 },
        { key: 'RTY_MAIN',    name: 'RTY Main (line)',    color: '#ec4899', width: 1.6, dash: 'dash' }, // hồng magenta (đổi từ rose để tránh trùng tông đỏ với Target)
      ],
      TARGET_MAIN, 'Target Main');

    // Panel 3 — SUB1
    //   Cột: S1 FPCB VI (xanh lá), S1 FVI (xanh lam)
    //   Đường: RTY Sub1 (vàng đậm)
    plotMixedPanel('rtyChartSUB1',
      [
        { key: 'SUB1_FPCB', name: 'S1 FPCB VI', color: '#1565C0' },
        { key: 'SUB1_FVI',  name: 'S1 FVI',     color: '#ef4444' }, // đỏ (trả về màu gốc theo ảnh tham chiếu)
      ],
      [
        { key: 'RTY_SUB1', name: 'RTY Sub1', color: tealAccent, width: 1.6 },
      ],
      TARGET_SUB1, 'Target Sub1');

    // Panel 4 — SUB2
    //   Cột: S2 Hook Bonding (xanh lá), S2 Oven Cure (xanh lam)
    //   Đường: RTY Sub2 (cam), S2 INDEX (vàng)
    plotMixedPanel('rtyChartSUB2',
      [
        { key: 'SUB2_HOOK', name: 'S2 Hook Bonding', color: '#1565C0' },
        { key: 'SUB2_OVEN', name: 'S2 Oven Cure',    color: '#ef4444' }, // đỏ (trả về màu gốc theo ảnh tham chiếu)
      ],
      [
        { key: 'RTY_SUB2',   name: 'RTY Sub2',  color: tealAccent, width: 1.6 },
        { key: 'SUB2_INDEX', name: 'S2 INDEX',  color: '#f59e0b',  width: 1.6, dash: 'dot' },
      ],
      TARGET_SUB2, 'Target Sub2');

    // ── Hàm vẽ biểu đồ mạng nhện (radar) cho các Model ──
    // FIX (spider-link-by-period, EPCC): trước đây dùng thẳng `xs`/`idxs` —
    // vốn CHỈ lọc theo Ngày bắt đầu/kết thúc, không quan tâm mốc đó có số
    // liệu THẬT hay không — nên khi đổi Ngày/Tuần/Tháng, chart vẫn vẽ đủ mọi
    // mốc trong khoảng lọc dù toàn bộ model ở mốc đó đều phải rơi mức
    // (fallback), trông như dữ liệu "không khớp" giữa các mốc/model.
    // Sửa theo đúng cách Chart 5 (Spider) của PerCapitaTab.tsx:
    //   labelsWithData = lọc mốc có ít nhất 1 model có số liệu THẬT đúng mốc
    //   → rồi mới lấy 8 mốc GẦN NHẤT (labelsWithData.slice(-8)).
    // Nhờ vậy trục Ngày/Tuần/Tháng luôn "link" đúng với dữ liệu thật đang có,
    // giống hệt nguyên tắc của biểu đồ Spider bên PerCapita.
    //
    // FIX #2 (spider-model-no-fabricated-point, đối chiếu Test4.xlsx, EPCC):
    // Trước đây MỌI model trong MODEL_SUMMARY luôn được vẽ, kể cả khi model
    // đó KHÔNG có bất kỳ số liệu Ngày/Tháng thật nào trong toàn bộ các mốc
    // đang xem — ví dụ SO1C30 S25 kiểm trong Test4.xlsx hoàn toàn KHÔNG có
    // dòng dữ liệu nào rơi vào tháng 7 (RTY_TTL.month[JUL] = null VÀ mọi
    // RTY_TTL.day của tháng 7 đều null). Khi đó code cũ rơi thẳng xuống mức
    // "Tổng hợp" (trung bình cả giai đoạn Jan→Jul) rồi vẽ như thể đó là số
    // của tháng 7 — gây hiểu lầm "có dữ liệu" trong khi thực chất không có.
    // Sửa: loại HẲN model đó khỏi trục theta nếu không có số liệu Ngày/Tháng
    // thật (mức 1 hoặc 2) ở BẤT KỲ mốc nào trong các mốc đang chọn — chỉ
    // model có ít nhất 1 mốc dữ liệu thật mới được vẽ (giống modelsWithData
    // của PerCapitaTab). Mức "Tổng hợp" (mức 3) chỉ còn dùng để bù cho ĐÚNG
    // 1-2 mốc lẻ tẻ thiếu của 1 model đã đủ điều kiện, không dùng để "tạo"
    // hẳn 1 model không hề có dữ liệu.
    const plotSpiderPanel = (
      elId: string,
      processType: 'TTL' | 'MAIN'
    ) => {
      const el = document.getElementById(elId);
      if (!el) return;

      const seriesKey = processType === 'TTL' ? 'RTY_TTL' : 'RTY_MAIN';
      const modelsAll = activeModelSummary.map(m => m.model);

      const candidates = xs.map((label, i) => ({ label, rawIndex: idxs[i] }));
      // EPCC (rty-ttl-spider-blank): nhiều file RTY không có cột TTL riêng (chỉ có
      // MAIN/SUB1/SUB2) → labelsWithData rỗng → purge → spider TTL trắng.
      // Fix: với TTL, nếu không có RTY_TTL thì fallback kiểm tra RTY_MAIN
      // (vì modelSummary.ttl.actual đã tính từ mainAct khi RTY_TTL null,
      // getSpiderValueWithFallback → mức 'summary' vẫn trả được giá trị).
      const fallbackSeriesKey = processType === 'TTL' ? 'RTY_MAIN' : null;
      const labelsWithData = candidates.filter(({ rawIndex }) =>
        modelsAll.some(m =>
          getSeriesValueForModel(m, seriesKey, viewMode, rawIndex, activeModelSeries) != null
          || (fallbackSeriesKey && getSeriesValueForModel(m, fallbackSeriesKey, viewMode, rawIndex, activeModelSeries) != null)
        )
      );
      const selected = labelsWithData.slice(-8);

      const setSpiderLegend = processType === 'TTL' ? setSpiderLegendTTL : setSpiderLegendMAIN;

      if (selected.length === 0) {
        window.Plotly.purge(elId);
        setSpiderLegend([]);
        return;
      }

      const hasRealDataAtAnySelectedPeriod = (m: string) => selected.some(({ label, rawIndex }) => {
        if (getSeriesValueForModel(m, seriesKey, viewMode, rawIndex, activeModelSeries) != null) return true;
        if (viewMode === 'day') {
          const mm = parseInt(label.split('/')[0], 10);
          const monthIdx = mm - 1;
          if (Number.isFinite(monthIdx) && monthIdx >= 0 && monthIdx < activeMonthLabels.length) {
            if (getSeriesValueForModel(m, seriesKey, 'month', monthIdx, activeModelSeries) != null) return true;
          }
        }
        // EPCC (rty-ttl-spider-blank): chấp nhận RTY_MAIN làm proxy cho TTL
        // khi file không có cột RTY_TTL riêng.
        if (processType === 'TTL' && fallbackSeriesKey &&
            getSeriesValueForModel(m, fallbackSeriesKey, viewMode, rawIndex, activeModelSeries) != null) return true;
        return false;
      });
      const models = modelsAll.filter(hasRealDataAtAnySelectedPeriod);

      if (models.length === 0) {
        window.Plotly.purge(elId);
        setSpiderLegend([]);
        return;
      }
      const theta = [...models, models[0]];

      const traces = selected.map(({ label, rawIndex }, i) => {
        // Dùng hàm có fallback để tránh null bị Plotly vẽ về tâm (r=0). Vì
        // `models` ở trên đã lọc bỏ model không hề có data thật, mức "Tổng
        // hợp" (mức 3) ở đây chỉ còn bù cho các mốc LẺ TẺ thiếu của model đã
        // đủ điều kiện — không còn tạo cả 1 model từ hư không.
        const points = models.map(m => getSpiderValueWithFallback(m, processType, viewMode, rawIndex, label, activeModelSummary, activeModelSeries));
        const r = points.map(p => p.value != null ? parseFloat((p.value * 100).toFixed(2)) : null);
        const noteText = points.map(p => SPIDER_SOURCE_LABEL[p.source]);
        // FIX (radar-label-no-marker, EPCC): trước đây gắn thêm dấu "*" vào
        // cuối số khi giá trị là rơi mức (không đúng mốc thật) — ở kích
        // thước chữ nhỏ trên radar, dấu "*" bị đọc nhầm thành dấu "/". Bỏ
        // hẳn ký hiệu này khỏi nhãn hiển thị trên chart; thông tin "rơi mức
        // hay không" vẫn còn đầy đủ trong tooltip (customdata/noteText) khi
        // hover, không mất thông tin, chỉ không in trực tiếp lên chart nữa.
        const dataLabels = points.map(p => p.value == null ? '' : `${(p.value * 100).toFixed(1)}`);

        r.push(r[0]);               // khép kín đa giác mạng nhện
        noteText.push(noteText[0]);
        // Điểm cuối là lặp lại điểm đầu (đóng vòng) — bỏ trống label để
        // không chồng số lên chính nó, giống PerCapitaTab.
        dataLabels.push('');

        const color = RADAR_COLORS[i % RADAR_COLORS.length];

        return {
          type: 'scatterpolar' as const,
          r,
          theta,
          fill: 'toself' as const,
          fillcolor: color + '1A', // ~10% opacity, giống cách PerCapitaTab tô nền đa giác
          mode: 'lines+markers+text' as const,
          marker: { color, size: 5 },
          line: { color, width: 2 },
          text: dataLabels,
          texttemplate: '%{text}',
          textposition: 'top center' as const,
          textfont: { color, size: 10.8 }, // +20% so với 9 gốc
          customdata: noteText,
          name: label,
          hovertemplate: `<b>${label}</b><br>%{theta}: %{r:.2f}%<br>%{customdata}<extra></extra>`
        };
      });

      // FIX (radar-legend-to-header, EPCC): trước đây legend nằm NGANG phía
      // trên bên trong ô chart (Plotly legend, orientation:'h', y:1.22) —
      // vẫn chiếm mất khoảng trống dọc của chart, khiến biểu đồ mạng nhện
      // nhỏ hơn mức có thể. Chuyển hẳn legend ra thanh TIÊU ĐỀ của panel
      // (giống 4 biểu đồ cột+đường bên dưới) qua state spiderLegendTTL/MAIN,
      // để nhường 100% chiều cao ô chart cho vòng tròn radar.
      setSpiderLegend(selected.map(({ label }, i) => ({ label, color: RADAR_COLORS[i % RADAR_COLORS.length] })));

      // Tự động tính toán range radar axis để không bị tràn
      const allR = traces.flatMap(t => t.r).filter((v): v is number => v != null);
      const minR = allR.length ? Math.max(0, Math.min(...allR)) : 0;
      const maxR = allR.length ? Math.min(100, Math.max(...allR)) : 100;
      const rMin = minR > 85 ? 85 : minR > 75 ? 75 : minR > 50 ? 50 : 0;
      const rMax = maxR < 99.8 ? 100 : 101;

      window.Plotly.newPlot(elId, traces, {
        paper_bgcolor: 'transparent',
        plot_bgcolor: 'transparent',
        font: { color: fontColor, size: 11.4 }, // +20% so với 9.5 gốc
        // FIX (radar-legend-to-header, EPCC): tắt hẳn legend nội bộ của
        // Plotly (đã chuyển ra thanh tiêu đề ở trên) và thu margin về mức
        // tối thiểu ở cả 4 phía để vòng tròn radar được vẽ TO NHẤT có thể
        // trong ô chart.
        margin: { t: 20, r: 20, b: 20, l: 20 },
        showlegend: false,
        polar: {
          bgcolor: 'transparent',
          radialaxis: {
            visible: true,
            range: [rMin, rMax],
            tickfont: { size: 9.6, color: fontColor }, // +20% so với 8 gốc
            gridcolor: gridColor,
            angle: 90,
            tickangle: 90
          },
          angularaxis: {
            tickfont: { size: 10.8, color: fontColor }, // +20% so với 9 gốc
            gridcolor: gridColor
          }
        }
      }, { displayModeBar: false, responsive: true });
    };

    // Vẽ 2 biểu đồ mạng nhện mới
    plotSpiderPanel('rtySpiderTTL', 'TTL');
    plotSpiderPanel('rtySpiderMAIN', 'MAIN');

  // EPCC (rty-chart-static-data-bug): thêm dynamicRtyData vào dependency
  // array để useEffect vẽ lại biểu đồ ngay khi có dữ liệu mới từ upload.
  }, [plotlyReady, activeTab, isLightMode, t, viewMode, startDate, endDate, selectedModel, xs, idxs, dynamicRtyData]);

  return (
    <div className={`second-dashboard rty-dashboard theme-${theme}`} style={{ padding: '0 16px 16px 16px', height: '100%', overflowY: 'auto', boxSizing: 'border-box' }}>
      {/* FIX (top-gap-zero, EPCC): trước đây padding: '16px' áp cho CẢ 4
          phía — tạo khoảng trống trắng phía trên cùng giữa mép viewport và
          thanh header ("HIỆU SUẤT RTY"). Bỏ riêng padding-top (còn 0), giữ
          nguyên trái/phải/dưới 16px — toàn bộ nội dung (thanh giờ, tiêu đề,
          2 tab, thanh filter, lưới biểu đồ) dịch chuyển sát lên top=0. */}
      <style>{`
        .rty-waiting-icon { animation: rty-pulse 1.6s ease-in-out infinite; display: inline-block; }
        @keyframes rty-pulse { 0%, 100% { opacity: 0.45; transform: scale(1); } 50% { opacity: 1; transform: scale(1.08); } }
        /* Cân chỉnh giống hệt Mục 2: không dùng sticky đè lên sidebar, không margin âm lệch lề */
        .rty-dashboard .dashboard-header-grid {
          background: #2F3A1D;
          border-radius: 14px;
          padding: 10px 16px;
          border: 1px solid rgba(0,0,0,0.18);
          box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        /* FIX (unify-header-color-C0EF6A): đồng hồ + tiêu đề trong khung
           header-grid (nền #2F3A1D) đổi sang màu cố định '#C0EF6A' theo
           yêu cầu — cùng tông với màu nhãn filter bên dưới, đảm bảo tương
           phản tốt trên nền xanh đậm ở mọi theme. */
        .rty-dashboard .dashboard-header-left { color: #C0EF6A !important; }
        .rty-dashboard .dashboard-header-title { color: #C0EF6A !important; }
        /* Chuẩn hóa khung 2 tab: copy NGUYÊN VĂN từ TargetActualDashboard
           (.second-dashboard .tab-container / .tab-btn) để đồng nhất tuyệt
           đối màu sắc + cỡ chữ giữa 2 dashboard. */
        .rty-dashboard .tab-container {
          display: flex;
          gap: 10px;
          margin-bottom: 8px;
        }
        .rty-dashboard .tab-btn {
          padding: 8px 18px;
          border-radius: 8px;
          font-weight: 700;
          font-size: 14px;
          cursor: pointer;
          transition: all 0.2s ease;
          border: 1px solid var(--border);
          background: rgba(30, 41, 59, 0.2);
          color: var(--text-2);
          display: inline-flex;
          align-items: center;
          gap: 6px;
          white-space: nowrap;
        }
        .rty-dashboard .tab-btn.active {
          background: linear-gradient(135deg, #1d4ed8 0%, #3b82f6 100%);
          color: #ffffff;
          border-color: #3b82f6;
          box-shadow: 0 4px 12px rgba(59, 130, 246, 0.35);
        }
        /* Để các tiêu đề sát viền trên cùng (top=0) trong khung biểu đồ */
        .rty-dashboard .chart-panel {
          padding: 0 !important;
          overflow: hidden;
        }
        .rty-dashboard .card-header-styled {
          border-radius: 12px 12px 0 0 !important;
          margin: 0 !important;
        }
        .rty-dashboard .chart-holder {
          padding: 8px 16px 12px 16px;
          box-sizing: border-box;
        }
      `}</style>

      {/* ── Header ── */}
      <div className="dashboard-header-grid">
        <div className="dashboard-header-left" style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-2)', fontWeight: 700, whiteSpace: 'nowrap' }}>
          <span aria-hidden="true">🕐</span>
          {formattedTime}
        </div>
        <h1 className="dashboard-header-title">{t.title}</h1>
        <div className="dashboard-header-right" />
      </div>

      {/* ── 3 Tab (thêm RTY Total, chuyển từ Mục 5 sang) ── */}
      <div className="tab-container">
        <button className={`tab-btn ${activeTab === 'summary' ? 'active' : ''}`} onClick={() => setActiveTab('summary')}>
          📊 {t.tab1}
        </button>
        {/* EPCC (rty-total-move-to-muc4): tab mới — đọc trực tiếp cột RTY %
            từ Test6.xlsx qua RtyTotalTab, có nút "Tải Excel RTY" RIÊNG bên
            trong, không dùng chung nút "Tải Excel" (đang disabled) của
            toolbar bên dưới. */}
        <button className={`tab-btn ${activeTab === 'rtyTotal' ? 'active' : ''}`} onClick={() => setActiveTab('rtyTotal')}>
          🎯 RTY Total
        </button>
        <button className={`tab-btn ${activeTab === 'merged' ? 'active' : ''}`} onClick={() => setActiveTab('merged')}>
          💼 {t.tab2}
        </button>
      </div>

      {/* EPCC (rty-total-move-to-muc4): ẩn toolbar filter Ngày/Model/Xem theo
          + nút "Tải Excel" (đang disabled, thuộc dữ liệu tĩnh Test4.xlsx)
          khi đang ở tab RTY Total — tab đó có toolbar RIÊNG của chính nó
          (đọc động Test6.xlsx), không liên quan gì tới toolbar này. */}
      {activeTab !== 'rtyTotal' && (
      <div className="topbar-dash" style={{
        display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '10px',
        background: '#2F3A1D', borderRadius: '14px', padding: '10px 14px',
        border: '1px solid rgba(0,0,0,0.18)', boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          <div style={{ width: '170px', flexShrink: 0 }} />
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px' }}>
            <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.startDate}</span>
            <span style={{ width: '130px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{t.endDate}</span>
            <span style={{ width: '160px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase' }}>{t.model}</span>
            <span style={{ width: '180px', textAlign: 'center', fontSize: '12px', fontWeight: 700, color: filterLabelColor, textTransform: 'uppercase' }}>{t.viewBy}</span>
          </div>
          <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
            <div style={{ width: '120px' }} />
          </div>
        </div>

        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          {/* FIX (move-showing-model-badge-inside-toolbar): dòng "Đang hiển
              thị: ⭐ MODEL (Hiệu suất tốt nhất)" trước đây nằm NGOÀI khung
              filter (phía trên khu vực 4 thẻ KPI) — theo yêu cầu, chuyển
              hẳn vào TRONG khung filter xanh đậm (#2F3A1D), đặt ở ô trống
              bên trái dòng 2 (cùng hàng với ô input ngày/model/xem theo)
              để gọn trong 1 khung duy nhất thay vì tách rời bên ngoài. */}
          <div style={{ width: '170px', minWidth: '170px', flexShrink: 0, display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: '4px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ fontSize: '11px', fontWeight: 600, color: filterLabelColor, whiteSpace: 'nowrap' }}>{t.kpiShowingModel}:</span>
              <span style={{
                display: 'inline-flex', alignItems: 'center', gap: '4px', padding: '3px 8px',
                borderRadius: '999px', fontWeight: 700, fontSize: '11px', whiteSpace: 'nowrap',
                background: selectedModel === activeBestModel ? 'rgba(16,185,129,0.22)' : 'rgba(14,165,233,0.22)',
                color: selectedModel === activeBestModel ? '#34d399' : '#38bdf8',
              }}>
                {selectedModel === activeBestModel && '⭐ '}
                {selectedModel || t.allOption}
                {selectedModel === activeBestModel && ` (${t.kpiBestLabel})`}
              </span>
            </div>
            <span style={{ fontSize: '11px', fontWeight: 600, color: '#22c55e', whiteSpace: 'nowrap' }}>
              📅 {lang === 'vi' ? 'Dữ liệu cập nhật đến' : lang === 'ko' ? '데이터 업데이트 기준일' : 'Data updated to'}: {activeLastUpdate}
            </span>
          </div>
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'center', flex: 1, margin: '0 24px', alignItems: 'center' }}>
            <input
              type="date"
              value={startDate}
              min={activeDataMinDate}
              max={activeDataMaxDate}
              onChange={e => setStartDate(e.target.value)}
              className="filter-date-input"
              style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
            />
            <input
              type="date"
              value={endDate}
              min={activeDataMinDate}
              max={activeDataMaxDate}
              onChange={e => setEndDate(e.target.value)}
              className="filter-date-input"
              style={{ width: '130px', minWidth: '130px', height: '38px', boxSizing: 'border-box', textAlign: 'center', padding: '8px 4px' }}
            />
            <select value={selectedModel} onChange={e => setSelectedModel(e.target.value)}
              className="header-filter-select" style={{ width: '160px', height: '38px' }}>
              <option value="">{t.allOption}</option>
              {activeModelSummarySorted.map(m => (
                <option key={m.model} value={m.model}>
                  {m.model === activeBestModel ? `⭐ ${m.model}` : m.model}
                </option>
              ))}
            </select>
            <div style={{ display: 'flex', gap: '0px', height: '38px', width: '180px', flexShrink: 0 }}>
              {(['day', 'week', 'month'] as const).map(mode => (
                <button key={mode} onClick={() => setViewMode(mode)}
                  style={{
                    flex: 1, padding: '6px 0', fontSize: '13px', fontWeight: 600,
                    borderRadius: mode === 'day' ? '6px 0 0 6px' : mode === 'month' ? '0 6px 6px 0' : '0',
                    border: '1px solid rgba(0,0,0,0.18)',
                    borderRight: mode !== 'month' ? 'none' : '1px solid rgba(0,0,0,0.18)',
                    background: viewMode === mode ? '#2e7d8c' : 'rgba(255,255,255,0.55)',
                    color: viewMode === mode ? '#ffffff' : '#7A5A2E',
                    cursor: 'pointer', height: '100%', display: 'inline-flex',
                    alignItems: 'center', justifyContent: 'center', whiteSpace: 'nowrap',
                  }}>
                  {mode === 'day' ? t.day : mode === 'week' ? t.week : t.month}
                </button>
              ))}
            </div>
          </div>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexShrink: 0 }}>
            <input 
              ref={fileInputRef}
              type="file" 
              accept=".xlsx, .xls" 
              style={{ display: 'none' }} 
              onChange={async (e) => {
                const file = e.target.files?.[0];
                if (!file) return;
                onSyncProgress?.({ bucket: 'RTY', done: 0, total: 0 });
                const reader = new FileReader();
                reader.onload = async (evt) => {
                  try {
                    const data = new Uint8Array(evt.target!.result as ArrayBuffer);
                    // EPCC (rty-upload-xlsx-global-missing): dùng thẳng XLSX đã import
                    // tĩnh ở đầu file, không cần `await import('xlsx')` nữa (trước đây
                    // biến này chỉ tồn tại cục bộ trong callback, không giúp gì cho
                    // parseWorkbookToRtySummaryData vì hàm đó đọc từ global).
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                    const dynamicResult = parseWorkbookToRtySummaryData(workbook);
                    if (dynamicResult && dynamicResult.dayLabels.length > 0) {
                      setDynamicRtyData(dynamicResult);
                      // EPCC (rty-empty-array-falsy-bug): nếu modelSummary rỗng sau parse
                      // (vẫn xảy ra khi file hoàn toàn không có Process/RTY column đúng chuẩn),
                      // KHÔNG đặt selectedModel về '' (gây 0/0 KPI) — giữ nguyên model đang chọn.
                      // Chỉ cập nhật selectedModel khi thực sự có model mới từ upload.
                      if (dynamicResult.modelSummary.length > 0) {
                        const sortedBest = [...dynamicResult.modelSummary].sort((a, b) => (b.ttl.actual - b.ttl.target) - (a.ttl.actual - a.ttl.target))[0]?.model || activeBestModel;
                        setSelectedModel(sortedBest);
                      }
                      setStartDate(dynamicResult.dataMinDate);
                      setEndDate(dynamicResult.dataMaxDate);
                      // Chỉ lưu cache khi có model data thực sự (tránh ghi đè cache hợp lệ
                      // bằng cache rỗng rồi gây 0/0 KPI ở lần reload tiếp theo).
                      if (dynamicResult.modelSummary.length > 0) {
                        await idbSetCacheSummary(IDB_KEY_RTY_SUMMARY_DATA, JSON.stringify(dynamicResult));
                      }
                    } else {
                      // EPCC (rty-upload-xlsx-global-missing): trước đây nhánh này im
                      // lặng bỏ qua khiến người dùng tưởng đã cập nhật xong. Giờ báo rõ
                      // để biết file đọc được nhưng không tìm thấy cột dữ liệu hợp lệ
                      // (cần cột Model/Process + Date/Period + RTY, hoặc các cột dạng
                      // ngày làm tiêu đề).
                      alert(lang === 'vi'
                        ? 'Đã đọc được tệp nhưng không tìm thấy dữ liệu RTY hợp lệ. Vui lòng kiểm tra lại cột Model/Process/Date/RTY trong file.'
                        : 'File was read but no valid RTY data rows were found. Please check the Model/Process/Date/RTY columns.');
                    }

                    if (onFileSelected) {
                      onFileSelected(file, workbook);
                    }
                  } catch (err) {
                    console.error('Lỗi đọc tệp Excel (RTY upload):', err);
                    alert('Lỗi đọc tệp Excel!');
                  } finally {
                    setTimeout(() => onSyncProgress?.(null), 1000);
                  }
                };
                reader.readAsArrayBuffer(file);
                e.target.value = '';
              }} 
            />
            <NeonButton
              className="btn btn-outline btn-sm"
              onClick={() => fileInputRef.current?.click()}
              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px', height: '38px', minWidth: '140px', padding: '0 14px', boxSizing: 'border-box', fontSize: '13px', cursor: 'pointer' }}
            >
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" width="15" height="15" style={{ flexShrink: 0 }}>
                <path d="M21 12a9 9 0 0 1-9 9c-2.52 0-4.93-1-6.74-2.74L3 16" />
                <path d="M3 12a9 9 0 0 1 9-9c2.52 0 4.93 1 6.74 2.74L21 8" />
                <path d="M3 16v5h5" />
                <path d="M16 3h5v5" />
              </svg>
              <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {lang === 'vi' ? 'Tải tệp lên' : lang === 'ko' ? '파일 업로드' : 'Upload File'}
              </span>
            </NeonButton>
          </div>
        </div>
      </div>
      )}

      {/* ══════════════ TAB MỚI: RTY TOTAL (chuyển từ Mục 5 sang) ══════════════ */}
      {activeTab === 'rtyTotal' && <RtyTotalTab theme={theme} lang={lang} onSyncProgress={onSyncProgress} />}

      {/* ══════════════ TAB 1: TÌNH HÌNH RTY ══════════════ */}
      {activeTab === 'summary' && (
        <>
          {/* 4 thẻ KPI — số liệu thật, tính từ MODEL_SUMMARY (đã lọc theo Model).
              Bộ màu chuẩn hóa theo đúng mục 2 (Báo cáo doanh số): xanh dương /
              xanh lá / tím / cam, cùng size & khoảng cách (kpi-grid, gap 12px).
              Mặc định hiển thị Model có hiệu suất tốt nhất (BEST_MODEL) — badge
              nhỏ bên dưới cho biết đang xem model nào, chọn model khác ở
              dropdown "Model" phía trên để số liệu nhảy theo. */}
          {/* Badge "Đang hiển thị: ⭐ MODEL (Hiệu suất tốt nhất)" đã được
              chuyển vào TRONG khung filter (topbar-dash) phía trên — xem
              FIX (move-showing-model-badge-inside-toolbar). Không lặp lại
              ở đây nữa để tránh hiển thị 2 lần. */}
          <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(4,1fr)', gap: '12px', marginBottom: '12px', width: '100%' }}>
            <div className="kpi-card" style={{ borderLeft: '4px solid #0ea5e9', background: 'linear-gradient(135deg, #0ea5e91a 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#0ea5e9" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 16, height: 16, flexShrink: 0 }}>
                  <path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" />
                  <polyline points="3.27 6.96 12 12.01 20.73 6.96" /><line x1="12" y1="22.08" x2="12" y2="12" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi1}</div>
              </div>
              <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px', flexWrap: 'wrap' }}>
                <div className="kpi-card-value" style={{ marginBottom: 0 }}>{(kpi.actual * 100).toFixed(2)}%</div>
                <div className="kpi-card-target">{t.kpi1Target}: {(kpi.target * 100).toFixed(1)}%</div>
              </div>
            </div>

            <div className="kpi-card" style={{ borderLeft: '4px solid #10b981', background: 'linear-gradient(135deg, #10b9811a 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#10b981" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 16, height: 16, flexShrink: 0 }}>
                  <circle cx="12" cy="12" r="10" /><circle cx="12" cy="12" r="6" /><circle cx="12" cy="12" r="2" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi2}</div>
              </div>
              <div className="kpi-card-value" style={{ marginBottom: 0 }}>{kpi.achieveRate.toFixed(1)}%</div>
            </div>

            <div className="kpi-card" style={{ borderLeft: '4px solid #8b5cf6', background: 'linear-gradient(135deg, #8b5cf61a 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#8b5cf6" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 16, height: 16, flexShrink: 0 }}>
                  <path d="M9 11l3 3L22 4" /><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi3}</div>
              </div>
              <div className="kpi-card-value" style={{ marginBottom: 0 }}>{kpi.meetingCount}/{kpi.totalCount}</div>
            </div>

            <div className="kpi-card" style={{ borderLeft: '4px solid #f59e0b', background: 'linear-gradient(135deg, #f59e0b1a 0%, rgba(30,41,59,0.4) 100%)' }}>
              <div className="kpi-card-header" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <svg viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ width: 16, height: 16, flexShrink: 0 }}>
                  <polyline points="23 6 13.5 15.5 8.5 10.5 1 18" /><polyline points="17 6 23 6 23 12" />
                </svg>
                <div className="kpi-card-label" style={{ marginBottom: 0 }}>{t.kpi4}</div>
              </div>
              <div className="kpi-card-value" style={{ marginBottom: 0, color: kpi.avgGapPct >= 0 ? '#10b981' : '#ef4444' }}>{kpi.avgGapPct >= 0 ? '+' : ''}{kpi.avgGapPct.toFixed(2)}pp</div>
            </div>
          </div>

          {/* Lưới 2×2 = 4 biểu đồ TTL/MAIN/SUB1/SUB2, đúng bố cục ảnh tham
              chiếu (ảnh 1). Model tham chiếu: SO3560 — model duy nhất có đủ
              dữ liệu thật cho cả Tháng/Tuần/Ngày trong Test4.xlsx. Trục X
              của cả 4 biểu đồ đổi theo "Xem theo" (Ngày/Tuần/Tháng) và luôn
              hiển thị tối đa 8 giá trị GẦN NHẤT trong khoảng Ngày bắt đầu/
              kết thúc đã chọn — không còn ô "chờ dữ liệu" vì cả 4 dòng đều
              có số liệu thật ở mọi cấp chi tiết. */}

          <div className="chart-grid" style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gridAutoRows: '370px', gap: '12px' }}>
            {([
              /* ── Chart 1 & 2: Biểu đồ mạng nhện (đưa lên đầu) ── */
              {
                id: 'rtySpiderTTL',
                rowLabel: 'TTL RTY THEO MODEL',
                accent: '#8b5cf6',
                bg: isLightMode ? '#d6c6fc' : 'rgba(139,92,246,0.14)',
                note: 'Biểu đồ so sánh CẢ 10 MODEL theo từng giai đoạn (không lọc theo Model đang chọn ở trên). Giá trị rơi mức (TB Tháng/Tổng hợp) do giai đoạn đó model không có số liệu thật đúng mốc — xem chi tiết khi hover vào từng điểm.',
                hideModelTag: true,
                // FIX (radar-legend-to-header, EPCC): legend giờ lấy từ
                // state ĐỘNG spiderLegendTTL (set trong plotSpiderPanel mỗi
                // lần vẽ lại) thay vì mảng rỗng — hiển thị ngay trên thanh
                // tiêu đề giống 4 biểu đồ bên dưới, nhường trọn ô chart cho
                // vòng tròn radar.
                legends: spiderLegendTTL.map(lg => ({ type: 'line' as const, label: lg.label, color: lg.color }))
              },
              {
                id: 'rtySpiderMAIN',
                rowLabel: 'MAIN RTY THEO MODEL',
                accent: '#06b6d4',
                bg: isLightMode ? '#a8e5f0' : 'rgba(6,182,212,0.14)',
                note: 'Biểu đồ so sánh CẢ 10 MODEL theo từng giai đoạn (không lọc theo Model đang chọn ở trên). Giá trị rơi mức (TB Tháng/Tổng hợp) do giai đoạn đó model không có số liệu thật đúng mốc — xem chi tiết khi hover vào từng điểm.',
                hideModelTag: true,
                // FIX (radar-legend-to-header, EPCC): tương tự panel TTL —
                // legend lấy từ state ĐỘNG spiderLegendMAIN.
                legends: spiderLegendMAIN.map(lg => ({ type: 'line' as const, label: lg.label, color: lg.color }))
              },
              /* ── Chart 3-6: Biểu đồ hỗn hợp cột+đường ── */
              {
                id: 'rtyChartTTL',
                rowLabel: 'TTL RTY',
                accent: '#f59e0b',
                bg: isLightMode ? '#fcddaa' : 'rgba(255,255,255,0.05)',
                note: undefined as string | undefined,
                hideModelTag: undefined as boolean | undefined,
                legends: [
                  { type: 'bar', label: 'RTY Sub1', color: '#1565C0' },
                  { type: 'bar', label: 'RTY Sub2', color: '#ef4444' },
                  { type: 'line', label: 'RTY Main', color: tealAccent },
                  { type: 'dashed', label: 'RTY TTL', color: '#8b5cf6' },
                  { type: 'dotted', label: 'Target RTY TTL 96.4%', color: '#ef4444' },
                ]
              },
              {
                id: 'rtyChartMAIN',
                rowLabel: 'MAIN',
                accent: '#f97316',
                bg: isLightMode ? '#fdcead' : 'rgba(249,115,22,0.14)',
                note: undefined as string | undefined,
                hideModelTag: undefined as boolean | undefined,
                legends: [
                  { type: 'bar', label: 'RTY Main', color: '#1565C0' },
                  { type: 'bar', label: 'Main FVI Final', color: '#ef4444' },
                  { type: 'line', label: 'Main Assy', color: tealAccent },
                  { type: 'line', label: 'Main Driving', color: '#f59e0b' },
                  { type: 'line', label: 'Main Tilt', color: '#a78bfa' },
                  { type: 'dashed', label: 'RTY Main (line)', color: '#ec4899' },
                  { type: 'dotted', label: 'Target Main 98.3%', color: '#ef4444' },
                ]
              },
              {
                id: 'rtyChartSUB1',
                rowLabel: 'SUB1',
                accent: '#10b981',
                bg: isLightMode ? '#abe7d3' : 'rgba(16,185,129,0.14)',
                note: undefined as string | undefined,
                hideModelTag: undefined as boolean | undefined,
                legends: [
                  { type: 'bar', label: 'S1 FPCB VI', color: '#1565C0' },
                  { type: 'bar', label: 'S1 FVI', color: '#ef4444' },
                  { type: 'line', label: 'RTY Sub1', color: tealAccent },
                  { type: 'dotted', label: 'Target Sub1 99.2%', color: '#ef4444' },
                ]
              },
              {
                id: 'rtyChartSUB2',
                rowLabel: 'SUB2',
                accent: '#3b82f6',
                bg: isLightMode ? '#bad3fc' : 'rgba(59,130,246,0.14)',
                note: undefined as string | undefined,
                hideModelTag: undefined as boolean | undefined,
                legends: [
                  { type: 'bar', label: 'S2 Hook Bonding', color: '#1565C0' },
                  { type: 'bar', label: 'S2 Oven Cure', color: '#ef4444' },
                  { type: 'line', label: 'RTY Sub2', color: tealAccent },
                  { type: 'dotted', label: 'S2 INDEX', color: '#f59e0b' },
                  { type: 'dotted', label: 'Target Sub2 98.8%', color: '#ef4444' },
                ]
              },
            ] as const).map(panel => (
              <div key={panel.id} className="panel chart-panel">
                <div className="card-header-styled" style={{
                  background: panel.bg, borderLeft: `4px solid ${panel.accent}`,
                  color: isLightMode ? '#1f2937' : 'var(--text-1)',
                  display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                  padding: '6px 10px', fontWeight: 700, fontSize: '12px', textTransform: 'uppercase',
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {/* FIX (remove-broken-note-icon, EPCC): icon tròn "i" (title
                        tooltip) trước đây bị render lỗi thành dấu X đỏ ở môi
                        trường build của người dùng (Antigravity/Vercel) — có
                        thể do font-icon override toàn cục trong theme. Bỏ hẳn
                        icon nổi, gắn tooltip `title` thẳng vào tên tiêu đề —
                        hover vẫn xem được ghi chú, không còn ký hiệu lạ trên UI. */}
                    <span title={panel.note}>{panel.rowLabel}</span>
                    {selectedModel && !panel.hideModelTag && (
                      <span style={{ fontSize: '10.5px', opacity: 0.75, textTransform: 'none' }}>({selectedModel})</span>
                    )}
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '4px', textTransform: 'none', fontWeight: 500, flexWrap: 'wrap', justifyContent: 'end' }}>
                    {panel.legends.map((lg, idx) => (
                      <RtyLegendItem key={idx} type={lg.type} label={lg.label} color={lg.color} />
                    ))}
                  </div>
                </div>
                <div className="chart-holder" style={{ height: '100%' }}>
                  <div id={panel.id} style={{ width: '100%', height: '100%', minHeight: 320 }} />
                </div>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ══════════════ TAB 2: RTY & CHI TIẾT (bảng thật, đã bỏ cột SL/Khách hàng) ══════════════ */}
      {activeTab === 'merged' && (
        <div className="panel merged-fill-panel" style={{ display: 'flex', flexDirection: 'column' }}>
          <div className="table-container merged-fill-table-container" style={{ display: 'flex', flexDirection: 'column', paddingTop: 0 }}>
            <div ref={tableScrollRef} className="table-scroll" style={{ overflowY: 'auto', overflowX: 'auto', maxHeight: `${tableScrollMaxHeight}px`, padding: '0 15px' }}>
              <table className="stat-table" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, marginTop: 0 }}>
                <thead style={{ position: 'sticky', top: 0, zIndex: 20 }}>
                  <tr style={{ background: 'var(--surface)', position: 'sticky', top: 0, zIndex: 20 }}>
                    <th colSpan={4} style={{ borderBottom: 'none', background: 'var(--surface)' }} />
                    <th colSpan={3} style={{ textAlign: 'center', background: isLightMode ? 'rgba(29, 78, 216, 0.95)' : 'rgba(29, 78, 216, 0.85)', color: '#fff', fontWeight: 800, fontSize: '16.5px', letterSpacing: '0.05em', textTransform: 'uppercase', padding: '10px 14px' }}>
                      {t.tblGroup2}
                    </th>
                  </tr>
                  <tr style={{ background: 'var(--surface)', position: 'sticky', top: '43px', zIndex: 20 }}>
                    {[t.colModel, t.colProcess, t.colType, t.colPeriod, t.colRtyTarget, t.colRtyActual, t.colRtyGap].map((h, i) => (
                      <th key={h} style={{ fontSize: '16.5px', fontWeight: 600, padding: '10px 12px', textAlign: i < 4 ? 'left' : 'right', color: 'var(--tbl-head-color)', background: 'var(--surface)' }}>
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pageRows.length === 0 && (
                    <tr>
                      <td colSpan={7} style={{ textAlign: 'center', padding: '28px', color: 'var(--text-3)', fontSize: '13.5px', fontWeight: 500 }}>
                        {t.tableEmpty}
                      </td>
                    </tr>
                  )}
                  {pageRows.map((r, i) => {
                    const gap = Math.round((r.actual - r.target) * 1000) / 10;
                    return (
                      <tr key={`${r.model}-${r.process}-${r.type}-${r.period}-${i}`} style={{ borderBottom: '1px solid var(--border)' }}>
                        <td style={{ padding: '8px 12px', fontWeight: 600 }}>{r.model}</td>
                        <td style={{ padding: '8px 12px' }}>{r.process}</td>
                        <td style={{ padding: '8px 12px' }}>{r.type}</td>
                        <td style={{ padding: '8px 12px' }}>{r.period}</td>
                        <td style={{ padding: '8px 12px', textAlign: 'right' }}>{(r.target * 100).toFixed(1)}%</td>
                        <td style={{ padding: '8px 12px', textAlign: 'right' }}>{(r.actual * 100).toFixed(2)}%</td>
                        <td style={{ padding: '8px 12px', textAlign: 'right', color: gap >= 0 ? '#10b981' : '#ef4444', fontWeight: 700 }}>
                          {gap >= 0 ? '+' : ''}{gap.toFixed(1)}pp
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div ref={paginationBarRef} className="pagination-bar" style={{ padding: '10px 15px', background: 'var(--surface)', borderTop: '1px solid var(--border-soft)', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '12px', flexWrap: 'wrap', flexShrink: 0 }}>
            <div className="pagination-info">
              {t.showing} {filteredRows.length === 0 ? 0 : (page - 1) * rowsPerPage + 1}-{Math.min(page * rowsPerPage, filteredRows.length)} {t.of} {filteredRows.length} {t.rows}
            </div>
            <div className="pagination-controls">
              <button className="page-btn" disabled={page === 1} onClick={() => setPage(1)}>«</button>
              <button className="page-btn" disabled={page === 1} onClick={() => setPage(p => Math.max(1, p - 1))}>‹</button>
              <button className="page-btn active">{page}</button>
              <button className="page-btn" disabled={page === totalPages} onClick={() => setPage(p => Math.min(totalPages, p + 1))}>›</button>
              <button className="page-btn" disabled={page === totalPages} onClick={() => setPage(totalPages)}>»</button>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <span style={{ fontSize: 11.5, color: 'var(--text-2)' }}>{t.rowsPerPage}</span>
              <select className="page-size-select" value={rowsPerPage} onChange={e => setRowsPerPage(Number(e.target.value))}>
                {[10, 25, 31, 50, 100].map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default RtyDashboard;
