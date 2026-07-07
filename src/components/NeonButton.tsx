import type { ButtonHTMLAttributes, ReactNode } from 'react';

type NeonColor = 'default' | 'pink' | 'cyan' | 'violet' | 'amber';

interface NeonButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  children: ReactNode;
  neonColor?: NeonColor;
}

// Bọc quanh <button> hiện có: giữ nguyên mọi className/props (btn, btn-outline,
// btn-ghost, btn-sm, onClick...), chỉ thêm class "neon" + 4 <span> cần cho
// hiệu ứng viền sáng chạy quanh nút (xem neon-buttons.css).
//
// Cách dùng — đổi:
//   <button className="btn btn-outline" onClick={handleReset}>← Tải file khác</button>
// thành:
//   <NeonButton className="btn btn-outline" onClick={handleReset}>← Tải file khác</NeonButton>
//
// Đổi màu neon (mặc định lấy theo --primary):
//   <NeonButton className="btn" neonColor="cyan">...</NeonButton>
export function NeonButton({
  children,
  className = '',
  neonColor = 'default',
  ...rest
}: NeonButtonProps) {
  const colorClass = neonColor !== 'default' ? `neon-${neonColor}` : '';
  return (
    <button className={`${className} neon ${colorClass}`.trim().replace(/\s+/g, ' ')} {...rest}>
      <span className="neon-line" />
      <span className="neon-line" />
      <span className="neon-line" />
      <span className="neon-line" />
      {children}
    </button>
  );
}
