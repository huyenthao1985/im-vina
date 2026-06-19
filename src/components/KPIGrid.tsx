import type { KPIData } from '../types';
import { fmtVND, fmtNum, fmtPct } from '../utils';

interface KPIGridProps {
  kpi: KPIData;
  mapping: { revenueCol: string; costCol: string; categoryCol: string; currency?: string };
}

const NONE_VALUE = '__none__';

const KPICard = ({
  icon, label, value, badge, badgeType, color, bgColor
}: {
  icon: string;
  label: string;
  value: string;
  badge?: string;
  badgeType?: 'positive' | 'negative' | 'neutral';
  color: string;
  bgColor: string;
}) => (
  <div className="kpi-card" style={{
    border: `1px solid ${bgColor}`,
    borderLeft: `4px solid ${color}`,
    background: `linear-gradient(135deg, ${bgColor} 0%, rgba(30, 41, 59, 0.4) 100%)`
  }}>
    <div className="kpi-card-header">
      <div className="kpi-card-icon" style={{ background: bgColor, color: color, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '34px', height: '34px', borderRadius: '10px', fontSize: '18px', flexShrink: 0 }}>{icon}</div>
      <div className="kpi-card-label" style={{ fontSize: '13px', fontWeight: '600', color: 'var(--text-2)', textTransform: 'uppercase', letterSpacing: '.04em' }}>{label}</div>
    </div>
    <div className="kpi-card-value" style={{ fontFamily: 'monospace', fontSize: '28px', fontWeight: '500', color: 'var(--text-0)' }}>{value}</div>
    {badge && (
      <span className={`kpi-card-badge badge-${badgeType || 'neutral'}`} style={{ marginTop: '6px' }}>{badge}</span>
    )}
  </div>
);

export const KPIGrid: React.FC<KPIGridProps> = ({ kpi, mapping }) => {
  const hasRevenue = mapping.revenueCol !== NONE_VALUE;
  const hasCost = mapping.costCol !== NONE_VALUE && kpi.hasCost;
  const hasCategory = mapping.categoryCol !== NONE_VALUE;

  const cards = [];

  // Row count always shown
  cards.push(
    <KPICard key="rows"
      icon="📋" label="Tổng số dòng"
      value={fmtNum(kpi.totalRows)}
      badge={hasCategory ? `${fmtNum(kpi.uniqueCategories)} kênh` : undefined}
      badgeType="neutral"
      color="var(--purple)" bgColor="var(--purple-soft)"
    />
  );

  if (hasRevenue) {
    cards.push(
      <KPICard key="revenue"
        icon="💰" label="Tổng Doanh thu"
        value={fmtVND(kpi.totalRevenue, mapping.currency)}
        color="var(--cyan)" bgColor="var(--cyan-soft)"
      />
    );
  }

  if (hasCost) {
    cards.push(
      <KPICard key="cost"
        icon="📉" label="Tổng Chi phí"
        value={fmtVND(kpi.totalCost, mapping.currency)}
        color="var(--amber)" bgColor="var(--amber-soft)"
      />
    );
  }

  if (hasRevenue && hasCost) {
    const profitPositive = kpi.totalProfit >= 0;
    cards.push(
      <KPICard key="profit"
        icon={profitPositive ? '📈' : '📉'}
        label="Lợi nhuận"
        value={fmtVND(kpi.totalProfit, mapping.currency)}
        badge={profitPositive ? '▲ Có lời' : '▼ Lỗ'}
        badgeType={profitPositive ? 'positive' : 'negative'}
        color={profitPositive ? 'var(--emerald)' : 'var(--rose)'}
        bgColor={profitPositive ? 'var(--emerald-soft)' : 'var(--rose-soft)'}
      />
    );

    const roiPositive = kpi.roi >= 0;
    cards.push(
      <KPICard key="roi"
        icon="🎯" label="ROI"
        value={fmtPct(kpi.roi)}
        badge={roiPositive ? '▲ Hiệu quả' : '▼ Kém'}
        badgeType={roiPositive ? 'positive' : 'negative'}
        color={roiPositive ? 'var(--purple)' : 'var(--rose)'}
        bgColor={roiPositive ? 'var(--purple-soft)' : 'var(--rose-soft)'}
      />
    );
  } else if (hasRevenue && !hasCost) {
    // Show avg revenue
    const avg = kpi.totalRows > 0 ? kpi.totalRevenue / kpi.totalRows : 0;
    cards.push(
      <KPICard key="avg"
        icon="📊" label="TB Doanh thu / dòng"
        value={fmtVND(avg, mapping.currency)}
        color="var(--emerald)" bgColor="var(--emerald-soft)"
      />
    );
  }

  // Fill up to 5 with categories count if needed
  if (cards.length < 5 && hasCategory) {
    cards.push(
      <KPICard key="cats"
        icon="🏷️" label="Số kênh / danh mục"
        value={fmtNum(kpi.uniqueCategories)}
        color="var(--indigo)" bgColor="rgba(129,140,248,0.12)"
      />
    );
  }

  return (
    <div className="kpi-grid" style={{ gridTemplateColumns: `repeat(${Math.min(cards.length, 5)}, 1fr)` }}>
      {cards}
    </div>
  );
};
