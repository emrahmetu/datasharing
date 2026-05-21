import { formatNumber, formatDelta } from '../utils/format.js';

export default function SummaryCards({ summary }) {
  if (!summary) return null;
  const { balances, totals, period } = summary;

  const cards = [
    { label: 'Dönem', value: `${period.start ?? '-'} → ${period.end ?? '-'}`, sub: `${period.days} gün` },
    { label: 'Açılış Bakiye TL', value: formatNumber(balances.beginningTL, { compact: true, currency: 'TL' }) },
    { label: 'Kapanış Bakiye TL', value: formatNumber(balances.closingTL, { compact: true, currency: 'TL' }), delta: balances.netChangeTL, currency: 'TL' },
    { label: 'Kapanış Bakiye USD', value: formatNumber(balances.closingUSD, { compact: true, currency: 'USD' }), delta: balances.netChangeUSD, currency: 'USD' },
    { label: 'Toplam Tahsilat TL', value: formatNumber(totals.collectionsTL, { compact: true, currency: 'TL' }) },
    { label: 'Toplam Ödeme TL', value: formatNumber(totals.paymentsTL, { compact: true, currency: 'TL' }) },
  ];

  return (
    <div className="cards">
      {cards.map((c, i) => (
        <div className="card" key={i}>
          <div className="label">{c.label}</div>
          <div className="value">{c.value}</div>
          {c.sub && <div className="muted" style={{ fontSize: 12, marginTop: 4 }}>{c.sub}</div>}
          {c.delta !== undefined && c.delta !== null && (
            <div className={`delta ${c.delta >= 0 ? 'pos' : 'neg'}`} style={{ fontSize: 13, marginTop: 4 }}>
              Net Δ {formatDelta(c.delta, c.currency)}
            </div>
          )}
        </div>
      ))}
    </div>
  );
}
