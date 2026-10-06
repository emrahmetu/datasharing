import { formatNumber } from '../utils/format.js';

export default function Report({ data, summary }) {
  if (!data || !summary) return null;

  const exportCsv = () => {
    const rows = [['Item', ...data.dates]];
    for (const li of data.lineItems) {
      rows.push([li.label, ...li.values.map(v => v ?? '')]);
    }
    const csv = rows.map(r => r.map(cell => {
      const s = String(cell ?? '');
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'cashflow-report.csv'; a.click();
    URL.revokeObjectURL(url);
  };

  const exportJson = () => {
    const blob = new Blob([JSON.stringify({ data, summary }, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'cashflow-report.json'; a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div>
      <div className="controls" style={{ marginBottom: 16 }}>
        <button className="primary" onClick={exportCsv}>CSV indir</button>
        <button className="ghost" onClick={exportJson}>JSON indir</button>
      </div>

      <h3 style={{ marginBottom: 8 }}>Dönem özeti</h3>
      <ul className="muted" style={{ lineHeight: 1.8 }}>
        <li>Dönem: <strong style={{ color: 'var(--text)' }}>{summary.period.start} → {summary.period.end}</strong> ({summary.period.days} gün)</li>
        <li>Açılış TL → Kapanış TL: <strong style={{ color: 'var(--text)' }}>
          {formatNumber(summary.balances.beginningTL, { compact: true })} → {formatNumber(summary.balances.closingTL, { compact: true })}
        </strong> (Net Δ {formatNumber(summary.balances.netChangeTL, { compact: true })})</li>
        <li>Açılış USD → Kapanış USD: <strong style={{ color: 'var(--text)' }}>
          {formatNumber(summary.balances.beginningUSD, { compact: true })} → {formatNumber(summary.balances.closingUSD, { compact: true })}
        </strong> (Net Δ {formatNumber(summary.balances.netChangeUSD, { compact: true })})</li>
        <li>Toplam Tahsilat TL: <strong style={{ color: 'var(--text)' }}>{formatNumber(summary.totals.collectionsTL, { compact: true })}</strong></li>
        <li>Toplam Ödeme TL: <strong style={{ color: 'var(--text)' }}>{formatNumber(summary.totals.paymentsTL, { compact: true })}</strong></li>
      </ul>
    </div>
  );
}
