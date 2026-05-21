import { useMemo } from 'react';
import {
  ResponsiveContainer, ComposedChart, Bar, Line, XAxis, YAxis, Tooltip, Legend, CartesianGrid,
} from 'recharts';

const findRow = (lineItems, re) => lineItems.find(li => re.test(li.label));

export default function CashflowChart({ data, currency = 'TL' }) {
  const series = useMemo(() => {
    if (!data?.dates) return [];
    const closing = findRow(data.lineItems, currency === 'TL' ? /^closing balance tl/i : /^closing balances? usd/i);
    const inflow = findRow(data.lineItems, currency === 'TL' ? /^total inflow tl/i : /^total inflow usd/i);
    const collections = findRow(data.lineItems, currency === 'TL' ? /^total tl$/i : /^total usd$/i);
    const payments = findRow(data.lineItems, currency === 'TL' ? /^grand total tl/i : /^grand total usd/i);

    return data.dates.map((d, i) => ({
      date: d,
      closing: closing?.values[i] ?? null,
      inflow: collections?.values[i] ?? null,
      payments: -(payments?.values[i] ?? 0),
    }));
  }, [data, currency]);

  if (!series.length) return <p className="muted">Veri yok.</p>;

  const fmt = (v) => Intl.NumberFormat('en-US', { notation: 'compact', maximumFractionDigits: 1 }).format(v);

  return (
    <div style={{ width: '100%', height: 360 }}>
      <ResponsiveContainer>
        <ComposedChart data={series} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
          <CartesianGrid stroke="#1e293b" strokeDasharray="3 3" />
          <XAxis dataKey="date" stroke="#94a3b8" tick={{ fontSize: 11 }} />
          <YAxis stroke="#94a3b8" tickFormatter={fmt} tick={{ fontSize: 11 }} />
          <Tooltip
            contentStyle={{ background: '#0f172a', border: '1px solid #1e293b' }}
            formatter={(v) => fmt(v)}
          />
          <Legend />
          <Bar dataKey="inflow" name={`Tahsilat ${currency}`} fill="#22c55e" />
          <Bar dataKey="payments" name={`Ödeme ${currency}`} fill="#f87171" />
          <Line type="monotone" dataKey="closing" name={`Kapanış ${currency}`} stroke="#38bdf8" strokeWidth={2} dot={false} />
        </ComposedChart>
      </ResponsiveContainer>
    </div>
  );
}
