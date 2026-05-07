import { useEffect, useState } from 'react';
import { ResponsiveContainer, LineChart, Line, XAxis, YAxis, Tooltip, CartesianGrid, Legend } from 'recharts';
import { formatNumber } from '../utils/format.js';

export default function Projection({ data }) {
  const [days, setDays] = useState(10);
  const [result, setResult] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const run = async () => {
    setLoading(true); setError(null);
    try {
      const res = await fetch('/api/project', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data, days }),
      });
      if (!res.ok) throw new Error((await res.json()).error || 'Projection failed');
      setResult(await res.json());
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { run(); /* eslint-disable-next-line */ }, [data]);

  const fmt = (v) => Intl.NumberFormat('en-US', { notation: 'compact', maximumFractionDigits: 1 }).format(v);

  return (
    <div>
      <div className="controls" style={{ marginBottom: 16 }}>
        <label>Gün sayısı</label>
        <input type="number" min={1} max={60} value={days} onChange={(e) => setDays(Number(e.target.value))} />
        <button className="primary" onClick={run} disabled={loading}>Projeksiyon hesapla</button>
        {result?.model && (
          <span className="muted">
            Avg Δ: {formatNumber(result.model.avgDailyDelta, { compact: true })} ·
            Avg Inflow: {formatNumber(result.model.avgInflow, { compact: true })} ·
            Avg Payments: {formatNumber(result.model.avgPayments, { compact: true })}
          </span>
        )}
      </div>
      {error && <div className="error">{error}</div>}
      {result?.projection?.length > 0 && (
        <div style={{ width: '100%', height: 320 }}>
          <ResponsiveContainer>
            <LineChart data={result.projection} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
              <CartesianGrid stroke="#1e293b" strokeDasharray="3 3" />
              <XAxis dataKey="day" stroke="#94a3b8" />
              <YAxis stroke="#94a3b8" tickFormatter={fmt} />
              <Tooltip contentStyle={{ background: '#0f172a', border: '1px solid #1e293b' }} formatter={(v) => fmt(v)} />
              <Legend />
              <Line type="monotone" dataKey="projectedClosingTL" name="Tahmini Kapanış TL" stroke="#38bdf8" strokeWidth={2} />
              <Line type="monotone" dataKey="projectedInflowTL" name="Tahmini Tahsilat TL" stroke="#22c55e" strokeWidth={2} strokeDasharray="4 4" />
              <Line type="monotone" dataKey="projectedPaymentsTL" name="Tahmini Ödeme TL" stroke="#f87171" strokeWidth={2} strokeDasharray="4 4" />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}
