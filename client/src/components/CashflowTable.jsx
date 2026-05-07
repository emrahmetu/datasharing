import { useMemo, useState } from 'react';
import { formatNumber } from '../utils/format.js';

const SECTION_ORDER = ['beginning', 'collections', 'collectionsTotal', 'totalInflow', 'payments', 'grandTotal', 'closing'];
const SECTION_LABELS = {
  beginning: 'Beginning Balance',
  collections: 'Collections',
  collectionsTotal: 'Collection Totals',
  totalInflow: 'Total Inflow',
  payments: 'Payments',
  grandTotal: 'Grand Total',
  closing: 'Closing Balance',
};

export default function CashflowTable({ data }) {
  const [filter, setFilter] = useState('');

  const grouped = useMemo(() => {
    const out = Object.fromEntries(SECTION_ORDER.map(s => [s, []]));
    for (const item of data.lineItems) {
      const key = SECTION_ORDER.includes(item.section) ? item.section : 'collections';
      if (filter && !item.label.toLowerCase().includes(filter.toLowerCase())) continue;
      out[key].push(item);
    }
    return out;
  }, [data, filter]);

  return (
    <div>
      <div className="spaced">
        <input
          placeholder="Satır ara…"
          value={filter}
          onChange={(e) => setFilter(e.target.value)}
          style={{ background: 'var(--panel-2)', color: 'var(--text)', border: '1px solid var(--border)', borderRadius: 6, padding: '6px 10px', minWidth: 240 }}
        />
        <span className="muted">{data.dates.length} sütun · {data.lineItems.length} satır</span>
      </div>
      <div className="scroll">
        <table>
          <thead>
            <tr>
              <th className="label-col">Item</th>
              {data.dates.map((d) => <th key={d}>{d}</th>)}
            </tr>
          </thead>
          <tbody>
            {SECTION_ORDER.map(section => {
              const rows = grouped[section];
              if (!rows.length) return null;
              return (
                <>
                  <tr key={`s-${section}`} className="section">
                    <td className="label-col" colSpan={data.dates.length + 1}>{SECTION_LABELS[section]}</td>
                  </tr>
                  {rows.map((row, i) => (
                    <tr key={`${section}-${i}`}>
                      <td className="label-col">{row.label}</td>
                      {data.dates.map((_, idx) => (
                        <td key={idx}>{formatNumber(row.values[idx])}</td>
                      ))}
                    </tr>
                  ))}
                </>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
