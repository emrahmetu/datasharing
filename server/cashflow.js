// Cashflow XLSX parser tailored for the consolidated daily format:
// Header row begins with "CONSOLIDATED" followed by date columns (e.g. "2 Feb Monday").
// Sections are introduced by section labels: "Beginning Balance TL/USD", "Collections",
// "TOTAL INFLOW TL/USD", "PAYMENTS", "Grand Total TL/USD", "Closing Balance TL/USD".

const NUMERIC_RE = /^-?\d+([.,]\d+)*$/;

function toNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (v instanceof Date) return null;
  const s = String(v).trim();
  if (!s || s === '__' || s === '-' || s.toLowerCase() === 'tatil') return null;
  const cleaned = s.replace(/,/g, '').replace(/\s+/g, '');
  if (!NUMERIC_RE.test(cleaned.replace(/,/g, ''))) {
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : null;
  }
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function isHeaderRow(row) {
  if (!row || !row.length) return false;
  const first = String(row[0] ?? '').trim().toUpperCase();
  return first === 'CONSOLIDATED' || first === 'PAYMENTS';
}

function buildDateColumns(headerRow) {
  // Skip the first cell (label) and produce date strings for the remaining columns.
  return headerRow.slice(1).map((cell, idx) => {
    if (cell instanceof Date) {
      return cell.toISOString().slice(0, 10);
    }
    const s = (cell ?? '').toString().trim();
    return s || `col_${idx + 1}`;
  });
}

export function parseCashflow(rows) {
  // Strip empty rows.
  const cleaned = rows.filter(r => r && r.some(c => c !== null && c !== undefined && String(c).trim() !== ''));

  let dates = [];
  let section = 'header';
  const lineItems = []; // { label, section, values: [] }

  for (const row of cleaned) {
    const label = String(row[0] ?? '').trim();
    const upper = label.toUpperCase();

    if (isHeaderRow(row)) {
      const cols = buildDateColumns(row);
      if (cols.length > dates.length) dates = cols;
      section = upper === 'PAYMENTS' ? 'payments' : 'inflow';
      continue;
    }

    if (!label) continue;

    if (/^collections$/i.test(label)) { section = 'collections'; continue; }
    if (/^payments$/i.test(label)) { section = 'payments'; continue; }

    const values = row.slice(1, dates.length + 1).map(toNumber);
    const item = { label, section, values };

    if (/^beginning balance/i.test(label)) item.section = 'beginning';
    else if (/^closing balance/i.test(label)) item.section = 'closing';
    else if (/^total inflow/i.test(label)) item.section = 'totalInflow';
    else if (/^total\s+(tl|usd)/i.test(label)) item.section = 'collectionsTotal';
    else if (/^grand total/i.test(label)) item.section = 'grandTotal';

    lineItems.push(item);
  }

  return { dates, lineItems };
}

export function summarize({ dates, lineItems }) {
  const findRow = (re) => lineItems.find(li => re.test(li.label));
  const beginningTL = findRow(/^beginning balance tl/i);
  const beginningUSD = findRow(/^beginning balance usd/i);
  const closingTL = findRow(/^closing balance tl/i);
  const closingUSD = findRow(/^closing balances? usd/i);
  const inflowTL = findRow(/^total inflow tl/i);
  const inflowUSD = findRow(/^total inflow usd/i);
  const grandTL = findRow(/^grand total tl/i);
  const grandUSD = findRow(/^grand total usd/i);

  const sumOf = (row) => row ? row.values.reduce((a, b) => a + (b || 0), 0) : 0;
  const last = (row) => row ? [...row.values].reverse().find(v => v !== null && v !== undefined) ?? null : null;
  const first = (row) => row ? row.values.find(v => v !== null && v !== undefined) ?? null : null;

  return {
    period: { start: dates[0] || null, end: dates[dates.length - 1] || null, days: dates.length },
    totals: {
      inflowTL: sumOf(inflowTL && { values: inflowTL.values.map((v, i) => (v ?? 0) - (i > 0 ? inflowTL.values[i - 1] ?? 0 : 0)) }) || null,
      collectionsTL: sumCollections(lineItems, 'TL'),
      collectionsUSD: sumCollections(lineItems, 'USD'),
      paymentsTL: sumOf(grandTL),
      paymentsUSD: sumOf(grandUSD),
    },
    balances: {
      beginningTL: first(beginningTL),
      beginningUSD: first(beginningUSD),
      closingTL: last(closingTL),
      closingUSD: last(closingUSD),
      netChangeTL: (last(closingTL) ?? 0) - (first(beginningTL) ?? 0),
      netChangeUSD: (last(closingUSD) ?? 0) - (first(beginningUSD) ?? 0),
    },
  };
}

function sumCollections(lineItems, currency) {
  // Use the explicit "Total TL"/"Total USD" rows that appear inside the Collections block.
  const target = currency === 'TL' ? /^total tl$/i : /^total usd$/i;
  const row = lineItems.find(li => target.test(li.label) && li.section === 'collectionsTotal');
  if (!row) return null;
  return row.values.reduce((a, b) => a + (b || 0), 0);
}

export function projectCashflow(parsed, days = 10) {
  const { dates, lineItems } = parsed;
  const closingTL = lineItems.find(li => /^closing balance tl/i.test(li.label));
  const closingUSD = lineItems.find(li => /^closing balances? usd/i.test(li.label));
  const inflowTL = lineItems.find(li => /^total inflow tl/i.test(li.label));
  const grandTL = lineItems.find(li => /^grand total tl/i.test(li.label));

  // Average daily net change from history.
  const series = (closingTL?.values || []).filter(v => v !== null);
  const deltas = [];
  for (let i = 1; i < series.length; i++) deltas.push(series[i] - series[i - 1]);
  const avgDelta = deltas.length ? deltas.reduce((a, b) => a + b, 0) / deltas.length : 0;

  const avgInflow = avg(inflowTL?.values);
  const avgPayments = avg(grandTL?.values);

  const projection = [];
  let lastClose = series.length ? series[series.length - 1] : null;
  for (let i = 1; i <= days; i++) {
    const projected = lastClose !== null ? lastClose + avgDelta : null;
    projection.push({
      day: i,
      projectedClosingTL: projected,
      projectedInflowTL: avgInflow,
      projectedPaymentsTL: avgPayments,
    });
    if (projected !== null) lastClose = projected;
  }

  return { projection, model: { avgDailyDelta: avgDelta, avgInflow, avgPayments } };
}

function avg(arr) {
  if (!arr) return 0;
  const v = arr.filter(x => x !== null && x !== undefined);
  if (!v.length) return 0;
  return v.reduce((a, b) => a + b, 0) / v.length;
}
