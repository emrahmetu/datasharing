export function formatNumber(value, opts = {}) {
  if (value === null || value === undefined || Number.isNaN(value)) return '—';
  const { compact = false, currency } = opts;
  const formatter = new Intl.NumberFormat('en-US', {
    notation: compact ? 'compact' : 'standard',
    maximumFractionDigits: compact ? 2 : 0,
  });
  const formatted = formatter.format(value);
  return currency ? `${formatted} ${currency}` : formatted;
}

export function formatDelta(value, currency) {
  if (value === null || value === undefined) return '—';
  const sign = value > 0 ? '+' : '';
  return `${sign}${formatNumber(value, { compact: true, currency })}`;
}
