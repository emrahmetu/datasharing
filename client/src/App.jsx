import { useState } from 'react';
import Dropzone from './components/Dropzone.jsx';
import SummaryCards from './components/SummaryCards.jsx';
import CashflowChart from './components/CashflowChart.jsx';
import CashflowTable from './components/CashflowTable.jsx';
import Projection from './components/Projection.jsx';
import Report from './components/Report.jsx';

export default function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [tab, setTab] = useState('overview');
  const [currency, setCurrency] = useState('TL');

  const handleFile = async (file) => {
    setLoading(true); setError(null);
    try {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch('/api/upload', { method: 'POST', body: fd });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Upload failed (${res.status})`);
      }
      setData(await res.json());
      setTab('overview');
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app">
      <header>
        <div>
          <h1>Cashflow Console</h1>
          <div className="sub">Konsolide günlük nakit akışı analizi · XLSX/CSV destekli</div>
        </div>
        {data && (
          <button className="ghost" onClick={() => { setData(null); setError(null); }}>
            Yeni dosya
          </button>
        )}
      </header>

      {error && <div className="error">{error}</div>}

      {!data && <Dropzone onFile={handleFile} loading={loading} />}

      {data && (
        <>
          <SummaryCards summary={data.summary} />

          <div className="panel">
            <div className="spaced">
              <div className="tabs">
                <button className={tab === 'overview' ? 'active' : ''} onClick={() => setTab('overview')}>Özet</button>
                <button className={tab === 'table' ? 'active' : ''} onClick={() => setTab('table')}>Tablo</button>
                <button className={tab === 'projection' ? 'active' : ''} onClick={() => setTab('projection')}>Tahmin</button>
                <button className={tab === 'report' ? 'active' : ''} onClick={() => setTab('report')}>Rapor</button>
              </div>
              {(tab === 'overview') && (
                <div className="tabs">
                  <button className={currency === 'TL' ? 'active' : ''} onClick={() => setCurrency('TL')}>TL</button>
                  <button className={currency === 'USD' ? 'active' : ''} onClick={() => setCurrency('USD')}>USD</button>
                </div>
              )}
            </div>

            {tab === 'overview' && <CashflowChart data={data} currency={currency} />}
            {tab === 'table' && <CashflowTable data={data} />}
            {tab === 'projection' && <Projection data={data} />}
            {tab === 'report' && <Report data={data} summary={data.summary} />}
          </div>
        </>
      )}
    </div>
  );
}
