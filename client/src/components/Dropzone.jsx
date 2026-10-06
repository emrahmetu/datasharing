import { useState, useRef } from 'react';

export default function Dropzone({ onFile, loading }) {
  const [drag, setDrag] = useState(false);
  const inputRef = useRef(null);

  const handleFiles = (files) => {
    if (!files || !files.length) return;
    const file = files[0];
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      alert('Lütfen bir XLSX, XLS veya CSV dosyası seçin.');
      return;
    }
    onFile(file);
  };

  return (
    <div
      className={`dropzone ${drag ? 'drag' : ''}`}
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={(e) => { e.preventDefault(); setDrag(false); handleFiles(e.dataTransfer.files); }}
    >
      <p><strong>{loading ? 'Yükleniyor…' : 'Cashflow XLSX dosyanızı sürükleyin veya tıklayın'}</strong></p>
      <p>Konsolide günlük format desteklenir (Beginning Balance, Collections, Payments, Closing Balance).</p>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        style={{ display: 'none' }}
        onChange={(e) => handleFiles(e.target.files)}
      />
    </div>
  );
}
