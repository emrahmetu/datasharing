import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as XLSX from 'xlsx';
import { parseCashflow, projectCashflow, summarize } from './cashflow.js';

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.get('/api/health', (_req, res) => res.json({ ok: true }));

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true });
    const sheetName = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null, raw: true });
    const parsed = parseCashflow(rows);
    res.json({ sheet: sheetName, ...parsed, summary: summarize(parsed) });
  } catch (err) {
    res.status(400).json({ error: 'Failed to parse file', detail: String(err.message || err) });
  }
});

app.post('/api/project', (req, res) => {
  const { data, days = 10 } = req.body || {};
  if (!data) return res.status(400).json({ error: 'Missing data' });
  try {
    const projection = projectCashflow(data, days);
    res.json(projection);
  } catch (err) {
    res.status(400).json({ error: 'Projection failed', detail: String(err.message || err) });
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`Cashflow server listening on http://localhost:${PORT}`);
});
