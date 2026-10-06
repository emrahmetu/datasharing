# Cashflow Program

Konsolide günlük nakit akışı (cashflow) için web tabanlı bir analiz aracı. XLSX şablonunuzu yükleyin; bakiye, tahsilat ve ödeme satırlarını otomatik ayrıştırır, görselleştirir, projeksiyon yapar ve rapor üretir.

## Desteklenen Format

Şablon ilk satırda `CONSOLIDATED` veya `PAYMENTS` etiketinden sonra tarih sütunlarını içermelidir (örn. `2 Feb Monday`). Aşağıdaki bölümler otomatik olarak tanınır:

- `Beginning Balance TL` / `Beginning Balance USD`
- `Collections` (alt satırlar: US/UK/TR Export Receivables, Domestic Collection, Electricity, Interest Income, VAT Refund, Loan Utilisation vs.)
- `Total TL` / `Total USD` (Collections toplamı)
- `TOTAL INFLOW TL` / `TOTAL INFLOW USD`
- `PAYMENTS` (alt satırlar: Loan Repayment, Payroll, Tax, Trade Payables, Export Freight, Fuel, Soda World, vb.)
- `Grand Total TL` / `Grand Total USD`
- `Closing Balance TL` / `Closing Balances USD`

Boş hücreler, `__`, `-` ve `Tatil` değerleri otomatik olarak yok sayılır.

## Mimari

```
/server   Node.js + Express + xlsx (parser, projection API)
/client   React + Vite + Recharts (upload, dashboard, charts, raporlar)
```

## Kurulum & Çalıştırma

```bash
npm run install:all     # workspace bağımlılıklarını kur
npm run dev             # server (4000) + client (5173) birlikte
```

Sonra `http://localhost:5173` adresine gidip XLSX dosyanızı yükleyin.

Yalnızca server: `npm run dev:server`
Yalnızca client: `npm run dev:client`

## Özellikler

- **Yükleme**: XLSX/XLS/CSV drag-and-drop
- **Özet**: Açılış/kapanış bakiyeleri, net Δ, tahsilat ve ödeme toplamları
- **Tablo**: Tüm satırların gün gün dökümü, satır arama
- **Grafikler**: Tahsilat (yeşil) + Ödeme (kırmızı) bar, Kapanış bakiyesi line chart, TL/USD geçişi
- **Tahmin**: Geçmiş ortalamadan günlük net Δ projeksiyonu (1–60 gün)
- **Rapor**: CSV ve JSON dışa aktarma

## API

| Endpoint | Method | Açıklama |
|---|---|---|
| `/api/health` | GET | Sağlık kontrolü |
| `/api/upload` | POST (multipart `file`) | XLSX/CSV ayrıştır, özet ile birlikte döndür |
| `/api/project` | POST `{data, days}` | Tahmin hesapla |
