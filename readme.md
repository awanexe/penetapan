# Dashboard SK — Hybrid Web App (GAS + Vercel)

Migrasi dari Google Apps Script monolitik (`HtmlService`) menjadi arsitektur **Hybrid Web App**:

- **Backend**: Google Apps Script (JSON API via `doGet`/`doPost`)
- **Database**: Google Sheets (spreadsheet aktif — script terikat)
- **Frontend**: statis di Vercel
- **Proxy**: Vercel Serverless Function (`api/proxy.mjs`) untuk mengatasi CORS & menyembunyikan URL GAS

Seluruh fitur lama dipertahankan: dashboard KPI, chart, grid OPD, halaman data tahunan, tombol Rekap Data, dan reveal animation.

---

## 📁 Struktur Proyek
