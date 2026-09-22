# PintarSaham Dashboard - Session Log

## 📌 Ringkasan Proyek
- **Nama Proyek**: PintarSaham Dashboard (`stock_dashboard`)
- **Repository Git**: [https://github.com/BismaYoga/stock_dashboard.git](https://github.com/BismaYoga/stock_dashboard.git)
- **Target Spreadsheet**: [Google Spreadsheet Data Dashboard](https://docs.google.com/spreadsheets/d/1tbkh-ulOyzm-SQm2H_cLEpZZDzQY7A0uB1v3fv-PAjE/edit?usp=sharing)
  - **Spreadsheet ID**: `1tbkh-ulOyzm-SQm2H_cLEpZZDzQY7A0uB1v3fv-PAjE`
- **Deployed Web App API URL**:
  `https://script.google.com/macros/s/AKfycby4yTQd8PJn4s3e4Fd6DfFVsHjyvoJBVOKr1hf6IxlWdU6x4lLDyWZfRTOwFnEPejAklQ/exec`

---

## 🏗️ Arsitektur Sistem (Decoupled)
Sistem ini menggunakan arsitektur **terpisah (decoupled)**:

1. **Backend (Google Apps Script - Spreadsheet)**:
   - `kode.gs`: Modul ETL resumable untuk membaca file Excel/Sheets valuasi di Google Drive (`Analisa Saham/Valuasi`), mengonversinya, dan menyusun sheet `Data` (157 kolom, 21 kuartal, formula harga live GOOGLEFINANCE).
   - `pintarsaham_dashboard.gs`: Bertindak sebagai **JSON API murni** (`ContentService`), melayani request:
     - `GET ?action=list` ➔ Mengembalikan daftar kode emiten `{ codes, total, withData }`.
     - `GET ?action=get&code=XXXX` ➔ Mengembalikan data fundamental, valuasi, harga wajar (Bear/Base/Bull), kinerja tahunan, YoY, dan balance sheet emiten.
   - *Catatan Penting*: Tidak ada file HTML di Google Apps Script.

2. **Frontend (Klien Web / Dashboard)**:
   - `dashboard.html`: Single-page application interaktif dengan Chart.js, picker kode saham, kartu KPI, dan kartu valuasi.
   - Mengambil data melalui HTTP `fetch` ke Web App API URL di atas.
   - Bisa dijalankan di browser lokal via `jalankan_dashboard.bat`, di-hosting di Hugging Face Spaces, Vercel, maupun GitHub Pages.

---

## ⏱️ Kronologi Percakapan & Keputusan Teknis

### Sesi 1: Review Awal Kode Sumber
- Memeriksa file: `kode.gs`, `pintarsaham_dashboard.gs`, dan `dashboard.html`.
- Mengidentifikasi kelebihan arsitektur auto-resume batching pada `kode.gs`.
- Menemukan potensi bug:
  - Case-sensitivity template HTML pada Apps Script (`'Dashboard'` vs `'dashboard'`).
  - Tidak adanya menu pembuka Web App di `onOpen`.
  - Kalkulasi valuasi PER yang membalik logika jika EPS bernilai negatif (perusahaan merugi).
  - Karakteristik emiten perbankan yang tidak memiliki pos Laba Kotor.
  - Bottleneck pembacaan 157 kolom sekaligus pada `listCodes()`.

### Sesi 2: Inisialisasi Git & Push Baseline
- Inisialisasi git pada subdirektori `pintarsaham_dashboard`.
- Menambahkan `README.md`.
- Melakukan *first commit* dan push ke branch `main` pada remote [https://github.com/BismaYoga/stock_dashboard.git](https://github.com/BismaYoga/stock_dashboard.git).

### Sesi 3: Integrasi Spreadsheet ID & Decoupling API
- User memberikan link Google Spreadsheet `1tbkh-ulOyzm-SQm2H_cLEpZZDzQY7A0uB1v3fv-PAjE`.
- Menanamkan `SPREADSHEET_ID` ke `CONFIG` (`kode.gs`) dan `DASH_CFG` (`pintarsaham_dashboard.gs`).
- User menginstruksikan bahwa **HTML tidak akan dideploy di Google Apps Script**, melainkan Apps Script hanya bertugas mengambil/menyajikan data.
- Mengubah `doGet(e)` di `pintarsaham_dashboard.gs` menjadi JSON API endpoint murni (`ContentService.MimeType.JSON`).
- Mengubah mekanisme komunikasi `dashboard.html` dari `google.script.run` menjadi HTTP `fetch(API_URL)`.

### Sesi 4: Verifikasi Deployment Backend & Perbaikan Error Origin
- User mendeploy script ke Web App dan membagikan URL:
  `https://script.google.com/macros/s/AKfycby4yTQd8PJn4s3e4Fd6DfFVsHjyvoJBVOKr1hf6IxlWdU6x4lLDyWZfRTOwFnEPejAklQ/exec`
- Dilakukan verifikasi API backend: **Berhasil 100%**, mengembalikan 69 emiten lengkap dengan metrik keuangan.
- Menyelesaikan pesan error:
  `Unsafe attempt to load URL file:///... 'file:' URLs are treated as unique security origins.`
  - **Penyebab**: Tag `<base target="_top">` bawaan Apps Script masih ada di `dashboard.html`.
  - **Solusi**: Menghapus tag `<base target="_top">`, menyetel default `API_URL` ke URL deploy user, dan membuat launcher `jalankan_dashboard.bat` untuk menjalankan local server `http://localhost:8080/dashboard.html`.

### Sesi 5: Kesepakatan Fokus Pengembangan
- **Status Backend**: `kode.gs` dan `pintarsaham_dashboard.gs` sudah selesai dideploy dan dalam kondisi stabil.
- **Fokus Selanjutnya**: Seluruh perubahan dan peningkatan berikutnya difokuskan pada antarmuka frontend di [`dashboard.html`](dashboard.html).

---

## 📋 Catatan Teknis untuk Pengembangan Frontend (`dashboard.html`)
- Endpoint API sudah terpasang secara default di variabel `DEFAULT_API_URL`.
- Bila user mengganti URL deployment, dapat diubah melalui tombol `⚙️ API` di kanan atas header (tersimpan di `localStorage`).
- Local server dapat diakses dengan menjalankan `jalankan_dashboard.bat`.
