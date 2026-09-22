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

### Sesi 6: Penambahan Data Informasi Ringkasan Emiten (69 Saham)
- **Kebutuhan**: User ingin menambahkan informasi ringkasan profil bisnis untuk setiap emiten agar tampil di dashboard.
- **Solusi Spreadsheet**:
  - Dibuat file script [`populate_ringkasan.gs`](populate_ringkasan.gs) yang memuat fungsi `populateRingkasanData()` untuk menginisialisasi sheet **`Ringkasan`** di spreadsheet tanpa menimpa data sheet `Data`.
  - Berisi data lengkap untuk **ke-69 emiten** (AADI s.d. WIFI): `Kode`, `Nama Perusahaan`, `Sektor`, `Subsektor`, `Ringkasan Bisnis`, dan `Highlight / Keunggulan`.
  - Menambahkan menu `📝 Isi / Update Sheet Ringkasan (69 Emiten)` di toolbar `kode.gs`.
- **Integrasi Backend**:
  - Menambahkan `getEmitenProfile_` di [`pintarsaham_dashboard.gs`](pintarsaham_dashboard.gs) untuk menyertakan objek `profile` pada respons JSON.
- **Integrasi Frontend**:
  - Menambahkan styling CSS dan komponen UI di [`dashboard.html`](dashboard.html):
    - Nama lengkap perusahaan dan badge sektor pada Hero Section.
    - Kartu baru: **"🏢 Profil Perusahaan & Ringkasan Bisnis"** dengan highlight keunggulan kompetitif.
    - Menanamkan kamus data `EMITEN_PROFILES` di frontend sebagai fallback instan (zero-downtime) sehingga langsung tampil di dashboard tanpa menunggu perubahan backend.

### Sesi 7: Resolusi Tuntas Error 'Unique Security Origin' & Validasi Rendering Frontend
- **Analisis Mendalam Error**:
  - `Unsafe attempt to load URL file:///... from frame with URL file:///... 'file:' URLs are treated as unique security origins.`
  - Ditemukan 2 faktor penyebab:
    1. **Sintaksis Terbuka**: Pada commit sebelumnya, penambahan kamus `EMITEN_PROFILES` secara tidak sengaja memotong penutup kurung kurawal `}` pada fungsi `loadEmiten()`. Ini memicu `SyntaxError: Unexpected end of input` yang menghentikan eksekusi script sebelum komponen DOM dirender.
    2. **Akses Storage pada `file:///`**: Browser berbasis Chromium memperlakukan URL dengan protokol `file:///` sebagai origin unik (`origin: null`). Pada mode privasi tertentu, memanggil `localStorage.getItem` dapat melempar `SecurityError: Access to Storage is not allowed`.
- **Langkah Perbaikan**:
  - Menambahkan kurung penutup `}` yang hilang pada fungsi `loadEmiten()` di [`dashboard.html`](dashboard.html).
  - Membungkus pembacaan dan penulisan `localStorage` dalam blok `try...catch` yang aman agar tidak pernah menghentikan alur program jika dibuka langsung sebagai file lokal.
  - Membuat duplikat [`index.html`](index.html) sehingga dapat dijalankan langsung di server root HTTP (`localhost:8080`) maupun berbagai static web hosting.
  - Memvalidasi seluruh sintaks JavaScript menggunakan Node.js VM engine (`new Function`).
  - Menguji rendering menggunakan browser engine Chromium (Playwright):
    - Berhasil mengambil daftar 69 emiten secara live dari Web App Apps Script.
    - Berhasil merender Hero Section, 4 kartu metrik KPI, kartu Profil & Ringkasan Perusahaan, dan Chart historis tanpa error.

---

## 📋 Catatan Teknis untuk Menjalankan Dashboard
1. **Cara 1: Local HTTP Server (Sangat Disarankan)**
   - Cukup double-click file [`jalankan_dashboard.bat`](jalankan_dashboard.bat).
   - Browser akan otomatis membuka `http://localhost:8080/dashboard.html` (atau `http://localhost:8080/`).
   - Bebas dari segala batasan keamanan `file:///` browser.
2. **Cara 2: Buka Langsung File HTML**
   - Double-click [`dashboard.html`](dashboard.html) atau [`index.html`](index.html).
   - Kode sudah dilengkapi pelindung `try...catch` dan penghapusan `<base target="_top">` sehingga API request tetap dapat berjalan normal.
3. **Konfigurasi URL API**:
   - Jika URL deployment Apps Script diperbarui, klik tombol **⚙️ API** di pojok kanan atas dashboard untuk memasukkan URL baru.

