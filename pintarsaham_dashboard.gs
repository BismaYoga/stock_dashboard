/**
 * ============================================================================
 * PintarSaham - Dashboard Web App
 * ============================================================================
 * Web app yang menampilkan dashboard interaktif emiten dengan dropdown kode,
 * KPI cards, valuasi (EPS/BVPS + Bear/Base/Bull), dan chart tahunan + YoY.
 *
 * DEPLOY:
 *   1) Apps Script Editor > Deploy > New deployment
 *   2) Type: Web app
 *   3) Execute as: Me  |  Who has access: Anyone with the link (atau lebih privat)
 *   4) Deploy > copy URL
 *
 * Data source: sheet "Data" yang dibangun oleh PintarSaham_DataDashboard.gs
 * (read-only, aman dijalankan paralel dengan Build/auto-resume).
 *
 * Catatan: angka di sheet Data dalam ribuan IDR. Saat untuk display & chart,
 * dikalikan 1000 untuk konversi ke IDR penuh.
 * ============================================================================
 */

const DASH_CFG = {
  DATA_SHEET: 'Data',
  // Currency conversion: data tersimpan dalam ribuan, dikalikan ini untuk IDR penuh
  SCALE_FACTOR: 1000,
  // ID Google Spreadsheet Data Dashboard
  SPREADSHEET_ID: '1tbkh-ulOyzm-SQm2H_cLEpZZDzQY7A0uB1v3fv-PAjE',
};

// 21 kuartal sama seperti di DataDashboard, urutan Q1 2026 -> Q1 2021
const DASH_QUARTERS = [
  'Q1 2026', 'Q4 2025', 'Q3 2025', 'Q2 2025', 'Q1 2025',
  'Q4 2024', 'Q3 2024', 'Q2 2024', 'Q1 2024',
  'Q4 2023', 'Q3 2023', 'Q2 2023', 'Q1 2023',
  'Q4 2022', 'Q3 2022', 'Q2 2022', 'Q1 2022',
  'Q4 2021', 'Q3 2021', 'Q2 2021', 'Q1 2021'
];

// Column index map (1-based) sesuai dengan DataDashboard
const DASH_COL = {
  KODE: 1,
  HARGA: 2,
  PENDAPATAN_START: 3,
  LABA_KOTOR_START: 24,
  LABA_USAHA_START: 45,
  LABA_BERSIH_START: 66,
  TOTAL_ASET_START: 87,
  TOTAL_LIAB_START: 108,
  TOTAL_EKUI_START: 129,
  EPS: 150, MIN_PE: 151, MEAN_PE: 152, MAX_PE: 153,
  BVPS: 154, MIN_PBV: 155, MEAN_PBV: 156, MAX_PBV: 157,
};

// ============================ WEB APP ENTRY =================================

function doGet() {
  let template;
  try {
    template = HtmlService.createTemplateFromFile('dashboard');
  } catch (e) {
    template = HtmlService.createTemplateFromFile('Dashboard');
  }
  return template
    .evaluate()
    .setTitle('PintarSaham Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Helper untuk mendapatkan instance Spreadsheet (by ID atau Active) */
function getSpreadsheet_() {
  if (DASH_CFG.SPREADSHEET_ID) {
    try {
      return SpreadsheetApp.openById(DASH_CFG.SPREADSHEET_ID);
    } catch (e) {
      Logger.log('Gagal openById(' + DASH_CFG.SPREADSHEET_ID + '): ' + e.message);
    }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  throw new Error('Spreadsheet tidak ditemukan. Pastikan DASH_CFG.SPREADSHEET_ID sudah benar atau script di-bind ke spreadsheet.');
}

/** Inject file lain (CSS/JS) ke template HTML utama. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================ MENU INTEGRATION ==============================

/**
 * Helper untuk menambahkan menu Dashboard di onOpen yang sudah ada.
 * Karena onOpen utama di PintarSaham_DataDashboard.gs, fungsi ini bisa
 * dipanggil dari sana, atau ditambahkan sebagai menu standalone.
 */
function openDashboardInfo() {
  const ui = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();
  const msg = url
    ? 'Dashboard web app:\n\n' + url + '\n\n' +
      'Tip: bookmark URL ini supaya akses cepat.'
    : 'Web app belum di-deploy.\n\nLangkah deploy:\n' +
      '  1. Apps Script Editor > Deploy > New deployment\n' +
      '  2. Type: Web app\n' +
      '  3. Execute as: Me  |  Who has access: pilih sesuai kebutuhan\n' +
      '  4. Deploy > copy URL';
  ui.alert('🌐 PintarSaham Dashboard', msg, ui.ButtonSet.OK);
}

// ============================ DATA API ======================================

/**
 * List semua kode emiten yang tersedia di sheet Data.
 * Filter: hanya kode yang punya data Pendapatan atau Laba Bersih (≥1 kuartal).
 * Return: { codes: [...], total: N, withData: M }
 */
function listCodes() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(DASH_CFG.DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return { codes: [], total: 0, withData: 0 };
  }
  const lastRow = sheet.getLastRow();
  // Optimasi: Hanya baca kolom yang diperlukan untuk cek ketersediaan data (Kode s.d. akhir kuartal Laba Bersih)
  const maxCol = DASH_COL.LABA_BERSIH_START - 1 + DASH_QUARTERS.length;
  const values = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();

  const codes = [];
  let withData = 0;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const code = String(row[0] || '').trim();
    if (!code) continue;
    // Cek apakah ada data Pendapatan atau Laba Bersih (≥1 kuartal non-empty)
    const hasPendapatan = hasAnyValue_(row, DASH_COL.PENDAPATAN_START - 1,
                                        DASH_COL.PENDAPATAN_START - 1 + DASH_QUARTERS.length);
    const hasLabaBersih = hasAnyValue_(row, DASH_COL.LABA_BERSIH_START - 1,
                                         DASH_COL.LABA_BERSIH_START - 1 + DASH_QUARTERS.length);
    const hasData = hasPendapatan || hasLabaBersih;
    if (hasData) withData++;
    codes.push({ code: code, hasData: hasData });
  }
  // Sort alfabetis
  codes.sort(function (a, b) { return a.code.localeCompare(b.code); });
  return { codes: codes, total: codes.length, withData: withData };
}

function hasAnyValue_(row, startIdx0, endIdx0) {
  for (let i = startIdx0; i < endIdx0 && i < row.length; i++) {
    const v = row[i];
    if (v !== '' && v !== null && v !== undefined && !isNaN(v)) return true;
  }
  return false;
}

/**
 * Ambil seluruh data untuk satu emiten, plus kalkulasi turunan
 * (harga wajar bear/base/bull, tahunan, YoY, margin).
 */
function getEmitenData(code) {
  if (!code) return { error: 'Kode emiten kosong' };
  code = String(code).trim().toUpperCase();

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(DASH_CFG.DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return { error: 'Sheet Data belum dibangun' };
  }
  const lastRow = sheet.getLastRow();
  // Cari row berdasarkan kode
  const codeColValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowIdx = -1;
  for (let i = 0; i < codeColValues.length; i++) {
    if (String(codeColValues[i][0]).trim().toUpperCase() === code) {
      rowIdx = i + 2; break;
    }
  }
  if (rowIdx < 0) return { error: 'Kode "' + code + '" tidak ditemukan' };

  // Ambil seluruh row + harga sekarang
  const row = sheet.getRange(rowIdx, 1, 1, 157).getValues()[0];
  const harga = parseFiniteNumber_(row[DASH_COL.HARGA - 1]);

  // Bangun kuartal data per metrik
  const metrics = buildQuarterlyMaps_(row);

  // Data ketersediaan: punya minimal 1 kuartal Pendapatan ATAU Laba Bersih?
  const totalCellPend = countEntries_(metrics.pendapatan);
  const totalCellLaba = countEntries_(metrics.labaBersih);
  const hasFinancialData = (totalCellPend > 0 || totalCellLaba > 0);

  // Valuasi
  const eps = parseFiniteNumber_(row[DASH_COL.EPS - 1]);
  const minPe = parseFiniteNumber_(row[DASH_COL.MIN_PE - 1]);
  const meanPe = parseFiniteNumber_(row[DASH_COL.MEAN_PE - 1]);
  const maxPe = parseFiniteNumber_(row[DASH_COL.MAX_PE - 1]);
  const bvps = parseFiniteNumber_(row[DASH_COL.BVPS - 1]);
  const minPbv = parseFiniteNumber_(row[DASH_COL.MIN_PBV - 1]);
  const meanPbv = parseFiniteNumber_(row[DASH_COL.MEAN_PBV - 1]);
  const maxPbv = parseFiniteNumber_(row[DASH_COL.MAX_PBV - 1]);

  // Harga wajar = EPS × multiplier ATAU BVPS × multiplier
  // Catatan: P/E hanya valid jika EPS > 0. Bila EPS <= 0 (perusahaan merugi),
  // valuasi PER di-set null (N/A) untuk mencegah target harga negatif.
  const epsIsPositive = (eps != null && eps > 0);
  const fairValue = {
    per: {
      bear: (epsIsPositive && minPe != null) ? eps * minPe : null,
      base: (epsIsPositive && meanPe != null) ? eps * meanPe : null,
      bull: (epsIsPositive && maxPe != null) ? eps * maxPe : null,
    },
    pbv: {
      bear: (bvps != null && minPbv != null) ? bvps * minPbv : null,
      base: (bvps != null && meanPbv != null) ? bvps * meanPbv : null,
      bull: (bvps != null && maxPbv != null) ? bvps * maxPbv : null,
    }
  };

  // Latest quarter & latest values (kuartal pertama yang ada datanya)
  const latestQuarter = findLatestQuarter_(metrics);
  const latestValues = pickLatestValues_(metrics, latestQuarter);

  // Tahunan (hanya tahun lengkap 4 kuartal)
  const tahunan = buildAnnualSeries_(metrics);
  // YoY current vs same quarter year ago
  const yoy = buildYoYComparison_(metrics, latestQuarter);
  // Balance sheet tahunan (Q4 tiap tahun + kuartal berjalan bila lebih baru)
  const balanceSeries = buildBalanceSeries_(metrics);

  // Total Aset, Liabilitas, Ekuitas (latest)
  const totalAset = latestValueOf_(metrics.totalAset);
  const totalLiab = latestValueOf_(metrics.totalLiab);
  const totalEkui = latestValueOf_(metrics.totalEkui);

  return {
    code: code,
    harga: harga,
    hasFinancialData: hasFinancialData,
    latestQuarter: latestQuarter,
    latest: latestValues,
    balance: { aset: totalAset, liabilitas: totalLiab, ekuitas: totalEkui },
    valuation: {
      eps: eps, bvps: bvps,
      pe: { min: minPe, mean: meanPe, max: maxPe },
      pbv: { min: minPbv, mean: meanPbv, max: maxPbv }
    },
    fairValue: fairValue,
    annual: tahunan,
    yoy: yoy,
    balanceSeries: balanceSeries  // {quarters, aset, liabilitas, ekuitas, ekuitasGrowth}
  };
}

// ============================ HELPERS =======================================

function parseFiniteNumber_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isFinite(n) ? n : null;
}

/** Bangun map kuartal -> value per metrik dari satu row sheet Data. */
function buildQuarterlyMaps_(row) {
  const metrics = {};
  const blocks = [
    { key: 'pendapatan', start: DASH_COL.PENDAPATAN_START },
    { key: 'labaKotor',  start: DASH_COL.LABA_KOTOR_START },
    { key: 'labaUsaha',  start: DASH_COL.LABA_USAHA_START },
    { key: 'labaBersih', start: DASH_COL.LABA_BERSIH_START },
    { key: 'totalAset',  start: DASH_COL.TOTAL_ASET_START },
    { key: 'totalLiab',  start: DASH_COL.TOTAL_LIAB_START },
    { key: 'totalEkui',  start: DASH_COL.TOTAL_EKUI_START },
  ];
  blocks.forEach(function (b) {
    const m = {};
    for (let i = 0; i < DASH_QUARTERS.length; i++) {
      const v = parseFiniteNumber_(row[b.start - 1 + i]);
      if (v !== null) m[DASH_QUARTERS[i]] = v;
    }
    metrics[b.key] = m;
  });
  return metrics;
}

function countEntries_(map) { return Object.keys(map || {}).length; }

/** Cari kuartal terbaru yang punya pendapatan ATAU laba bersih. */
function findLatestQuarter_(metrics) {
  for (let i = 0; i < DASH_QUARTERS.length; i++) {
    const q = DASH_QUARTERS[i];
    if ((metrics.pendapatan && q in metrics.pendapatan) ||
        (metrics.labaBersih && q in metrics.labaBersih)) {
      return q;
    }
  }
  return null;
}

/** Cari kuartal terbaru di map (sembarang metrik), berguna untuk total aset dll. */
function latestValueOf_(map) {
  if (!map) return null;
  for (let i = 0; i < DASH_QUARTERS.length; i++) {
    if (DASH_QUARTERS[i] in map) {
      return { quarter: DASH_QUARTERS[i], value: map[DASH_QUARTERS[i]] };
    }
  }
  return null;
}

function pickLatestValues_(metrics, q) {
  if (!q) return { pendapatan: null, labaKotor: null, labaUsaha: null, labaBersih: null };
  return {
    pendapatan: (metrics.pendapatan && q in metrics.pendapatan) ? metrics.pendapatan[q] : null,
    labaKotor:  (metrics.labaKotor  && q in metrics.labaKotor)  ? metrics.labaKotor[q]  : null,
    labaUsaha:  (metrics.labaUsaha  && q in metrics.labaUsaha)  ? metrics.labaUsaha[q]  : null,
    labaBersih: (metrics.labaBersih && q in metrics.labaBersih) ? metrics.labaBersih[q] : null,
  };
}

/**
 * Buat seri tahunan dari Q4 ke belakang: hanya tahun yang punya 4 kuartal LENGKAP.
 * Mulai dari tahun yang Q4-nya tersedia dan paling baru.
 * Margin = labaBersih / pendapatan (sebagai persentase).
 */
function buildAnnualSeries_(metrics) {
  const years = [];
  const pendapatan = [];
  const labaBersih = [];
  const margin = [];

  // Cari tahun-tahun yang Q4-nya ada di pendapatan atau labaBersih
  const candidateYears = {};
  DASH_QUARTERS.forEach(function (q) {
    if (q.indexOf('Q4 ') === 0) {
      const yr = parseInt(q.split(' ')[1], 10);
      candidateYears[yr] = true;
    }
  });
  const yearList = Object.keys(candidateYears).map(Number).sort(function (a, b) { return b - a; });

  yearList.forEach(function (yr) {
    const qs = ['Q1 ' + yr, 'Q2 ' + yr, 'Q3 ' + yr, 'Q4 ' + yr];
    // Pendapatan: semua 4 kuartal harus ada
    const pendVals = qs.map(function (q) { return metrics.pendapatan[q]; });
    const labaVals = qs.map(function (q) { return metrics.labaBersih[q]; });
    const pendComplete = pendVals.every(function (v) { return v != null; });
    const labaComplete = labaVals.every(function (v) { return v != null; });
    if (!pendComplete && !labaComplete) return;  // skip jika dua-duanya tidak lengkap

    years.push(yr);
    const pendSum = pendComplete ? pendVals.reduce(function (a, b) { return a + b; }, 0) : null;
    const labaSum = labaComplete ? labaVals.reduce(function (a, b) { return a + b; }, 0) : null;
    pendapatan.push(pendSum);
    labaBersih.push(labaSum);
    // Margin: hitung apa adanya seperti permintaan user
    const m = (pendSum != null && pendSum !== 0 && labaSum != null)
              ? (labaSum / pendSum) * 100 : null;
    margin.push(m);
  });

  // Reverse supaya kiri-ke-kanan kronologis (tahun lama -> baru)
  years.reverse(); pendapatan.reverse(); labaBersih.reverse(); margin.reverse();
  return { years: years, pendapatan: pendapatan, labaBersih: labaBersih, margin: margin };
}

/** YoY: kuartal terbaru vs kuartal sama tahun lalu. */
function buildYoYComparison_(metrics, latestQuarter) {
  if (!latestQuarter) return null;
  const parts = latestQuarter.split(' ');  // ["Q1","2026"]
  const qNum = parts[0];
  const yr = parseInt(parts[1], 10);
  const prevQ = qNum + ' ' + (yr - 1);
  if (DASH_QUARTERS.indexOf(prevQ) < 0) return null;

  const pendCurr = (metrics.pendapatan && latestQuarter in metrics.pendapatan)
                   ? metrics.pendapatan[latestQuarter] : null;
  const pendPrev = (metrics.pendapatan && prevQ in metrics.pendapatan)
                   ? metrics.pendapatan[prevQ] : null;
  const labaCurr = (metrics.labaBersih && latestQuarter in metrics.labaBersih)
                   ? metrics.labaBersih[latestQuarter] : null;
  const labaPrev = (metrics.labaBersih && prevQ in metrics.labaBersih)
                   ? metrics.labaBersih[prevQ] : null;

  if (pendCurr == null && pendPrev == null && labaCurr == null && labaPrev == null) return null;

  return {
    currentQuarter: latestQuarter,
    prevYearQuarter: prevQ,
    pendapatanCurr: pendCurr,
    pendapatanPrev: pendPrev,
    labaCurr: labaCurr,
    labaPrev: labaPrev,
  };
}

/**
 * Ambil seri tahunan balance sheet (Aset, Liabilitas, Ekuitas) berbasis Q4
 * tiap tahun. Plus, kalau ada kuartal terbaru yang lebih baru dari Q4 terakhir
 * yang tersedia, append di paling kanan sebagai "kuartal berjalan".
 *
 * Contoh output untuk emiten dengan data sampai Q1 2026:
 *   labels: ['Q4 2021', 'Q4 2022', 'Q4 2023', 'Q4 2024', 'Q4 2025', 'Q1 2026']
 *
 * Contoh output untuk emiten yang sudah tutup buku Q4 saja, terbaru Q4 2025:
 *   labels: ['Q4 2021', 'Q4 2022', 'Q4 2023', 'Q4 2024', 'Q4 2025']
 *
 * Pertumbuhan ekuitas: dihitung antar titik (period-over-period) — antara dua
 * Q4 berdekatan ini setara YoY, dan antara Q4 terakhir → kuartal berjalan
 * mencerminkan perubahan parsial year. Titik pertama growth = null.
 */
function buildBalanceSeries_(metrics) {
  // 1) Kumpulkan semua Q4 yang punya data balance (aset/liab/ekui salah satu)
  const q4Entries = [];   // [{ quarter, year }]
  DASH_QUARTERS.forEach(function (q) {
    if (q.indexOf('Q4 ') !== 0) return;
    const has = (metrics.totalAset && q in metrics.totalAset) ||
                (metrics.totalLiab && q in metrics.totalLiab) ||
                (metrics.totalEkui && q in metrics.totalEkui);
    if (has) {
      q4Entries.push({ quarter: q, year: parseInt(q.split(' ')[1], 10) });
    }
  });
  // Urutkan kronologis (tahun ke tahun)
  q4Entries.sort(function (a, b) { return a.year - b.year; });

  // 2) Cari kuartal balance terbaru di seluruh data (bisa jadi bukan Q4)
  let latestQ = null;
  for (let i = 0; i < DASH_QUARTERS.length; i++) {
    const q = DASH_QUARTERS[i];
    const has = (metrics.totalAset && q in metrics.totalAset) ||
                (metrics.totalLiab && q in metrics.totalLiab) ||
                (metrics.totalEkui && q in metrics.totalEkui);
    if (has) { latestQ = q; break; }
  }

  // 3) Bangun daftar label final: semua Q4 + kuartal berjalan bila lebih baru
  const labels = q4Entries.map(function (e) { return e.quarter; });
  if (latestQ && labels.indexOf(latestQ) < 0) {
    // latestQ bukan Q4, dan lebih baru dari Q4 terakhir (karena dia ditemukan
    // duluan saat scan dari Q1 2026 ke belakang). Append.
    labels.push(latestQ);
  }

  if (labels.length === 0) {
    return { quarters: [], aset: [], liabilitas: [], ekuitas: [], ekuitasGrowth: [] };
  }

  // 4) Map nilai per label
  const aset = labels.map(function (q) {
    return (metrics.totalAset && q in metrics.totalAset) ? metrics.totalAset[q] : null;
  });
  const liab = labels.map(function (q) {
    return (metrics.totalLiab && q in metrics.totalLiab) ? metrics.totalLiab[q] : null;
  });
  const ekui = labels.map(function (q) {
    return (metrics.totalEkui && q in metrics.totalEkui) ? metrics.totalEkui[q] : null;
  });

  // 5) Pertumbuhan ekuitas period-over-period
  const ekuitasGrowth = ekui.map(function (cur, i) {
    if (i === 0) return null;
    const prev = ekui[i - 1];
    if (cur == null || prev == null || prev === 0) return null;
    return ((cur - prev) / Math.abs(prev)) * 100;
  });

  return {
    quarters: labels,
    aset: aset,
    liabilitas: liab,
    ekuitas: ekui,
    ekuitasGrowth: ekuitasGrowth
  };
}