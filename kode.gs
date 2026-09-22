/**
 * ============================================================================
 * PintarSaham - Data Dashboard Builder (v2 - Incremental & Resumable)
 * ============================================================================
 * Membaca seluruh file valuasi (PER & PBV) di "Analisa Saham/Valuasi" (rekursif)
 * lalu mengkonsolidasikan data mentah tiap emiten ke sheet "Data" pada
 * spreadsheet "Data Dashboard".
 *
 * ARSITEKTUR (v2):
 *   - Phase 1 (sekali, ~1-3 menit): enumerasi semua file Drive, buat "plan"
 *     di sheet tersembunyi "_PintarSaham_Plan", tulis header + kolom Kode +
 *     formula GOOGLEFINANCE untuk semua emiten ke sheet "Data".
 *   - Phase 2 (berulang sampai habis): proses file per file, langsung tulis
 *     ke sheet "Data" setelah dibaca. Status disimpan di plan sheet.
 *
 *   ✅ Jika timeout (6 menit Apps Script), data yang sudah ditulis tetap aman.
 *   ✅ Jalankan ulang menu "Build Data Dashboard" untuk melanjutkan dari file
 *      terakhir yang belum diproses.
 *   ✅ Rate limit (User Rate Limit Exceeded) ditangani dengan retry+backoff
 *      dan delay antar konversi.
 *
 * STRUKTUR KOLOM OUTPUT:
 *   1)   Kode
 *   2)   Harga Sekarang (=GOOGLEFINANCE)
 *   3-23)    Pendapatan Q1 2026 .. Pendapatan Q1 2021       (PER file)
 *   24-44)   Laba Kotor Q1 2026 .. Q1 2021                  (PER file)
 *   45-65)   Laba Usaha Q1 2026 .. Q1 2021                  (PER file)
 *   66-86)   Laba Bersih Tahun Berjalan Q1 2026 .. Q1 2021  (PER file)
 *   87-107)  Total Aset Q1 2026 .. Q1 2021                  (PBV file)
 *   108-128) Total Liabilitas Q1 2026 .. Q1 2021            (PBV file)
 *   129-149) Total Ekuitas Q1 2026 .. Q1 2021               (PBV file)
 *   150-153) EPS, Min PE, Mean PE, Max PE                   (PER file)
 *   154-157) BVPS, Min PBV, Mean PBV, Max PBV               (PBV file)
 *
 * ============================================================================
 * PERSIAPAN:
 * ----------------------------------------------------------------------------
 * Script ini *container-bound* di spreadsheet "Data Dashboard" (buka via
 * Extensions > Apps Script dari spreadsheet tersebut).
 *
 *   1) Sidebar Apps Script Editor: "Services" (+) > "Drive API"
 *      (v2 maupun v3 sama-sama didukung).
 *   2) Save (Ctrl+S), lalu reload tab spreadsheet "Data Dashboard".
 *   3) Menu "🟦 PintarSaham" akan muncul di toolbar.
 *
 * ALUR PEMAKAIAN:
 *   1) Klik "Preview Files" untuk cek daftar emiten yang akan diproses.
 *   2) Klik "Build Data Dashboard" -> jalankan SAMPAI SELESAI (otomatis lanjut
 *      dari progress sebelumnya tiap kali di-run).
 *   3) Jika ingin mulai ulang dari awal, klik "Reset Progress".
 * ============================================================================
 */

// ============================ CONFIG ========================================
const CONFIG = {
  ROOT_FOLDER_NAME: 'Analisa Saham',
  VALUASI_FOLDER_NAME: 'Valuasi',
  PER_FOLDER_NAME: 'Valuasi Berdasarkan PER',
  PBV_FOLDER_NAME: 'Valuasi Berdasarkan PBV',

  // ID folder "Analisa Saham" (opsional). Pakai ini jika ada beberapa folder
  // bernama sama di Drive Anda yang menyebabkan auto-detect salah pilih.
  // Cara dapat: buka folder di Drive, lihat URL drive.google.com/drive/folders/XXXX
  // -> bagian XXXX adalah ID-nya. Kosongkan untuk auto-detect by name.
  ANALISA_SAHAM_FOLDER_ID: '',

  DATA_SHEET_NAME: 'Data',
  PLAN_SHEET_NAME: '_PintarSaham_Plan', // hidden helper sheet
  LOG_SHEET_NAME:  '_PintarSaham_Log',  // hidden log issue sheet
  PRICE_TICKER_PREFIX: 'IDX:',

  // Whitelist emiten: hanya emiten yang kodenya terdaftar di sheet ini
  // (kolom A, mulai A2) yang akan diproses. Kosongkan WHITELIST_SHEET_NAME
  // ('') untuk memproses SEMUA file di Drive (perilaku lama).
  WHITELIST_SHEET_NAME: 'Emiten',
  WHITELIST_COLUMN: 1,       // kolom A
  WHITELIST_START_ROW: 2,    // mulai baris 2 (baris 1 = header)

  // Batasi waktu run agar berhenti rapi sebelum Apps Script timeout (6 menit).
  // 270 detik = 4.5 menit, sisakan 1.5 menit untuk cleanup + schedule trigger.
  MAX_RUNTIME_SECONDS: 270,

  // Delay antar konversi file (ms) untuk menghindari Drive API rate limit.
  // Naikkan jika sering kena "User Rate Limit Exceeded".
  CONVERT_DELAY_MS: 800,

  // Maksimum percobaan ulang per file saat rate limit / error transient.
  MAX_RETRIES: 4,

  // Hapus file XLSX hasil konversi sementara setelah dibaca.
  CLEAN_TEMP_FILES: true,

  // ⚡ Auto-resume: setelah run berhenti karena timeout, otomatis schedule
  // trigger waktu untuk lanjutkan run berikutnya. Anda tinggal klik "Build"
  // SEKALI, lalu biarkan script berjalan otomatis sampai semua selesai.
  AUTO_RESUME: true,
  AUTO_RESUME_DELAY_SECONDS: 60,
};

const QUARTERS = [
  'Q1 2026', 'Q4 2025', 'Q3 2025', 'Q2 2025', 'Q1 2025',
  'Q4 2024', 'Q3 2024', 'Q2 2024', 'Q1 2024',
  'Q4 2023', 'Q3 2023', 'Q2 2023', 'Q1 2023',
  'Q4 2022', 'Q3 2022', 'Q2 2022', 'Q1 2022',
  'Q4 2021', 'Q3 2021', 'Q2 2021', 'Q1 2021'
];

// Catatan: tiap label DI-NORMALISASI (lowercase + collapse whitespace + trim)
// sebelum dicocokkan. Variant tanpa "K" ditambahkan untuk menangani kasus
// data Stockbit di-paste lalu user Replace All "K" (untuk hilangkan suffix
// ribuan) yang accidental juga menghapus K dari label.
const METRIC_PER = [
  { key: 'Pendapatan',                  labels: ['total pendapatan', 'pendapatan'] },
  { key: 'Laba Kotor',                  labels: ['laba kotor', 'laba otor'] },  // K hilang
  { key: 'Laba Usaha',                  labels: ['laba usaha'] },
  { key: 'Laba Bersih Tahun Berjalan',  labels: ['laba bersih tahun berjalan'] },
];

const METRIC_PBV = [
  { key: 'Total Aset',        labels: ['total aset'] },
  { key: 'Total Liabilitas',  labels: ['total liabilitas'] },
  { key: 'Total Ekuitas',     labels: ['total ekuitas', 'total euitas'] },  // k hilang
];

const SHEET_NAME_CANDIDATES = {
  KUARTAL: [
    'Laporan Kuartal', 'Laporan Kuartal ',
    'Laporan per Kuartal', 'Laporan per Kuartal ',
    'Laporan Kuartalan',
  ],
  PROYEKSI: ['Proyeksi', 'Proyeksi '],
};

// Column index map (1-based, untuk sheet "Data")
const COL = {
  KODE: 1,
  HARGA: 2,
  PENDAPATAN_START: 3,    // 3..23
  LABA_KOTOR_START: 24,   // 24..44
  LABA_USAHA_START: 45,   // 45..65
  LABA_BERSIH_START: 66,  // 66..86
  TOTAL_ASET_START: 87,   // 87..107
  TOTAL_LIAB_START: 108,  // 108..128
  TOTAL_EKUI_START: 129,  // 129..149
  EPS: 150,
  MIN_PE: 151,
  MEAN_PE: 152,
  MAX_PE: 153,
  BVPS: 154,
  MIN_PBV: 155,
  MEAN_PBV: 156,
  MAX_PBV: 157,
};
const TOTAL_COLS = 157;

// ============================ MENU ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🟦 PintarSaham')
    .addItem('▶ Build Data Dashboard (auto-resume)', 'buildDataDashboard')
    .addItem('⏸ Stop Auto-Resume', 'stopAutoResume')
    .addSeparator()
    .addItem('🔍 Preview Files', 'previewFiles')
    .addItem('📊 Lihat Progress', 'showProgress')
    .addItem('📋 Lihat Issue Log', 'showIssues')
    .addItem('⏰ Status Trigger Auto-Resume', 'showTriggers')
    .addItem('🔧 Diagnose Folder', 'diagnoseFolder')
    .addSeparator()
    .addItem('🔁 Retry Files dengan Issue', 'retryIssuedFiles')
    .addItem('🔁 Retry Empty Files (deteksi otomatis)', 'retryEmptyFiles')
    .addItem('🔁 Retry by Code (manual)', 'retryByCode')
    .addSeparator()
    .addItem('♻ Reset Progress (mulai ulang)', 'resetProgress')
    .addToUi();
}

// ============================ MAIN ==========================================

/**
 * Fungsi utama. Jalankan berulang sampai semua file selesai diproses.
 * Bisa dipanggil manual via menu, atau otomatis via time-based trigger
 * (auto-resume) bila CONFIG.AUTO_RESUME aktif.
 */
function buildDataDashboard() {
  // Lock untuk mencegah run paralel (manual + trigger bersamaan).
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) {
    const note = 'Run lain sedang berjalan — kemungkinan besar auto-resume ' +
                 'trigger yang sedang aktif (artinya: auto-resume BERFUNGSI ✓). ' +
                 'Tunggu sampai selesai (max ~5 menit). ' +
                 'Cek menu "⏰ Status Trigger Auto-Resume" untuk detail.';
    Logger.log(note);
    try {
      const ss = locateDashboard_();
      ss.toast(note, '🟦 PintarSaham — run aktif', 15);
    } catch (e) {}
    return;
  }

  try {
    runBuildCycle_();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function runBuildCycle_() {
  const t0 = Date.now();
  Logger.log('=== buildDataDashboard mulai: ' + new Date(t0) + ' ===');

  const ss = locateDashboard_();
  const dataSheet = ensureDataSheet_(ss);
  let planSheet = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);

  // ---- Phase 1: setup plan & header bila belum ada ----
  let plan;
  if (!planSheet || planSheet.getLastRow() < 2) {
    ss.toast('Phase 1: enumerasi file Drive...', '🟦 PintarSaham', 5);
    plan = setupPlanAndHeader_(ss, dataSheet);
    planSheet = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);
    Logger.log('Phase 1 selesai. ' + plan.length + ' emiten terdaftar.');
  } else {
    plan = readPlan_(planSheet);
    Logger.log('Plan ditemukan: ' + plan.length + ' emiten.');
  }

  // ---- Phase 2: proses file satu per satu ----
  let processedThisRun = 0;
  let partialThisRun = 0;
  let errors = 0;
  let stoppedByTimeout = false;

  for (let i = 0; i < plan.length; i++) {
    const item = plan[i];

    // Cek batas waktu
    const elapsed = (Date.now() - t0) / 1000;
    if (elapsed > CONFIG.MAX_RUNTIME_SECONDS) {
      stoppedByTimeout = true;
      Logger.log('Stop karena MAX_RUNTIME tercapai (' + elapsed.toFixed(0) + 's).');
      break;
    }

    // Skip kalau kedua file sudah diproses
    const needPer = item.perFileId && !item.perDone;
    const needPbv = item.pbvFileId && !item.pbvDone;
    if (!needPer && !needPbv) continue;

    const remainingMin = Math.max(0, (CONFIG.MAX_RUNTIME_SECONDS - elapsed) / 60);
    ss.toast('[' + (i + 1) + '/' + plan.length + '] ' + item.code +
             ' (~' + remainingMin.toFixed(1) + ' min tersisa)',
             '🟦 PintarSaham', 3);

    // Proses PER
    if (needPer) {
      const r = processOneFile_(item, 'PER', dataSheet);
      if (r.success) {
        markStatus_(planSheet, item.planRow, 'PER');
        processedThisRun++;
        if (r.status !== 'OK') {
          logIssue_(ss, item.code, 'PER', r.status, r.issues, r.filename);
          partialThisRun++;
        }
      } else {
        logIssue_(ss, item.code, 'PER', 'ERROR', r.issues, r.filename);
        errors++;
      }
      Utilities.sleep(CONFIG.CONVERT_DELAY_MS);
    }

    // Cek timeout lagi sebelum PBV
    if ((Date.now() - t0) / 1000 > CONFIG.MAX_RUNTIME_SECONDS) {
      stoppedByTimeout = true;
      break;
    }

    // Proses PBV
    if (needPbv) {
      const r = processOneFile_(item, 'PBV', dataSheet);
      if (r.success) {
        markStatus_(planSheet, item.planRow, 'PBV');
        processedThisRun++;
        if (r.status !== 'OK') {
          logIssue_(ss, item.code, 'PBV', r.status, r.issues, r.filename);
          partialThisRun++;
        }
      } else {
        logIssue_(ss, item.code, 'PBV', 'ERROR', r.issues, r.filename);
        errors++;
      }
      Utilities.sleep(CONFIG.CONVERT_DELAY_MS);
    }
  }

  // ---- Summary ----
  const dur = ((Date.now() - t0) / 1000).toFixed(1);
  const remaining = countRemaining_(planSheet);

  // Bersihkan trigger lama supaya tidak menumpuk.
  cleanupResumeTriggers_();

  const stats = processedThisRun + ' file diproses' +
                (partialThisRun > 0 ? ' (' + partialThisRun + ' partial)' : '') +
                (errors > 0 ? ', ' + errors + ' error' : '');

  let msg;
  let scheduledNext = false;
  if (remaining === 0) {
    msg = '🎉 SELESAI! Semua file sudah diproses (' + plan.length + ' emiten). ' +
          'Run terakhir: ' + dur + 's, ' + stats + '.';
  } else if (stoppedByTimeout) {
    msg = 'Pause setelah ' + dur + 's. ' + stats +
          '. Tersisa ~' + remaining + ' file.';
    if (CONFIG.AUTO_RESUME) {
      scheduleResume_();
      scheduledNext = true;
      msg += ' ⏰ Auto-resume dalam ' + CONFIG.AUTO_RESUME_DELAY_SECONDS + 's.';
    } else {
      msg += ' Jalankan "Build" lagi untuk melanjutkan.';
    }
  } else {
    msg = 'Selesai run dalam ' + dur + 's. ' + stats +
          '. Tersisa ~' + remaining + ' file.';
  }
  Logger.log('=== ' + msg + ' ===');
  // Toast (non-blocking). Akan terlihat bila tab terbuka, terlewat bila tidak — aman.
  ss.toast(msg, '🟦 PintarSaham', 30);

  // Alert (blocking) HANYA saat fully selesai. Saat resume-cycle tidak dipakai
  // supaya tidak hang menunggu user klik OK & kena hard timeout.
  if (remaining === 0) {
    try {
      SpreadsheetApp.getUi().alert('🎉 PintarSaham — Selesai',
                                    msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {}
  }
}

/**
 * Jadwalkan run berikutnya via time-based trigger.
 */
function scheduleResume_() {
  ScriptApp.newTrigger('buildDataDashboard')
    .timeBased()
    .after(CONFIG.AUTO_RESUME_DELAY_SECONDS * 1000)
    .create();
  Logger.log('Auto-resume trigger dijadwalkan dalam ' +
             CONFIG.AUTO_RESUME_DELAY_SECONDS + ' detik.');
}

/**
 * Hapus semua trigger auto-resume yang ada (mencegah duplikat).
 */
function cleanupResumeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'buildDataDashboard') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (removed > 0) Logger.log('Bersihkan ' + removed + ' trigger auto-resume lama.');
}

/**
 * Manual stop: hapus semua trigger auto-resume yang sudah dijadwalkan.
 * Dipanggil dari menu kalau user mau pause permanently.
 */
/**
 * Tampilkan log issue. Unhide sheet log dan switch ke sheet itu supaya user
 * bisa langsung browse daftar file yang bermasalah dengan detail issue-nya.
 */
/**
 * Identifikasi file yang setelah diproses datanya kosong total (PER atau PBV),
 * lalu tandai untuk retry. Berguna untuk men-trace alasan file tidak terbaca
 * (terutama untuk file yang diproses SEBELUM logging code aktif).
 */
function retryEmptyFiles() {
  const ui = SpreadsheetApp.getUi();
  const ss = locateDashboard_();
  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  const planSheet = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);

  if (!dataSheet || !planSheet ||
      dataSheet.getLastRow() < 2 || planSheet.getLastRow() < 2) {
    ui.alert('Belum ada data atau plan untuk diperiksa.');
    return;
  }

  const totalRows = Math.min(dataSheet.getLastRow() - 1, planSheet.getLastRow() - 1);
  const dataValues = dataSheet.getRange(2, 1, totalRows, TOTAL_COLS).getValues();
  const planDone = planSheet.getRange(2, 4, totalRows, 2).getValues(); // PER_Done, PBV_Done

  const isEmpty = function (v) { return v === '' || v === null || v === undefined; };
  const allEmpty = function (row, startCol1, endCol1) {
    for (let c = startCol1 - 1; c <= endCol1 - 1; c++) {
      if (!isEmpty(row[c])) return false;
    }
    return true;
  };

  const perRetry = [], pbvRetry = [];

  for (let i = 0; i < totalRows; i++) {
    const row = dataValues[i];
    const code = row[0];
    const planRow = i + 2;
    const perDone = planDone[i][0] === true || planDone[i][0] === 'TRUE';
    const pbvDone = planDone[i][1] === true || planDone[i][1] === 'TRUE';

    // PER kosong total: SEMUA cells kuartalan + EPS/PE empty, dan sudah ditandai Done
    if (perDone) {
      const perKuartalEmpty = allEmpty(row, COL.PENDAPATAN_START,
                                       COL.LABA_BERSIH_START + QUARTERS.length - 1);
      const perProjEmpty = allEmpty(row, COL.EPS, COL.MAX_PE);
      if (perKuartalEmpty && perProjEmpty) {
        perRetry.push({ code: code, planRow: planRow });
      }
    }

    // PBV kosong total
    if (pbvDone) {
      const pbvKuartalEmpty = allEmpty(row, COL.TOTAL_ASET_START,
                                        COL.TOTAL_EKUI_START + QUARTERS.length - 1);
      const pbvProjEmpty = allEmpty(row, COL.BVPS, COL.MAX_PBV);
      if (pbvKuartalEmpty && pbvProjEmpty) {
        pbvRetry.push({ code: code, planRow: planRow });
      }
    }
  }

  if (perRetry.length === 0 && pbvRetry.length === 0) {
    ui.alert('🎉 Tidak Ada File Kosong',
             'Semua file yang sudah diproses memiliki setidaknya sebagian data.\n\n' +
             'Untuk lihat issue PARSIAL (data hanya sebagian), klik ' +
             '"📋 Lihat Issue Log".',
             ui.ButtonSet.OK);
    return;
  }

  // Preview
  let msg = 'File dengan DATA KOSONG TOTAL setelah diproses:\n\n';
  msg += '• PER: ' + perRetry.length + ' file\n';
  msg += '• PBV: ' + pbvRetry.length + ' file\n\n';

  if (perRetry.length > 0) {
    const sample = perRetry.slice(0, 15).map(function (x) { return x.code; });
    msg += 'PER (15 pertama): ' + sample.join(', ');
    if (perRetry.length > 15) msg += ', ... +' + (perRetry.length - 15) + ' lagi';
    msg += '\n\n';
  }
  if (pbvRetry.length > 0) {
    const sample = pbvRetry.slice(0, 15).map(function (x) { return x.code; });
    msg += 'PBV (15 pertama): ' + sample.join(', ');
    if (pbvRetry.length > 15) msg += ', ... +' + (pbvRetry.length - 15) + ' lagi';
    msg += '\n\n';
  }

  msg += 'Tandai file-file ini untuk RETRY?\n' +
         'Saat run berikutnya (Build / auto-resume), file ini akan ' +
         'diproses ulang dan issue log akan mencatat alasannya.';

  const resp = ui.alert('🔁 Retry Empty Files', msg, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  // Clear Done flag dengan batch write (lebih efisien)
  perRetry.forEach(function (x) {
    planSheet.getRange(x.planRow, 4).setValue(false);
  });
  pbvRetry.forEach(function (x) {
    planSheet.getRange(x.planRow, 5).setValue(false);
  });

  const total = perRetry.length + pbvRetry.length;
  ss.toast('Ditandai ' + total + ' file untuk retry. ' +
           'Klik "Build Data Dashboard" untuk mulai re-process.',
           '🟦 PintarSaham', 10);
}

function showIssues() {
  const ss = locateDashboard_();
  const log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!log || log.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('📋 Issue Log',
      'Belum ada issue ter-log. Semua file sejauh ini terbaca tanpa masalah.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  log.showSheet();
  ss.setActiveSheet(log);
  const totalIssues = log.getLastRow() - 1;
  // Quick summary
  const data = log.getRange(2, 4, totalIssues, 1).getValues();
  const byStatus = {};
  data.forEach(function (r) {
    const s = r[0];
    byStatus[s] = (byStatus[s] || 0) + 1;
  });
  const summary = Object.keys(byStatus).map(function (k) {
    return k + ': ' + byStatus[k];
  }).join(', ');
  ss.toast('Total issue ter-log: ' + totalIssues + ' (' + summary + ')',
           '📋 Issue Log', 10);
}

/**
 * Tandai file-file yang ada di issue log untuk di-retry pada Build berikutnya.
 * Berguna setelah script di-update untuk menangani format baru (label variant,
 * nama sheet alternatif, dll) — file PARTIAL/EMPTY/ERROR sebelumnya mungkin
 * sekarang bisa terbaca dengan benar.
 *
 * Cara kerja:
 *   1) Baca semua entri unik (code, type) dari _PintarSaham_Log
 *   2) Set kolom PER_Done atau PBV_Done di _PintarSaham_Plan kembali ke FALSE
 *   3) Bersihkan issue log (akan diisi ulang pada Build berikutnya)
 *   4) User tinggal klik "Build Data Dashboard" untuk mulai retry
 */
function retryIssuedFiles() {
  const ui = SpreadsheetApp.getUi();
  const ss = locateDashboard_();
  const log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  const plan = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);

  if (!log || log.getLastRow() < 2) {
    ui.alert('🔁 Retry', 'Tidak ada file di issue log untuk di-retry.', ui.ButtonSet.OK);
    return;
  }
  if (!plan || plan.getLastRow() < 2) {
    ui.alert('🔁 Retry', 'Plan sheet kosong. Jalankan Build dulu.', ui.ButtonSet.OK);
    return;
  }

  // Ekstrak (code, type) unik dari log
  const logData = log.getRange(2, 1, log.getLastRow() - 1, 4).getValues();
  const pairs = {};
  logData.forEach(function (row) {
    const code = String(row[1] || '').trim();
    const type = String(row[2] || '').trim();
    if (!code || (type !== 'PER' && type !== 'PBV')) return;
    pairs[code + '|' + type] = { code: code, type: type };
  });
  const pairList = Object.keys(pairs).map(function (k) { return pairs[k]; });

  if (pairList.length === 0) {
    ui.alert('🔁 Retry', 'Tidak ada entri valid di issue log.', ui.ButtonSet.OK);
    return;
  }

  // Konfirmasi
  const sample = pairList.slice(0, 10).map(function (p) {
    return p.code + ' (' + p.type + ')';
  }).join(', ');
  const more = pairList.length > 10 ? ' ... +' + (pairList.length - 10) + ' lainnya' : '';
  const confirm = ui.alert(
    '🔁 Retry Files dengan Issue',
    'Akan menandai ' + pairList.length + ' file dari issue log untuk di-retry ' +
    'pada Build berikutnya.\n\n' +
    'Contoh: ' + sample + more + '\n\n' +
    'Setelah konfirmasi:\n' +
    '  1) Status Done di-reset untuk file-file tsb\n' +
    '  2) Issue log dibersihkan\n' +
    '  3) Jalankan menu "▶ Build Data Dashboard" untuk memproses ulang\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // Build lookup: code → planRow
  const planData = plan.getRange(2, 1, plan.getLastRow() - 1, 1).getValues();
  const codeToRow = {};
  planData.forEach(function (row, idx) {
    codeToRow[String(row[0]).trim()] = idx + 2;
  });

  let updated = 0;
  const notFound = [];
  pairList.forEach(function (p) {
    const row = codeToRow[p.code];
    if (!row) { notFound.push(p.code + ' (' + p.type + ')'); return; }
    const col = p.type === 'PER' ? 4 : 5;   // PER_Done = col 4, PBV_Done = col 5
    plan.getRange(row, col).setValue(false);
    plan.getRange(row, 6).setValue('');     // clear Updated timestamp
    updated++;
  });

  // Bersihkan log: hanya keep header
  log.clear();
  log.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Kode', 'Type',
                                        'Status', 'Filename', 'Issues']]);
  log.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#fce8b2');
  log.setFrozenRows(1);

  let msg = '✓ ' + updated + ' file ditandai untuk retry. Issue log dibersihkan.';
  if (notFound.length > 0) {
    msg += '\n\n⚠ ' + notFound.length + ' kode di log tidak ada di plan: ' +
           notFound.slice(0, 8).join(', ') +
           (notFound.length > 8 ? ', ...' : '');
  }
  msg += '\n\nLangkah berikut: klik "▶ Build Data Dashboard" untuk memproses ulang.';
  ui.alert('🔁 Retry', msg, ui.ButtonSet.OK);
}

/**
 * Retry by code list. Berguna ketika user ingin menjalankan ulang emiten
 * spesifik (mis: file di Drive sudah diperbaiki manual, atau ingin re-verify
 * data tertentu) tanpa harus reset progress total.
 */
function retryByCode() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '🔁 Retry by Code',
    'Masukkan kode emiten yang ingin di-retry, pisahkan dengan koma atau baris baru. ' +
    'PER dan PBV keduanya akan di-retry.\n\nContoh: BBRI, ANTM, HOKI',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const input = resp.getResponseText().trim();
  if (!input) return;

  const codes = input.split(/[,\s\n\r]+/)
    .map(function (s) { return s.trim().toUpperCase(); })
    .filter(function (s) { return s.length > 0; });

  if (codes.length === 0) {
    ui.alert('Tidak ada kode valid dari input.');
    return;
  }

  const ss = locateDashboard_();
  const plan = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);
  if (!plan || plan.getLastRow() < 2) {
    ui.alert('Plan sheet belum ada. Jalankan Build dulu untuk inisialisasi.');
    return;
  }

  const planData = plan.getRange(2, 1, plan.getLastRow() - 1, 1).getValues();
  const codeToRow = {};
  planData.forEach(function (row, idx) {
    codeToRow[String(row[0]).trim().toUpperCase()] = idx + 2;
  });

  let updated = 0;
  const notFound = [];
  codes.forEach(function (c) {
    const row = codeToRow[c];
    if (!row) { notFound.push(c); return; }
    plan.getRange(row, 4).setValue(false);  // PER_Done
    plan.getRange(row, 5).setValue(false);  // PBV_Done
    plan.getRange(row, 6).setValue('');
    updated++;
  });

  let msg = '✓ ' + updated + ' kode ditandai untuk retry (PER + PBV).';
  if (notFound.length > 0) {
    msg += '\n\n⚠ Kode tidak ditemukan di plan: ' + notFound.join(', ') +
           '\nPastikan ejaan benar dan file ada di Drive.';
  }
  msg += '\n\nLangkah berikut: klik "▶ Build Data Dashboard" untuk memproses ulang.';
  ui.alert('🔁 Retry by Code', msg, ui.ButtonSet.OK);
}

function stopAutoResume() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'buildDataDashboard') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  const ss = locateDashboard_();
  const msg = count > 0
    ? 'Dibatalkan: ' + count + ' trigger auto-resume. Jalankan "Build" manual untuk lanjut.'
    : 'Tidak ada trigger auto-resume yang aktif.';
  ss.toast(msg, '🟦 PintarSaham', 8);
  try {
    SpreadsheetApp.getUi().alert('🟦 PintarSaham', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
}

/**
 * Diagnostic: tampilkan status trigger auto-resume + lock status.
 * Membantu cek apakah ada run aktif sedang berjalan.
 */
function showTriggers() {
  const ui = SpreadsheetApp.getUi();
  const all = ScriptApp.getProjectTriggers();
  const ours = all.filter(function (t) {
    return t.getHandlerFunction() === 'buildDataDashboard';
  });

  // Snapshot lock status: tryLock(0) -> success berarti lock SAAT INI bebas
  const lock = LockService.getDocumentLock();
  const lockFree = lock.tryLock(0);
  if (lockFree) { try { lock.releaseLock(); } catch (e) {} }
  const locked = !lockFree;

  let msg = '';
  msg += locked ? '🔒 Lock: TERKUNCI — ada run aktif sedang processing\n'
                : '🔓 Lock: BEBAS — tidak ada run aktif saat ini\n';
  msg += '⏰ Trigger auto-resume aktif: ' + ours.length + '\n';
  msg += 'Total trigger di project: ' + all.length + '\n\n';

  if (locked) {
    msg += '💡 Run sedang berjalan. Tunggu max ~5 menit sebelum klik "Build" manual. ' +
           'Cek Apps Script > Executions untuk detail run aktif.\n\n';
  }

  if (ours.length === 0) {
    msg += 'Tidak ada trigger auto-resume terjadwal.\n';
    msg += 'Kemungkinan: build sudah selesai, atau trigger sudah fire & terhapus, ' +
           'atau belum klik "Build" sejak update.';
  } else {
    ours.forEach(function (t, i) {
      msg += 'Trigger #' + (i + 1) + ':\n';
      msg += '  Handler: ' + t.getHandlerFunction() + '\n';
      msg += '  Event:   ' + t.getEventType() + '\n';
      msg += '  Source:  ' + t.getTriggerSource() + '\n\n';
    });
    if (!locked) {
      msg += '⏰ Trigger akan fire dalam 1–3 menit (timing oleh Google).';
    }
  }

  msg += '\n\nUntuk lihat run history aktual:\n' +
         '  Apps Script Editor > sidebar kiri > ⏱ Executions';

  Logger.log(msg);
  ui.alert('⏰ Status', msg, ui.ButtonSet.OK);
}

/**
 * Reset progress: hapus plan sheet & kosongkan Data sheet, supaya run berikutnya
 * mulai dari nol. Minta konfirmasi user dulu.
 */
function resetProgress() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '♻ Reset Progress',
    'Tindakan ini akan menghapus plan sheet dan mengosongkan sheet "' +
    CONFIG.DATA_SHEET_NAME + '". Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = locateDashboard_();
  const plan = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);
  if (plan) ss.deleteSheet(plan);
  const log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (log) ss.deleteSheet(log);
  const data = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  if (data) data.clearContents();
  cleanupResumeTriggers_();  // hentikan auto-resume yang mungkin sedang antri
  ss.toast('Progress di-reset. Jalankan "Build Data Dashboard" untuk mulai dari awal.',
           '🟦 PintarSaham', 8);
}

/**
 * Tampilkan ringkasan progress dari plan sheet.
 */
function showProgress() {
  const ss = locateDashboard_();
  const ps = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);
  if (!ps || ps.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Belum ada plan. Jalankan "Build Data Dashboard" dulu.');
    return;
  }
  const plan = readPlan_(ps);
  let totalPer = 0, donePer = 0, totalPbv = 0, donePbv = 0;
  plan.forEach(function (it) {
    if (it.perFileId) { totalPer++; if (it.perDone) donePer++; }
    if (it.pbvFileId) { totalPbv++; if (it.pbvDone) donePbv++; }
  });
  const pctPer = totalPer ? ((donePer / totalPer) * 100).toFixed(1) : '0.0';
  const pctPbv = totalPbv ? ((donePbv / totalPbv) * 100).toFixed(1) : '0.0';
  const remaining = (totalPer - donePer) + (totalPbv - donePbv);
  const msg = 'Total emiten: ' + plan.length + '\n' +
              'PER:  ' + donePer + '/' + totalPer + ' (' + pctPer + '%)\n' +
              'PBV:  ' + donePbv + '/' + totalPbv + ' (' + pctPbv + '%)\n' +
              'File tersisa: ' + remaining;
  SpreadsheetApp.getUi().alert('📊 Progress', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Preview daftar file yang akan diproses (tidak menulis ke sheet).
 */
function previewFiles() {
  const ss = locateDashboard_();
  ss.toast('Menelusuri folder...', '🟦 PintarSaham Preview', 5);
  const whitelist = readWhitelist_(ss);
  const folders = locateValuasiFolders_();
  let perFiles = collectValuationFiles_(folders.perFolder, 'PER');
  let pbvFiles = collectValuationFiles_(folders.pbvFolder, 'PBV');

  const perRaw = perFiles.length, pbvRaw = pbvFiles.length;
  if (whitelist) {
    perFiles = perFiles.filter(function (it) { return whitelist.has(it.code); });
    pbvFiles = pbvFiles.filter(function (it) { return whitelist.has(it.code); });
  }

  Logger.log('=== Preview ===');
  if (whitelist) {
    Logger.log('Whitelist AKTIF: ' + whitelist.size + ' kode. ' +
               'PER ' + perRaw + '->' + perFiles.length + ', ' +
               'PBV ' + pbvRaw + '->' + pbvFiles.length + ' setelah filter.');
  } else {
    Logger.log('Whitelist tidak aktif (memproses semua file).');
  }
  Logger.log('Total PER: ' + perFiles.length);
  perFiles.forEach(function (it, i) {
    Logger.log('  ' + (i + 1) + '. ' + it.code + ' -- ' + it.file.getName() +
               ' [' + (it.file.getMimeType() === MimeType.GOOGLE_SHEETS ? 'GSheet' : 'XLSX') + ']');
  });
  Logger.log('Total PBV: ' + pbvFiles.length);
  pbvFiles.forEach(function (it, i) {
    Logger.log('  ' + (i + 1) + '. ' + it.code + ' -- ' + it.file.getName() +
               ' [' + (it.file.getMimeType() === MimeType.GOOGLE_SHEETS ? 'GSheet' : 'XLSX') + ']');
  });

  const perSet = {}; perFiles.forEach(function (it) { perSet[it.code] = true; });
  const pbvSet = {}; pbvFiles.forEach(function (it) { pbvSet[it.code] = true; });
  const onlyPer = Object.keys(perSet).filter(function (k) { return !pbvSet[k]; });
  const onlyPbv = Object.keys(pbvSet).filter(function (k) { return !perSet[k]; });
  if (onlyPer.length) Logger.log('Hanya PER (tanpa PBV): ' + onlyPer.join(', '));
  if (onlyPbv.length) Logger.log('Hanya PBV (tanpa PER): ' + onlyPbv.join(', '));

  // Kode di whitelist yang tidak ada file-nya
  if (whitelist) {
    const found = {};
    perFiles.concat(pbvFiles).forEach(function (it) { found[it.code] = true; });
    const missing = [];
    whitelist.forEach(function (c) { if (!found[c]) missing.push(c); });
    if (missing.length) Logger.log('⚠ Whitelist tanpa file di Drive: ' + missing.join(', '));
  }

  const wlNote = whitelist ? ' (whitelist: ' + whitelist.size + ' kode)' : '';
  ss.toast('PER: ' + perFiles.length + ', PBV: ' + pbvFiles.length + wlNote +
           '. Detail di View > Executions.', '🟦 PintarSaham Preview', 10);
}

// ============================ PHASE 1: SETUP ================================

/**
 * Enumerasi semua file Drive, buat plan sheet & header sheet Data,
 * tulis kolom Kode + GOOGLEFINANCE untuk semua emiten.
 * Return: array plan items.
 */
function setupPlanAndHeader_(ss, dataSheet) {
  // Baca whitelist emiten (bila dikonfigurasi)
  const whitelist = readWhitelist_(ss);
  if (whitelist) {
    Logger.log('Whitelist aktif: ' + whitelist.size + ' kode terdaftar di sheet "' +
               CONFIG.WHITELIST_SHEET_NAME + '"');
  } else {
    Logger.log('Whitelist tidak aktif: memproses SEMUA file di Drive.');
  }

  const folders = locateValuasiFolders_();
  Logger.log('Enumerasi PER...');
  let perFiles = collectValuationFiles_(folders.perFolder, 'PER');
  Logger.log('  -> ' + perFiles.length + ' file (sebelum filter)');
  Logger.log('Enumerasi PBV...');
  let pbvFiles = collectValuationFiles_(folders.pbvFolder, 'PBV');
  Logger.log('  -> ' + pbvFiles.length + ' file (sebelum filter)');

  // Filter berdasarkan whitelist
  if (whitelist) {
    perFiles = perFiles.filter(function (it) { return whitelist.has(it.code); });
    pbvFiles = pbvFiles.filter(function (it) { return whitelist.has(it.code); });
    Logger.log('Setelah filter whitelist: PER=' + perFiles.length +
               ', PBV=' + pbvFiles.length);
  }

  // Gabungkan jadi mapping per-kode
  const byCode = {};
  perFiles.forEach(function (it) {
    byCode[it.code] = byCode[it.code] || { code: it.code };
    byCode[it.code].perFileId = it.file.getId();
  });
  pbvFiles.forEach(function (it) {
    byCode[it.code] = byCode[it.code] || { code: it.code };
    byCode[it.code].pbvFileId = it.file.getId();
  });

  const codes = Object.keys(byCode).sort();

  // Diagnostik: kode di whitelist yang TIDAK ketemu file-nya di Drive
  if (whitelist) {
    const missing = [];
    whitelist.forEach(function (code) {
      if (!byCode[code]) missing.push(code);
    });
    if (missing.length > 0) {
      Logger.log('⚠ ' + missing.length + ' kode di whitelist tidak ada file PER/PBV di Drive: ' +
                 missing.join(', '));
    }
  }

  const plan = codes.map(function (code, idx) {
    const e = byCode[code];
    return {
      code: code,
      perFileId: e.perFileId || '',
      pbvFileId: e.pbvFileId || '',
      perDone: false,
      pbvDone: false,
      planRow: idx + 2,  // row 1 = header
    };
  });

  // ---- Tulis plan sheet (hidden) ----
  let ps = ss.getSheetByName(CONFIG.PLAN_SHEET_NAME);
  if (ps) ss.deleteSheet(ps);
  ps = ss.insertSheet(CONFIG.PLAN_SHEET_NAME);
  const planHeader = ['Kode', 'PER_FileID', 'PBV_FileID', 'PER_Done', 'PBV_Done', 'Updated'];
  const planRows = [planHeader].concat(plan.map(function (p) {
    return [p.code, p.perFileId, p.pbvFileId, false, false, ''];
  }));
  ps.getRange(1, 1, planRows.length, planHeader.length).setValues(planRows);
  ps.getRange(1, 1, 1, planHeader.length).setFontWeight('bold').setBackground('#dadce0');
  ps.setFrozenRows(1);
  ps.hideSheet();

  // ---- Tulis header sheet Data ----
  const header = buildHeader_();
  dataSheet.clearContents();
  dataSheet.getRange(1, 1, 1, TOTAL_COLS).setValues([header]);
  dataSheet.getRange(1, 1, 1, TOTAL_COLS)
           .setFontWeight('bold')
           .setBackground('#0f2240')
           .setFontColor('#ffffff')
           .setHorizontalAlignment('center');
  dataSheet.setFrozenRows(1);
  dataSheet.setFrozenColumns(2);

  // ---- Tulis Kode + formula GOOGLEFINANCE untuk semua emiten ----
  if (plan.length > 0) {
    const kodeCol = plan.map(function (p) { return [p.code]; });
    const hargaCol = plan.map(function (p) {
      return ['=IFERROR(GOOGLEFINANCE("' + CONFIG.PRICE_TICKER_PREFIX + p.code + '","price"),)'];
    });
    dataSheet.getRange(2, COL.KODE, plan.length, 1).setValues(kodeCol);
    dataSheet.getRange(2, COL.HARGA, plan.length, 1).setValues(hargaCol);
  }

  SpreadsheetApp.flush();
  return plan;
}

function buildHeader_() {
  const h = new Array(TOTAL_COLS);
  h[COL.KODE - 1] = 'Kode';
  h[COL.HARGA - 1] = 'Harga Sekarang';
  const blocks = [
    { name: 'Pendapatan',                 start: COL.PENDAPATAN_START },
    { name: 'Laba Kotor',                 start: COL.LABA_KOTOR_START },
    { name: 'Laba Usaha',                 start: COL.LABA_USAHA_START },
    { name: 'Laba Bersih Tahun Berjalan', start: COL.LABA_BERSIH_START },
    { name: 'Total Aset',                 start: COL.TOTAL_ASET_START },
    { name: 'Total Liabilitas',           start: COL.TOTAL_LIAB_START },
    { name: 'Total Ekuitas',              start: COL.TOTAL_EKUI_START },
  ];
  blocks.forEach(function (b) {
    for (let i = 0; i < QUARTERS.length; i++) {
      h[b.start - 1 + i] = b.name + ' ' + QUARTERS[i];
    }
  });
  h[COL.EPS - 1] = 'EPS';
  h[COL.MIN_PE - 1] = 'Min PE';
  h[COL.MEAN_PE - 1] = 'Mean PE';
  h[COL.MAX_PE - 1] = 'Max PE';
  h[COL.BVPS - 1] = 'BVPS';
  h[COL.MIN_PBV - 1] = 'Min PBV';
  h[COL.MEAN_PBV - 1] = 'Mean PBV';
  h[COL.MAX_PBV - 1] = 'Max PBV';
  return h;
}

// ============================ PHASE 2: PROCESS ==============================

/**
 * Proses satu file (PER atau PBV) dan langsung tulis hasilnya ke sheet Data.
 * Return: { success, status, issues, filename }
 *   - success = false : file gagal dibuka, akan retry di run berikutnya
 *   - status = 'OK'      : semua data terbaca
 *   - status = 'PARTIAL' : sebagian data terbaca, sebagian missing
 *   - status = 'EMPTY'   : file terbuka tapi tidak ada data berguna terbaca
 *   - status = 'ERROR'   : exception terjadi
 */
function processOneFile_(item, type, dataSheet) {
  const fileId = type === 'PER' ? item.perFileId : item.pbvFileId;
  const tag = '[' + type + ' ' + item.code + ']';
  const tempIds = [];
  const issues = [];
  let filename = '';

  try {
    const file = DriveApp.getFileById(fileId);
    filename = file.getName();
    Logger.log(tag + ' processing: ' + filename);

    const ss = openAsSpreadsheetWithRetry_(file, tempIds);
    const kuartal = findSheet_(ss, SHEET_NAME_CANDIDATES.KUARTAL);
    const proyeksi = findSheet_(ss, SHEET_NAME_CANDIDATES.PROYEKSI);
    const metricDefs = type === 'PER' ? METRIC_PER : METRIC_PBV;

    let metrics = {};
    let proj = { current: null, min: null, mean: null, max: null };

    // --- Process sheet Kuartal ---
    if (!kuartal) {
      const sheetNames = ss.getSheets().map(function (s) { return s.getName(); });
      issues.push('Sheet kuartal tidak ditemukan. Sheet yang ada: [' +
                  sheetNames.join(' | ') + ']');
    } else {
      metrics = readQuarterlyMetrics_(kuartal, metricDefs);
      const totalCells = countTotalCells_(metrics);
      if (totalCells === 0) {
        const sample = sampleColumnALabels_(kuartal, 15);
        issues.push('Sheet "' + kuartal.getName() +
                    '" ditemukan, tapi 0 cell data terbaca. ' +
                    'Label kolom A: [' + sample + ']');
      } else {
        const missing = metricDefs
          .filter(function (m) { return !metrics[m.key]; })
          .map(function (m) { return m.key; });
        if (missing.length > 0) {
          const sample = sampleColumnALabels_(kuartal, 15);
          issues.push('Metrik tidak terbaca di "' + kuartal.getName() +
                      '": [' + missing.join(' | ') + ']. ' +
                      'Label kolom A: [' + sample + ']');
        }
      }
    }

    // --- Process sheet Proyeksi ---
    if (!proyeksi) {
      issues.push('Sheet "Proyeksi" tidak ditemukan');
    } else {
      proj = readProjection_(proyeksi);
      const allEmpty = (proj.current == null && proj.min == null &&
                        proj.mean == null && proj.max == null);
      if (allEmpty) {
        issues.push('Sheet "' + proyeksi.getName() +
                    '" ditemukan, tapi EPS/BVPS & Min/Mean/Max kosong');
      }
    }

    // --- Tulis ke sheet Data (parsial atau lengkap) ---
    if (type === 'PER') {
      writePerRow_(dataSheet, item.planRow, metrics, proj);
    } else {
      writePbvRow_(dataSheet, item.planRow, metrics, proj);
    }

    // --- Tentukan status ---
    let status = 'OK';
    if (issues.length > 0) {
      const totalCells = countTotalCells_(metrics);
      const hadProj = (proj.current != null || proj.min != null ||
                       proj.mean != null || proj.max != null);
      status = (totalCells === 0 && !hadProj) ? 'EMPTY' : 'PARTIAL';
      Logger.log(tag + ' [' + status + '] ' + issues.join(' | '));
    }

    return { success: true, status: status, issues: issues, filename: filename };
  } catch (e) {
    const msg = String(e.message || e);
    Logger.log('ERROR ' + tag + ': ' + msg);
    return { success: false, status: 'ERROR', issues: [msg], filename: filename };
  } finally {
    if (CONFIG.CLEAN_TEMP_FILES) {
      tempIds.forEach(function (id) {
        try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
      });
    }
  }
}

function countTotalCells_(metrics) {
  return Object.keys(metrics).reduce(function (acc, k) {
    return acc + Object.keys(metrics[k] || {}).length;
  }, 0);
}

/**
 * Ambil daftar label non-kosong dari kolom A sheet, untuk diagnostic logging.
 * Trim leading whitespace tapi pertahankan kasus aslinya supaya user bisa
 * lihat exact text yang ada di file.
 */
function sampleColumnALabels_(sheet, limit) {
  const lastRow = Math.min(sheet.getLastRow(), 50);
  if (lastRow < 1) return '';
  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  const labels = [];
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v == null || v === '') continue;
    const s = String(v).replace(/\s+/g, ' ').trim();
    if (s) labels.push(s);
    if (labels.length >= limit) break;
  }
  return labels.join(' | ') + (labels.length >= limit ? ' | ...' : '');
}

/**
 * Tulis issue ke sheet log tersembunyi "_PintarSaham_Log".
 */
function logIssue_(ss, code, type, status, issues, filename) {
  let log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    log.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Kode', 'Type',
                                          'Status', 'Filename', 'Issues']]);
    log.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#fce8b2');
    log.setFrozenRows(1);
    log.setColumnWidth(1, 150);
    log.setColumnWidth(5, 220);
    log.setColumnWidth(6, 600);
    log.hideSheet();
  }
  log.appendRow([new Date(), code, type, status,
                 filename || '', (issues || []).join(' | ')]);
}

/**
 * Tulis row PER ke sheet Data:
 *   - cols 3..86  (Pendapatan, Laba Kotor, Laba Usaha, Laba Bersih × 21Q)
 *   - cols 150..153 (EPS, Min PE, Mean PE, Max PE)
 */
function writePerRow_(dataSheet, rowIdx, metrics, proj) {
  const blocks = [
    { key: 'Pendapatan',                 start: COL.PENDAPATAN_START },
    { key: 'Laba Kotor',                 start: COL.LABA_KOTOR_START },
    { key: 'Laba Usaha',                 start: COL.LABA_USAHA_START },
    { key: 'Laba Bersih Tahun Berjalan', start: COL.LABA_BERSIH_START },
  ];
  // Range cols 3..86 (84 kolom)
  const numCols = COL.LABA_BERSIH_START + QUARTERS.length - COL.PENDAPATAN_START;
  const row = new Array(numCols).fill('');
  blocks.forEach(function (b) {
    const qmap = metrics[b.key] || {};
    for (let i = 0; i < QUARTERS.length; i++) {
      const v = qmap[QUARTERS[i]];
      if (v !== undefined && v !== null) {
        row[b.start - COL.PENDAPATAN_START + i] = v;
      }
    }
  });
  dataSheet.getRange(rowIdx, COL.PENDAPATAN_START, 1, numCols).setValues([row]);
  dataSheet.getRange(rowIdx, COL.EPS, 1, 4).setValues([[
    safe_(proj.current), safe_(proj.min), safe_(proj.mean), safe_(proj.max)
  ]]);
}

/**
 * Tulis row PBV ke sheet Data:
 *   - cols 87..149 (Total Aset, Total Liabilitas, Total Ekuitas × 21Q)
 *   - cols 154..157 (BVPS, Min PBV, Mean PBV, Max PBV)
 */
function writePbvRow_(dataSheet, rowIdx, metrics, proj) {
  const blocks = [
    { key: 'Total Aset',       start: COL.TOTAL_ASET_START },
    { key: 'Total Liabilitas', start: COL.TOTAL_LIAB_START },
    { key: 'Total Ekuitas',    start: COL.TOTAL_EKUI_START },
  ];
  // Range cols 87..149 (63 kolom)
  const numCols = COL.TOTAL_EKUI_START + QUARTERS.length - COL.TOTAL_ASET_START;
  const row = new Array(numCols).fill('');
  blocks.forEach(function (b) {
    const qmap = metrics[b.key] || {};
    for (let i = 0; i < QUARTERS.length; i++) {
      const v = qmap[QUARTERS[i]];
      if (v !== undefined && v !== null) {
        row[b.start - COL.TOTAL_ASET_START + i] = v;
      }
    }
  });
  dataSheet.getRange(rowIdx, COL.TOTAL_ASET_START, 1, numCols).setValues([row]);
  dataSheet.getRange(rowIdx, COL.BVPS, 1, 4).setValues([[
    safe_(proj.current), safe_(proj.min), safe_(proj.mean), safe_(proj.max)
  ]]);
}

/**
 * Tandai status di plan sheet (kolom PER_Done / PBV_Done).
 */
function markStatus_(planSheet, planRow, type) {
  const col = type === 'PER' ? 4 : 5;
  planSheet.getRange(planRow, col).setValue(true);
  planSheet.getRange(planRow, 6).setValue(new Date()); // Updated
}

// ============================ FILE I/O ======================================

/**
 * Buka file sebagai Spreadsheet dengan retry+backoff untuk rate limit errors.
 */
function openAsSpreadsheetWithRetry_(file, tempIds) {
  let lastErr = null;
  for (let attempt = 0; attempt < CONFIG.MAX_RETRIES; attempt++) {
    try {
      return openAsSpreadsheet_(file, tempIds);
    } catch (e) {
      lastErr = e;
      const msg = String(e.message || e);
      const isRateLimit = /rate limit|userRateLimit|quota|too many requests|503|500/i.test(msg);
      if (!isRateLimit || attempt === CONFIG.MAX_RETRIES - 1) throw e;
      const wait = Math.pow(2, attempt) * 2000 + Math.floor(Math.random() * 1000); // 2s, 4s, 8s, ...
      Logger.log('Rate limit (' + msg.substring(0, 80) + '), retry attempt ' +
                 (attempt + 1) + ' setelah ' + wait + 'ms');
      Utilities.sleep(wait);
    }
  }
  throw lastErr;
}

function openAsSpreadsheet_(file, tempIds) {
  const mime = file.getMimeType();
  if (mime === MimeType.GOOGLE_SHEETS) {
    return SpreadsheetApp.openById(file.getId());
  }
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('Advanced Drive Service belum diaktifkan. ' +
                    'Apps Script Editor > Services > + > Drive API.');
  }
  const blob = file.getBlob();
  const tmpName = '__TEMP_PintarSaham_' + Utilities.getUuid();
  let tmpId = null;

  if (typeof Drive.Files.create === 'function') {
    const created = Drive.Files.create(
      { name: tmpName, mimeType: MimeType.GOOGLE_SHEETS },
      blob
    );
    tmpId = created.id;
  } else if (typeof Drive.Files.insert === 'function') {
    const created = Drive.Files.insert(
      { title: tmpName, mimeType: MimeType.GOOGLE_SHEETS },
      blob,
      { convert: true }
    );
    tmpId = created.id;
  } else {
    throw new Error('Drive API tidak menyediakan method create() maupun insert().');
  }

  tempIds.push(tmpId);
  return SpreadsheetApp.openById(tmpId);
}

// ============================ PLAN SHEET I/O ================================

function readPlan_(planSheet) {
  const data = planSheet.getRange(2, 1, planSheet.getLastRow() - 1, 6).getValues();
  return data.map(function (row, idx) {
    return {
      code: String(row[0]),
      perFileId: String(row[1] || ''),
      pbvFileId: String(row[2] || ''),
      perDone: row[3] === true || row[3] === 'TRUE',
      pbvDone: row[4] === true || row[4] === 'TRUE',
      planRow: idx + 2,
    };
  });
}

function countRemaining_(planSheet) {
  if (!planSheet || planSheet.getLastRow() < 2) return 0;
  const data = planSheet.getRange(2, 2, planSheet.getLastRow() - 1, 4).getValues();
  let remaining = 0;
  data.forEach(function (row) {
    const perId = row[0], pbvId = row[1], perDone = row[2], pbvDone = row[3];
    if (perId && !perDone) remaining++;
    if (pbvId && !pbvDone) remaining++;
  });
  return remaining;
}

// ============================ DRIVE NAVIGATION ==============================

function locateDashboard_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('Script harus dijalankan sebagai container-bound di "Data Dashboard".');
}

function ensureDataSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.DATA_SHEET_NAME);
  return sh;
}

function locateValuasiFolders_() {
  let root = null;

  // Opsi 1: pakai ID langsung jika dikonfigurasi (paling reliable)
  if (CONFIG.ANALISA_SAHAM_FOLDER_ID) {
    try {
      root = DriveApp.getFolderById(CONFIG.ANALISA_SAHAM_FOLDER_ID);
    } catch (e) {
      throw new Error('CONFIG.ANALISA_SAHAM_FOLDER_ID tidak valid atau tidak bisa diakses: ' +
                      e.message);
    }
  } else {
    // Opsi 2: cari SEMUA folder bernama "Analisa Saham", pilih yang punya
    // subfolder Valuasi (terutama yang lengkap dengan PER & PBV).
    const candidates = [];
    const it = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
    while (it.hasNext()) candidates.push(it.next());

    if (candidates.length === 0) {
      throw new Error('Folder "' + CONFIG.ROOT_FOLDER_NAME + '" tidak ditemukan di Drive Anda. ' +
                      'Pastikan folder sudah ada dan akun yang menjalankan script punya akses.');
    }

    // Skoring tiap kandidat: 2 = punya Valuasi + PER + PBV, 1 = punya Valuasi, 0 = tidak
    let best = null, bestScore = -1;
    for (let i = 0; i < candidates.length; i++) {
      const c = candidates[i];
      const valuasi = getChildFolderByName_(c, CONFIG.VALUASI_FOLDER_NAME);
      let score = 0;
      if (valuasi) {
        score = 1;
        if (getChildFolderByName_(valuasi, CONFIG.PER_FOLDER_NAME) &&
            getChildFolderByName_(valuasi, CONFIG.PBV_FOLDER_NAME)) {
          score = 2;
        }
      }
      if (score > bestScore) { best = c; bestScore = score; }
    }

    if (bestScore < 2) {
      // Tidak ada yang punya struktur lengkap. Beri error informatif.
      const lines = candidates.map(function (c, i) {
        const valuasi = getChildFolderByName_(c, CONFIG.VALUASI_FOLDER_NAME);
        const v = valuasi ? 'Valuasi:✓' : 'Valuasi:✗';
        return '  ' + (i + 1) + ') ID=' + c.getId() + ' ' + v + ' URL=' + c.getUrl();
      }).join('\n');
      throw new Error('Ditemukan ' + candidates.length + ' folder "' + CONFIG.ROOT_FOLDER_NAME +
                      '" tetapi tidak ada yang berisi struktur lengkap (Valuasi > PER & PBV).\n' +
                      'Kandidat:\n' + lines + '\n\n' +
                      'Solusi: jalankan menu "🔧 Diagnose Folder" untuk detail, lalu ' +
                      'set CONFIG.ANALISA_SAHAM_FOLDER_ID ke ID folder yang benar.');
    }

    root = best;
    Logger.log('Menggunakan folder root: ' + root.getName() + ' (ID: ' + root.getId() +
               ', dipilih dari ' + candidates.length + ' kandidat)');
  }

  const valuasi = getChildFolderByName_(root, CONFIG.VALUASI_FOLDER_NAME);
  if (!valuasi) {
    throw new Error('Subfolder "' + CONFIG.VALUASI_FOLDER_NAME + '" tidak ada di "' +
                    root.getName() + '" (ID: ' + root.getId() + ').');
  }
  const perFolder = getChildFolderByName_(valuasi, CONFIG.PER_FOLDER_NAME);
  const pbvFolder = getChildFolderByName_(valuasi, CONFIG.PBV_FOLDER_NAME);
  if (!perFolder || !pbvFolder) {
    // List subfolder yang ada untuk diagnosa typo / spasi
    const subs = [];
    const sit = valuasi.getFolders();
    while (sit.hasNext()) subs.push('"' + sit.next().getName() + '"');
    const missing = [];
    if (!perFolder) missing.push('"' + CONFIG.PER_FOLDER_NAME + '"');
    if (!pbvFolder) missing.push('"' + CONFIG.PBV_FOLDER_NAME + '"');
    throw new Error('Subfolder ' + missing.join(' dan ') + ' tidak ada di "' +
                    valuasi.getName() + '".\nSubfolder yang ada:\n  ' + subs.join('\n  '));
  }
  return { perFolder: perFolder, pbvFolder: pbvFolder };
}

/**
 * Diagnostic: walk down folder tree dan tampilkan isi tiap level.
 * Membantu identifikasi typo, multiple folders, atau struktur yang berbeda.
 */
function diagnoseFolder() {
  const ui = SpreadsheetApp.getUi();
  const rootCandidates = [];
  const it = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  while (it.hasNext()) rootCandidates.push(it.next());

  if (rootCandidates.length === 0) {
    ui.alert('🔧 Diagnose',
             'Tidak ada folder bernama "' + CONFIG.ROOT_FOLDER_NAME + '" di Drive Anda.',
             ui.ButtonSet.OK);
    return;
  }

  let msg = 'Ditemukan ' + rootCandidates.length + ' folder "' +
            CONFIG.ROOT_FOLDER_NAME + '":\n\n';
  let bestRoot = null, bestScore = -1, bestValuasi = null;

  rootCandidates.forEach(function (root, ri) {
    msg += '━━━ #' + (ri + 1) + ' "' + root.getName() + '" ━━━\n';
    msg += 'ID: ' + root.getId() + '\n';

    // Cari semua subfolder "Valuasi" di root (mungkin ada > 1!)
    const valuasiList = [];
    const vIt = root.getFoldersByName(CONFIG.VALUASI_FOLDER_NAME);
    while (vIt.hasNext()) valuasiList.push(vIt.next());

    if (valuasiList.length === 0) {
      msg += '✗ Tidak ada subfolder "' + CONFIG.VALUASI_FOLDER_NAME + '"\n';
      msg += 'Isi "' + root.getName() + '":\n';
      msg += listSubfolderNames_(root, 30) + '\n';
    } else {
      if (valuasiList.length > 1) {
        msg += '⚠ Ada ' + valuasiList.length + ' folder bernama "Valuasi"!\n';
      }
      valuasiList.forEach(function (v, vi) {
        const label = valuasiList.length > 1 ? ' #' + (vi + 1) : '';
        msg += '\nValuasi' + label + ' (ID: ' + v.getId() + ')\n';

        const per = getChildFolderByName_(v, CONFIG.PER_FOLDER_NAME);
        const pbv = getChildFolderByName_(v, CONFIG.PBV_FOLDER_NAME);
        msg += '  PER "' + CONFIG.PER_FOLDER_NAME + '": ' + (per ? '✓' : '✗') + '\n';
        msg += '  PBV "' + CONFIG.PBV_FOLDER_NAME + '": ' + (pbv ? '✓' : '✗') + '\n';

        // Selalu tampilkan isi Valuasi supaya user bisa lihat nama persisnya
        msg += '  Isi folder:\n' + indent_(listSubfolderNames_(v, 30), '    ');

        const score = (per && pbv) ? 2 : (per || pbv) ? 1 : 0;
        if (score > bestScore) {
          bestScore = score;
          bestRoot = root;
          bestValuasi = v;
        }
      });
    }
    msg += '\n';
  });

  if (bestScore === 2 && bestRoot) {
    msg += '━━━ REKOMENDASI ━━━\n';
    msg += 'Folder lengkap ditemukan. Untuk mengunci pilihan, set di CONFIG:\n\n';
    msg += "ANALISA_SAHAM_FOLDER_ID: '" + bestRoot.getId() + "',\n";
  } else {
    msg += '━━━ INFO ━━━\n';
    msg += 'Belum ada folder dengan struktur lengkap (Valuasi + PER + PBV).\n';
    msg += 'Bandingkan daftar di atas dengan nama yang diharapkan:\n';
    msg += '  - "' + CONFIG.PER_FOLDER_NAME + '"\n';
    msg += '  - "' + CONFIG.PBV_FOLDER_NAME + '"\n';
    msg += 'Jika nama berbeda (misal beda spasi/case), sesuaikan CONFIG ' +
           'atau rename foldernya di Drive.\n';
  }

  Logger.log(msg);
  // Alert maks ~2000 char, log selalu lengkap di Executions
  const shown = msg.length > 1800 ? (msg.substring(0, 1800) +
                                      '\n... (terpotong, lihat lengkap di Executions Log)') : msg;
  ui.alert('🔧 Diagnose Folder', shown, ui.ButtonSet.OK);
}

function listSubfolderNames_(folder, limit) {
  const subs = folder.getFolders();
  const lines = [];
  let count = 0;
  while (subs.hasNext()) {
    if (count >= limit) {
      lines.push('  - ... (terpotong di ' + limit + ' folder)');
      break;
    }
    lines.push('  - "' + subs.next().getName() + '"');
    count++;
  }
  if (lines.length === 0) lines.push('  (folder kosong / tidak ada subfolder)');
  return lines.join('\n');
}

function indent_(text, prefix) {
  return text.split('\n').map(function (l) { return prefix + l; }).join('\n') + '\n';
}

function getChildFolderByName_(parent, name) {
  // Exact match terlebih dahulu (paling cepat)
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  // Fallback: trimmed + case-insensitive (untuk menangani spasi tersembunyi/case)
  const target = String(name).trim().toLowerCase();
  const all = parent.getFolders();
  while (all.hasNext()) {
    const f = all.next();
    if (String(f.getName()).trim().toLowerCase() === target) return f;
  }
  return null;
}

/**
 * Baca daftar kode emiten whitelist dari sheet (kolom A, mulai baris 2).
 * Return: Set kode (uppercase) atau null jika whitelist tidak diaktifkan
 * atau sheet tidak ditemukan / kosong.
 */
function readWhitelist_(ss) {
  if (!CONFIG.WHITELIST_SHEET_NAME) return null;  // fitur dimatikan
  const sheet = ss.getSheetByName(CONFIG.WHITELIST_SHEET_NAME);
  if (!sheet) {
    Logger.log('⚠ Sheet whitelist "' + CONFIG.WHITELIST_SHEET_NAME +
               '" tidak ditemukan. Memproses SEMUA file.');
    return null;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.WHITELIST_START_ROW) {
    Logger.log('⚠ Sheet whitelist kosong (tidak ada data dari baris ' +
               CONFIG.WHITELIST_START_ROW + '). Memproses SEMUA file.');
    return null;
  }
  const numRows = lastRow - CONFIG.WHITELIST_START_ROW + 1;
  const values = sheet.getRange(CONFIG.WHITELIST_START_ROW, CONFIG.WHITELIST_COLUMN,
                                numRows, 1).getValues();
  const set = new Set();
  values.forEach(function (row) {
    const raw = row[0];
    if (raw == null || raw === '') return;
    // Normalisasi: trim, uppercase, ambil token pertama (kalau ada "BBCA - Bank ...")
    let code = String(raw).trim().toUpperCase();
    // Buang teks setelah spasi/tanda hubung (kalau user tulis "BBCA Bank Central Asia")
    const m = code.match(/^([A-Z0-9]{3,5})\b/);
    if (m) code = m[1];
    if (code) set.add(code);
  });
  if (set.size === 0) {
    Logger.log('⚠ Sheet whitelist tidak berisi kode valid. Memproses SEMUA file.');
    return null;
  }
  return set;
}

function collectValuationFiles_(folder, type) {
  const out = [];
  const seenIds = {};
  const stack = [folder];
  while (stack.length > 0) {
    const f = stack.pop();
    const files = f.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fid = file.getId();
      if (seenIds[fid]) continue;
      seenIds[fid] = true;
      const code = extractCode_(file.getName(), type);
      if (code) out.push({ code: code, file: file });
    }
    const subs = f.getFolders();
    while (subs.hasNext()) stack.push(subs.next());
  }
  return out;
}

function extractCode_(filename, type) {
  let base = filename.replace(/\.(xlsx|xls|gsheet)$/i, '').trim();
  if (/petunjuk|template|contoh/i.test(base)) return null;
  const re = new RegExp('^([A-Z0-9]{3,5})[\\s_\\-\\.]+' + type + '\\b', 'i');
  const m = base.match(re);
  return m ? m[1].toUpperCase() : null;
}

// ============================ DATA EXTRACTION ===============================

function findSheet_(ss, candidates) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (candidates.indexOf(sheets[i].getName()) >= 0) return sheets[i];
  }
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName().trim();
    for (let j = 0; j < candidates.length; j++) {
      if (n === candidates[j].trim()) return sheets[i];
    }
  }
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName().toLowerCase().trim();
    for (let j = 0; j < candidates.length; j++) {
      if (n.indexOf(candidates[j].toLowerCase().trim()) === 0) return sheets[i];
    }
  }
  return null;
}

/**
 * Normalisasi label cell agar perbandingan robust terhadap:
 *   - case (uppercase/lowercase)
 *   - leading/trailing whitespace
 *   - tabs / multi-space (mis: "      Aset Lancar" -> "aset lancar")
 *   - non-breaking space (U+00A0)
 */
function normalizeLabel_(s) {
  return String(s == null ? '' : s)
    .toLowerCase()
    .replace(/\s+/g, ' ')  // collapse semua whitespace (tab, nbsp, multi-space)
    .trim();
}

function readQuarterlyMetrics_(sheet, metricDefs) {
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return {};
  const header = data[0];
  const quarterCol = {};
  for (let c = 1; c < header.length; c++) {
    const v = String(header[c] == null ? '' : header[c]).trim();
    if (/^Q[1-4]\s*\d{4}$/.test(v)) {
      const norm = v.replace(/\s+/g, ' ');
      quarterCol[norm] = c;
    }
  }
  const result = {};
  for (let mi = 0; mi < metricDefs.length; mi++) {
    const md = metricDefs[mi];
    let foundRow = -1;

    // Exact normalized match SAJA. Tidak ada startsWith fallback supaya
    // tidak salah ambil "Total Aset Lancar" sebagai "Total Aset",
    // atau "Pendapatan Bunga" sebagai "Pendapatan".
    for (let r = 1; r < data.length; r++) {
      const label = normalizeLabel_(data[r][0]);
      if (!label) continue;
      if (md.labels.indexOf(label) >= 0) { foundRow = r; break; }
    }

    if (foundRow >= 0) {
      const row = data[foundRow];
      result[md.key] = {};
      for (let qi = 0; qi < QUARTERS.length; qi++) {
        const q = QUARTERS[qi];
        if (q in quarterCol) {
          const v = row[quarterCol[q]];
          if (v === '-' || v === '' || v === null || v === undefined) continue;
          if (typeof v === 'number') {
            result[md.key][q] = v;
          } else {
            const n = parseFloat(String(v).replace(/,/g, ''));
            if (!isNaN(n)) result[md.key][q] = n;
          }
        }
      }
    }
  }
  return result;
}

/**
 * Baca data dari sheet Proyeksi.
 *
 * Strukturnya konsisten di semua file: ada cell berlabel "Stock" yang dipakai
 * sebagai anchor, dan Current EPS/BVPS selalu 2 kolom ke kanan dari "Stock"
 * (1 baris di bawah). Min/Mean/Max headers ada di baris terpisah dengan 3
 * cell berturutan, values langsung di baris bawahnya.
 *
 * Posisi absolut bisa BERGESER antar file (mis: SIDO PBV pakai col D bukan C).
 * Jadi kita cari "Stock" anywhere dalam area atas, lalu pakai posisinya
 * sebagai anchor; demikian pula Min/Mean/Max dicari sebagai 3 cell berturutan
 * di mana saja, bukan pada kolom tetap.
 */
function readProjection_(sheet) {
  const data = sheet.getDataRange().getValues();
  let current = null, min = null, mean = null, max = null;
  if (data.length === 0) return { current: null, min: null, mean: null, max: null };

  // ---- 1) Cari "Stock" di area atas, lalu Current EPS/BVPS = 2 kolom ke kanan, 1 baris di bawah ----
  let stockRow = -1, stockCol = -1;
  const stockMaxRows = Math.min(data.length, 25);
  const stockMaxCols = 12;
  for (let r = 0; r < stockMaxRows; r++) {
    const row = data[r] || [];
    const ncol = Math.min(row.length, stockMaxCols);
    for (let c = 0; c < ncol; c++) {
      const v = row[c];
      if (v != null && String(v).trim().toLowerCase() === 'stock') {
        stockRow = r;
        stockCol = c;
        break;
      }
    }
    if (stockRow >= 0) break;
  }
  if (stockRow >= 0 && stockRow + 1 < data.length) {
    const valueRow = data[stockRow + 1] || [];
    const targetCol = stockCol + 2;
    if (targetCol < valueRow.length) {
      current = valueRow[targetCol];
    }
  }

  // ---- 2) Cari header Min|Mean|Max sebagai 3 cell berturutan, nilai di baris bawah ----
  const mmmMaxRows = Math.min(data.length, 30);
  const mmmMaxCols = 15;
  for (let r = 0; r < mmmMaxRows; r++) {
    const row = data[r] || [];
    const ncol = Math.min(row.length, mmmMaxCols);
    for (let c = 0; c + 2 < ncol; c++) {
      const a = String(row[c]     == null ? '' : row[c]    ).trim().toLowerCase();
      const b = String(row[c + 1] == null ? '' : row[c + 1]).trim().toLowerCase();
      const d = String(row[c + 2] == null ? '' : row[c + 2]).trim().toLowerCase();
      if (a === 'min' && b === 'mean' && d === 'max') {
        if (r + 1 < data.length) {
          const valRow = data[r + 1] || [];
          min  = valRow[c];
          mean = valRow[c + 1];
          max  = valRow[c + 2];
        }
        return { current: current, min: min, mean: mean, max: max };
      }
    }
  }

  return { current: current, min: min, mean: mean, max: max };
}

function safe_(v) {
  return (v === undefined || v === null) ? '' : v;
}