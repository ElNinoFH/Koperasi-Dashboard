/**
 * ============================================================
 *  KOPERASI DASHBOARD - Web App (Google Apps Script)
 *  Code.gs : Router, konfigurasi global, dan helper utama
 * ============================================================
 *
 * Arsitektur Hybrid:
 *  - Spreadsheet baru terstruktur dijadikan DATABASE
 *  - Laporan format lama di-generate otomatis (read-only) untuk arsip/cetak
 *
 * Cara pakai pertama kali:
 *  1. Jalankan fungsi `setupDatabase()` sekali dari editor Apps Script.
 *  2. Deploy sebagai Web App (Deploy > New deployment > Web app).
 */

// ====================== KONFIGURASI GLOBAL ======================

/**
 * Nama-nama sheet yang menyusun database.
 */
const SHEETS = {
  MASTER: 'MasterBahan',
  STOK: 'StokHarian',
  TRANSAKSI: 'Transaksi',
  PEMBAYARAN: 'Pembayaran',
  SESI_TIM: 'SesiTim',
  CONFIG: 'Config'
};

/**
 * Daftar fase/kelompok penerima distribusi (urutan = urutan tampil di UI).
 */
const FASE_LIST = [
  'Fase PS',
  'Fase A',
  'Fase B',
  'Fase C',
  'SFS',
  'Makan Malam',
  'Sarapan PJ',
  'Guru',
  'IPAK'
];

/**
 * Kategori bahan beserta prefix kode-nya.
 */
const KATEGORI = {
  'Karbohidrat': 'K',
  'Protein': 'P',
  'Sayuran': 'S',
  'Buah': 'B',
  'Bumbu Dapur': 'BD'
};

const TIMEZONE = 'Asia/Jakarta';
const APP_TITLE = 'Koperasi Dashboard';

// ====================== ROUTER WEB APP ======================

/**
 * Entry point web app.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appTitle = APP_TITLE;
  return template
    .evaluate()
    .setTitle(APP_TITLE)
    .setFaviconUrl('https://ssl.gstatic.com/docs/script/images/favicon.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Menggabungkan file HTML (untuk memisahkan CSS & JS dari Index).
 * Dipakai di Index.html via <?!= include('Styles') ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================== HELPER SPREADSHEET ======================

/**
 * Mengembalikan Spreadsheet yang dipakai sebagai database.
 *
 * Urutan prioritas:
 *  1. Spreadsheet aktif  — kalau script di-bind ke Sheets (container-bound)
 *  2. ID tersimpan       — kalau sudah pernah dibuat sebelumnya (standalone)
 *  3. Buat baru          — pertama kali dijalankan sebagai standalone
 */
function getSpreadsheet() {
  // 1. Container-bound: ada spreadsheet aktif
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  // 2. Standalone: cek apakah ID sudah pernah disimpan
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('DB_SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      // ID tidak valid — lanjut buat baru
    }
  }

  // 3. Buat spreadsheet baru dan simpan ID-nya untuk dipakai selamanya
  const ss = SpreadsheetApp.create('Koperasi Dashboard — Database');
  props.setProperty('DB_SPREADSHEET_ID', ss.getId());

  // Pindahkan ke folder yang sama dengan script (opsional, biar rapi)
  try {
    const scriptFile = DriveApp.getFileById(ScriptApp.getScriptId());
    const parents = scriptFile.getParents();
    if (parents.hasNext()) {
      const folder = parents.next();
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    }
  } catch (e) { /* tidak kritis */ }

  Logger.log('Spreadsheet baru dibuat: ' + ss.getUrl());
  return ss;
}

/**
 * Mengambil sheet berdasarkan nama. Lempar error bila belum di-setup.
 */
function getSheet(name) {
  const sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" belum ada. Jalankan setupDatabase() terlebih dahulu.');
  }
  return sheet;
}

/**
 * Membaca seluruh isi sheet sebagai array of object (key = header baris pertama).
 */
function readSheetAsObjects(name) {
  const sheet = getSheet(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const obj = {};
    let isEmpty = true;
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = values[i][j];
      if (values[i][j] !== '' && values[i][j] !== null) isEmpty = false;
    }
    if (!isEmpty) {
      obj._row = i + 1; // nomor baris asli di sheet (1-based)
      rows.push(obj);
    }
  }
  return rows;
}

// ====================== HELPER UMUM ======================

/**
 * Format tanggal menjadi string "yyyy-MM-dd" sesuai timezone.
 */
function formatDate(date) {
  if (!date) date = new Date();
  if (typeof date === 'string') date = new Date(date);
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd');
}

/**
 * Format tanggal menjadi label manusiawi: "Sabtu, 27 Juni 2026".
 */
function formatDateLabel(date) {
  if (!date) date = new Date();
  if (typeof date === 'string') date = new Date(date + 'T00:00:00');
  const hari = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return hari[date.getDay()] + ', ' + date.getDate() + ' ' +
    bulan[date.getMonth()] + ' ' + date.getFullYear();
}

/**
 * Membulatkan angka uang ke ribuan terdekat (mengikuti perilaku spreadsheet lama).
 */
function bulatkanRibuan(nilai) {
  return Math.round(nilai / 1000) * 1000;
}

/**
 * Membuat ID unik berbasis timestamp.
 */
function generateId(prefix) {
  return (prefix || 'ID') + '-' + new Date().getTime() + '-' +
    Math.floor(Math.random() * 1000);
}

/**
 * Mengembalikan tanggal hari ini (string yyyy-MM-dd) untuk konsumsi front-end.
 */
function getToday() {
  return {
    date: formatDate(new Date()),
    label: formatDateLabel(new Date())
  };
}
