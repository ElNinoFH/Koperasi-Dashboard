/**
 * ============================================================
 *  Setup.gs : Inisialisasi database (sheet) + seed MasterBahan
 * ============================================================
 *
 *  >>> JALANKAN setupDatabase() SEKALI dari editor Apps Script <<<
 *
 *  Membuat 6 sheet:
 *   - MasterBahan  : data master bahan (sumber kebenaran tunggal)
 *   - StokHarian   : ledger stok per tanggal (carry-over otomatis)
 *   - Transaksi    : distribusi per bahan per fase (format TALL)
 *   - Pembayaran   : status bayar per fase per tanggal
 *   - SesiTim      : siapa operator & tim koperasi per tanggal
 *   - Config       : penyimpanan key-value (mis. daftar anggota tetap)
 */

// ====================== HEADER TIAP SHEET ======================

const HEADERS = {
  MasterBahan: ['KodeBarang', 'NamaBahan', 'Kategori', 'Satuan', 'HargaBeli', 'HargaJual', 'StokMin', 'Aktif'],
  StokHarian: ['Tanggal', 'KodeBarang', 'NamaBahan', 'Kategori', 'Satuan', 'StokAwal', 'Masuk', 'Keluar', 'StokAkhir', 'Operator', 'UpdatedAt'],
  Transaksi: ['ID', 'Tanggal', 'KodeBarang', 'NamaBahan', 'Kategori', 'Satuan', 'Fase', 'Jumlah', 'HargaSatuan', 'TotalHarga', 'Operator', 'CreatedAt'],
  Pembayaran: ['Tanggal', 'Fase', 'TotalBelanja', 'Dibayar', 'Sisa', 'Status', 'UpdatedAt'],
  SesiTim: ['Tanggal', 'Operator', 'AnggotaTim', 'CreatedAt', 'UpdatedAt'],
  Config: ['Key', 'Value']
};

// ====================== FUNGSI UTAMA SETUP ======================

/**
 * Inisialisasi seluruh database. Aman dijalankan berulang
 * (tidak menimpa sheet yang sudah punya data master).
 */
function setupDatabase() {
  const ss = getSpreadsheet();

  Object.keys(HEADERS).forEach(function (sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    // Tulis header bila baris pertama kosong
    const firstCell = sheet.getRange(1, 1).getValue();
    if (firstCell === '' || firstCell === null) {
      const headers = HEADERS[sheetName];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#7C3AED')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  });

  // Seed master bahan hanya bila masih kosong
  const masterSheet = ss.getSheetByName(SHEETS.MASTER);
  if (masterSheet.getLastRow() < 2) {
    seedMasterBahan();
  }

  // Hapus sheet default "Sheet1" bila kosong
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(def);
  }

  // Simpan default config anggota tetap (kosong, bisa diisi user)
  setConfig('ANGGOTA_TETAP', getConfig('ANGGOTA_TETAP') || '');

  SpreadsheetApp.flush();
  return 'Setup selesai. ' + getMasterBahanRaw().length + ' bahan siap digunakan.';
}

/**
 * Mengisi MasterBahan dengan data dari spreadsheet lama.
 */
function seedMasterBahan() {
  const sheet = getSheet(SHEETS.MASTER);
  const data = SEED_BAHAN.map(function (b) {
    return [b[0], b[1], b[2], b[3], b[4], b[4], b[5] || 0, true];
    //      kode, nama, kategori, satuan, hargaBeli, hargaJual(=beli), stokMin, aktif
  });
  sheet.getRange(2, 1, data.length, HEADERS.MasterBahan.length).setValues(data);
  return data.length;
}

// ====================== CONFIG KEY-VALUE ======================

function setConfig(key, value) {
  const sheet = getSheet(SHEETS.CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function getConfig(key) {
  const sheet = getSheet(SHEETS.CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return '';
}

// ====================== DATA SEED (dari spreadsheet lama) ======================
// Format: [Kode, Nama, Kategori, Satuan, HargaBeli, StokMin]

const SEED_BAHAN = [
  // ---------- KARBOHIDRAT ----------
  ['K01', 'Beras', 'Karbohidrat', 'liter', 14500, 5],
  ['K02', 'Kentang', 'Karbohidrat', 'kg', 20000, 0],
  ['K03', 'Tepung sagu', 'Karbohidrat', 'bks', 7000, 0],
  ['K04', 'Tepung tapioka', 'Karbohidrat', 'Bks', 38000, 0],
  ['K05', 'Tepung ketan', 'Karbohidrat', 'sachet', 21000, 0],
  ['K06', 'Bihun', 'Karbohidrat', 'bungkus', 11000, 0],
  ['K07', 'Ubi', 'Karbohidrat', 'Kg', 15000, 0],
  ['K08', 'Ubi madu', 'Karbohidrat', 'kg', 12000, 0],
  ['K09', 'Ketan Hitam', 'Karbohidrat', 'kg', 42000, 0],
  ['K10', 'Tepung beras', 'Karbohidrat', 'bks', 10000, 0],
  ['K11', 'Beras merah', 'Karbohidrat', 'kg', 19000, 0],
  ['K12', 'Singkong', 'Karbohidrat', 'kg', 10000, 0],
  ['K13', 'Talas', 'Karbohidrat', 'kg', 17000, 0],

  // ---------- PROTEIN ----------
  ['P01', 'Ayam Potong 10', 'Protein', 'ekor', 52000, 2],
  ['P02', 'Ayam potong 12', 'Protein', 'ekor', 57000, 2],
  ['P03', 'Telur puyuh', 'Protein', 'buah', 0, 0],
  ['P04', 'Tahu Cina', 'Protein', 'kotak', 7000, 0],
  ['P05', 'Tempe', 'Protein', 'papan', 7000, 0],
  ['P06', 'Daging sapi', 'Protein', 'kg', 157000, 0],
  ['P07', 'Daging sapi teriaki', 'Protein', 'kg', 127000, 0],
  ['P08', 'Kacang tanah kupas', 'Protein', 'kg', 37000, 0],
  ['P09', 'Ikan kembung', 'Protein', 'ekor', 7000, 0],
  ['P10', 'biji wijen', 'Protein', 'kg', 63000, 0],
  ['P11', 'Tofu', 'Protein', 'kg', 10000, 0],
  ['P12', 'keju', 'Protein', 'pcs', 17000, 0],
  ['P13', 'Telur', 'Protein', 'butir', 2000, 0],
  ['P14', 'ikan dori', 'Protein', 'kg', 61000, 0],
  ['P15', 'tahu kuning', 'Protein', 'buah', 1500, 0],
  ['P16', 'ikan nila', 'Protein', 'kg', 42000, 0],
  ['P17', 'tahu putih', 'Protein', 'pack', 6000, 0],
  ['P18', 'edamame', 'Protein', 'kg', 38000, 0],
  ['P19', 'ayam filet', 'Protein', 'kg', 55000, 0],
  ['P20', 'daging rawon', 'Protein', 'kg', 112000, 0],
  ['P21', 'tetelan', 'Protein', 'kg', 87000, 0],
  ['P22', 'daging has dalam', 'Protein', 'kg', 137000, 0],
  ['P23', 'Ikan Patin', 'Protein', 'kg', 42000, 0],
  ['P24', 'Udang Kupas', 'Protein', 'kg', 97000, 0],
  ['P25', 'Ikan Tuna', 'Protein', 'kg', 47000, 0],
  ['P26', 'kedelai', 'Protein', 'Kg', 18000, 0],
  ['P27', 'Ayam potong', 'Protein', 'Kg', 6000, 0],

  // ---------- SAYURAN ----------
  ['S01', 'Bayam', 'Sayuran', 'ikat', 7000, 0],
  ['S02', 'Wortel', 'Sayuran', 'kg', 25000, 0],
  ['S03', 'Kembang kol', 'Sayuran', 'kg', 32000, 0],
  ['S04', 'Buncis', 'Sayuran', 'kg', 22000, 0],
  ['S05', 'Kacang merah', 'Sayuran', 'kg', 22000, 0],
  ['S06', 'Daun bawang + seledri', 'Sayuran', 'ikat', 5000, 0],
  ['S07', 'Oyong', 'Sayuran', 'buah', 3000, 0],
  ['S08', 'Toge kecil', 'Sayuran', 'kg', 22000, 0],
  ['S09', 'Pokcoy', 'Sayuran', 'kg', 20000, 0],
  ['S10', 'Jamur kuping', 'Sayuran', 'kg', 26000, 0],
  ['S11', 'Timun', 'Sayuran', 'kg', 17000, 0],
  ['S12', 'Biji jagung', 'Sayuran', 'buah', 0, 0],
  ['S13', 'Kol', 'Sayuran', 'KG', 16000, 0],
  ['S14', 'Daun pandan', 'Sayuran', '5 lembar', 3000, 0],
  ['S15', 'Tomat', 'Sayuran', 'kg', 24000, 0],
  ['S16', 'Melinjo', 'Sayuran', 'kg', 30000, 0],
  ['S17', 'Kecambah', 'Sayuran', 'kg', 24000, 0],
  ['S18', 'Sayur Asem', 'Sayuran', 'pcs', 6000, 0],
  ['S19', 'Kangkung', 'Sayuran', 'iket', 6000, 0],
  ['S20', 'Selada', 'Sayuran', 'kg', 50000, 0],
  ['S21', 'Sawi putih', 'Sayuran', 'kg', 20000, 0],
  ['S22', 'Cesim', 'Sayuran', 'kg', 13000, 0],
  ['S23', 'baby corn', 'Sayuran', 'kg', 22000, 0],
  ['S24', 'sawi hijau', 'Sayuran', 'kg', 17000, 0],
  ['S25', 'kemangi', 'Sayuran', 'iket', 2000, 0],
  ['S26', 'kacang panjang', 'Sayuran', 'kg', 20000, 0],
  ['S27', 'jagung', 'Sayuran', 'buah', 7000, 0],
  ['S28', 'labu siam', 'Sayuran', 'kg', 17000, 0],
  ['S29', 'terong', 'Sayuran', 'buah', 3000, 0],
  ['S30', 'Sawi Hijau', 'Sayuran', 'kg', 18000, 0],
  ['S31', 'Labu kuning', 'Sayuran', 'kg', 15000, 0],
  ['S32', 'Toge', 'Sayuran', 'kg', 20000, 0],
  ['S33', 'Daun pisang', 'Sayuran', 'gulung', 4000, 0],

  // ---------- BUAH ----------
  ['B01', 'Pepaya', 'Buah', 'kg', 18000, 0],
  ['B02', 'Melon', 'Buah', 'kg', 28000, 0],
  ['B03', 'Semangka', 'Buah', 'kg', 20000, 0],
  ['B04', 'Pisang lampung', 'Buah', 'sisir', 11000, 0],
  ['B05', 'Kismis', 'Buah', 'bks', 0, 0],
  ['B06', 'Jambu kristal', 'Buah', 'kg', 28000, 0],
  ['B07', 'Nanas madu', 'Buah', 'buah', 5000, 0],
  ['B08', 'Lemon', 'Buah', 'kg', 33000, 0],
  ['B09', 'Jeruk nipis', 'Buah', 'kg', 1700, 0],
  ['B10', 'Kelengkeng', 'Buah', 'kg', 52000, 0],
  ['B11', 'Jambu Biji', 'Buah', 'kg', 25000, 0],
  ['B12', 'nangka kupas', 'Buah', 'kg', 42000, 0],
  ['B13', 'kelapa parut', 'Buah', 'buah', 18000, 0],
  ['B14', 'Apel', 'Buah', 'kg', 42000, 0],
  ['B15', 'Jeruk', 'Buah', 'kg', 25000, 0],
  ['B16', 'pir', 'Buah', 'KG', 28000, 0],
  ['B17', 'Mangga', 'Buah', 'kg', 38000, 0],
  ['B18', 'Jeruk Medan', 'Buah', 'kg', 29000, 0],
  ['B19', 'Jambu Air', 'Buah', 'kg', 26000, 0],
  ['B20', 'Bengkoang', 'Buah', 'Kg', 20000, 0],
  ['B21', 'pisang ambon', 'Buah', 'sisir', 32000, 0],
  ['B22', 'Alpukat', 'Buah', 'kg', 28000, 0],
  ['B23', 'stroberi', 'Buah', 'pack', 12000, 0],
  ['B24', 'pisang uli', 'Buah', 'sisir', 12000, 0],
  ['B25', 'Kurma', 'Buah', 'Kg', 24000, 0],
  ['B26', 'jeruk limau', 'Buah', 'buah', 1200, 0],
  ['B27', 'belimbing', 'Buah', 'gr', 18000, 0],
  ['B28', 'Nanas', 'Buah', 'buah', 12000, 0],
  ['B29', 'buah naga', 'Buah', 'buah', 30000, 0],
  ['B30', 'Salak', 'Buah', 'Buah', 1500, 0],

  // ---------- BUMBU DAPUR ----------
  ['BD01', 'Bawang merah', 'Bumbu Dapur', 'kg', 75000, 0],
  ['BD02', 'Bawang putih', 'Bumbu Dapur', 'kg', 55000, 0],
  ['BD03', 'Bawang Bombay', 'Bumbu Dapur', 'buah', 7000, 0],
  ['BD04', 'Cabe ijo besar', 'Bumbu Dapur', 'kg', 55000, 0],
  ['BD05', 'Cabe keriting', 'Bumbu Dapur', 'kg', 80000, 0],
  ['BD06', 'Cabe Rawit', 'Bumbu Dapur', 'kg', 85000, 0],
  ['BD07', 'Cabe merah besar', 'Bumbu Dapur', 'kg', 85000, 0],
  ['BD08', 'Cabe ijo kriting', 'Bumbu Dapur', 'buah', 61000, 0],
  ['BD09', 'Tomat ijo', 'Bumbu Dapur', 'buah', 700, 0],
  ['BD10', 'Temu Kunci', 'Bumbu Dapur', 'bonggol', 1000, 0],
  ['BD11', 'Kunyit', 'Bumbu Dapur', 'kg', 20000, 0],
  ['BD12', 'Biji Pala', 'Bumbu Dapur', 'buah', 3000, 0],
  ['BD13', 'Garam reina', 'Bumbu Dapur', 'kantong', 6000, 0],
  ['BD14', 'Gula Pasir', 'Bumbu Dapur', 'kg', 20000, 0],
  ['BD15', 'Ketumbar', 'Bumbu Dapur', 'kg', 51000, 0],
  ['BD16', 'Lengkuas', 'Bumbu Dapur', 'kg', 22000, 0],
  ['BD17', 'Kencur', 'Bumbu Dapur', 'kg', 51000, 0],
  ['BD18', 'chiaseed', 'Bumbu Dapur', 'KG', 100000, 0],
  ['BD19', 'Kemiri', 'Bumbu Dapur', 'kg', 80000, 0],
  ['BD20', 'Jahe', 'Bumbu Dapur', 'kg', 32000, 0],
  ['BD21', 'Terasi', 'Bumbu Dapur', 'sachet', 500, 0],
  ['BD22', 'Daun salam/jeruk', 'Bumbu Dapur', 'per 5 lembar', 1000, 0],
  ['BD23', 'Lada bubuk', 'Bumbu Dapur', 'bks', 4000, 0],
  ['BD24', 'Sereh', 'Bumbu Dapur', 'per batang', 2000, 0],
  ['BD25', 'Sarung tangan', 'Bumbu Dapur', 'kotak', 10000, 0],
  ['BD26', 'Gula merah', 'Bumbu Dapur', 'kg', 22000, 0],
  ['BD27', 'Bumbu Dapur', 'Bumbu Dapur', 'bks', 12000, 0],
  ['BD28', 'Asam kandis', 'Bumbu Dapur', 'gr', 12000, 0],
  ['BD29', 'Agar Swallow', 'Bumbu Dapur', 'bks', 5500, 0],
  ['BD30', 'Nutrijel', 'Bumbu Dapur', 'bks', 7000, 0],
  ['BD31', 'Minyak', 'Bumbu Dapur', 'liter', 24000, 0],
  ['BD32', 'Kecap asin', 'Bumbu Dapur', 'botol', 15000, 0],
  ['BD33', 'Santan', 'Bumbu Dapur', 'bks', 9000, 0],
  ['BD34', 'Garam masak kecil', 'Bumbu Dapur', 'kg', 3500, 0],
  ['BD35', 'Cengkeh', 'Bumbu Dapur', 'pcs', 2000, 0],
  ['BD36', 'Kapulaga', 'Bumbu Dapur', 'bks', 1500, 0],
  ['BD37', 'Kayu secang', 'Bumbu Dapur', 'kg', 90000, 0],
  ['BD38', 'Kayu manis', 'Bumbu Dapur', 'kg', 167000, 0],
  ['BD39', 'bunga lawang', 'Bumbu Dapur', 'bungkus', 6000, 0],
  ['BD40', 'Teh cap botol', 'Bumbu Dapur', 'pcs', 500, 0],
  ['BD41', 'Garam masak besar', 'Bumbu Dapur', 'pcs', 6000, 0],
  ['BD42', 'Citric acid', 'Bumbu Dapur', 'pcs', 0, 0],
  ['BD43', 'mutiara', 'Bumbu Dapur', 'kg', 0, 0],
  ['BD44', 'kerupuk udang', 'Bumbu Dapur', 'kg', 0, 0],
  ['BD45', 'gula halus', 'Bumbu Dapur', 'bks', 0, 0],
  ['BD46', 'vanili', 'Bumbu Dapur', 'bks', 1000, 0],
  ['BD47', 'bumbu pecel', 'Bumbu Dapur', 'bks', 0, 0],
  ['BD48', 'asem jawa', 'Bumbu Dapur', 'bks', 8000, 0],
  ['BD49', 'Lada hitam', 'Bumbu Dapur', 'kg', 200000, 0]
];
