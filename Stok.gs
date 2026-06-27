/**
 * ============================================================
 *  Stok.gs : Manajemen stok harian (ledger)
 * ============================================================
 *  Logika kunci:
 *    StokAkhir = StokAwal + Masuk - Keluar
 *    StokAwal hari ini = StokAkhir hari sebelumnya (carry-over)
 *    Keluar otomatis dihitung dari total Transaksi pada tanggal tsb.
 */

/**
 * Mengambil StokAkhir terakhir yang tercatat untuk sebuah bahan
 * sebelum (atau pada) tanggal tertentu. Untuk carry-over.
 * @return {number}
 */
function getStokAkhirSebelum(kode, tanggal) {
  const rows = readSheetAsObjects(SHEETS.STOK).filter(function (r) {
    return r.KodeBarang === kode && formatDate(r.Tanggal) < tanggal;
  });
  if (rows.length === 0) return 0;
  rows.sort(function (a, b) {
    return formatDate(a.Tanggal) < formatDate(b.Tanggal) ? 1 : -1;
  });
  return Number(rows[0].StokAkhir) || 0;
}

/**
 * Mengambil baris stok untuk tanggal tertentu (semua bahan yang punya record).
 * @return {Object} map kode -> baris stok
 */
function getStokMapByTanggal(tanggal) {
  const tgl = formatDate(tanggal);
  const rows = readSheetAsObjects(SHEETS.STOK).filter(function (r) {
    return formatDate(r.Tanggal) === tgl;
  });
  const map = {};
  rows.forEach(function (r) { map[r.KodeBarang] = r; });
  return map;
}

/**
 * Snapshot stok lengkap untuk sebuah tanggal (gabungan master + ledger + carry-over).
 * Bahan tanpa record di tanggal tsb tetap tampil dengan StokAwal = carry-over.
 * @return {Array<Object>}
 */
function getSnapshotStok(tanggal) {
  const tgl = formatDate(tanggal || new Date());
  const master = getMasterBahan(true);
  const stokMap = getStokMapByTanggal(tgl);
  const keluarMap = getTotalKeluarByTanggal(tgl); // dari Transaksi

  return master.map(function (b) {
    const rec = stokMap[b.kode];
    let stokAwal, masuk;
    if (rec) {
      stokAwal = Number(rec.StokAwal) || 0;
      masuk = Number(rec.Masuk) || 0;
    } else {
      stokAwal = getStokAkhirSebelum(b.kode, tgl);
      masuk = 0;
    }
    const keluar = keluarMap[b.kode] || 0;
    const stokAkhir = stokAwal + masuk - keluar;
    const stokKritis = b.stokMin > 0 && stokAkhir <= b.stokMin;

    return {
      kode: b.kode,
      nama: b.nama,
      kategori: b.kategori,
      satuan: b.satuan,
      hargaBeli: b.hargaBeli,
      hargaJual: b.hargaJual,
      stokMin: b.stokMin,
      stokAwal: stokAwal,
      masuk: masuk,
      keluar: keluar,
      stokAkhir: stokAkhir,
      kritis: stokKritis,
      habis: stokAkhir <= 0
    };
  });
}

/**
 * Menambah / mengatur stok masuk untuk sebuah bahan pada tanggal tertentu.
 * Bila record belum ada, dibuat dengan StokAwal = carry-over.
 *
 * @param {Object} data { tanggal, kode, masuk, stokAwalOverride?, operator }
 */
function inputStokMasuk(data) {
  const tgl = formatDate(data.tanggal || new Date());
  const bahan = getBahanByKode(data.kode);
  if (!bahan) throw new Error('Bahan tidak ditemukan: ' + data.kode);

  const sheet = getSheet(SHEETS.STOK);
  const rows = readSheetAsObjects(SHEETS.STOK);
  const existing = rows.find(function (r) {
    return r.KodeBarang === data.kode && formatDate(r.Tanggal) === tgl;
  });

  const masuk = Number(data.masuk) || 0;
  const now = new Date();
  const operator = data.operator || '';

  if (existing) {
    const stokAwal = data.stokAwalOverride != null
      ? Number(data.stokAwalOverride)
      : (Number(existing.StokAwal) || 0);
    const masukTotal = masuk; // set absolut (bukan akumulasi) agar mudah dikoreksi
    sheet.getRange(existing._row, 6).setValue(stokAwal);   // StokAwal
    sheet.getRange(existing._row, 7).setValue(masukTotal); // Masuk
    sheet.getRange(existing._row, 10).setValue(operator);  // Operator
    sheet.getRange(existing._row, 11).setValue(now);       // UpdatedAt
    // StokAkhir & Keluar akan dihitung ulang oleh rekalkulasi
  } else {
    const stokAwal = data.stokAwalOverride != null
      ? Number(data.stokAwalOverride)
      : getStokAkhirSebelum(data.kode, tgl);
    sheet.appendRow([
      tgl, bahan.kode, bahan.nama, bahan.kategori, bahan.satuan,
      stokAwal, masuk, 0, stokAwal + masuk, operator, now
    ]);
  }

  rekalkulasiStok(tgl);
  SpreadsheetApp.flush();
  return getSnapshotStok(tgl);
}

/**
 * Memastikan setiap bahan punya baris stok pada tanggal tertentu.
 * Membuat baris carry-over untuk bahan yang belum ada record-nya.
 * Dipanggil saat membuka transaksi hari tsb.
 */
function pastikanBarisStok(tanggal, operator) {
  const tgl = formatDate(tanggal);
  const master = getMasterBahan(true);
  const stokMap = getStokMapByTanggal(tgl);
  const sheet = getSheet(SHEETS.STOK);
  const now = new Date();
  const baru = [];

  master.forEach(function (b) {
    if (!stokMap[b.kode]) {
      const stokAwal = getStokAkhirSebelum(b.kode, tgl);
      baru.push([tgl, b.kode, b.nama, b.kategori, b.satuan, stokAwal, 0, 0, stokAwal, operator || '', now]);
    }
  });

  if (baru.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, baru.length, HEADERS.StokHarian.length).setValues(baru);
  }
  return baru.length;
}

/**
 * Menghitung ulang kolom Keluar & StokAkhir berdasarkan Transaksi pada tanggal tsb.
 */
function rekalkulasiStok(tanggal) {
  const tgl = formatDate(tanggal);
  const sheet = getSheet(SHEETS.STOK);
  const keluarMap = getTotalKeluarByTanggal(tgl);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (formatDate(data[i][0]) === tgl) {
      const kode = data[i][1];
      const stokAwal = Number(data[i][5]) || 0;
      const masuk = Number(data[i][6]) || 0;
      const keluar = keluarMap[kode] || 0;
      const stokAkhir = stokAwal + masuk - keluar;
      sheet.getRange(i + 1, 8).setValue(keluar);    // Keluar
      sheet.getRange(i + 1, 9).setValue(stokAkhir); // StokAkhir
    }
  }
}

/**
 * Daftar bahan dengan stok kritis pada tanggal tertentu.
 */
function getStokKritis(tanggal) {
  return getSnapshotStok(tanggal).filter(function (b) {
    return b.kritis || b.habis;
  });
}
