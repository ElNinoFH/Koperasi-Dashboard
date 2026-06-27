/**
 * ============================================================
 *  MasterBahan.gs : CRUD data master bahan
 * ============================================================
 *  MasterBahan adalah sumber kebenaran tunggal (single source
 *  of truth) untuk daftar bahan, satuan, dan harga.
 */

/**
 * Mengambil seluruh baris master (raw, termasuk yang tidak aktif).
 * @return {Array<Object>}
 */
function getMasterBahanRaw() {
  return readSheetAsObjects(SHEETS.MASTER);
}

/**
 * Mengambil daftar bahan untuk konsumsi front-end.
 * @param {boolean} hanyaAktif jika true, sembunyikan bahan non-aktif
 * @return {Array<Object>}
 */
function getMasterBahan(hanyaAktif) {
  const rows = getMasterBahanRaw();
  const list = rows.map(function (r) {
    return {
      kode: r.KodeBarang,
      nama: r.NamaBahan,
      kategori: r.Kategori,
      satuan: r.Satuan,
      hargaBeli: Number(r.HargaBeli) || 0,
      hargaJual: Number(r.HargaJual) || 0,
      stokMin: Number(r.StokMin) || 0,
      aktif: r.Aktif === true || r.Aktif === 'TRUE' || r.Aktif === 'true',
      row: r._row
    };
  });
  if (hanyaAktif) {
    return list.filter(function (b) { return b.aktif; });
  }
  return list;
}

/**
 * Mencari satu bahan berdasarkan kode.
 * @return {Object|null}
 */
function getBahanByKode(kode) {
  const list = getMasterBahan(false);
  return list.find(function (b) { return b.kode === kode; }) || null;
}

/**
 * Menambah bahan baru. Kode dibuat otomatis berdasarkan kategori
 * jika tidak disediakan.
 * @param {Object} data { kode?, nama, kategori, satuan, hargaBeli, hargaJual?, stokMin? }
 */
function tambahBahan(data) {
  if (!data.nama || String(data.nama).trim() === '') {
    throw new Error('Nama bahan wajib diisi.');
  }
  if (!data.kategori || !KATEGORI[data.kategori]) {
    throw new Error('Kategori tidak valid.');
  }

  const sheet = getSheet(SHEETS.MASTER);
  let kode = data.kode && String(data.kode).trim();
  if (!kode) {
    kode = generateKodeBerikutnya(data.kategori);
  }

  // Pastikan kode unik
  if (getBahanByKode(kode)) {
    throw new Error('Kode "' + kode + '" sudah dipakai.');
  }

  const hargaBeli = Number(data.hargaBeli) || 0;
  const hargaJual = Number(data.hargaJual) || hargaBeli;

  sheet.appendRow([
    kode,
    String(data.nama).trim(),
    data.kategori,
    data.satuan || '',
    hargaBeli,
    hargaJual,
    Number(data.stokMin) || 0,
    true
  ]);
  SpreadsheetApp.flush();
  return getBahanByKode(kode);
}

/**
 * Memperbarui bahan yang sudah ada (berdasarkan kode).
 */
function updateBahan(data) {
  const bahan = getBahanByKode(data.kode);
  if (!bahan) throw new Error('Bahan dengan kode "' + data.kode + '" tidak ditemukan.');

  const sheet = getSheet(SHEETS.MASTER);
  const row = bahan.row;

  sheet.getRange(row, 2).setValue(data.nama != null ? String(data.nama).trim() : bahan.nama);
  sheet.getRange(row, 3).setValue(data.kategori || bahan.kategori);
  sheet.getRange(row, 4).setValue(data.satuan != null ? data.satuan : bahan.satuan);
  sheet.getRange(row, 5).setValue(data.hargaBeli != null ? Number(data.hargaBeli) : bahan.hargaBeli);
  sheet.getRange(row, 6).setValue(data.hargaJual != null ? Number(data.hargaJual) : bahan.hargaJual);
  sheet.getRange(row, 7).setValue(data.stokMin != null ? Number(data.stokMin) : bahan.stokMin);
  if (data.aktif != null) sheet.getRange(row, 8).setValue(!!data.aktif);

  SpreadsheetApp.flush();
  return getBahanByKode(data.kode);
}

/**
 * Menonaktifkan bahan (soft delete agar riwayat transaksi tetap valid).
 */
function nonaktifkanBahan(kode) {
  return updateBahan({ kode: kode, aktif: false });
}

/**
 * Membuat kode berikutnya untuk suatu kategori (mis. K14, P28, ...).
 */
function generateKodeBerikutnya(kategori) {
  const prefix = KATEGORI[kategori];
  const list = getMasterBahan(false).filter(function (b) {
    return b.kategori === kategori;
  });
  let maxNum = 0;
  list.forEach(function (b) {
    const m = String(b.kode).match(new RegExp('^' + prefix + '(\\d+)$'));
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxNum) maxNum = n;
    }
  });
  const next = maxNum + 1;
  // Lebar angka mengikuti prefix: BD pakai 2 digit, lainnya 2 digit juga
  return prefix + (next < 10 ? '0' + next : '' + next);
}
