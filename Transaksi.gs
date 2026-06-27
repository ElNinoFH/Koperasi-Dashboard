/**
 * ============================================================
 *  Transaksi.gs : Input distribusi bahan per fase
 * ============================================================
 *  Format TALL: 1 baris = 1 bahan untuk 1 fase pada 1 tanggal.
 *    TotalHarga = Jumlah x HargaSatuan
 *  Inilah inti sistem (menggantikan kolom-kolom fase di sheet lama).
 */

/**
 * Mengambil semua transaksi pada tanggal tertentu.
 * @return {Array<Object>}
 */
function getTransaksiByTanggal(tanggal) {
  const tgl = formatDate(tanggal);
  return readSheetAsObjects(SHEETS.TRANSAKSI).filter(function (r) {
    return formatDate(r.Tanggal) === tgl;
  });
}

/**
 * Map kode bahan -> total jumlah keluar (dari semua fase) pada tanggal tsb.
 * Dipakai oleh Stok.gs untuk menghitung "Keluar".
 * @return {Object}
 */
function getTotalKeluarByTanggal(tanggal) {
  const trx = getTransaksiByTanggal(tanggal);
  const map = {};
  trx.forEach(function (t) {
    const kode = t.KodeBarang;
    map[kode] = (map[kode] || 0) + (Number(t.Jumlah) || 0);
  });
  return map;
}

/**
 * Matriks distribusi untuk UI input: per bahan, jumlah tiap fase.
 * Menggabungkan snapshot stok + transaksi tersimpan.
 * @return {Object} { tanggal, fase, bahan: [ {..., distribusi:{Fase:jumlah}} ] }
 */
function getMatriksDistribusi(tanggal) {
  const tgl = formatDate(tanggal || new Date());
  const snapshot = getSnapshotStok(tgl);
  const trx = getTransaksiByTanggal(tgl);

  // Susun map: kode -> { fase -> jumlah }
  const distMap = {};
  trx.forEach(function (t) {
    if (!distMap[t.KodeBarang]) distMap[t.KodeBarang] = {};
    distMap[t.KodeBarang][t.Fase] = Number(t.Jumlah) || 0;
  });

  const bahan = snapshot.map(function (b) {
    const dist = {};
    let totalKeluar = 0;
    FASE_LIST.forEach(function (f) {
      const j = (distMap[b.kode] && distMap[b.kode][f]) || 0;
      dist[f] = j;
      totalKeluar += j;
    });
    return {
      kode: b.kode,
      nama: b.nama,
      kategori: b.kategori,
      satuan: b.satuan,
      hargaSatuan: b.hargaBeli,
      stokAwal: b.stokAwal,
      masuk: b.masuk,
      stokAkhir: b.stokAkhir,
      stokMin: b.stokMin,
      kritis: b.kritis,
      distribusi: dist,
      totalKeluar: totalKeluar
    };
  });

  return {
    tanggal: tgl,
    tanggalLabel: formatDateLabel(tgl),
    fase: FASE_LIST,
    bahan: bahan,
    ringkasan: hitungRingkasanFase(tgl)
  };
}

/**
 * Menyimpan distribusi untuk satu bahan (seluruh fase sekaligus).
 * Menghapus baris lama bahan+tanggal, lalu menulis ulang yang > 0.
 *
 * @param {Object} data { tanggal, kode, distribusi: {Fase: jumlah}, operator }
 */
function simpanDistribusiBahan(data) {
  const tgl = formatDate(data.tanggal || new Date());
  const bahan = getBahanByKode(data.kode);
  if (!bahan) throw new Error('Bahan tidak ditemukan: ' + data.kode);

  const sheet = getSheet(SHEETS.TRANSAKSI);

  // Hapus transaksi lama untuk kode+tanggal ini (dari bawah ke atas).
  const data2 = sheet.getDataRange().getValues();
  for (let i = data2.length - 1; i >= 1; i--) {
    if (formatDate(data2[i][1]) === tgl && data2[i][2] === data.kode) {
      sheet.deleteRow(i + 1);
    }
  }

  // Tulis baris baru untuk fase dengan jumlah > 0.
  const now = new Date();
  const operator = data.operator || '';
  const dist = data.distribusi || {};
  const baris = [];
  FASE_LIST.forEach(function (f) {
    const jumlah = Number(dist[f]) || 0;
    if (jumlah > 0) {
      const total = jumlah * bahan.hargaBeli;
      baris.push([
        generateId('TRX'), tgl, bahan.kode, bahan.nama, bahan.kategori,
        bahan.satuan, f, jumlah, bahan.hargaBeli, total, operator, now
      ]);
    }
  });
  if (baris.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, baris.length, HEADERS.Transaksi.length).setValues(baris);
  }

  rekalkulasiStok(tgl);
  hitungUlangPembayaran(tgl);
  SpreadsheetApp.flush();

  return {
    kode: data.kode,
    ringkasan: hitungRingkasanFase(tgl)
  };
}

/**
 * Menyimpan banyak distribusi sekaligus (batch).
 * @param {Object} data { tanggal, items: [ {kode, distribusi} ], operator }
 */
function simpanDistribusiBatch(data) {
  const tgl = formatDate(data.tanggal || new Date());
  (data.items || []).forEach(function (item) {
    simpanDistribusiBahan({
      tanggal: tgl,
      kode: item.kode,
      distribusi: item.distribusi,
      operator: data.operator
    });
  });
  return getMatriksDistribusi(tgl);
}

/**
 * Menghitung total belanja per fase pada tanggal tertentu.
 * @return {Object} { perFase:{Fase:{total,totalBulat}}, grandTotal, grandTotalBulat }
 */
function hitungRingkasanFase(tanggal) {
  const trx = getTransaksiByTanggal(tanggal);
  const perFase = {};
  FASE_LIST.forEach(function (f) { perFase[f] = { total: 0, totalBulat: 0 }; });

  trx.forEach(function (t) {
    const f = t.Fase;
    if (perFase[f]) {
      perFase[f].total += Number(t.TotalHarga) || 0;
    }
  });

  let grandTotal = 0;
  let grandTotalBulat = 0;
  FASE_LIST.forEach(function (f) {
    perFase[f].totalBulat = bulatkanRibuan(perFase[f].total);
    grandTotal += perFase[f].total;
    grandTotalBulat += perFase[f].totalBulat;
  });

  return {
    perFase: perFase,
    grandTotal: grandTotal,
    grandTotalBulat: grandTotalBulat
  };
}

/**
 * Validasi: apakah ada bahan yang total distribusinya melebihi stok tersedia.
 * @return {Array<Object>} daftar pelanggaran
 */
function validasiStokTransaksi(tanggal) {
  const matriks = getMatriksDistribusi(tanggal);
  const pelanggaran = [];
  matriks.bahan.forEach(function (b) {
    const tersedia = b.stokAwal + b.masuk;
    if (b.totalKeluar > tersedia) {
      pelanggaran.push({
        kode: b.kode,
        nama: b.nama,
        tersedia: tersedia,
        diminta: b.totalKeluar,
        satuan: b.satuan
      });
    }
  });
  return pelanggaran;
}
