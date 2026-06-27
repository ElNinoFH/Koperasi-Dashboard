/**
 * ============================================================
 *  Laporan.gs : Pembayaran, rekap finansial, & laporan format lama
 * ============================================================
 */

// ====================== PEMBAYARAN ======================

/**
 * Menghitung ulang baris Pembayaran berdasarkan ringkasan fase.
 * Mempertahankan nilai "Dibayar" yang sudah diinput user.
 */
function hitungUlangPembayaran(tanggal) {
  const tgl = formatDate(tanggal);
  const ringkasan = hitungRingkasanFase(tgl);
  const sheet = getSheet(SHEETS.PEMBAYARAN);
  const rows = readSheetAsObjects(SHEETS.PEMBAYARAN);
  const now = new Date();

  FASE_LIST.forEach(function (f) {
    const totalBulat = ringkasan.perFase[f].totalBulat;
    const existing = rows.find(function (r) {
      return formatDate(r.Tanggal) === tgl && r.Fase === f;
    });

    if (existing) {
      const dibayar = Number(existing.Dibayar) || 0;
      const sisa = totalBulat - dibayar;
      sheet.getRange(existing._row, 3).setValue(totalBulat);            // TotalBelanja
      sheet.getRange(existing._row, 5).setValue(sisa);                  // Sisa
      sheet.getRange(existing._row, 6).setValue(statusBayar(totalBulat, dibayar)); // Status
      sheet.getRange(existing._row, 7).setValue(now);
    } else if (totalBulat > 0) {
      sheet.appendRow([tgl, f, totalBulat, 0, totalBulat, 'Belum Bayar', now]);
    }
  });
}

/**
 * Menentukan label status pembayaran.
 */
function statusBayar(total, dibayar) {
  if (total <= 0) return '-';
  if (dibayar >= total) return 'Lunas';
  if (dibayar > 0) return 'Kurang';
  return 'Belum Bayar';
}

/**
 * Mengambil data pembayaran per fase pada tanggal tertentu.
 * @return {Array<Object>}
 */
function getPembayaran(tanggal) {
  const tgl = formatDate(tanggal);
  hitungUlangPembayaran(tgl);
  const rows = readSheetAsObjects(SHEETS.PEMBAYARAN).filter(function (r) {
    return formatDate(r.Tanggal) === tgl;
  });
  return rows.map(function (r) {
    return {
      fase: r.Fase,
      totalBelanja: Number(r.TotalBelanja) || 0,
      dibayar: Number(r.Dibayar) || 0,
      sisa: Number(r.Sisa) || 0,
      status: r.Status
    };
  });
}

/**
 * Mencatat pembayaran untuk sebuah fase.
 * @param {Object} data { tanggal, fase, dibayar }
 */
function catatPembayaran(data) {
  const tgl = formatDate(data.tanggal);
  hitungUlangPembayaran(tgl);
  const sheet = getSheet(SHEETS.PEMBAYARAN);
  const rows = readSheetAsObjects(SHEETS.PEMBAYARAN);
  const existing = rows.find(function (r) {
    return formatDate(r.Tanggal) === tgl && r.Fase === data.fase;
  });
  if (!existing) throw new Error('Belum ada belanja untuk fase ' + data.fase);

  const total = Number(existing.TotalBelanja) || 0;
  const dibayar = Number(data.dibayar) || 0;
  sheet.getRange(existing._row, 4).setValue(dibayar);
  sheet.getRange(existing._row, 5).setValue(total - dibayar);
  sheet.getRange(existing._row, 6).setValue(statusBayar(total, dibayar));
  sheet.getRange(existing._row, 7).setValue(new Date());
  SpreadsheetApp.flush();
  return getPembayaran(tgl);
}

// ====================== REKAP FINANSIAL ======================

/**
 * Ringkasan finansial harian lengkap untuk dashboard.
 * @return {Object}
 */
function getRekapHarian(tanggal) {
  const tgl = formatDate(tanggal || new Date());
  const trx = getTransaksiByTanggal(tgl);
  const ringkasan = hitungRingkasanFase(tgl);
  const pembayaran = getPembayaran(tgl);
  const snapshot = getSnapshotStok(tgl);

  // Total Modal (HPP) = nilai seluruh stok yang tersedia hari ini (StokAwal+Masuk) x hargaBeli.
  let totalModal = 0;
  snapshot.forEach(function (b) {
    totalModal += (b.stokAwal + b.masuk) * b.hargaBeli;
  });

  // Total Penjualan = nilai bahan yang terdistribusi (keluar) x hargaJual.
  let totalPenjualan = 0;
  trx.forEach(function (t) {
    const bahan = getBahanByKodeCached(t.KodeBarang);
    const hargaJual = bahan ? bahan.hargaJual : (Number(t.HargaSatuan) || 0);
    totalPenjualan += (Number(t.Jumlah) || 0) * hargaJual;
  });

  const totalDibayar = pembayaran.reduce(function (s, p) { return s + p.dibayar; }, 0);
  const totalKekurangan = pembayaran.reduce(function (s, p) { return s + Math.max(0, p.sisa); }, 0);

  return {
    tanggal: tgl,
    tanggalLabel: formatDateLabel(tgl),
    totalPenjualan: totalPenjualan,
    totalModal: totalModal,
    totalBelanja: ringkasan.grandTotalBulat,
    keuntungan: totalPenjualan - totalModal,
    totalDibayar: totalDibayar,
    totalKekurangan: totalKekurangan,
    perFase: ringkasan.perFase,
    pembayaran: pembayaran,
    jumlahTransaksi: trx.length
  };
}

// Cache sederhana untuk getBahanByKode (mengurangi pembacaan sheet berulang).
let _bahanCache = null;
function getBahanByKodeCached(kode) {
  if (!_bahanCache) {
    _bahanCache = {};
    getMasterBahan(false).forEach(function (b) { _bahanCache[b.kode] = b; });
  }
  return _bahanCache[kode] || null;
}

/**
 * Rekap mingguan: agregasi 7 hari berakhir pada tanggal (atau rentang custom).
 * @param {string} tanggalAkhir
 * @param {number} jumlahHari default 7
 */
function getRekapMingguan(tanggalAkhir, jumlahHari) {
  const akhir = tanggalAkhir ? new Date(tanggalAkhir + 'T00:00:00') : new Date();
  const n = jumlahHari || 7;
  const harian = [];
  let totPenjualan = 0, totModal = 0, totBelanja = 0;

  for (let i = n - 1; i >= 0; i--) {
    const d = new Date(akhir);
    d.setDate(d.getDate() - i);
    const tgl = formatDate(d);
    const rekap = getRekapHarian(tgl);
    harian.push({
      tanggal: tgl,
      label: Utilities.formatDate(d, TIMEZONE, 'EEE dd/MM'),
      penjualan: rekap.totalPenjualan,
      modal: rekap.totalModal,
      belanja: rekap.totalBelanja,
      keuntungan: rekap.keuntungan
    });
    totPenjualan += rekap.totalPenjualan;
    totModal += rekap.totalModal;
    totBelanja += rekap.totalBelanja;
  }

  return {
    dari: harian[0] ? harian[0].tanggal : formatDate(akhir),
    sampai: formatDate(akhir),
    harian: harian,
    totalPenjualan: totPenjualan,
    totalModal: totModal,
    totalBelanja: totBelanja,
    totalKeuntungan: totPenjualan - totModal
  };
}

// ====================== GENERATE LAPORAN FORMAT LAMA ======================

/**
 * Membuat sheet laporan dengan format identik spreadsheet lama
 * (read-only, untuk arsip & cetak). Mengembalikan URL ke sheet tsb.
 *
 * @param {string} tanggal
 * @return {Object} { sheetName, url }
 */
function generateLaporanFormatLama(tanggal) {
  const tgl = formatDate(tanggal || new Date());
  const ss = getSpreadsheet();
  const namaSheet = 'Laporan_' + tgl;

  // Hapus sheet lama bila sudah ada agar selalu fresh.
  const lama = ss.getSheetByName(namaSheet);
  if (lama) ss.deleteSheet(lama);
  const sheet = ss.insertSheet(namaSheet);

  const snapshot = getSnapshotStok(tgl);
  const matriks = getMatriksDistribusi(tgl);
  const sesi = getSesiTim(tgl);

  // Map kode -> distribusi & data tambahan.
  const bahanMap = {};
  matriks.bahan.forEach(function (b) { bahanMap[b.kode] = b; });

  const fase = FASE_LIST;
  // --- Header utama ---
  const totalKolom = 8 + fase.length * 2; // info dasar + (jumlah,harga) per fase
  sheet.getRange(1, 1).setValue('Anggota Koperasi');
  sheet.getRange(1, 9).setValue('Laporan Belanja Harian');
  sheet.getRange(1, totalKolom).setValue(formatDateLabel(tgl));
  if (sesi) {
    sheet.getRange(2, 1).setValue('Operator: ' + sesi.operator);
    sheet.getRange(2, 9).setValue('Tim: ' + sesi.anggota.join(', '));
  }

  // --- Baris header kolom ---
  const headerRow = 4;
  const headerInfo = ['Kode', 'Nama Bahan', 'Stok Awal', 'Masuk', 'Stok Akhir', 'Satuan', 'Harga', 'HPP'];
  let header = headerInfo.slice();
  fase.forEach(function (f) { header.push(f); header.push(''); });
  sheet.getRange(headerRow, 1, 1, header.length).setValues([header]);

  const subRow = headerRow + 1;
  let sub = ['', '', '', '', '', '', '', ''];
  fase.forEach(function () { sub.push('Jumlah'); sub.push('Harga'); });
  sheet.getRange(subRow, 1, 1, sub.length).setValues([sub]);

  // --- Isi data per kategori ---
  let r = subRow + 1;
  const totalPerFase = {};
  fase.forEach(function (f) { totalPerFase[f] = 0; });

  Object.keys(KATEGORI).forEach(function (kat) {
    // Baris judul kategori
    sheet.getRange(r, 1).setValue(kat).setFontWeight('bold').setBackground('#EDE9FE');
    r++;
    snapshot.filter(function (b) { return b.kategori === kat; }).forEach(function (b) {
      const bm = bahanMap[b.kode] || { distribusi: {} };
      const hpp = (b.stokAwal + b.masuk) * b.hargaBeli;
      const baris = [
        b.kode, b.nama, b.stokAwal, b.masuk, b.stokAkhir, b.satuan, b.hargaBeli, hpp
      ];
      fase.forEach(function (f) {
        const j = (bm.distribusi && bm.distribusi[f]) || 0;
        const h = j * b.hargaBeli;
        baris.push(j > 0 ? j : '');
        baris.push(h > 0 ? h : 0);
        totalPerFase[f] += h;
      });
      sheet.getRange(r, 1, 1, baris.length).setValues([baris]);
      r++;
    });
  });

  // --- Baris total per fase ---
  r++;
  const totalRow = ['Total Belanja Per Fase', '', '', '', '', '', '', ''];
  fase.forEach(function (f) {
    totalRow.push('');
    totalRow.push(bulatkanRibuan(totalPerFase[f]));
  });
  sheet.getRange(r, 1, 1, totalRow.length).setValues([totalRow]);
  sheet.getRange(r, 1).setFontWeight('bold');
  r++;

  // --- Ringkasan finansial ---
  const rekap = getRekapHarian(tgl);
  r++;
  const ringkasanData = [
    ['Total Penjualan', rekap.totalPenjualan],
    ['Total Modal (HPP)', rekap.totalModal],
    ['Total Keuntungan', rekap.keuntungan],
    ['Total Dibayar', rekap.totalDibayar],
    ['Total Kekurangan', rekap.totalKekurangan]
  ];
  ringkasanData.forEach(function (row) {
    sheet.getRange(r, 1).setValue(row[0]).setFontWeight('bold');
    sheet.getRange(r, 8).setValue(row[1]);
    r++;
  });

  // --- Format tampilan ---
  sheet.getRange(headerRow, 1, 2, header.length)
    .setFontWeight('bold').setBackground('#7C3AED').setFontColor('#FFFFFF');
  sheet.setFrozenRows(subRow);
  sheet.autoResizeColumns(1, Math.min(header.length, 12));
  // Format kolom uang
  const lastRow = sheet.getLastRow();
  sheet.getRange(subRow + 1, 7, lastRow - subRow, 1).setNumberFormat('"Rp"#,##0');
  sheet.getRange(subRow + 1, 8, lastRow - subRow, 1).setNumberFormat('"Rp"#,##0');
  fase.forEach(function (f, idx) {
    const col = 9 + idx * 2 + 1; // kolom harga fase
    sheet.getRange(subRow + 1, col, lastRow - subRow, 1).setNumberFormat('"Rp"#,##0');
  });

  SpreadsheetApp.flush();
  return {
    sheetName: namaSheet,
    url: ss.getUrl() + '#gid=' + sheet.getSheetId()
  };
}
