/**
 * ============================================================
 *  Auth.gs : Manajemen sesi tim harian
 * ============================================================
 *
 *  BUKAN login berbasis password. Ini adalah verifikasi
 *  "siapa yang bertugas hari ini":
 *    - Operator  : satu orang yang menginput data hari ini
 *    - AnggotaTim: seluruh tim koperasi yang bertugas hari ini
 *                  (operator otomatis ikut tergabung)
 *
 *  Disimpan per tanggal di sheet SesiTim sehingga tercatat
 *  siapa yang bertanggung jawab atas data setiap harinya.
 */

/**
 * Memulai / memperbarui sesi tim untuk suatu tanggal.
 *
 * @param {Object} data { tanggal, operator, anggota: [..] }
 * @return {Object} sesi tim yang tersimpan
 */
function mulaiSesiTim(data) {
  if (!data || !data.operator || String(data.operator).trim() === '') {
    throw new Error('Nama penginput data (operator) wajib diisi.');
  }

  const tanggal = formatDate(data.tanggal || new Date());
  const operator = String(data.operator).trim();

  // Bersihkan & gabungkan anggota; operator wajib termasuk.
  let anggota = Array.isArray(data.anggota) ? data.anggota : [];
  anggota = anggota
    .map(function (n) { return String(n).trim(); })
    .filter(function (n) { return n !== ''; });

  // Pastikan operator ada di daftar tim (otomatis disertakan).
  const sudahAda = anggota.some(function (n) {
    return n.toLowerCase() === operator.toLowerCase();
  });
  if (!sudahAda) {
    anggota.unshift(operator);
  }

  // Hilangkan duplikat (case-insensitive) sambil menjaga urutan.
  const seen = {};
  anggota = anggota.filter(function (n) {
    const key = n.toLowerCase();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });

  const sheet = getSheet(SHEETS.SESI_TIM);
  const rows = readSheetAsObjects(SHEETS.SESI_TIM);
  const now = new Date();
  const anggotaStr = anggota.join(', ');

  // Cari sesi pada tanggal yang sama -> update, jika tidak -> buat baru.
  const existing = rows.find(function (r) { return formatDate(r.Tanggal) === tanggal; });
  if (existing) {
    sheet.getRange(existing._row, 2).setValue(operator);    // Operator
    sheet.getRange(existing._row, 3).setValue(anggotaStr);  // AnggotaTim
    sheet.getRange(existing._row, 5).setValue(now);         // UpdatedAt
  } else {
    sheet.appendRow([tanggal, operator, anggotaStr, now, now]);
  }

  // Simpan anggota sebagai "anggota tetap" untuk saran di hari berikutnya.
  perbaruiAnggotaTetap(anggota);

  SpreadsheetApp.flush();
  return getSesiTim(tanggal);
}

/**
 * Mengambil sesi tim pada tanggal tertentu.
 * @return {Object|null}
 */
function getSesiTim(tanggal) {
  const tgl = formatDate(tanggal || new Date());
  const rows = readSheetAsObjects(SHEETS.SESI_TIM);
  const sesi = rows.find(function (r) { return formatDate(r.Tanggal) === tgl; });
  if (!sesi) return null;
  return {
    tanggal: tgl,
    tanggalLabel: formatDateLabel(tgl),
    operator: sesi.Operator,
    anggota: String(sesi.AnggotaTim || '')
      .split(',')
      .map(function (n) { return n.trim(); })
      .filter(function (n) { return n !== ''; })
  };
}

/**
 * Status awal aplikasi: cek apakah sesi hari ini sudah dibuat.
 * Dipakai front-end untuk memutuskan tampilkan halaman auth atau dashboard.
 */
function getInitState() {
  const today = formatDate(new Date());
  const sesi = getSesiTim(today);
  return {
    today: today,
    todayLabel: formatDateLabel(today),
    sesiHariIni: sesi,                       // null jika belum diisi
    anggotaTetap: getAnggotaTetap(),         // saran nama untuk autocomplete
    fase: FASE_LIST,
    kategori: Object.keys(KATEGORI)
  };
}

// ====================== ANGGOTA TETAP (SARAN) ======================

/**
 * Daftar anggota yang pernah tercatat -> dipakai sebagai saran cepat.
 */
function getAnggotaTetap() {
  const raw = getConfig('ANGGOTA_TETAP');
  if (!raw) return [];
  return String(raw)
    .split(',')
    .map(function (n) { return n.trim(); })
    .filter(function (n) { return n !== ''; });
}

/**
 * Menambahkan nama-nama baru ke daftar anggota tetap (tanpa duplikat).
 */
function perbaruiAnggotaTetap(namaBaru) {
  const current = getAnggotaTetap();
  const seen = {};
  current.forEach(function (n) { seen[n.toLowerCase()] = n; });
  namaBaru.forEach(function (n) {
    if (n && !seen[n.toLowerCase()]) seen[n.toLowerCase()] = n;
  });
  const gabung = Object.keys(seen).map(function (k) { return seen[k]; });
  setConfig('ANGGOTA_TETAP', gabung.join(', '));
}

/**
 * Menghapus satu nama dari daftar anggota tetap.
 */
function hapusAnggotaTetap(nama) {
  const current = getAnggotaTetap().filter(function (n) {
    return n.toLowerCase() !== String(nama).trim().toLowerCase();
  });
  setConfig('ANGGOTA_TETAP', current.join(', '));
  return getAnggotaTetap();
}
