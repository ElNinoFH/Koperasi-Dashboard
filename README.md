# Koperasi Dashboard

Web App **Google Apps Script** untuk kasir koperasi harian — menggantikan input manual di spreadsheet lama dengan dashboard modern yang cepat, dengan kalkulasi otomatis, manajemen stok, pencatatan pembayaran, dan rekap finansial.

Tema visual mengikuti logo (gradasi **ungu–cyan**) dengan sentuhan modern-futuristik (glassmorphism halus).

---

## ✨ Fitur Utama

| Modul | Fungsi |
|---|---|
| **Gerbang Kode Akses (PIN)** | Halaman pertama sebelum apa pun lain bisa diakses. Mencegah orang yang sekadar tahu link web app untuk melihat/mengubah data keuangan koperasi (lihat bagian [Keamanan](#-keamanan)). |
| **Verifikasi Tim Harian** | Sebelum masuk, kasir mencatat siapa **penginput data (operator)** dan **seluruh tim** yang bertugas hari ini. Operator otomatis tergabung dalam daftar tim. |
| **Dashboard** | Ringkasan penjualan, modal (HPP), laba/rugi, total belanja, grafik distribusi per fase, status pembayaran, dan peringatan stok kritis. |
| **Transaksi Harian** | Input distribusi bahan ke 9 fase/kelompok (Fase PS, A, B, C, SFS, Makan Malam, Sarapan PJ, Guru, IPAK) dengan kalkulasi harga real-time dan validasi stok. |
| **Manajemen Stok** | Stok realtime dengan **carry-over otomatis** antar hari, input stok masuk, tambah bahan baru, dan deteksi stok minimum. |
| **Laporan & Rekap** | Ringkasan finansial harian, rekap 7 hari terakhir, dan **generate laporan dengan format identik spreadsheet lama** untuk arsip/cetak. |

---

## 🏛️ Arsitektur (Hybrid)

Spreadsheet baru yang **terstruktur** dijadikan database. Laporan format lama tetap bisa di-generate kapan saja untuk arsip dan cetak.

```
┌──────────────────────────────────────────────┐
│         WEB APP (Dashboard Kasir)             │
│   Index.html · Styles.html · Script.html      │
└────────────────────┬─────────────────────────┘
                     │ google.script.run
┌────────────────────▼─────────────────────────┐
│        BACKEND (Google Apps Script)           │
│  Code · Auth · MasterBahan · Stok ·           │
│  Transaksi · Laporan · Setup                  │
└────────────────────┬─────────────────────────┘
                     │ SpreadsheetApp
┌────────────────────▼─────────────────────────┐
│          GOOGLE SHEETS (Database)             │
│  MasterBahan · StokHarian · Transaksi ·       │
│  Pembayaran · SesiTim · Config                │
│  + Laporan_[tgl] (format lama, auto-generated)│
└──────────────────────────────────────────────┘
```

### Struktur Sheet

| Sheet | Isi |
|---|---|
| `MasterBahan` | Daftar bahan, kategori, satuan, harga beli/jual, stok minimum (sumber kebenaran tunggal). |
| `StokHarian` | Ledger stok per tanggal: StokAwal, Masuk, Keluar, StokAkhir (carry-over otomatis). |
| `Transaksi` | Format **TALL**: 1 baris = 1 bahan untuk 1 fase pada 1 tanggal. |
| `Pembayaran` | Total belanja, dibayar, sisa, & status per fase per tanggal. |
| `SesiTim` | Operator & anggota tim yang bertugas per tanggal. |
| `Config` | Penyimpanan key-value (mis. daftar anggota tetap untuk autocomplete). |

### Rumus Inti (mengikuti spreadsheet lama)

```
StokAkhir       = StokAwal + Masuk - Keluar
Keluar          = Σ Jumlah semua fase (dari Transaksi)
TotalHarga fase = Jumlah × HargaSatuan
Total per Fase  = Σ TotalHarga (dibulatkan ke ribuan terdekat)
Total Modal     = Σ (StokAwal + Masuk) × HargaBeli
Total Penjualan = Σ Jumlah keluar × HargaJual
Keuntungan      = Total Penjualan − Total Modal
```

---

## 🚀 Cara Deploy

> Aplikasi ini berjalan sebagai **Container-bound script** di sebuah Google Spreadsheet.

### 1. Siapkan Spreadsheet & Project
1. Buat Google Spreadsheet baru (ini akan menjadi database).
2. Menu **Extensions → Apps Script**.
3. Salin file dari repo ini ke project Apps Script (hanya 2 file):
   - `Code.gs` — seluruh backend
   - `Index.html` — seluruh frontend (HTML + CSS + JS)
   - Salin juga isi `appsscript.json` (aktifkan **Project Settings → Show "appsscript.json"**).

### 2. Inisialisasi Database
1. Di editor Apps Script, pilih fungsi **`setupDatabase`** lalu klik **Run**.
2. Berikan izin yang diminta.
3. Setelah selesai, spreadsheet akan berisi sheet `MasterBahan` (terisi ±159 bahan) dan sheet lainnya.

### 3. Deploy sebagai Web App
1. Klik **Deploy → New deployment**.
2. Pilih tipe **Web app**.
3. Setelan:
   - **Execute as**: `Me` (pemilik)
   - **Who has access**: sesuai kebutuhan (mis. `Anyone` atau `Anyone within organization`)
4. Klik **Deploy** dan salin **URL Web App**. Bagikan ke kasir.

### Menggunakan `clasp` (opsional)
```bash
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "Koperasi Dashboard"
clasp push
```
Lalu jalankan `setupDatabase` & buat deployment dari editor.

---

## 📖 Alur Pakai Kasir

1. **Buka URL** → muncul **gerbang kode akses (PIN)**.
2. Masukkan PIN yang benar → muncul halaman verifikasi tim.
3. Isi **nama penginput data** (operator) → otomatis masuk daftar tim.
4. Tambah **anggota tim lain** yang bertugas (Enter untuk menambah).
5. Klik **Masuk Dashboard**.
6. **Transaksi** → input jumlah distribusi tiap bahan ke tiap fase → **Simpan Laporan**.
7. **Stok** → input stok masuk / tambah bahan baru bila perlu.
8. **Dashboard** → cek ringkasan & catat pembayaran tiap fase.
9. **Laporan** → **Buat Laporan Format Lama** untuk arsip/cetak.

---

## 🔒 Keamanan

Web app ini di-deploy dengan `appsscript.json` → `"access": "ANYONE_ANONYMOUS"`.
Ini **disengaja** — aplikasi dipakai kasir/tim koperasi lewat link, tanpa
mengharuskan mereka login akun Google. Konsekuensinya: **siapa pun yang
tahu link deployment bisa membuka halamannya**, jadi ada lapisan keamanan
tambahan di level aplikasi:

- **Gerbang PIN**: halaman pertama yang muncul sekarang adalah input kode
  akses (PIN), sebelum halaman verifikasi tim/dashboard bisa diakses.
  PIN divalidasi lewat fungsi backend `verifyAccessCode(pin)`, yang
  mengembalikan token sesi sederhana (disimpan di `CacheService`, berlaku
  6 jam) bila PIN benar.
- **Validasi ganda di backend**: fungsi-fungsi yang mengubah data
  keuangan/stok (`tambahBahan`, `updateBahan`, `catatPembayaran`,
  `simpanDistribusiBatch`, `simpanDistribusiBahan`, `inputStokMasuk`,
  `mulaiSesiTim`) juga memvalidasi token sesi tersebut lewat
  `requireAccess_()` di sisi server — jadi gerbang PIN di frontend tidak
  bisa dilewati begitu saja lewat console browser.
- **PIN disimpan di `PropertiesService.getScriptProperties()`**, bukan
  hardcode di kode, supaya tetap aman meski repo GitHub ini publik.

### PIN default — WAJIB DIGANTI

Saat pertama kali dipakai (belum pernah ada PIN tersimpan), aplikasi
otomatis memakai **PIN default `1234`**. Ini **BUKAN** PIN yang aman untuk
produksi — siapa pun yang membaca kode/README ini juga tahu PIN-nya.

**Ganti PIN sebelum dipakai mencatat data keuangan sungguhan**, dengan
salah satu cara:

1. **Lewat aplikasi**: masuk dengan PIN lama (`1234` bila belum pernah
   diganti), lalu panggil `changeAccessCode(token, 'pinBaruAnda')` — token
   didapat otomatis setelah `verifyAccessCode` sukses. (Saat ini belum ada
   tombol UI khusus untuk ini — bisa dipanggil lewat console browser saat
   sudah login, atau ditambahkan tombol "Ganti PIN" di pengaturan bila
   diperlukan.)
2. **Manual dari editor Apps Script** (paling mudah): buka
   **Extensions → Apps Script**, lalu jalankan sekali baris berikut lewat
   fungsi sementara atau langsung di **Execution log**:
   ```js
   PropertiesService.getScriptProperties().setProperty('ACCESS_PIN', 'pinBaruAnda');
   ```

Catatan: PIN ini adalah lapisan keamanan **sederhana** untuk mencegah
orang iseng yang kebetulan mendapat link, bukan pengganti otentikasi
sungguhan — jangan bagikan PIN di tempat publik, dan ganti berkala bila
perlu.

---

## 🎨 Palet Warna

| Token | Hex | Pemakaian |
|---|---|---|
| Violet | `#8B5CF6` | Aksen utama, header |
| Violet Light | `#A78BFA` | Gradien, highlight |
| Cyan | `#2DD4BF` | Aksen sekunder, nilai aktif |
| Blue | `#60A5FA` | Gradien tengah |
| Success | `#34D399` | Lunas, stok aman |
| Warning | `#FBBF24` | Stok menipis |
| Danger | `#F87171` | Rugi, stok habis |

---

## 📁 Struktur File

```
Koperasi-Dashboard/
├── appsscript.json      # Manifest (timezone, webapp config)
├── Code.gs              # SELURUH backend (setup, auth, stok, transaksi, laporan)
└── Index.html           # SELURUH frontend (HTML + CSS + JS dalam 1 file)
```

---

*Dibuat berdasarkan analisis "Laporan Koperasi Harian". Format laporan lama tetap dapat diproduksi otomatis sehingga kebiasaan lama tidak terganggu.*
