# Koperasi Dashboard

Web App **Google Apps Script** untuk kasir koperasi harian — menggantikan input manual di spreadsheet lama dengan dashboard modern yang cepat, dengan kalkulasi otomatis, manajemen stok, pencatatan pembayaran, dan rekap finansial.

Tema visual mengikuti logo (gradasi **ungu–cyan**) dengan sentuhan modern-futuristik (glassmorphism halus).

---

## ✨ Fitur Utama

| Modul | Fungsi |
|---|---|
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
3. Salin semua file `.gs` dan `.html` dari repo ini ke project Apps Script:
   - `Code.gs`, `Setup.gs`, `Auth.gs`, `MasterBahan.gs`, `Stok.gs`, `Transaksi.gs`, `Laporan.gs`
   - `Index.html`, `Styles.html`, `Script.html`
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

1. **Buka URL** → muncul halaman verifikasi tim.
2. Isi **nama penginput data** (operator) → otomatis masuk daftar tim.
3. Tambah **anggota tim lain** yang bertugas (Enter untuk menambah).
4. Klik **Masuk Dashboard**.
5. **Transaksi** → input jumlah distribusi tiap bahan ke tiap fase → **Simpan Laporan**.
6. **Stok** → input stok masuk / tambah bahan baru bila perlu.
7. **Dashboard** → cek ringkasan & catat pembayaran tiap fase.
8. **Laporan** → **Buat Laporan Format Lama** untuk arsip/cetak.

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
├── Code.gs              # Router doGet, konstanta global, helper
├── Setup.gs             # Inisialisasi DB + seed master bahan
├── Auth.gs              # Sesi tim harian (operator + anggota)
├── MasterBahan.gs       # CRUD master bahan
├── Stok.gs              # Logic stok harian + carry-over
├── Transaksi.gs         # Distribusi per fase + kalkulasi
├── Laporan.gs           # Pembayaran, rekap, generate format lama
├── Index.html           # Struktur SPA
├── Styles.html          # Tema ungu-cyan futuristik
└── Script.html          # Client-side logic
```

---

*Dibuat berdasarkan analisis "Laporan Koperasi Harian". Format laporan lama tetap dapat diproduksi otomatis sehingga kebiasaan lama tidak terganggu.*
