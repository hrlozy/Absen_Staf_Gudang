# Absen_Staf_Gudang

Sistem Presensi & Dashboard KPI Staf Packing Gudang (LOZY PCP Service).

## 🚀 Fitur Utama
- **Frontend**: Live Barcode/QR Scanner berbasis kamera WebRTC real-time + Autocomplete Nama Karyawan + Pemilihan Shift (Shift 1, Shift 2, Shift 3).
- **Backend & Database**: Google Apps Script & Google Sheets (`Shift 1`, `Shift 2`, `Shift 3`, `Master_Karyawan`, `Pengaturan`).
- **Dashboard KPI**: Rekap kedisiplinan staf, kehadiran harian, dan grafik ketepatan waktu.
- **Hosting**: GitHub Pages (HTTPS Top-Level Domain untuk akses kamera langsung).

## 📁 Struktur File
- `index.html`: Halaman utama presensi dan scanner live barcode.
- `dashboard.html`: Halaman monitoring kedisiplinan dan rekap kehadiran.
- `Kode.gs`: Script backend untuk Google Apps Script (menangani API endpoint dan penyimpanan spreadsheet).
