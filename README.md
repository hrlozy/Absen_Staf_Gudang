# Absen_Staf_Gudang

Sistem Presensi & Dashboard KPI Staf Packing Gudang (LOZY PCP Service).

## Fitur Utama
- **Autentikasi Password Pribadi**: Setiap staf wajib memasukkan password/PIN pribadi mereka sebelum dapat melakukan pemindaian presensi. Password tersimpan aman pada sheet `Master_Karyawan`.
- **Verifikasi Barcode Posisi Kerja**: Pemindaian kamera memvalidasi Barcode Posisi di meja/stasiun kerja staf. Presensi hanya berhasil jika barcode yang dipindai sesuai dengan Posisi staf yang terdaftar di database.
- **Presensi Dual-Mode**: Mendukung presensi **Masuk (Check-In)** dan **Pulang (Check-Out)** dengan auto-detect shift kerja cerdas.
- **Live Barcode / QR Scanner**: Berbasis kamera WebRTC real-time dengan laser guidance + audio feedback beep, serta opsi fallback foto jika kamera perangkat terkendala.
- **Validasi Geofencing**: Menggunakan formula Haversine untuk menghitung radius jarak fisik staf ke gudang secara presisi.
- **Dashboard KPI & Real-Time Log**:
  - Rekap statistik kehadiran harian (Karyawan Aktif, Hadir Hari Ini, Pulang Hari Ini, dan per-shift).
  - Distribusi shift kerja via Chart.js Doughnut Chart.
  - Leaderboard akumulasi poin kehadiran staf.
  - Tabel aktivitas log presensi hari ini lengkap dengan filter pencarian instan.
  - Fitur **Export CSV** (kompatibel langsung dengan Microsoft Excel & Google Sheets).
- **Hosting**: Fleksibel berjalan murni di Google Apps Script Web App maupun di-hosting di GitHub Pages.

## Struktur Database `Master_Karyawan`
| ID Karyawan (A) | Nama (B) | Departemen (C) | Posisi (D) | Shift Assigned (E) | Barcode Posisi (F) | Password (G) |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| LOZY001 | Fulanah | People Ops | Staff Packing | Shift 1 | Staff Packing | 123456 |

## Struktur File
- [`index.html`](index.html): Halaman antarmuka presensi staf gudang (input nama, input password pribadi, selector tipe presensi & shift, live scanner posisi).
- [`dashboard.html`](dashboard.html): Halaman dashboard monitoring kedisiplinan, leaderboard poin, log aktivitas harian, dan ekspor CSV.
- [`Kode.gs`](Kode.gs): Script backend Google Apps Script yang menangani verifikasi password, validasi barcode posisi, geofencing, routing API, dan rekaman spreadsheet.
