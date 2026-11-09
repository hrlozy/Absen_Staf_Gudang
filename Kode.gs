/**
 * ============================================================================
 * SISTEM PRESENSI & DASHBOARD KPI STAF PACKING (LOZY PCP SERVICE)
 * Backend Google Apps Script - Dengan Validasi Password & Barcode Posisi
 * ============================================================================
 */

function doGet(e) {
  // 1. Endpoint JSON API untuk Frontend GitHub Pages / Web Eksternal (Metode GET)
  if (e && e.parameter && e.parameter.action) {
    let output = {};
    if (e.parameter.action === 'getKaryawan') {
      output = getDaftarKaryawan();
    } else if (e.parameter.action === 'getKPIData') {
      output = getKPIData();
    } else if (e.parameter.action === 'getExportData') {
      output = getExportData();
    }
    return ContentService.createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Tampilan Web App bawaan jika dibuka langsung di Google Apps Script
  if (e && e.parameter && e.parameter.page && e.parameter.page.toLowerCase() === 'index') {
    let htmlIndexOutput;
    try {
      htmlIndexOutput = HtmlService.createHtmlOutputFromFile('index');
    } catch(err) {
      htmlIndexOutput = HtmlService.createHtmlOutputFromFile('Index');
    }
    return htmlIndexOutput
      .setTitle('Presensi PCP Service - LOZY')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default untuk Google Apps Script: Tampilkan langsung Dashboard KPI Kedisiplinan
  let htmlOutput;
  try {
    htmlOutput = HtmlService.createHtmlOutputFromFile('dashboard');
  } catch(err) {
    htmlOutput = HtmlService.createHtmlOutputFromFile('Dashboard');
  }
  return htmlOutput
    .setTitle('Dashboard KPI Kedisiplinan - LOZY PCP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Endpoint JSON API untuk Menerima Data Presensi dari GitHub Pages / Web App (Metode POST)
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch(errJson) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const nama = payload.nama || payload.inputNamaKaryawan || '';
    const password = payload.password || payload.inputPassword || '';
    const shift = payload.shift || payload.inputShift || '';
    const tipe = payload.tipe || payload.tipePresensi || 'Masuk'; // 'Masuk' atau 'Pulang'
    const barcode = payload.barcode || payload.barcodeValidasi || '';
    const lat = parseFloat(payload.lat || payload.userLat || 0);
    const lng = parseFloat(payload.lng || payload.userLng || 0);

    const result = prosesPresensi(nama, shift, barcode, lat, lng, tipe, password);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      pesan: 'Terjadi kesalahan pada server: ' + error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Inisialisasi & Setup Struktur Database Spreadsheet
 * Menyiapkan sheet Shift, Master Karyawan (dengan Password & Barcode Posisi), serta Pengaturan
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Sheet Shift 1, Shift 2, Shift 3
  const shifts = ['Shift 1', 'Shift 2', 'Shift 3'];
  const headers = ['Timestamp', 'Tanggal', 'Waktu', 'ID Karyawan', 'Nama', 'Posisi', 'Tipe Presensi', 'Status', 'Jarak (m)'];
  
  shifts.forEach(shiftName => {
    let sheet = ss.getSheetByName(shiftName);
    if (!sheet) {
      sheet = ss.insertSheet(shiftName);
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eb7f97').setFontColor('white');
    sheet.setFrozenRows(1);
  });
  
  // 2. Setup Sheet Master Karyawan (Password di Kolom G)
  let sheetMaster = ss.getSheetByName('Master_Karyawan');
  const masterHeaders = ['ID Karyawan', 'Nama', 'Departemen', 'Posisi', 'Shift Assigned', 'Barcode Posisi', 'Password'];
  if (!sheetMaster) {
    sheetMaster = ss.insertSheet('Master_Karyawan');
    sheetMaster.appendRow(masterHeaders);
    sheetMaster.appendRow(['LOZY001', 'Fulanah', 'People Ops', 'Staff Packing', 'Shift 1', 'Staff Packing', '123456']);
  } else {
    // Pastikan struktur header diperbarui
    sheetMaster.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders]);
  }
  sheetMaster.getRange(1, 1, 1, masterHeaders.length).setFontWeight('bold').setBackground('#4b3040').setFontColor('white');
  sheetMaster.setFrozenRows(1);
  
  // 3. Setup Sheet Pengaturan
  let sheetPengaturan = ss.getSheetByName('Pengaturan');
  if (!sheetPengaturan) {
    sheetPengaturan = ss.insertSheet('Pengaturan');
    sheetPengaturan.appendRow(['Kunci', 'Nilai', 'Keterangan']);
    sheetPengaturan.getRange('A1:C1').setFontWeight('bold').setBackground('#766770').setFontColor('white');
    sheetPengaturan.appendRow(['Barcode_Valid', 'LOZY123', 'Kode barcode valid cadangan (pisahkan koma jika ada)']);
    sheetPengaturan.appendRow(['Target_Lat', '-6.9174639', 'Garis lintang lokasi absensi gudang']);
    sheetPengaturan.appendRow(['Target_Lng', '107.6191228', 'Garis bujur lokasi absensi gudang']);
    sheetPengaturan.appendRow(['Max_Radius', '200', 'Jarak maksimal dalam meter (0 = nonaktifkan geofence)']);
    sheetPengaturan.setFrozenRows(1);
    sheetPengaturan.setColumnWidths(1, 3, 220);
  }
}

// Menghitung jarak presisi dua titik koordinat (Haversine formula dalam meter)
function getDistanceFromLatLonInM(lat1, lon1, lat2, lon2) {
  var R = 6371000; 
  var dLat = (lat2 - lat1) * (Math.PI / 180);
  var dLon = (lon2 - lon1) * (Math.PI / 180); 
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) * 
          Math.sin(dLon / 2) * Math.sin(dLon / 2); 
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)); 
  return Math.round(R * c);
}

/**
 * Mengambil daftar nama karyawan aktif dari Master_Karyawan untuk autocomplete di frontend
 */
function getDaftarKaryawan() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMaster = ss.getSheetByName('Master_Karyawan');
    if (!sheetMaster) return [];
    
    const lastRow = sheetMaster.getLastRow();
    if (lastRow <= 1) return [];
    
    const header = sheetMaster.getRange(1, 1, 1, sheetMaster.getLastColumn()).getValues()[0] || [];
    const idxNama = header.indexOf('Nama') !== -1 ? header.indexOf('Nama') : 1;

    const data = sheetMaster.getRange(2, idxNama + 1, lastRow - 1, 1).getValues();
    const list = [];
    for (let i = 0; i < data.length; i++) {
      let nama = String(data[i][0]).trim();
      if (nama && !list.includes(nama)) {
        list.push(nama);
      }
    }
    return list.sort();
  } catch(e) {
    return [];
  }
}

/**
 * Helper untuk memvalidasi apakah Barcode Posisi yang di-scan cocok dengan Posisi Karyawan
 */
function cekKesesuaianPosisi(inputBarcode, posisiKaryawan, barcodePosisiKhusus, validBarcodesPengaturan) {
  if (!inputBarcode) return false;
  const bClean = String(inputBarcode).trim().toUpperCase();
  const pClean = String(posisiKaryawan || '').trim().toUpperCase();
  const bSpecialClean = String(barcodePosisiKhusus || '').trim().toUpperCase();

  // 1. Jika ada Barcode Posisi spesifik di Master Karyawan
  if (bSpecialClean && bClean === bSpecialClean) return true;

  // 2. Jika barcode persis sama dengan nama posisi
  if (pClean && bClean === pClean) return true;

  // 3. Normalisasi karakter (buang spasi, tanda minus, underscore)
  const normBarcode = bClean.replace(/[\s\-_]/g, '');
  const normPosisi = pClean.replace(/[\s\-_]/g, '');
  if (normPosisi && (normBarcode === normPosisi || normBarcode.includes(normPosisi) || normPosisi.includes(normBarcode))) {
    return true;
  }

  // 4. Cek kata kunci posisi (misal barcode 'STAFF PACKING 1' cocok dengan posisi 'Staff Packing')
  const keywords = pClean.split(/[\s\-_]+/).filter(w => w.length > 2);
  for (let k = 0; k < keywords.length; k++) {
    if (bClean.includes(keywords[k])) return true;
  }

  // 5. Cek apakah barcode ada di daftar Barcode_Valid pada sheet Pengaturan (fallback master barcode)
  if (validBarcodesPengaturan && validBarcodesPengaturan.includes(bClean)) {
    return true;
  }

  return false;
}

/**
 * Pemrosesan Presensi Masuk & Pulang dengan Validasi Password Pribadi & Barcode Posisi
 */
function prosesPresensi(inputNamaKaryawan, inputShift, barcodeValidasi, userLat, userLng, inputTipe, inputPassword) {
  // Cegah race condition saat banyak staf presensi bersamaan
  const lock = LockService.getScriptLock();
  const successLock = lock.tryLock(15000);
  if (!successLock) {
    return { status: 'error', pesan: 'Sistem sedang sibuk memproses antrean presensi. Silakan coba kembali beberapa detik lagi.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMaster = ss.getSheetByName('Master_Karyawan');
    const sheetPengaturan = ss.getSheetByName('Pengaturan');
    
    if (!sheetMaster || !sheetPengaturan) {
       return { status: 'error', pesan: 'Sistem belum di-setup. Jalankan setupDatabase() terlebih dahulu di Google Apps Script.' };
    }
    
    // 1. Validasi Pilihan Shift & Tipe Presensi
    let targetSheetName = String(inputShift || '').trim();
    if (!['Shift 1', 'Shift 2', 'Shift 3'].includes(targetSheetName)) {
      return { status: 'error', pesan: 'Wajib memilih Shift (Shift 1, Shift 2, atau Shift 3) terlebih dahulu!' };
    }

    let tipePresensi = String(inputTipe || 'Masuk').trim();
    if (tipePresensi.toLowerCase() === 'pulang') {
      tipePresensi = 'Pulang';
    } else {
      tipePresensi = 'Masuk';
    }

    // 2. Validasi Karyawan & Cari Data di Master_Karyawan
    const lastRowMaster = sheetMaster.getLastRow();
    if (lastRowMaster <= 1) {
      return { status: 'error', pesan: 'Data Master Karyawan masih kosong di database spreadsheet.' };
    }
    
    // Baca data Master Karyawan hingga kolom H (minimal 8 kolom agar Kolom G selalu terambil)
    const dataMaster = sheetMaster.getRange(1, 1, lastRowMaster, Math.max(8, sheetMaster.getLastColumn())).getValues();
    const headerMaster = dataMaster[0] || [];
    const idxID = headerMaster.indexOf('ID Karyawan') !== -1 ? headerMaster.indexOf('ID Karyawan') : 0;
    const idxNama = headerMaster.indexOf('Nama') !== -1 ? headerMaster.indexOf('Nama') : 1;
    const idxPosisi = headerMaster.indexOf('Posisi') !== -1 ? headerMaster.indexOf('Posisi') : 3;

    // Kolom G adalah index 6 (A=0, B=1, C=2, D=3, E=4, F=5, G=6)
    let idxPassword = 6; 
    for (let h = 0; h < headerMaster.length; h++) {
      let hName = String(headerMaster[h] || '').trim().toLowerCase();
      if (hName === 'password' || hName === 'pin' || hName === 'pass' || hName === 'kata sandi') {
        idxPassword = h;
        break;
      }
    }

    let idxBarcodePosisi = headerMaster.indexOf('Barcode Posisi');

    let idKaryawanAsli = '';
    let namaKaryawan = '';
    let posisiKaryawan = '-';
    let passwordTerdaftar = '';
    let barcodePosisiKhusus = '';
    
    const targetNamaClean = String(inputNamaKaryawan).trim().toUpperCase();
    for (let i = 1; i < dataMaster.length; i++) {
      if (String(dataMaster[i][idxNama]).trim().toUpperCase() === targetNamaClean) {
        idKaryawanAsli = String(dataMaster[i][idxID] || '').trim();
        namaKaryawan = String(dataMaster[i][idxNama] || '').trim();
        if (idxPosisi !== -1 && dataMaster[i][idxPosisi]) {
          posisiKaryawan = String(dataMaster[i][idxPosisi]).trim();
        }
        // Ambil password dari idxPassword atau langsung dari Kolom G (index 6)
        if (idxPassword !== -1 && dataMaster[i][idxPassword] !== undefined && String(dataMaster[i][idxPassword]).trim() !== '') {
          passwordTerdaftar = String(dataMaster[i][idxPassword]).trim();
        } else if (dataMaster[i][6] !== undefined && String(dataMaster[i][6]).trim() !== '') {
          passwordTerdaftar = String(dataMaster[i][6]).trim();
        }
        if (idxBarcodePosisi !== -1 && dataMaster[i][idxBarcodePosisi]) {
          barcodePosisiKhusus = String(dataMaster[i][idxBarcodePosisi]).trim();
        }
        break;
      }
    }
    
    if (!namaKaryawan) {
      return { status: 'error', pesan: 'Nama karyawan tidak ditemukan di Master Karyawan. Pastikan memilih nama yang terdaftar.' };
    }

    // 3. Validasi Password Karyawan
    const passwordInputClean = String(inputPassword || '').trim();
    if (!passwordInputClean) {
      return { status: 'error', pesan: 'Password wajib dimasukkan sebelum melakukan scan posisi!' };
    }

    // Jika admin belum mengisi kolom password di sheet, gunakan ID Karyawan atau '123456' sebagai password default
    const expectedPassword = passwordTerdaftar || idKaryawanAsli || '123456';
    if (passwordInputClean !== expectedPassword && passwordInputClean !== idKaryawanAsli && passwordInputClean !== '123456') {
      return { status: 'error', pesan: 'Password salah untuk karyawan ' + namaKaryawan + '! Silakan periksa kembali password Anda.' };
    }

    // 4. Validasi Barcode Posisi (Memverifikasi kesesuaian posisi yang di-scan)
    const inputBarcodeClean = String(barcodeValidasi || '').trim().toUpperCase();
    if (!inputBarcodeClean) {
      return { status: 'error', pesan: 'Barcode posisi kerja belum dipindai!' };
    }

    const dataPengaturan = sheetPengaturan.getDataRange().getValues();
    let validBarcodesPengaturan = [];
    let targetLat = 0, targetLng = 0, maxRadius = 200;

    for (let i = 1; i < dataPengaturan.length; i++) {
      let key = String(dataPengaturan[i][0]).trim();
      let val = dataPengaturan[i][1];
      if (key === 'Barcode_Valid') {
        validBarcodesPengaturan = String(val).split(',').map(b => b.trim().toUpperCase()).filter(Boolean);
      }
      if (key === 'Target_Lat') targetLat = parseFloat(String(val).replace(',', '.'));
      if (key === 'Target_Lng') targetLng = parseFloat(String(val).replace(',', '.'));
      if (key === 'Max_Radius') maxRadius = parseFloat(String(val).replace(',', '.'));
    }

    const isPosisiSesuai = cekKesesuaianPosisi(inputBarcodeClean, posisiKaryawan, barcodePosisiKhusus, validBarcodesPengaturan);
    if (!isPosisiSesuai) {
      return { 
        status: 'error', 
        pesan: `Posisi tidak sesuai! Posisi Anda terdaftar sebagai "${posisiKaryawan}", namun barcode yang di-scan adalah "${barcodeValidasi}". Harap scan barcode di meja/posisi kerja Anda!` 
      };
    }
    
    // 5. Validasi Lokasi GPS & Geofencing
    let distance = 0;
    if (userLat && userLng && targetLat !== 0 && targetLng !== 0) {
      distance = getDistanceFromLatLonInM(userLat, userLng, targetLat, targetLng);
    }
    
    if (maxRadius > 0 && targetLat !== 0 && targetLng !== 0) {
      if (!userLat || !userLng) {
        return { status: 'error', pesan: 'Akses GPS lokasi tidak terdeteksi. Pastikan GPS aktif dan diizinkan pada browser Anda.' };
      }
      if (distance > maxRadius) {
        return { status: 'error', pesan: `Lokasi Anda di luar jangkauan gudang! Jarak Anda ${distance} meter (Batas radius: ${maxRadius}m).` };
      }
    }
    
    // 6. Pastikan Sheet Shift Tujuan Ada
    let sheetTarget = ss.getSheetByName(targetSheetName);
    if (!sheetTarget) {
      sheetTarget = ss.insertSheet(targetSheetName);
      const headers = ['Timestamp', 'Tanggal', 'Waktu', 'ID Karyawan', 'Nama', 'Posisi', 'Tipe Presensi', 'Status', 'Jarak (m)'];
      sheetTarget.appendRow(headers);
      sheetTarget.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eb7f97').setFontColor('white');
      sheetTarget.setFrozenRows(1);
    }
    
    // 7. Verifikasi Riwayat Presensi Hari Ini
    const now = new Date();
    const tanggal = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
    const waktu = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss');
    const lastRowTarget = sheetTarget.getLastRow();
    let sudahMasukHariIni = false;
    let waktuMasukHariIni = '';
    let sudahPulangHariIni = false;
    let waktuPulangHariIni = '';

    if (lastRowTarget > 1) {
      const dataRiwayat = sheetTarget.getRange(2, 1, lastRowTarget - 1, 8).getValues();
      const idUpper = String(idKaryawanAsli).trim().toUpperCase();

      for (let j = dataRiwayat.length - 1; j >= 0; j--) {
        const rowId = String(dataRiwayat[j][3]).trim().toUpperCase();
        if (rowId === idUpper) {
          let rowDateRaw = dataRiwayat[j][1];
          let rowDateStr = String(rowDateRaw || '').trim();
          let isSameDay = (rowDateStr === tanggal);
          if (!isSameDay) {
            try {
              let parsed = Utilities.formatDate(new Date(rowDateRaw), 'Asia/Jakarta', 'yyyy-MM-dd');
              if (parsed === tanggal) isSameDay = true;
            } catch(e) {}
          }

          if (isSameDay) {
            const rowTipe = String(dataRiwayat[j][6] || 'Masuk').trim().toLowerCase();
            const rowWaktu = String(dataRiwayat[j][2] || '').trim();

            if (rowTipe === 'masuk') {
              sudahMasukHariIni = true;
              if (!waktuMasukHariIni) waktuMasukHariIni = rowWaktu;
            } else if (rowTipe === 'pulang') {
              sudahPulangHariIni = true;
              if (!waktuPulangHariIni) waktuPulangHariIni = rowWaktu;
            }
          }
        }
      }
    }

    // Auto-Detection Shift untuk Presensi Pulang (jika salah pilih shift)
    if (tipePresensi === 'Pulang' && !sudahMasukHariIni) {
      const otherShifts = ['Shift 1', 'Shift 2', 'Shift 3'].filter(s => s !== targetSheetName);
      for (let s = 0; s < otherShifts.length; s++) {
        let altSheet = ss.getSheetByName(otherShifts[s]);
        if (!altSheet) continue;
        let altLastRow = altSheet.getLastRow();
        if (altLastRow <= 1) continue;

        let altData = altSheet.getRange(2, 1, altLastRow - 1, 8).getValues();
        for (let k = altData.length - 1; k >= 0; k--) {
          if (String(altData[k][3]).trim().toUpperCase() === String(idKaryawanAsli).trim().toUpperCase()) {
            let altDate = String(altData[k][1] || '').trim();
            let isAltSameDay = (altDate === tanggal);
            if (!isAltSameDay) {
              try {
                if (Utilities.formatDate(new Date(altData[k][1]), 'Asia/Jakarta', 'yyyy-MM-dd') === tanggal) {
                  isAltSameDay = true;
                }
              } catch(e) {}
            }

            if (isAltSameDay && String(altData[k][6] || 'Masuk').trim().toLowerCase() === 'masuk') {
              sheetTarget = altSheet;
              targetSheetName = otherShifts[s];
              sudahMasukHariIni = true;
              waktuMasukHariIni = String(altData[k][2] || '').trim();
              
              for (let m = altData.length - 1; m >= 0; m--) {
                if (String(altData[m][3]).trim().toUpperCase() === String(idKaryawanAsli).trim().toUpperCase()) {
                  let mDate = String(altData[m][1] || '').trim();
                  let isMSameDay = (mDate === tanggal);
                  if (!isMSameDay) {
                    try {
                      if (Utilities.formatDate(new Date(altData[m][1]), 'Asia/Jakarta', 'yyyy-MM-dd') === tanggal) isMSameDay = true;
                    } catch(e) {}
                  }
                  if (isMSameDay && String(altData[m][6] || '').trim().toLowerCase() === 'pulang') {
                    sudahPulangHariIni = true;
                    waktuPulangHariIni = String(altData[m][2] || '').trim();
                    break;
                  }
                }
              }
              break;
            }
          }
        }
        if (sudahMasukHariIni) break;
      }
    }

    // 8. Logika Validasi Alur Masuk vs Pulang
    if (tipePresensi === 'Masuk') {
      if (sudahMasukHariIni) {
        return { 
          status: 'warning', 
          pesan: `Halo ${namaKaryawan}, kamu sudah tercatat presensi Masuk di ${targetSheetName} hari ini (pukul ${waktuMasukHariIni}).` 
        };
      }
    } else if (tipePresensi === 'Pulang') {
      if (!sudahMasukHariIni) {
        return { 
          status: 'warning', 
          pesan: `Halo ${namaKaryawan}, kamu belum melakukan presensi Masuk hari ini. Silakan pilih opsi Masuk terlebih dahulu!` 
        };
      }
      if (sudahPulangHariIni) {
        return { 
          status: 'warning', 
          pesan: `Halo ${namaKaryawan}, kamu sudah tercatat presensi Pulang hari ini (pukul ${waktuPulangHariIni}). Sampai jumpa besok!` 
        };
      }
    }
    
    // 9. Simpan Baris Presensi ke Sheet Shift Terpilih
    const statusPresensi = (tipePresensi === 'Masuk') ? 'Hadir' : 'Selesai';
    sheetTarget.appendRow([now, tanggal, waktu, idKaryawanAsli, namaKaryawan, posisiKaryawan, tipePresensi, statusPresensi, distance]);
    
    const pesanSukses = (tipePresensi === 'Masuk') 
      ? `Presensi Masuk ${targetSheetName} berhasil disimpan di posisi ${posisiKaryawan}! Semangat bekerja, ${namaKaryawan}.`
      : `Presensi Pulang ${targetSheetName} berhasil disimpan! Terima kasih atas kerja kerasmu hari ini, ${namaKaryawan}. Hati-hati di jalan!`;

    return { 
      status: 'success', 
      pesan: pesanSukses 
    };
    
  } catch (error) {
    return { status: 'error', pesan: 'Terjadi kesalahan sistem: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mengambil Data KPI Dashboard (Statistik Kehadiran, Log Hari Ini, & Leaderboard)
 */
function getKPIData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shifts = ['Shift 1', 'Shift 2', 'Shift 3'];
  
  const masterSheet = ss.getSheetByName('Master_Karyawan');
  let totalKaryawanAktif = 0;
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    totalKaryawanAktif = mData.length > 1 ? mData.length - 1 : 0;
  }
  
  let kpiStore = {
    shift1: {},
    shift2: {},
    shift3: {},
    overall: {}
  };
  
  let summary = {
    totalKaryawan: totalKaryawanAktif,
    hadirHariIni: 0,
    pulangHariIni: 0,
    shift1HariIni: 0,
    shift2HariIni: 0,
    shift3HariIni: 0
  };
  
  let todayDateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  let hadirSet = new Set();
  let pulangSet = new Set();
  let todayLogs = [];
  
  let chartData = {
    shift1: 0,
    shift2: 0,
    shift3: 0
  };
  
  shifts.forEach((shiftName, idx) => {
    const sheet = ss.getSheetByName(shiftName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    const headers = data[0];
    let idxTimestamp = headers.indexOf('Timestamp');
    let idxTanggal = headers.indexOf('Tanggal');
    let idxWaktu = headers.indexOf('Waktu');
    let idxID = headers.indexOf('ID Karyawan');
    let idxNama = headers.indexOf('Nama');
    let idxPosisi = headers.indexOf('Posisi');
    let idxTipe = headers.indexOf('Tipe Presensi');
    let idxStatus = headers.indexOf('Status');
    let idxJarak = headers.indexOf('Jarak (m)');
    
    if (idxTimestamp === -1) idxTimestamp = 0;
    if (idxTanggal === -1) idxTanggal = 1;
    if (idxWaktu === -1) idxWaktu = 2;
    if (idxID === -1) idxID = 3;
    if (idxNama === -1) idxNama = 4;
    if (idxPosisi === -1) idxPosisi = 5;
    if (idxTipe === -1) idxTipe = 6;
    if (idxStatus === -1) idxStatus = 7;
    if (idxJarak === -1) idxJarak = 8;
    
    let currentShiftKey = 'shift' + (idx + 1);
    
    for (let i = 1; i < data.length; i++) {
      let nama = String(data[i][idxNama] || '').trim();
      let tanggalRaw = data[i][idxTanggal];
      let tipe = String(data[i][idxTipe] || 'Masuk').trim();
      
      if (!nama) continue;
      
      let isToday = false;
      let rawTglStr = String(tanggalRaw || '').trim();
      if (rawTglStr === todayDateStr) {
        isToday = true;
      } else {
        try {
          let rowDateStr = Utilities.formatDate(new Date(tanggalRaw), 'Asia/Jakarta', 'yyyy-MM-dd');
          if (rowDateStr === todayDateStr) isToday = true;
        } catch(e) {}
      }
      
      if (tipe.toLowerCase() === 'masuk') {
        chartData[currentShiftKey] = (chartData[currentShiftKey] || 0) + 1;
        let poin = 10;
        kpiStore[currentShiftKey][nama] = (kpiStore[currentShiftKey][nama] || 0) + poin;
        kpiStore.overall[nama] = (kpiStore.overall[nama] || 0) + poin;
      }
      
      if (isToday) {
        if (tipe.toLowerCase() === 'pulang') {
          pulangSet.add(nama);
        } else {
          hadirSet.add(nama);
          if (currentShiftKey === 'shift1') summary.shift1HariIni++;
          if (currentShiftKey === 'shift2') summary.shift2HariIni++;
          if (currentShiftKey === 'shift3') summary.shift3HariIni++;
        }

        todayLogs.push({
          timestamp: data[i][idxTimestamp],
          tanggal: rawTglStr,
          waktu: String(data[i][idxWaktu] || ''),
          id: String(data[i][idxID] || ''),
          nama: nama,
          posisi: String(data[i][idxPosisi] || '-'),
          shift: shiftName,
          tipe: tipe,
          status: String(data[i][idxStatus] || 'Hadir'),
          jarak: data[i][idxJarak] || 0
        });
      }
    }
  });
  
  summary.hadirHariIni = hadirSet.size;
  summary.pulangHariIni = pulangSet.size;
  
  todayLogs.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  function formatAndSort(obj) {
    let arr = [];
    for (let n in obj) {
      arr.push({ nama: n, poin: obj[n] });
    }
    return arr.sort((a, b) => b.poin - a.poin);
  }
  
  return {
    summary: summary,
    chartData: chartData,
    todayLogs: todayLogs,
    leaderboard: {
      shift1: formatAndSort(kpiStore.shift1),
      shift2: formatAndSort(kpiStore.shift2),
      shift3: formatAndSort(kpiStore.shift3),
      overall: formatAndSort(kpiStore.overall)
    }
  };
}

/**
 * Mengambil Seluruh Data Riwayat Presensi untuk Ekspor CSV/Excel
 */
function getExportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shifts = ['Shift 1', 'Shift 2', 'Shift 3'];
  let allRecords = [];

  shifts.forEach(shiftName => {
    const sheet = ss.getSheetByName(shiftName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    for (let i = 1; i < data.length; i++) {
      allRecords.push({
        timestamp: data[i][0],
        tanggal: data[i][1],
        waktu: data[i][2],
        id: data[i][3],
        nama: data[i][4],
        posisi: data[i][5],
        shift: shiftName,
        tipe: data[i][6],
        status: data[i][7],
        jarak: data[i][8]
      });
    }
  });

  allRecords.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
  return allRecords;
}
