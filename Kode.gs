function doGet(e) {
  // 1. Endpoint JSON API untuk Frontend GitHub Pages (Metode GET)
  if (e && e.parameter && e.parameter.action) {
    let output = {};
    if (e.parameter.action === 'getKaryawan') {
      output = getDaftarKaryawan();
    } else if (e.parameter.action === 'getKPIData') {
      output = getKPIData();
    }
    return ContentService.createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Tampilan Web App bawaan jika dibuka langsung di Google Apps Script
  // Jika parameter ?page=index diberikan secara eksplisit, baru tampilkan form presensi
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
 * Endpoint JSON API untuk Menerima Data Presensi dari GitHub Pages (Metode POST)
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch(errJson) {
        // Fallback jika dikirim via URL-encoded form
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const nama = payload.nama || payload.inputNamaKaryawan || '';
    const shift = payload.shift || payload.inputShift || '';
    const barcode = payload.barcode || payload.barcodeValidasi || '';
    const lat = parseFloat(payload.lat || payload.userLat || 0);
    const lng = parseFloat(payload.lng || payload.userLng || 0);

    const result = prosesPresensi(nama, shift, barcode, lat, lng);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      pesan: 'Terjadi kesalahan pada server: ' + error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// JALANKAN FUNGSI INI SEKALI SAJA UNTUK MENG-GENERATE HEADER SPREADSHEET
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Sheet Shift 1, Shift 2, Shift 3
  const shifts = ['Shift 1', 'Shift 2', 'Shift 3'];
  shifts.forEach(shiftName => {
    let sheet = ss.getSheetByName(shiftName);
    if (!sheet) {
      sheet = ss.insertSheet(shiftName);
    }
    // Paksa update header ke versi terbaru (Menambahkan Posisi)
    const headers = ['Timestamp', 'Tanggal', 'Waktu', 'ID Karyawan', 'Nama', 'Posisi', 'Tipe Presensi', 'Waktu Masuk', 'Jarak (m)'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eb7f97').setFontColor('white');
    sheet.setFrozenRows(1);
  });
  
  // 2. Setup Sheet Master Karyawan
  let sheetMaster = ss.getSheetByName('Master_Karyawan');
  if (!sheetMaster) {
    sheetMaster = ss.insertSheet('Master_Karyawan');
    sheetMaster.appendRow(['ID Karyawan', 'Nama', 'Departemen', 'Posisi', 'Shift Assigned']);
    sheetMaster.appendRow(['LOZY001', 'Fulanah', 'People Ops', 'Staff Packing', 'Shift 1']);
  } else {
    // Paksa update header ke versi terbaru (Menambahkan Posisi di antara Departemen dan Shift)
    sheetMaster.getRange('A1:E1').setValues([['ID Karyawan', 'Nama', 'Departemen', 'Posisi', 'Shift Assigned']]);
  }
  sheetMaster.getRange('A1:E1').setFontWeight('bold').setBackground('#4b3040').setFontColor('white');
  sheetMaster.setFrozenRows(1);
  
  // 3. Setup Sheet Pengaturan
  let sheetPengaturan = ss.getSheetByName('Pengaturan');
  if (!sheetPengaturan) {
    sheetPengaturan = ss.insertSheet('Pengaturan');
    sheetPengaturan.appendRow(['Kunci', 'Nilai', 'Keterangan']);
    sheetPengaturan.getRange('A1:C1').setFontWeight('bold').setBackground('#766770').setFontColor('white');
    sheetPengaturan.appendRow(['Barcode_Valid', 'LOZY123', 'Satu barcode ini untuk semua shift']);
    sheetPengaturan.appendRow(['Target_Lat', '-6.9174639', 'Garis lintang lokasi absensi']);
    sheetPengaturan.appendRow(['Target_Lng', '107.6191228', 'Garis bujur lokasi absensi']);
    sheetPengaturan.appendRow(['Max_Radius', '200', 'Jarak maksimal dalam meter (Saran: 200-300 untuk indoor)']);
    sheetPengaturan.setFrozenRows(1);
    sheetPengaturan.setColumnWidths(1, 3, 200);
  }
}

// Menghitung jarak (Haversine formula)
function getDistanceFromLatLonInM(lat1, lon1, lat2, lon2) {
  var R = 6371000; 
  var dLat = (lat2-lat1) * (Math.PI/180);
  var dLon = (lon2-lon1) * (Math.PI/180); 
  var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
          Math.cos(lat1 * (Math.PI/180)) * Math.cos(lat2 * (Math.PI/180)) * 
          Math.sin(dLon/2) * Math.sin(dLon/2); 
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)); 
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
    
    const data = sheetMaster.getRange(2, 2, lastRow - 1, 1).getValues();
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
 * Pemrosesan Presensi Masuk (Check-In) dengan Shift yang Dipilih oleh Karyawan
 */
function prosesPresensi(inputNamaKaryawan, inputShift, barcodeValidasi, userLat, userLng) {
  // Cegah race condition saat banyak karyawan presensi bersamaan
  const lock = LockService.getScriptLock();
  const successLock = lock.tryLock(15000); // Tunggu antrean hingga 15 detik
  if (!successLock) {
    return { status: 'error', pesan: 'Sistem sedang sibuk memproses presensi lain. Silakan coba kembali sesaat lagi.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMaster = ss.getSheetByName('Master_Karyawan');
    const sheetPengaturan = ss.getSheetByName('Pengaturan');
    
    if (!sheetMaster || !sheetPengaturan) {
       return { status: 'error', pesan: 'Sistem belum di-setup. Jalankan setupDatabase() terlebih dahulu.' };
    }
    
    // 1. Validasi Pilihan Shift
    const targetSheetName = String(inputShift || '').trim();
    if (!['Shift 1', 'Shift 2', 'Shift 3'].includes(targetSheetName)) {
      return { status: 'error', pesan: 'Wajib memilih Shift (Shift 1, Shift 2, atau Shift 3) terlebih dahulu!' };
    }
    
    // 2. Cek Pengaturan (Barcode & Lokasi Geofencing)
    const dataPengaturan = sheetPengaturan.getDataRange().getValues();
    let validBarcode = '', targetLat = 0, targetLng = 0, maxRadius = 200;
    for (let i = 1; i < dataPengaturan.length; i++) {
      let key = String(dataPengaturan[i][0]).trim();
      let val = dataPengaturan[i][1];
      if (key === 'Barcode_Valid') validBarcode = String(val).trim();
      if (key === 'Target_Lat') targetLat = parseFloat(String(val).replace(',', '.'));
      if (key === 'Target_Lng') targetLng = parseFloat(String(val).replace(',', '.'));
      if (key === 'Max_Radius') maxRadius = parseFloat(String(val).replace(',', '.'));
    }
    
    // Validasi Barcode
    if (String(barcodeValidasi).trim() !== validBarcode) {
      return { status: 'error', pesan: 'Barcode Validasi tidak cocok atau salah!' };
    }
    
    // Validasi Lokasi GPS & Geofencing
    let distance = 0;
    if (userLat && userLng && targetLat !== 0 && targetLng !== 0) {
      distance = getDistanceFromLatLonInM(userLat, userLng, targetLat, targetLng);
    }
    
    // Geofencing aktif jika koordinat target dan maxRadius > 0
    if (maxRadius > 0 && targetLat !== 0 && targetLng !== 0) {
      if (!userLat || !userLng) {
        return { status: 'error', pesan: 'Akses lokasi tidak ditemukan. Pastikan GPS aktif dan diizinkan pada perangkat Anda.' };
      }
      if (distance > maxRadius) {
        return { status: 'error', pesan: `Lokasi Anda di luar jangkauan! Jarak Anda ${distance} meter (Maksimal radius: ${maxRadius}m).` };
      }
    }
    
    // 3. Validasi Karyawan & Ambil Posisi dari Master_Karyawan
    const lastRowMaster = sheetMaster.getLastRow();
    if (lastRowMaster <= 1) {
      return { status: 'error', pesan: 'Data Master Karyawan masih kosong di spreadsheet.' };
    }
    
    const dataMaster = sheetMaster.getRange(1, 1, lastRowMaster, Math.max(5, sheetMaster.getLastColumn())).getValues();
    let idKaryawanAsli = '';
    let namaKaryawan = '';
    let posisiKaryawan = '-';
    
    const headerMaster = dataMaster[0] || [];
    const idxPosisi = headerMaster.indexOf('Posisi');
    
    const targetNamaClean = String(inputNamaKaryawan).trim().toUpperCase();
    for (let i = 1; i < dataMaster.length; i++) {
      if (String(dataMaster[i][1]).trim().toUpperCase() === targetNamaClean) {
        idKaryawanAsli = dataMaster[i][0];
        namaKaryawan = dataMaster[i][1];
        if (idxPosisi !== -1 && dataMaster[i][idxPosisi]) {
          posisiKaryawan = String(dataMaster[i][idxPosisi]).trim();
        }
        break;
      }
    }
    
    if (!namaKaryawan) {
      return { status: 'error', pesan: 'Nama karyawan tidak ditemukan di Master Karyawan. Pastikan nama sesuai dengan yang terdaftar.' };
    }
    
    // 4. Pastikan Sheet Shift Tujuan Ada
    let sheetTarget = ss.getSheetByName(targetSheetName);
    if (!sheetTarget) {
      sheetTarget = ss.insertSheet(targetSheetName);
      const headers = ['Timestamp', 'Tanggal', 'Waktu', 'ID Karyawan', 'Nama', 'Posisi', 'Tipe Presensi', 'Waktu Masuk', 'Jarak (m)'];
      sheetTarget.appendRow(headers);
      sheetTarget.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eb7f97').setFontColor('white');
      sheetTarget.setFrozenRows(1);
    }
    
    const now = new Date();
    const tanggal = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
    const waktu = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss');
    
    // 5. Cek Duplikasi / Anti Spam pada Sheet Shift Tersebut
    const lastRowTarget = sheetTarget.getLastRow();
    if (lastRowTarget > 1) {
      let startRow = Math.max(2, lastRowTarget - 100);
      let numRows = lastRowTarget - startRow + 1;
      let data = sheetTarget.getRange(startRow, 1, numRows, 9).getValues();
      
      for (let j = data.length - 1; j >= 0; j--) {
        // Kolom index 3 adalah ID Karyawan (D)
        if (String(data[j][3]).trim().toUpperCase() === String(idKaryawanAsli).trim().toUpperCase()) {
          let rowTimestamp = new Date(data[j][0]).getTime();
          let diffMinutes = (now.getTime() - rowTimestamp) / 60000;
          
          // Cegah double klik / scan bertubi-tubi dalam 5 menit
          if (diffMinutes < 5) {
            return { 
              status: 'warning', 
              pesan: `Halo ${namaKaryawan}, kamu baru saja melakukan presensi di ${targetSheetName}.` 
            };
          }
          
          // Cek tanggal presensi sebelumnya pada shift ini
          let rowDateRaw = data[j][1];
          let rowDateStr = String(rowDateRaw || '').trim();
          let isSameDay = (rowDateStr === tanggal);
          if (!isSameDay) {
            try {
              let parsed = Utilities.formatDate(new Date(rowDateRaw), 'Asia/Jakarta', 'yyyy-MM-dd');
              if (parsed === tanggal) isSameDay = true;
            } catch(e) {}
          }
          
          if (isSameDay) {
            let prevWaktu = data[j][2] || '';
            return { 
              status: 'warning', 
              pesan: `Halo ${namaKaryawan}, kamu sudah tercatat presensi di ${targetSheetName} hari ini (pukul ${prevWaktu}).` 
            };
          }
          break;
        }
      }
    }
    
    // 6. Status Kehadiran Fleksibel (Tidak dipengaruhi oleh jam shift, jam berapapun presensi diakui Hadir)
    const statusPresensi = 'Hadir';
    
    // 7. Simpan Baris Presensi ke Sheet Shift Pilihan Karyawan
    sheetTarget.appendRow([now, tanggal, waktu, idKaryawanAsli, namaKaryawan, posisiKaryawan, 'Masuk', statusPresensi, distance]);
    
    return { 
      status: 'success', 
      pesan: `Presensi ${targetSheetName} berhasil disimpan! Semangat bekerja, ${namaKaryawan}.` 
    };
    
  } catch (error) {
    return { status: 'error', pesan: 'Terjadi kesalahan sistem: ' + error.message };
  } finally {
    lock.releaseLock();
  }
}

// FUNGSI UNTUK MENGAMBIL DATA KPI DASHBOARD (FOKUS KEHADIRAN & DISTRIBUSI SHIFT FLEKSIBEL)
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
    shift1HariIni: 0,
    shift2HariIni: 0,
    shift3HariIni: 0
  };
  
  let todayDateStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  let hadirSet = new Set();
  
  let chartData = {
    shift1: 0,
    shift2: 0,
    shift3: 0
  };
  
  shifts.forEach((shiftName, idx) => {
    const sheet = ss.getSheetByName(shiftName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // cuma header
    
    const headers = data[0];
    let idxNama = headers.indexOf('Nama');
    let idxTanggal = headers.indexOf('Tanggal');
    
    // Fallback index jika format kolom bergeser
    if (idxNama === -1) idxNama = 4;
    if (idxTanggal === -1) idxTanggal = 1;
    
    let currentShiftKey = 'shift' + (idx + 1);
    
    for (let i = 1; i < data.length; i++) {
      let nama = String(data[i][idxNama]).trim();
      let tanggalRaw = data[i][idxTanggal];
      
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
      
      // Akumulasi total kehadiran per shift untuk grafik distribusi
      chartData[currentShiftKey] = (chartData[currentShiftKey] || 0) + 1;
      
      if (isToday) {
        hadirSet.add(nama);
        if (currentShiftKey === 'shift1') summary.shift1HariIni++;
        if (currentShiftKey === 'shift2') summary.shift2HariIni++;
        if (currentShiftKey === 'shift3') summary.shift3HariIni++;
      }
      
      // Poin kehadiran fleksibel (10 poin tiap kehadiran)
      let poin = 10;
      kpiStore[currentShiftKey][nama] = (kpiStore[currentShiftKey][nama] || 0) + poin;
      kpiStore.overall[nama] = (kpiStore.overall[nama] || 0) + poin;
    }
  });
  
  summary.hadirHariIni = hadirSet.size;
  
  // Format dan sort data
  function formatAndSort(obj) {
    let arr = [];
    for (let n in obj) {
      arr.push({ nama: n, poin: obj[n] });
    }
    // Sort descending (poin / jumlah kehadiran tertinggi)
    return arr.sort((a, b) => b.poin - a.poin);
  }
  
  return {
    summary: summary,
    chartData: chartData,
    leaderboard: {
      shift1: formatAndSort(kpiStore.shift1),
      shift2: formatAndSort(kpiStore.shift2),
      shift3: formatAndSort(kpiStore.shift3),
      overall: formatAndSort(kpiStore.overall)
    }
  };
}