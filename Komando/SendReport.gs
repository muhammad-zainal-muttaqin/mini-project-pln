// Debug override: send all messages to this number for testing
var DEBUG_OVERRIDE = false;
var DEBUG_PHONE_RAW = '087778651293';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Send Report 📄')
    .addItem('Send Report 📩', 'sendReport')
    .addItem('Test Send Report 🧪', 'testSendReport')
    .addToUi();
}

function sendReport() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Tidak dapat memperoleh lock. Proses sudah berjalan.");
    return;
  }

  try {
    // Ambil semua sheet yang diperlukan
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName('REPORT02');
    var msgSheet = ss.getSheetByName('MSGSENDER');
    var logSheet = ss.getSheetByName('LOG');
    
    if (!reportSheet) {
      Logger.log("Sheet 'REPORT02' tidak ditemukan!");
      return;
    }
    if (!msgSheet) {
      Logger.log("Sheet 'MSGSENDER' tidak ditemukan!");
      return;
    }
    if (!logSheet) {
      Logger.log("Sheet 'LOG' tidak ditemukan!");
      return;
    }

    // Update timestamp di sheet sebelum screenshot
    reportSheet.getRange('C5').setValue(new Date());
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    // Buat nama file berdasarkan tanggal
    var reportDateObj = reportSheet.getRange('C3').getValue();
    var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
    var now = new Date();
    var genDateTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var fileName = 'KOMANDO_' + reportDateStr + '_' + genDateTimeStr + '.png';

    // PENTING: Buat screenshot dua kali untuk memastikan identik
    // Screenshot pertama untuk WhatsApp
    var blobForWhatsApp = createReportScreenshot(reportSheet);
    if (!blobForWhatsApp) {
      Logger.log("Gagal membuat screenshot untuk WhatsApp");
      return;
    }
    blobForWhatsApp.setName(fileName);
    
    // Screenshot kedua untuk Drive
    var blobForDrive = createReportScreenshot(reportSheet);
    if (!blobForDrive) {
      Logger.log("Gagal membuat screenshot untuk Drive");
      return;
    }
    blobForDrive.setName(fileName);
    
    // Upload gambar ke Google Drive
    var publicUrl = uploadImageToDrive(blobForDrive);
    Logger.log("Gambar berhasil diupload ke Drive: " + publicUrl);

    // Ambil data untuk logging
    var logColumnA = reportSheet.getRange('C2').getValue();
    var startDate = reportSheet.getRange('C3').getValue();
    var endDate = reportSheet.getRange('C4').getValue();
    var logColumnB = formatDateRange(startDate, endDate);

    // Ambil token dan data penerima
    var token = msgSheet.getRange('C5').getValue();
    var recipientData = msgSheet.getRange(7, 3, msgSheet.getLastRow() - 6, 3).getValues(); // C7:E

    // Kirim pesan ke setiap penerima
    recipientData.forEach(function(row) {
      var phoneNumber = getTargetPhone(row[0]); // C
      var recipientName = row[1]; // D
      var messageText = row[2]; // E

      // Skip jika data tidak lengkap
      if (!phoneNumber || !messageText || phoneNumber === '' || messageText === '') {
        return;
      }

      try {
        // Setelah uploadImageToDrive()…
        var fileId = publicUrl.match(/id=([^&]+)/)[1];
        var driveBlob = DriveApp.getFileById(fileId).getBlob();
        driveBlob.setName(fileName);
        
        // Kirim pesan dengan gambar dari Drive link sebagai attachment
        var payload = {
          target:   phoneNumber,
          message:  messageText,
          file:     screenshotBlob,   // atau driveBlob kalau mau pakai file hasil upload
          filename: fileName
        };
        
        var options = {
          method: 'post',
          headers: { 'Authorization': token },
          payload: payload,
          muteHttpExceptions: true
        };

        var response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
        Logger.log('Pesan berhasil dikirim ke ' + phoneNumber + ': ' + response.getContentText());
        
        // Log ke sheet LOG
        logSheet.appendRow([logColumnA, logColumnB, phoneNumber, recipientName, messageText, 'SENT', new Date()]);
        
      } catch (error) {
        Logger.log('Error mengirim pesan ke ' + phoneNumber + ': ' + error);
        logSheet.appendRow([logColumnA, logColumnB, phoneNumber, recipientName, messageText, 'FAILED', new Date()]);
      }
      
      // Delay untuk menghindari rate limiting
      Utilities.sleep(3000);
    });

    Logger.log("Proses pengiriman laporan selesai");
    
  } finally {
    lock.releaseLock();
  }
}

function getTargetPhone(rawPhone) {
  return formatPhone(DEBUG_OVERRIDE ? DEBUG_PHONE_RAW : rawPhone);
}

function formatPhone(num) {
  if (!num) return '';
  num = num.toString().trim();
  if (num.startsWith('+628')) {
    return '628' + num.substring(4);
  } else if (num.startsWith('08')) {
    return '628' + num.substring(2);
  } else if (num.startsWith('628')) {
    return num;
  }
  return num;
}

function formatDate(date) {
  if (!date) return '';
  
  // Jika date adalah string, coba konversi ke Date object
  if (typeof date === 'string') {
    date = new Date(date);
  }
  
  // Validasi apakah date adalah Date object yang valid
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return date.toString(); // Return string asli jika tidak bisa dikonversi
  }
  
  var day = date.getDate();
  var month = date.getMonth();
  var year = date.getFullYear();
  
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  
  return day + " " + monthNames[month] + " " + year;
}

function formatDateRange(startDate, endDate) {
  return formatDate(startDate) + " s.d. " + formatDate(endDate);
}

function createReportScreenshot(reportSheet) {
  try {
    // Pastikan semua perubahan di sheet tercommit sebelum screenshot
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    // Cari baris terakhir yang benar-benar berisi data di kolom B (ORG_1)
    // Karena kolom B selalu berisi data untuk setiap baris organisasi
    var lastRow = findLastDataRow(reportSheet);
    
    // PENTING: Hanya ambil sampai kolom G (kolom ke-7) untuk laporan
    // Jangan ambil kolom H dan seterusnya yang berisi daftar pegawai
    var lastCol = 7; // Kolom G = kolom ke-7
    
    // Pastikan range mencakup semua elemen penting:
    // - Header info (B1:C5)
    // - Logo & header summary (B8:G10) 
    // - Header tabel (B12:G12)
    // - Data sampai baris terakhir yang berisi data, tapi hanya sampai kolom G
    var range = reportSheet.getRange(1, 1, lastRow, lastCol);
    
    Logger.log("Range laporan: A1:" + getColumnLetter(lastCol) + lastRow);
    Logger.log("Total rows: " + lastRow + ", Total columns: " + lastCol + " (sampai kolom G saja)");
    
    // Buat chart sebagai tabel dengan range yang tepat
    var chart = reportSheet.newChart()
      .setChartType(Charts.ChartType.TABLE)
      .addRange(range)
      .setPosition(1, 1, 0, 0)
      .setOption('width', 1200)  // Ukuran yang sesuai untuk kolom A-G
      .setOption('height', Math.max(800, lastRow * 25)) // Tinggi dinamis berdasarkan jumlah baris
      .setOption('backgroundColor', 'white')
      .setOption('legend', {position: 'none'})
      .setOption('enableInteractivity', false)
      .setOption('alternatingRowStyle', false) // Nonaktifkan alternating untuk menjaga format asli
      .setOption('allowHtml', true)
      .setOption('showRowNumber', false)
      .setOption('page', 'disable') // Nonaktifkan pagination agar semua data tampil
      .build();
    
    // Ambil chart sebagai gambar
    var chartImage = chart.getBlob();
    
    Logger.log("Screenshot berhasil dibuat, ukuran: " + chartImage.getBytes().length + " bytes");
    return chartImage;
    
  } catch (error) {
    Logger.log("Error saat membuat screenshot: " + error);
    return null;
  }
}

// Helper function untuk convert nomor kolom ke huruf
function getColumnLetter(columnNumber) {
  var columnLetter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

// Helper function untuk mencari baris terakhir yang berisi data
function findLastDataRow(reportSheet) {
  // Ambil semua data di kolom B (ORG_1) dari baris 1 sampai 100
  var columnBData = reportSheet.getRange('B1:B100').getValues();
  
  // Cari dari bawah ke atas untuk menemukan baris terakhir yang tidak kosong
  for (var i = columnBData.length - 1; i >= 0; i--) {
    if (columnBData[i][0] !== '' && columnBData[i][0] !== null && columnBData[i][0] !== undefined) {
      var lastDataRow = i + 1; // +1 karena array dimulai dari 0 tapi baris dimulai dari 1
      Logger.log("Baris terakhir yang berisi data: " + lastDataRow);
      return lastDataRow;
    }
  }
  
  // Jika tidak ditemukan data, return minimal 30 baris untuk safety
  Logger.log("Tidak ditemukan data, menggunakan 30 baris default");
  return 30;
}

function uploadImageToDrive(imageBlob) {
  var folder = DriveApp.getFolderById("1WLXGZeinrbBPt6Si3Qhn7lME9UYinQiT");
  var file = folder.createFile(imageBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Gunakan export=download untuk direct download
  return "https://drive.google.com/uc?export=download&id=" + file.getId();
}

function testSendReport() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Tidak dapat memperoleh lock. Proses sudah berjalan.");
    return;
  }

  try {
    Logger.log("=== MULAI TEST SEND REPORT ===");
    
    // Ambil semua sheet yang diperlukan
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName('REPORT02');
    var msgSheet = ss.getSheetByName('MSGSENDER');
    var logSheet = ss.getSheetByName('LOG');
    
    if (!reportSheet) {
      Logger.log("ERROR: Sheet 'REPORT02' tidak ditemukan!");
      return;
    }
    if (!msgSheet) {
      Logger.log("ERROR: Sheet 'MSGSENDER' tidak ditemukan!");
      return;
    }
    if (!logSheet) {
      Logger.log("ERROR: Sheet 'LOG' tidak ditemukan!");
      return;
    }
    
    Logger.log("✅ Semua sheet berhasil ditemukan");

    // Update timestamp di sheet sebelum screenshot
    reportSheet.getRange('C5').setValue(new Date());
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    // Buat nama file berdasarkan tanggal
    var reportDateObj = reportSheet.getRange('C3').getValue();
    var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
    var now = new Date();
    var genDateTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var fileName = 'TEST_KOMANDO_' + reportDateStr + '_' + genDateTimeStr + '.png';

    // Hapus pembuatan blobForWhatsApp dan blobForDrive
    // Ambil screenshot dan upload ke Drive
    // Ambil satu screenshot
    Logger.log("📸 Membuat screenshot laporan...");
    var screenshotBlob = createReportScreenshot(reportSheet);
    if (!screenshotBlob) {
      Logger.log("ERROR: Gagal membuat screenshot laporan");
      return;
    }
    screenshotBlob.setName(fileName);
    Logger.log("✅ Screenshot berhasil dibuat");

    // Upload gambar ke Google Drive
    Logger.log("☁️ Mengupload screenshot ke Google Drive...");
    var publicUrl = uploadImageToDrive(screenshotBlob);
    Logger.log("✅ Screenshot berhasil diupload ke Drive: " + publicUrl);

    // Ambil data untuk logging
    var logColumnA = reportSheet.getRange('C2').getValue();
    var startDate = reportSheet.getRange('C3').getValue();
    var endDate = reportSheet.getRange('C4').getValue();
    var logColumnB = formatDateRange(startDate, endDate);
    Logger.log("✅ Data logging: " + logColumnA + " | " + logColumnB);

    // Ambil token
    var token = msgSheet.getRange('C5').getValue();
    if (!token) {
      Logger.log("ERROR: Token tidak ditemukan di C5");
      return;
    }
    Logger.log("✅ Token berhasil diambil");

    // Data test - kirim ke nomor test
    var testPhone = '087778651293';
    var testName = 'Test User';
    var testMessage = '🧪 TEST LAPORAN KOMANDO\n\nIni adalah test pengiriman laporan otomatis.\n\nTanggal: ' + formatDateRange(startDate, endDate);
    
    Logger.log("📱 Mengirim test ke: " + testPhone);
    Logger.log("💬 Pesan: " + testMessage);

    try {
      // Format nomor telepon
      var formattedPhone = formatPhone(testPhone);
      Logger.log("📞 Nomor terformat: " + formattedPhone);
      
      // Setelah uploadImageToDrive()…
      // Dapatkan blob langsung dari file Drive yang baru di-upload
      var fileId = publicUrl.match(/id=([^&]+)/)[1];
      var driveBlob = DriveApp.getFileById(fileId).getBlob();
      driveBlob.setName(fileName);
      
      // Kirim pesan dengan gambar file Drive sebagai attachment
      var payload = {
        target: formattedPhone,
        message: testMessage,
        file:    driveBlob,
        filename: fileName
      };
      
      var options = {
        method: 'post',
        headers: { 'Authorization': token },
        payload: payload,
        muteHttpExceptions: true
      };

      Logger.log("🚀 Mengirim pesan...");
      var response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
      var responseText = response.getContentText();
      Logger.log('✅ Pesan berhasil dikirim: ' + responseText);
      
      // Log ke sheet LOG
      logSheet.appendRow([logColumnA, logColumnB, formattedPhone, testName, testMessage, 'TEST_SENT', new Date()]);
      Logger.log("✅ Log berhasil disimpan");
      
    } catch (error) {
      Logger.log('❌ Error mengirim pesan: ' + error);
      logSheet.appendRow([logColumnA, logColumnB, formattedPhone, testName, testMessage, 'TEST_FAILED', new Date()]);
    }

    Logger.log("=== TEST SEND REPORT SELESAI ===");
    
  } finally {
    lock.releaseLock();
  }
}