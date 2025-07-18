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

    // Ambil screenshot tabel REPORT02
    var blob = createReportScreenshot(reportSheet);
    if (!blob) {
      Logger.log("Gagal membuat screenshot");
      return;
    }

    // Buat nama file berdasarkan tanggal
    var reportDateObj = reportSheet.getRange('C3').getValue();
    var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
    var genDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    var fileName = 'KOMANDO_' + reportDateStr + '_' + genDateStr + '.png';
    blob.setName(fileName);

    // Upload gambar ke Google Drive
    var publicUrl = uploadImageToDrive(blob);
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
        // Kirim pesan dengan gambar sebagai attachment
        var payload = {
          target: phoneNumber,
          message: messageText,
          file: blob,
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
    // Deteksi range dinamis - ambil dari A1 sampai data terakhir
    var lastRow = reportSheet.getLastRow();
    
    // PENTING: Hanya ambil sampai kolom G (kolom ke-7) untuk laporan
    // Jangan ambil kolom H dan seterusnya yang berisi daftar pegawai
    var lastCol = 7; // Kolom G = kolom ke-7
    
    // Pastikan range mencakup semua elemen penting:
    // - Header info (B1:C5)
    // - Logo & header summary (B8:G10) 
    // - Header tabel (B12:G12)
    // - Data sampai baris terakhir, tapi hanya sampai kolom G
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

function uploadImageToDrive(imageBlob) {
  var folder = DriveApp.getFolderById("1WLXGZeinrbBPt6Si3Qhn7lME9UYinQiT");
  var file = folder.createFile(imageBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?id=" + file.getId();
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

    // Ambil screenshot tabel REPORT02
    Logger.log("📸 Membuat screenshot...");
    var blob = createReportScreenshot(reportSheet);
    if (!blob) {
      Logger.log("ERROR: Gagal membuat screenshot");
      return;
    }
    Logger.log("✅ Screenshot berhasil dibuat");

    // Buat nama file berdasarkan tanggal
    var reportDateObj = reportSheet.getRange('C3').getValue();
    var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
    var genDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    var fileName = 'TEST_KOMANDO_' + reportDateStr + '_' + genDateStr + '.png';
    blob.setName(fileName);
    Logger.log("✅ Nama file: " + fileName);

    // Upload gambar ke Google Drive
    Logger.log("☁️ Mengupload ke Google Drive...");
    var publicUrl = uploadImageToDrive(blob);
    Logger.log("✅ Gambar berhasil diupload ke Drive: " + publicUrl);

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
      
      // Kirim pesan dengan gambar sebagai attachment
      var payload = {
        target: formattedPhone,
        message: testMessage,
        file: blob,
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