// Debug override: send all messages to this number for testing
var DEBUG_OVERRIDE = false;
var DEBUG_PHONE_RAW = '087778651293';

function getTargetPhone(rawPhone) {
  return formatPhone(DEBUG_OVERRIDE ? DEBUG_PHONE_RAW : rawPhone);
}

function generateReport02() {
  // Ambil screenshot tabel REPORT02 sebagai blob
  var blob = getReportScreenshot();
  if (!blob) {
    Logger.log("Gagal ambil screenshot");
    return;
  }
  // Ambil spreadsheet dan tanggal report
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName('REPORT02');
  if (!reportSheet) {
    Logger.log("Sheet 'REPORT02' tidak ditemukan!");
    return;
  }
  var reportDateObj = reportSheet.getRange('C3').getValue();
  var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
  var genDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var fileName = 'KOMANDO_' + reportDateStr + '_' + genDateStr + '.png';
  blob.setName(fileName);
  // Ambil data untuk mengirim pesan
  var sheet = ss.getSheetByName('MSGSENDER');
  if (!sheet) {
    Logger.log("Sheet 'MSGSENDER' tidak ditemukan!");
    return;
  }
  var token = sheet.getRange('C5').getValue();
  var data = sheet.getRange(7, 3, sheet.getLastRow() - 6, 3).getValues(); // C7:E

  // Setup logging for generateReport02
  var logSheet = ss.getSheetByName('LOG');
  if (!logSheet) {
    Logger.log("Sheet 'LOG' tidak ditemukan!");
  }
  var logColumnA = reportSheet.getRange('C2').getValue();
  var startDate = reportSheet.getRange('C3').getValue();
  var endDate = reportSheet.getRange('C4').getValue();
  var logColumnB = formatDateRange(startDate, endDate);

  // Kirim screenshot sebagai lampiran ke setiap nomor
  data.forEach(function(row) {
    var phone = getTargetPhone(row[0]);
    var messageText = row[2];
    if (!phone || !messageText) return;
    try {
      var payload = {
        target:   phone,
        message:  messageText,
        file:     blob,
        filename: fileName
      };
      var options = {
        method: 'post',
        headers: { 'Authorization': token },
        payload: payload,
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
      Logger.log('Pesan ke ' + phone + ': ' + response.getContentText());
      if (logSheet) logSheet.appendRow([logColumnA, logColumnB, phone, row[1], messageText, 'SENT', new Date()]);
    } catch (e) {
      Logger.log('Error kirim ke ' + phone + ': ' + e);
      if (logSheet) logSheet.appendRow([logColumnA, logColumnB, phone, row[1], messageText, 'FAILED', new Date()]);
    }
    // Jeda untuk menghindari rate limit
    Utilities.sleep(3000);
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Generate 📄')
    .addItem('Generate Report 📄🔍', 'generateReport02AndUpload')
    .addItem('Send Message 📩', 'sendWhatsAppMessages')
    .addToUi();
}

function sendWhatsAppMessages() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Tidak dapat memperoleh lock. Proses sudah berjalan.");
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("MSGSENDER");
    var logSheet = ss.getSheetByName("LOG");
    var reportSheet = ss.getSheetByName("REPORT02");
    if (!sheet) {
      Logger.log("Sheet 'MSGSENDER' tidak ditemukan!");
      return;
    }
    if (!logSheet) {
      Logger.log("Sheet 'LOG' tidak ditemukan!");
      return;
    }
    if (!reportSheet) {
      Logger.log("Sheet 'REPORT02' tidak ditemukan!");
      return;
    }

    // Ambil data dari sheet REPORT02 untuk kolom A dan B di LOG
    var logColumnA = reportSheet.getRange("C2").getValue();
    var startDate = reportSheet.getRange("C3").getValue();
    var endDate = reportSheet.getRange("C4").getValue();
    
    // Debug: log tipe data yang diambil
    Logger.log("Start Date Type: " + typeof startDate + ", Value: " + startDate);
    Logger.log("End Date Type: " + typeof endDate + ", Value: " + endDate);
    
    var logColumnB = formatDateRange(startDate, endDate);

    var urlFonnte = "https://api.fonnte.com/send";
    var tokenFonnte = sheet.getRange("C5").getValue();

    var dataRange = sheet.getRange(7, 3, sheet.getLastRow() - 6, 3).getValues(); // C7:E
    for (var i = 0; i < dataRange.length; i++) {
      var phoneNumber = getTargetPhone(dataRange[i][0]); // C
      var recipientName = dataRange[i][1]; // D
      var message = dataRange[i][2]; // E

      // Skip rows that don't have phone number or message
      if (!phoneNumber || !message || phoneNumber === '' || message === '') {
        continue;
      }

      if (phoneNumber && message) {
        var payload = {
          "target": phoneNumber,
          "message": message
        };
        var options = {
          "method": "post",
          "headers": { "Authorization": tokenFonnte },
          "payload": payload,
          "muteHttpExceptions": true
        };

        try {
          var response = UrlFetchApp.fetch(urlFonnte, options);
          Logger.log("Pesan terkirim ke " + phoneNumber + ": " + response.getContentText());
          // Log ke LOG sheet
          logSheet.appendRow([logColumnA, logColumnB, phoneNumber, recipientName, message, "SENT", new Date()]);
          
          // Delay 3 detik antara setiap pengiriman untuk menghindari rate limiting
          Utilities.sleep(3000);
        } catch (error) {
          Logger.log("Error saat mengirim pesan: " + error);
          logSheet.appendRow([logColumnA, logColumnB, phoneNumber, recipientName, message, "FAILED", new Date()]);
          
          // Delay juga jika error untuk konsistensi
          Utilities.sleep(3000);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
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

function getReportScreenshot() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName("REPORT02");
  
  if (!reportSheet) {
    Logger.log("Sheet 'REPORT02' tidak ditemukan!");
    return null;
  }
  
  try {
    // Cek data terlebih dahulu
    var range = reportSheet.getRange("A1:H30");
    var data = range.getValues();
    Logger.log("Data rows: " + data.length + ", Data columns: " + data[0].length);
    Logger.log("Sample data: " + JSON.stringify(data[0]));
    
    // Gunakan range lengkap untuk menampilkan semua data
    var chart = reportSheet.newChart()
      .setChartType(Charts.ChartType.TABLE)
      .addRange(range)
      .setPosition(1, 1, 0, 0)
      .setOption('width', 1200)
      .setOption('height', 800)
      .setOption('backgroundColor', 'white')
      .setOption('legend', {position: 'none'})
      .setOption('enableInteractivity', false)
      .setOption('alternatingRowStyle', true)
      .setOption('allowHtml', true)
      .setOption('showRowNumber', false)
      .setOption('page', 'enable')
      .setOption('pageSize', 30)
      .build();
    
    // Get the chart as image without inserting it
    var chartImage = chart.getBlob();
    
    Logger.log("Chart image berhasil dibuat, ukuran: " + chartImage.getBytes().length + " bytes");
    return chartImage;
    
  } catch (error) {
    Logger.log("Error saat membuat screenshot: " + error);
    return null;
  }
}

function uploadImageAndGetDirectLink(imageBlob) {
  var folder = DriveApp.getFolderById("1WLXGZeinrbBPt6Si3Qhn7lME9UYinQiT");
  var file = folder.createFile(imageBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?id=" + file.getId();
}

function sendWhatsAppWithPublicUrl(phoneNumber, message, publicUrl, tokenFonnte) {
  var urlFonnte = "https://api.fonnte.com/send";
  var payload = {
    "target": phoneNumber,
    "message": message,
    "url": publicUrl,
    "filename": "laporan.png" // beri tahu API tipe file, agar format PNG dikenali
    // "countryCode": "62" // optional
  };
  var options = {
    "method": "post",
    "headers": { "Authorization": tokenFonnte },
    "payload": payload,
    "muteHttpExceptions": true
  };
  var response = UrlFetchApp.fetch(urlFonnte, options);
  Logger.log(response.getContentText());
}

function sendReportWithScreenshot() {
  var screenshot = getReportScreenshot();
  if (screenshot) {
    var imageUrl = uploadImageAndGetDirectLink(screenshot); // Sudah menghasilkan direct link Google Drive
    Logger.log("Public URL (Drive Direct Download): " + imageUrl);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("MSGSENDER");
    var tokenFonnte = sheet.getRange("C5").getValue();
    var phoneNumber = "6287778651293";
    var message = "Berikut laporan harian Anda, silakan cek gambar di bawah ini.";

    sendWhatsAppWithPublicUrl(phoneNumber, message, imageUrl, tokenFonnte);
    Logger.log("Pesan WhatsApp dengan gambar sudah dikirim!");
  } else {
    Logger.log("Gagal membuat screenshot!");
  }
}

function generateReport02WithUploadAndLog() {
  // Only upload the screenshot to Drive
  var blob = getReportScreenshot();
  if (!blob) {
    Logger.log("Gagal ambil screenshot");
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName('REPORT02');
  if (!reportSheet) {
    Logger.log("Sheet 'REPORT02' tidak ditemukan!");
    return;
  }
  // Build filename
  var reportDateObj = reportSheet.getRange('C3').getValue();
  var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
  var genDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var fileName = 'KOMANDO_' + reportDateStr + '_' + genDateStr + '.png';
  blob.setName(fileName);
  // Upload to Drive
  var publicUrl = uploadImageAndGetDirectLink(blob);
  Logger.log("Uploaded to Drive: " + publicUrl);
}

// Wrapper: run original send + upload when Generate Report is pressed
function generateReport02AndUpload() {
  generateReport02();
  generateReport02WithUploadAndLog();
}