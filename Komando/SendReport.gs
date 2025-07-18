// Debug override: send all messages to this number for testing
var DEBUG_OVERRIDE = false;
var DEBUG_PHONE_RAW = '';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Send Report 📄')
    .addItem('Send only TO 📩', 'sendReportTOOnly')
    .addItem('Send TO & CC 📩📋', 'sendReportTOAndCC')
    .addItem('Test Send Report 🧪', 'testSendReport')
    .addToUi();
}

function sendReportTOOnly() {
  sendReportWithFilter('TO');
}

function sendReportTOAndCC() {
  sendReportWithFilter('ALL');
}

function sendReportWithFilter(filterType) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Could not acquire lock. Process is already running.");
    return;
  }

  try {
    // Get all required sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName('REPORT02');
    var msgSheet = ss.getSheetByName('MSGSENDER');
    var logSheet = ss.getSheetByName('LOG');
    
    if (!reportSheet) {
      Logger.log("Sheet 'REPORT02' not found!");
      return;
    }
    if (!msgSheet) {
      Logger.log("Sheet 'MSGSENDER' not found!");
      return;
    }
    if (!logSheet) {
      Logger.log("Sheet 'LOG' not found!");
      return;
    }

    // Update timestamp in sheet before taking screenshot
    reportSheet.getRange('B5').setValue(new Date());
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    // Create file name based on date
    var reportDateObj = reportSheet.getRange('B3').getValue();
    var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
    var now = new Date();
    var genDateTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var fileName = 'KOMANDO_' + reportDateStr + '_' + genDateTimeStr + '.png';

    // IMPORTANT: Capture screenshots twice to ensure they are identical
    // First screenshot for WhatsApp
    var blobForWhatsApp = createReportScreenshot(reportSheet);
    if (!blobForWhatsApp) {
      Logger.log("Failed to create screenshot for WhatsApp");
      return;
    }
    blobForWhatsApp.setName(fileName);
    
    // Second screenshot for Drive
    var blobForDrive = createReportScreenshot(reportSheet);
    if (!blobForDrive) {
      Logger.log("Failed to create screenshot for Drive");
      return;
    }
    blobForDrive.setName(fileName);
    
    // Upload image to Google Drive
    var publicUrl = uploadImageToDrive(blobForDrive);
    Logger.log("Image successfully uploaded to Drive: " + publicUrl);

    // Get data for logging
    var logColumnA = reportSheet.getRange('B2').getValue();
    var startDate = reportSheet.getRange('B3').getValue();
    var endDate = reportSheet.getRange('B4').getValue();
    var logColumnB = formatDateRange(startDate, endDate);

    // Retrieve token and recipient data
    var token = msgSheet.getRange('C5').getValue();
    var recipientData = msgSheet.getRange(7, 2, msgSheet.getLastRow() - 6, 4).getValues(); // B7:E (B=type, C=phone, D=name, E=message)

    var sentCount = 0;
    var skippedCount = 0;

    // Send message to each recipient based on filter
    recipientData.forEach(function(row) {
      var recipientType = row[0]; // B column (TO/CC)
      var phoneNumber = getTargetPhone(row[1]); // C column
      var recipientName = row[2]; // D column
      var messageText = row[3]; // E column

      // Skip if data is incomplete
      if (!phoneNumber || !messageText || phoneNumber === '' || messageText === '') {
        return;
      }

      // Apply filter logic
      if (filterType === 'TO' && recipientType !== 'TO') {
        skippedCount++;
        Logger.log('Skipped ' + recipientType + ' recipient: ' + recipientName);
        return;
      }

      try {
        // After uploading to Drive...
        var fileId = publicUrl.match(/id=([^&]+)/)[1];
        var driveBlob = DriveApp.getFileById(fileId).getBlob();
        driveBlob.setName(fileName);
        
        // Send message with image file as attachment
        var payload = {
          target:   phoneNumber,
          message:  messageText,
          file:     blobForWhatsApp,   // or driveBlob if you want to use the uploaded file
          filename: fileName
        };
        
        var options = {
          method: 'post',
          headers: { 'Authorization': token },
          payload: payload,
          muteHttpExceptions: true
        };

        var response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
        Logger.log('Message sent to ' + phoneNumber + ' (' + recipientType + '): ' + response.getContentText());
        
        // Log to LOG sheet
        logSheet.appendRow([logColumnA, logColumnB, phoneNumber, recipientName, messageText, 'SENT_' + recipientType, new Date()]);
        sentCount++;
        
      } catch (error) {
        Logger.log('Error sending message to ' + phoneNumber + ' (' + recipientType + '): ' + error);
        logSheet.appendRow([logColumnA, logColumnB, phoneNumber, recipientName, messageText, 'FAILED_' + recipientType, new Date()]);
      }
      
      // Delay to avoid rate limiting
      Utilities.sleep(3000);
    });

    Logger.log("Report sending process completed. Sent: " + sentCount + ", Skipped: " + skippedCount);
    
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
  
  // If date is a string, try to convert to Date object
  if (typeof date === 'string') {
    date = new Date(date);
  }
  
  // Validate if date is a valid Date object
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return date.toString(); // Return original string if cannot convert
  }
  
  var day = date.getDate();
  var month = date.getMonth();
  var year = date.getFullYear();
  
  var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  
  return day + " " + monthNames[month] + " " + year;
}

function formatDateRange(startDate, endDate) {
  return formatDate(startDate) + " to " + formatDate(endDate);
}

function createReportScreenshot(reportSheet) {
  try {
    // Ensure sheet changes are applied
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    // Determine last data row and fixed columns
    var lastRow = findLastDataRow(reportSheet);
    var lastCol = 6; // up to column F

    // Read data range A1:F(lastRow)
    var dataVals = reportSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    
    // Build rows array with blank rows after row 5 and 8 (0-based index 4 and 7)
    var combined = [];
    dataVals.forEach(function(row, i) {
      combined.push(row);
      if (i === 4 || i === 7) {
        combined.push(Array(lastCol).fill(''));
      }
    });

    // Build DataTable for chart
    var dtBuilder = Charts.newDataTable();
    for (var c = 0; c < lastCol; c++) {
      dtBuilder.addColumn(Charts.ColumnType.STRING, '');
    }
    // Add each row individually to DataTableBuilder
    combined.forEach(function(row) {
      dtBuilder.addRow(row);
    });
    var dataTable = dtBuilder.build();

    // Create table chart using Charts service
    var chart = Charts.newTableChart()
      .setDataTable(dataTable)
      .setOption('width', 1200)
      .setOption('height', Math.max(800, combined.length * 25))
      .setOption('backgroundColor', 'white')
      .setOption('enableInteractivity', false)
      .setOption('allowHtml', true)
      .setOption('showRowNumber', false)
      .setOption('page', 'disable')
      .build();

    // Return chart image blob
    return chart.getBlob();
  } catch (error) {
    Logger.log('Error creating screenshot: ' + error);
    return null;
  }
}

// Helper function to convert column number to letter
function getColumnLetter(columnNumber) {
  var columnLetter = '';
  while (columnNumber > 0) {
    var remainder = (columnNumber - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnLetter;
}

// Helper function to find the last row with data
function findLastDataRow(reportSheet) {
  // Retrieve all data in column F (F1:F100)
  var columnFData = reportSheet.getRange('F1:F100').getValues();
  
  // Search from bottom up to find the last non-empty row
  for (var i = columnFData.length - 1; i >= 0; i--) {
    if (columnFData[i][0] !== '' && columnFData[i][0] !== null && columnFData[i][0] !== undefined) {
      var lastDataRow = i + 1; // +1 because array is 0-based but rows start at 1
      Logger.log("Last row with data based on Column F: " + lastDataRow);
      return lastDataRow;
    }
  }
  
  // If no data found, return a default of 30 rows for safety
  Logger.log("No data found in Column F, using default 30 rows");
  return 30;
}

function uploadImageToDrive(imageBlob) {
  var folder = DriveApp.getFolderById("1WLXGZeinrbBPt6Si3Qhn7lME9UYinQiT");
  var file = folder.createFile(imageBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Use export=download for direct download
  return "https://drive.google.com/uc?export=download&id=" + file.getId();
}

function testSendReport() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Could not acquire lock. Process is already running.");
    return;
  }

  try {
    Logger.log("=== START TEST SEND REPORT ===");
    
    // Get all required sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName('REPORT02');
    var msgSheet = ss.getSheetByName('MSGSENDER');
    var logSheet = ss.getSheetByName('LOG');
    
    if (!reportSheet) {
      Logger.log("ERROR: Sheet 'REPORT02' not found!");
      return;
    }
    if (!msgSheet) {
      Logger.log("ERROR: Sheet 'MSGSENDER' not found!");
      return;
    }
    if (!logSheet) {
      Logger.log("ERROR: Sheet 'LOG' not found!");
      return;
    }
    
    Logger.log("✅ All sheets found");

    // Update timestamp in sheet before taking screenshot
    reportSheet.getRange('B5').setValue(new Date());
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    // Create file name based on date
    var reportDateObj = reportSheet.getRange('B3').getValue();
    var reportDateStr = Utilities.formatDate(reportDateObj, Session.getScriptTimeZone(), 'yyyyMMdd');
    var now = new Date();
    var genDateTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var fileName = 'TEST_KOMANDO_' + reportDateStr + '_' + genDateTimeStr + '.png';

    // Take one screenshot and upload to Drive
    Logger.log("📸 Creating report screenshot...");
    var screenshotBlob = createReportScreenshot(reportSheet);
    if (!screenshotBlob) {
      Logger.log("ERROR: Failed to create report screenshot");
      return;
    }
    screenshotBlob.setName(fileName);
    Logger.log("✅ Screenshot created");

    // Upload image to Google Drive
    Logger.log("☁️ Uploading screenshot to Google Drive...");
    var publicUrl = uploadImageToDrive(screenshotBlob);
    Logger.log("✅ Screenshot uploaded to Drive: " + publicUrl);

    // Get data for logging
    var logColumnA = reportSheet.getRange('B2').getValue();
    var startDate = reportSheet.getRange('B3').getValue();
    var endDate = reportSheet.getRange('B4').getValue();
    var logColumnB = formatDateRange(startDate, endDate);
    Logger.log("✅ Logging data: " + logColumnA + " | " + logColumnB);

    // Retrieve token
    var token = msgSheet.getRange('C5').getValue();
    if (!token) {
      Logger.log("ERROR: Token not found in C5");
      return;
    }
    Logger.log("✅ Token retrieved");

    // Test data - sending to test number
    var testPhone = '087778651293';
    var testName = 'Test User';
    var testMessage = '🧪 KOMANDO REPORT TEST\n\nThis is an automated report sending test.\n\nDate: ' + formatDateRange(startDate, endDate);
    
    Logger.log("📱 Sending test to: " + testPhone);
    Logger.log("💬 Message: " + testMessage);

    try {
      // Format phone number
      var formattedPhone = formatPhone(testPhone);
      Logger.log("📞 Formatted number: " + formattedPhone);
      
      // Get blob directly from the newly uploaded Drive file
      var fileId = publicUrl.match(/id=([^&]+)/)[1];
      var driveBlob = DriveApp.getFileById(fileId).getBlob();
      driveBlob.setName(fileName);
      
      // Send message with Drive file as attachment
      var payload = {
        target: formattedPhone,
        message: testMessage,
        file:    screenshotBlob,
        filename: fileName
      };
      
      var options = {
        method: 'post',
        headers: { 'Authorization': token },
        payload: payload,
        muteHttpExceptions: true
      };

      Logger.log("🚀 Sending message...");
      var response = UrlFetchApp.fetch('https://api.fonnte.com/send', options);
      var responseText = response.getContentText();
      Logger.log('✅ Message sent: ' + responseText);
      
      // Log to LOG sheet
      logSheet.appendRow([logColumnA, logColumnB, formattedPhone, testName, testMessage, 'TEST_SENT', new Date()]);
      Logger.log("✅ Log saved");
      
    } catch (error) {
      Logger.log('❌ Error sending message: ' + error);
      logSheet.appendRow([logColumnA, logColumnB, formattedPhone, testName, testMessage, 'TEST_FAILED', new Date()]);
    }

    Logger.log("=== TEST SEND REPORT COMPLETED ===");
    
  } finally {
    lock.releaseLock();
  }
}