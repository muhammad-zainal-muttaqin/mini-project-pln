/**
 * @fileoverview Google Apps Script for sending reports from a Google Sheet via WhatsApp.
 * This script creates a report from a sheet, saves it as an image to Google Drive,
 * and sends it to a list of recipients using the Fonnte API.
 */

// =================================================================
// CONFIGURATION
// =================================================================

/**
 * Script configuration settings.
 * @const
 */
const CONFIG = {
  SHEET_NAMES: {
    REPORT: 'REPORT02',
    MESSAGES: 'MSGSENDER',
    LOG: 'LOG',
  },
  DRIVE_FOLDER_ID: '1WLXGZeinrbBPt6Si3Qhn7lME9UYinQiT', // Consider making this configurable outside the script for multi-environment deployments.
  API: {
    URL: 'https://api.fonnte.com/send',
    TOKEN_CELL: 'C5',
  },
  DEBUG: {
    ENABLED: false,
    PHONE_NUMBER: '', // E.g., '6281234567890'
  },
  LOCK_TIMEOUT: 5000, // 5 seconds
  SLEEP_TIME: {
    FLUSH: 500, // ms to wait after flushing spreadsheet changes
    API_RATE_LIMIT: 3000, // ms to wait between API calls
  },
  // New constants for cell references and ranges
  CELL_REFERENCES: {
    REPORT_TIMESTAMP: 'B5',
    REPORT_LOG_A: 'B2',
    REPORT_START_DATE: 'B3',
    REPORT_END_DATE: 'B4',
  },
  RECIPIENT_RANGE: {
    START_ROW: 7,
    START_COL: 2,
    NUM_COLS: 4, // recipientType, rawPhone, name, message
  },
  REPORT_LAST_COLUMN_INDEX: 6, // Column F
};

// =================================================================
// UI & MENU
// =================================================================

/**
 * Adds a custom menu to the Google Sheets UI when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Send Report 📄')
    .addItem('Send only TO 📩', 'sendReportTOOnly')
    .addItem('Send TO & CC 📩📋', 'sendReportTOAndCC')
    .addSeparator()
    .addItem('Test Send Report 🧪', 'testSendReport')
    .addToUi();
}

/**
 * Entry point for sending the report to 'TO' recipients only.
 */
function sendReportTOOnly() {
  sendReport('TO');
}

/**
 * Entry point for sending the report to 'TO' and 'CC' recipients.
 */
function sendReportTOAndCC() {
  sendReport('ALL');
}

/**
 * Entry point for sending a test report.
 */
function testSendReport() {
  sendReport('TEST');
}

// =================================================================
// MAIN REPORTING LOGIC
// =================================================================

/**
 * Main function to generate and send the report.
 * @param {string} filterType - The recipient filter type ('TO', 'ALL', or 'TEST').
 */
function sendReport(filterType) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_TIMEOUT)) {
    Logger.log('Could not acquire lock. Process is already running.');
    return;
  }

  try {
    Logger.log(`=== STARTING REPORT PROCESS (Filter: ${filterType}) ===`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getRequiredSheets(ss);
    
    updateTimestamp(sheets.report);

    const reportDetails = getReportDetails(sheets.report);
    const fileName = createFileName(reportDetails.date, filterType === 'TEST');
    
    const screenshotBlob = createReportScreenshot(sheets.report);
    screenshotBlob.setName(fileName);

    const publicUrl = uploadImageToDrive(screenshotBlob, fileName);
    Logger.log('Image successfully uploaded to Drive: ' + publicUrl);

    if (filterType === 'TEST') {
      sendTestMessage(sheets, reportDetails, screenshotBlob, fileName);
    } else {
      sendMessagesToRecipients(sheets, reportDetails, screenshotBlob, fileName, publicUrl, filterType);
    }

    Logger.log('=== REPORT PROCESS COMPLETED ===');
  } catch (error) {
    Logger.log(`ERROR in sendReport: ${error.message}\n${error.stack}`);
  } finally {
    lock.releaseLock();
  }
}

// =================================================================
// HELPER FUNCTIONS
// =================================================================

/**
 * Retrieves all required sheets from the spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet.
 * @returns {{report: GoogleAppsScript.Spreadsheet.Sheet, messages: GoogleAppsScript.Spreadsheet.Sheet, log: GoogleAppsScript.Spreadsheet.Sheet}}
 * @throws {Error} If a required sheet is not found.
 */
function getRequiredSheets(ss) {
  const sheets = {
    report: ss.getSheetByName(CONFIG.SHEET_NAMES.REPORT),
    messages: ss.getSheetByName(CONFIG.SHEET_NAMES.MESSAGES),
    log: ss.getSheetByName(CONFIG.SHEET_NAMES.LOG),
  };

  for (const [name, sheet] of Object.entries(sheets)) {
    if (!sheet) {
      throw new Error(`Sheet '${CONFIG.SHEET_NAMES[name.toUpperCase()]}' not found!`);
    }
  }
  return sheets;
}

/**
 * Updates the timestamp on the report sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} reportSheet
 */
function updateTimestamp(reportSheet) {
  reportSheet.getRange('B5').setValue(new Date());
  SpreadsheetApp.flush();
  Utilities.sleep(CONFIG.SLEEP_TIME.FLUSH);
}

/**
 * Gets report metadata from the report sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} reportSheet
 * @returns {{logA: string, logB: string, date: Date}}
 */
function getReportDetails(reportSheet) {
  const logA = reportSheet.getRange(CONFIG.CELL_REFERENCES.REPORT_LOG_A).getValue();
  const startDate = reportSheet.getRange(CONFIG.CELL_REFERENCES.REPORT_START_DATE).getValue();
  const endDate = reportSheet.getRange(CONFIG.CELL_REFERENCES.REPORT_END_DATE).getValue();
  return {
    logA: logA,
    logB: formatDateRange(startDate, endDate),
    date: startDate,
  };
}

/**
 * Creates a unique file name for the report image.
 * @param {Date} reportDate - The date of the report.
 * @param {boolean} isTest - Whether this is a test report.
 * @returns {string} The generated file name.
 */
function createFileName(reportDate, isTest = false) {
  const reportDateStr = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), 'yyyyMMdd');
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const prefix = isTest ? 'TEST_KOMANDO_' : 'KOMANDO_';
  return `${prefix}${reportDateStr}_${nowStr}.png`;
}

/**
 * Sends messages to all relevant recipients.
 * @param {object} sheets - The collection of required sheets.
 * @param {object} reportDetails - The details for logging.
 * @param {GoogleAppsScript.Base.Blob} blob - The report image blob.
 * @param {string} fileName - The name of the image file.
 * @param {string} publicUrl - The public URL of the image on Google Drive.
 * @param {string} filterType - The recipient filter ('TO' or 'ALL').
 */
function sendMessagesToRecipients(sheets, reportDetails, blob, fileName, publicUrl, filterType) {
  const token = sheets.messages.getRange(CONFIG.API.TOKEN_CELL).getValue();
  if (!token) {
    throw new Error(`API Token not found in cell ${CONFIG.API.TOKEN_CELL}`);
  }

  const recipientData = sheets.messages.getRange(
    CONFIG.RECIPIENT_RANGE.START_ROW,
    CONFIG.RECIPIENT_RANGE.START_COL,
    sheets.messages.getLastRow() - (CONFIG.RECIPIENT_RANGE.START_ROW - 1),
    CONFIG.RECIPIENT_RANGE.NUM_COLS
  ).getValues();
  let sentCount = 0;
  let skippedCount = 0;

  recipientData.forEach(row => {
    const [recipientType, rawPhone, name, message] = row;

    if (!rawPhone || !message) return;

    if (filterType === 'TO' && recipientType !== 'TO') {
      skippedCount++;
      Logger.log(`Skipped ${recipientType} recipient: ${name}`);
      return;
    }

    const phoneNumber = getTargetPhone(rawPhone);
    try {
      sendMessage(token, phoneNumber, message, blob, fileName);
      logResult(sheets.log, reportDetails, phoneNumber, name, message, `SENT_${recipientType}`);
      sentCount++;
    } catch (error) {
      Logger.log(`Error sending to ${phoneNumber} (${name}): ${error.message}`);
      logResult(sheets.log, reportDetails, phoneNumber, name, message, `FAILED_${recipientType}`);
    }
    Utilities.sleep(CONFIG.SLEEP_TIME.API_RATE_LIMIT);
  });

  Logger.log(`Process completed. Sent: ${sentCount}, Skipped: ${skippedCount}`);
}

/**
 * Sends a single test message.
 * @param {object} sheets - The collection of required sheets.
 * @param {object} reportDetails - The details for logging.
 * @param {GoogleAppsScript.Base.Blob} blob - The report image blob.
 * @param {string} fileName - The name of the image file.
 */
function sendTestMessage(sheets, reportDetails, blob, fileName) {
  const token = sheets.messages.getRange(CONFIG.API.TOKEN_CELL).getValue();
  const testPhone = '087778651293'; // Hardcoded for testing
  const testName = 'Test User';
  const testMessage = `🧪 KOMANDO REPORT TEST\n\nThis is an automated report sending test.\n\nDate: ${reportDetails.logB}`;
  const phoneNumber = formatPhone(testPhone);

  try {
    Logger.log(`Sending test message to ${phoneNumber}`);
    sendMessage(token, phoneNumber, testMessage, blob, fileName);
    logResult(sheets.log, reportDetails, phoneNumber, testName, testMessage, 'TEST_SENT');
    Logger.log('Test message sent successfully.');
  } catch (error) {
    Logger.log(`Error sending test message: ${error.message}`);
    logResult(sheets.log, reportDetails, phoneNumber, testName, testMessage, 'TEST_FAILED');
  }
}

/**
 * Sends a message via the Fonnte API.
 * @param {string} token - The API authorization token.
 * @param {string} target - The recipient's phone number.
 * @param {string} message - The text message.
 * @param {GoogleAppsScript.Base.Blob} fileBlob - The file to attach.
 * @param {string} filename - The name of the attached file.
 */
function sendMessage(token, target, message, fileBlob, filename) {
  const payload = { target, message, file: fileBlob, filename };
  const options = {
    method: 'post',
    headers: { 'Authorization': token },
    payload: payload,
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(CONFIG.API.URL, options);
  Logger.log(`Message sent to ${target}. Response: ${response.getContentText()}`);
}

/**
 * Logs the result of a send attempt to the LOG sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet
 * @param {object} reportDetails - Contains logA and logB.
 * @param {string} phone - The recipient's phone number.
 * @param {string} name - The recipient's name.
 * @param {string} message - The message sent.
 * @param {string} status - The result status (e.g., 'SENT_TO', 'FAILED_CC').
 */
function logResult(logSheet, reportDetails, phone, name, message, status) {
  logSheet.appendRow([reportDetails.logA, reportDetails.logB, phone, name, message, status, new Date()]);
}

// =================================================================
// UTILITY FUNCTIONS
// =================================================================

/**
 * Returns the target phone number, using a debug override if enabled.
 * @param {string} rawPhone - The raw phone number from the sheet.
 * @returns {string} The formatted phone number.
 */
function getTargetPhone(rawPhone) {
  const phone = CONFIG.DEBUG.ENABLED ? CONFIG.DEBUG.PHONE_NUMBER : rawPhone;
  return formatPhone(phone);
}

/**
 * Formats a phone number to the '62...' standard.
 * @param {string|number} num - The phone number to format.
 * @returns {string} The formatted phone number.
 */
function formatPhone(num) {
  if (!num) return '';
  const strNum = String(num).trim();
  if (strNum.startsWith('+628')) return '628' + strNum.substring(4);
  if (strNum.startsWith('08')) return '628' + strNum.substring(2);
  return strNum;
}

/**
 * Formats a date range string.
 * @param {Date} startDate
 * @param {Date} endDate
 * @returns {string} Formatted date range (e.g., "1 Jan 2023 to 31 Jan 2023").
 */
function formatDateRange(startDate, endDate) {
  return `${formatDate(startDate)} to ${formatDate(endDate)}`;
}

/**
 * Formats a single date object into "D MMM YYYY" format.
 * @param {Date|string} date - The date to format.
 * @returns {string} The formatted date string.
 */
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') {
    date = new Date(date);
  }
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return String(date);
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'd MMM yyyy');
}

/**
 * Creates a screenshot of the report range as a blob.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} reportSheet
 * @returns {GoogleAppsScript.Base.Blob} The image blob of the report chart.
 * @throws {Error} If the screenshot creation fails.
 */
function createReportScreenshot(reportSheet) {
  try {
    const lastRow = findLastDataRow(reportSheet, 'F', 100);
    const lastCol = CONFIG.REPORT_LAST_COLUMN_INDEX; 
    const data = reportSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    const dataTable = Charts.newDataTable();
    for (let c = 0; c < lastCol; c++) {
      dataTable.addColumn(Charts.ColumnType.STRING, '');
    }
    data.forEach(row => dataTable.addRow(row));

    const chart = Charts.newTableChart()
      .setDataTable(dataTable)
      .setOption('width', 1200)
      .setOption('height', Math.max(800, data.length * 25 + 100)) // Add padding
      .setOption('allowHtml', true)
      .setOption('showRowNumber', false)
      .build();

    return chart.getBlob();
  } catch (error) {
    throw new Error(`Failed to create screenshot: ${error.message}`);
  }
}

/**
 * Finds the last row with data in a specific column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to search.
 * @param {string} column - The column letter to check (e.g., 'F').
 * @param {number} maxRows - The maximum number of rows to check.
 * @returns {number} The last row number with data.
 */
function findLastDataRow(sheet, column, maxRows) {
  const values = sheet.getRange(`${column}1:${column}${maxRows}`).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0]) {
      return i + 1;
    }
  }
  return 30; // Default fallback
}

/**
 * Uploads a blob to a specific Google Drive folder and returns its public URL.
 * @param {GoogleAppsScript.Base.Blob} blob - The file blob to upload.
 * @param {string} fileName - The desired file name.
 * @returns {string} The public download URL of the file.
 */
function uploadImageToDrive(blob, fileName) {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const file = folder.createFile(blob.setName(fileName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://drive.google.com/uc?export=download&id=${file.getId()}`;
}
