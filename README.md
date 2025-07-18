# Mini Project PLN - Komando Report & WhatsApp Sender

A Google Apps Script solution to:

- Capture a dynamic screenshot of the `REPORT02` sheet in Google Sheets
- Upload the screenshot to Google Drive
- Send the same screenshot via WhatsApp to a list of recipients using the Fonnte API
- Log every send attempt in a dedicated `LOG` sheet

---

## Features

- **Flexible Reporting**: Custom 📄 menu in Google Sheets with multiple sending options
  - **Send only TO**: Send to group recipients only
  - **Send TO & CC**: Send to both group and individual recipients
  - **Test Send Report**: Verify setup with a test message
- **Dynamic Range Detection**: Automatically finds the last row of data and captures only columns A–F
- **Image Consistency**: Ensures the image sent to WhatsApp is identical to the one uploaded to Drive
- **Configurable & Testable**: Built-in test function to validate message sending
- **Full Logging**: Appends a row per message in the `LOG` sheet with detailed status and timestamp

---

## Prerequisites

1. A Google Workspace account with access to Google Sheets and Apps Script
2. A spreadsheet with three sheets named exactly:
   - `REPORT02` (your report data)
   - `MSGSENDER` (WhatsApp token & recipients)
   - `LOG` (message send history)
3. A valid Fonnte API token (placed in cell B5 of `MSGSENDER`)
4. A Google Drive folder ID where screenshots will be uploaded (configure in script)

---

## Installation & Setup

1. **Open your Google Spreadsheet**.
2. **Tools → Script editor** and replace the default code with the contents of `Komando/SendReport.gs`.
3. In `uploadImageToDrive(imageBlob)`, replace the folder ID with yours:
   ```js
   var folder = DriveApp.getFolderById("YOUR_DRIVE_FOLDER_ID");
   ```
4. Set your Fonnte API token in `MSGSENDER!B5`.
5. Populate `MSGSENDER` rows B7:E with:
   - Column B: Recipient type (TO/CC)
   - Column C: Phone number
   - Column D: Recipient name (optional)
   - Column E: Message text
6. Ensure `REPORT02` has:
   - Cell B2: logColumnA label (e.g. report name)
   - Cells B3/B4: start and end dates
   - Data table in columns A–F starting row 12
7. Save the script. Reload the spreadsheet to see the **Send Report 📄** menu.

---

## Usage

### Sending Options
1. **Send only TO 📩**
   - Sends messages ONLY to recipients marked as "TO" in column B
   - Typically used for sending to group recipients

2. **Send TO & CC 📩📋**
   - Sends messages to ALL recipients (both "TO" and "CC")
   - Used for comprehensive communication including groups and individuals

3. **Test Send Report 🧪**
   - Sends a test message to a predefined number
   - Helps verify setup and API connectivity

### Logging
- Each message send attempt is logged in the `LOG` sheet
- Logs include:
  - Report name
  - Date range
  - Recipient details
  - Send status (SENT_TO, SENT_CC, FAILED_TO, FAILED_CC)
  - Timestamp

---

## How It Works

- `onOpen()` – adds a custom menu with three items
- `sendReportTOOnly()` – sends to TO recipients
- `sendReportTOAndCC()` – sends to all recipients
- `testSendReport()` – sends a test message
- `createReportScreenshot()` – generates a table chart image
- `uploadImageToDrive()` – uploads image and returns a shareable link

---

## Troubleshooting

- **Image Mismatch**: Ensure `flush()` and short `sleep()` are in place before screenshot
- **Stale Data**: Timestamp in `REPORT02!B5` is updated right before snapshot
- **Fonnte URL Issues**: If direct URL fails, send the blob (`file` payload) instead
- **Lock Errors**: If a process is still running, the script logs and exits

---

## Contributing

1. Fork the project
2. Create a feature branch
3. Submit a pull request with a clear description

---

## License

MIT © [Mini Project-2025/07/19]