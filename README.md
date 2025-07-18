# Mini Project PLN - Komando Report & WhatsApp Sender

A Google Apps Script solution to:

- Capture a dynamic screenshot of the `REPORT02` sheet in Google Sheets
- Upload the screenshot to Google Drive
- Send the same screenshot via WhatsApp to a list of recipients using the Fonnte API
- Log every send attempt in a dedicated `LOG` sheet

---

## Features

- **Single-Click Report**: Custom 📄 menu in Google Sheets with `Send Report` and `Test Send Report` options.
- **Dynamic Range Detection**: Automatically finds the last row of data and captures only columns A–F.
- **Image Consistency**: Ensures the image sent to WhatsApp is identical to the one uploaded to Drive.
- **Configurable & Testable**: Built-in `testSendReport` function sends a test message to a debug number.
- **Full Logging**: Appends a row per message in the `LOG` sheet with status and timestamp.

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
5. Populate `MSGSENDER` rows C7:E with recipient phone, name (optional), and message text.
6. Ensure `REPORT02` has:
   - Cell B2: logColumnA label (e.g. report name)
   - Cells B3/B4: start and end dates
   - Data table in columns A–F starting row 12
7. Save the script. Reload the spreadsheet to see the **Send Report 📄** menu.

---

## Usage

### Send Report (Production)
1. In the spreadsheet UI, click **Send Report 📄 → Send Report 📩**.
2. The script will:
   - Update the timestamp in `REPORT02!B5`
   - Capture a fresh screenshot of your report table
   - Upload it to Drive and retrieve its file blob
   - Send the image via WhatsApp to each recipient in `MSGSENDER`
   - Log each send attempt in `LOG`

### Test Send Report (Debug)
1. Click **Send Report 📄 → Test Send Report 🧪**.
2. Sends a test message (and image) to `DEBUG_PHONE_RAW` or a hard-coded test number.
3. Ideal for verifying your setup without hitting live recipients.

---

## How It Works

- `onOpen()` – adds a custom menu with two items
- `sendReport()` – main function for production sends
- `testSendReport()` – similar flow, targets a single debug number
- `createReportScreenshot()` – flushes changes, auto-detects data range in columns A-F, builds a table chart and returns a PNG blob
- `uploadImageToDrive()` – uploads the blob to Drive, sets sharing to anyone-with-link, returns a direct download URL
- Helpers:
  - `findLastDataRow()` – finds the last non-empty row in column F
  - `formatphone()`, `formatDate()`, `formatDateRange()` for formatting

---

## Troubleshooting

- **Image Mismatch**: Ensure `flush()` and short `sleep()` are in place before screenshot.
- **Stale Data**: Timestamp in `REPORT02!B5` is updated right before snapshot.
- **Fonnte URL Issues**: If direct URL fails, send the blob (`file` payload) instead of `url`.
- **Lock Errors**: If a process is still running, the script logs and exits.

---

## Contributing

1. Fork the project
2. Create a feature branch
3. Submit a pull request with a clear description

---

## License

MIT © [Mini Project-2025/07/19]