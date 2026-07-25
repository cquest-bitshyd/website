// Google Apps Script image proxy for the c-QUEST site.
// 1) Paste this into a new Apps Script project.
// 2) Deploy as a web app: Execute as Me, Who has access: Anyone.
// 3) Set window.CQUEST_IMAGE_PROXY_BASE to the web app URL in the page.

function doGet(e) {
  const id = e && e.parameter && e.parameter.id ? e.parameter.id : '';
  if (!id) {
    return ContentService.createTextOutput('Missing image id')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    const file = DriveApp.getFileById(id);
    const blob = file.getBlob();
    const contentType = blob.getContentType() || 'image/jpeg';
    const payload = contentType + '|' + Utilities.base64Encode(blob.getBytes());
    return ContentService.createTextOutput(payload)
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput('Error: ' + error.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function convertSheetImageLinksToProxy() {
  const SPREADSHEET_ID = 'PUT_YOUR_SHEET_ID_HERE';
  const SHEET_NAME = 'Sheet1';
  const IMAGE_COL = 4;
  const PROXY_BASE = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  const values = sh.getDataRange().getValues();

  for (let r = 1; r < values.length; r++) {
    const currentValue = values[r][IMAGE_COL - 1] || '';
    const fileId = extractDriveId(currentValue);
    if (fileId) {
      values[r][IMAGE_COL - 1] = `${PROXY_BASE}?id=${fileId}`;
    }
  }

  sh.getRange(1, 1, values.length, values[0].length).setValues(values);
}

function extractDriveId(url) {
  if (!url) return '';
  const matches = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return matches && matches[1] ? matches[1] : '';
}
