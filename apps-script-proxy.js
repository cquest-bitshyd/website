// Google Apps Script image proxy for the c-QUEST site.
// 1) Paste this into a new Apps Script project.
// 2) Deploy as a web app: Execute as Me, Who has access: Anyone.
// 3) Use the custom menu to set your deployed web app URL.

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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('c-QUEST Images')
    .addItem('Set Proxy URL', 'promptForProxyBase')
    .addItem('Convert Drive Links to Proxy URLs', 'convertSheetImageLinksToProxy')
    .addToUi();
}

function promptForProxyBase() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Deploy URL',
    'Paste the Google Apps Script web app URL ending in /exec',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const value = response.getResponseText().trim();
    if (value) {
      PropertiesService.getScriptProperties().setProperty('PROXY_BASE', value);
      ui.alert('Proxy URL saved. You can now convert your image links.');
    }
  }
}

function getProxyBase() {
  return PropertiesService.getScriptProperties().getProperty('PROXY_BASE') || '';
}

function convertSheetImageLinksToProxy() {
  const proxyBase = getProxyBase();
  if (!proxyBase) {
    SpreadsheetApp.getUi().alert('No proxy URL is set. Use the menu to set it first.');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    SpreadsheetApp.getUi().alert('The sheet is empty.');
    return;
  }

  const headers = values[0] || [];
  const imageColumns = findImageColumns(headers);
  let updated = false;

  for (let row = 1; row < values.length; row++) {
    for (const col of imageColumns) {
      const currentValue = values[row][col] || '';
      const fileId = extractDriveId(currentValue);
      if (fileId) {
        values[row][col] = `${proxyBase}?id=${fileId}`;
        updated = true;
      }
    }
  }

  if (!updated) {
    SpreadsheetApp.getUi().alert('No Drive image links were found to convert.');
    return;
  }

  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  SpreadsheetApp.getUi().alert('Image links were converted to proxy URLs.');
}

function onEdit(e) {
  const proxyBase = getProxyBase();
  if (!proxyBase) return;

  const range = e && e.range ? e.range : null;
  if (!range) return;

  const editedValue = range.getValue();
  if (typeof editedValue !== 'string' || !editedValue) return;

  const sheet = range.getSheet();
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const imageColumns = findImageColumns(headerRow);
  const isImageColumn = imageColumns.includes(range.getColumn() - 1);
  if (!isImageColumn) return;

  const fileId = extractDriveId(editedValue);
  if (!fileId) return;

  range.setValue(`${proxyBase}?id=${fileId}`);
}

function findImageColumns(headers) {
  return headers.reduce((cols, header, index) => {
    const label = String(header || '').toLowerCase();
    if (label.includes('image') || label.includes('img') || label.includes('photo') || label.includes('cover')) {
      cols.push(index);
    }
    return cols;
  }, []);
}

function extractDriveId(url) {
  if (!url || typeof url !== 'string') return '';
  const matches = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return matches && matches[1] ? matches[1] : '';
}
