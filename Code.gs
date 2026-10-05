const SPREADSHEET_ID = '1rNhv5eqqhQ3JcrvN7JenRMvjgOYXwf35KuSPAPPK_nM';

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request body is required.');
    }

    const payload = JSON.parse(e.postData.contents);
    const records = Array.isArray(payload) ? payload : [payload];
    if (!records.length || records.some(function (record) {
      return !record || typeof record !== 'object' || Array.isArray(record) ||
        Object.keys(record).length === 0;
    })) {
      throw new Error('Send a JSON object or an array of JSON objects.');
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
      const lastColumn = sheet.getLastColumn();
      const headers = lastColumn
        ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String)
        : [];

      records.forEach(function (record) {
        Object.keys(record).forEach(function (key) {
          if (headers.indexOf(key) === -1) headers.push(key);
        });
      });

      if (lastColumn === 0 || headers.length > lastColumn) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }

      const rows = records.map(function (record) {
        return headers.map(function (header) {
          const value = record[header];
          if (value === null || value === undefined) return '';
          return typeof value === 'object' ? JSON.stringify(value) : value;
        });
      });

      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
        .setValues(rows);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse({ ok: true, rowsAdded: records.length });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}