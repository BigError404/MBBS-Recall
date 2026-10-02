/**
 * MBBS Recall Pro — Google Apps Script backend (safe migration prototype)
 *
 * Deploy as a Web App:
 *   Execute as: Me
 *   Who has access: Only myself
 *
 * This is intentionally read-only. It never edits cells, formats, formulas,
 * validations, or records. Use the bound spreadsheet ID only.
 */
const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1wSA5eYsTYCZE-q2CdGOiQshrr-bdTE-63Lo_iw3lBLE',
  MASTER_SHEET: 'Master Entry',
  REVISION_SHEET: 'Revision Log',
  MAX_RECORDS: 5000
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MBBS Recall Pro — Sheets Backend')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Health check: returns metadata only. */
function getBackendStatus() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return {
    ok: true,
    title: ss.getName(),
    spreadsheetIdSuffix: CONFIG.SPREADSHEET_ID.slice(-6),
    sheets: ss.getSheets().map(s => ({
      name: s.getName(),
      lastRow: s.getLastRow(),
      lastColumn: s.getLastColumn()
    })),
    mode: 'READ_ONLY'
  };
}

/**
 * Return non-empty Master Entry records, reading only A:I input columns.
 * Formula/calculated columns J:AC are not read or modified.
 */
function getMasterRecords(limit) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MASTER_SHEET);
  if (!sheet) throw new Error('Master Entry sheet was not found.');
  const lastRow = Math.min(sheet.getLastRow(), CONFIG.MAX_RECORDS + 2);
  if (lastRow < 3) return { headers: [], records: [], count: 0, truncated: false };

  const headerValues = sheet.getRange(2, 1, 1, 9).getDisplayValues()[0];
  const rows = sheet.getRange(3, 1, lastRow - 2, 9).getDisplayValues();
  const records = rows.map((r, i) => ({ sheetRow: i + 3, values: r }))
    .filter(x => x.values.some(v => String(v).trim() !== ''));
  const safeLimit = Math.max(1, Math.min(Number(limit) || 200, 1000));
  return {
    headers: headerValues,
    records: records.slice(0, safeLimit),
    count: records.length,
    truncated: records.length > safeLimit || sheet.getLastRow() > lastRow
  };
}

/** Return recent revision-log rows as display values only. */
function getRecentRevisions(limit) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.REVISION_SHEET);
  if (!sheet) throw new Error('Revision Log sheet was not found.');
  const lastRow = Math.min(sheet.getLastRow(), CONFIG.MAX_RECORDS + 1);
  if (lastRow < 2) return { headers: [], records: [], count: 0, truncated: false };
  const width = Math.min(sheet.getLastColumn(), 11);
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, width).getDisplayValues();
  const records = rows.map((r, i) => ({ sheetRow: i + 2, values: r }))
    .filter(x => x.values.some(v => String(v).trim() !== ''));
  const safeLimit = Math.max(1, Math.min(Number(limit) || 100, 500));
  return {
    headers,
    records: records.slice(-safeLimit).reverse(),
    count: records.length,
    truncated: records.length > safeLimit || sheet.getLastRow() > lastRow
  };
}
