/** MBBS Recall Pro — bound Google Apps Script backend.
 * Deploy as a web app owned by the spreadsheet owner. All browser calls use
 * google.script.run (same Apps Script origin); no OAuth tokens or API secrets in the PWA.
 * Master Entry: row 2 headers, data starts row 3, A:AC (29 columns).
 * Revision Log: row 1 headers, data starts row 2, A:K (11 columns).
 */
const SPREADSHEET_ID = '1wSA5eYsTYCZE-q2CdGOiQshrr-bdTE-63Lo_iw3lBLE';
const CFG = Object.freeze({
  MASTER: 'Master Entry', MASTER_HEADER_ROW: 2, MASTER_FIRST_ROW: 3, MASTER_INPUT_COLS: 9,
  REVISION: 'Revision Log', REVISION_HEADER_ROW: 1, REVISION_FIRST_ROW: 2, REVISION_INPUT_COLS: 4,
  ARCHIVE: '_MBBS Recall Archive', LOCK_MS: 15000
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MBBS Recall Pro — Sheet Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getBootstrap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) throw new Error('Bind this Apps Script project to the MBBS Recall Google Sheet.');
  const master = requireSheet_(ss, CFG.MASTER);
  const revision = requireSheet_(ss, CFG.REVISION);
  return {
    title: ss.getName(),
    masterHeaders: master.getRange(CFG.MASTER_HEADER_ROW, 1, 1, CFG.MASTER_INPUT_COLS).getDisplayValues()[0],
    revisionHeaders: revision.getRange(CFG.REVISION_HEADER_ROW, 1, 1, CFG.REVISION_INPUT_COLS).getDisplayValues()[0],
    masterCount: Math.max(0, lastDataRow_(master, CFG.MASTER_FIRST_ROW, 1) - CFG.MASTER_FIRST_ROW + 1),
    revisionCount: Math.max(0, lastDataRow_(revision, CFG.REVISION_FIRST_ROW, 1) - CFG.REVISION_FIRST_ROW + 1),
    mode: 'Connected to bound spreadsheet; server-side writes; formulas outside input columns are protected.'
  };
}

function listRecords(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = requireSheet_(ss, CFG.MASTER);
  const last = lastDataRow_(sh, CFG.MASTER_FIRST_ROW, 1);
  if (last < CFG.MASTER_FIRST_ROW) return [];
  const cap = Math.max(1, Math.min(Number(limit) || 100, 300));
  const start = Math.max(CFG.MASTER_FIRST_ROW, last - cap + 1);
  const vals = sh.getRange(start, 1, last - start + 1, CFG.MASTER_INPUT_COLS).getDisplayValues();
  return vals.map((r, i) => ({row: start + i, id: r[0], entryDate: r[1], subject: r[2], unit: r[3], question: r[4], answer: r[5], source: r[6], notes: r[7], priority: r[8], snapshot: r.slice(0, CFG.MASTER_INPUT_COLS)}))
    .filter(r => r.snapshot.some(v => String(v).trim() !== '')).reverse();
}

function addRecord(input) {
  return withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = requireSheet_(ss, CFG.MASTER);
    const data = validateRecord_(input);
    const last = lastDataRow_(sh, CFG.MASTER_FIRST_ROW, 1);
    const row = Math.max(CFG.MASTER_FIRST_ROW, last + 1);
    const id = nextId_(sh);
    const values = [[id, dateValue_(data.entryDate, null, ''), data.subject, data.unit, data.question, data.answer, data.source, data.notes, data.priority]];
    sh.getRange(row, 1, 1, CFG.MASTER_INPUT_COLS).setValues(values);
    SpreadsheetApp.flush();
    return {ok: true, id, row, message: 'IMP saved to Master Entry. Calculated columns were not written.'};
  });
}

function editRecord(input) {
  return withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = requireSheet_(ss, CFG.MASTER);
    const row = Number(input && input.row);
    if (!Number.isInteger(row) || row < CFG.MASTER_FIRST_ROW || row > sh.getMaxRows()) throw new Error('Invalid source row.');
    const current = sh.getRange(row, 1, 1, CFG.MASTER_INPUT_COLS).getDisplayValues()[0];
    if (String(current[0]) !== String(input.id)) throw new Error('Conflict: this row no longer contains that IMP ID. Reload before editing.');
    if (!Array.isArray(input.expectedSnapshot) || input.expectedSnapshot.length !== CFG.MASTER_INPUT_COLS ||
        current.some((v, i) => String(v) !== String(input.expectedSnapshot[i] ?? ''))) {
      throw new Error('Conflict: this IMP changed since it was opened. Reload the record; no changes were saved.');
    }
    const data = validateRecord_(input.record);
    const values = [[current[0], dateValue_(data.entryDate, sh.getRange(row, 2).getValue(), current[1]), data.subject, data.unit, data.question, data.answer, data.source, data.notes, data.priority]];
    sh.getRange(row, 1, 1, CFG.MASTER_INPUT_COLS).setValues(values);
    SpreadsheetApp.flush();
    return {ok: true, id: current[0], row, message: 'IMP updated. ID and calculated columns were preserved.'};
  });
}

function archiveRecord(input) {
  return withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = requireSheet_(ss, CFG.MASTER);
    const row = Number(input && input.row);
    if (!Number.isInteger(row) || row < CFG.MASTER_FIRST_ROW || row > sh.getMaxRows()) throw new Error('Invalid source row.');
    const current = sh.getRange(row, 1, 1, CFG.MASTER_INPUT_COLS).getDisplayValues()[0];
    if (String(current[0]) !== String(input.id)) throw new Error('Conflict: row ID changed. Reload before archiving.');
    if (!Array.isArray(input.expectedSnapshot) || input.expectedSnapshot.length !== CFG.MASTER_INPUT_COLS ||
        current.some((v, i) => String(v) !== String(input.expectedSnapshot[i] ?? ''))) {
      throw new Error('Conflict: this IMP changed since it was opened. Nothing was archived.');
    }
    if (!current.some(v => String(v).trim())) throw new Error('This row is already empty.');
    let archive = ss.getSheetByName(CFG.ARCHIVE);
    if (!archive) {
      archive = ss.insertSheet(CFG.ARCHIVE);
      archive.getRange(1, 1, 1, 11).setValues([['Archived At','Original Sheet','Original Row','IMP ID','Entry Date','Subject','Unit','Question','Answer Key Points','Source','Notes / Priority']]);
      archive.hideSheet();
    }
    archive.appendRow([new Date(), CFG.MASTER, row, current[0], current[1], current[2], current[3], current[4], current[5], current[6], 'Notes: '+current[7]+' | Priority: '+current[8]]);
    // Clear only user-input columns A:I. Never delete the row or touch formula columns J:AC.
    sh.getRange(row, 1, 1, CFG.MASTER_INPUT_COLS).clearContent();
    SpreadsheetApp.flush();
    return {ok: true, id: current[0], row, message: 'Archived recoverably; input cells A:I cleared, formula columns untouched.'};
  });
}

function addRevision(input) {
  return withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rev = requireSheet_(ss, CFG.REVISION);
    const id = String(input && input.id || '').trim();
    const result = String(input && input.result || '').trim().toUpperCase();
    const notes = String(input && input.notes || '').trim();
    if (!id) throw new Error('Choose an IMP ID.');
    if (!['PASS','FAIL','PARTIAL'].includes(result)) throw new Error('Result must be Pass, Fail, or Partial.');
    const master = requireSheet_(ss, CFG.MASTER);
    const lastMaster = lastDataRow_(master, CFG.MASTER_FIRST_ROW, 1);
    const ids = lastMaster >= CFG.MASTER_FIRST_ROW ? master.getRange(CFG.MASTER_FIRST_ROW,1,lastMaster-CFG.MASTER_FIRST_ROW+1,1).getDisplayValues().flat() : [];
    if (!ids.some(v => String(v) === id)) throw new Error('IMP ID not found in Master Entry.');
    const row = Math.max(CFG.REVISION_FIRST_ROW, lastDataRow_(rev, CFG.REVISION_FIRST_ROW, 1) + 1);
    rev.getRange(row, 1, 1, CFG.REVISION_INPUT_COLS).setValues([[input.date || new Date(), id, result, notes]]);
    SpreadsheetApp.flush();
    return {ok: true, row, message: 'Revision saved. Formula columns E:K were not written.'};
  });
}

function getRecentRevisions(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = requireSheet_(ss, CFG.REVISION);
  const last = lastDataRow_(sh, CFG.REVISION_FIRST_ROW, 1);
  if (last < CFG.REVISION_FIRST_ROW) return [];
  const cap = Math.max(1, Math.min(Number(limit) || 50, 200));
  const start = Math.max(CFG.REVISION_FIRST_ROW, last - cap + 1);
  const vals = sh.getRange(start, 1, last - start + 1, CFG.REVISION_INPUT_COLS).getDisplayValues();
  return vals.map((r,i)=>({row:start+i,date:r[0],id:r[1],result:r[2],notes:r[3]})).filter(r=>r.date||r.id||r.result||r.notes).reverse();
}

function validateRecord_(x) {
  if (!x || typeof x !== 'object') throw new Error('Record fields are missing.');
  const clean = k => String(x[k] ?? '').trim();
  const subject=clean('subject'), unit=clean('unit'), question=clean('question');
  if (!subject) throw new Error('Subject is required.');
  if (!question) throw new Error('Question / IMP is required.');
  const priority=clean('priority') || 'Regular';
  if (!['Regular','Must-Do','Late'].includes(priority)) throw new Error('Priority must be Regular, Must-Do, or Late.');
  return {entryDate: clean('entryDate'), subject, unit, question, answer:clean('answer'), source:clean('source'), notes:clean('notes'), priority};
}
function nextId_(sh) {
  const last=lastDataRow_(sh,CFG.MASTER_FIRST_ROW,1);
  if(last<CFG.MASTER_FIRST_ROW)return 1;
  const ids=sh.getRange(CFG.MASTER_FIRST_ROW,1,last-CFG.MASTER_FIRST_ROW+1,1).getDisplayValues().flat().map(v=>Number(v)).filter(n=>Number.isSafeInteger(n)&&n>0);
  return (ids.length?Math.max(...ids):0)+1;
}
function lastDataRow_(sh, firstRow, col) {
  const last=sh.getLastRow();
  if(last<firstRow)return firstRow-1;
  const vals=sh.getRange(firstRow,col,last-firstRow+1,1).getDisplayValues();
  for(let i=vals.length-1;i>=0;i--)if(String(vals[i][0]).trim()!=='')return firstRow+i;
  return firstRow-1;
}
function dateValue_(input, existingRaw, existingDisplay) {
  const value = String(input || '').trim();
  if (!value) return existingRaw || new Date();
  if (existingDisplay && value === String(existingDisplay)) return existingRaw;
  // Native date input uses yyyy-mm-dd; parse locally to avoid timezone date shifts.
  const m = /^(\\d{4})-(\\d{2})-(\\d{2})$/.exec(value);
  if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]), 12, 0, 0);
  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return parsed;
  throw new Error('Entry Date must be a valid date.');
}
function requireSheet_(ss,name){if(!ss)throw new Error('Spreadsheet context unavailable.');const sh=ss.getSheetByName(name);if(!sh)throw new Error('Required tab not found: '+name);return sh;}
function withLock_(fn){const lock=LockService.getScriptLock();if(!lock.tryLock(CFG.LOCK_MS))throw new Error('Another save is in progress. Wait a moment and retry.');try{return fn();}finally{lock.releaseLock();}}
