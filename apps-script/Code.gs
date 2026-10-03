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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
    if (!archive) archive = ss.insertSheet(CFG.ARCHIVE);
    ensureArchiveLayout_(archive);
    const rawCurrent = sh.getRange(row, 1, 1, CFG.MASTER_INPUT_COLS).getValues()[0];
    const archiveRow = Math.max(2, archive.getLastRow() + 1);
    // Write and verify the recoverable copy before clearing the live row.
    archive.getRange(archiveRow, 1, 1, 12).setValues([[new Date(), CFG.MASTER, row, ...rawCurrent]]);
    SpreadsheetApp.flush();
    const verified = archive.getRange(archiveRow, 4, 1, CFG.MASTER_INPUT_COLS).getValues()[0];
    if (!sameCells_(verified, rawCurrent)) {
      archive.getRange(archiveRow, 1, 1, 13).clearContent();
      SpreadsheetApp.flush();
      throw new Error('Archive copy verification failed for one or more fields. The Master Entry row was left untouched.');
    }
    // Clear only user-input columns A:I. Never delete the row or touch formula columns J:AC.
    sh.getRange(row, 1, 1, CFG.MASTER_INPUT_COLS).clearContent();
    SpreadsheetApp.flush();
    return {ok: true, id: current[0], row, message: 'Archived recoverably; input cells A:I cleared, formula columns untouched.'};
  });
}


function listArchive(limit) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(CFG.ARCHIVE);
  if (!sh || sh.getLastRow() < 2) return [];
  ensureArchiveLayout_(sh);
  const cap = Math.max(1, Math.min(Number(limit) || 50, 200));
  const last = sh.getLastRow(), start = Math.max(2, last - cap + 1);
  return sh.getRange(start, 1, last - start + 1, 13).getDisplayValues().map((r,i)=>({
    archiveRow:start+i, archivedAt:r[0], originalSheet:r[1], originalRow:r[2],
    id:r[3], entryDate:r[4], subject:r[5], unit:r[6], question:r[7],
    answer:r[8], source:r[9], notes:r[10], priority:r[11],
    snapshot:r.slice(3,12), archiveStatus:r[12] || ''
  })).filter(r=>!r.archiveStatus).reverse();
}

function restoreArchived(input) {
  return withLock_(() => {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const archive = ss.getSheetByName(CFG.ARCHIVE);
    const master = requireSheet_(ss, CFG.MASTER);
    const ar = Number(input && input.archiveRow);
    if (!archive || !Number.isInteger(ar) || ar < 2 || ar > archive.getLastRow()) throw new Error('Invalid archive row.');
    ensureArchiveLayout_(archive);
    const saved = archive.getRange(ar, 1, 1, 13).getValues()[0];
    if (String(saved[3]) !== String(input.id)) throw new Error('Archive entry changed. Refresh the archive list.');
    if (String(saved[12] || '').startsWith('RESTORED ')) throw new Error('This archive entry is already marked restored. Reload the archive list.');
    const original = saved.slice(3,12);
    if (!original[0]) throw new Error('Archive entry has no IMP ID.');
    const last = lastDataRow_(master, CFG.MASTER_FIRST_ROW, 1);
    const masterValues = last >= CFG.MASTER_FIRST_ROW
      ? master.getRange(CFG.MASTER_FIRST_ROW, 1, last-CFG.MASTER_FIRST_ROW+1, CFG.MASTER_INPUT_COLS).getValues()
      : [];
    const duplicateIndex = masterValues.findIndex(r => String(r[0]) === String(original[0]));
    if (duplicateIndex >= 0) {
      const existing = masterValues[duplicateIndex];
      if (!sameCells_(existing, original)) {
        throw new Error('This IMP ID already exists with different content. Nothing was changed; compare Master Entry and the archive before deciding which copy to keep.');
      }
      const status = 'RESTORED '+new Date().toISOString();
      archive.getRange(ar,13).setValue(status);
      SpreadsheetApp.flush();
      return {ok:true,id:original[0],row:CFG.MASTER_FIRST_ROW+duplicateIndex,message:'This exact IMP already exists in Master Entry. No duplicate was created; the archive entry was marked recovered.'};
    }
    const originalRow = Number(saved[2]);
    const originalRowValues = Number.isInteger(originalRow) && originalRow >= CFG.MASTER_FIRST_ROW && originalRow <= master.getMaxRows()
      ? master.getRange(originalRow,1,1,CFG.MASTER_INPUT_COLS).getDisplayValues()[0] : null;
    const target = originalRowValues && originalRowValues.every(v=>!String(v).trim())
      ? originalRow : Math.max(CFG.MASTER_FIRST_ROW,last+1);
    master.getRange(target,1,1,CFG.MASTER_INPUT_COLS).setValues([original]);
    SpreadsheetApp.flush();
    const restoredCheck = master.getRange(target,1,1,CFG.MASTER_INPUT_COLS).getValues()[0];
    if (!sameCells_(restoredCheck, original)) throw new Error('Restore verification failed for one or more fields. The archive copy was retained; check Master Entry before retrying.');
    // Mark only after all nine restored input fields have been verified.
    archive.getRange(ar,13).setValue('RESTORED '+new Date().toISOString());
    SpreadsheetApp.flush();
    return {ok:true,id:original[0],row:target,message:'Restored and verified in Master Entry. Archive copy retained; formula columns were not written.'};
  });
}

function addRevision(input) {
  return withLock_(() => {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rev = requireSheet_(ss, CFG.REVISION);
    const id = String(input && input.id || '').trim();
    const resultInput = String(input && input.result || '').trim().toLowerCase();
    const resultMap = {pass: 'Pass', fail: 'Fail', partial: 'Partial'};
    const result = resultMap[resultInput] || '';
    const notes = String(input && input.notes || '').trim();
    if (!id) throw new Error('Choose an IMP ID.');
    if (!['Pass','Fail','Partial'].includes(result)) throw new Error('Result must be Pass, Fail, or Partial.');
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
  const ids=last>=CFG.MASTER_FIRST_ROW
    ? sh.getRange(CFG.MASTER_FIRST_ROW,1,last-CFG.MASTER_FIRST_ROW+1,1).getDisplayValues().flat().map(Number).filter(n=>Number.isSafeInteger(n)&&n>0)
    : [];
  // Include archived IDs so archiving the highest ID can never cause ID reuse.
  const archive=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CFG.ARCHIVE);
  if(archive && archive.getLastRow()>=2) {
    const archived=archive.getRange(2,4,archive.getLastRow()-1,1).getDisplayValues().flat().map(Number).filter(n=>Number.isSafeInteger(n)&&n>0);
    ids.push(...archived);
  }
  return (ids.length?Math.max(...ids):0)+1;
}
function lastDataRow_(sh, firstRow, col) {
  // One batched read is faster than dozens of small range reads on mobile workflows.
  // Scan in memory because the Revision Log may have formulas prefilled to row 10,000.
  const last = sh.getLastRow();
  if (last < firstRow) return firstRow - 1;
  const vals = sh.getRange(firstRow, col, last - firstRow + 1, 1).getDisplayValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return firstRow + i;
  }
  return firstRow - 1;
}
function ensureArchiveLayout_(sh) {
  // Older archive tabs may have only A:L. Restore status uses column M.
  if (sh.getMaxColumns() < 13) sh.insertColumnsAfter(sh.getMaxColumns(), 13 - sh.getMaxColumns());
  const headers = ['Archived At','Original Sheet','Original Row','IMP ID','Entry Date','Subject','Unit','Question','Answer Key Points','Source','Notes','Priority','Archive Status'];
  const current = sh.getRange(1, 1, 1, 13).getDisplayValues()[0];
  if (current.every(v => !String(v).trim())) sh.getRange(1, 1, 1, 13).setValues([headers]);
  else if (!String(current[12] || '').trim()) sh.getRange(1, 13).setValue(headers[12]);
}
function sameCells_(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
  const key = v => v instanceof Date ? 'DATE:' + v.getTime() : (v === null || v === undefined ? '' : String(v));
  return a.every((v, i) => key(v) === key(b[i]));
}
function dateValue_(input, existingRaw, existingDisplay) {
  const value = String(input || '').trim();
  if (!value) return existingRaw || new Date();
  if (existingDisplay && value === String(existingDisplay)) return existingRaw;
  // Native date input uses yyyy-mm-dd; parse locally to avoid timezone date shifts.
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]), 12, 0, 0);
  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) return parsed;
  throw new Error('Entry Date must be a valid date.');
}
function requireSheet_(ss,name){if(!ss)throw new Error('Spreadsheet context unavailable.');const sh=ss.getSheetByName(name);if(!sh)throw new Error('Required tab not found: '+name);return sh;}
function withLock_(fn){const lock=LockService.getScriptLock();if(!lock.tryLock(CFG.LOCK_MS))throw new Error('Another save is in progress. Wait a moment and retry.');try{return fn();}finally{lock.releaseLock();}}


/**
 * Returns the full subject/unit catalog from Units Config, not just subjects
 * that already have an IMP in Master Entry. Existing Master Entry values are
 * unioned in so legacy units remain selectable while the catalog is maintained.
 * Units Config layout: subject headers in row 1 (A:L), units in rows 2:201.
 */
function getSubjectUnitOptions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = requireSheet_(ss, CFG.MASTER);
  const config = ss.getSheetByName('Units Config');
  if (!config) throw new Error('Required tab not found: Units Config. Subject and unit choices must come from the workbook catalog, not only from test IMPs.');
  const headers = config.getRange(1, 1, 1, Math.min(12, config.getMaxColumns())).getDisplayValues()[0];
  const catalogRows = config.getRange(2, 1, Math.min(200, Math.max(1, config.getMaxRows() - 1)), headers.length).getDisplayValues();
  const unitsBySubject = {};
  const subjects = new Set();
  headers.forEach((value, col) => {
    const subject = String(value || '').trim();
    if (!subject) return;
    subjects.add(subject);
    unitsBySubject[subject] = [];
    catalogRows.forEach(row => {
      const unit = String(row[col] || '').trim();
      if (unit && !unitsBySubject[subject].includes(unit)) unitsBySubject[subject].push(unit);
    });
  });

  // Preserve any subjects/units already used in records, even if a legacy
  // unit was removed from Units Config; this avoids breaking old IMP edits.
  const last = lastDataRow_(master, CFG.MASTER_FIRST_ROW, 1);
  if (last >= CFG.MASTER_FIRST_ROW) {
    const rows = master.getRange(CFG.MASTER_FIRST_ROW, 3, last-CFG.MASTER_FIRST_ROW+1, 2).getDisplayValues();
    rows.forEach(([subjectValue, unitValue]) => {
      const subject = String(subjectValue || '').trim();
      const unit = String(unitValue || '').trim();
      if (!subject) return;
      subjects.add(subject);
      if (!unitsBySubject[subject]) unitsBySubject[subject] = [];
      if (unit && !unitsBySubject[subject].includes(unit)) unitsBySubject[subject].push(unit);
    });
  }

  Object.keys(unitsBySubject).forEach(subject => unitsBySubject[subject].sort((a,b)=>a.localeCompare(b)));
  return {
    subjects: Array.from(subjects).sort((a,b)=>a.localeCompare(b)),
    unitsBySubject: unitsBySubject
  };
}
