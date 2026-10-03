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

function doGet(e) {
  if (e && e.parameter && e.parameter.bridge === '1') {
    return HtmlService.createHtmlOutputFromFile('Bridge')
      .setTitle('MBBS Recall Pro — Secure Sheets Bridge')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
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


/**
 * Full read-only snapshot for the offline-first PWA sync layer.
 * Reads only user-input columns and catalog values; formulas are never exported
 * as writable data. The bridge exposes this function only through an allowlist.
 */
function getSyncSnapshot() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const master = requireSheet_(ss, CFG.MASTER);
  const revision = requireSheet_(ss, CFG.REVISION);
  const masterLast = lastDataRow_(master, CFG.MASTER_FIRST_ROW, 1);
  const masterRows = masterLast >= CFG.MASTER_FIRST_ROW
    ? master.getRange(CFG.MASTER_FIRST_ROW, 1, masterLast-CFG.MASTER_FIRST_ROW+1, CFG.MASTER_INPUT_COLS).getValues() : [];
  const masterDisplay = masterLast >= CFG.MASTER_FIRST_ROW
    ? master.getRange(CFG.MASTER_FIRST_ROW, 1, masterLast-CFG.MASTER_FIRST_ROW+1, CFG.MASTER_INPUT_COLS).getDisplayValues() : [];
  const cards = [];
  masterRows.forEach((r, i) => {
    if (!r.some(v => String(v === null || v === undefined ? '' : v).trim() !== '')) return;
    const d = masterDisplay[i];
    cards.push({
      row: CFG.MASTER_FIRST_ROW+i, id: String(d[0] || ''),
      entryDate: isoDate_(r[1], d[1]), subject: String(d[2] || ''),
      unit: String(d[3] || ''), question: String(d[4] || ''),
      answer: String(d[5] || ''), source: String(d[6] || ''),
      notes: String(d[7] || ''), priority: String(d[8] || 'Regular'),
      snapshot: d.slice(0, CFG.MASTER_INPUT_COLS)
    });
  });
  const revLast = lastDataRow_(revision, CFG.REVISION_FIRST_ROW, 1);
  const revRows = revLast >= CFG.REVISION_FIRST_ROW
    ? revision.getRange(CFG.REVISION_FIRST_ROW, 1, revLast-CFG.REVISION_FIRST_ROW+1, CFG.REVISION_INPUT_COLS).getValues() : [];
  const revDisplay = revLast >= CFG.REVISION_FIRST_ROW
    ? revision.getRange(CFG.REVISION_FIRST_ROW, 1, revLast-CFG.REVISION_FIRST_ROW+1, CFG.REVISION_INPUT_COLS).getDisplayValues() : [];
  const reviews = [];
  revRows.forEach((r, i) => {
    if (!r.some(v => String(v === null || v === undefined ? '' : v).trim() !== '')) return;
    const d = revDisplay[i];
    reviews.push({
      row: CFG.REVISION_FIRST_ROW+i, logDate: isoDate_(r[0], d[0]),
      impId: String(d[1] || ''), result: String(d[2] || '').trim().toUpperCase(),
      notes: String(d[3] || ''), snapshot: d.slice(0, CFG.REVISION_INPUT_COLS)
    });
  });
  const config = requireSheet_(ss, 'Units Config');
  const width = Math.min(12, Math.max(1, config.getMaxColumns()));
  const headers = config.getRange(1, 1, 1, width).getDisplayValues()[0];
  const rowCount = Math.min(200, Math.max(1, config.getMaxRows()-1));
  const unitsRows = config.getRange(2, 1, rowCount, width).getDisplayValues();
  const unitsBySubject = {};
  headers.forEach((h, col) => {
    const subject = String(h || '').trim();
    if (!subject) return;
    unitsBySubject[subject] = [];
    unitsRows.forEach(row => {
      const unit = String(row[col] || '').trim();
      if (unit && !unitsBySubject[subject].includes(unit)) unitsBySubject[subject].push(unit);
    });
  });
  Object.keys(unitsBySubject).forEach(s => unitsBySubject[s].sort((a,b)=>a.localeCompare(b)));
  return {title:ss.getName(), fetchedAt:new Date().toISOString(), master:cards, revisions:reviews, units:unitsBySubject};
}

/**
 * Applies a reviewed change plan under one script lock. Updates require the
 * exact nine-cell snapshot read earlier. New cards and revisions use stable
 * client keys so retrying after a lost browser response does not duplicate them.
 * Deletions are intentionally unsupported; use the recoverable archive manager.
 */
function syncApplyChanges(plan) {
  return withLock_(() => {
    if (!plan || typeof plan !== 'object') throw new Error('Sync change plan is missing.');
    const adds = Array.isArray(plan.adds) ? plan.adds : [];
    const updates = Array.isArray(plan.updates) ? plan.updates : [];
    const revisions = Array.isArray(plan.revisions) ? plan.revisions : [];
    if (adds.length > 1000 || updates.length > 1000 || revisions.length > 5000) throw new Error('Sync batch exceeds the safe batch limit.');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const master = requireSheet_(ss, CFG.MASTER), rev = requireSheet_(ss, CFG.REVISION);
    const props = PropertiesService.getScriptProperties();
    const addKeys = new Set(), updateRows = new Set();
    const preparedAdds = adds.map(a => {
      const key = String(a && a.clientKey || '').trim();
      if (!/^[A-Za-z0-9_-]{1,100}$/.test(key) || addKeys.has(key)) throw new Error('Every new IMP needs a unique, stable sync key.');
      addKeys.add(key);
      const record = validateRecord_(a.record);
      return {key, record, existingId: props.getProperty('MBBS_SYNC_ADD_'+key)};
    });
    const preparedUpdates = updates.map(u => {
      const row = Number(u && u.row), id = String(u && u.id || '');
      if (!Number.isInteger(row) || row < CFG.MASTER_FIRST_ROW || row > master.getMaxRows()) throw new Error('Sync update has an invalid sheet row.');
      if (updateRows.has(row)) throw new Error('Sync plan contains duplicate updates for one sheet row.');
      updateRows.add(row);
      const current = master.getRange(row,1,1,CFG.MASTER_INPUT_COLS).getDisplayValues()[0];
      if (String(current[0]) !== id) throw new Error('Sync conflict: an IMP ID changed rows. Refresh the snapshot and retry.');
      if (!Array.isArray(u.expectedSnapshot) || u.expectedSnapshot.length !== CFG.MASTER_INPUT_COLS ||
          current.some((v,i)=>String(v)!==String(u.expectedSnapshot[i] ?? ''))) {
        throw new Error('Sync conflict: an IMP changed in Google Sheets after the preview. No sync writes were started.');
      }
      return {row,id,current,record:validateRecord_(u.record)};
    });
    const preparedRevisions = revisions.map(r => {
      const key = String(r && r.clientKey || '').trim();
      if (!/^[A-Za-z0-9_-]{1,100}$/.test(key)) throw new Error('Every new revision needs a stable sync key.');
      const date = String(r.logDate || '').trim(), resultInput = String(r.result || '').trim().toLowerCase();
      const resultMap = {pass:'Pass',fail:'Fail',partial:'Partial'};
      const result = resultMap[resultInput] || '';
      const dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
      const parsedDay = dm ? new Date(Number(dm[1]),Number(dm[2])-1,Number(dm[3]),12,0,0) : null;
      if (!dm || !parsedDay || parsedDay.getFullYear()!==Number(dm[1]) || parsedDay.getMonth()!==Number(dm[2])-1 ||
          parsedDay.getDate()!==Number(dm[3]) || !result) throw new Error('Sync revision has an invalid calendar date or result.');
      const impClientKey = String(r.impClientKey || '').trim();
      const impId = String(r.impId || '').trim();
      if (!impId && (!impClientKey || !addKeys.has(impClientKey))) throw new Error('Sync revision refers to an unknown IMP.');
      return {key, date, result, notes:String(r.notes || '').trim(), impId, impClientKey};
    });
    const activeIds = new Set();
    const masterLast = lastDataRow_(master, CFG.MASTER_FIRST_ROW, 1);
    if (masterLast >= CFG.MASTER_FIRST_ROW) master.getRange(CFG.MASTER_FIRST_ROW,1,masterLast-CFG.MASTER_FIRST_ROW+1,1).getDisplayValues().flat().forEach(v=>{if(v)activeIds.add(String(v));});
    preparedAdds.forEach(a => {
      if (a.existingId && !activeIds.has(String(a.existingId))) throw new Error('A previous sync add is recorded but its IMP is missing from Master Entry. Resolve the archive before retrying.');
      if (a.existingId) {
        const idValues = master.getRange(CFG.MASTER_FIRST_ROW,1,Math.max(0,masterLast-CFG.MASTER_FIRST_ROW+1),1).getDisplayValues().flat();
        const idx = idValues.findIndex(v=>String(v)===String(a.existingId));
        if (idx < 0) throw new Error('A previous sync add could not be located. Refresh the snapshot before retrying.');
        const row = CFG.MASTER_FIRST_ROW+idx;
        const raw = master.getRange(row,1,1,CFG.MASTER_INPUT_COLS).getValues()[0];
        const shown = master.getRange(row,1,1,CFG.MASTER_INPUT_COLS).getDisplayValues()[0];
        const d = a.record;
        if (isoDate_(raw[1],shown[1]) !== d.entryDate || shown[2] !== d.subject || shown[3] !== d.unit ||
            shown[4] !== d.question || shown[5] !== d.answer || shown[6] !== d.source || shown[7] !== d.notes || shown[8] !== d.priority) {
          throw new Error('A previously started sync add now has different content. Refresh the Sheet and resolve the IMP before retrying; no duplicate was created.');
        }
      }
    });
    preparedUpdates.forEach(u => { if (!activeIds.has(u.id)) throw new Error('Sync update refers to an IMP that is no longer active.'); });
    preparedRevisions.forEach(r => {
      if (r.impId && !activeIds.has(r.impId) && !preparedAdds.some(a=>a.key===r.impId)) throw new Error('Sync revision refers to an IMP ID that is not active.');
      if (props.getProperty('MBBS_SYNC_REV_'+r.key)) return;
    });

    const idMap = {}, added = [], updated = [], revisionRows = [];
    preparedAdds.forEach(a => {
      if (a.existingId) { idMap[a.key] = String(a.existingId); return; }
      const id = nextId_(master), last = lastDataRow_(master, CFG.MASTER_FIRST_ROW, 1);
      const row = Math.max(CFG.MASTER_FIRST_ROW, last+1), d = a.record;
      master.getRange(row,1,1,CFG.MASTER_INPUT_COLS).setValues([[id,dateValue_(d.entryDate,null,''),d.subject,d.unit,d.question,d.answer,d.source,d.notes,d.priority]]);
      SpreadsheetApp.flush();
      const check = master.getRange(row,1,1,CFG.MASTER_INPUT_COLS).getDisplayValues()[0];
      if (String(check[0]) !== String(id) || check[2] !== d.subject || check[3] !== d.unit || check[4] !== d.question ||
          check[5] !== d.answer || check[6] !== d.source || check[7] !== d.notes || check[8] !== d.priority) {
        throw new Error('A new IMP could not be verified. Do not retry blindly; refresh the sheet snapshot first.');
      }
      props.setProperty('MBBS_SYNC_ADD_'+a.key, String(id));
      idMap[a.key] = String(id); added.push({clientKey:a.key,id:String(id),row});
    });
    preparedUpdates.forEach(u => {
      const d=u.record, oldDate=master.getRange(u.row,2).getValue();
      master.getRange(u.row,1,1,CFG.MASTER_INPUT_COLS).setValues([[u.id,dateValue_(d.entryDate,oldDate,u.current[1]),d.subject,d.unit,d.question,d.answer,d.source,d.notes,d.priority]]);
      SpreadsheetApp.flush();
      const check=master.getRange(u.row,1,1,CFG.MASTER_INPUT_COLS).getDisplayValues()[0];
      if(String(check[0])!==u.id || check[2]!==d.subject || check[3]!==d.unit || check[4]!==d.question || check[5]!==d.answer || check[6]!==d.source || check[7]!==d.notes || check[8]!==d.priority) throw new Error('An IMP update failed verification. Refresh the snapshot before retrying.');
      updated.push({id:u.id,row:u.row});
    });
    const revLast = lastDataRow_(rev, CFG.REVISION_FIRST_ROW, 1);
    const existingRevRows = revLast >= CFG.REVISION_FIRST_ROW
      ? rev.getRange(CFG.REVISION_FIRST_ROW,1,revLast-CFG.REVISION_FIRST_ROW+1,CFG.REVISION_INPUT_COLS).getValues() : [];
    const existingRevDisplay = revLast >= CFG.REVISION_FIRST_ROW
      ? rev.getRange(CFG.REVISION_FIRST_ROW,1,revLast-CFG.REVISION_FIRST_ROW+1,CFG.REVISION_INPUT_COLS).getDisplayValues() : [];
    const existingSigs = new Set();
    existingRevRows.forEach((r,i)=>existingSigs.add([isoDate_(r[0],existingRevDisplay[i][0]),String(existingRevDisplay[i][1]||''),String(existingRevDisplay[i][2]||'').trim().toUpperCase(),String(existingRevDisplay[i][3]||'').trim()].join('|')));
    preparedRevisions.forEach(r => {
      const oldId = r.impClientKey ? idMap[r.impClientKey] : r.impId;
      if (!oldId) throw new Error('A new revision could not be mapped to its saved IMP ID.');
      if (props.getProperty('MBBS_SYNC_REV_'+r.key)) { revisionRows.push({clientKey:r.key,duplicate:true}); return; }
      const sig=[r.date,String(oldId),r.result.toUpperCase(),r.notes].join('|');
      if (existingSigs.has(sig)) {
        props.setProperty('MBBS_SYNC_REV_'+r.key,'existing');
        revisionRows.push({clientKey:r.key,duplicate:true});
        return;
      }
      const row=Math.max(CFG.REVISION_FIRST_ROW,lastDataRow_(rev,CFG.REVISION_FIRST_ROW,1)+1);
      rev.getRange(row,1,1,CFG.REVISION_INPUT_COLS).setValues([[dateValue_(r.date,null,''),Number(oldId),r.result,r.notes]]);
      SpreadsheetApp.flush();
      const check=rev.getRange(row,1,1,CFG.REVISION_INPUT_COLS).getDisplayValues()[0];
      if(String(check[1])!==String(oldId) || String(check[2]).toUpperCase()!==r.result.toUpperCase() || String(check[3])!==r.notes) throw new Error('A revision write failed verification. Refresh the sheet snapshot before retrying.');
      props.setProperty('MBBS_SYNC_REV_'+r.key,'saved');
      existingSigs.add(sig); revisionRows.push({clientKey:r.key,row});
    });
    return {ok:true,idMap,added,updated,revisions:revisionRows,message:'Sync batch completed; each write was verified and formula columns were untouched.'};
  });
}

function isoDate_(raw, display) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return String(raw.getFullYear())+'-'+String(raw.getMonth()+1).padStart(2,'0')+'-'+String(raw.getDate()).padStart(2,'0');
  }
  const s=String(display || raw || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d=new Date(s);
  if (Number.isNaN(d.getTime())) return '';
  return String(d.getFullYear())+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
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
  if (m) {
    const year=Number(m[1]), month=Number(m[2])-1, day=Number(m[3]);
    const parsed=new Date(year,month,day,12,0,0);
    if(parsed.getFullYear()!==year || parsed.getMonth()!==month || parsed.getDate()!==day) throw new Error('Entry Date must be a real calendar date.');
    return parsed;
  }
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
