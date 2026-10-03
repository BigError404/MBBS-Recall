/* MBBS Recall Pro — staged Google Sheets read-only UI.
 * This is intentionally a connection/inspection test only. It never writes to Sheets
 * and never imports remote rows into local study data.
 */
(() => {
  'use strict';
  const DEFAULT_CLIENT_ID = '237396931415-v9qn32rpikjngn6tba2ljsv260rcp6aj.apps.googleusercontent.com';
  const DEFAULT_SPREADSHEET_ID = '1wSA5eYsTYCZE-q2CdGOiQshrr-bdTE-63Lo_iw3lBLE';
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  const cfg = { clientId: DEFAULT_CLIENT_ID, spreadsheetId: DEFAULT_SPREADSHEET_ID, connected: false, sheets: [] };
  let busy = false;

  function panel() {
    if (document.querySelector('#googleSheetsPanel')) return;
    const settings = document.querySelector('#view');
    if (!settings || !settings.querySelector('h2')) return;
    const section = document.createElement('section');
    section.className = 'panel';
    section.id = 'googleSheetsPanel';
    section.innerHTML = `
      <h2>Google Sheets connection — read-only test</h2>
      <div class="notice warn">Safety mode: this only checks access and reads sheet names. It does not modify your Google Sheet or overwrite local study data. Your access token stays in memory and is not saved.</div>
      <div class="formgrid">
        <label class="formfield full">OAuth Client ID<input id="gsClientId" class="field" autocomplete="off" spellcheck="false" value="${esc(cfg.clientId)}"></label>
        <label class="formfield full">Google Spreadsheet ID<input id="gsSpreadsheetId" class="field" autocomplete="off" spellcheck="false" value="${esc(cfg.spreadsheetId)}"></label>
      </div>
      <div class="toolbar" style="margin-top:10px">
        <button class="btn primary" id="gsConnect" type="button">Connect Google account</button>
        <button class="btn" id="gsInspect" type="button" disabled>Check workbook tabs</button>
        <button class="btn" id="gsDisconnect" type="button" disabled>Disconnect</button>
      </div>
      <div id="gsStatus" class="notice" role="status">Not connected. Connect while online; Google may ask you to approve read-only access.</div>
      <div class="toolbar" style="margin-top:10px"><button class="btn" id="gsHeaders" type="button" disabled>Inspect key sheet headers (read-only)</button></div>
      <div id="gsHeaderResults"></div>
      <div id="gsSheetList"></div>
      <p class="footnote">The connection requests read-only spreadsheet access. If Google shows an OAuth test-user warning, confirm you added the Google account that owns the spreadsheet as a test user in the same Cloud project.</p>`;
    const danger = Array.from(settings.querySelectorAll('h2')).find(h => h.textContent.includes('Danger zone'));
    (danger && danger.closest('section')) ? danger.closest('section').before(section) : settings.append(section);
    section.querySelector('#gsConnect').addEventListener('click', connect);
    section.querySelector('#gsInspect').addEventListener('click', inspect);
    section.querySelector('#gsDisconnect').addEventListener('click', disconnect);
    section.querySelector('#gsHeaders').addEventListener('click', inspectHeaders);
  }
  function status(message, type='') {
    const el = document.querySelector('#gsStatus');
    if (el) { el.className = 'notice' + (type ? ' ' + type : ''); el.textContent = message; }
  }
  function inputs() {
    cfg.clientId = document.querySelector('#gsClientId')?.value.trim() || '';
    cfg.spreadsheetId = document.querySelector('#gsSpreadsheetId')?.value.trim() || '';
  }
  function controls() {
    const connected = Boolean(window.MBBSGoogleSheets?.isConnected());
    const connect = document.querySelector('#gsConnect'), inspect = document.querySelector('#gsInspect'), disconnectBtn = document.querySelector('#gsDisconnect'), headersBtn = document.querySelector('#gsHeaders');
    if (connect) { connect.disabled = busy; connect.textContent = connected ? 'Reconnect Google account' : 'Connect Google account'; }
    if (inspect) inspect.disabled = busy || !connected;
    if (disconnectBtn) disconnectBtn.disabled = busy || !connected;
    if (headersBtn) headersBtn.disabled = busy || !connected;
  }
  async function connect() {
    if (busy) return;
    inputs();
    busy = true; controls(); status('Opening Google authorization…');
    try {
      if (!window.MBBSGoogleSheets) throw new Error('Google Sheets module did not load. Reload the app while online.');
      await window.MBBSGoogleSheets.connect(cfg.clientId);
      cfg.connected = true;
      status('Connected. Next, check workbook tabs. No data has been changed.', 'good');
    } catch (e) { status(e.message || String(e), 'error'); }
    finally { busy = false; controls(); }
  }
  async function inspect() {
    if (busy) return;
    inputs();
    busy = true; controls(); status('Reading spreadsheet metadata…');
    try {
      const book = await window.MBBSGoogleSheets.getSpreadsheet(cfg.spreadsheetId);
      cfg.sheets = book.sheets;
      const root = document.querySelector('#gsSheetList');
      if (root) root.innerHTML = `<div class="notice good">Access confirmed: <strong>${esc(book.title)}</strong> · ${book.sheets.length} tabs. This was a metadata-only read.</div><div class="tablewrap"><table><thead><tr><th>Tab</th><th>Rows</th><th>Columns</th></tr></thead><tbody>${book.sheets.map(s => `<tr><td>${esc(s.title)}</td><td>${esc(s.rowCount)}</td><td>${esc(s.columnCount)}</td></tr>`).join('')}</tbody></table></div><p><a href="${esc(book.spreadsheetUrl)}" target="_blank" rel="noopener noreferrer">Open Google Sheet</a></p>`;
      status('Workbook metadata loaded successfully. Still read-only; local data is untouched.', 'good');
    } catch (e) { status(e.message || String(e), 'error'); }
    finally { busy = false; controls(); }
  }

  async function inspectHeaders() {
    if (busy) return;
    inputs();
    busy = true; controls();
    status('Mapping workbook headers in one read-only batch…');
    const root = document.querySelector('#gsHeaderResults');
    try {
      if (!cfg.sheets.length) {
        const book = await window.MBBSGoogleSheets.getSpreadsheet(cfg.spreadsheetId);
        cfg.sheets = book.sheets;
      }
      // Read only rows 1–2 of every tab, in one batch request. Never reads study records.
      const ranges = [];
      const owners = [];
      for (const sheet of cfg.sheets) {
        const safeTitle = "'" + sheet.title.replace(/'/g, "''") + "'";
        ranges.push(safeTitle + '!1:2');
        owners.push(sheet);
      }
      const response = await window.MBBSGoogleSheets.readRanges(cfg.spreadsheetId, ranges);
      const valueRanges = response.valueRanges || [];
      const letter = i => {
        let n = i + 1, out = '';
        while (n > 0) { const rem = (n - 1) % 26; out = String.fromCharCode(65 + rem) + out; n = Math.floor((n - 1) / 26); }
        return out;
      };
      const nonEmpty = row => (row || []).filter(v => String(v ?? '').trim() !== '').length;
      const report = owners.map((sheet, i) => {
        const values = valueRanges[i]?.values || [];
        const row1 = values[0] || [], row2 = values[1] || [];
        const isDataTab = /^(Master Entry|Revision Log|TODAY DUE|THIS WEEK|Data Health|Dashboard|START HERE)$/i.test(sheet.title) ||
          /^(Medicine|Surgery|Obstetrics|Gynaecology|Pediatrics|Paediatrics|Pathology|Pharmacology|Microbiology|Anatomy|Physiology|Biochemistry|Radiology|Psychiatry|Community Medicine|Forensic Medicine)$/i.test(sheet.title);
        const preferredRow = /^Master Entry$/i.test(sheet.title) || /^(Medicine|Surgery|Obstetrics|Gynaecology|Pediatrics|Paediatrics|Pathology|Pharmacology|Microbiology|Anatomy|Physiology|Biochemistry|Radiology|Psychiatry|Community Medicine|Forensic Medicine)$/i.test(sheet.title) ? 2 : 1;
        const primary = preferredRow === 2 ? row2 : row1;
        const secondary = preferredRow === 2 ? row1 : row2;
        const primaryCount = nonEmpty(primary), secondaryCount = nonEmpty(secondary);
        const useRow = primaryCount ? preferredRow : (secondaryCount ? (preferredRow === 1 ? 2 : 1) : preferredRow);
        const chosen = useRow === 1 ? row1 : row2;
        const headers = chosen.map((value, col) => ({ col: letter(col), value: String(value ?? '').trim() })).filter(cell => cell.value);
        return { title: sheet.title, row: useRow, headers, row1Count: nonEmpty(row1), row2Count: nonEmpty(row2), isDataTab };
      });
      const key = report.filter(x => x.isDataTab);
      const other = report.filter(x => !x.isDataTab);
      // Mobile-first compact layout: never put one-character-wide table cells on screen.
      const render = item => '<section class="panel" style="margin:10px 0;padding:12px">' +
        '<h3 style="margin:0 0 6px;font-size:1rem">' + esc(item.title) + '</h3>' +
        '<p class="footnote" style="margin:0 0 8px">Candidate headers: row ' + item.row +
        ' · filled cells: row 1=' + item.row1Count + ', row 2=' + item.row2Count + '</p>' +
        (item.headers.length ? '<div style="display:flex;flex-wrap:wrap;gap:6px">' +
          item.headers.map(cell => '<span style="display:inline-block;max-width:100%;padding:6px 9px;border:1px solid var(--line,#d5dde5);border-radius:8px;overflow-wrap:anywhere;background:var(--surface,#fff)"><strong>' +
          esc(cell.col) + '</strong> — ' + esc(cell.value) + '</span>').join('') +
          '</div>' : '<p>No non-empty cells in rows 1–2.</p>') +
        '</section>';
      if (root) root.innerHTML = '<div class="notice good">Header check complete: rows 1–2 inspected across ' + report.length + ' tabs in one read-only batch. Results are displayed as compact mobile-friendly labels. No study records were read, imported, or changed.</div>' +
        '<h3>Core study and dashboard tabs (' + key.length + ')</h3>' + key.map(render).join('') +
        '<details><summary>Other workbook tabs (' + other.length + ')</summary>' + other.map(render).join('') + '</details>';
      status('Workbook header map generated in one batch. No records were imported or changed.', 'good');
    } catch (e) {
      status(e.message || String(e), 'error');
    } finally { busy = false; controls(); }
  }
  function disconnect() {
    window.MBBSGoogleSheets?.disconnect();
    cfg.connected = false;
    status('Disconnected. No spreadsheet changes were made.');
    controls();
  }
  const observer = new MutationObserver(() => {
    if (document.querySelector('#view') && document.querySelector('#view h1')?.textContent?.includes('Settings')) panel();
  });
  function start() {
    observer.observe(document.querySelector('#view') || document.body, { childList: true, subtree: true });
    if (document.querySelector('#view h1')?.textContent?.includes('Settings')) panel();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();