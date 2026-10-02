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
    status('Reading the actual header row from selected workbook tabs…');
    // Master Entry row 1 contains instructions, not field names. Its actual headers are on row 2.
    const targets = [
      { title: 'Master Entry', row: 2 },
      { title: 'Revision Log', row: 1 },
      { title: 'TODAY DUE', row: 1 },
      { title: 'THIS WEEK', row: 1 },
      { title: 'Data Health', row: 1 }
    ];
    const root = document.querySelector('#gsHeaderResults');
    try {
      const results = [];
      for (const target of targets) {
        try {
          const range = "'" + target.title.replace(/'/g, "''") + "'!" + target.row + ":" + target.row;
          const data = await window.MBBSGoogleSheets.readRange(cfg.spreadsheetId, range);
          results.push({ title: target.title, row: target.row, headers: (data.values && data.values[0]) || [], error: '' });
        } catch (e) {
          results.push({ title: target.title, row: target.row, headers: [], error: e.message || String(e) });
        }
      }
      if (root) root.innerHTML = '<div class="notice good">Header inspection complete. Master Entry uses row 2 for headers; the other selected tabs use row 1. No study records were imported or changed.</div>' +
        results.map(r => '<section class="panel"><h3>' + esc(r.title) + ' — header row ' + r.row + '</h3>' +
          (r.error ? '<p class="notice warn">' + esc(r.error) + '</p>' :
            (r.headers.length ? '<div class="tablewrap"><table><tbody>' + r.headers.map((h, i) => '<tr><th>' + String.fromCharCode(65 + i) + '</th><td>' + esc(h) + '</td></tr>').join('') + '</tbody></table></div>' : '<p>No non-empty header cells returned.</p>')) +
          '</section>').join('');
      status('Correct header rows inspected. Review these field names before any data mapping.', 'good');
    } catch (e) {
      status(e.message || String(e), 'error');
    } finally {
      busy = false; controls();
    }
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