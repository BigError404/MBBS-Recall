/* MBBS Recall Pro — Google Sheets read-only connection bootstrap.
 * No client secret. Access tokens remain in memory and are never persisted.
 * This module intentionally starts READ-ONLY; do not broaden scopes until the
 * workbook mapping and write-conflict safeguards have been tested.
 */
(() => {
  'use strict';
  const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets.readonly';
  let tokenClient = null;
  let accessToken = null;
  let configuredClientId = '';

  function loadIdentityServices() {
    if (window.google?.accounts?.oauth2) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const existing = document.querySelector('script[data-google-identity-services]');
      if (existing) {
        existing.addEventListener('load', resolve, { once: true });
        existing.addEventListener('error', () => reject(new Error('Google sign-in library failed to load.')), { once: true });
        return;
      }
      const script = document.createElement('script');
      script.src = 'https://accounts.google.com/gsi/client';
      script.async = true;
      script.defer = true;
      script.dataset.googleIdentityServices = 'true';
      script.onload = resolve;
      script.onerror = () => reject(new Error('Could not load Google sign-in library. Check your internet connection.'));
      document.head.appendChild(script);
    });
  }

  async function connect(clientId) {
    if (!clientId || !clientId.endsWith('.apps.googleusercontent.com')) {
      throw new Error('Enter a valid Google OAuth Client ID first.');
    }
    await loadIdentityServices();
    configuredClientId = clientId.trim();
    return new Promise((resolve, reject) => {
      tokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: configuredClientId,
        scope: SHEETS_SCOPE,
        include_granted_scopes: true,
        callback: response => {
          if (response?.error) {
            accessToken = null;
            reject(new Error('Google authorization failed: ' + response.error));
            return;
          }
          if (!response?.access_token) {
            reject(new Error('Google did not return an access token.'));
            return;
          }
          accessToken = response.access_token;
          resolve({ connected: true, scope: SHEETS_SCOPE });
        },
        error_callback: error => reject(new Error(error?.message || 'Google sign-in popup failed. Allow popups and try again.'))
      });
      tokenClient.requestAccessToken({ prompt: 'consent' });
    });
  }

  async function api(url, options = {}) {
    if (!accessToken) throw new Error('Connect your Google account first.');
    const response = await fetch(url, {
      ...options,
      headers: { ...(options.headers || {}), Authorization: 'Bearer ' + accessToken }
    });
    if (response.status === 401) {
      accessToken = null;
      throw new Error('Google authorization expired. Reconnect and retry.');
    }
    if (!response.ok) {
      let detail = '';
      try { detail = (await response.json()).error?.message || ''; } catch (_) {}
      throw new Error('Google Sheets API error (' + response.status + ')' + (detail ? ': ' + detail : '.'));
    }
    return response.json();
  }

  async function getSpreadsheet(spreadsheetId) {
    if (!/^[a-zA-Z0-9_-]{20,}$/.test(spreadsheetId || '')) throw new Error('Enter a valid spreadsheet ID.');
    const data = await api('https://sheets.googleapis.com/v4/spreadsheets/' +
      encodeURIComponent(spreadsheetId) + '?fields=spreadsheetId,spreadsheetUrl,properties(title),sheets(properties(sheetId,title,index,gridProperties(rowCount,columnCount)))');
    return {
      spreadsheetId: data.spreadsheetId,
      spreadsheetUrl: data.spreadsheetUrl,
      title: data.properties?.title || '',
      sheets: (data.sheets || []).map(s => ({
        sheetId: s.properties.sheetId,
        title: s.properties.title,
        index: s.properties.index,
        rowCount: s.properties.gridProperties?.rowCount,
        columnCount: s.properties.gridProperties?.columnCount
      }))
    };
  }

  async function readRanges(spreadsheetId, ranges) {
    if (!/^[a-zA-Z0-9_-]{20,}$/.test(spreadsheetId || '')) throw new Error('Enter a valid spreadsheet ID.');
    if (!Array.isArray(ranges) || ranges.length < 1 || ranges.length > 100 ||
        ranges.some(r => typeof r !== 'string' || !r.trim() || r.length > 300)) {
      throw new Error('Provide between 1 and 100 explicit spreadsheet ranges.');
    }
    const params = new URLSearchParams({ valueRenderOption: 'FORMATTED_VALUE', dateTimeRenderOption: 'FORMATTED_STRING' });
    ranges.forEach(range => params.append('ranges', range));
    return api('https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(spreadsheetId) +
      '/values:batchGet?' + params.toString());
  }

  async function readRange(spreadsheetId, range) {
    if (!range || typeof range !== 'string') throw new Error('Provide an explicit sheet range, e.g. Master Entry!A1:Z20.');
    const data = await readRanges(spreadsheetId, [range]);
    return { range: data.valueRanges?.[0]?.range || range, values: data.valueRanges?.[0]?.values || [] };
  }

  function disconnect() {
    if (accessToken && window.google?.accounts?.oauth2?.revoke) {
      try { window.google.accounts.oauth2.revoke(accessToken, () => {}); } catch (_) {}
    }
    accessToken = null;
    tokenClient = null;
  }

  window.MBBSGoogleSheets = Object.freeze({
    connect, disconnect, getSpreadsheet, readRange, readRanges,
    isConnected: () => Boolean(accessToken),
    scope: SHEETS_SCOPE
  });
})();
