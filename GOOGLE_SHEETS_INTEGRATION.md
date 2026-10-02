# Google Sheets integration — staged rollout

## Current stage: read-only connection panel

Branch: `sheets-sync-safe`  
Draft PR: https://github.com/BigError404/MBBS-Recall/pull/1

The test branch loads `google-sheets-readonly.js` and `google-sheets-ui.js`. In Settings, the connection panel can request consent, read spreadsheet metadata, and display tab names plus row/column counts.

It requests only this OAuth scope:

`https://www.googleapis.com/auth/spreadsheets.readonly`

## Automated verification

GitHub Actions run: https://github.com/BigError404/MBBS-Recall/actions/runs/36994588004

Status: **passed** for JavaScript syntax checks and the read-only safety contract. The workflow checks that:
- The OAuth scope stays read-only.
- No Sheets update/append/clear/batch-update endpoint is introduced.
- The helper does not use localStorage, sessionStorage, or IndexedDB for credentials.
- The first-stage UI only requests spreadsheet metadata, not study rows.
- Script load order and service-worker cache entries are present.

These are code-level checks only. They do **not** prove that Google OAuth consent or a real Sheets API request works in a browser.

## Safety boundaries

- No client secret is present in frontend code.
- Access tokens remain in memory and are not stored in IndexedDB/localStorage.
- The current UI does not write to Google Sheets.
- It does not import remote records into local study data.
- Production `main` and the live root app are unchanged; integration code is on the test branch.
- Google authorization requires an internet connection.

## Browser test still pending

The test branch does not yet have a separate deployed preview URL. Do not use the live production URL to test these changes: it still serves the production branch. Do not merge this draft PR merely to make the panel appear.

To complete a real browser test, deploy this branch to a separate preview host/path with HTTPS and configure that exact origin in the OAuth client's **Authorized JavaScript origins**. Keep the production root app and its deployment source unchanged. The existing origin `https://bigerror404.github.io` is valid only for pages served on that origin; origins cannot include a path.

Once the preview exists:
1. Keep the OAuth app in Testing and add the spreadsheet-owning Google account as a test user.
2. Open Settings while online and tap **Connect Google account**.
3. Approve the read-only request for the correct account.
4. Tap **Check workbook tabs** and confirm the expected workbook title and tab names.
5. Confirm local study records remain unchanged and the sheet itself was not modified.

Expected errors to investigate: origin mismatch, app not configured for this account, API not enabled in this project, popup blocked, or insufficient spreadsheet access.

## Next stages

1. Complete real browser OAuth and metadata verification.
2. Validate actual Master Entry and Revision Log headers/ranges before reading study rows.
3. Map sheet data to the app model and add an explicit preview/dry-run; never auto-import.
4. Consider writes only after a separate test plan. Use append/idempotency keys and version checks; never blindly overwrite entire tabs.
5. Test offline behavior, duplicate prevention, partial failures, conflict recovery, and backup restore.
6. Enable production only after explicit review.

## OAuth configuration

- Existing web application origin: `https://bigerror404.github.io`
- No client secret in browser code.
- Keep app in Testing during validation.
- The Google account owning the spreadsheet must have access to it.
- Read-only authorization cannot write to the spreadsheet; a future write stage would require a separate scope and renewed consent.
