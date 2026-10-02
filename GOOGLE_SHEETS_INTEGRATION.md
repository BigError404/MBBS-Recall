# Google Sheets integration — staged rollout

## Current stage: read-only connection panel

Branch: `sheets-sync-safe`

The test branch loads `google-sheets-readonly.js` and `google-sheets-ui.js`. In Settings, the connection panel can request consent, read spreadsheet metadata, and display tab names plus row/column counts.

It requests only this OAuth scope:

`https://www.googleapis.com/auth/spreadsheets.readonly`

Safety boundaries:
- No client secret is present in frontend code.
- Access tokens remain in memory and are not stored in IndexedDB/localStorage.
- The current UI does not write to Google Sheets.
- It does not import remote records into local study data.
- The production `main` branch is unchanged; these files are on the test branch only.
- The service worker caches both integration scripts for app-shell availability. Google authorization itself requires an internet connection.

## Test procedure

1. Confirm Google Sheets API is enabled in the same Cloud project as the OAuth client.
2. Keep the OAuth app in Testing and add the spreadsheet-owning Google account as a test user.
3. Deploy this branch only to a private/temporary preview if possible; do not merge into production until the test passes.
4. Open Settings while online and tap **Connect Google account**.
5. Approve the read-only request for the correct Google account.
6. Tap **Check workbook tabs** and confirm the expected workbook title and tab names appear.
7. Confirm the app's local study records remain unchanged and the sheet itself was not modified.

Expected errors to investigate: origin mismatch, app not configured for this account, API not enabled in this project, popup blocked, or insufficient spreadsheet access.

## Next stages

1. Validate the actual workbook's Master Entry and Revision Log headers/ranges before reading study rows.
2. Map sheet data to the app model and add an explicit preview/dry-run; never auto-import.
3. Consider writes only after a separate test plan. Use append/idempotency keys and version checks; never blindly overwrite entire tabs.
4. Test offline behavior, duplicate prevention, partial failures, conflict recovery, and backup restore.
5. Enable production only after explicit review.

## OAuth configuration

- Web application origin: `https://bigerror404.github.io`
- No client secret in browser code.
- Keep app in Testing during validation.
- The Google account owning the spreadsheet must have access to it.
- Read-only authorization cannot write to the spreadsheet; a future write stage would require a separate scope and renewed consent.
