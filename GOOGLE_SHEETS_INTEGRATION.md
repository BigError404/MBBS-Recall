# Google Sheets integration — staged rollout

## Current stage: read-only OAuth helper

Branch: `sheets-sync-safe`

The file `google-sheets-readonly.js` implements a small browser-side helper using Google Identity Services token flow. It does not contain a client secret and does not persist access tokens. It requests only:

`https://www.googleapis.com/auth/spreadsheets.readonly`

It can:
- request user consent with an OAuth Client ID,
- read spreadsheet metadata (title and tab names),
- read an explicitly specified range,
- disconnect/revoke the in-memory token.

## Important: not yet wired into the app UI

This is an isolated first step, not a finished synchronization feature. It is intentionally not loaded by `index.html` and cannot change the existing app behavior. No app data is sent anywhere by this helper unless a developer explicitly calls its API after authorization.

## Next stages

1. Add a Settings panel for Client ID and spreadsheet ID, without saving access tokens.
2. Load this helper and perform metadata + read-only range tests.
3. Map the real workbook's Master Entry and Revision Log layouts to the app data model; verify formulas and column headers before mapping.
4. Add writes only after a separate test plan is accepted. Use append/idempotency keys and a revision/version check to detect conflicts; never blindly overwrite entire tabs.
5. Test offline queue/retry, duplicate prevention, partial failures, and backup restore.
6. Only then consider enabling the integration in production.

## OAuth configuration notes

- Use a Web application OAuth client with JavaScript origin `https://bigerror404.github.io`.
- Do not put a client secret in frontend code. The GIS browser token flow uses the client ID and user consent.
- Keep the OAuth app in Testing and add the owner's Google account as a test user.
- Google Sheets API must be enabled in the same Cloud project as the OAuth client.
- A read-only test cannot write to the spreadsheet. A later write stage will require the appropriate scope and renewed consent.
