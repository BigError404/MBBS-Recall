# MBBS Recall Pro v7 — workbook-aligned offline PWA

A mobile-first rebuild of the earlier primitive PWA, aligned to `MBBS_SRS_GOOGLE_SHEETS_ANDROID_FIXED_v11_FINAL.xlsx`.

## Important data-safety note

- The v4 source had no persistent database despite its README claiming IndexedDB. v5 introduced IndexedDB with localStorage fallback; v6 retains those storage keys to preserve existing local data when upgrading in the same browser profile.
- Export a full JSON backup regularly, especially before uninstalling the PWA, clearing browser/site data, changing devices, importing data, or resetting the app.
- Local browser storage is device/browser-specific; it is not cloud sync and does not automatically sync with Google Sheets.
- The included seed is extracted from the workbook file available during this rebuild (one IMP, one revision, all 12 subjects, configured units and interval settings). It is only the starting snapshot; import your current `.xlsx` workbook if it has changed since then.

## Features

- **New in v7:** import the original `.xlsx` workbook directly offline, including Master Entry inputs, Revision Log inputs, all 12 subject unit lists, interval settings, and exam date when set; Master Entry now exposes a broader set of workbook helper/calculated columns. **Retained from v6:** sort IMPs by ID, subject/unit, due date, priority, or failure count; select individual rows, visible-page rows, or all filtered matches; bulk-delete selected IMPs with linked revision cleanup and Undo.
- Workbook-aligned dashboard, TODAY DUE, THIS WEEK, Master Entry, Revision Log, Weak Topics, Unit Progress, 12 subject views, Unit Management, interval settings, and exam-date cap.
- Create, edit, search, filter, paginate, export and delete IMPs. Deleting an IMP asks for confirmation and removes associated revision rows so they cannot become orphaned; an in-session Undo snapshot is offered.
- Create, edit and delete revision rows; grades from Study append a revision and update the schedule.
- PASS/FAIL/PARTIAL normalization, date validation, unit validation, duplicate-ID detection, stage/streak calculations, exam cap, and data-quality flags.
- Full JSON backup/restore and CSV import/export for Master Entry and Revision Log.
- Responsive phone layout, keyboard controls for study, offline asset cache, no external fonts/CDN dependencies. The XLSX reader requires modern-browser `DecompressionStream` support for raw DEFLATE; test on your Android Chrome after deployment.

## Upgrade from v5 / v6

- Upload all files in this ZIP to the root of the existing GitHub Pages repository, replacing same-named files.
- The IndexedDB and localStorage keys are intentionally retained so the app can read existing v5 local data in the same browser profile. Keep a full JSON backup before upgrading.
- The service-worker cache name is updated to v7 so cached app assets refresh.

## Run / install

1. Upload the contents of this folder to the **root of your existing GitHub Pages repository**, replacing same-named old app files. Keep all files together; do not upload only `index.html`.
2. Commit the changes and wait for GitHub Pages to deploy.
3. Open the Pages URL in Chrome on Android, allow it to load once online, then Chrome menu → **Install app** / **Add to Home screen**.
4. Open Settings → **Export full JSON backup** before importing or making large changes.

For local testing, serve this folder with `python -m http.server 8080` or `npx serve .`; PWA installation/service workers require HTTP(S), not a direct `file://` URL.

## Importing the workbook

- Open Settings or Workbook → Master Entry and choose Import, then select the original `.xlsx` file. Import parsing is local and offline; it does not upload the workbook anywhere.
- The importer reads Master Entry input columns A:I, Revision Log input columns A:D, Units Config, Settings - Intervals, and the Dashboard exam date. It skips calculated/formula columns because the app recomputes scheduling from the imported inputs.
- Import merges cards by ID (matching IDs are replaced; unrelated app cards remain), appends nonduplicate workbook reviews, and updates unit lists and interval settings. A nonblank exam date replaces the app setting. A confirmation screen describes these effects before changes are applied.
- CSV import remains available for Master Entry and Revision Log exports. Full JSON backup/restore remains the most complete way to move all app data. Neither `.xlsx` nor CSV imports copy cell formatting, dropdown validation, charts, or the spreadsheet renderer itself.

## Limits and parity scope

The app reproduces the main scheduling rules and key workbook data views, not Google Sheets itself. v7 adds direct offline workbook import and a wider Master Entry table; v6 introduced better bulk data management but does not provide live Google Sheets sync, pixel-identical spreadsheet behavior, every formula helper column, or all cell-level dropdown/formatting behavior. Workbook-style capacity warnings show 1,000 IMPs / 9,999 revisions. This is a client-side app and does not guarantee data recovery if device storage is removed.

## Developer tests

The `tests/` folder contains offline automated checks; it is not needed for normal studying. With Node.js installed, run `node tests/engine.test.js`, `node tests/app-regression.test.js`, and `node tests/xlsx-import.test.js` from this folder. These checks do not replace testing on your Android device.
