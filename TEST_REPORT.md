# MBBS Recall Pro v7 — validation report

Date: 2026-10-02  
Reference workbook: `MBBS_SRS_GOOGLE_SHEETS_ANDROID_FIXED_v11_FINAL.xlsx`  
Scope: workbook structure and input-field extraction, scheduling engine tests, app source regressions, synthetic XLSX-import fixture, syntax checks, and deterministic workload simulations. This report does not claim complete Google Sheets parity or certification on a physical Android device.

## Workbook analysis and import coverage

The reference workbook contains 26 sheets. The app importer reads the user-input fields from `Master Entry` columns A:I and `Revision Log` columns A:D, the 12 subject columns in `Units Config`, the 3 interval sets in `Settings - Intervals`, and a nonblank exam date from `Dashboard!B3`.

The supplied workbook was parsed during the build and returned:

- 1 IMP row and 1 revision row
- 12 subjects and 97 configured units
- Interval values matching the workbook's Regular, Must-Do, and Late sets
- No exam date set in `Dashboard!B3`

This means the supplied file is a starting snapshot with one actual IMP/revision row; it is not a workbook full of pre-populated study questions. Importing does not copy formulas, formatting, charts, dropdown validation, or the spreadsheet renderer. The app recalculates schedule/helper values from the imported input data.

## v7 changes

- Added an offline `.xlsx` importer without external CDN/dependency requirements. It reads ordinary, unencrypted OOXML `.xlsx` files using ZIP decompression and XML parsing APIs in modern browsers.
- Import preview/confirmation explains merge behavior: matching IMP IDs are replaced, unrelated local cards remain, duplicate revision rows are skipped, unit lists and intervals are replaced from the workbook, and a nonblank exam date is applied.
- Added input validation for duplicate IDs inside an imported workbook before changing app state.
- Expanded Master Entry into a horizontally scrollable workbook-style table with input fields and computed fields: entry date, subject, unit, question, answer, source, last revision, last result, attempts, passes, fails, streak, stage, next due, status, notes, priority, fail rate, days left, data-quality flags, and edit/delete actions.
- Regenerated `workbook-seed.js` and `workbook-seed.json` directly from the supplied v11 workbook's input fields, Units Config, and interval settings.
- Bumped the service-worker cache to v7 and added `xlsx-import.js` to the offline core cache. The existing v5 IndexedDB database and localStorage fallback identifiers are retained for upgrade continuity.

## Automated checks

- Scheduling engine: **27/27 passed**, including date validation, PASS/PARTIAL/FAIL intervals, streak behavior, exam cap, duplicate IDs, data-quality exclusions, CSV parsing, formula-injection protection, and 180-day simulations.
- App regressions: **11/11 passed**, including upgrade storage keys, v7 cache assets, bulk selection/deletion/undo, sorting, Master Entry columns, XLSX wiring, full helper-column CSV export, and PWA manifest settings.
- XLSX parser synthetic fixture: **12 assertions passed**, covering record fields, revisions, 12-subject mapping, intervals, unit extraction, blank exam date, and Excel serial date conversion.
- Real source workbook parse: **1 IMP, 1 revision, 12 subjects, 97 units** extracted successfully.
- JavaScript syntax checks and manifest JSON parsing were run on the final source.

## Import merge behavior and safety

The importer does not silently clear the existing database. It merges workbook IMPs by ID and warns that same-ID records will be replaced. Existing revision history is retained; exact matching date/ID/result/notes rows are deduplicated. The workbook's configured units and intervals become the active settings, so older cards using renamed/removed units may be flagged and need manual migration. Back up full JSON before import.

## Remaining limitations / not certified

- Live Google Sheets synchronization is not implemented.
- This is not a pixel-identical spreadsheet and does not render every workbook helper cell, formula, dropdown, chart, or conditional format.
- Direct XLSX parsing depends on browser support for `DecompressionStream('deflate-raw')`; the importer gives an error if the browser cannot use it. Current Android Chrome should be tested directly.
- A headless browser smoke test could not be completed in the build environment. Physical Android testing is still required for import UI, persistence after reload, installation, offline relaunch, and service-worker update behavior.
- Browser storage remains local to the browser/device. Export JSON backups regularly, especially before imports and app upgrades.
