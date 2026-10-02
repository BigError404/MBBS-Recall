# MBBS Recall — Offline-first Android PWA

A free, installable progressive web app for MBBS spaced repetition, seeded from `MBBS_SRS_GOOGLE_SHEETS_ANDROID_FIXED_v11_FINAL.xlsx`.

## What's included
- Extracted the workbook's entered card fields, review-log entries, 12 subjects, configured units, and priority-specific SRS intervals into `seed-data.json`.
- Local-first storage with IndexedDB (localStorage fallback).
- Dashboard, due queue, PASS / PARTIAL / FAIL grading, review history, subject progress, weak-topic analytics, card search/edit, and exam-date cap.
- JSON full backups and restore; CSV card import/export for round-tripping with Google Sheets.
- Offline service worker and installable PWA manifest.

## Install on Android
A PWA's service worker and install prompt require a secure origin (HTTPS) or localhost. Opening `index.html` directly from Downloads will not provide reliable offline caching or installation.

Free static hosting option:
1. Create a GitHub repository and upload the contents of this folder (all files at the repository root).
2. In repository **Settings → Pages**, choose deploy from the main branch and root folder.
3. Open the resulting `https://...github.io/.../` link in Chrome on Android while online once.
4. Choose Chrome menu → **Install app** or **Add to Home screen**.
5. Open the app once while online so the app shell and seed data are cached. Then test airplane mode.

No server or paid database is required for local use. Optional cloud sync is not configured in this initial version. To move data between devices, export a JSON backup and restore it on the other device.

## Importing future workbook updates
The baseline data from the latest workbook is already included. For later updates, export the `Master Entry` sheet from Google Sheets as CSV, keeping these columns if possible: `ID`, `Entry Date`, `Subject`, `Unit / System`, `Question / IMP`, `Answer Key Points`, `Source`, `Notes`, `Priority`. In the app, use **Settings → Import cards CSV**. It merges by ID and updates matching cards. This app does not execute Excel/Google Sheets formulas or Apps Script; it uses its own local SRS engine.

To transfer reviews from a newer workbook, export `Revision Log` as CSV and use a future dedicated history-import flow or restore a full JSON backup exported by this app. The initial workbook's existing revision entries are already seeded.

## Data and SRS notes
- Existing workbook entries are imported without running formulas; computed stage/due state is reconstructed from the review history.
- PASS advances the pass streak/stage up to stage 5; FAIL schedules again based on the workbook's interval settings; PARTIAL preserves stage and uses the configured partial interval.
- Review dates use the device's local calendar date.
- Local browser storage can be erased by clearing site data. Export backups regularly.

## Files
- `index.html`, `styles.css`, `app.js`: app UI and logic
- `manifest.json`, `icon.svg`, `sw.js`: PWA install/offline shell
- `seed-data.json`: extracted workbook baseline
- `templates/cards-import-template.csv`: import example
