# MBBS Recall Pro — Sheets Preview

This folder is a separate preview build for testing Google OAuth and read-only Google Sheets metadata access.

- Preview URL after this preview-only folder is published by GitHub Pages: `https://bigerror404.github.io/MBBS-Recall/preview/`
- The app shows a visible PREVIEW ONLY banner.
- Its IndexedDB and localStorage key is `mbbs-recall-sheets-preview-v1`, separate from production app storage.
- Its service worker uses a separate cache namespace and does not delete production caches.
- Google Sheets access requests only `spreadsheets.readonly`.
- It reads spreadsheet metadata/tab names only; no study rows are imported and no spreadsheet cells are written.
- The production app files at the repository root are not changed by this preview-only folder.

This preview is not yet published. It becomes available only after this separate preview PR is reviewed/merged and GitHub Pages finishes deploying. Do not merge the main integration PR as part of this step.
