# MBBS Recall Pro — Apps Script Sheet Manager (alternate backend)

This is a separate, server-side Google Apps Script web app. It avoids the GitHub Pages browser OAuth/header-inspection path entirely: the UI is served by Apps Script and calls the bound spreadsheet through `google.script.run`.

## Install (one-time)
1. Open the **native Google Sheet** that already has the verified Drive backups.
2. Extensions → Apps Script.
3. Replace the editor's default code with `Code.gs` from this folder. Add an HTML file named **Index** and paste `Index.html`.
4. Save. In Project Settings, confirm the project is bound to the correct spreadsheet. No client secret or OAuth client ID is used.
5. In Apps Script, select `doGet` and click **Run** once; review and approve Google permissions. If the editor prompts for a function selection, choose `getBootstrap` first.
6. Deploy → New deployment → Web app. Execute as **Me**; access should be **Only myself** for initial testing. Deploy and open the generated URL while signed in to the spreadsheet-owning account.

## Safety design
- This is a separate interface; it does not change the GitHub Pages app or local PWA storage.
- Master Entry writes are restricted to A:I; formula/calculated columns J:AC are never written.
- Revision Log writes are restricted to A:D; formula columns E:K are never written.
- Edits compare the original A:I snapshot before saving; concurrent changes cause a conflict error rather than overwrite.
- Archive copies the full nine input fields to a hidden `_MBBS Recall Archive` tab and clears only A:I; it does not delete a row or touch J:AC. A Restore control can put the record back (original row if still empty, otherwise the next available row); it refuses to duplicate an existing ID.
- A document lock serializes writes. Inputs are validated. IDs are generated from existing numeric IDs.
- The first version does not bulk import/export, sync the PWA local database, or delete revisions.

## Important preflight
The expected layout is based on the workbook screenshots and the connected workbook metadata: Master Entry headers in row 2, data from row 3, user inputs A:I; Revision Log headers row 1, data from row 2, user inputs A:D. Before using Add/Edit/Archive, verify these rows/columns in the native sheet. First test Add with a disposable record, verify formula columns remain intact, then test archive and restore manually from the hidden archive. Never test against important records first.

## Limits
This is a working alternate Sheet Manager, not yet a synchronization engine for the GitHub PWA. Keep using the current backup trigger system. Do not share the web-app URL publicly; initial access should be restricted to yourself.
