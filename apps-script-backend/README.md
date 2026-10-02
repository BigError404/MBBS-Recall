# MBBS Recall Pro — Apps Script backend alternative

This is an isolated, **read-only** prototype. It does not replace or change the existing GitHub Pages app, its local database, or the spreadsheet backup setup.

## Why this alternative
The browser OAuth/API inspection UI was difficult to use on mobile. Apps Script can access the spreadsheet server-side, and the HTML uses `google.script.run` instead of the browser Sheets API. This avoids the broken header-inspection layout and browser-side OAuth token flow.

## Deploy privately (required)
1. Open the native Google Sheet owned by your study account.
2. Open **Extensions → Apps Script**. This creates a script bound to that spreadsheet.
3. Replace the editor's starter code with the contents of `Code.gs` in this folder. Add an HTML file named exactly `Index` and paste `Index.html` into it.
4. Save the project with a name such as **MBBS Recall Pro — Private Backend**.
5. Select **Deploy → New deployment → Web app**.
6. Set **Execute as: Me** and **Who has access: Only myself**. Do not choose “Anyone” or “Anyone with a Google account”.
7. Deploy, review Google's requested permissions, and open the web-app URL while signed into the same Google account that owns the workbook.
8. Tap **Check workbook**, then **Load study records** and **Load recent revisions**.

## Safety limits in this prototype
- All operations are read-only.
- The workbook ID is fixed in `Code.gs`; the browser cannot supply another spreadsheet ID.
- Master Entry reads only columns A:I (input fields); it does not read or write calculated columns J:AC.
- It displays at most 200 study records and 100 recent revision rows per request.
- There are no add/edit/delete controls and no write methods.
- Never deploy this personal-data app with public access.

## Current scope / known limitations
This is a focused data-access prototype, not yet the complete replacement for the PWA. It validates direct, mobile-friendly reads through Apps Script. After those reads are confirmed, the next stage can implement carefully validated CRUD operations with explicit confirmation and backup checks, followed by a mobile UI for study/revision workflows. Do not merge it into the production PWA or assume sync exists until that later work is completed and tested.
