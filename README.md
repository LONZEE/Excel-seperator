# Excel Splitter (External Reference Id)

Quick and dirty web app to upload an Excel file and split rows by `External Reference Id`.

This app now works as a static site, so it can be deployed on GitHub Pages.

## What it does
- Accepts `.xlsx` and `.xls` files.
- Reads the first sheet.
- Finds a column named `External Reference Id` (or similar like `external_reference_id`).
- Lets you preview rows in a table directly on the page.
- Groups rows by that value.
- Downloads a ZIP with one Excel file per External Reference Id.
- Runs entirely in the browser for GitHub Pages compatibility.

## Run
```bash
npm install
npm start
```

Then open: http://localhost:3000

For GitHub Pages, the site entry is [index.html](index.html).

## Use
- Click **Preview Table** to view all rows in-browser.
- Use the dropdowns and search box to filter by External Reference Id or any other header field.
- Click **Split & Download ZIP** to generate one file per External Reference Id.

## GitHub Pages
- GitHub Pages cannot run `server.js` or any Node/Express backend.
- The upload, preview, filtering, and ZIP split all run client-side now.
- The repo now includes a root [index.html](index.html), which GitHub Pages can serve directly.
- In GitHub, go to **Settings > Pages** and deploy from your `main` branch, `/ (root)` folder.

## Notes
- Rows with missing External Reference Id go into `MISSING_EXTERNAL_REFERENCE_ID.xlsx`.
- This is intentionally minimal and not hardened for production.
