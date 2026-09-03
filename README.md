# Excel Splitter (External Reference Id)

Quick and dirty web app to upload an Excel file and split rows by `External Reference Id`.

## What it does
- Accepts `.xlsx` and `.xls` files.
- Reads the first sheet.
- Finds a column named `External Reference Id` (or similar like `external_reference_id`).
- Lets you preview rows in a table directly on the page.
- Groups rows by that value.
- Downloads a ZIP with one Excel file per External Reference Id.

## Run
```bash
npm install
npm start
```

Then open: http://localhost:3000

## Use
- Click **Preview Table** to view up to the first 100 rows in-browser.
- Click **Split & Download ZIP** to generate one file per External Reference Id.

## Notes
- Rows with missing External Reference Id go into `MISSING_EXTERNAL_REFERENCE_ID.xlsx`.
- This is intentionally minimal and not hardened for production.
