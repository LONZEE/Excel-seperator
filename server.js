const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const archiver = require('archiver');

const app = express();
const port = 3000;

const upload = multer({ storage: multer.memoryStorage() });

app.use(express.static('public'));

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function sanitizeFileName(value) {
  return String(value)
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .slice(0, 100) || 'missing_external_reference_id';
}

function normalizeExternalReferenceValue(value) {
  const raw = String(value || '').trim();

  if (!raw) {
    return {
      key: 'missing_external_reference_id',
      label: 'MISSING_EXTERNAL_REFERENCE_ID'
    };
  }

  const numericPattern = /^\d+(\.\d+)?$/;
  const label = numericPattern.test(raw) ? raw : raw.toLowerCase();

  return {
    key: raw.toLowerCase(),
    label
  };
}

function parseWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });

  if (!workbook.SheetNames.length) {
    throw new Error('Excel file has no sheets.');
  }

  const firstSheetName = workbook.SheetNames[0];
  const firstSheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

  if (!rows.length) {
    throw new Error('Excel sheet is empty.');
  }

  const headers = Object.keys(rows[0]);
  const targetHeader = headers.find((h) => {
    const normalized = normalizeHeader(h);
    return normalized === 'externalreferenceid' || normalized === 'externalreference';
  });

  if (!targetHeader) {
    throw new Error('Could not find a column named External Reference Id.');
  }

  return { rows, headers, targetHeader, firstSheetName };
}

app.post('/preview', upload.single('excelFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Please upload an Excel file.' });
  }

  try {
    const { rows, headers, targetHeader, firstSheetName } = parseWorkbook(req.file.buffer);
    const idGroups = new Map();

    for (const row of rows) {
      const normalized = normalizeExternalReferenceValue(row[targetHeader]);
      if (!idGroups.has(normalized.key)) {
        idGroups.set(normalized.key, { key: normalized.key, label: normalized.label, count: 0 });
      }
      idGroups.get(normalized.key).count += 1;
    }

    const externalReferenceIds = Array.from(idGroups.values()).sort((a, b) =>
      a.label.localeCompare(b.label)
    );

    return res.json({
      sheetName: firstSheetName,
      totalRows: rows.length,
      uniqueExternalReferenceIds: externalReferenceIds.length,
      targetHeader,
      headers,
      externalReferenceIds,
      rows
    });
  } catch (error) {
    return res.status(400).json({ error: error.message });
  }
});

app.post('/split', upload.single('excelFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Please upload an Excel file.' });
  }

  try {
    const { rows, targetHeader } = parseWorkbook(req.file.buffer);

    const grouped = new Map();

    for (const row of rows) {
      const normalized = normalizeExternalReferenceValue(row[targetHeader]);

      if (!grouped.has(normalized.key)) {
        grouped.set(normalized.key, {
          label: normalized.label,
          rows: []
        });
      }

      grouped.get(normalized.key).rows.push(row);
    }

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="split_by_external_reference_id.zip"'
    );

    const archive = archiver('zip', { zlib: { level: 9 } });
    archive.on('error', (err) => {
      res.status(500).end(`Zip creation failed: ${err.message}`);
    });

    archive.pipe(res);

    for (const group of grouped.values()) {
      const outWorkbook = XLSX.utils.book_new();
      const outSheet = XLSX.utils.json_to_sheet(group.rows);
      XLSX.utils.book_append_sheet(outWorkbook, outSheet, 'Rows');

      const outputBuffer = XLSX.write(outWorkbook, {
        type: 'buffer',
        bookType: 'xlsx'
      });

      const safeId = sanitizeFileName(group.label);
      archive.append(outputBuffer, { name: `${safeId}.xlsx` });
    }

    archive.finalize();
  } catch (error) {
    res.status(400).json({ error: `Failed to process file: ${error.message}` });
  }
});

app.listen(port, () => {
  console.log(`Excel splitter running at http://localhost:${port}`);
});
