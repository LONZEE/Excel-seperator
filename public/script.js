const form = document.getElementById('uploadForm');
const statusEl = document.getElementById('status');
const previewSection = document.getElementById('previewSection');
const previewMeta = document.getElementById('previewMeta');
const previewTable = document.getElementById('previewTable');
const idFilter = document.getElementById('idFilter');
const headerFilter = document.getElementById('headerFilter');
const searchInput = document.getElementById('searchInput');
const filterMeta = document.getElementById('filterMeta');

let previewData = null;
let loadedFileKey = null;

function normalizeExternalReferenceValue(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    return 'missing_external_reference_id';
  }
  return raw.toLowerCase();
}

function formatExternalReferenceValue(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    return 'MISSING_EXTERNAL_REFERENCE_ID';
  }

  if (/^\d+(\.\d+)?$/.test(raw)) {
    return raw;
  }

  return raw.toLowerCase();
}

function sanitizeFileName(value) {
  return String(value)
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .slice(0, 100) || 'missing_external_reference_id';
}

function getFileKey(file) {
  return [file.name, file.size, file.lastModified].join(':');
}

async function readFileAsArrayBuffer(file) {
  return file.arrayBuffer();
}

async function loadWorkbookData(file) {
  const fileKey = getFileKey(file);
  if (previewData && loadedFileKey === fileKey) {
    return previewData;
  }

  const buffer = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(buffer, { type: 'array' });

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
  const targetHeader = headers.find((header) => {
    const normalized = String(header || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');

    return normalized === 'externalreferenceid' || normalized === 'externalreference';
  });

  if (!targetHeader) {
    throw new Error('Could not find a column named External Reference Id.');
  }

  const idGroups = new Map();

  for (const row of rows) {
    const key = normalizeExternalReferenceValue(row[targetHeader]);
    const label = formatExternalReferenceValue(row[targetHeader]);

    if (!idGroups.has(key)) {
      idGroups.set(key, { key, label, count: 0 });
    }

    idGroups.get(key).count += 1;
  }

  const payload = {
    sheetName: firstSheetName,
    totalRows: rows.length,
    uniqueExternalReferenceIds: idGroups.size,
    targetHeader,
    headers,
    externalReferenceIds: Array.from(idGroups.values()).sort((a, b) => a.label.localeCompare(b.label)),
    rows
  };

  previewData = payload;
  loadedFileKey = fileKey;
  return payload;
}

function downloadBlob(blob, fileName) {
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.URL.revokeObjectURL(url);
}

async function splitWorkbook(payload) {
  const grouped = new Map();

  for (const row of payload.rows) {
    const key = normalizeExternalReferenceValue(row[payload.targetHeader]);
    const label = formatExternalReferenceValue(row[payload.targetHeader]);

    if (!grouped.has(key)) {
      grouped.set(key, { label, rows: [] });
    }

    grouped.get(key).rows.push(row);
  }

  const zip = new JSZip();

  for (const group of grouped.values()) {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(group.rows);
    XLSX.utils.book_append_sheet(workbook, sheet, 'Rows');

    const output = XLSX.write(workbook, {
      type: 'array',
      bookType: 'xlsx'
    });

    zip.file(`${sanitizeFileName(group.label)}.xlsx`, output);
  }

  return zip.generateAsync({ type: 'blob' });
}

function clearPreview() {
  previewData = null;
  loadedFileKey = null;
  previewMeta.textContent = '';
  filterMeta.textContent = '';
  idFilter.innerHTML = '';
  headerFilter.innerHTML = '';
  searchInput.value = '';
  previewTable.innerHTML = '';
  previewSection.classList.add('hidden');
}

function renderTable(headers, rows) {
  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');

  for (const header of headers) {
    const th = document.createElement('th');
    th.textContent = header;
    headRow.appendChild(th);
  }

  thead.appendChild(headRow);

  const tbody = document.createElement('tbody');
  for (const row of rows) {
    const tr = document.createElement('tr');
    for (const header of headers) {
      const td = document.createElement('td');
      td.textContent = String(row[header] ?? '');
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  previewTable.innerHTML = '';
  previewTable.appendChild(thead);
  previewTable.appendChild(tbody);
}

function getFilteredRows(data, selectedId, selectedHeader, query) {
  const normalizedQuery = String(query || '').trim().toLowerCase();

  return data.rows.filter((row) => {
    const idMatches =
      selectedId === '__ALL__' ||
      normalizeExternalReferenceValue(row[data.targetHeader]) === selectedId;

    if (!idMatches) {
      return false;
    }

    if (!normalizedQuery) {
      return true;
    }

    if (selectedHeader === '__ANY__') {
      return data.headers.some((header) =>
        String(row[header] ?? '')
          .toLowerCase()
          .includes(normalizedQuery)
      );
    }

    return String(row[selectedHeader] ?? '')
      .toLowerCase()
      .includes(normalizedQuery);
  });
}

function renderFilteredRows() {
  if (!previewData) {
    return;
  }

  const selectedId = idFilter.value || '__ALL__';
  const selectedHeader = headerFilter.value || '__ANY__';
  const query = searchInput.value || '';
  const rows = getFilteredRows(previewData, selectedId, selectedHeader, query);
  renderTable(previewData.headers, rows);

  const idText =
    selectedId === '__ALL__' ? 'all External Reference Ids' : `External Reference Id: ${selectedId}`;
  const headerText =
    selectedHeader === '__ANY__' ? 'all headers' : `header: ${selectedHeader}`;
  const queryText = query.trim() ? `query: "${query.trim()}"` : 'no text filter';

  filterMeta.textContent = `Showing ${rows.length} rows (${idText}, ${headerText}, ${queryText}).`;
}

function renderPreview(payload) {
  const { totalRows, uniqueExternalReferenceIds, targetHeader, sheetName, externalReferenceIds } = payload;

  previewData = payload;
  previewMeta.textContent = `Sheet: ${sheetName} | Rows: ${totalRows} | Unique External Reference Ids: ${uniqueExternalReferenceIds} | Matched Column: ${targetHeader}`;

  idFilter.innerHTML = '';

  const allOption = document.createElement('option');
  allOption.value = '__ALL__';
  allOption.textContent = 'All External Reference Ids';
  idFilter.appendChild(allOption);

  for (const id of externalReferenceIds) {
    const option = document.createElement('option');
    option.value = id.key;
    option.textContent = `${id.label} (${id.count})`;
    idFilter.appendChild(option);
  }

  headerFilter.innerHTML = '';

  const anyOption = document.createElement('option');
  anyOption.value = '__ANY__';
  anyOption.textContent = 'Any Header Field';
  headerFilter.appendChild(anyOption);

  for (const header of payload.headers) {
    const option = document.createElement('option');
    option.value = header;
    option.textContent = header;
    headerFilter.appendChild(option);
  }

  idFilter.value = '__ALL__';
  headerFilter.value = '__ANY__';
  searchInput.value = '';
  renderFilteredRows();
  previewSection.classList.remove('hidden');
}

idFilter.addEventListener('change', () => {
  renderFilteredRows();
});

headerFilter.addEventListener('change', () => {
  renderFilteredRows();
});

searchInput.addEventListener('input', () => {
  renderFilteredRows();
});

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  const fileInput = document.getElementById('excelFile');
  const file = fileInput.files[0];

  if (!file) {
    statusEl.textContent = 'Pick a file first.';
    return;
  }

  const action = event.submitter?.value || 'split';
  statusEl.textContent = action === 'preview' ? 'Loading preview...' : 'Processing...';

  const formData = new FormData();

  try {
    const payload = await loadWorkbookData(file);

    if (action === 'preview') {
      renderPreview(payload);
      statusEl.textContent = `Preview ready. Loaded ${payload.rows.length} rows.`;
      return;
    }

    const blob = await splitWorkbook(payload);
    downloadBlob(blob, 'split_by_external_reference_id.zip');

    statusEl.textContent = 'Done. ZIP downloaded.';
  } catch (error) {
    clearPreview();
    statusEl.textContent = error.message;
  }
});
