const form = document.getElementById('uploadForm');
const statusEl = document.getElementById('status');
const previewSection = document.getElementById('previewSection');
const previewMeta = document.getElementById('previewMeta');
const previewTable = document.getElementById('previewTable');
const idFilter = document.getElementById('idFilter');
const headerFilter = document.getElementById('headerFilter');
const searchInput = document.getElementById('searchInput');
const filterMeta = document.getElementById('filterMeta');
const columnOptions = document.getElementById('columnOptions');
const showAllColumnsButton = document.getElementById('showAllColumns');
const hideAllColumnsButton = document.getElementById('hideAllColumns');
const presetIdColumnsButton = document.getElementById('presetIdColumns');
const presetCustomerColumnsButton = document.getElementById('presetCustomerColumns');
const googleSheetUrlInput = document.getElementById('googleSheetUrl');
const loadGoogleSheetButton = document.getElementById('loadGoogleSheet');

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

function isSummableMoneyColumn(header) {
  return /(amount|balance|premium|charge|total|payment|paid|due|cost|price|fee|credit|debit)/i.test(
    String(header || '')
  );
}

function parseNumericValue(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  const cleaned = String(value ?? '')
    .trim()
    .replace(/[$,\s]/g, '')
    .replace(/[()]/g, '');

  if (!cleaned) {
    return null;
  }

  const isNegative = /^\(.*\)$/.test(String(value ?? '').trim()) || /^-/.test(cleaned);
  const parsed = Number.parseFloat(cleaned.replace(/^-/, ''));

  if (!Number.isFinite(parsed)) {
    return null;
  }

  return isNegative ? -parsed : parsed;
}

function formatAmount(value) {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(value);
}

function getFileKey(file) {
  return [file.name, file.size, file.lastModified].join(':');
}

async function readFileAsArrayBuffer(file) {
  return file.arrayBuffer();
}

function createPreviewData(workbook, sourceName) {
  if (!workbook.SheetNames.length) {
    throw new Error('The file has no sheets.');
  }

  const firstSheetName = workbook.SheetNames[0];
  const firstSheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

  if (!rows.length) {
    throw new Error('The selected sheet is empty.');
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

  return {
    sheetName: sourceName || firstSheetName,
    totalRows: rows.length,
    uniqueExternalReferenceIds: idGroups.size,
    targetHeader,
    headers,
    externalReferenceIds: Array.from(idGroups.values()).sort((a, b) => a.label.localeCompare(b.label)),
    rows
  };
}

async function loadWorkbookData(file) {
  const fileKey = getFileKey(file);
  if (previewData && loadedFileKey === fileKey) {
    return previewData;
  }

  const buffer = await readFileAsArrayBuffer(file);
  const workbook = XLSX.read(buffer, { type: 'array' });
  const payload = createPreviewData(workbook, file.name);

  previewData = payload;
  loadedFileKey = fileKey;
  return payload;
}

function getGoogleSheetExportUrl(url) {
  const parsed = new URL(url);
  const sharedSheetMatch = parsed.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
  const publishedSheetMatch = parsed.pathname.match(/\/spreadsheets\/d\/e\/([^/]+)/);
  const gidMatch = parsed.hash.match(/(?:^#|[?&])gid=(\d+)/);
  const gid = parsed.searchParams.get('gid') || (gidMatch ? gidMatch[1] : '0');

  if (publishedSheetMatch) {
    const publishedSheetId = publishedSheetMatch[1];
    return `https://docs.google.com/spreadsheets/d/e/${publishedSheetId}/pub?output=csv&gid=${encodeURIComponent(gid)}`;
  }

  if (!sharedSheetMatch) {
    throw new Error('Paste a Google Sheets sharing link.');
  }

  const sheetId = sharedSheetMatch[1];
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${encodeURIComponent(gid)}`;
}

async function loadGoogleSheet(url) {
  const exportUrl = getGoogleSheetExportUrl(url);
  let response;

  try {
    response = await fetch(exportUrl);
  } catch {
    throw new Error(
      'Google blocked this sheet from browser import. In Google Sheets, use File > Share > Publish to web, choose CSV, then paste the published link.'
    );
  }

  if (!response.ok) {
    throw new Error(
      'Google Sheet could not be loaded. Use File > Share > Publish to web, choose CSV, then paste the published link.'
    );
  }

  const csv = await response.text();

  if (/<!doctype html|<html/i.test(csv)) {
    throw new Error(
      'Google returned a sign-in or access page. Use File > Share > Publish to web, choose CSV, then paste the published link.'
    );
  }

  const workbook = XLSX.read(csv, { type: 'string' });
  const payload = createPreviewData(workbook, 'Google Sheet');

  previewData = payload;
  loadedFileKey = `google-sheet:${exportUrl}`;
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
  window.setTimeout(() => window.URL.revokeObjectURL(url), 1000);
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
  columnOptions.innerHTML = '';
  previewTable.innerHTML = '';
  previewSection.classList.add('hidden');
}

function renderTable(headers, rows) {
  const visibleHeaders = headers.filter((header) => {
    const checkbox = columnOptions.querySelector(`[data-header="${CSS.escape(header)}"]`);
    return !checkbox || checkbox.checked;
  });

  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');

  for (const header of visibleHeaders) {
    const th = document.createElement('th');
    th.textContent = header;
    headRow.appendChild(th);
  }

  thead.appendChild(headRow);

  const tbody = document.createElement('tbody');
  for (const row of rows) {
    const tr = document.createElement('tr');
    for (const header of visibleHeaders) {
      const td = document.createElement('td');
      td.textContent = String(row[header] ?? '');
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  const tfoot = document.createElement('tfoot');
  const footerRow = document.createElement('tr');
  let hasSummableColumn = false;

  for (const [index, header] of visibleHeaders.entries()) {
    const td = document.createElement('td');

    if (index === 0) {
      td.textContent = 'Totals';
    }

    if (isSummableMoneyColumn(header)) {
      const total = rows.reduce((sum, row) => {
        const numericValue = parseNumericValue(row[header]);
        return numericValue === null ? sum : sum + numericValue;
      }, 0);

      td.textContent = formatAmount(total);
      hasSummableColumn = true;
    }

    footerRow.appendChild(td);
  }

  previewTable.innerHTML = '';
  previewTable.appendChild(thead);
  previewTable.appendChild(tbody);
  if (hasSummableColumn) {
    tfoot.appendChild(footerRow);
    previewTable.appendChild(tfoot);
  }
}

function renderColumnOptions(headers) {
  columnOptions.innerHTML = '';

  for (const header of headers) {
    const label = document.createElement('label');
    label.className = 'columnOption';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = true;
    checkbox.dataset.header = header;
    checkbox.addEventListener('change', () => {
      const checkedCount = columnOptions.querySelectorAll('input:checked').length;
      if (!checkedCount) {
        checkbox.checked = true;
        return;
      }
      renderFilteredRows();
    });

    const text = document.createElement('span');
    text.textContent = header;

    label.appendChild(checkbox);
    label.appendChild(text);
    columnOptions.appendChild(label);
  }
}

function setVisibleColumns(predicate) {
  const checkboxes = Array.from(columnOptions.querySelectorAll('input[type="checkbox"]'));
  let checkedCount = 0;

  for (const checkbox of checkboxes) {
    checkbox.checked = predicate(checkbox.dataset.header);
    if (checkbox.checked) {
      checkedCount += 1;
    }
  }

  if (!checkedCount && checkboxes.length > 0) {
    checkboxes[0].checked = true;
  }

  renderFilteredRows();
}

function isIdColumn(header) {
  return /(id|reference|ref|number|code|key)/i.test(header);
}

function isCustomerColumn(header) {
  return /(customer|member|client|account|person|name|email|phone|address|city|state|zip)/i.test(header);
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

  renderColumnOptions(payload.headers);
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

showAllColumnsButton.addEventListener('click', () => {
  setVisibleColumns(() => true);
});

hideAllColumnsButton.addEventListener('click', () => {
  setVisibleColumns(() => false);
});

presetIdColumnsButton.addEventListener('click', () => {
  setVisibleColumns((header) => isIdColumn(header));
});

presetCustomerColumnsButton.addEventListener('click', () => {
  setVisibleColumns((header) => isCustomerColumn(header));
});

loadGoogleSheetButton.addEventListener('click', async () => {
  const url = googleSheetUrlInput.value.trim();

  if (!url) {
    statusEl.textContent = 'Paste a Google Sheets sharing link first.';
    return;
  }

  statusEl.textContent = 'Loading Google Sheet...';

  try {
    const payload = await loadGoogleSheet(url);
    renderPreview(payload);
    statusEl.textContent = `Google Sheet loaded. ${payload.rows.length} rows ready.`;
  } catch (error) {
    clearPreview();
    statusEl.textContent = error.message;
  }
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
