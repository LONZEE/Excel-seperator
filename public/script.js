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

function normalizeExternalReferenceValue(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    return 'missing_external_reference_id';
  }
  return raw.toLowerCase();
}

function clearPreview() {
  previewData = null;
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

async function parseErrorMessage(response, fallback) {
  let message = fallback;
  try {
    const payload = await response.json();
    message = payload.error || message;
  } catch {
    // Ignore JSON parse errors for non-JSON error responses.
  }
  return message;
}

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
  formData.append('excelFile', file);

  try {
    if (action === 'preview') {
      const response = await fetch('/preview', {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        const message = await parseErrorMessage(response, 'Preview failed.');
        throw new Error(message);
      }

      const payload = await response.json();
      renderPreview(payload);
      statusEl.textContent = `Preview ready. Loaded ${payload.rows.length} rows.`;
      return;
    }

    const response = await fetch('/split', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      const message = await parseErrorMessage(response, 'Upload failed.');
      throw new Error(message);
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'split_by_external_reference_id.zip';
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    statusEl.textContent = 'Done. ZIP downloaded.';
  } catch (error) {
    clearPreview();
    statusEl.textContent = error.message;
  }
});
