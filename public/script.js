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

// Payroll elements
const payCycleInput = document.getElementById('payCycle');
const commissionTableBody = document.getElementById('commissionTableBody');
const syncGoogleSheetButton = document.getElementById('syncGoogleSheetButton');
const googleWebhookUrlInput = document.getElementById('googleWebhookUrl');
const syncStatusEl = document.getElementById('syncStatus');
const exportSummaryExcelBtn = document.getElementById('exportSummaryExcelBtn');

// Hard set list of trainers
const HARD_SET_TRAINERS = [
  'Nick',
  'Erick',
  'Tete',
  'Magali',
  'Sandra',
  'Jose',
  'Mario',
  'Anthony'
];

// OPTIONAL: Hardcode your Google Sheet Web App URL here so you never have to paste it anywhere
const DEFAULT_GOOGLE_SHEET_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbxTOKCLNP4WhnM09AXqIha0V3SINIvfGfOYyW2-vK9mDSQJLeTASlNZ9_7SV0H2-8o7/exec';

// Default commission percentages
const DEFAULT_COMMISSIONS = {
  'Nick': 50,
  'Erick': 50,
  'Tete': 50,
  'Magali': 50,
  'Sandra': 50,
  'Jose': 50,
  'Mario': 50,
  'Anthony': 50
};

// Load saved rates from localStorage if available
function getTrainerCommissionRates() {
  try {
    const saved = localStorage.getItem('trainer_commissions');
    if (saved) {
      return Object.assign({}, DEFAULT_COMMISSIONS, JSON.parse(saved));
    }
  } catch {
    // ignore
  }
  return Object.assign({}, DEFAULT_COMMISSIONS);
}

function saveTrainerCommissionRate(trainer, rate) {
  try {
    const rates = getTrainerCommissionRates();
    rates[trainer] = Number(rate) || 0;
    localStorage.setItem('trainer_commissions', JSON.stringify(rates));
  } catch {
    // ignore
  }
}

// Load saved Webhook URL if available
function getSavedWebhookUrl() {
  try {
    return localStorage.getItem('google_sheet_webhook_url') || DEFAULT_GOOGLE_SHEET_WEBHOOK_URL;
  } catch {
    return DEFAULT_GOOGLE_SHEET_WEBHOOK_URL;
  }
}

function saveWebhookUrl(url) {
  try {
    localStorage.setItem('google_sheet_webhook_url', url.trim());
  } catch {
    // ignore
  }
}

if (googleWebhookUrlInput) {
  googleWebhookUrlInput.value = getSavedWebhookUrl();
  googleWebhookUrlInput.addEventListener('change', () => {
    saveWebhookUrl(googleWebhookUrlInput.value);
  });
}

// Map trainer names for case-insensitive lookup
const TRAINER_LOOKUP = new Map();
for (const trainer of HARD_SET_TRAINERS) {
  TRAINER_LOOKUP.set(trainer.trim().toLowerCase(), trainer.trim());
}

let previewData = null;
let loadedFileKey = null;

/**
 * Categorizes an External Reference Id value:
 * If it matches one of the hard-set trainers (case-insensitive), returns the canonical Trainer name.
 * Otherwise, it automatically routes into "member".
 */
function routeExternalReferenceId(value) {
  const raw = String(value || '').trim();
  if (!raw) {
    return {
      key: 'member',
      label: 'member'
    };
  }

  const normalized = raw.toLowerCase();
  if (TRAINER_LOOKUP.has(normalized)) {
    const canonical = TRAINER_LOOKUP.get(normalized);
    return {
      key: canonical.toLowerCase(),
      label: canonical
    };
  }

  // Any other value (numbers, missing, unlisted name, etc.) goes to member
  return {
    key: 'member',
    label: 'member'
  };
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

  // Identify the primary amount column if present
  const amountHeader = headers.find((h) => /^amount$/i.test(h)) ||
                       headers.find((h) => isSummableMoneyColumn(h)) || null;

  const idGroups = new Map();

  for (const row of rows) {
    const routed = routeExternalReferenceId(row[targetHeader]);
    row._assignedGroup = routed.label;

    if (!idGroups.has(routed.key)) {
      idGroups.set(routed.key, { key: routed.key, label: routed.label, count: 0 });
    }

    idGroups.get(routed.key).count += 1;
  }

  // Sort: Trainers first alphabetically, then "member" at the end
  const sortedGroups = Array.from(idGroups.values()).sort((a, b) => {
    if (a.key === 'member') return 1;
    if (b.key === 'member') return -1;
    return a.label.localeCompare(b.label);
  });

  return {
    sheetName: sourceName || firstSheetName,
    totalRows: rows.length,
    uniqueExternalReferenceIds: idGroups.size,
    targetHeader,
    amountHeader,
    headers,
    externalReferenceIds: sortedGroups,
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
    const routed = routeExternalReferenceId(row[payload.targetHeader]);

    if (!grouped.has(routed.key)) {
      grouped.set(routed.key, { label: routed.label, rows: [] });
    }

    grouped.get(routed.key).rows.push(row);
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
  if (commissionTableBody) commissionTableBody.innerHTML = '';
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
    const routed = routeExternalReferenceId(row[data.targetHeader]);
    const idMatches =
      selectedId === '__ALL__' ||
      routed.key === selectedId;

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
    selectedId === '__ALL__' ? 'all Assigned Groups' : `Group: ${selectedId}`;
  const headerText =
    selectedHeader === '__ANY__' ? 'all headers' : `header: ${selectedHeader}`;
  const queryText = query.trim() ? `query: "${query.trim()}"` : 'no text filter';

  filterMeta.textContent = `Showing ${rows.length} rows (${idText}, ${headerText}, ${queryText}).`;
}

/**
 * Calculates and renders the Payroll & Commission Summary Card
 */
function renderPayrollSummary(payload) {
  if (!commissionTableBody) return;
  commissionTableBody.innerHTML = '';

  const rates = getTrainerCommissionRates();
  const amountCol = payload.amountHeader;

  // Aggregate stats per trainer + member
  const stats = new Map();
  for (const trainer of HARD_SET_TRAINERS) {
    stats.set(trainer.toLowerCase(), {
      name: trainer,
      count: 0,
      totalSales: 0,
      rate: rates[trainer] !== undefined ? rates[trainer] : 50
    });
  }

  let memberCount = 0;
  let memberTotal = 0;

  for (const row of payload.rows) {
    const routed = routeExternalReferenceId(row[payload.targetHeader]);
    const val = amountCol ? (parseNumericValue(row[amountCol]) || 0) : 0;

    if (stats.has(routed.key)) {
      const item = stats.get(routed.key);
      item.count += 1;
      item.totalSales += val;
    } else {
      memberCount += 1;
      memberTotal += val;
    }
  }

  let totalPayoutAllTrainers = 0;
  let totalSalesAllTrainers = 0;
  let totalSessionsAllTrainers = 0;

  // Render each trainer row
  for (const trainer of HARD_SET_TRAINERS) {
    const item = stats.get(trainer.toLowerCase());
    const payout = (item.totalSales * item.rate) / 100;
    const gymRetained = item.totalSales - payout;

    totalPayoutAllTrainers += payout;
    totalSalesAllTrainers += item.totalSales;
    totalSessionsAllTrainers += item.count;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${item.name}</strong></td>
      <td class="text-center">${item.count}</td>
      <td class="text-end">$${formatAmount(item.totalSales)}</td>
      <td class="text-center" style="max-width: 110px;">
        <div class="input-group input-group-sm">
          <input type="number" min="0" max="100" step="1" class="form-control text-center trainer-rate-input" data-trainer="${item.name}" value="${item.rate}">
          <span class="input-group-text">%</span>
        </div>
      </td>
      <td class="text-end text-success fw-bold">$${formatAmount(payout)}</td>
      <td class="text-end text-muted">$${formatAmount(gymRetained)}</td>
    `;
    commissionTableBody.appendChild(tr);
  }

  // Member row
  const trMember = document.createElement('tr');
  trMember.className = 'table-light';
  trMember.innerHTML = `
    <td><span class="badge bg-secondary">member</span> <small class="text-muted">(General non-trainer sales)</small></td>
    <td class="text-center">${memberCount}</td>
    <td class="text-end">$${formatAmount(memberTotal)}</td>
    <td class="text-center text-muted">—</td>
    <td class="text-end text-muted">$0.00</td>
    <td class="text-end fw-bold">$${formatAmount(memberTotal)}</td>
  `;
  commissionTableBody.appendChild(trMember);

  // Grand total row
  const grandTotalSales = totalSalesAllTrainers + memberTotal;
  const grandGymRetained = grandTotalSales - totalPayoutAllTrainers;

  const trTotal = document.createElement('tr');
  trTotal.className = 'table-primary fw-bold';
  trTotal.innerHTML = `
    <td>TOTAL PAYROLL SUMMARY</td>
    <td class="text-center">${totalSessionsAllTrainers + memberCount}</td>
    <td class="text-end">$${formatAmount(grandTotalSales)}</td>
    <td class="text-center">—</td>
    <td class="text-end text-success">$${formatAmount(totalPayoutAllTrainers)}</td>
    <td class="text-end text-primary">$${formatAmount(grandGymRetained)}</td>
  `;
  commissionTableBody.appendChild(trTotal);

  // Attach listener to rate inputs
  const rateInputs = commissionTableBody.querySelectorAll('.trainer-rate-input');
  rateInputs.forEach(input => {
    input.addEventListener('change', (e) => {
      const tName = e.target.dataset.trainer;
      const newRate = Number(e.target.value) || 0;
      saveTrainerCommissionRate(tName, newRate);
      renderPayrollSummary(payload);
    });
  });
}

function renderPreview(payload) {
  const { totalRows, uniqueExternalReferenceIds, targetHeader, sheetName, externalReferenceIds } = payload;

  previewData = payload;
  previewMeta.textContent = `Sheet: ${sheetName} | Total Rows: ${totalRows} | Groups: ${uniqueExternalReferenceIds} (Hard-set Trainers + member catch-all) | Column: ${targetHeader}`;

  idFilter.innerHTML = '';

  const allOption = document.createElement('option');
  allOption.value = '__ALL__';
  allOption.textContent = 'All Assigned Groups';
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
  renderPayrollSummary(payload);
  previewSection.classList.remove('hidden');
}

// Build Exportable Excel Summary Workbook
function generatePayrollSummaryWorkbook(payload, cycle) {
  const rates = getTrainerCommissionRates();
  const amountCol = payload.amountHeader;
  const wb = XLSX.utils.book_new();

  // Tab 1: Summary Sheet
  const summaryRows = [
    ['PAYROLL & COMMISSION SUMMARY', ''],
    ['Pay Cycle:', cycle || 'N/A'],
    ['Generated on:', new Date().toLocaleString()],
    ['', ''],
    ['Trainer', 'Sessions/Txns', 'Total Sales ($)', 'Commission Rate', 'Trainer Payout ($)', 'Gym Retained ($)']
  ];

  let totalSalesAll = 0;
  let totalPayoutAll = 0;
  let totalSessionsAll = 0;

  for (const trainer of HARD_SET_TRAINERS) {
    const tRows = payload.rows.filter(r => routeExternalReferenceId(r[payload.targetHeader]).label === trainer);
    const count = tRows.length;
    const sales = tRows.reduce((sum, r) => sum + (amountCol ? (parseNumericValue(r[amountCol]) || 0) : 0), 0);
    const rate = rates[trainer] !== undefined ? rates[trainer] : 50;
    const payout = (sales * rate) / 100;
    const gymRetained = sales - payout;

    totalSalesAll += sales;
    totalPayoutAll += payout;
    totalSessionsAll += count;

    summaryRows.push([trainer, count, sales, `${rate}%`, payout, gymRetained]);
  }

  // Member summary
  const memberRows = payload.rows.filter(r => routeExternalReferenceId(r[payload.targetHeader]).label === 'member');
  const memberCount = memberRows.length;
  const memberSales = memberRows.reduce((sum, r) => sum + (amountCol ? (parseNumericValue(r[amountCol]) || 0) : 0), 0);
  summaryRows.push(['member (General non-trainer sales)', memberCount, memberSales, '0%', 0, memberSales]);

  summaryRows.push(['', '', '', '', '', '']);
  summaryRows.push(['TOTAL', totalSessionsAll + memberCount, totalSalesAll + memberSales, '', totalPayoutAll, (totalSalesAll + memberSales) - totalPayoutAll]);

  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, summarySheet, 'Payroll Summary');

  // Individual tabs for each trainer + member
  for (const trainer of HARD_SET_TRAINERS) {
    const tRows = payload.rows.filter(r => routeExternalReferenceId(r[payload.targetHeader]).label === trainer);
    if (tRows.length) {
      const sheet = XLSX.utils.json_to_sheet(tRows);
      XLSX.utils.book_append_sheet(wb, sheet, trainer);
    }
  }

  if (memberRows.length) {
    const mSheet = XLSX.utils.json_to_sheet(memberRows);
    XLSX.utils.book_append_sheet(wb, mSheet, 'member');
  }

  return wb;
}

// Sync to Master Google Sheet via Webhook (Apps Script)
if (syncGoogleSheetButton) {
  syncGoogleSheetButton.addEventListener('click', async () => {
    if (!previewData) {
      syncStatusEl.textContent = 'Please load a workbook first.';
      syncStatusEl.className = 'text-danger small mt-2';
      return;
    }

    const webhookUrl = (googleWebhookUrlInput ? googleWebhookUrlInput.value.trim() : '') || getSavedWebhookUrl();
    if (!webhookUrl) {
      syncStatusEl.textContent = 'Please provide a Google Apps Script Web App URL first.';
      syncStatusEl.className = 'text-danger small mt-2';
      return;
    }

    const cycle = payCycleInput ? payCycleInput.value.trim() : '';
    const rates = getTrainerCommissionRates();
    const amountCol = previewData.amountHeader;

    const trainerPayload = [];
    for (const trainer of HARD_SET_TRAINERS) {
      const tRows = previewData.rows.filter(r => routeExternalReferenceId(r[previewData.targetHeader]).label === trainer);
      const count = tRows.length;
      const sales = tRows.reduce((sum, r) => sum + (amountCol ? (parseNumericValue(r[amountCol]) || 0) : 0), 0);
      const rate = rates[trainer] !== undefined ? rates[trainer] : 50;
      const payout = (sales * rate) / 100;
      const gymRetained = sales - payout;

      trainerPayload.push({
        name: trainer,
        count: count,
        totalSales: sales,
        commissionRate: rate,
        payout: payout,
        gymRetained: gymRetained
      });
    }

    const memberRows = previewData.rows.filter(r => routeExternalReferenceId(r[previewData.targetHeader]).label === 'member');
    const memberCount = memberRows.length;
    const memberTotal = memberRows.reduce((sum, r) => sum + (amountCol ? (parseNumericValue(r[amountCol]) || 0) : 0), 0);

    const postBody = {
      payCycle: cycle || 'Current Cycle',
      trainers: trainerPayload,
      memberCount: memberCount,
      memberTotal: memberTotal,
      headers: previewData.headers,
      rawTransactions: previewData.rows
    };

    syncStatusEl.textContent = 'Saving to Master Google Sheet...';
    syncStatusEl.className = 'text-primary small mt-2';

    try {
      await fetch(webhookUrl, {
        method: 'POST',
        mode: 'no-cors', // allows cross-origin submission to Apps Script without CORS blockage
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(postBody)
      });

      syncStatusEl.textContent = `✓ Payroll for cycle "${cycle || 'Current'}" sent to Google Sheet! Check your master spreadsheet.`;
      syncStatusEl.className = 'text-success fw-bold small mt-2';
    } catch (err) {
      syncStatusEl.textContent = 'Error syncing to Google Sheet: ' + err.message;
      syncStatusEl.className = 'text-danger small mt-2';
    }
  });
}

// Export formatted Excel Summary workbook button
if (exportSummaryExcelBtn) {
  exportSummaryExcelBtn.addEventListener('click', () => {
    if (!previewData) {
      statusEl.textContent = 'Please load a workbook first.';
      return;
    }

    const cycle = payCycleInput ? payCycleInput.value.trim() : 'Payroll';
    const wb = generatePayrollSummaryWorkbook(previewData, cycle);
    const output = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([output], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const safeCycle = sanitizeFileName(cycle.replace(/\//g, '-'));
    downloadBlob(blob, `Payroll_Summary_${safeCycle}.xlsx`);
    statusEl.textContent = `Downloaded Payroll Summary Excel for ${cycle}.`;
  });
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
  const file = fileInput && fileInput.files ? fileInput.files[0] : null;

  if (!file && !previewData) {
    statusEl.textContent = 'Please choose a file or click "Load sheet" above first.';
    return;
  }

  const action = event.submitter?.value || 'split';
  statusEl.textContent = action === 'preview' ? 'Loading preview...' : 'Processing...';

  try {
    let payload = previewData;

    if (file) {
      payload = await loadWorkbookData(file);
    }

    if (!payload) {
      throw new Error('No data loaded. Choose a file or load a Google Sheet first.');
    }

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
