// ── State ──────────────────────────────────────────────────────────────────
const state = {
  purchase: null,  // { count, missingIsbn13Count, fileName }
  library: null,
  results: null,
  activeFilter: null, // 'ISBN13' | '서명' | '서명+출판사' | null
};

// ── DOM refs ───────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

const purchaseFile = $('purchaseFile');
const libraryFile = $('libraryFile');
const compareBtn = $('compareBtn');
const apiKeyInput = $('apiKey');

// ── API Key visibility toggle ───────────────────────────────────────────────
$('toggleApiKey').addEventListener('click', () => {
  apiKeyInput.type = apiKeyInput.type === 'password' ? 'text' : 'password';
});
apiKeyInput.type = 'password';

// ── File Upload ────────────────────────────────────────────────────────────
function setupDropZone(dropZoneId, inputId, type) {
  const dropZone = $(dropZoneId);
  const input = $(inputId);

  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) uploadFile(file, type);
  });
  dropZone.addEventListener('click', e => {
    if (!e.target.classList.contains('btn-upload') && e.target.tagName !== 'LABEL') {
      input.click();
    }
  });
  input.addEventListener('change', () => {
    if (input.files[0]) uploadFile(input.files[0], type);
  });
}

setupDropZone('purchaseDropZone', 'purchaseFile', 'purchase');
setupDropZone('libraryDropZone', 'libraryFile', 'library');

$('purchaseRemove').addEventListener('click', () => resetFile('purchase'));
$('libraryRemove').addEventListener('click', () => resetFile('library'));

function resetFile(type) {
  state[type] = null;
  $(`${type}Info`).classList.add('hidden');
  $(`${type}DropZone`).classList.remove('hidden');
  $(`${type}File`).value = '';
  updateCompareBtn();
}

async function uploadFile(file, type) {
  if (!file.name.match(/\.xlsx?$/i)) {
    showToast('엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.', 'error');
    return;
  }

  const dropZone = $(`${type}DropZone`);
  const info = $(`${type}Info`);
  dropZone.style.opacity = '0.5';

  const formData = new FormData();
  formData.append('file', file);
  formData.append('type', type);

  try {
    const res = await fetch('/api/upload', { method: 'POST', body: formData });
    const data = await res.json();

    if (!res.ok) throw new Error(data.error || '업로드 실패');

    state[type] = { ...data, fileName: file.name };

    // Show info panel
    dropZone.classList.add('hidden');
    $(`${type}FileName`).textContent = `📄 ${file.name}`;
    $(`${type}Count`).textContent = data.count.toLocaleString();
    $(`${type}Missing`).textContent = data.missingIsbn13Count.toLocaleString();

    const missingWrap = $(`${type}MissingWrap`);
    missingWrap.style.display = data.missingIsbn13Count > 0 ? '' : 'none';

    // Column detection
    const colDiv = $(`${type}Columns`);
    const cols = data.detectedColumns;
    colDiv.innerHTML =
      `<strong>감지된 열:</strong> ` +
      [
        cols.isbn13 ? `ISBN13 → <em>${cols.isbn13}</em>` : '<span style="color:#e53e3e">ISBN13 열 없음</span>',
        cols.title ? `서명 → <em>${cols.title}</em>` : '<span style="color:#e53e3e">서명 열 없음</span>',
        cols.publisher ? `출판사 → <em>${cols.publisher}</em>` : '출판사 열 없음',
      ].join(' &nbsp;|&nbsp; ');

    info.classList.remove('hidden');
    updateCompareBtn();
  } catch (err) {
    showToast(err.message, 'error');
  } finally {
    dropZone.style.opacity = '';
  }
}

function updateCompareBtn() {
  compareBtn.disabled = !(state.purchase && state.library);
}

// ── Compare ────────────────────────────────────────────────────────────────
compareBtn.addEventListener('click', startCompare);

async function startCompare() {
  const progressCard = $('progressCard');
  const resultsCard = $('resultsCard');

  compareBtn.disabled = true;
  progressCard.classList.remove('hidden');
  resultsCard.classList.add('hidden');
  $('progressSteps').innerHTML = '';

  const apiKey = apiKeyInput.value.trim();

  try {
    const res = await fetch('/api/compare', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ apiKey }),
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error);
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });

      const messages = buffer.split('\n\n');
      buffer = messages.pop();

      for (const msg of messages) {
        const eventMatch = msg.match(/^event: (.+)/m);
        const dataMatch = msg.match(/^data: (.+)/m);
        if (!eventMatch || !dataMatch) continue;
        const event = eventMatch[1].trim();
        const data = JSON.parse(dataMatch[1]);
        handleSSEEvent(event, data);
      }
    }
  } catch (err) {
    showToast(err.message, 'error');
    compareBtn.disabled = false;
  }
}

function handleSSEEvent(event, data) {
  if (event === 'progress') {
    updateProgress(data);
  } else if (event === 'result') {
    showResults(data);
    $('compareBtn').disabled = false;
  } else if (event === 'error') {
    showToast(data.message, 'error');
    $('compareBtn').disabled = false;
  }
}

const progressState = {};

function updateProgress(data) {
  const steps = $('progressSteps');
  const key = data.step;

  if (!progressState[key]) {
    progressState[key] = document.createElement('div');
    progressState[key].className = 'progress-step';
    steps.appendChild(progressState[key]);
  }

  const el = progressState[key];
  const pct = data.total ? Math.round((data.current / data.total) * 100) : 100;
  const icon = pct === 100 ? '✅' : '⏳';

  el.innerHTML = `
    <span class="step-icon">${icon}</span>
    <span class="step-msg">${data.message}</span>
    ${data.total ? `
      <div class="progress-bar-wrap">
        <div class="progress-bar" style="width:${pct}%"></div>
      </div>
      <span class="step-pct">${pct}%</span>
    ` : ''}
  `;
}

// ── Results Display ────────────────────────────────────────────────────────
function showResults(data) {
  state.results = data;
  state.activeFilter = null;

  const { duplicates, purchaseCount, libraryCount } = data;

  // Summary chips
  $('resultsSummary').innerHTML = `
    <span class="summary-chip chip-red">중복 ${duplicates.length.toLocaleString()}권</span>
    <span class="summary-chip chip-blue">구입 예정 ${purchaseCount.toLocaleString()}권</span>
    <span class="summary-chip chip-green">소장 ${libraryCount.toLocaleString()}권</span>
  `;

  // Match method breakdown
  const methods = {};
  for (const d of duplicates) {
    methods[d.matchMethod] = (methods[d.matchMethod] || 0) + 1;
  }

  const methodStyles = {
    'ISBN13': 'method-isbn',
    '서명': 'method-title',
    '서명+출판사': 'method-title-pub',
  };

  const breakdown = $('matchBreakdown');
  breakdown.innerHTML = '';
  for (const [method, count] of Object.entries(methods)) {
    const chip = document.createElement('span');
    chip.className = `method-chip ${methodStyles[method] || ''} active`;
    chip.textContent = `${method} 기준: ${count}권`;
    chip.dataset.method = method;
    chip.addEventListener('click', () => toggleFilter(method));
    breakdown.appendChild(chip);
  }

  const resultsCard = $('resultsCard');
  resultsCard.classList.remove('hidden');

  if (duplicates.length === 0) {
    $('duplicatesSection').classList.add('hidden');
    $('noDuplicateMsg').classList.remove('hidden');
  } else {
    $('duplicatesSection').classList.remove('hidden');
    $('noDuplicateMsg').classList.add('hidden');
    renderTable(duplicates);
  }

  // Download buttons
  $('downloadBar').classList.remove('hidden');
  if (data.hasEnrichedLibrary) {
    $('btnDownloadEnriched').classList.remove('hidden');
  } else {
    $('btnDownloadEnriched').classList.add('hidden');
  }

  // Scroll to results
  resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function toggleFilter(method) {
  state.activeFilter = state.activeFilter === method ? null : method;

  // Update chip active state
  document.querySelectorAll('.method-chip').forEach(chip => {
    if (state.activeFilter === null) {
      chip.classList.add('active');
    } else {
      chip.classList.toggle('active', chip.dataset.method === state.activeFilter);
    }
  });

  applyFilters();
}

$('searchFilter').addEventListener('input', applyFilters);

function applyFilters() {
  const query = $('searchFilter').value.toLowerCase();
  const { duplicates } = state.results;

  const filtered = duplicates.filter(d => {
    if (state.activeFilter && d.matchMethod !== state.activeFilter) return false;
    if (!query) return true;
    return (
      (d.purchase.isbn13 || '').includes(query) ||
      (d.purchase.title || '').toLowerCase().includes(query) ||
      (d.purchase.publisher || '').toLowerCase().includes(query) ||
      (d.library.title || '').toLowerCase().includes(query) ||
      (d.library.publisher || '').toLowerCase().includes(query)
    );
  });

  renderTable(filtered);
}

const methodBadgeStyle = {
  'ISBN13': 'background:#e9d8fd;color:#553c9a',
  '서명': 'background:#fefcbf;color:#744210',
  '서명+출판사': 'background:#fed7e2;color:#97266d',
};

function renderTable(rows) {
  const tbody = $('duplicatesBody');
  const noResults = $('noResults');

  if (rows.length === 0) {
    tbody.innerHTML = '';
    noResults.classList.remove('hidden');
    return;
  }
  noResults.classList.add('hidden');

  tbody.innerHTML = rows.map((d, i) => `
    <tr>
      <td>${i + 1}</td>
      <td class="isbn">${d.purchase.isbn13 || d.library.isbn13 || '—'}</td>
      <td class="title">${esc(d.purchase.title) || '—'}</td>
      <td class="pub">${esc(d.purchase.publisher) || '—'}</td>
      <td class="title">${esc(d.library.title) || '—'}</td>
      <td class="pub">${esc(d.library.publisher) || '—'}</td>
      <td class="method"><span style="${methodBadgeStyle[d.matchMethod] || ''}">${d.matchMethod}</span></td>
    </tr>
  `).join('');
}

// ── Helpers ────────────────────────────────────────────────────────────────
function esc(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function showToast(message, type = 'info') {
  // Simple inline notification
  let toast = document.querySelector('.toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.className = 'toast';
    toast.style.cssText = `
      position:fixed; bottom:1.5rem; left:50%; transform:translateX(-50%);
      padding:.75rem 1.5rem; border-radius:10px; font-size:.88rem; font-weight:600;
      z-index:9999; max-width:90vw; text-align:center; box-shadow:0 4px 16px rgba(0,0,0,.15);
      animation: fadein .2s;
    `;
    document.body.appendChild(toast);
  }
  toast.style.background = type === 'error' ? '#fc8181' : '#48bb78';
  toast.style.color = '#fff';
  toast.textContent = message;
  toast.style.display = 'block';
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { toast.style.display = 'none'; }, 4000);
}
