// ── State ─────────────────────────────────────────────────────────────────────
const state = {
  purchase: null,
  library:  null,
  results:  null,
  dupFilter: null,
  apiKey: null,
};

// ── 알라딘 API (브라우저 직접 호출) ───────────────────────────────────────────
// output=js 응답이 jQuery.fn({...}) 형태이므로 script 태그 주입으로 처리
// jQuery.fn 충돌 방지를 위해 요청을 직렬화(queue)
const _aladinQueue = (() => {
  let tail = Promise.resolve();
  return fn => { tail = tail.then(fn); return tail; };
})();

function fetchIsbn13ByTitle(title, apiKey) {
  return _aladinQueue(() => new Promise(resolve => {
    const prev = window.jQuery?.fn;
    if (!window.jQuery) window.jQuery = {};

    let done = false;
    const finish = isbn13 => {
      if (done) return;
      done = true;
      window.jQuery.fn = prev;
      if (script.parentNode) script.remove();
      clearTimeout(timer);
      resolve(isbn13 || '');
    };

    window.jQuery.fn = data => finish(data?.item?.[0]?.isbn13 || '');
    const timer = setTimeout(() => finish(''), 6000);

    const params = new URLSearchParams({
      ttbkey: apiKey, Query: title, QueryType: 'Title',
      MaxResults: 1, start: 1, SearchTarget: 'Book',
      output: 'js', Version: '20131101',
    });
    const script = document.createElement('script');
    script.src = `https://www.aladin.co.kr/ttb/api/ItemSearch.aspx?${params}`;
    script.onerror = () => finish('');
    document.head.appendChild(script);
  }));
}

async function enrichBooks(books, onProgress) {
  const apiKey = state.apiKey;
  if (!apiKey || !books?.length) return [];
  const results = [];
  for (let i = 0; i < books.length; i++) {
    const isbn13 = await fetchIsbn13ByTitle(books[i].title, apiKey);
    results.push({ idx: books[i].idx, isbn13 });
    onProgress(i + 1, books.length, results.filter(r => r.isbn13).length);
  }
  return results;
}

const $ = id => document.getElementById(id);
const compareBtn = $('compareBtn');

// ── API Status indicator ──────────────────────────────────────────────────────
(async () => {
  const el   = $('apiStatus');
  const text = $('apiStatusText');
  try {
    const cfg = await (await fetch('/api/config')).json();
    state.apiKey = cfg.apiKey || null;
    if (cfg.apiKey) {
      el.classList.add('ok');
      text.textContent = '알라딘 API 연동됨';
    } else {
      el.classList.add('fail');
      text.textContent = '알라딘 API 미설정';
    }
  } catch (_) {
    el.classList.add('fail');
    text.textContent = '서버 연결 오류';
  }
})();

// ── File Upload ───────────────────────────────────────────────────────────────
function setupDropZone(dropZoneId, inputId, type) {
  const dz    = $(dropZoneId);
  const input = $(inputId);

  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('dragover');
    if (e.dataTransfer.files[0]) uploadFile(e.dataTransfer.files[0], type);
  });
  dz.addEventListener('click', e => {
    if (!e.target.classList.contains('btn-upload') && e.target.tagName !== 'LABEL') input.click();
  });
  input.addEventListener('change', () => { if (input.files[0]) uploadFile(input.files[0], type); });
}

setupDropZone('purchaseDropZone', 'purchaseFile', 'purchase');
setupDropZone('libraryDropZone',  'libraryFile',  'library');

$('purchaseRemove').addEventListener('click', () => resetFile('purchase'));
$('libraryRemove').addEventListener('click',  () => resetFile('library'));

function resetFile(type) {
  state[type] = null;
  $(`${type}Info`).classList.add('hidden');
  $(`${type}DropZone`).classList.remove('hidden');
  $(`${type}File`).value = '';
  updateCompareBtn();
}

async function uploadFile(file, type) {
  if (!file.name.match(/\.xlsx?$/i)) { toast('엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.', 'error'); return; }

  const dz = $(`${type}DropZone`);
  dz.style.opacity = '0.5';

  try {
    const fd = new FormData();
    fd.append('file', file);
    fd.append('type', type);
    const res  = await fetch('/api/upload', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error);

    state[type] = { ...data, fileName: file.name, booksWithoutIsbn13: data.booksWithoutIsbn13 || [] };
    dz.classList.add('hidden');

    $(`${type}FileName`).textContent = `📄 ${file.name}`;
    $(`${type}Count`).textContent    = data.count.toLocaleString();
    $(`${type}Missing`).textContent  = data.missingIsbn13Count.toLocaleString();
    $(`${type}MissingWrap`).style.display = data.missingIsbn13Count > 0 ? '' : 'none';

    const c = data.detectedColumns;
    $(`${type}Columns`).innerHTML =
      '<strong>감지된 열:</strong> ' + [
        c.isbn13    ? `ISBN13 → <em>${c.isbn13}</em>`       : '<span style="color:#e53e3e">ISBN13 열 없음</span>',
        c.title     ? `서명 → <em>${c.title}</em>`          : '<span style="color:#e53e3e">서명 열 없음</span>',
        c.publisher ? `출판사 → <em>${c.publisher}</em>`    : '출판사 열 없음',
      ].join(' &nbsp;|&nbsp; ');

    $(`${type}Info`).classList.remove('hidden');
    updateCompareBtn();
  } catch (err) {
    toast(err.message, 'error');
  } finally {
    dz.style.opacity = '';
  }
}

function updateCompareBtn() {
  compareBtn.disabled = !(state.purchase && state.library);
}

// ── Compare ───────────────────────────────────────────────────────────────────
compareBtn.addEventListener('click', startCompare);

async function startCompare() {
  compareBtn.disabled = true;
  $('progressCard').classList.remove('hidden');
  $('resultsCard').classList.add('hidden');
  $('progressSteps').innerHTML = '';
  Object.keys(progressEls).forEach(k => delete progressEls[k]);

  try {
    // ── Step 1 & 2: 브라우저에서 알라딘 API 직접 호출하여 ISBN13 보강 ──
    const enrichedPurchase = await enrichStep('purchase');
    const enrichedLibrary  = await enrichStep('library');

    // ── Step 3: 서버에 비교 요청 ────────────────────────────────────────
    updateProgress({ step: 'comparing', message: '중복 검사 중...' });
    const res = await fetch('/api/compare', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ enrichedPurchase, enrichedLibrary }),
    });
    if (!res.ok) { const e = await res.json(); throw new Error(e.error); }

    const reader  = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const parts = buf.split('\n\n');
      buf = parts.pop();
      for (const msg of parts) {
        const ev   = msg.match(/^event: (.+)/m)?.[1]?.trim();
        const data = msg.match(/^data: (.+)/m)?.[1];
        if (ev && data) handleSSE(ev, JSON.parse(data));
      }
    }
  } catch (err) {
    toast(err.message, 'error');
    compareBtn.disabled = false;
  }
}

async function enrichStep(type) {
  const books = state[type]?.booksWithoutIsbn13 || [];
  if (!books.length || !state.apiKey) return [];
  const label = type === 'purchase' ? '구입 예정 목록' : '소장 목록';
  updateProgress({ step: `enrich_${type}`, message: `${label} ISBN13 조회 중...`, current: 0, total: books.length });
  return enrichBooks(books, (cur, tot, found) =>
    updateProgress({ step: `enrich_${type}`, message: `${label} ISBN13 조회 중...`, current: cur, total: tot, enriched: found })
  );
}

const progressEls = {};
function handleSSE(event, data) {
  if (event === 'progress') updateProgress(data);
  else if (event === 'result') { showResults(data); compareBtn.disabled = false; }
  else if (event === 'error')  { toast(data.message, 'error'); compareBtn.disabled = false; }
}

function updateProgress(data) {
  if (!progressEls[data.step]) {
    progressEls[data.step] = document.createElement('div');
    progressEls[data.step].className = 'progress-step';
    $('progressSteps').appendChild(progressEls[data.step]);
  }
  const el  = progressEls[data.step];
  const pct = data.total ? Math.round((data.current / data.total) * 100) : 100;
  const failNote = (data.failed > 0)
    ? `<span style="color:#e53e3e;font-size:.8rem;margin-left:.5rem">⚠ ${data.failed}건 실패 (${data.lastError})</span>`
    : '';
  el.innerHTML = `
    <span class="step-icon">${pct === 100 ? '✅' : '⏳'}</span>
    <span class="step-msg">${data.message}</span>${failNote}
    ${data.total ? `
      <div class="progress-bar-wrap"><div class="progress-bar" style="width:${pct}%"></div></div>
      <span class="step-pct">${pct}%</span>` : ''}
  `;
}

// ── Results ───────────────────────────────────────────────────────────────────
function showResults(data) {
  state.results   = data;
  state.dupFilter = null;

  const { duplicates, nonDuplicates, purchaseCount, libraryCount, pEnrichedCount, lEnrichedCount } = data;

  // Summary
  $('resultsSummary').innerHTML = `
    <span class="summary-chip chip-red">중복 ${duplicates.length.toLocaleString()}권</span>
    <span class="summary-chip chip-blue2">구입 가능 ${nonDuplicates.length.toLocaleString()}권</span>
    <span class="summary-chip chip-blue">구입 예정 ${purchaseCount.toLocaleString()}권</span>
    <span class="summary-chip chip-green">소장 ${libraryCount.toLocaleString()}권</span>
  `;

  // Match breakdown chips
  const methods = {};
  duplicates.forEach(d => { methods[d.matchMethod] = (methods[d.matchMethod] || 0) + 1; });
  const bd = $('matchBreakdown');
  bd.innerHTML = '';
  Object.entries(methods).forEach(([m, cnt]) => {
    const chip = document.createElement('span');
    chip.className  = `method-chip ${methodClass[m] || ''} active`;
    chip.textContent = `${m}: ${cnt}권`;
    chip.dataset.method = m;
    chip.addEventListener('click', () => toggleDupFilter(m));
    bd.appendChild(chip);
  });

  renderDupTable(duplicates);
  renderAvTable(nonDuplicates);

  // Download buttons
  if (pEnrichedCount > 0) $('dlPurchase').classList.remove('hidden');
  if (lEnrichedCount > 0) $('dlLibrary').classList.remove('hidden');

  $('resultsCard').classList.remove('hidden');
  $('resultsCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ── Tabs ──────────────────────────────────────────────────────────────────────
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.add('hidden'));
    tab.classList.add('active');
    $(`panel${cap(tab.dataset.tab)}`).classList.remove('hidden');
  });
});
const cap = s => s.charAt(0).toUpperCase() + s.slice(1);

// ── Duplicate table ───────────────────────────────────────────────────────────
const methodClass = { 'ISBN13': 'method-isbn', '서명': 'method-title', '서명+출판사': 'method-title-pub' };
const methodStyle = { 'ISBN13': 'background:#e9d8fd;color:#553c9a', '서명': 'background:#fefcbf;color:#744210', '서명+출판사': 'background:#fed7e2;color:#97266d' };

function toggleDupFilter(m) {
  state.dupFilter = state.dupFilter === m ? null : m;
  document.querySelectorAll('.method-chip').forEach(c =>
    c.classList.toggle('active', state.dupFilter === null || c.dataset.method === state.dupFilter)
  );
  applyDupFilter();
}

$('dupSearch').addEventListener('input', applyDupFilter);

function applyDupFilter() {
  const q = $('dupSearch').value.toLowerCase();
  const rows = state.results.duplicates.filter(d => {
    if (state.dupFilter && d.matchMethod !== state.dupFilter) return false;
    if (!q) return true;
    return [d.purchase.isbn13, d.purchase.title, d.purchase.publisher,
            d.library.title,   d.library.publisher].some(v => (v||'').toLowerCase().includes(q));
  });
  renderDupTable(rows);
}

function renderDupTable(rows) {
  const body = $('dupBody');
  $('dupNoResults').classList.toggle('hidden', rows.length > 0);
  body.innerHTML = rows.map((d, i) => `
    <tr>
      <td>${i+1}</td>
      <td class="isbn">${d.purchase.isbn13 || d.library.isbn13 || '—'}</td>
      <td class="title">${esc(d.purchase.title) || '—'}</td>
      <td class="pub">${esc(d.purchase.publisher) || '—'}</td>
      <td class="title">${esc(d.library.title) || '—'}</td>
      <td class="pub">${esc(d.library.publisher) || '—'}</td>
      <td><span style="${methodStyle[d.matchMethod]||''};padding:.2rem .55rem;border-radius:999px;font-size:.75rem;font-weight:600">${d.matchMethod}</span></td>
    </tr>`).join('');
}

// ── Available table ───────────────────────────────────────────────────────────
$('avSearch').addEventListener('input', applyAvFilter);

function applyAvFilter() {
  const q = $('avSearch').value.toLowerCase();
  const rows = state.results.nonDuplicates.filter(d =>
    !q || [d.purchase.isbn13, d.purchase.title, d.purchase.publisher]
      .some(v => (v||'').toLowerCase().includes(q))
  );
  renderAvTable(rows);
}

function renderAvTable(rows) {
  const body = $('avBody');
  $('avNoResults').classList.toggle('hidden', rows.length > 0);
  body.innerHTML = rows.map((d, i) => `
    <tr>
      <td>${i+1}</td>
      <td class="isbn">${d.purchase.isbn13 || '—'}</td>
      <td class="title">${esc(d.purchase.title) || '—'}</td>
      <td class="pub">${esc(d.purchase.publisher) || '—'}</td>
    </tr>`).join('');
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function esc(s) {
  return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function toast(msg, type = 'info') {
  let el = document.querySelector('.toast');
  if (!el) {
    el = document.createElement('div');
    el.className = 'toast';
    el.style.cssText = 'position:fixed;bottom:1.5rem;left:50%;transform:translateX(-50%);padding:.75rem 1.5rem;border-radius:10px;font-size:.88rem;font-weight:600;z-index:9999;max-width:90vw;text-align:center;box-shadow:0 4px 16px rgba(0,0,0,.15)';
    document.body.appendChild(el);
  }
  el.style.background = type === 'error' ? '#fc8181' : '#48bb78';
  el.style.color = '#fff';
  el.textContent = msg;
  el.style.display = 'block';
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.style.display = 'none'; }, 4000);
}
