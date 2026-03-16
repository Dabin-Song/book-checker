require('dotenv').config();
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const axios = require('axios');
const http = require('http');
const fs = require('fs');

const httpAgent = new http.Agent({ keepAlive: true, maxSockets: 15 });

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.static('public'));

// ─── In-memory store ─────────────────────────────────────────────────────────
// purchase / library 각각: { books, headers, enrichedBooks }
// _lastResults: { duplicates, nonDuplicates }
const store = { purchase: null, library: null, _lastResults: null };

// ─── Column detection ─────────────────────────────────────────────────────────
const ISBN13_KEYS  = ['isbn13', 'isbn-13', 'isbn 13', '도서번호', 'isbn'];
const TITLE_KEYS   = ['서명', '도서명', '제목', '책이름', '서 명', 'title', '표제'];
const PUBLISHER_KEYS = ['출판사', '발행처', '발행사', '출 판 사', 'publisher'];

function normalizeKey(k) { return String(k).trim().toLowerCase().replace(/\s+/g, ' '); }

function findColumn(obj, candidates) {
  const keys = Object.keys(obj);
  for (const c of candidates) {
    const m = keys.find(k => normalizeKey(k) === c);
    if (m) return m;
  }
  for (const c of candidates) {
    const m = keys.find(k => normalizeKey(k).includes(c));
    if (m) return m;
  }
  return null;
}

function extractFields(row) {
  const isbn13Key    = findColumn(row, ISBN13_KEYS);
  const titleKey     = findColumn(row, TITLE_KEYS);
  const publisherKey = findColumn(row, PUBLISHER_KEYS);

  let isbn13 = isbn13Key ? String(row[isbn13Key]).trim() : '';
  isbn13 = isbn13.replace(/[-\s]/g, '');
  if (isbn13.length !== 13 || !/^\d+$/.test(isbn13)) isbn13 = '';

  return {
    isbn13,
    title:     titleKey     ? String(row[titleKey]).trim()     : '',
    publisher: publisherKey ? String(row[publisherKey]).trim() : '',
    _raw: row,
    _isbn13Key:    isbn13Key,
    _titleKey:     titleKey,
    _publisherKey: publisherKey,
  };
}

// ─── Excel parsing ────────────────────────────────────────────────────────────
function parseExcel(filePath) {
  const wb = XLSX.readFile(filePath, { type: 'file', cellDates: true, raw: false });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  if (!rawRows.length) return { rows: [], headers: [] };
  const headers = Object.keys(rawRows[0]);
  const rows = rawRows.filter(r => Object.values(r).some(v => String(v).trim() !== ''));
  return { rows, headers };
}

// ─── Aladin API ───────────────────────────────────────────────────────────────
async function fetchIsbn13ByTitle(title, apiKey) {
  try {
    const res = await axios.get('http://www.aladin.co.kr/ttb/api/ItemSearch.aspx', {
      params: {
        ttbkey: apiKey, Query: title, QueryType: 'Title',
        MaxResults: 1, start: 1, SearchTarget: 'Book',
        output: 'js', Version: '20131101',
      },
      timeout: 5000,
      httpAgent,
    });
    const item = res.data?.item?.[0];
    if (item?.isbn13) return String(item.isbn13).replace(/[-\s]/g, '');
  } catch (_) { /* ignore */ }
  return '';
}

async function enrichBooks(books, apiKey, onProgress) {
  const CONCURRENCY = 10;
  const targets = books
    .map((b, idx) => ({ b, idx }))
    .filter(({ b }) => !b.isbn13 && b.title);

  let done = 0, enriched = 0;

  async function runPool(tasks, concurrency, worker) {
    let i = 0;
    const next = async () => {
      if (i >= tasks.length) return;
      const t = tasks[i++];
      await worker(t);
      await next();
    };
    await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, next));
  }

  await runPool(targets, CONCURRENCY, async ({ b, idx }) => {
    const isbn13 = await fetchIsbn13ByTitle(b.title, apiKey);
    if (isbn13) { books[idx].isbn13 = isbn13; enriched++; }
    done++;
    onProgress?.(done, targets.length, enriched);
  });

  return { total: targets.length, enriched };
}

// ─── Duplicate detection ──────────────────────────────────────────────────────
function normalizeTitle(t) {
  return String(t).trim().toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^\w\s가-힣ㄱ-ㅎㅏ-ㅣ]/g, '');
}

function compare(purchaseBooks, libraryBooks) {
  const byIsbn  = new Map();
  const byTitle = new Map();
  const byTitlePub = new Map();

  for (const b of libraryBooks) {
    if (b.isbn13) byIsbn.set(b.isbn13, b);
    const nt = normalizeTitle(b.title);
    if (nt) {
      if (!byTitle.has(nt)) byTitle.set(nt, b);
      const k = `${nt}||${normalizeTitle(b.publisher)}`;
      if (!byTitlePub.has(k)) byTitlePub.set(k, b);
    }
  }

  const duplicates    = [];
  const nonDuplicates = [];

  for (const p of purchaseBooks) {
    let match = null, method = '';

    if (p.isbn13 && byIsbn.has(p.isbn13))           { match = byIsbn.get(p.isbn13);      method = 'ISBN13'; }
    if (!match) {
      const nt = normalizeTitle(p.title);
      if (nt && byTitle.has(nt))                    { match = byTitle.get(nt);            method = '서명'; }
    }
    if (!match) {
      const k = `${normalizeTitle(p.title)}||${normalizeTitle(p.publisher)}`;
      if (byTitlePub.has(k))                        { match = byTitlePub.get(k);          method = '서명+출판사'; }
    }

    if (match) duplicates.push({ purchase: p, library: match, matchMethod: method });
    else       nonDuplicates.push({ purchase: p });
  }

  return { duplicates, nonDuplicates };
}

// ─── Excel export ─────────────────────────────────────────────────────────────
const HEADER_FILL  = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const HEADER_FONT  = { bold: true };

async function buildEnrichedWorkbook(type) {
  const d = store[type];
  if (!d) return null;
  const books   = d.enrichedBooks || d.books;
  const headers = d.headers;

  const wb    = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet(type === 'purchase' ? '구입예정목록' : '소장목록');
  const hr    = sheet.addRow(headers);
  hr.font = HEADER_FONT;
  hr.fill = HEADER_FILL('FFE2EFDA');

  for (const book of books) {
    const row = { ...book._raw };
    if (book._isbn13Key && book.isbn13) row[book._isbn13Key] = book.isbn13;
    sheet.addRow(headers.map(h => row[h] ?? ''));
  }
  sheet.columns.forEach(c => { c.width = 22; });
  return wb;
}

async function buildResultsWorkbook() {
  if (!store._lastResults) return null;
  const { duplicates, nonDuplicates } = store._lastResults;
  const wb = new ExcelJS.Workbook();

  // Sheet 1: 중복 목록
  const s1  = wb.addWorksheet('중복목록');
  const h1  = s1.addRow(['No.', 'ISBN13', '구입예정 도서명', '구입예정 출판사', '소장 도서명', '소장 출판사', '대조 기준']);
  h1.font = HEADER_FONT; h1.fill = HEADER_FILL('FFFCE4D6');
  duplicates.forEach((d, i) => s1.addRow([
    i + 1,
    d.purchase.isbn13 || d.library.isbn13 || '',
    d.purchase.title, d.purchase.publisher,
    d.library.title,  d.library.publisher,
    d.matchMethod,
  ]));
  s1.columns = [{ width:5 },{ width:16 },{ width:36 },{ width:18 },{ width:36 },{ width:18 },{ width:14 }];

  // Sheet 2: 구입 가능 목록
  const s2  = wb.addWorksheet('구입가능목록');
  const h2  = s2.addRow(['No.', 'ISBN13', '도서명', '출판사']);
  h2.font = HEADER_FONT; h2.fill = HEADER_FILL('FFDAE8FC');
  nonDuplicates.forEach((d, i) => s2.addRow([
    i + 1,
    d.purchase.isbn13 || '',
    d.purchase.title,
    d.purchase.publisher,
  ]));
  s2.columns = [{ width:5 },{ width:16 },{ width:40 },{ width:20 }];

  return wb;
}

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/api/config', (_, res) =>
  res.json({ hasServerKey: !!process.env.ALADIN_API_KEY })
);

app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const type = req.body.type;
    if (!['purchase', 'library'].includes(type))
      return res.status(400).json({ error: '잘못된 파일 유형입니다.' });

    const { rows, headers } = parseExcel(req.file.path);
    const books = rows.map(extractFields);

    store[type] = { books, headers, enrichedBooks: null };

    const sample = books[0] || {};
    fs.unlink(req.file.path, () => {});

    res.json({
      count: books.length,
      missingIsbn13Count: books.filter(b => !b.isbn13).length,
      detectedColumns: {
        isbn13:    sample._isbn13Key    || null,
        title:     sample._titleKey     || null,
        publisher: sample._publisherKey || null,
      },
      headers,
    });
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: '파일 파싱 중 오류가 발생했습니다: ' + err.message });
  }
});

// SSE: enrich → compare
app.post('/api/compare', async (req, res) => {
  if (!store.purchase || !store.library)
    return res.status(400).json({ error: '두 파일 모두 업로드해주세요.' });

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const send = (event, data) =>
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);

  try {
    const apiKey = req.body.apiKey || process.env.ALADIN_API_KEY || '';

    // 작업 대상 배열 (원본 복사 — enrichedBooks 에 별도 저장)
    const purchaseBooks = store.purchase.books.map(b => ({ ...b }));
    const libraryBooks  = store.library.books.map(b =>  ({ ...b }));

    const pMissing = purchaseBooks.filter(b => !b.isbn13).length;
    const lMissing = libraryBooks.filter(b =>  !b.isbn13).length;

    // ── Step 1: 구입 예정 목록 ISBN13 보강 ──────────────────────────────
    if (pMissing > 0 && apiKey) {
      send('progress', { step: 'enrich_purchase', message: `구입 예정 목록 ISBN13 조회 중... (${pMissing}건)`, current: 0, total: pMissing });
      await enrichBooks(purchaseBooks, apiKey, (cur, tot, enriched) =>
        send('progress', { step: 'enrich_purchase', message: '구입 예정 목록 ISBN13 조회 중...', current: cur, total: tot, enriched })
      );
    }
    // enriched 결과를 store 에 저장 (다운로드용)
    store.purchase.enrichedBooks = purchaseBooks;

    // ── Step 2: 소장 목록 ISBN13 보강 ───────────────────────────────────
    if (lMissing > 0 && apiKey) {
      send('progress', { step: 'enrich_library', message: `소장 목록 ISBN13 조회 중... (${lMissing}건)`, current: 0, total: lMissing });
      await enrichBooks(libraryBooks, apiKey, (cur, tot, enriched) =>
        send('progress', { step: 'enrich_library', message: '소장 목록 ISBN13 조회 중...', current: cur, total: tot, enriched })
      );
    }
    store.library.enrichedBooks = libraryBooks;

    // ── Step 3: 비교 (ISBN13 우선) ──────────────────────────────────────
    send('progress', { step: 'comparing', message: '중복 검사 중...' });
    const { duplicates, nonDuplicates } = compare(purchaseBooks, libraryBooks);

    const slim = b => ({ isbn13: b.isbn13, title: b.title, publisher: b.publisher });
    store._lastResults = {
      duplicates:    duplicates.map(d => ({ purchase: slim(d.purchase), library: slim(d.library), matchMethod: d.matchMethod })),
      nonDuplicates: nonDuplicates.map(d => ({ purchase: slim(d.purchase) })),
    };

    send('result', {
      ...store._lastResults,
      purchaseCount:    purchaseBooks.length,
      libraryCount:     libraryBooks.length,
      pEnrichedCount:   pMissing,
      lEnrichedCount:   lMissing,
    });
  } catch (err) {
    console.error('Compare error:', err);
    send('error', { message: '비교 중 오류가 발생했습니다: ' + err.message });
  }

  res.end();
});

// ── Downloads ────────────────────────────────────────────────────────────────
app.get('/api/download/enriched/:type', async (req, res) => {
  const type = req.params.type;
  if (!['purchase', 'library'].includes(type) || !store[type])
    return res.status(400).json({ error: '데이터가 없습니다.' });

  const wb = await buildEnrichedWorkbook(type);
  const label = type === 'purchase' ? '구입예정목록_ISBN13보완' : '소장목록_ISBN13보완';
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(label)}.xlsx`);
  await wb.xlsx.write(res);
  res.end();
});

app.get('/api/download/results', async (req, res) => {
  if (!store._lastResults)
    return res.status(400).json({ error: '먼저 중복 검사를 실행해주세요.' });

  const wb = await buildResultsWorkbook();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent('중복검사_결과')}.xlsx`);
  await wb.xlsx.write(res);
  res.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Book Checker running at http://localhost:${PORT}`));
