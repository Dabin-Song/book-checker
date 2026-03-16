const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const axios = require('axios');
const http = require('http');
const path = require('path');
const fs = require('fs');

// Reuse TCP connections across Aladin API requests
const httpAgent = new http.Agent({ keepAlive: true, maxSockets: 15 });

const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.static('public'));

// In-memory storage for parsed file data
const uploadedData = { purchase: null, library: null };

// ─── Column name normalization ───────────────────────────────────────────────

const ISBN13_KEYS = ['isbn13', 'isbn-13', 'isbn 13', '도서번호', 'isbn'];
const TITLE_KEYS = ['서명', '도서명', '제목', '책이름', '서 명', 'title', '표제'];
const PUBLISHER_KEYS = ['출판사', '발행처', '발행사', '출 판 사', 'publisher'];

function normalizeKey(key) {
  return String(key).trim().toLowerCase().replace(/\s+/g, ' ');
}

function findColumn(obj, candidates) {
  const keys = Object.keys(obj);
  for (const candidate of candidates) {
    const match = keys.find(k => normalizeKey(k) === candidate);
    if (match) return match;
  }
  // Partial match fallback
  for (const candidate of candidates) {
    const match = keys.find(k => normalizeKey(k).includes(candidate));
    if (match) return match;
  }
  return null;
}

function extractFields(row) {
  const isbn13Key = findColumn(row, ISBN13_KEYS);
  const titleKey = findColumn(row, TITLE_KEYS);
  const publisherKey = findColumn(row, PUBLISHER_KEYS);

  let isbn13 = isbn13Key ? String(row[isbn13Key]).trim() : '';
  // Normalize ISBN13: remove hyphens, keep only digits
  isbn13 = isbn13.replace(/[-\s]/g, '');
  if (isbn13.length !== 13 || !/^\d+$/.test(isbn13)) isbn13 = '';

  return {
    isbn13,
    title: titleKey ? String(row[titleKey]).trim() : '',
    publisher: publisherKey ? String(row[publisherKey]).trim() : '',
    _raw: row,
    _isbn13Key: isbn13Key,
    _titleKey: titleKey,
    _publisherKey: publisherKey,
  };
}

// ─── Excel parsing ───────────────────────────────────────────────────────────

function parseExcel(filePath) {
  const workbook = XLSX.readFile(filePath, { type: 'file', cellDates: true, raw: false });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // sheet_to_json with header:1 gives array-of-arrays; with header:"A" gives key-based
  // Use defval:'' to fill blanks, raw:false for string values
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });

  if (rawRows.length === 0) return { rows: [], headers: [] };

  const headers = Object.keys(rawRows[0]);
  const rows = rawRows.filter(row =>
    Object.values(row).some(v => String(v).trim() !== '')
  );

  return { rows, headers };
}

// ─── Aladin API ──────────────────────────────────────────────────────────────

async function fetchIsbn13ByTitle(title, apiKey) {
  try {
    const url = 'http://www.aladin.co.kr/ttb/api/ItemSearch.aspx';
    const res = await axios.get(url, {
      params: {
        ttbkey: apiKey,
        Query: title,
        QueryType: 'Title',
        MaxResults: 1,
        start: 1,
        SearchTarget: 'Book',
        output: 'js',
        Version: '20131101',
      },
      timeout: 5000,
      httpAgent,
    });
    const data = res.data;
    if (data && data.item && data.item.length > 0) {
      const isbn13 = data.item[0].isbn13 || '';
      return isbn13 ? String(isbn13).replace(/[-\s]/g, '') : '';
    }
  } catch (e) {
    // ignore
  }
  return '';
}

async function enrichWithIsbn13(books, apiKey, progressCallback) {
  const CONCURRENCY = 10;
  let enriched = 0;
  let completed = 0;

  const targets = books
    .map((book, idx) => ({ book, idx }))
    .filter(({ book }) => !book.isbn13 && book.title);

  const total = books.length;

  // Run tasks with a fixed concurrency pool
  async function runPool(tasks, concurrency, worker) {
    let i = 0;
    async function next() {
      if (i >= tasks.length) return;
      const task = tasks[i++];
      await worker(task);
      await next();
    }
    await Promise.all(Array.from({ length: concurrency }, next));
  }

  await runPool(targets, CONCURRENCY, async ({ book, idx }) => {
    const isbn13 = await fetchIsbn13ByTitle(book.title, apiKey);
    if (isbn13) {
      books[idx].isbn13 = isbn13;
      enriched++;
    }
    completed++;
    if (progressCallback) progressCallback(completed, targets.length, enriched);
  });

  if (progressCallback) progressCallback(total, total, enriched);
  return enriched;
}

// ─── Duplicate detection ─────────────────────────────────────────────────────

function normalizeTitle(t) {
  return String(t)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^\w\s가-힣ㄱ-ㅎㅏ-ㅣ]/g, '');
}

function findDuplicates(purchaseBooks, libraryBooks) {
  const results = [];

  // Build lookup maps for library
  const libraryByIsbn = new Map();
  const libraryByTitle = new Map();
  const libraryByTitlePublisher = new Map();

  for (const book of libraryBooks) {
    if (book.isbn13) libraryByIsbn.set(book.isbn13, book);
    const nt = normalizeTitle(book.title);
    if (nt) {
      if (!libraryByTitle.has(nt)) libraryByTitle.set(nt, book);
      const key = `${nt}||${normalizeTitle(book.publisher)}`;
      if (!libraryByTitlePublisher.has(key)) libraryByTitlePublisher.set(key, book);
    }
  }

  for (const pBook of purchaseBooks) {
    let matchedLibBook = null;
    let matchMethod = '';

    // 1st priority: ISBN13
    if (pBook.isbn13 && libraryByIsbn.has(pBook.isbn13)) {
      matchedLibBook = libraryByIsbn.get(pBook.isbn13);
      matchMethod = 'ISBN13';
    }

    // 2nd priority: title
    if (!matchedLibBook) {
      const nt = normalizeTitle(pBook.title);
      if (nt && libraryByTitle.has(nt)) {
        matchedLibBook = libraryByTitle.get(nt);
        matchMethod = '서명';
      }
    }

    // 3rd priority: title + publisher
    if (!matchedLibBook) {
      const key = `${normalizeTitle(pBook.title)}||${normalizeTitle(pBook.publisher)}`;
      if (libraryByTitlePublisher.has(key)) {
        matchedLibBook = libraryByTitlePublisher.get(key);
        matchMethod = '서명+출판사';
      }
    }

    if (matchedLibBook) {
      results.push({
        purchase: pBook,
        library: matchedLibBook,
        matchMethod,
      });
    }
  }

  return results;
}

// ─── Routes ──────────────────────────────────────────────────────────────────

app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    const fileType = req.body.type; // 'purchase' | 'library'
    if (!['purchase', 'library'].includes(fileType)) {
      return res.status(400).json({ error: '잘못된 파일 유형입니다.' });
    }

    const { rows, headers } = parseExcel(req.file.path);
    const books = rows.map(extractFields);

    uploadedData[fileType] = { books, headers, originalRows: rows };

    // Detect column names for feedback
    const sample = books[0] || {};
    const detectedColumns = {
      isbn13: sample._isbn13Key || null,
      title: sample._titleKey || null,
      publisher: sample._publisherKey || null,
    };

    const missingIsbn13Count = books.filter(b => !b.isbn13).length;

    // Clean up uploaded file
    fs.unlink(req.file.path, () => {});

    res.json({
      count: books.length,
      missingIsbn13Count,
      detectedColumns,
      headers,
    });
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: '파일 파싱 중 오류가 발생했습니다: ' + err.message });
  }
});

// SSE endpoint for comparison with progress
app.post('/api/compare', async (req, res) => {
  const { apiKey } = req.body;

  if (!uploadedData.purchase || !uploadedData.library) {
    return res.status(400).json({ error: '두 파일 모두 업로드해주세요.' });
  }

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const send = (event, data) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  try {
    const purchaseBooks = uploadedData.purchase.books.map(b => ({ ...b }));
    const libraryBooks = uploadedData.library.books.map(b => ({ ...b }));

    const purchaseMissing = purchaseBooks.filter(b => !b.isbn13).length;
    const libraryMissing = libraryBooks.filter(b => !b.isbn13).length;

    // Enrich purchase list if needed
    if (purchaseMissing > 0 && apiKey) {
      send('progress', { step: 'enrich_purchase', message: `구입 예정 목록 ISBN13 조회 중... (${purchaseMissing}건, 동시 10건)`, current: 0, total: purchaseMissing });
      await enrichWithIsbn13(purchaseBooks, apiKey, (current, total, enriched) => {
        send('progress', { step: 'enrich_purchase', message: `구입 예정 목록 ISBN13 조회 중...`, current, total, enriched });
      });
    }

    // Enrich library list if needed
    if (libraryMissing > 0 && apiKey) {
      send('progress', { step: 'enrich_library', message: `소장 목록 ISBN13 조회 중... (${libraryMissing}건, 동시 10건)`, current: 0, total: libraryMissing });
      await enrichWithIsbn13(libraryBooks, apiKey, (current, total, enriched) => {
        send('progress', { step: 'enrich_library', message: `소장 목록 ISBN13 조회 중...`, current, total, enriched });
      });
    }

    send('progress', { step: 'comparing', message: '중복 검사 중...' });

    const duplicates = findDuplicates(purchaseBooks, libraryBooks);

    // Serialize only display fields
    const serializeBook = (b) => ({
      isbn13: b.isbn13,
      title: b.title,
      publisher: b.publisher,
    });

    send('result', {
      duplicates: duplicates.map(d => ({
        purchase: serializeBook(d.purchase),
        library: serializeBook(d.library),
        matchMethod: d.matchMethod,
      })),
      purchaseCount: purchaseBooks.length,
      libraryCount: libraryBooks.length,
      purchaseEnriched: purchaseMissing,
      libraryEnriched: libraryMissing,
    });
  } catch (err) {
    console.error('Compare error:', err);
    send('error', { message: '비교 중 오류가 발생했습니다: ' + err.message });
  }

  res.end();
});

// Download enriched Excel
app.post('/api/download-enriched', async (req, res) => {
  const { type } = req.body;
  if (!uploadedData[type]) {
    return res.status(400).json({ error: '파일이 업로드되지 않았습니다.' });
  }

  const { books, headers } = uploadedData[type];
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');

  // Add headers
  sheet.addRow(headers);

  // Add rows with updated ISBN13
  for (const book of books) {
    const row = { ...book._raw };
    if (book._isbn13Key && book.isbn13) {
      row[book._isbn13Key] = book.isbn13;
    }
    sheet.addRow(headers.map(h => row[h] ?? ''));
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="enriched_${type}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Book Checker running at http://localhost:${PORT}`);
});
