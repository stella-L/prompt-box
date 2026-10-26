// 프롬프트 상자 — 구글 시트 API
// 시트의 [확장 프로그램 → Apps Script]에 붙여넣고 웹 앱으로 배포하세요. (README 참고)

const TOKEN = '여기에-비밀-토큰을-넣으세요';
const SHEET = 'prompts';
const COLS = ['id', 'title', 'body', 'source', 'tags', 'memo', 'fav', 'created', 'updated'];
const HEADERS = ['id', '제목', '프롬프트', '출처', '태그', '메모', '즐겨찾기', '만든 날', '수정한 날'];

function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.token !== TOKEN) throw new Error('토큰이 맞지 않아요');
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = getSheet();
      if (req.action === 'list') res = { items: readAll(sh) };
      else if (req.action === 'save') { upsert(sh, req.item); res = {}; }
      else if (req.action === 'delete') { remove(sh, req.id); res = {}; }
      else if (req.action === 'bulk') { appendMissing(sh, req.items || []); res = {}; }
      else throw new Error('알 수 없는 요청');
    } finally {
      lock.releaseLock();
    }
    res.ok = true;
  } catch (err) {
    res = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) sh = ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, COLS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), COLS.length).setNumberFormat('@'); // '='로 시작해도 수식이 되지 않게
    sh.setColumnWidth(3, 400);
    sh.hideColumns(1);
  }
  return sh;
}

function toRow(it) {
  return [
    it.id,
    it.title || '',
    it.body || '',
    it.source || '',
    (it.tags || []).join(', '),
    it.memo || '',
    it.fav ? '★' : '',
    fmtDate(it.created),
    fmtDate(it.updated),
  ];
}

function fromRow(r) {
  return {
    id: String(r[0]),
    title: String(r[1]),
    body: String(r[2]),
    source: String(r[3]),
    tags: String(r[4]).split(',').map(s => s.trim().replace(/^#/, '')).filter(Boolean),
    memo: String(r[5]),
    fav: String(r[6]).trim() !== '',
    created: parseDate(r[7]),
    updated: parseDate(r[8]),
  };
}

function fmtDate(ms) {
  return Utilities.formatDate(new Date(ms || Date.now()), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function parseDate(v) {
  if (v instanceof Date) return v.getTime();
  const t = Date.parse(String(v).replace(' ', 'T'));
  return isNaN(t) ? 0 : t;
}

function readAll(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const range = sh.getRange(2, 1, last - 1, COLS.length);
  const rows = range.getValues();
  let fixed = false;
  // 시트에서 직접 추가한 줄은 id가 비어 있으니 채워준다
  rows.forEach(r => {
    if (!r[0] && String(r[2]).trim()) {
      r[0] = Utilities.getUuid();
      if (!r[7]) r[7] = fmtDate();
      if (!r[8]) r[8] = r[7];
      fixed = true;
    }
  });
  if (fixed) range.setValues(rows);
  return rows.filter(r => r[0] && String(r[2]).trim()).map(fromRow);
}

function findRow(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function writeRows(sh, row, values) {
  sh.getRange(row, 1, values.length, COLS.length).setNumberFormat('@').setValues(values);
}

function upsert(sh, item) {
  if (!item || !item.id) throw new Error('잘못된 데이터');
  const row = findRow(sh, item.id);
  writeRows(sh, row > 0 ? row : sh.getLastRow() + 1, [toRow(item)]);
}

function remove(sh, id) {
  const row = findRow(sh, id);
  if (row > 0) sh.deleteRow(row);
}

function appendMissing(sh, items) {
  const last = sh.getLastRow();
  const ids = new Set(last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])));
  const rows = items.filter(it => it && it.id && !ids.has(it.id)).map(toRow);
  if (rows.length) writeRows(sh, last + 1, rows);
}
