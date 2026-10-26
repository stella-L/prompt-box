// 프롬프트 상자 — 구글 시트 API
// 시트의 [확장 프로그램 → Apps Script]에 붙여넣고 웹 앱으로 배포하세요. (README 참고)

const TOKEN = '여기에-비밀-토큰을-넣으세요';

// 종류별 시트 탭. main 칸이 비어 있는 줄은 빈 줄로 본다
const KINDS = {
  prompts: {
    cols: ['id', 'title', 'body', 'source', 'tags', 'memo', 'fav', 'created', 'updated'],
    headers: ['id', '제목', '프롬프트', '출처', '태그', '메모', '즐겨찾기', '만든 날', '수정한 날'],
    main: 'body',
  },
  links: {
    cols: ['id', 'title', 'url', 'tags', 'memo', 'fav', 'created', 'updated'],
    headers: ['id', '제목', '링크', '태그', '메모', '즐겨찾기', '만든 날', '수정한 날'],
    main: 'url',
  },
};

// 브라우저로 URL을 열면 권한 승인 후 이 화면이 보인다
function doGet() {
  Object.keys(KINDS).forEach(getSheet);
  return ContentService.createTextOutput('✅ 프롬프트 상자 연결 준비 완료! 이 창은 닫아도 돼요.');
}

function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.token !== TOKEN) throw new Error('토큰이 맞지 않아요');
    const kind = req.kind || 'prompts';
    if (!KINDS[kind]) throw new Error('알 수 없는 종류');
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (req.action === 'all') res = { prompts: readAll('prompts'), links: readAll('links') };
      else if (req.action === 'list') res = { items: readAll(kind) };
      else if (req.action === 'save') { upsert(kind, req.item); res = {}; }
      else if (req.action === 'delete') { remove(kind, req.id); res = {}; }
      else if (req.action === 'bulk') { appendMissing(kind, req.items || []); res = {}; }
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

function getSheet(kind) {
  const { cols, headers, main } = KINDS[kind];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(kind);
  if (!sh) sh = ss.insertSheet(kind);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@'); // '='로 시작해도 수식이 되지 않게
    sh.setColumnWidth(cols.indexOf(main) + 1, 400);
    sh.hideColumns(1);
  }
  return sh;
}

function toRow(kind, it) {
  return KINDS[kind].cols.map(c => {
    if (c === 'tags') return (it.tags || []).join(', ');
    if (c === 'fav') return it.fav ? '★' : '';
    if (c === 'created' || c === 'updated') return fmtDate(it[c]);
    return it[c] || '';
  });
}

function fromRow(kind, r) {
  const it = {};
  KINDS[kind].cols.forEach((c, i) => {
    if (c === 'tags') it.tags = String(r[i]).split(',').map(s => s.trim().replace(/^#/, '')).filter(Boolean);
    else if (c === 'fav') it.fav = String(r[i]).trim() !== '';
    else if (c === 'created' || c === 'updated') it[c] = parseDate(r[i]);
    else it[c] = String(r[i]);
  });
  return it;
}

function fmtDate(ms) {
  return Utilities.formatDate(new Date(ms || Date.now()), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function parseDate(v) {
  if (v instanceof Date) return v.getTime();
  const t = Date.parse(String(v).replace(' ', 'T'));
  return isNaN(t) ? 0 : t;
}

function readAll(kind) {
  const sh = getSheet(kind);
  const { cols, main } = KINDS[kind];
  const m = cols.indexOf(main), cr = cols.indexOf('created'), up = cols.indexOf('updated');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const range = sh.getRange(2, 1, last - 1, cols.length);
  const rows = range.getValues();
  let fixed = false;
  // 시트에서 직접 추가한 줄은 id가 비어 있으니 채워준다
  rows.forEach(r => {
    if (!r[0] && String(r[m]).trim()) {
      r[0] = Utilities.getUuid();
      if (!r[cr]) r[cr] = fmtDate();
      if (!r[up]) r[up] = r[cr];
      fixed = true;
    }
  });
  if (fixed) range.setValues(rows);
  return rows.filter(r => r[0] && String(r[m]).trim()).map(r => fromRow(kind, r));
}

function findRow(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

// 서식 있는 텍스트로 쓰면 '=', '-' 등으로 시작해도 수식으로 바뀌지 않는다
function writeRows(sh, row, values) {
  const rich = values.map(r => r.map(v => SpreadsheetApp.newRichTextValue().setText(String(v)).build()));
  sh.getRange(row, 1, values.length, values[0].length).setNumberFormat('@').setRichTextValues(rich);
}

function upsert(kind, item) {
  if (!item || !item.id) throw new Error('잘못된 데이터');
  const sh = getSheet(kind);
  const row = findRow(sh, item.id);
  writeRows(sh, row > 0 ? row : sh.getLastRow() + 1, [toRow(kind, item)]);
}

function remove(kind, id) {
  const sh = getSheet(kind);
  const row = findRow(sh, id);
  if (row > 0) sh.deleteRow(row);
}

function appendMissing(kind, items) {
  const sh = getSheet(kind);
  const last = sh.getLastRow();
  const ids = new Set(last < 2 ? [] : sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])));
  const rows = items.filter(it => it && it.id && !ids.has(it.id)).map(it => toRow(kind, it));
  if (rows.length) writeRows(sh, last + 1, rows);
}
