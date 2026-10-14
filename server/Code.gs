/* ============================================================
   숲속 박자 음악회 — 기록 저장 서버 (Google Apps Script)
   · 로그인 확인은 생각 오락실 통합 계정 서버에 물어봐요.
   · 기록은 이 스프레드시트의 '기록' 시트에 닉네임마다 한 줄씩 저장해요.
   설치: 스프레드시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여 넣고
         배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
   ============================================================ */
var HUB_API = 'https://script.google.com/macros/s/AKfycbzZlmUD9M4rfo4Y4eEYgNjZH0N8SppUA1Otuo9yMEOqcLfklw8XbS5ZmXj1D3KZY4EG/exec';
var SHEET = '기록';
var HEAD = ['닉네임', '저장 시각', '별', '통과 리듬', '코인', '모은 친구', '무한 마디', '무한 점수', '데이터'];
var MAX_DATA = 45000;

function doGet() { return out({ ok: true, msg: '숲속 박자 음악회 서버가 켜져 있어요.' }); }

function doPost(e) {
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'rank') return out(rank());
    var who = verify(b.nick, b.pin);
    if (!who.ok) return out(who);
    if (b.action === 'load') return out(load(who.nick));
    if (b.action === 'save') return out(store(who.nick, b.data));
    return out({ ok: false, error: '알 수 없는 요청이에요.' });
  } catch (err) {
    return out({ ok: false, error: '서버 오류: ' + err, retry: true });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 통합 계정 확인 (6시간 동안 기억) ---------- */
function verify(nick, pin) {
  nick = String(nick || '').trim(); pin = String(pin || '').trim();
  if (!nick || !/^\d{4}$/.test(pin)) return { ok: false, code: 'PIN', error: '로그인 정보가 없어요.' };
  var cache = CacheService.getScriptCache();
  var key = 'v:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, nick + '|' + pin));
  var hit = cache.get(key);
  if (hit) return { ok: true, nick: hit };
  var r = UrlFetchApp.fetch(HUB_API, {
    method: 'post', contentType: 'text/plain;charset=utf-8', muteHttpExceptions: true, followRedirects: true,
    payload: JSON.stringify({ action: 'me', nick: nick, pin: pin })
  });
  var j = {};
  try { j = JSON.parse(r.getContentText()); } catch (e) { return { ok: false, error: '생각 오락실 서버에 연결하지 못했어요.', retry: true }; }
  if (!j.ok) return { ok: false, code: j.code || 'PIN', error: j.error || '생각 오락실에서 다시 로그인해 주세요.' };
  cache.put(key, j.nick || nick, 21600);
  return { ok: true, nick: j.nick || nick };
}

/* ---------- 시트 ---------- */
function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function findRow(sh, nick) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var names = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (var i = 0; i < names.length; i++) if (names[i][0] === nick) return i + 2;
  return 0;
}

function load(nick) {
  var sh = sheet(), row = findRow(sh, nick);
  if (!row) return { ok: true, nick: nick, data: null };
  var raw = sh.getRange(row, 9).getValue();
  var data = null;
  try { data = JSON.parse(raw); } catch (e) {}
  return { ok: true, nick: nick, data: data };
}

function store(nick, data) {
  if (!data || typeof data !== 'object') return { ok: false, error: '저장할 기록이 없어요.' };
  var raw = JSON.stringify(data);
  if (raw.length > MAX_DATA) return { ok: false, error: '기록이 너무 커요.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet(), row = findRow(sh, nick);
    var stars = 0, pass = 0;
    (data.best || []).forEach(function (s) { (s || []).forEach(function (v) { stars += +v || 0; }); });
    (data.pass || []).forEach(function (s) { (s || []).forEach(function (v) { if (v) pass++; }); });
    var dex = data.dex || {}, kinds = 0;
    for (var k in dex) if (dex[k] > 0) kinds++;
    var inf = data.infBest || {};
    var vals = [[nick, new Date(), stars, pass, +data.coins || 0, kinds, +inf.bars || 0, +inf.score || 0, raw]];
    if (!row) row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, HEAD.length).setValues(vals);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- 무한 음표 순위 (나중에 순위표에 씀) ---------- */
function rank() {
  var sh = sheet(), n = sh.getLastRow();
  if (n < 2) return { ok: true, list: [] };
  var v = sh.getRange(2, 1, n - 1, 8).getValues();
  var list = v.filter(function (r) { return r[6] > 0; })
    .map(function (r) { return { nick: String(r[0]), bars: r[6], score: r[7] }; })
    .sort(function (a, b) { return b.bars - a.bars || b.score - a.score; })
    .slice(0, 30);
  return { ok: true, list: list };
}
