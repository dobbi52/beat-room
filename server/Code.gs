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
    if (row) {   // 선생님이 정리한 기록보다 옛날 기록이면 받지 않고 서버 기록을 돌려줘요
      var old = null;
      try { old = JSON.parse(sh.getRange(row, 9).getValue()); } catch (e) {}
      if (old && (old.rev || 0) > (data.rev || 0)) return { ok: false, code: 'REV', data: old };
    }
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

/* ============================================================
   한 번만 실행: 코인 정리 (2026-10-15)
   Apps Script 편집기 위쪽에서 함수 '코인정리'를 고르고 ▶ 실행을 눌러요.
   · 바꾸기 전 '기록' 시트를 '백업 …' 시트로 복사해 둬요.
   · 모든 학생 코인 → 0
   · 새 코인 규칙으로 정상적으로 뽑을 수 있는 횟수보다 많이 뽑은 학생은 뽑은 횟수의 반을 지워요.
     (겹친 것 먼저, 보통 등급 먼저 지우고 아주 특별 친구는 최대한 남겨요)
   · rev를 올려서, 학생 기기에 남은 옛 기록이 다시 덮어쓰지 못하게 해요.
   ============================================================ */
var TIER_OF = {"1":0,"2":0,"3":0,"6":0,"7":0,"8":0,"11":0,"12":0,"13":0,"16":0,"17":0,"18":0,"21":0,"22":0,"23":0,"26":0,"27":0,"28":0,"29":0,"31":0,"32":0,"33":0,"34":0,"36":0,"37":0,"38":0,"39":0,"41":0,"42":0,"43":0,"44":0,"46":0,"47":0,"48":0,"49":0,"5":2,"10":2,"15":2,"20":2,"25":2,"4":1,"9":1,"14":1,"19":1,"24":1,"30":1,"35":1,"40":1,"45":1,"50":1};

function fairDraws(d) {
  var c = 50;
  (d.pass || []).forEach(function (s, i) { (s || []).forEach(function (v) { if (v) c += 30 + (i + 1) * 10; }); });
  (d.best || []).forEach(function (s) { (s || []).forEach(function (v) { c += 10 * (+v || 0); }); });
  return Math.floor(c / 50);
}

function trimDex(d) {
  var dex = d.dex || {}, draws = 0, list = [];
  for (var id in dex) {
    var n = +dex[id] || 0; draws += n;
    for (var k = n; k >= 1; k--) list.push({ id: id, k: k, tier: TIER_OF[id] || 0 });
  }
  if (draws <= fairDraws(d)) return 0;
  var cut = Math.floor(draws / 2);
  list.sort(function (a, b) {               // 지울 순서: 겹친 것 → 낮은 등급 → 나중에 뽑힌 것
    var da = a.k > 1 ? 0 : 1, db = b.k > 1 ? 0 : 1;
    return da - db || a.tier - b.tier || b.k - a.k;
  });
  for (var i = 0; i < cut; i++) {
    var it = list[i];
    dex[it.id] = (dex[it.id] || 0) - 1;
    if (dex[it.id] <= 0) delete dex[it.id];
  }
  d.dex = dex;
  var owned = Object.keys(dex);
  var want = (d.stage || []).length;
  var st = (d.stage || []).filter(function (id) { return dex[id] > 0; });
  owned.sort(function (a, b) { return (TIER_OF[b] || 0) - (TIER_OF[a] || 0); })
    .forEach(function (id) { if (st.length < want && st.indexOf(+id) < 0 && st.indexOf(id) < 0) st.push(+id); });
  d.stage = st;
  return cut;
}

function 코인정리() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('cleanup-2026-10-15')) { Logger.log('이미 정리했어요. 다시 하지 않아요.'); return; }
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = sheet();
    sh.copyTo(ss).setName('백업 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd HHmm'));
    var n = sh.getLastRow();
    if (n < 2) return;
    var rng = sh.getRange(2, 1, n - 1, HEAD.length), v = rng.getValues(), now = Date.now(), trimmed = 0;
    for (var r = 0; r < v.length; r++) {
      var d = null;
      try { d = JSON.parse(v[r][8]); } catch (e) {}
      if (!d) continue;
      if (trimDex(d) > 0) trimmed++;
      d.coins = 0; d.rev = (d.rev || 0) + 1; d.t = now;
      var kinds = 0; for (var k in d.dex) if (d.dex[k] > 0) kinds++;
      v[r][4] = 0; v[r][5] = kinds; v[r][8] = JSON.stringify(d);
    }
    rng.setValues(v);
    props.setProperty('cleanup-2026-10-15', String(now));
    Logger.log('정리 끝: ' + v.length + '명 코인 0, 동물 반 지운 학생 ' + trimmed + '명');
  } finally { lock.releaseLock(); }
}
