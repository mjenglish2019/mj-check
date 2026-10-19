/**
 * 엠제이 점검실 서버 (구글 시트 + Apps Script) — 2026-10-14
 *
 * 쓰는 법
 *  1) 이 코드를 시트의 [확장 프로그램 → Apps Script]에 통째로 붙여 넣고 저장
 *  2) 위쪽 함수 목록에서 '처음설정'을 골라 [실행] (처음 한 번 권한 허용)
 *  3) [배포 → 새 배포 → 웹 앱] 실행: 나 / 액세스: 모든 사용자 → 배포 → 웹 앱 주소 복사
 *
 * 시트
 *  - 학생 : 아이디 · 비밀번호 · 이름(학교+학년) · 메모   ← 원장님이 관리
 *  - 결과 : 시험이 끝날 때마다 한 줄씩 자동으로 쌓임
 *  - 오늘 : 오늘 결과만 모아 보기(자동)
 *  - 알림장 : 학생이 앱에서 쓴 알림장이 날짜별로 한 줄씩 쌓임(자동, 고치지 마세요)
 *      상태 = 작성 중 / 확정 / 자동 잠금. 확정하거나 날짜가 지나면(밤 12시) 서버가 수정을 거부함
 *  - 상태 : 학생별 기록 저장용(숨김, 고치지 마세요)
 */

var SHEET_STU = "학생", SHEET_RES = "결과", SHEET_TODAY = "오늘", SHEET_STATE = "상태";
var RES_HEAD = ["시간", "아이디", "이름", "구분", "교재", "단원", "점수", "맞은 수", "결과", "재시험", "틀린 것", "화면 이탈", "조교 확인 영작"];
var CHUNK = 40000;          // 셀 하나에 넣는 글자 수(시트 한도 5만 자보다 작게)
var MAX_FAIL = 10;          // 비밀번호 10번 틀리면 10분 잠금

/* 알림장 (2026-10-19) */
var SHEET_DIARY = "알림장";
var DIARY_HEAD = ["날짜", "아이디", "이름", "단어", "듣기", "독해", "문법", "구문해석", "영작", "기타", "상태", "확정 시각", "마지막 수정"];
var DIARY_KEYS = ["vocab", "listen", "read", "gram", "syn", "write", "etc"];   // 시트 D~J열과 같은 순서 (앱 DIARY_F와 같게)
var DIARY_MAX = 300;        // 칸 하나에 쓸 수 있는 글자 수
var DIARY_SCAN = 6000;      // 시트 맨 아래에서 몇 줄까지 찾아볼지
var DIARY_SHOW = 30;        // 앱에 돌려주는 최근 알림장 개수
var TZ = "Asia/Seoul";
var ST_DRAFT = "작성 중", ST_DONE = "확정", ST_AUTO = "자동 잠금";

/* ---------- 처음 한 번 실행 ---------- */
function 처음설정() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(SHEET_STU);
  if (!s) {
    s = ss.insertSheet(SHEET_STU, 0);
    s.getRange("A:B").setNumberFormat("@");
    s.appendRow(["아이디", "비밀번호", "이름(학교+학년)", "메모"]);
    s.appendRow(["홍길동", "1234", "홍길동(배명고2)", "예시예요. 지우고 학생을 넣어 주세요"]);
    s.setFrozenRows(1);
  }
  var r = ss.getSheetByName(SHEET_RES);
  if (!r) {
    r = ss.insertSheet(SHEET_RES);
    r.appendRow(RES_HEAD);
    r.setFrozenRows(1);
    r.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm");
  }
  var t = ss.getSheetByName(SHEET_TODAY);
  if (!t) {
    t = ss.insertSheet(SHEET_TODAY);
    t.getRange("A1").setFormula(
      "=IFERROR(QUERY('" + SHEET_RES + "'!A:M,\"select * where A >= date '\"&TEXT(TODAY(),\"yyyy-mm-dd\")&\"' order by A desc\",1),\"오늘 결과가 아직 없어요\")");
    t.getRange("A:A").setNumberFormat("hh:mm");
  }
  var st = ss.getSheetByName(SHEET_STATE);
  if (!st) {
    st = ss.insertSheet(SHEET_STATE);
    st.appendRow(["아이디", "저장 시각", "기록(고치지 마세요)"]);
    st.hideSheet();
  }
  var dy = ss.getSheetByName(SHEET_DIARY);
  if (!dy) {
    dy = ss.insertSheet(SHEET_DIARY);
    dy.getRange("A:M").setNumberFormat("@");   // 전부 글자로 저장(=로 시작하는 글이 수식이 되지 않게)
    dy.appendRow(DIARY_HEAD);
    dy.setFrozenRows(1);
  }
  secret_();
}

/* ---------- 웹 앱 입구 ---------- */
function doGet() {
  return json_({ ok: true, app: "mj-check" });
}
function doPost(e) {
  var out;
  try {
    var q = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    out = handle_(q);
  } catch (err) {
    out = { ok: false, code: "server", msg: String(err && err.message || err) };
  }
  return json_(out);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(q) {
  if (q.act === "login") {
    var cache = CacheService.getScriptCache(), fk = "fail:" + key_(q.id);
    var fails = Number(cache.get(fk) || 0);
    if (fails >= MAX_FAIL) return { ok: false, code: "locked" };
    var stu = findStudent_(q.id);
    if (!stu || !pwOk_(stu, q.pw)) {
      cache.put(fk, String(fails + 1), 600);
      return { ok: false, code: "badlogin" };
    }
    cache.remove(fk);
    return { ok: true, id: stu.id, name: stu.name, token: token_(stu), state: loadState_(stu.id) };
  }
  var me = findStudent_(q.id);
  if (!me || q.token !== token_(me)) return { ok: false, code: "auth" };
  if (q.act === "push") {
    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      var cur = loadState_(me.id) || {};
      var merged = mergeState(cur, q.state || {});
      var have = {}, rids = (cur.rids || []).slice(), rows = [];
      rids.forEach(function (x) { have[x] = 1; });
      (q.results || []).forEach(function (r) {
        if (r && r.at && !have[r.at]) { have[r.at] = 1; rids.push(r.at); rows.push(resRow_(me, r)); }
      });
      merged.rids = rids.slice(-400);
      if (rows.length) {
        var sh = sheet_(SHEET_RES);
        sh.getRange(sh.getLastRow() + 1, 1, rows.length, RES_HEAD.length).setValues(rows);
      }
      saveState_(me.id, merged);
      return { ok: true, name: me.name, state: merged, saved: rows.length };
    } finally {
      lock.releaseLock();
    }
  }
  if (q.act === "diary_get" || q.act === "diary_save" || q.act === "diary_confirm") return diary_(me, q);
  return { ok: false, code: "act" };
}

/* ---------- 알림장 (2026-10-19) ----------
   규칙: 그날(한국 시간) 알림장만 쓸 수 있고, 확정을 누르거나 날짜가 바뀌면 서버가 수정·삭제를 거부한다.
   날짜가 지났는데 확정 못 한 알림장은 상태가 '자동 잠금'으로 바뀐다(아무 학생이나 앱을 쓸 때 정리됨). */
function todayKst_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"); }
function nowKst_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm"); }
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd") : String(v == null ? "" : v).trim(); }
function tstr_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm") : String(v == null ? "" : v).trim(); }

function diaryWindow_(sh) {
  var n = sh.getLastRow();
  if (n < 2) return { first: 2, vals: [] };
  var first = Math.max(2, n - DIARY_SCAN + 1);
  var vals = sh.getRange(first, 1, n - first + 1, DIARY_HEAD.length).getValues();
  for (var i = 0; i < vals.length; i++) { vals[i][0] = dstr_(vals[i][0]); vals[i][10] = String(vals[i][10] || "").trim(); }
  return { first: first, vals: vals };
}
function sweepDiary_(sh, win, today) {
  for (var i = 0; i < win.vals.length; i++) {
    var v = win.vals[i];
    if (v[0] && v[0] < today && v[10] === ST_DRAFT) { v[10] = ST_AUTO; sh.getRange(win.first + i, 11).setValue(ST_AUTO); }
  }
}
function diaryFind_(win, id, date) {
  for (var i = win.vals.length - 1; i >= 0; i--) if (win.vals[i][0] === date && key_(win.vals[i][1]) === key_(id)) return i;
  return -1;
}
function entryOf_(v) {
  var f = {};
  for (var k = 0; k < DIARY_KEYS.length; k++) f[DIARY_KEYS[k]] = String(v[3 + k] == null ? "" : v[3 + k]);
  return { d: dstr_(v[0]), f: f, st: String(v[10] || ST_DRAFT), conf: tstr_(v[11]), upd: tstr_(v[12]) };
}
function diaryList_(win, id) {
  var out = [];
  for (var i = win.vals.length - 1; i >= 0 && out.length < DIARY_SHOW; i--)
    if (key_(win.vals[i][1]) === key_(id)) out.push(entryOf_(win.vals[i]));
  out.sort(function (a, b) { return a.d < b.d ? 1 : a.d > b.d ? -1 : 0; });
  return out;
}
function cleanFields_(f) {
  var out = [];
  for (var k = 0; k < DIARY_KEYS.length; k++) out.push(String((f && f[DIARY_KEYS[k]]) || "").replace(/\r/g, "").trim().slice(0, DIARY_MAX));
  return out;
}
function diary_(me, q) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var sh = sheet_(SHEET_DIARY), today = todayKst_();
    var win = diaryWindow_(sh);
    sweepDiary_(sh, win, today);
    if (q.act === "diary_get") return { ok: true, today: today, entries: diaryList_(win, me.id) };
    var date = String(q.date || "");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, code: "date" };
    if (date > today) return { ok: false, code: "future" };
    if (date < today) return { ok: false, code: "locked" };       // 날짜가 지난 알림장은 고칠 수 없음
    var idx = diaryFind_(win, me.id, date);
    if (idx >= 0 && win.vals[idx][10] !== ST_DRAFT) return { ok: false, code: "locked" };   // 확정·자동 잠금
    var f = cleanFields_(q.f), confirm = q.act === "diary_confirm";
    if (confirm && !f.some(function (x) { return x !== ""; })) return { ok: false, code: "empty" };
    var now = nowKst_();
    var row = [date, me.id, me.name].concat(f, [confirm ? ST_DONE : ST_DRAFT, confirm ? now : "", now]);
    var r = idx >= 0 ? win.first + idx : sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, DIARY_HEAD.length).setValues([row]);
    return { ok: true, entry: entryOf_(row) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 학생 ---------- */
function key_(v) { return String(v == null ? "" : v).replace(/\s+/g, "").toLowerCase(); }
function findStudent_(id) {
  var k = key_(id);
  if (!k) return null;
  var v = sheet_(SHEET_STU).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (key_(v[i][0]) === k && String(v[i][1]).trim() !== "") {
      var sid = String(v[i][0]).trim();
      return { id: sid, pw: v[i][1], name: String(v[i][2] || "").trim() || sid };
    }
  }
  return null;
}
function pwOk_(stu, pw) {
  var typed = String(pw == null ? "" : pw).trim();
  if (typeof stu.pw === "number") return typed !== "" && Number(typed) === stu.pw;
  return String(stu.pw).trim() === typed;
}
function secret_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty("SECRET");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty("SECRET", s); }
  return s;
}
/* 비밀번호를 바꾸면 토큰도 바뀌어 그 학생은 다시 로그인해야 함 */
function token_(stu) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(stu.id + "|" + String(stu.pw).trim(), secret_()));
}

/* ---------- 결과 한 줄 ---------- */
function txt_(s) {
  s = String(s == null ? "" : s).slice(0, 3000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function sec_(n) { n = Number(n) || 0; return n >= 60 ? Math.round(n / 60) + "분" : n + "초"; }
function resRow_(me, r) {
  var T = { vocab: "단어", grammar: "문법", reading: "독해", naesin: "내신" };
  var st = r.pass ? "통과" : r.final ? "조교 확인" : "재시험";
  var wrong = (r.wrong && r.wrong.length) ? r.wrong.join(", ") : (r.wrongN ? "틀린 문제 " + r.wrongN + "개" : "");
  var d = new Date(r.at);
  return [isNaN(d) ? new Date() : d, txt_(me.id), txt_(me.name), T[r.type] || txt_(r.type), txt_(r.bookName), txt_(r.unitName),
    Number(r.pct) || 0, txt_((r.correct || 0) + "/" + (r.total || 0)), st, r.retryNo ? r.retryNo + "차" : "",
    txt_(wrong), r.away ? r.away + "회 (" + sec_(r.awaySec) + ")" : "", txt_((r.pend || []).join(" / "))];
}

/* ---------- 학생별 기록 저장 (상태 시트, 숨김) ---------- */
function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(name);
  if (!s) { 처음설정(); s = ss.getSheetByName(name); }
  return s;
}
function findRow_(sh, id) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === "id:" + id) return i + 2;
  return 0;
}
function loadState_(id) {
  var sh = sheet_(SHEET_STATE), row = findRow_(sh, id);
  if (!row) return null;
  var w = sh.getLastColumn();
  if (w < 3) return null;
  var parts = sh.getRange(row, 3, 1, w - 2).getValues()[0];
  var s = parts.map(function (p) { p = String(p || ""); return p.charAt(0) === "~" ? p.slice(1) : p; }).join("");
  try { return s ? JSON.parse(s) : null; } catch (e) { return null; }
}
function saveState_(id, st) {
  var sh = sheet_(SHEET_STATE), s = JSON.stringify(st), parts = [];
  for (var i = 0; i < s.length; i += CHUNK) parts.push("~" + s.slice(i, i + CHUNK));
  var row = findRow_(sh, id) || sh.getLastRow() + 1;
  var width = Math.max(sh.getLastColumn(), 2 + parts.length);
  var vals = ["id:" + id, new Date()].concat(parts);
  while (vals.length < width) vals.push("");
  sh.getRange(row, 1, 1, width).setValues([vals]);
}

/* ---------- 기록 합치기 (앱 index.html의 mergeState와 똑같이 유지) ---------- */
function mergeState(a, b) {
  a = a || {}; b = b || {};
  var day = [a.day || "", b.day || ""].sort().pop();
  var A = a.day === day ? a : {}, B = b.day === day ? b : {};
  var rank = function (l) { return !l ? -1 : (l.state === "pass" || l.state === "final") ? 1e6 : (l.retryNo || 0) * 2 + (l.state === "open" ? 1 : 0); };
  var lock = {}, k;
  for (k in (A.lock || {})) lock[k] = A.lock[k];
  for (k in (B.lock || {})) {
    var x = lock[k], y = B.lock[k], rx = rank(x), ry = rank(y);
    if (ry > rx) lock[k] = y;
    else if (ry === rx && x) {
      var m = {}, j; for (j in y) m[j] = y[j];
      m.away = Math.max(x.away || 0, y.away || 0); m.awaySec = Math.max(x.awaySec || 0, y.awaySec || 0);
      m.quitAt = Math.max(x.quitAt || 0, y.quitAt || 0); m.qDone = Math.max(x.qDone || 0, y.qDone || 0);
      lock[k] = m;
    }
  }
  var seen = {}, log = [];
  (A.log || []).concat(B.log || []).forEach(function (e) { if (e && e.at && !seen[e.at]) { seen[e.at] = 1; log.push(e); } });
  log.sort(function (p, q) { return p.at < q.at ? -1 : p.at > q.at ? 1 : 0; });
  var hp = function (h) { return !h ? -1 : h.pass ? 3 : h.final ? 2 : 1; };
  var hist = {};
  for (k in (a.hist || {})) hist[k] = a.hist[k];
  for (k in (b.hist || {})) {
    var hx = hist[k], hy = b.hist[k];
    if (!hx || hp(hy) > hp(hx) || (hp(hy) === hp(hx) && (hy.at || "") >= (hx.at || ""))) hist[k] = hy;
  }
  var wb = (b.wrongAt || 0) >= (a.wrongAt || 0) ? b : a;
  var sb = (b.selAt || 0) >= (a.selAt || 0) ? b : a;
  var rids = (a.rids || []).concat(b.rids || []).filter(function (x, i, arr) { return arr.indexOf(x) === i; }).slice(-400);
  return {
    v: 1, day: day, hist: hist, wrong: wb.wrong || [], wrongAt: wb.wrongAt || 0, sel: sb.sel || {}, selAt: sb.selAt || 0,
    lock: lock, log: log, sent: [A.sent || "", B.sent || ""].sort().pop() || null, rids: rids,
    updated: Math.max(a.updated || 0, b.updated || 0)
  };
}
