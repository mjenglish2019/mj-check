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
 *  - 모의고사 : 주간 모의고사 정답(한 시험 한 줄). 앱 '모의고사 OMR → 선생님: 정답 등록'에서 넣거나 여기에 직접 써도 됨
 *      정답 칸은 35241 53214 … 처럼 1~5 숫자만(띄어쓰기 자유), 3점 문항 칸은 21, 29, 31 …
 *  - 모의고사 답안 : 학생이 앱 OMR로 낸 답안(한 학생·한 시험 한 줄). 줄을 지우면 그 학생이 다시 낼 수 있음
 */

var SHEET_STU = "학생", SHEET_RES = "결과", SHEET_TODAY = "오늘", SHEET_STATE = "상태";
var RES_HEAD = ["시간", "아이디", "이름", "구분", "교재", "단원", "점수", "맞은 수", "결과", "재시험", "틀린 것", "화면 이탈", "조교 확인 영작"];
var CHUNK = 40000;          // 셀 하나에 넣는 글자 수(시트 한도 5만 자보다 작게)
var MAX_FAIL = 10;          // 비밀번호 10번 틀리면 10분 잠금

/* 알림장 (2026-10-19) */
var SHEET_DIARY = "알림장";
var DIARY_HEAD = ["날짜", "아이디", "이름", "단어", "듣기", "독해", "문법", "구문해석", "영작", "기타", "상태", "확정 시각", "마지막 수정", "완료 확인", "확인 시각", "미완료"];
var DIARY_LABELS = DIARY_HEAD.slice(3, 10);   // 단어 ~ 기타 (칸 이름)
var DIARY_BODY = 13;        // A~M열(예전 학생 입력부). N~P열은 선생님 검사(완료 칸 이름, 검사 시각, 미완료 칸 이름)
var ST_CHECK = "검사";      // 선생님 검사로 만들어진 줄
var SHEET_SET = "설정";     // B1 = 선생님 확인 비밀번호
var DIARY_KEYS = ["vocab", "listen", "read", "gram", "syn", "write", "etc"];   // 시트 D~J열과 같은 순서 (앱 DIARY_F와 같게)
var DIARY_MAX = 300;        // 칸 하나에 쓸 수 있는 글자 수
var DIARY_SCAN = 6000;      // 시트 맨 아래에서 몇 줄까지 찾아볼지
var DIARY_SHOW = 30;        // 앱에 돌려주는 최근 알림장 개수
var TZ = "Asia/Seoul";

/* 주간 모의고사 OMR (2026-11-25) */
var SHEET_MOCK = "모의고사", SHEET_MANS = "모의고사 답안";
var MOCK_HEAD = ["시험 날짜", "시험 이름", "대상(비우면 전체)", "정답", "3점 문항", "듣기 문항 수", "메모", "등록 시각"];
var MANS_HEAD = ["제출 시각", "아이디", "이름", "시험 날짜", "시험 이름", "점수", "맞은 수", "듣기", "독해", "틀린 번호", "학생 답안"];
var MOCK_OPEN_DAYS = 14;    // 시험 날짜부터 며칠 동안 학생 화면에 보일지
var MOCK_KEEP_DAYS = 60;    // 이미 낸 시험 결과를 학생 화면에 며칠 동안 보여 줄지
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
    dy.getRange("A:P").setNumberFormat("@");   // 전부 글자로 저장(=로 시작하는 글이 수식이 되지 않게)
    dy.appendRow(DIARY_HEAD);
    dy.setFrozenRows(1);
  }
  var se = ss.getSheetByName(SHEET_SET);
  if (!se) {
    se = ss.insertSheet(SHEET_SET);
    se.getRange("B:B").setNumberFormat("@");
    se.appendRow(["선생님 확인 비밀번호 (숫자 4자리를 오른쪽 칸에 적어 주세요)", ""]);
  }
  var mk = ss.getSheetByName(SHEET_MOCK);
  if (!mk) {
    mk = ss.insertSheet(SHEET_MOCK);
    mk.getRange("A:H").setNumberFormat("@");   // 정답(35241…)이 큰 숫자로 바뀌지 않게 전부 글자로
    mk.appendRow(MOCK_HEAD);
    mk.setFrozenRows(1);
  }
  var ma = ss.getSheetByName(SHEET_MANS);
  if (!ma) {
    ma = ss.insertSheet(SHEET_MANS);
    ma.getRange("A:K").setNumberFormat("@");
    ma.appendRow(MANS_HEAD);
    ma.setFrozenRows(1);
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
  if (q.act === "diary_get" || q.act === "diary_save" || q.act === "diary_confirm" || q.act === "diary_check") return diary_(me, q);
  if (/^mock_/.test(String(q.act || ""))) return mock_(me, q);
  return { ok: false, code: "act" };
}

/* ---------- 주간 모의고사 OMR (2026-11-25) ----------
   선생님: 시험마다 정답·3점 문항·대상을 등록(선생님 확인 비밀번호 = 설정 B1).
   학생: 앱에서 OMR처럼 답을 칠해 한 번만 제출 → 서버가 채점해 '모의고사 답안'·'결과' 시트에 기록.
   학생 화면에는 점수·틀린 번호만 보내고 정답은 보내지 않음. 정답을 고치면 이미 낸 학생도 새 정답으로 다시 채점됨. */
var CIRCLE_ = { "①": 1, "②": 2, "③": 3, "④": 4, "⑤": 5 };
function mockKey_(v) {                    // 정답 칸 → [3,5,2,…]. 숫자로 바뀐 칸(지수 표기 등)은 못 씀 → null
  if (typeof v === "number") return null;
  var s = String(v == null ? "" : v), out = [];
  for (var i = 0; i < s.length; i++) {
    var c = s.charAt(i);
    if (CIRCLE_[c]) out.push(CIRCLE_[c]);
    else if (c >= "1" && c <= "5") out.push(Number(c));
  }
  return out;
}
function mockNums_(v) {
  var out = [];
  String(v == null ? "" : v).split(/[^0-9]+/).forEach(function (x) { var n = Number(x); if (x && n > 0 && out.indexOf(n) < 0) out.push(n); });
  return out.sort(function (a, b) { return a - b; });
}
function mockExam_(v, row) {
  var key = mockKey_(v[3]), n = key ? key.length : 0;
  var ls = String(v[5] == null ? "" : v[5]).trim(), listen = ls === "" ? (n === 45 ? 17 : 0) : Math.max(0, Math.min(n, Number(ls) || 0));
  var date = dstr_(v[0]), name = String(v[1] == null ? "" : v[1]).trim();
  return { row: row, mid: date + "|" + name, date: date, name: name, target: String(v[2] == null ? "" : v[2]).trim(),
    key: key || [], bad: !key, n: n, three: mockNums_(v[4]).filter(function (x) { return x <= n; }), listen: listen, memo: String(v[6] == null ? "" : v[6]) };
}
function mockExams_() {
  var sh = sheet_(SHEET_MOCK), n = sh.getLastRow(), out = [];
  if (n < 2) return out;
  var vals = sh.getRange(2, 1, n - 1, MOCK_HEAD.length).getValues();
  for (var i = 0; i < vals.length; i++) {
    var e = mockExam_(vals[i], i + 2);
    if (/^\d{4}-\d{2}-\d{2}$/.test(e.date) && e.name) out.push(e);
  }
  return out;
}
function mockFor_(me, e) {               // 대상 칸: 비우면 전체. "고1, 배명고2"처럼 쓰면 이름(학교+학년)이나 아이디에 그 말이 들어간 학생만
  var t = e.target.split(/[,，、\/]+/).map(key_).filter(function (x) { return x; });
  if (!t.length) return true;
  var nm = key_(me.name), id = key_(me.id);
  return t.some(function (x) { return nm.indexOf(x) >= 0 || id === x; });
}
function mockAns_(s, n) {                // 학생 답안 글자 "35021…"(0 = 안 칠함) → 배열
  var a = String(s == null ? "" : s).replace(/[^0-5]/g, "").split("").map(Number);
  while (a.length < n) a.push(0);
  return a.slice(0, n);
}
function mockGrade_(e, ans) {
  var three = {}, has3 = e.three.length > 0, sc = 0, mx = 0, c = 0, lc = 0, wrong = [];
  e.three.forEach(function (x) { three[x] = 1; });
  for (var i = 0; i < e.n; i++) {
    var p = has3 ? (three[i + 1] ? 3 : 2) : 1, ok = ans[i] === e.key[i];
    mx += p;
    if (ok) { sc += p; c++; if (i < e.listen) lc++; } else wrong.push(i + 1);
  }
  return { score: has3 ? sc : (e.n ? Math.round(c * 100 / e.n) : 0), max: has3 ? mx : 100, scaled: !has3, correct: c, total: e.n,
    lc: lc, lt: e.listen, rc: c - lc, rt: e.n - e.listen, wrong: wrong };
}
function mockScoreTxt_(g) { return g.scaled ? "환산 " + g.score + "점" : g.max === 100 ? g.score + "점" : g.score + "/" + g.max + "점"; }
function mockSubs_() {
  var sh = sheet_(SHEET_MANS), n = sh.getLastRow();
  if (n < 2) return { sh: sh, vals: [] };
  return { sh: sh, vals: sh.getRange(2, 1, n - 1, MANS_HEAD.length).getValues() };
}
function mockSubKey_(v) { return dstr_(v[3]) + "|" + String(v[4] == null ? "" : v[4]).trim(); }
function mockPub_(e) { return { mid: e.mid, date: e.date, name: e.name, n: e.n, listen: e.listen }; }
function mockAdmin_(e, subs) {
  var cnt = 0, sum = 0;
  subs.vals.forEach(function (v) { if (mockSubKey_(v) === e.mid) { cnt++; sum += mockGrade_(e, mockAns_(v[10], e.n)).score; } });
  return { mid: e.mid, date: e.date, name: e.name, target: e.target, key: e.key.join(""), bad: e.bad, n: e.n, three: e.three, listen: e.listen, memo: e.memo,
    cnt: cnt, avg: cnt ? Math.round(sum / cnt * 10) / 10 : null };
}
function mockRegrade_(e, subs) {         // 정답을 고친 뒤 '모의고사 답안' 시트의 점수 칸도 새로 맞춤
  subs.vals.forEach(function (v, i) {
    if (mockSubKey_(v) !== e.mid) return;
    var g = mockGrade_(e, mockAns_(v[10], e.n));
    subs.sh.getRange(i + 2, 6, 1, 5).setValues([[mockScoreTxt_(g), g.correct + "/" + g.total, g.lt ? g.lc + "/" + g.lt : "", g.lt ? g.rc + "/" + g.rt : "", g.wrong.join(", ")]]);
  });
}
function mock_(me, q) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var today = todayKst_(), exams = mockExams_(), subs = mockSubs_();
    if (q.act === "mock_list") {
      var from = Utilities.formatDate(new Date(Date.now() - MOCK_OPEN_DAYS * 864e5), TZ, "yyyy-MM-dd");
      var keep = Utilities.formatDate(new Date(Date.now() - MOCK_KEEP_DAYS * 864e5), TZ, "yyyy-MM-dd");
      var mine = {};
      subs.vals.forEach(function (v) { if (key_(v[1]) === key_(me.id)) mine[mockSubKey_(v)] = v; });
      var out = [];
      exams.forEach(function (e) {
        if (e.bad || !e.n) return;
        var v = mine[e.mid], x = mockPub_(e);
        if (v) {
          if (e.date < keep) return;
          var ans = mockAns_(v[10], e.n), g = mockGrade_(e, ans);
          g.at = tstr_(v[0]); g.ans = ans.join(""); x.done = g;
          out.push(x);
        } else if (e.date <= today && e.date >= from && mockFor_(me, e)) out.push(x);
      });
      out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : a.name < b.name ? -1 : 1; });
      return { ok: true, today: today, exams: out };
    }
    if (q.act === "mock_submit") {
      var ex = null;
      exams.forEach(function (e) { if (e.mid === String(q.mid || "") && !e.bad && e.n) ex = e; });
      if (!ex) return { ok: false, code: "noexam" };
      if (ex.date > today || !mockFor_(me, ex)) return { ok: false, code: "noexam" };
      for (var i = 0; i < subs.vals.length; i++)
        if (key_(subs.vals[i][1]) === key_(me.id) && mockSubKey_(subs.vals[i]) === ex.mid) return { ok: false, code: "done" };
      var raw = Array.isArray(q.ans) ? q.ans : [];
      var ans = []; for (var j = 0; j < ex.n; j++) { var a = Number(raw[j]) || 0; ans.push(a >= 1 && a <= 5 ? a : 0); }
      var g = mockGrade_(ex, ans), at = nowKst_();
      subs.sh.appendRow([at, txt_(me.id), txt_(me.name), ex.date, txt_(ex.name), mockScoreTxt_(g), g.correct + "/" + g.total,
        g.lt ? g.lc + "/" + g.lt : "", g.lt ? g.rc + "/" + g.rt : "", g.wrong.join(", "), "'" + ans.join("")]);
      var rs = sheet_(SHEET_RES);
      rs.appendRow([new Date(), txt_(me.id), txt_(me.name), "모의고사", txt_(ex.name), ex.date, g.scaled ? g.score : Math.round(g.score * 100 / g.max),
        g.correct + "/" + g.total, "제출", "", g.wrong.length ? "틀린 번호 " + g.wrong.join(", ") : "", "", ""]);
      g.at = at; g.ans = ans.join("");
      var x2 = mockPub_(ex); x2.done = g;
      return { ok: true, exam: x2 };
    }
    /* 선생님 */
    var bad = pinCheck_(me, q.pin);
    if (bad) return bad;
    var list = function () { return exams.slice().reverse().slice(0, 40).map(function (e) { return mockAdmin_(e, subs); }); };
    if (q.act === "mock_admin_list") return { ok: true, today: today, exams: list() };
    if (q.act === "mock_admin_save" || q.act === "mock_admin_del") {
      var sh = sheet_(SHEET_MOCK), orig = null;
      exams.forEach(function (e) { if (e.mid === String(q.orig || "")) orig = e; });
      var cntOf = function (e) { return subs.vals.filter(function (v) { return mockSubKey_(v) === e.mid; }).length; };
      if (q.act === "mock_admin_del") {
        if (!orig) return { ok: false, code: "noexam" };
        if (cntOf(orig)) return { ok: false, code: "hassubs" };
        sh.deleteRow(orig.row);
        exams = mockExams_();
        return { ok: true, exams: list() };
      }
      var date = String(q.date || ""), name = String(q.name || "").replace(/\s+/g, " ").trim().slice(0, 60);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, code: "date" };
      if (!name) return { ok: false, code: "name" };
      var key = (Array.isArray(q.key) ? q.key : []).map(Number);
      if (!key.length || key.length > 100 || key.some(function (k) { return !(k >= 1 && k <= 5); })) return { ok: false, code: "key" };
      var mid = date + "|" + name;
      var dup = exams.some(function (e) { return e.mid === mid && (!orig || e.row !== orig.row); });
      if (dup) return { ok: false, code: "dup" };
      if (orig && orig.mid !== mid && cntOf(orig)) return { ok: false, code: "rename" };
      var three = (Array.isArray(q.three) ? q.three : []).map(Number).filter(function (x) { return x >= 1 && x <= key.length; });
      var ls = Math.max(0, Math.min(key.length, Number(q.listen) || 0));
      var g5 = function (k) { var s = ""; for (var i = 0; i < k.length; i++) s += (i && i % 5 === 0 ? " " : "") + k[i]; return s; };
      var row = [date, txt_(name), txt_(String(q.target || "").trim().slice(0, 80)), g5(key), three.join(", "), String(ls), txt_(String(q.memo || "").slice(0, 200)), nowKst_()];
      var r = orig ? orig.row : sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, MOCK_HEAD.length).setNumberFormat("@").setValues([row]);
      var ne = mockExam_(row, r);
      if (orig) mockRegrade_(ne, subs);
      exams = mockExams_();
      return { ok: true, mid: mid, exams: list() };
    }
    return { ok: false, code: "act" };
  } finally {
    lock.releaseLock();
  }
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
/* 선생님이 '완료 확인'한 칸: 시트에는 "단어, 듣기"처럼 칸 이름으로, 앱에는 키 배열로 */
function parseCk_(s) {
  var out = [], parts = String(s == null ? "" : s).split(/[,，、\s]+/);
  for (var i = 0; i < parts.length; i++) {
    var k = DIARY_LABELS.indexOf(parts[i]);
    if (k >= 0 && out.indexOf(DIARY_KEYS[k]) < 0) out.push(DIARY_KEYS[k]);
  }
  return out;
}
function ckText_(keys) {
  var out = [];
  for (var k = 0; k < DIARY_KEYS.length; k++) if (keys.indexOf(DIARY_KEYS[k]) >= 0) out.push(DIARY_LABELS[k]);
  return out.join(", ");
}
function entryOf_(v) {
  var f = {};
  for (var k = 0; k < DIARY_KEYS.length; k++) f[DIARY_KEYS[k]] = String(v[3 + k] == null ? "" : v[3 + k]);
  return { d: dstr_(v[0]), f: f, st: String(v[10] || ST_DRAFT), conf: tstr_(v[11]), upd: tstr_(v[12]), ck: parseCk_(v[13]), ckt: tstr_(v[14]), nk: parseCk_(v[15]) };
}
/* 선생님 확인 비밀번호: 시트 '설정' B1. 10번 틀리면 그 학생 계정에서 10분 동안 막음 */
function pinCheck_(me, pin) {
  var cache = CacheService.getScriptCache(), fk = "pinfail:" + key_(me.id), fails = Number(cache.get(fk) || 0);
  if (fails >= MAX_FAIL) return { ok: false, code: "pinlocked" };
  var cell = sheet_(SHEET_SET).getRange(1, 2).getValues()[0][0];
  var real = String(cell == null ? "" : cell).trim();
  if (!real) return { ok: false, code: "nopin" };
  if (String(pin == null ? "" : pin).trim() !== real) { cache.put(fk, String(fails + 1), 600); return { ok: false, code: "badpin" }; }
  cache.remove(fk);
  return null;
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
    if (String(sh.getRange(1, 16).getValue() || "") !== DIARY_HEAD[15]) {   // 예전 시트에는 N~P열 제목을 붙임
      sh.getRange("N:P").setNumberFormat("@");
      sh.getRange(1, 14, 1, 3).setValues([[DIARY_HEAD[13], DIARY_HEAD[14], DIARY_HEAD[15]]]);
    }
    var win = diaryWindow_(sh);
    sweepDiary_(sh, win, today);
    if (q.act === "diary_get") return { ok: true, today: today, entries: diaryList_(win, me.id) };
    if (q.act === "diary_check") {          // 선생님 검사 (2026-10-20): 칸 하나를 완료(ok)·미완료(no)·지움("")으로. 지난 날짜도 됨. 줄이 없으면 만듦
      var bad = pinCheck_(me, q.pin);
      if (bad) return bad;
      var d2 = String(q.date || "");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d2)) return { ok: false, code: "date" };
      if (d2 > today) return { ok: false, code: "future" };
      var kk = DIARY_KEYS.indexOf(String(q.key || ""));
      if (kk < 0) return { ok: false, code: "key" };
      var st2 = q.state === "ok" ? "ok" : q.state === "no" ? "no" : "";
      var i2 = diaryFind_(win, me.id, d2), v2, r2;
      if (i2 >= 0) { v2 = win.vals[i2]; r2 = win.first + i2; }
      else {
        v2 = [d2, me.id, me.name, "", "", "", "", "", "", "", ST_CHECK, "", nowKst_(), "", "", ""];
        r2 = sh.getLastRow() + 1;
        sh.getRange(r2, 1, 1, DIARY_BODY).setValues([v2.slice(0, DIARY_BODY)]);
      }
      var ck2 = parseCk_(v2[13]), nk2 = parseCk_(v2[15]), key2 = DIARY_KEYS[kk];
      ck2 = ck2.filter(function (x) { return x !== key2; }); nk2 = nk2.filter(function (x) { return x !== key2; });
      if (st2 === "ok") ck2.push(key2); else if (st2 === "no") nk2.push(key2);
      v2[13] = ckText_(ck2); v2[14] = nowKst_(); v2[15] = ckText_(nk2);
      sh.getRange(r2, 14, 1, 3).setValues([[v2[13], v2[14], v2[15]]]);
      return { ok: true, entry: entryOf_(v2) };
    }
    var date = String(q.date || "");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, code: "date" };
    if (date > today) return { ok: false, code: "future" };
    if (date < today) return { ok: false, code: "locked" };       // 날짜가 지난 알림장은 고칠 수 없음
    var idx = diaryFind_(win, me.id, date);
    if (idx >= 0 && win.vals[idx][10] !== ST_DRAFT) return { ok: false, code: "locked" };   // 확정·자동 잠금
    var f = cleanFields_(q.f), confirm = q.act === "diary_confirm";
    if (confirm && !f.some(function (x) { return x !== ""; })) return { ok: false, code: "empty" };
    var now = nowKst_();
    var oldCk = idx >= 0 ? parseCk_(win.vals[idx][13]) : [], newCk = [];
    for (var m = 0; m < DIARY_KEYS.length; m++)      // 선생님이 확인한 뒤 내용이 바뀐 칸은 확인이 풀림
      if (oldCk.indexOf(DIARY_KEYS[m]) >= 0 && String(win.vals[idx][3 + m]) === f[m] && f[m] !== "") newCk.push(DIARY_KEYS[m]);
    var row = [date, me.id, me.name].concat(f, [confirm ? ST_DONE : ST_DRAFT, confirm ? now : "", now]);   // DIARY_BODY(13)칸
    var r = idx >= 0 ? win.first + idx : sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, DIARY_BODY).setValues([row]);           // 선생님 확인 열(N~O)은 건드리지 않음
    if (newCk.length !== oldCk.length) sh.getRange(r, 14).setValue(ckText_(newCk));
    return { ok: true, entry: entryOf_(row.concat([ckText_(newCk), idx >= 0 ? win.vals[idx][14] : "", idx >= 0 ? win.vals[idx][15] : ""])) };
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
