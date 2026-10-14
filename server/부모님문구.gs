/**
 * 엠제이 점검실 — 부모님 문구 만들기 (2026-10-14)
 *
 * 결과 탭의 기록을 학생마다 "부모님께 보낼 카톡 문구"로 정리해 '부모님문구' 탭에 써 줍니다.
 * 시트 위쪽 메뉴 [점검실]에서 오늘 / 최근 7일 / 기간 직접 입력을 고르면 됩니다.
 *
 * ※ 이 파일은 Apps Script에 '새 스크립트 파일'로 붙여 넣고 저장만 하면 됩니다.
 *    웹 앱을 다시 배포할 필요는 없어요. (학생 앱과는 상관없는 기능)
 */
var PM_SHEET = "부모님문구";

function onOpen() {
  SpreadsheetApp.getUi().createMenu("점검실")
    .addItem("부모님 문구 만들기 — 오늘", "부모님문구_오늘")
    .addItem("부모님 문구 만들기 — 최근 7일", "부모님문구_7일")
    .addItem("부모님 문구 만들기 — 기간 직접 입력", "부모님문구_기간")
    .addToUi();
}

function 부모님문구_오늘() { var t = pm_day_(new Date(), 0); pm_make_(t, t); }
function 부모님문구_7일() { pm_make_(pm_day_(new Date(), -6), pm_day_(new Date(), 0)); }
function 부모님문구_기간() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt("부모님 문구 — 기간", "예) 10/8~10/14   (하루만 보려면 10/14)", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var p = pm_parseRange_(r.getResponseText(), new Date());
  if (!p) { ui.alert("날짜를 알아보지 못했어요. 10/8~10/14 처럼 써 주세요."); return; }
  pm_make_(p[0], p[1]);
}

/* ---------- 날짜 ---------- */
function pm_day_(d, add) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + (add || 0)); }
function pm_md_(d) { return (d.getMonth() + 1) + "/" + d.getDate(); }
function pm_key_(d) { return d.getFullYear() + "-" + (d.getMonth() + 1) + "-" + d.getDate(); }
function pm_parseRange_(s, now) {
  var m = String(s || "").replace(/\s+/g, "").match(/^(\d{1,2})[\/.\-](\d{1,2})(?:[~\-](\d{1,2})[\/.\-](\d{1,2}))?$/);
  if (!m) return null;
  function mk(mo, da) {
    var d = new Date(now.getFullYear(), Number(mo) - 1, Number(da));
    if (d > pm_day_(now, 0)) d = new Date(now.getFullYear() - 1, Number(mo) - 1, Number(da));
    return d;
  }
  var a = mk(m[1], m[2]), b = m[3] ? mk(m[3], m[4]) : a;
  if (isNaN(a) || isNaN(b)) return null;
  return a <= b ? [a, b] : [b, a];
}

/* ---------- 만들기 ---------- */
function pm_make_(from, to) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = ss.getSheetByName("결과"), stu = ss.getSheetByName("학생");
  if (!res || !stu) { SpreadsheetApp.getUi().alert("'결과' 탭이나 '학생' 탭을 찾지 못했어요."); return; }
  var rows = pm_build_(stu.getDataRange().getValues(), res.getDataRange().getValues(), from, to, new Date());
  var sh = ss.getSheetByName(PM_SHEET) || ss.insertSheet(PM_SHEET);
  sh.clear();
  var title = (pm_key_(from) === pm_key_(to) ? pm_md_(from) : pm_md_(from) + "~" + pm_md_(to)) +
    " 기간 · 만든 시각 " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d HH:mm") +
    " · G열 문구를 복사해 부모님께 보내세요";
  sh.getRange(1, 1).setValue(title).setFontWeight("bold");
  var head = ["이름", "아이디", "공부한 날", "단원", "통과", "화면 이탈", "부모님께 보낼 문구"];
  sh.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#dde4ff");
  if (rows.length) sh.getRange(3, 1, rows.length, head.length).setValues(rows);
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 80); sh.setColumnWidths(3, 4, 70); sh.setColumnWidth(7, 420);
  sh.getRange(3, 7, Math.max(rows.length, 1), 1).setWrap(true).setVerticalAlignment("top");
  sh.getRange(3, 1, Math.max(rows.length, 1), 6).setVerticalAlignment("top");
  sh.setFrozenRows(2);
  ss.setActiveSheet(sh);
  ss.toast("학생 " + rows.length + "명 문구를 만들었어요.", "부모님문구", 5);
}

/* 학생 탭·결과 탭 값 → 부모님문구 탭 줄들 (테스트하기 쉽게 시트와 분리) */
function pm_build_(stuVals, resVals, from, to, now) {
  var T0 = pm_day_(from, 0), T1 = pm_day_(to, 1), oneDay = pm_key_(from) === pm_key_(to), todayK = pm_key_(now);
  var norm = function (v) { return String(v == null ? "" : v).replace(/\s+/g, "").toLowerCase(); };
  var students = [], byKey = {};
  for (var i = 1; i < stuVals.length; i++) {
    var id = String(stuVals[i][0] || "").trim();
    if (!id || id === "원장") continue;
    var s = { id: id, name: String(stuVals[i][2] || "").trim() || id, days: {} };
    students.push(s); byKey[norm(id)] = s;
  }
  for (var j = 1; j < resVals.length; j++) {
    var r = resVals[j], at = r[0] instanceof Date ? r[0] : new Date(r[0]);
    if (isNaN(at) || at < T0 || at >= T1) continue;
    var st = byKey[norm(r[1])];
    if (!st) continue; // 학생 탭에 없는 아이디(퇴원생 등)는 빼요
    var dk = pm_key_(at);
    var day = st.days[dk] || (st.days[dk] = { d: pm_day_(at, 0), units: {}, order: [] });
    var uk = r[3] + "|" + r[4] + "|" + r[5];
    if (!day.units[uk]) { day.units[uk] = { type: r[3], book: r[4], unit: r[5], tries: [] }; day.order.push(uk); }
    var aw = String(r[11] || ""), an = aw.match(/(\d+)회/), as = aw.match(/\((\d+)(초|분)\)/);
    day.units[uk].tries.push({ at: at, pct: Number(r[6]) || 0, res: String(r[8] || ""),
      away: an ? Number(an[1]) : 0, awaySec: as ? Number(as[1]) * (as[2] === "분" ? 60 : 1) : 0 });
  }
  var WD = "일월화수목금토", out = [];
  students.forEach(function (s) {
    var keys = Object.keys(s.days).sort(function (a, b) { return s.days[a].d - s.days[b].d; });
    var nUnit = 0, nPass = 0, nFinal = 0, aw = 0, awS = 0, body = [];
    keys.forEach(function (k) {
      var day = s.days[k];
      if (!oneDay) body.push("", "▶ " + pm_md_(day.d) + " (" + WD.charAt(day.d.getDay()) + ")");
      day.order.forEach(function (uk) {
        var u = day.units[uk]; u.tries.sort(function (a, b) { return a.at - b.at; });
        var pass = u.tries.some(function (t) { return t.res === "통과"; });
        var fin = !pass && u.tries.some(function (t) { return t.res === "조교 확인"; });
        var a = 0, as2 = 0; u.tries.forEach(function (t) { a += t.away; as2 += t.awaySec; });
        nUnit++; if (pass) nPass++; else if (fin) nFinal++; aw += a; awS += as2;
        var stTxt = pass ? "통과" : fin ? "조교 선생님과 다시 확인" : (k === todayK ? "재시험 중" : "재시험 미완료");
        body.push("■ " + u.type + " · " + u.book + (u.unit ? " " + u.unit : ""));
        body.push("  " + u.tries.map(function (t) { return t.pct; }).join(" → ") + "점 (" + u.tries.length + "회) · " + stTxt +
          (a ? " · 화면 이탈 " + a + "회 (" + pm_sec_(as2) + ")" : ""));
      });
    });
    var range = oneDay ? pm_md_(from) : pm_md_(from) + "~" + pm_md_(to);
    var lines = ["[엠제이영어학원] " + s.name + " 학생 " + (oneDay ? "오늘" : "") + " 점검 결과 (" + range + ")"];
    lines[0] = lines[0].replace("학생  점검", "학생 점검");
    if (!nUnit) lines.push("", oneDay ? "오늘은 본 시험이 없습니다." : "이 기간에 본 시험이 없습니다.");
    else {
      var nDays = keys.length, total = Math.round((pm_day_(to, 1) - pm_day_(from, 0)) / 864e5);
      lines.push("", (oneDay ? "" : "공부한 날 " + nDays + "일 / " + total + "일 · ") + nUnit + "단원 (통과 " + nPass +
        (nFinal ? " · 조교 확인 " + nFinal : "") + (nUnit - nPass - nFinal ? " · 아직 통과 전 " + (nUnit - nPass - nFinal) : "") + ")");
      if (aw) lines.push("시험 중 화면 이탈 " + aw + "회 (" + pm_sec_(awS) + ")");
      if (oneDay) lines.push("");
      lines = lines.concat(body);
    }
    lines.push("", "통과 기준은 100점이며, 틀린 문제는 재시험(3회)으로 다시 확인합니다.");
    out.push({ has: nUnit > 0, row: [s.name, s.id, nUnit ? keys.length : 0, nUnit, nPass, aw ? aw + "회" : "", lines.join("\n")] });
  });
  out.sort(function (a, b) { return (b.has - a.has) || (a.row[0] < b.row[0] ? -1 : a.row[0] > b.row[0] ? 1 : 0); });
  return out.map(function (o) { return o.row; });
}
function pm_sec_(n) { n = Number(n) || 0; return n >= 60 ? Math.round(n / 60) + "분" : n + "초"; }
