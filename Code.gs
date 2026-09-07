/**
 * 워드클라우드 수집 백엔드
 * 고양 스마트시티즌 · 제미나이·노트북LM으로 끝내는 AI 업무 자동화
 *
 * 구글 시트에 붙여넣고 웹앱으로 배포하면 응답 수집 주소가 생깁니다.
 * 자세한 순서는 README.md 를 보세요.
 */

// ── 설정 ──────────────────────────────
var SHEET_NAME = '응답';      // 응답이 쌓일 시트 이름
var ADMIN_PIN  = '2026';      // '전체 지우기' 를 누를 때 필요한 비밀번호
// ─────────────────────────────────────


/** 응답 시트를 가져오고, 없으면 만듭니다. */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['시각', '참여자ID', '고른 항목']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** JSON 으로 응답합니다. */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 모든 요청이 여기로 들어옵니다.
 *   ?action=submit&id=xxx&words=엑셀 취합|보고서 작성
 *   ?action=list
 *   ?action=clear&pin=2026
 */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var action = p.action || 'list';

  try {
    if (action === 'submit') return submit_(p);
    if (action === 'clear')  return clear_(p);
    return list_();
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 참여자가 고른 항목을 한 줄 추가합니다. */
function submit_(p) {
  var raw = (p.words || '').trim();
  if (!raw) return json_({ ok: false, error: '보낼 항목이 없습니다' });

  // 항목은 | 로 구분해서 받습니다. 빈 값과 너무 긴 값은 걸러냅니다.
  var words = raw.split('|')
    .map(function (w) { return w.trim().replace(/\s+/g, ' '); })
    .filter(function (w) { return w.length > 0 && w.length <= 20; })
    .slice(0, 10);

  if (!words.length) return json_({ ok: false, error: '보낼 항목이 없습니다' });

  // 여러 명이 동시에 눌러도 줄이 섞이지 않도록 잠급니다.
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    getSheet_().appendRow([new Date(), p.id || '', words.join(', ')]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, saved: words.length });
}

/** 지금까지 모인 응답을 세어서 돌려줍니다. */
function list_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return json_({ ok: true, people: 0, answers: 0, tally: {} });

  var rows = sh.getRange(2, 3, last - 1, 1).getValues();  // C열(고른 항목)만
  var tally = {}, answers = 0;

  rows.forEach(function (r) {
    String(r[0] || '').split(',').forEach(function (w) {
      w = w.trim();
      if (!w) return;
      tally[w] = (tally[w] || 0) + 1;
      answers++;
    });
  });

  return json_({ ok: true, people: last - 1, answers: answers, tally: tally });
}

/** 응답을 모두 지웁니다. 비밀번호가 맞아야 합니다. */
function clear_(p) {
  if (p.pin !== ADMIN_PIN) return json_({ ok: false, error: '비밀번호가 맞지 않습니다' });
  var sh = getSheet_();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  return json_({ ok: true, cleared: true });
}
