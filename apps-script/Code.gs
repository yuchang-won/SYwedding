/**
 * 모바일 청첩장 방명록 — Google Apps Script
 * 스프레드시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣고 웹 앱으로 배포하세요.
 *
 * 시트 구조 (처음 실행하면 자동 생성):
 *   A: 작성일시 | B: 이름 | C: 메시지 | D: 숨김
 * D열(숨김)에 아무 글자나 적으면 그 메시지는 청첩장에서 안 보입니다. (스팸·악성 글 관리용)
 */

const SHEET_NAME = '방명록';
const NAME_MAX = 10;
const MSG_MAX = 50;
const LIST_MAX = 300;          // 청첩장에 내려주는 최대 개수
const DAILY_LIMIT = 500;       // 하루 최대 등록 수 (도배 방지용 안전장치)

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['작성일시', '이름', '메시지', '숨김']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 글자 수는 이모지까지 한 글자로 셉니다
function clip_(s, max) {
  return Array.from(String(s || '').trim().replace(/\s+/g, ' ')).slice(0, max).join('');
}

// 시트에서 수식으로 실행되지 않도록 (=, +, -, @ 로 시작하는 값)
function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function list_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .filter(r => r[0] && !r[3])
    .map(r => ({ at: new Date(r[0]).toISOString(), name: String(r[1]), msg: String(r[2]) }))
    .reverse()
    .slice(0, LIST_MAX);
}

// 조회: GET
function doGet() {
  try {
    return json_({ ok: true, list: list_() });
  } catch (err) {
    return json_({ ok: false, error: 'read_failed' });
  }
}

// 등록: POST (body = JSON 문자열, Content-Type: text/plain)
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // 숨은 칸이 채워져 있으면 봇 → 저장하지 않고 성공처럼 응답
    if (d.hp) return json_({ ok: true, list: list_() });

    const name = clip_(d.name, NAME_MAX);
    const msg = clip_(d.msg, MSG_MAX);
    if (!name || !msg) return json_({ ok: false, error: 'empty' });

    // 하루 등록 수 제한
    const cache = CacheService.getScriptCache();
    const key = 'cnt-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
    const cnt = Number(cache.get(key) || 0);
    if (cnt >= DAILY_LIMIT) return json_({ ok: false, error: 'limit' });
    cache.put(key, String(cnt + 1), 60 * 60 * 24);

    sheet_().appendRow([new Date(), safe_(name), safe_(msg), '']);
    return json_({ ok: true, list: list_() });
  } catch (err) {
    return json_({ ok: false, error: 'write_failed' });
  } finally {
    lock.releaseLock();
  }
}
