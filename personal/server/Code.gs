/**
 * 개인 점수판 — 데이터 저장용 Google Apps Script
 *
 * 사용 방법
 *  1) 아래 PIN 값을 선생님만 아는 숫자(4~6자리)로 바꾸세요.
 *  2) 저장(Ctrl+S)한 뒤, 위쪽 함수 선택에서 setup 을 고르고 [실행]을 한 번 누르세요.
 *  3) [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → [배포]
 *  4) 나오는 "웹 앱 URL"을 복사해서 Claude에게 알려 주세요.
 *
 * 시트 구성
 *  - "학생" 시트: 번호, 별명
 *  - "기록" 시트: ID, 시각, 번호, 점수, 사유   (점수 합계는 이 기록을 더해서 계산)
 */

const PIN = '0000';          // ← 꼭 바꾸세요! (점수·별명을 바꿀 때 필요한 선생님 비밀번호)
const STUDENT_COUNT = 19;    // 학생 수
const TZ = 'Asia/Seoul';

/** 처음 한 번 실행: 시트를 만들고 권한을 승인합니다. */
function setup() {
  sheets_();
}

function sheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let st = ss.getSheetByName('학생');
  if (!st) {
    st = ss.insertSheet('학생');
    const rows = [['번호', '별명']];
    for (let i = 1; i <= STUDENT_COUNT; i++) rows.push([i, i + '번']);
    st.getRange(1, 1, rows.length, 2).setValues(rows);
    st.getRange('B:B').setNumberFormat('@');
    st.setFrozenRows(1);
  }
  let lg = ss.getSheetByName('기록');
  if (!lg) {
    lg = ss.insertSheet('기록');
    lg.getRange(1, 1, 1, 5).setValues([['ID', '시각', '번호', '점수', '사유']]);
    lg.getRange('E:E').setNumberFormat('@');
    lg.setFrozenRows(1);
  }
  return { st: st, lg: lg };
}

/* ---------- 읽기 (누구나) ---------- */

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'history') {
      return out_({ ok: true, history: history_(Number(p.no)) });
    }
    return out_(Object.assign({ ok: true }, state_()));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function state_() {
  const s = sheets_();
  const st = s.st, lg = s.lg;

  const sRows = st.getLastRow() - 1;
  const students = sRows > 0 ? st.getRange(2, 1, sRows, 2).getValues() : [];

  const totals = {};
  const lRows = lg.getLastRow() - 1;
  if (lRows > 0) {
    lg.getRange(2, 3, lRows, 2).getValues().forEach(function (r) {
      const no = Number(r[0]);
      totals[no] = (totals[no] || 0) + (Number(r[1]) || 0);
    });
  }

  return {
    students: students
      .filter(function (r) { return r[0] !== '' && !isNaN(Number(r[0])); })
      .map(function (r) {
        const no = Number(r[0]);
        return { no: no, nick: String(r[1]), total: totals[no] || 0 };
      })
      .sort(function (a, b) { return a.no - b.no; }),
    records: Math.max(lRows, 0)
  };
}

function history_(no) {
  const lg = sheets_().lg;
  const n = lg.getLastRow() - 1;
  if (n <= 0) return [];
  return lg.getRange(2, 1, n, 5).getValues()
    .filter(function (r) { return Number(r[2]) === no; })
    .map(function (r) {
      const t = r[1] instanceof Date ? r[1] : new Date(r[1]);
      return {
        id: String(r[0]),
        time: isNaN(t) ? '' : Utilities.formatDate(t, TZ, "yyyy-MM-dd'T'HH:mm:ss"),
        delta: Number(r[3]) || 0,
        reason: String(r[4] || '')
      };
    })
    .reverse(); // 최신 기록이 먼저
}

/* ---------- 쓰기 (선생님 PIN 필요) ---------- */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (String(body.pin) !== String(PIN)) return out_({ ok: false, error: 'pin' });

    lock.waitLock(15000);
    const s = sheets_();

    if (body.action === 'add') {
      const no = Math.round(Number(body.no));
      const delta = Math.round(Number(body.delta));
      if (!(no >= 1 && no <= 99) || !delta || Math.abs(delta) > 100) {
        return out_({ ok: false, error: 'bad request' });
      }
      s.lg.appendRow([
        Utilities.getUuid().slice(0, 8),
        new Date(),
        no,
        delta,
        safe_(body.reason, 20)
      ]);
    } else if (body.action === 'nick') {
      const no = Math.round(Number(body.no));
      const row = findRow_(s.st, 1, no);
      if (!row) return out_({ ok: false, error: 'no student' });
      const nick = safe_(body.nick, 10) || (no + '번');
      s.st.getRange(row, 2).setValue(nick);
    } else if (body.action === 'delete') {
      const row = findRow_(s.lg, 1, String(body.id));
      if (row) s.lg.deleteRow(row);
    } else if (body.action !== 'check') {
      return out_({ ok: false, error: 'unknown action' });
    }

    SpreadsheetApp.flush();
    return out_(Object.assign({ ok: true }, state_()));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ---------- 도움 함수 ---------- */

function findRow_(sheet, col, value) {
  const n = sheet.getLastRow() - 1;
  if (n <= 0) return 0;
  const vals = sheet.getRange(2, col, n, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(value)) return i + 2;
  }
  return 0;
}

// 글자 수 제한 + 수식으로 해석되지 않도록 처리
function safe_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
