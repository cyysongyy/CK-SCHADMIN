// =====================================================
// 辰光國小 校務系統 — 學生名冊 GAS 腳本
// 目的：原本 STUDENT_ROSTER（含學生真實姓名）是直接寫死在 index.html
// 裡的，這個 repo 是公開的，等於姓名一直被公開放在 GitHub 上。
// 改成這支 GAS 腳本後，姓名只存在你自己的 Google 試算表（私有），
// index.html 只在執行期用網路請求去讀，git 裡完全不會出現姓名。
//
// 貼到 script.google.com 後：
//   1. 這個檔案本身「不含」任何學生姓名，可以直接貼上，不用擔心。
//   2. 姓名要用「一次性匯入」的方式貼進試算表——這一步我會另外給你
//      一份「只給你看、不會進 git」的匯入用程式碼，裡面才有真實姓名，
//      你貼到 Apps Script 編輯器執行一次就好，執行完就刪掉那段。
//      Apps Script 編輯器是你自己 Google 帳號底下的東西，跟 GitHub
//      完全無關，不會外流。
//   3. 部署 → 新增部署作業 → 網頁應用程式
//      - 執行身分：我（Me）
//      - 誰可以存取：所有人（Anyone）
//   4. 複製部署網址，貼到 index.html 的 GAS_ROSTER_URL 常數
// =====================================================

const ROSTER_SHEET_NAME = '學生名冊';
const ROSTER_COLS = ['cls', 'gradeNum', 'seat', 'sid', 'name'];

// ── doGet：供 index.html fetch 呼叫，回傳整份名冊 ──────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'roster';
  if (action === 'roster') {
    return jsonResponse({ ok: true, roster: getRoster() });
  }
  return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
}

// ── doPost：新增/整批覆寫名冊（給以後要做管理介面時用）────
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonResponse({ ok: false, error: 'Invalid JSON' }); }
  if (body.action === 'saveRoster') {
    saveRoster(body.roster || []);
    return jsonResponse({ ok: true, message: '已儲存名冊，共 ' + (body.roster || []).length + ' 筆' });
  }
  if (body.action === 'addStudent') {
    addStudent(body.student || {});
    return jsonResponse({ ok: true, message: '已新增學生' });
  }
  return jsonResponse({ ok: false, error: 'Unknown action: ' + body.action });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ROSTER_SHEET_NAME);
    sh.getRange(1, 1, 1, ROSTER_COLS.length).setValues([ROSTER_COLS]);
    sh.getRange(1, 1, 1, ROSTER_COLS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getRoster() {
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(r => r[0]).map(r => {
    const o = {};
    ROSTER_COLS.forEach((c, i) => { o[c] = r[i]; });
    o.gradeNum = Number(o.gradeNum);
    o.seat = Number(o.seat);
    o.sid = String(o.sid);
    return o;
  });
}

function saveRoster(roster) {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, ROSTER_COLS.length).clearContent();
  if (!roster.length) return;
  const rows = roster.map(s => ROSTER_COLS.map(c => (s[c] !== undefined ? s[c] : '')));
  sh.getRange(2, 1, rows.length, ROSTER_COLS.length).setValues(rows);
}

function addStudent(s) {
  const sh = getSheet_();
  sh.appendRow(ROSTER_COLS.map(c => (s[c] !== undefined ? s[c] : '')));
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
