// =====================================================
// 辰光國小 校務系統 — 週三進修 GAS 腳本 v3
// 貼到 script.google.com 取代舊版後：
//   1. 選「updateFromDrive」函式 → 點「執行」→ 授權 Drive + Sheets
//   2. 部署 → 網路應用程式 → 執行身份：我 → 存取：所有人
//   3. 複製部署 URL 填入 index.html 的 GAS_TRAINING_URL
// =====================================================

const FOLDER_ID = '1FAb8TBPYhbroQIigcUYjHqzasAX20sC6';
const SHEET_ID  = '1RknaPoxdbJ9tisudv6cJRZl-dyVWr-H3ICw-GHPRnnk';

// ── doGet：供首頁 fetch 呼叫 ──────────────────────────
function doGet(e) {
  try {
    const forceUpdate = e && e.parameter && e.parameter.force === '1';

    // 非強制更新：先看試算表快取是否為本週資料
    if (!forceUpdate) {
      const cached = readFromSheet();
      if (cached.found && isCurrentWeek(cached.wedDate)) {
        return jsonResponse(cached);
      }
    }

    // 快取過期或強制更新：從 Drive 讀取並寫入試算表
    const result = getWeeklyTrainingFromDrive();
    if (result.found) writeToSheet(result);
    return jsonResponse(result);

  } catch (err) {
    return jsonResponse({ found: false, error: err.message });
  }
}

// ── 手動執行這個函式完成授權 & 立刻更新本週資料 ─────────
// （在 Apps Script 編輯器選此函式 → 執行 → 點「授權」）
function updateFromDrive() {
  const result = getWeeklyTrainingFromDrive();
  if (result.found) {
    writeToSheet(result);
    Logger.log('✅ 寫入成功：' + result.trainingName + '  日期：' + result.wedDate);
    Logger.log('來源檔案：' + result.fileName);
  } else {
    Logger.log('❌ 找不到本週資料：' + JSON.stringify(result));
  }
}

// ── 試算表讀取 ────────────────────────────────────────
function readFromSheet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const row   = sheet.getRange(2, 1, 1, 6).getValues()[0];
  if (!row || !row[0]) return { found: false };
  return {
    found:        true,
    wedDate:      row[0],
    trainingName: row[1],
    location:     row[2],
    source:       row[3],
    updatedAt:    row[4],
    note:         row[5]
  };
}

// ── 試算表寫入 ────────────────────────────────────────
function writeToSheet(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  sheet.getRange(1, 1, 1, 6).setValues([[
    'wedDate', 'trainingName', 'location', 'source', 'updatedAt', 'note'
  ]]);
  sheet.getRange(2, 1, 1, 6).setValues([[
    data.wedDate      || '',
    data.trainingName || '',
    data.location     || '',
    data.source       || '',
    new Date().toISOString().split('T')[0],
    data.huangReport  || ''
  ]]);
}

// ── 判斷日期是否為本週三 ──────────────────────────────
function isCurrentWeek(wedDateStr) {
  if (!wedDateStr) return false;
  const today  = new Date();
  const sunday = new Date(today);
  sunday.setDate(today.getDate() - today.getDay());
  sunday.setHours(0, 0, 0, 0);
  const friday = new Date(sunday);
  friday.setDate(sunday.getDate() + 5);
  friday.setHours(23, 59, 59, 0);
  const wedDate = new Date(wedDateStr);
  return wedDate >= sunday && wedDate <= friday;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// 從 Drive 資料夾讀取本週夕會報告
// =====================================================
function getWeeklyTrainingFromDrive() {
  const today     = new Date();
  const sunday    = new Date(today);
  sunday.setDate(today.getDate() - today.getDay());

  const wednesday = new Date(sunday);
  wednesday.setDate(sunday.getDate() + 3);

  const monday = new Date(sunday);
  monday.setDate(sunday.getDate() + 1);

  const mm    = String(wednesday.getMonth() + 1).padStart(2, '0');
  const dd    = String(wednesday.getDate()).padStart(2, '0');
  const monMm = String(monday.getMonth() + 1).padStart(2, '0');
  const monDd = String(monday.getDate()).padStart(2, '0');
  const wedDateStr = `${wednesday.getFullYear()}-${mm}-${dd}`;

  // 搜尋策略：週一日期 → 週三日期 → 最新檔案
  let targetFile = null;
  const searchKeys = [monMm + monDd, mm + dd];

  for (const key of searchKeys) {
    try {
      const iter = DriveApp.searchFiles(
        `'${FOLDER_ID}' in parents and title contains '${key}' and trashed = false`
      );
      if (iter.hasNext()) { targetFile = iter.next(); break; }
    } catch (e) { /* 繼續嘗試 */ }
  }

  // 找不到本週 → 取資料夾內最新的 .docx
  if (!targetFile) {
    try {
      const iter = DriveApp.searchFiles(
        `'${FOLDER_ID}' in parents and title contains '.docx' and trashed = false`
      );
      let latest = null;
      while (iter.hasNext()) {
        const f = iter.next();
        if (!latest || f.getLastUpdated() > latest.getLastUpdated()) latest = f;
      }
      targetFile = latest;
    } catch (e) {
      return { found: false, message: 'Drive 搜尋失敗：' + e.message + '（請確認已完成授權）' };
    }
  }

  if (!targetFile) {
    return { found: false, message: '本週夕會報告尚未上傳到 Drive 資料夾' };
  }

  // 讀取文件純文字
  let content = '';
  try {
    content = targetFile.getAs('text/plain').getDataAsString('UTF-8');
  } catch (e) {
    return { found: false, message: '讀取 docx 失敗：' + e.message };
  }

  const lines = content.split('\n');
  let huangLines = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line.includes('黃雅玲')) {
      huangLines.push(line);
      for (let j = 1; j <= 8; j++) {
        if (lines[i + j]) huangLines.push(lines[i + j].trim());
      }
      break;
    }
  }

  const huangReport = huangLines.join(' ');

  // 解析週三進修
  let trainingName = '';
  let trainingTime = '';
  const wedMonth = parseInt(mm, 10);
  const wedDay   = parseInt(dd, 10);

  const patterns = [
    // 「週三進修-1330舞蹈韻律VS體能訓練」
    /週三進修[-\s]*(\d{3,4})?\s*([^\d。，、\n]{2,30})/,
    // 「6/18週三進修...」
    new RegExp(`${wedMonth}[/／]${wedDay}[^\\n]*週三進修[^\\n]*`),
    // 「週三進修...」任意格式
    /週三進修[^\n]*/,
  ];

  for (const pat of patterns) {
    const m = pat.exec(huangReport);
    if (m) {
      if (m[2]) {
        // 有時間+名稱
        if (m[1] && m[1].length >= 3) {
          trainingTime = m[1].substring(0, 2) + ':' + m[1].substring(2);
        }
        trainingName = m[2].replace(/[。，、].*/g, '').trim();
      } else {
        trainingName = m[0]
          .replace(/.*週三進修[-\s]*\d*\s*/, '')
          .replace(/[。，、].*/g, '')
          .trim();
      }
      if (trainingName && trainingName.length > 1) break;
    }
  }

  if (!trainingName || trainingName.length <= 1) {
    trainingName = '（請查閱夕會報告）';
  }

  return {
    found:        true,
    fileName:     targetFile.getName(),
    wedDate:      wedDateStr,
    trainingName: trainingName,
    location:     trainingTime,
    source:       `${targetFile.getName()}，黃雅玲主任報告`,
    huangReport:  huangReport.substring(0, 400)
  };
}
