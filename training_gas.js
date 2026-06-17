// =====================================================
// 辰光國小 校務系統 — 週三進修自動讀取腳本
// Google Apps Script  (貼到 script.google.com)
// =====================================================

const FOLDER_ID = '1FAb8TBPYhbroQIigcUYjHqzasAX20sC6';

function doGet() {
  try {
    const result = getWeeklyTraining();
    const output = ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
    return output;
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ found: false, error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getWeeklyTraining() {
  // 週次以週日為起點（5/31 日 ~ 6/5 五）
  // 本週三 = 本週日 + 3 天，無論今天是哪天都固定
  const today = new Date();
  const sunday = new Date(today);
  sunday.setDate(today.getDate() - today.getDay()); // 回到本週日

  const wednesday = new Date(sunday);
  wednesday.setDate(sunday.getDate() + 3); // 本週三

  const monday = new Date(sunday);
  monday.setDate(sunday.getDate() + 1);   // 本週一（夕會報告通常以週一日期命名）

  const mm    = String(wednesday.getMonth() + 1).padStart(2, '0');
  const dd    = String(wednesday.getDate()).padStart(2, '0');
  const monMm = String(monday.getMonth() + 1).padStart(2, '0');
  const monDd = String(monday.getDate()).padStart(2, '0');
  const wedDateStr = `${wednesday.getFullYear()}-${mm}-${dd}`;

  // 搜尋夕會報告檔案（使用 searchFiles，支援共用資料夾）
  let targetFile = null;

  // 先找週一日期（檔案通常命名如「第15週-0601.docx」），再找週三日期
  const searchKeys = [monMm + monDd, mm + dd];
  for (const key of searchKeys) {
    const iter = DriveApp.searchFiles(
      `'${FOLDER_ID}' in parents and title contains '${key}' and trashed = false`
    );
    if (iter.hasNext()) {
      targetFile = iter.next();
      break;
    }
  }

  // 找不到就取資料夾內最新修改的 docx 檔案
  if (!targetFile) {
    const iter = DriveApp.searchFiles(
      `'${FOLDER_ID}' in parents and title contains '.docx' and trashed = false`
    );
    let latest = null;
    while (iter.hasNext()) {
      const f = iter.next();
      if (!latest || f.getLastUpdated() > latest.getLastUpdated()) latest = f;
    }
    targetFile = latest;
  }

  if (!targetFile) {
    return { found: false, message: '找不到夕會報告，請確認 Drive 資料夾內有檔案' };
  }

  // 讀取檔案純文字內容
  const blob = targetFile.getAs('text/plain');
  const content = blob.getDataAsString('UTF-8');

  // 找黃雅玲主任的報告欄位
  const lines = content.split('\n');
  let huangLines = [];
  let inHuang = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!inHuang && line.includes('黃雅玲')) {
      inHuang = true;
      huangLines.push(line);
      // 抓接下來最多 4 行（報告事項可能跨行）
      for (let j = 1; j <= 4; j++) {
        if (lines[i + j]) huangLines.push(lines[i + j].trim());
      }
      break;
    }
  }

  const huangReport = huangLines.join(' ');

  // 解析週三進修內容
  // 格式範例：「2.6/3週三進修食農教育-中廊」「3.6/10吉他及兒童權利公約」
  let trainingName = '';
  let trainingNote = '';

  const wedMonth = parseInt(mm, 10);
  const wedDay   = parseInt(dd, 10);

  // 嘗試比對「M/D週三進修...」或「M/D...進修...」
  const patterns = [
    new RegExp(`${wedMonth}/${wedDay}[\\s]*週三進修[\\s]*([^\\d。，、\\n]+)`),
    new RegExp(`${wedMonth}[/／]${wedDay}[^週]*週三[^\\n]*`),
  ];

  for (const pat of patterns) {
    const m = pat.exec(huangReport);
    if (m) {
      trainingName = m[1] ? m[1].trim() : m[0].replace(/^.*週三進修\s*/, '').trim();
      trainingName = trainingName.replace(/[。，、].*/g, '').trim();
      trainingNote = m[0].trim();
      break;
    }
  }

  // 若仍找不到，回傳完整報告讓使用者判斷
  if (!trainingName) {
    trainingName = '（請查看夕會報告）';
    trainingNote = huangReport.substring(0, 200);
  }

  return {
    found: true,
    fileName: targetFile.getName(),
    wedDate: wedDateStr,
    trainingName: trainingName,
    huangReport: huangReport.substring(0, 300),
    note: `來源：${targetFile.getName()}，黃雅玲主任報告`
  };
}
