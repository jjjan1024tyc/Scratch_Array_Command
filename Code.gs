/**
 * Scratch 清單實驗室 — 成績上傳（Google Apps Script）
 *
 * 使用方式：
 * 1. 在 Google 試算表選「擴充功能 → Apps Script」，建立綁定在這份試算表的專案。
 * 2. 把這個檔案的內容貼到 Code.gs；新增 HTML 檔，命名為「index」，貼上 index.html 的內容。
 * 3. 「部署 → 新增部署作業 → 網頁應用程式」
 *    執行身分：我
 *    存取權：「（學校網域）中的所有人」
 * 4. 把部署後的網址給學生。學生用學校帳號開啟，成績會寫進這份試算表的「成績」工作表。
 */

const SHEET_NAME = '成績';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Scratch 清單實驗室')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 取得學生的帳號（同網域的使用者才取得到） */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * 儲存成績。每位學生一列，以帳號辨識。
 * 只有新的總分大於或等於原本的總分時才覆蓋，
 * 所以「再玩一次」中途上傳的較低分數，不會蓋掉之前的最佳成績。
 * data = { total: 總分, solved: 完成題數, scores: [每題得分，沒答對是 ''] }
 */
function saveScores(data) {
  const email = Session.getActiveUser().getEmail();   // 帳號以伺服器取得的為準
  if (!email) throw new Error('未取得學校帳號');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    const scores = Array.isArray(data.scores) ? data.scores : [];
    const total = Number(data.total) || 0;
    const solved = Number(data.solved) || 0;

    if (sheet.getLastRow() === 0) {
      const header = ['上傳時間', '帳號', '完成題數', '總分'];
      scores.forEach(function (_, i) { header.push('第' + (i + 1) + '題'); });
      sheet.appendRow(header);
      sheet.setFrozenRows(1);
    }

    const row = [new Date(), email, solved, total].concat(scores);
    const last = sheet.getLastRow();
    const emails = last > 1
      ? sheet.getRange(2, 2, last - 1, 1).getValues().map(function (r) { return r[0]; })
      : [];
    const idx = emails.indexOf(email);

    if (idx === -1) {
      sheet.appendRow(row);
      return 'added';
    }
    const rowNum = idx + 2;
    const oldTotal = Number(sheet.getRange(rowNum, 4).getValue()) || 0;
    if (total >= oldTotal) {
      sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
      return 'updated';
    }
    return 'kept';   // 原本的成績比較高，保留
  } finally {
    lock.releaseLock();
  }
}
