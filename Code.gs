/**
 * ============================================================================
 * 🎓 教學訓練計畫主持人工作坊 — Google Apps Script 後端串接程式 (Code.gs)
 * ============================================================================
 * 試算表網址：https://docs.google.com/spreadsheets/d/1DXp8rcOjWTABT6Dg2Q6zdTXD7lhIay7DP2TNcGsBrto/edit?gid=0#gid=0
 * 目標分頁名稱：工作坊報名資料
 *
 * 【表頭欄位順序 (A1:I1)】
 * A1: 報名時間
 * B1: 姓名
 * C1: 機構名
 * D1: 職稱
 * E1: 負責的職類
 * F1: 是否擔任教學訓練計畫主持人
 * G1: 參與方式
 * H1: Email
 * I1: 聯繫電話
 * ============================================================================
 */

const SPREADSHEET_ID = '1DXp8rcOjWTABT6Dg2Q6zdTXD7lhIay7DP2TNcGsBrto';
const SHEET_NAME = '工作坊報名資料';

const HEADERS = [
  '報名時間',
  '姓名',
  '機構名',
  '職稱',
  '負責的職類',
  '是否擔任教學訓練計畫主持人',
  '參與方式',
  'Email',
  '聯繫電話'
];

/**
 * 取得或自動建立「工作坊報名資料」分頁與表頭
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  // 若尚無「工作坊報名資料」分頁，自動建立
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // 檢查第一列是否已有表頭，若無則自動寫入並美化為紫色漸層風格表頭
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.appendRow(HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange
      .setBackground('#4c1d95')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 手動初始化函式：可在 Apps Script 編輯器上方點選「執行」先建立好分頁與表頭
 */
function setupSheet() {
  const sheet = getOrCreateSheet();
  Logger.log('✅ 已成功初始化分頁：「' + sheet.getName() + '」與表頭欄位！');
}

/**
 * 處理網頁表單 POST 送出資料
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const sheet = getOrCreateSheet();
    let data = {};

    // 支援 JSON 格式 (text/plain 或 application/json) 與傳統表單參數
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // 產生台灣時區時間戳記
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    // 組裝要寫入的一列資料 (對應 A~I 欄)
    const rowData = [
      timestamp,
      data.name || '',
      data.organization || '',
      data.jobTitle || '',
      data.category || '',
      data.isDirector || '',
      data.attendanceMode || '',
      data.email || '',
      data.phone ? "'" + String(data.phone) : '' // 加上單引號避免電話開頭 0 被吃掉
    ];

    sheet.appendRow(rowData);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: '報名資料已成功匯入「工作坊報名資料」試算表！',
        timestamp: timestamp,
        data: data
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 處理 GET 請求：可讀取目前「工作坊報名資料」分頁的所有報名紀錄供網頁視覺化儀表板同步
 */
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'success',
          count: 0,
          records: []
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const records = rows.slice(1).map(row => ({
      timestamp: row[0],
      name: row[1],
      organization: row[2],
      jobTitle: row[3],
      category: row[4],
      isDirector: row[5],
      attendanceMode: row[6],
      email: row[7],
      phone: row[8]
    }));

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        count: records.length,
        records: records
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
