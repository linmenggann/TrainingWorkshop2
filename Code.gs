/**
 * Google Apps Script - 教學訓練計畫主持人工作坊報名表單
 *
 * 【使用說明】
 * 1. 開啟 Google Sheets: https://docs.google.com/spreadsheets/d/1rFbD11t6LRkYgzNzlL41EapiFe5RZ4TheJZdfnGecpE/edit
 * 2. 將第一個分頁名稱改為「工作坊報名資料」
 * 3. 在第一列（表頭）依序填入：
 *    A1: 時間戳記
 *    B1: 姓名
 *    C1: 機構名
 *    D1: 職稱
 *    E1: 負責的職類
 *    F1: 是否擔任教學訓練計畫主持人
 *    G1: 參與方式
 *    H1: Email
 *    I1: 聯繫電話
 * 4. 點選上方選單「擴充功能」>「Apps Script」
 * 5. 將此檔案內容貼入編輯器中（取代預設的 myFunction）
 * 6. 點選「部署」>「新增部署作業」
 *    - 類型選「網頁應用程式」
 *    - 執行身分：「我」
 *    - 誰可以存取：「所有人」
 * 7. 點選「部署」，授權後取得網址
 * 8. 將網址貼到 index.html 中的 APPS_SCRIPT_URL 變數
 */

// 報名人數上限
var MAX_CAPACITY = 4;

// 處理 POST 請求（表單提交）
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("工作坊報名資料");

    if (!sheet) {
      return createResponse(false, "找不到「工作坊報名資料」分頁");
    }

    // 檢查是否已額滿
    var currentCount = Math.max(0, sheet.getLastRow() - 1);
    if (currentCount >= MAX_CAPACITY) {
      return createResponse(false, "報名已額滿，無法受理新報名。");
    }

    var data = JSON.parse(e.postData.contents);

    var timestamp = new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" });

    sheet.appendRow([
      timestamp,
      data.name       || "",
      data.org        || "",
      data.title      || "",
      data.category   || "",
      data.isHost     || "",
      data.joinMethod || "",
      data.email      || "",
      data.phone      || ""
    ]);

    return createResponse(true, "報名成功！");
  } catch (error) {
    return createResponse(false, "發生錯誤：" + error.message);
  }
}

// 處理 GET 請求（回傳報名資料供儀表板使用）
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("工作坊報名資料");

    if (!sheet) {
      return createResponse(false, "找不到「工作坊報名資料」分頁");
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      var output = JSON.stringify({ success: true, data: [], headers: [] });
      return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
    }

    var headers = sheet.getRange(1, 1, 1, 9).getValues()[0];
    var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    var data = rows.map(function(row) {
      return {
        timestamp:  row[0],
        name:       row[1],
        org:        row[2],
        title:      row[3],
        category:   row[4],
        isHost:     row[5],
        joinMethod: row[6],
        email:      row[7],
        phone:      row[8]
      };
    });

    var isFull = data.length >= MAX_CAPACITY;
    var output = JSON.stringify({ success: true, data: data, total: data.length, maxCapacity: MAX_CAPACITY, isFull: isFull });
    return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return createResponse(false, "發生錯誤：" + error.message);
  }
}

// 建立 JSON 回應（支援 CORS）
function createResponse(success, message) {
  var output = JSON.stringify({ success: success, message: message });
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
