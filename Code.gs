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

// 處理 POST 請求（表單提交）
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("工作坊報名資料");

    if (!sheet) {
      return createResponse(false, "找不到「工作坊報名資料」分頁");
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

// 處理 GET 請求（測試用）
function doGet(e) {
  return createResponse(true, "Apps Script 已就緒！");
}

// 建立 JSON 回應（支援 CORS）
function createResponse(success, message) {
  var output = JSON.stringify({ success: success, message: message });
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
