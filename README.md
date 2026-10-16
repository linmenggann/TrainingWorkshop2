# 🎓 教學訓練計畫主持人工作坊 — 視覺化報名網頁與 Google Sheets 串接

## 📌 目標 Google Sheets 資訊
- **試算表網址**：[開啟 Google Sheets 試算表](https://docs.google.com/spreadsheets/d/1DXp8rcOjWTABT6Dg2Q6zdTXD7lhIay7DP2TNcGsBrto/edit?gid=0#gid=0)
- **Spreadsheet ID**：`1DXp8rcOjWTABT6Dg2Q6zdTXD7lhIay7DP2TNcGsBrto`
- **分頁名稱**：`工作坊報名資料`

---

## 📋 一、Google Sheets 表頭（放在「工作坊報名資料」分頁第 1 列 `A1:I1`）

您可以直接複製下方整行文字（含 Tab 分隔），在 Google Sheets「**工作坊報名資料**」分頁點選 **`A1` 儲存格** 貼上，即可自動填入 `A1` 到 `I1`：

```text
報名時間	姓名	機構名	職稱	負責的職類	是否擔任教學訓練計畫主持人	參與方式	Email	聯繫電話
```

| 儲存格 | 表頭名稱 | 對應網頁欄位 (`name`) | 範例內容 |
| :---: | :--- | :--- | :--- |
| **A1** | `報名時間` | `timestamp` (系統自動產生) | `2026/01/25 09:00:00` |
| **B1** | `姓名` | `name` | `王小明` |
| **C1** | `機構名` | `organization` | `奇美醫院` |
| **D1** | `職稱` | `jobTitle` | `護理長` |
| **E1** | `負責的職類` | `category` | `護理` |
| **F1** | `是否擔任教學訓練計畫主持人` | `isDirector` | `是` |
| **G1** | `參與方式` | `attendanceMode` | `實體` |
| **H1** | `Email` | `email` | `example@mail.chimei.org.tw` |
| **I1** | `聯繫電話` | `phone` | `0912-345-678` |

> 💡 **小提示**：若您使用專案內的 `Code.gs`，當試算表尚未建立「工作坊報名資料」分頁或第一列空白時，程式也會**自動建立分頁並填入紫色表頭**！

---

## ⚙️ 二、Google Apps Script 部署步驟 (`Code.gs`)

1. 開啟您的 [Google Sheets 試算表](https://docs.google.com/spreadsheets/d/1DXp8rcOjWTABT6Dg2Q6zdTXD7lhIay7DP2TNcGsBrto/edit?gid=0#gid=0)。
2. 點選上方選單 **「擴充功能 (Extensions)」** ➔ **「Apps Script」**。
3. 將編輯器內預設的程式碼清空，貼上本專案 [`Code.gs`](./Code.gs) 的完整程式碼並按 `Ctrl + S` 儲存。
4. （選用）在上方函式下拉選單選擇 `setupSheet`，點選 **「執行」** 並完成授權，即可自動建立「`工作坊報名資料`」分頁與表頭。
5. 點選右上角 **「部署 (Deploy)」** ➔ **「新增部署作業 (New deployment)」**：
   - **選取類型**：點選齒輪圖示 ➔ **「網頁應用程式 (Web app)」**
   - **執行身分 (Execute as)**：**「我 (Me)」**
   - **誰可以存取 (Who has access)**：**「所有人 (Anyone)」**
6. 點選 **「部署」** 後，複製產生的 **「網頁應用程式網址 (`https://script.google.com/macros/s/.../exec`)」**。
7. 回到 `index.html` 網頁下方的 **「☁️ Google Sheets 串接與表頭設定」** 區塊，貼上該網址並點選 **「💾 儲存串接網址」**（或直接填入 `index.html` 內的 `DEFAULT_GAS_WEB_APP_URL` 變數），即可開始將報名資料即時匯入 Google Sheets！