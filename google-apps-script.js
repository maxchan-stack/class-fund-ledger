/**
 * 班費紀錄系統 — Google Apps Script 雙向同步服務程式碼
 * 
 * 部署教學：
 * 1. 在您的 Google 試算表選單中點選「擴充功能」 -> 「Apps Script」。
 * 2. 清空原本的 `Code.gs` 內容，並將此段程式碼貼入。
 * 3. 點選上方「儲存」按鈕。
 * 4. 點選右上角「部署」 -> 「新增部署」。
 * 5. 類型選擇「網頁應用程式」（Web App）。
 * 6. 設定：
 *    - 說明：班費系統同步 API
 *    - 執行身分：我（您的 Google 帳號）
 *    - 誰有存取權：所有人（Anyone）— **必須設定為所有人，前端瀏覽器才能呼叫**
 * 7. 點選「部署」，並授予權限。
 * 8. 複製產生的「網頁應用程式 URL」（格式通常為 https://script.google.com/macros/s/.../exec）。
 * 9. 將此網址貼入班費系統的「試算表備份」設定欄位中。
 */

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var transSheet = getOrCreateSheet(ss, "transactions", ["id", "type", "date", "category", "item", "unitPrice", "qty", "amount", "payee", "term", "note"]);
  var rosterSheet = getOrCreateSheet(ss, "roster", ["seat", "name"]);
  var settingsSheet = getOrCreateSheet(ss, "settings", ["key", "value"]);
  
  var data = {
    transactions: readSheetData(transSheet),
    roster: readSheetData(rosterSheet),
    settings: readSettingsData(settingsSheet)
  };
  
  return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    
    if (action === "sync" || action === "push") {
      if (postData.transactions) {
        var transSheet = getOrCreateSheet(ss, "transactions", ["id", "type", "date", "category", "item", "unitPrice", "qty", "amount", "payee", "term", "note"]);
        writeSheetData(transSheet, postData.transactions, ["id", "type", "date", "category", "item", "unitPrice", "qty", "amount", "payee", "term", "note"]);
      }
      if (postData.roster) {
        var rosterSheet = getOrCreateSheet(ss, "roster", ["seat", "name"]);
        writeSheetData(rosterSheet, postData.roster, ["seat", "name"]);
      }
      if (postData.settings) {
        var settingsSheet = getOrCreateSheet(ss, "settings", ["key", "value"]);
        writeSettingsData(settingsSheet, postData.settings);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Sync successful" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Invalid action" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 輔助函式：取得或建立 Sheet，若新建則自動寫入標頭
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1); // 凍結首行
  }
  return sheet;
}

// 讀取一般 Sheet 的資料轉為物件陣列
function readSheetData(sheet) {
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  var headers = rows[0];
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var item = {};
    var hasContent = false;
    for (var j = 0; j < headers.length; j++) {
      item[headers[j]] = row[j];
      if (row[j] !== "") hasContent = true;
    }
    if (hasContent) {
      // 數值型態欄位還原為數字
      if (item.amount !== undefined) item.amount = Number(item.amount) || 0;
      if (item.unitPrice !== undefined) item.unitPrice = Number(item.unitPrice) || 0;
      if (item.qty !== undefined) item.qty = Number(item.qty) || 0;
      data.push(item);
    }
  }
  return data;
}

// 讀取設定頁面的資料
function readSettingsData(sheet) {
  var rows = sheet.getDataRange().getValues();
  var settings = {};
  if (rows.length <= 1) return settings;
  for (var i = 1; i < rows.length; i++) {
    var key = rows[i][0];
    var val = rows[i][1];
    if (key) {
      try {
        settings[key] = JSON.parse(val);
      } catch (e) {
        settings[key] = val;
      }
    }
  }
  return settings;
}

// 將物件陣列寫入 Sheet（覆蓋寫入）
function writeSheetData(sheet, dataList, headers) {
  sheet.clearContents();
  sheet.appendRow(headers);
  if (!dataList || dataList.length === 0) return;
  var rows = [];
  for (var i = 0; i < dataList.length; i++) {
    var item = dataList[i];
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      var val = item[headers[j]];
      row.push(val === undefined || val === null ? "" : val);
    }
    rows.push(row);
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// 將設定寫入 Sheet
function writeSettingsData(sheet, settingsObj) {
  sheet.clearContents();
  sheet.appendRow(["key", "value"]);
  var keys = Object.keys(settingsObj);
  if (keys.length === 0) return;
  var rows = [];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var val = settingsObj[key];
    if (typeof val === "object") {
      val = JSON.stringify(val);
    }
    rows.push([key, val]);
  }
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}
