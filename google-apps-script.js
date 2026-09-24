/**
 * 教材費紀錄系統 — Google Apps Script 雙向同步服務程式碼
 * 
 * 部署教學：
 * 1. 在您的 Google 試算表選單中點選「擴充功能」 -> 「Apps Script」。
 * 2. 清空原本的 `Code.gs` 內容，並將此段程式碼貼入。
 * 3. 點選上方「儲存」按鈕。
 * 4. 點選右上角「部署」 -> 「新增部署」。
 * 5. 類型選擇「網頁應用程式」（Web App）。
 * 6. 設定：
 *    - 說明：教材費系統同步 API
 *    - 執行身分：我（您的 Google 帳號）
 *    - 誰有存取權：所有人（Anyone）— **必須設定為所有人，前端瀏覽器才能呼叫**
 * 7. 點選「部署」，並授予權限。
 * 8. 複製產生的「網頁應用程式 URL」（格式通常為 https://script.google.com/macros/s/.../exec）。
 * 9. 將此網址貼入教材費系統的「試算表備份」設定欄位中。
 */

function doGet(e) {
  try {
    var ss = getActiveSpreadsheetWithFallback(e);
    
    // 解析參數，支援家長端 GET 請求驗證與導師驗證
    var action = e && e.parameter && e.parameter.action ? e.parameter.action : "";
    var auth = e && e.parameter && e.parameter.auth ? String(e.parameter.auth) : "";
    
    // ── 教師密碼驗證（輕量路徑：僅讀取 settings，不讀取名冊與帳本） ──
    if (action === "teacherAuth") {
      var isTeacherAuth = checkTeacherAuth(ss, auth);
      return ContentService.createTextOutput(JSON.stringify({
        success: isTeacherAuth,
        isTeacher: isTeacherAuth,
        error: isTeacherAuth ? null : "密碼錯誤"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ── 家長查詢驗證（優化路徑：名冊僅讀取一次並傳入 parentAuth） ──
    if (action === "parentAuth") {
      var seat = e && e.parameter && e.parameter.seat ? String(e.parameter.seat) : "";
      var pin = e && e.parameter && e.parameter.pin ? String(e.parameter.pin) : "";
      var termParam = e && e.parameter && e.parameter.term ? String(e.parameter.term) : "";
      
      if (!seat || !pin) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "請提供座號與 PIN 碼" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      // 讀取名冊一次，供 parentAuth 內部使用與後續去識別化
      var rosterSheet = getOrCreateSheet(ss, "roster", ["seat", "name"]);
      var rawRoster = readSheetData(rosterSheet);
      
      var authResult = parentAuth(ss, seat, pin, termParam, rawRoster);
      // 對家長端的傳回內容中的備註與項目名稱進行去識別化過濾，確保不洩露其他同學姓名
      if (authResult.success && authResult.data) {
        if (authResult.data.expenses) {
          authResult.data.expenses.forEach(function(t) {
            if (t.note) t.note = maskNamesInText(t.note, rawRoster);
            if (t.item) t.item = maskNamesInText(t.item, rawRoster);
          });
        }
        if (authResult.data.incomes) {
          authResult.data.incomes.forEach(function(t) {
            if (t.note) t.note = maskNamesInText(t.note, rawRoster);
          });
        }
      }
      return ContentService.createTextOutput(JSON.stringify(authResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // ── 一般同步請求（需名冊） ──
    var rosterSheet = getOrCreateSheet(ss, "roster", ["seat", "name"]);
    var rawRoster = readSheetData(rosterSheet);
    
    // 驗證是否為導師身分
    var isTeacher = checkTeacherAuth(ss, auth);
    
    // 讀取第一個工作表作為主帳本
    var sheet = ss.getSheets()[0];
    try {
      setupSheetHeadersAndFormulas(sheet);
    } catch(e) {
      Logger.log("doGet setup headers error: " + e.toString());
    }
    
    var transactions = readAllTransactions(sheet);
    var settingsSheet = getOrCreateSheet(ss, "settings", ["key", "value"]);
    var settings = readSettingsData(settingsSheet);
    
    // 如果不是導師，則強制對名冊與交易明細進行去識別化，且隱藏密碼以維護個資安全
    if (!isTeacher) {
      var anonymizedRoster = rawRoster.map(function(r) {
        return { seat: String(r.seat), name: maskName(r.name) };
      });
      
      var anonymizedTransactions = transactions.map(function(t) {
        return {
          id: t.id,
          type: t.type,
          date: t.date,
          source: t.type === "income" ? maskNamesInText(t.source, rawRoster) : "",
          amount: t.amount,
          seat: t.seat !== undefined && t.seat !== null ? String(t.seat) : "",
          category: t.type === "expense" ? t.category : undefined,
          item: t.type === "expense" ? maskNamesInText(t.item, rawRoster) : undefined,
          payee: t.type === "expense" ? maskNamesInText(t.payee, rawRoster) : "",
          term: t.term,
          note: maskNamesInText(t.note, rawRoster)
        };
      });
      
      // 隱藏設定中的教師密碼
      if (settings.pin) {
        settings.pin = "HIDDEN";
      }
      
      var data = {
        transactions: anonymizedTransactions,
        roster: anonymizedRoster,
        settings: settings
      };
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: data, anonymized: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 導師模式：返回完整包含姓名與細節的明細
    var data = {
      transactions: transactions,
      roster: rawRoster,
      settings: settings
    };
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 去識別化中文姓名（例如：王小明 -> 王○明，李二 -> 李○） ──
function maskName(name) {
  if (!name) return "";
  var len = name.length;
  if (len <= 2) {
    return name.charAt(0) + "○";
  }
  return name.charAt(0) + "○" + name.substring(2);
}

// ── 批次遮罩文字區塊中提及的所有學生姓名，維護個資安全 ──
function maskNamesInText(text, roster) {
  if (!text) return "";
  var masked = text;
  for (var i = 0; i < roster.length; i++) {
    var name = roster[i].name;
    if (name && name.length >= 2) {
      // 避免匹配到單字或重疊
      var re = new RegExp(name, "g");
      masked = masked.replace(re, maskName(name));
    }
  }
  return masked;
}

// ── 驗證導師讀取權限（比對 settings 中的 pin） ──
function checkTeacherAuth(ss, auth) {
  try {
    var settingsSheet = ss.getSheetByName("settings");
    if (!settingsSheet) return true; // 尚未建立 settings 工作表，預設放行

    var settings = readSettingsData(settingsSheet);
    var storedPin = settings.pin;

    if (!storedPin || String(storedPin).trim() === "HIDDEN") return true; // 尚未設定或密碼損毀，預設放行
    if (!auth) return false;

    var storedPinStr = String(storedPin).trim();
    var authStr = String(auth).trim();

    // 判斷試算表中的 PIN 是否已雜湊（SHA-256 hex 為 64 字元）
    var isHashed = (storedPinStr.length === 64 && /^[0-9a-f]+$/.test(storedPinStr));

    if (isHashed) {
      // 已雜湊：將傳入的明碼 auth 雜湊後再比對
      return hashTeacherPin(authStr) === storedPinStr;
    } else {
      // 向下相容：試算表仍為舊版明碼格式
      var plainMatch = (storedPinStr === authStr);
      if (plainMatch) {
        // 自動升級：將明碼 PIN 雜湊後寫回試算表
        try {
          settings.pin = hashTeacherPin(authStr);
          writeSettingsData(settingsSheet, settings);
          Logger.log("[checkTeacherAuth] PIN 已自動升級為雜湊儲存格式");
        } catch (upgradeErr) {
          Logger.log("[checkTeacherAuth] PIN 升級失敗（不影響本次驗證）：" + upgradeErr.toString());
        }
      }
      return plainMatch;
    }
  } catch (e) {
    Logger.log("checkTeacherAuth error: " + e.toString());
  }
  return false;
}

function doPost(e) {
  try {
    var ss = getActiveSpreadsheetWithFallback(e);
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    
    if (action === "teacherAuth") {
      var auth = (postData.auth !== undefined && postData.auth !== null) ? String(postData.auth) : "";
      var isTeacherAuth = checkTeacherAuth(ss, auth);
      return ContentService.createTextOutput(JSON.stringify({
        success: isTeacherAuth,
        isTeacher: isTeacherAuth,
        error: isTeacherAuth ? null : "密碼錯誤"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "sync" || action === "push") {
      var auth = (postData.auth !== undefined && postData.auth !== null) ? String(postData.auth) : "";
      if (!checkTeacherAuth(ss, auth)) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "驗證失敗，無權限寫入試算表" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var sheet = ss.getSheets()[0];
      try {
        setupSheetHeadersAndFormulas(sheet);
      } catch(e) {
        Logger.log("doPost setup headers error: " + e.toString());
      }
      
      // 1. 寫入交易紀錄（分流寫入主帳本，採局部區塊寫入，避免互相覆蓋欄位）
      if (postData.transactions) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 5) {
          // 清空第 5 列以下的所有資料（涵蓋至 S 欄 UUID 欄）
          sheet.getRange(5, 1, lastRow - 4, 19).clearContent();
        }
        
        var transList = postData.transactions;
        
        // 分流：處理收入
        var incomes = transList.filter(function(t) { return t.type === "income"; });
        if (incomes.length > 0) {
          var incomeRowsMain = [];
          var incomeRowsMeta = [];
          for (var i = 0; i < incomes.length; i++) {
            var t = incomes[i];
            // A: 日期, B: 內容(來源), C: 金額, D: 座號 (4 欄)
            incomeRowsMain.push([
              t.date || "",
              t.source || "",
              t.amount || 0,
              t.seat !== undefined && t.seat !== null ? String(t.seat) : ""
            ]);
            // N: 學期, O: 備註 (2 欄)
            incomeRowsMeta.push([
              t.term || "",
              t.note || ""
            ]);
          }
          sheet.getRange(5, 1, incomeRowsMain.length, 4).setValues(incomeRowsMain);
          sheet.getRange(5, 14, incomeRowsMeta.length, 2).setValues(incomeRowsMeta);
          
          // 同時寫入收入交易的 UUID 至 R 欄（第 18 欄）
          var incomeUUIDs = incomes.map(function(t) { return [t.id || '']; });
          sheet.getRange(5, 18, incomeUUIDs.length, 1).setValues(incomeUUIDs);
        }
        
        // 分流：處理支出
        var expenses = transList.filter(function(t) { return t.type === "expense"; });
        if (expenses.length > 0) {
          var expenseRowsMain = [];
          var expenseRowsMeta = [];
          for (var i = 0; i < expenses.length; i++) {
            var t = expenses[i];
            // F: 日期 (Col 6), G: 類別, H: 項目, I: 單價, J: 數量, K: 金額, L: 取款人, M: 座號 (8 欄)
            expenseRowsMain.push([
              t.date || "",
              t.category || "",
              t.item || "",
              t.unitPrice || 0,
              t.qty || 0,
              t.amount || 0,
              t.payee !== undefined && t.payee !== null ? String(t.payee) : "",
              t.seat !== undefined && t.seat !== null ? String(t.seat) : ""
            ]);
            // P: 學期 (Col 16), Q: 備註 (Col 17) (2 欄)
            expenseRowsMeta.push([
              t.term || "",
              t.note || ""
            ]);
          }
          sheet.getRange(5, 6, expenseRowsMain.length, 8).setValues(expenseRowsMain);
          sheet.getRange(5, 16, expenseRowsMeta.length, 2).setValues(expenseRowsMeta);
          
          // 同時寫入支出交易的 UUID 至 S 欄（第 19 欄）
          var expenseUUIDs = expenses.map(function(t) { return [t.id || '']; });
          sheet.getRange(5, 19, expenseUUIDs.length, 1).setValues(expenseUUIDs);
        }
      }
      
      // 2. 寫入名冊（背景額外工作表）
      if (postData.roster) {
        var rosterSheet = getOrCreateSheet(ss, "roster", ["seat", "name"]);
        writeSheetData(rosterSheet, postData.roster, ["seat", "name"]);
      }
      
      // 3. 寫入設定（背景額外工作表）
      if (postData.settings) {
        var settingsSheet = getOrCreateSheet(ss, "settings", ["key", "value"]);
        var existingSettings = readSettingsData(settingsSheet);
        
        // 若前端傳來 HIDDEN，保留試算表中現有的 PIN，不覆寫
        if (postData.settings.pin === "HIDDEN" && existingSettings.pin) {
          postData.settings.pin = existingSettings.pin;
        }

        // 若前端傳來明碼 PIN（非 HIDDEN、非空、非已雜湊的 64 字元 hex），將其雜湊後再寫入
        var incomingPin = postData.settings.pin;
        if (incomingPin && incomingPin !== "HIDDEN") {
          var incomingPinStr = String(incomingPin).trim();
          var alreadyHashed = (incomingPinStr.length === 64 && /^[0-9a-f]+$/.test(incomingPinStr));
          if (!alreadyHashed) {
            postData.settings.pin = hashTeacherPin(incomingPinStr);
          }
        }

        writeSettingsData(settingsSheet, postData.settings);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Sync successful" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "resetParentPin") {
      var auth = (postData.auth !== undefined && postData.auth !== null) ? String(postData.auth) : "";
      if (!checkTeacherAuth(ss, auth)) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "驗證失敗，無權限重設密碼" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      var seat = postData.seat !== undefined && postData.seat !== null ? String(postData.seat) : "";
      if (!seat) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "請提供欲重設之座號" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      var resetSuccess = deleteCredentialBySeat(ss, seat);
      return ContentService.createTextOutput(JSON.stringify({ success: true, reset: resetSuccess }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "parentAuth") {
      var seat = postData.seat !== undefined && postData.seat !== null ? String(postData.seat) : "";
      var pin = postData.pin !== undefined && postData.pin !== null ? String(postData.pin) : "";
      var termParam = postData.term || "";

      if (!seat || !pin) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "請提供座號與 PIN 碼" }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var authResult = parentAuth(ss, seat, pin, termParam);
      return ContentService.createTextOutput(JSON.stringify(authResult))
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

// 輔助函式：刪除特定座號的家長 credentials 紀錄進行密碼重設
function deleteCredentialBySeat(ss, seat) {
  var credSheet = getOrCreateSheet(ss, "credentials", ["seat", "pinHash", "updatedAt"]);
  var credentials = readSheetData(credSheet);
  var nextCreds = [];
  var found = false;
  for (var i = 0; i < credentials.length; i++) {
    if (String(credentials[i].seat) !== String(seat)) {
      nextCreds.push(credentials[i]);
    } else {
      found = true;
    }
  }
  if (found) {
    writeSheetData(credSheet, nextCreds, ["seat", "pinHash", "updatedAt"]);
  }
  return found;
}

// ===================================================================
// 家長登入 / 繳費狀態查詢（parentAuth）
// ===================================================================

// 主流程：驗證座號 + PIN，首次登入自動建立 PIN，成功後回傳去識別化的家長視角資料
function parentAuth(ss, seat, pin, termParam, cachedRoster) {
  var credSheet = getOrCreateSheet(ss, "credentials", ["seat", "pinHash", "updatedAt"]);
  var credentials = readSheetData(credSheet); // [{seat, pinHash, updatedAt}, ...]

  var existing = null;
  for (var i = 0; i < credentials.length; i++) {
    if (String(credentials[i].seat) === seat) {
      existing = credentials[i];
      break;
    }
  }

  var pinHash = hashPin(seat, pin);
  var isNewSetup = false;

  if (existing) {
    if (String(existing.pinHash) !== pinHash) {
      return { success: false, error: "PIN 碼錯誤" };
    }
  } else {
    // 首次登入：建立新的 PIN 紀錄
    var allCreds = credentials.concat([{
      seat: seat,
      pinHash: pinHash,
      updatedAt: formatJSDate(new Date())
    }]);
    writeSheetData(credSheet, allCreds, ["seat", "pinHash", "updatedAt"]);
    isNewSetup = true;
  }

  // 確認該座號存在於名冊（優先使用已快取的名冊，避免重複讀取）
  var roster = cachedRoster;
  if (!roster) {
    var rosterSheet = getOrCreateSheet(ss, "roster", ["seat", "name"]);
    roster = readSheetData(rosterSheet);
  }
  var studentEntry = null;
  for (var j = 0; j < roster.length; j++) {
    if (String(roster[j].seat) === seat) {
      studentEntry = roster[j];
      break;
    }
  }
  if (!studentEntry) {
    return { success: false, error: "查無此座號，請確認座號是否正確" };
  }

  // 讀取交易資料（沿用 doGet 的讀取邏輯）
  var sheet = ss.getSheets()[0];
  var transactions = readAllTransactions(sheet);

  // 決定查詢學期：優先使用前端傳入的 term，否則取交易紀錄中最新出現的學期
  var term = termParam;
  if (!term) {
    for (var k = transactions.length - 1; k >= 0; k--) {
      if (transactions[k].term) {
        term = transactions[k].term;
        break;
      }
    }
  }

  // 應繳金額設定：settings 裡的 duesConfig，格式為 { "學期字串": 金額 }
  var settingsSheet = getOrCreateSheet(ss, "settings", ["key", "value"]);
  var settings = readSettingsData(settingsSheet);
  var duesConfig = settings.duesConfig || {};
  var amountDue = Number(duesConfig[term]) || 0;

  // 加總該座號、該學期的所有 income 交易（個人預繳/儲值）
  var studentIncomes = [];
  var amountPaid = 0;
  for (var m = 0; m < transactions.length; m++) {
    var t = transactions[m];
    if (t.type === "income" && String(t.seat) === seat && (!term || t.term === term)) {
      amountPaid += Number(t.amount) || 0;
      studentIncomes.push({
        id: t.id,
        date: t.date,
        amount: Number(t.amount) || 0,
        term: t.term,
        note: t.note || ""
      });
    }
  }

  // 加總該座號、該學期的所有 expense 交易（個人教材扣款）
  var studentExpenses = [];
  var amountSpent = 0;
  for (var n = 0; n < transactions.length; n++) {
    var exp = transactions[n];
    if (exp.type === "expense" && String(exp.seat) === seat && (!term || exp.term === term)) {
      amountSpent += Number(exp.amount) || 0;
      studentExpenses.push({
        id: exp.id,
        date: exp.date,
        category: exp.category || "教材費",
        item: exp.item || "",
        amount: Number(exp.amount) || 0,
        term: exp.term,
        note: exp.note || ""
      });
    }
  }

  var studentBalance = amountPaid - amountSpent;

  var status = "unpaid";
  if (amountDue > 0) {
    if (amountPaid >= amountDue) status = "paid";
    else if (amountPaid > 0) status = "partial";
  } else if (amountPaid > 0) {
    status = "paid"; // 尚未設定金額但已有繳費紀錄，視為已繳
  }

  return {
    success: true,
    isNewSetup: isNewSetup,
    data: {
      myChild: {
        seat: seat,
        name: studentEntry.name,
        payment: {
          seat: seat,
          term: term,
          amountDue: amountDue,
          amountPaid: amountPaid,
          amountSpent: amountSpent,
          balance: studentBalance,
          status: status
        }
      },
      incomes: studentIncomes,
      expenses: studentExpenses,
      balance: studentBalance
    }
  };
}

// 將 doGet 裡讀取交易的邏輯抽成共用函式，供 parentAuth 使用
function readAllTransactions(sheet) {
  var transactions = [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 5) return transactions;

  var incomeData = sheet.getRange(5, 1, lastRow - 4, 18).getValues();
  var incomeDisplay = sheet.getRange(5, 1, lastRow - 4, 1).getDisplayValues();
  for (var i = 0; i < incomeData.length; i++) {
    var row = incomeData[i];
    var dateVal = formatJSDate(incomeDisplay[i][0]);
    if (dateVal) {
      // row[17] 對應第 18 欄（R 欄），優先使用已持久化的 UUID
      var incomeUUID = (row[17] && String(row[17]).trim()) ? String(row[17]).trim() : ("income_" + i);
      transactions.push({
        id: incomeUUID,
        type: "income",
        date: dateVal,
        source: row[1] || "",
        amount: Number(row[2]) || 0,
        seat: row[3] !== undefined && row[3] !== "" ? String(row[3]) : "",
        term: row[13] || "",
        note: row[14] || ""
      });
    }
  }

  // 從 F 欄（第 6 欄）延伸讀取至 S 欄（第 19 欄），共 14 欄
  var expenseData = sheet.getRange(5, 6, lastRow - 4, 14).getValues();
  var expenseDisplay = sheet.getRange(5, 6, lastRow - 4, 1).getDisplayValues();
  for (var i = 0; i < expenseData.length; i++) {
    var row = expenseData[i];
    var dateVal = formatJSDate(expenseDisplay[i][0]);
    if (dateVal) {
      // row[13] 對應第 19 欄（S 欄，支出 UUID）
      var expenseUUID = (row[13] && String(row[13]).trim()) ? String(row[13]).trim() : ("expense_" + i);
      transactions.push({
        id: expenseUUID,
        type: "expense",
        date: dateVal,
        category: row[1] || "",
        item: row[2] || "",
        unitPrice: Number(row[3]) || 0,
        qty: Number(row[4]) || 0,
        amount: Number(row[5]) || 0,
        payee: row[6] !== undefined && row[6] !== "" ? String(row[6]) : "",
        seat: row[7] !== undefined && row[7] !== "" ? String(row[7]) : "",
        term: row[10] || "",
        note: row[11] || ""
      });
    }
  }

  return transactions;
}

// 將 PIN 碼以 SHA-256 雜湊，混入座號作為簡易 salt，避免明碼存於試算表
function hashPin(seat, pin) {
  var raw = "cls-fund-salt:" + seat + ":" + pin;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  var hex = "";
  for (var i = 0; i < digest.length; i++) {
    var byte = (digest[i] + 256) % 256;
    hex += ("0" + byte.toString(16)).slice(-2);
  }
  return hex;
}

// 將導師 PIN 碼以 SHA-256 雜湊儲存，避免明碼存於試算表
function hashTeacherPin(pin) {
  var raw = 'cls-fund-teacher-salt:' + pin;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < digest.length; i++) {
    var b = (digest[i] + 256) % 256;
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return hex;
}

// 取得活動中試算表，若為獨立 standalone 部署則從 URL 參數或 postBody 讀取對應的試算表
function getActiveSpreadsheetWithFallback(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  
  var url = "";
  if (e && e.parameter && e.parameter.url) {
    url = e.parameter.url;
  }
  
  if (!url && e && e.postData && e.postData.contents) {
    try {
      var postData = JSON.parse(e.postData.contents);
      if (postData.sheetUrl) url = postData.sheetUrl;
    } catch(err) {}
  }
  
  if (url) {
    return SpreadsheetApp.openByUrl(url);
  }
  
  throw new Error("找不到試算表！若您使用的是獨立版 Apps Script，請於設定中確定已填寫試算表網址。");
}
// 輔助函式：將 Google 試算表回傳之 Date 物件或字串，統一格式化為 YYYY-MM-DD 字串
function formatJSDate(val) {
  if (!val) return "";
  if (val && (val instanceof Date || typeof val.getFullYear === "function" || Object.prototype.toString.call(val) === "[object Date]")) {
    var y = val.getFullYear();
    var m = ("0" + (val.getMonth() + 1)).slice(-2);
    var d = ("0" + val.getDate()).slice(-2);
    return y + "-" + m + "-" + d;
  }
  var str = String(val).trim();
  if (str.indexOf("T") !== -1) {
    str = str.split("T")[0];
  }
  str = str.replace(/\//g, "-");
  return str;
}

// 輔助函式：自動建立並補齊第一個工作表的欄位標頭、總額公式與排版格式
// 使用「只初始化一次」防衛邏輯：若 A1 已有內容則直接跳過，避免每次同步都呼叫大量 Sheets API
function setupSheetHeadersAndFormulas(sheet) {
  try {
    // 檢查 A1 是否已有值，有值代表已初始化，直接略過所有設定
    var existingA1 = sheet.getRange("A1").getValue();
    if (existingA1 && String(existingA1).trim() !== "") {
      return;
    }

    // ── 以下只有在試算表首次建立時才會執行 ──

    // 1. 大標題
    var rA1 = sheet.getRange("A1");
    rA1.setValue("導師班級教材費專戶紀錄表");
    rA1.setFontWeight("bold");
    rA1.setFontSize(13);

    // 2. 統計公式
    var rA2 = sheet.getRange("A2");
    rA2.setValue("總收入：");
    rA2.setFontWeight("bold");

    var rB2 = sheet.getRange("B2");
    rB2.setFormula("=SUM(C5:C)");
    rB2.setFontWeight("bold");

    var rC2 = sheet.getRange("C2");
    rC2.setValue("總支出：");
    rC2.setFontWeight("bold");

    var rD2 = sheet.getRange("D2");
    rD2.setFormula("=SUM(K5:K)");
    rD2.setFontWeight("bold");

    var rF2 = sheet.getRange("F2");
    rF2.setValue("教材費結餘：");
    rF2.setFontWeight("bold");

    var rG2 = sheet.getRange("G2");
    rG2.setFormula("=B2-D2");
    rG2.setFontWeight("bold");
    rG2.setFontColor("#2d6a4f");

    // 3. 第 4 列欄位標頭
    var rIncHead = sheet.getRange(4, 1, 1, 4);
    rIncHead.setValues([["日期", "來源/項目", "金額", "座號"]]);
    rIncHead.setFontWeight("bold");
    rIncHead.setBackground("#e8f5e9");

    var rSep = sheet.getRange(4, 5);
    rSep.setValue(" ");
    rSep.setBackground("#f5f5f5");

    var rExpHead = sheet.getRange(4, 6, 1, 8);
    rExpHead.setValues([["日期", "類別", "項目", "單價", "數量", "金額", "取款人", "座號"]]);
    rExpHead.setFontWeight("bold");
    rExpHead.setBackground("#ffebee");

    sheet.getRange(4, 14).setValue("學期 (收)").setFontWeight("bold").setBackground("#e8f5e9");
    sheet.getRange(4, 15).setValue("備註 (收)").setFontWeight("bold").setBackground("#e8f5e9");
    sheet.getRange(4, 16).setValue("學期 (支)").setFontWeight("bold").setBackground("#ffebee");
    sheet.getRange(4, 17).setValue("備註 (支)").setFontWeight("bold").setBackground("#ffebee");
    sheet.getRange(4, 18).setValue("UUID (收)").setFontWeight("bold").setBackground("#e3f2fd");
    sheet.getRange(4, 19).setValue("UUID (支)").setFontWeight("bold").setBackground("#fce4ec");

  } catch (err) {
    Logger.log("setupSheetHeadersAndFormulas error: " + err.toString());
  }
}