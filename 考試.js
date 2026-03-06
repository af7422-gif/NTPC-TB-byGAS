  // 批次更新題目出題/答對次數
  function updateQuestionStatsBatch(updates, unit) {
    // 🚫 如果單位是衛生局 → 直接跳過（維持你原本邏輯）
    if (unit === "衛生局") return;

    if (!updates || updates.length === 0) return;

    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // 最多等 10 秒，避免併發衝突

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("題庫");
      if (!sheet) return;

      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const idxAsked    = headers.indexOf("出題次數");
      const idxCorrect  = headers.indexOf("答對次數");
      const idxAccuracy = headers.indexOf("答對率");

      if (idxAsked < 0 || idxCorrect < 0 || idxAccuracy < 0) return;

      updates.forEach(u => {
        if (!u || typeof u.serialNo !== "number") return;

        // ⚠️ 序號 = 1 → 第 2 列
        const row = u.serialNo + 1;

        const askedCell   = sheet.getRange(row, idxAsked + 1);
        const correctCell = sheet.getRange(row, idxCorrect + 1);

        let asked   = Number(askedCell.getValue())   || 0;
        let correct = Number(correctCell.getValue()) || 0;

        if (u.asked)   asked++;
        if (u.correct) correct++;

        const accuracy = asked > 0 ? correct / asked : 0;

        // 一次寫回（仍在 lock 內）
        sheet
          .getRange(row, idxAsked + 1, 1, 3)
          .setValues([[asked, correct, accuracy]]);
      });

    } finally {
      // ✅ 一定釋放鎖，避免整個系統卡死
      lock.releaseLock();
    }
  }

  //紀錄帳號答題情形
  function updateExamStats(userAccount, action) {
    if (!userAccount) return;

    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // 最多等 10 秒，避免併發衝突

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const accSheet = ss.getSheetByName("帳號資訊");
      if (!accSheet) return;

      const accData = accSheet.getDataRange().getValues();
      if (accData.length < 2) return;

      const accHeaders = accData[0];
      const idxAcc = accHeaders.indexOf("公務帳號");
      const idxAns = accHeaders.indexOf("答題次數");   // 實際上是「出卷次數」
      const idxPass = accHeaders.indexOf("通過次數");
      const idxRate = accHeaders.indexOf("合格率");
      const idxLastLogin = accHeaders.indexOf("上次登入時間");

      if (idxAcc < 0 || idxAns < 0 || idxPass < 0 || idxRate < 0) return;

      const targetAcc = String(userAccount).trim().toLowerCase();

      for (let r = 1; r < accData.length; r++) {
        const acc = String(accData[r][idxAcc]).trim().toLowerCase();
        if (acc !== targetAcc) continue;

        let totalAns  = Number(accData[r][idxAns])  || 0;
        let totalPass = Number(accData[r][idxPass]) || 0;

        if (action === "issued") {
          totalAns++;     // 系統已出題一次
        }

        if (action === "passed") {
          totalPass++;    // 考卷通過一次

          // ✅ 更新上次登入時間（仍在 lock 內）
          if (idxLastLogin >= 0) {
            accSheet.getRange(r + 1, idxLastLogin + 1).setValue(new Date());
          }
        }

        // ✅ 寫回統計數
        accSheet.getRange(r + 1, idxAns + 1).setValue(totalAns);
        accSheet.getRange(r + 1, idxPass + 1).setValue(totalPass);

        // ✅ 重算合格率
        const rate = totalAns > 0 ? totalPass / totalAns : 0;
        accSheet.getRange(r + 1, idxRate + 1).setValue(rate);

        break; // 👈 找到帳號就結束
      }

    } finally {
      // ✅ 一定釋放鎖，避免整個系統被卡住
      lock.releaseLock();
    }
  }

  //因為快速通關維護得上次通過時間
  function updateLastLogin(userAccount) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("帳號資訊");
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxAcc = headers.indexOf("公務帳號");
    const idxLastLogin = headers.indexOf("上次登入時間");

    if (idxAcc < 0 || idxLastLogin < 0) return;

    for (let r = 1; r < data.length; r++) {
      if (String(data[r][idxAcc]).trim().toLowerCase() === String(userAccount).trim().toLowerCase()) {
        sheet.getRange(r+1, idxLastLogin+1).setValue(new Date());  // 寫入現在時間
        break;
      }
    }
  }

  /**
   * 題庫專用：上傳 base64 圖片至「試算表名稱_上傳圖片/題庫_參考圖」資料夾
   * 並回傳可直接嵌入的 Google Drive 圖片網址
   * @param {string} base64 - data:image/png;base64,... 或 data:image/jpeg;base64,...
   * @param {string} folderName - 子資料夾名稱（建議固定為 "題庫_參考圖"）
   * @param {string} filename - 題目序號或題目文字，將用於命名
   * @return {string} 可直接顯示的圖片網址
   */
  function uploadQuestionImageBase64(base64, folderName, filename) {
    if (!base64 || typeof base64 !== "string" || !base64.includes(',')) {
      throw new Error("無效的 base64 圖片資料");
    }

    // === 判斷格式與副檔名 ===
    const contentType = base64.startsWith('data:image/jpeg') ? 'image/jpeg' : 'image/png';
    const ext = contentType === 'image/jpeg' ? '.jpg' : '.png';
    const name = filename ? filename + ext : 'question_image' + ext;

    // === 轉成 Blob ===
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64.split(',')[1]),
      contentType,
      name
    );

    // === 找出目前試算表的主資料夾 ===
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();

    // === 建立主資料夾（試算表名稱_上傳圖片） ===
    const mainFolderName = ss.getName() + "_上傳圖片";
    let mainFolder;
    const mainFolders = parentFolder.getFoldersByName(mainFolderName);
    mainFolder = mainFolders.hasNext() ? mainFolders.next() : parentFolder.createFolder(mainFolderName);

    // === 建立題庫專屬子資料夾（folderName，如「題庫_參考圖」） ===
    let targetFolder;
    const subFolders = mainFolder.getFoldersByName(folderName);
    targetFolder = subFolders.hasNext() ? subFolders.next() : mainFolder.createFolder(folderName);

    // === 建立檔案與設定分享權限 ===
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ✅ 回傳可直接嵌入顯示的圖片網址（Drive thumbnail）
    return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1600`;
  }


  //將題目寫進題庫
  function addQuestion(newQ) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("題庫");
    if (!sheet) throw new Error("❌ 找不到題庫工作表");

    // 讀取表頭
    const headers = sheet.getDataRange().getValues()[0];

    // 找「序號」欄位
    const idxSerial = headers.indexOf("序號");
    if (idxSerial < 0) throw new Error("❌ 題庫缺少『序號』欄");

    // 🔹 序號 = 當前最後一列 index - 1
    const serialNo = sheet.getLastRow();   

    // 預設的審核資訊 JSON
    const defaultReviewInfo = {
      reviewers: {}
    };

    // 組成一列資料（依照 headers 順序）
    const row = headers.map(h => {
      if (h === "序號") return serialNo;
      if (h === "啟用狀態") return "✏️ 待審核"; 
      if (h === "審核資訊") return newQ[h] || JSON.stringify(defaultReviewInfo);
      return (newQ[h] !== undefined) ? newQ[h] : "";
    });

    // 寫入
    sheet.appendRow(row);

    // 回傳：這一列完整資料
    return row;
  }



  //紀錄作答資訊
  function saveExamRecord(record) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // 最多等 10 秒，避免併發衝突

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("答題紀錄");
      if (!sheet) return;

      // === 格式化提交時間 ===
      const ts = Utilities.formatDate(
        new Date(record.提交時間),
        Session.getScriptTimeZone(),
        "yyyy/MM/dd HH:mm:ss"
      );

      // === 取出現有表頭（⚠️ 一定要在 lock 裡） ===
      let headers = sheet.getDataRange().getValues()[0] || [];

      // 確保有基本欄位（新增「公務帳號」）
      const baseCols = [
        "提交時間",
        "公務帳號",
        "單位",
        "用戶名稱",
        "作答題數",
        "答對題數",
        "答題內容",
        "是否合格",
        "答題/答對次數"
      ];

      baseCols.forEach(col => {
        if (!headers.includes(col)) {
          headers.push(col);
          sheet.getRange(1, headers.length).setValue(col);
        }
      });

      // === 組合出處統計（這份考卷內的答題/答對次數） ===
      const stats = {}; // { "出處": { 答題次數, 答對次數 } }
      (record.答題內容 || []).forEach(item => {
        const src = String(item.出處 || "").trim();
        if (!src) return;
        if (!stats[src]) stats[src] = { 答題次數: 0, 答對次數: 0 };
        stats[src].答題次數++;
        if (item.是否答對) stats[src].答對次數++;
      });

      // === 準備要寫入的資料 ===
      const rowValues = new Array(headers.length).fill("");

      rowValues[headers.indexOf("提交時間")] = ts;
      rowValues[headers.indexOf("公務帳號")] = record.公務帳號;
      rowValues[headers.indexOf("單位")] = record.單位;
      rowValues[headers.indexOf("用戶名稱")] = record.用戶名稱;
      rowValues[headers.indexOf("作答題數")] = record.作答題數;
      rowValues[headers.indexOf("答對題數")] = record.答對題數;
      rowValues[headers.indexOf("答題內容")] = JSON.stringify(record.答題內容 || []);
      rowValues[headers.indexOf("是否合格")] = record.是否合格 ? "合格" : "不合格";
      rowValues[headers.indexOf("答題/答對次數")] = JSON.stringify(stats);

      // === 寫入一行完整紀錄（⚠️ appendRow 必須在 lock 裡） ===
      sheet.appendRow(rowValues);

    } finally {
      // ✅ 不論成功或失敗，一定釋放鎖
      lock.releaseLock();
    }
  }

  //寫入審核題目
  function updateReviewInfo(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("題庫");
    const values = sheet.getDataRange().getValues();
    const headers = values[0];

    const idxSerial = headers.indexOf("序號");
    const idxReview = headers.indexOf("審核資訊");
    const idxStatus = headers.indexOf("啟用狀態");

    const rowIdx = values.findIndex((r, i) => i > 0 && String(r[idxSerial]) === String(data.serialNo));
    if (rowIdx === -1) throw new Error("找不到序號 " + data.serialNo);

    let reviewInfo;
    try {
      reviewInfo = JSON.parse(values[rowIdx][idxReview] || '{"reviewers":{}}');
    } catch (e) {
      reviewInfo = { reviewers: {} };
    }
    if (!reviewInfo.reviewers || typeof reviewInfo.reviewers !== "object") {
      reviewInfo.reviewers = {};
    }

    const key = String(data.reviewer).trim();
    if (!reviewInfo.reviewers[key]) reviewInfo.reviewers[key] = {};

    reviewInfo.reviewers[key].decision = data.decision;
    reviewInfo.reviewers[key].reason = data.reason;
    reviewInfo.reviewers[key].time = data.time;

    sheet.getRange(rowIdx + 1, idxReview + 1).setValue(JSON.stringify(reviewInfo));

    const allReviewers = Object.keys(reviewInfo.reviewers || {});
    const decisions = allReviewers.map(name => reviewInfo.reviewers[name]?.decision || "");
    const validDecisions = decisions.filter(d => d);

    let newStatus = "✏️ 待審核";

    if (validDecisions.includes("退件")) {
      newStatus = "❌ 退件";
    } else if (validDecisions.length === allReviewers.length && validDecisions.every(d => d === "通過")) {
      newStatus = "✅ 啟用";
    }

    sheet.getRange(rowIdx + 1, idxStatus + 1).setValue(newStatus);
  }

  /**
   * 命題者重新送審題目
   * @param {Object} updatedQ - 包含序號、出處、題型、題目、A~E、答案、啟用狀態
   */
function resubmitQuestion(updatedQ) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("題庫");
  if (!sheet) throw new Error("❌ 找不到題庫工作表");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxSerial = headers.indexOf("序號");
  const idxReview = headers.indexOf("審核資訊");
  const idxStatus = headers.indexOf("啟用狀態");
  const idxAuthor = headers.indexOf("命題者");

  // === 找到序號對應的列 ===
  const rowIndex = data.findIndex((r, i) => i > 0 && String(r[idxSerial]) === String(updatedQ["序號"]));
  if (rowIndex === -1) throw new Error("找不到序號：" + updatedQ["序號"]);
  const rowNum = rowIndex + 1; // 真實表格列號
  const oldRow = data[rowIndex];
  const author = updatedQ["命題者"] || (idxAuthor >= 0 ? String(oldRow[idxAuthor] || "").trim() : "");

  // === 建立新的審核資訊：清空所有 reviewer 狀態 ===
  const reviewInfo = { reviewers: {} };

  // 找出所有「衛生局」人員作為預設審核者（排除命題者）
  const accSheet = ss.getSheetByName("帳號資訊");
  const accData = accSheet ? accSheet.getDataRange().getValues() : [];
  const accHeaders = accData[0] || [];
  const idxUnit = accHeaders.indexOf("單位名稱");
  const idxName = accHeaders.indexOf("姓名");

  if (idxUnit >= 0 && idxName >= 0) {
    accData.slice(1).forEach(row => {
      const unit = String(row[idxUnit] || "").trim();
      const name = String(row[idxName] || "").trim();
      if (unit === "衛生局" && name && name !== author) {
        reviewInfo.reviewers[name] = { decision: "", reason: "", time: "" };
      }
    });
  }

  // === 處理答案欄（兼容新舊版本）
  let answerValue = updatedQ["答案"];
  try {
    // 若答案是物件就轉成 JSON 字串（例如含 hint、refs）
    if (typeof answerValue === "object") {
      answerValue = JSON.stringify(answerValue);
    } else if (typeof answerValue === "string") {
      // 嘗試解析是否為合法 JSON（若已是 JSON 格式則保留）
      JSON.parse(answerValue);
    }
  } catch (e) {
    // 若不是 JSON 也沒關係，保留原樣
    // console.warn("答案不是 JSON 格式，維持文字格式");
  }

  // === 更新主要欄位 ===
  const fieldsToUpdate = ["出處", "題型", "題目", "A", "B", "C", "D", "E"];
  fieldsToUpdate.forEach(f => {
    const colIdx = headers.indexOf(f);
    if (colIdx >= 0 && updatedQ[f] !== undefined) {
      sheet.getRange(rowNum, colIdx + 1).setValue(updatedQ[f]);
    }
  });

  // ✅ 單獨更新答案（避免 JSON 被破壞）
  const idxAns = headers.indexOf("答案");
  if (idxAns >= 0 && answerValue !== undefined) {
    sheet.getRange(rowNum, idxAns + 1).setValue(answerValue);
  }

  // === 重設啟用狀態為「✏️ 待審核」 ===
  if (idxStatus >= 0) {
    sheet.getRange(rowNum, idxStatus + 1).setValue("✏️ 待審核");
  }

  // === 重置審核資訊 JSON ===
  if (idxReview >= 0) {
    sheet.getRange(rowNum, idxReview + 1).setValue(JSON.stringify(reviewInfo));
  }

  return { success: true, message: "題目已重新送審（命題者未列入審核者）" };
}


