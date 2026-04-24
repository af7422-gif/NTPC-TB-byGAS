function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

  function doGet(e) {
    var tmpl = HtmlService.createTemplateFromFile('api_index');
    tmpl.serviceUrl = ScriptApp.getService().getUrl();

    // 設定您要的圖示 URL，加上 '&format=png' 確保 Apps Script 正常運作
    var faviconUrl = "https://i.meee.com.tw/fBKiHSL.png"; 

    var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    if (ssId === "1_h_z0JljrDQxwg04HrRR6rPNexTk0-VrZkdwwkCACyjcXhY2IcZNoPPG") {
      var output = tmpl.evaluate()
                      .setTitle('銷案審核助手')
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      output.setFaviconUrl(faviconUrl); // 設定圖示
      return output;
    } else {
      var output = tmpl.evaluate()
                      .setTitle('銷案審核助手')
                      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      output.setFaviconUrl(faviconUrl); // 設定圖示
      return output;
    }
  }

  // 驗證登入帳密並更新登入記錄
  function checkLogin(account, password) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("帳號資訊");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idx = col => headers.indexOf(col);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (
        String(row[idx("公務帳號")]).toLowerCase() === String(account).toLowerCase() &&
        row[idx("密碼")] === password
      ) {
        // ===== 新增這段：更新登入欄位 =====
        const now = new Date();
        // 登入次數
        const loginCountIdx = idx("登入次數");
        if (loginCountIdx >= 0) {
          const count = Number(row[loginCountIdx] || 0) + 1;
          sheet.getRange(i + 1, loginCountIdx + 1).setValue(count);
        }
        return "OK";   // <--- 一定要有這行！
        // ===== End =====
      }
    }
    // 登入失敗
    return;
  }


function goToMain() {
  const template = HtmlService.createTemplateFromFile('main');
  return template.evaluate().getContent();
}

function getUrl() {
  var url = ScriptApp.getService().getUrl();
  return url;
}

function getAllAccounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('帳號資訊');
  const fieldMap = getFieldIndexMap(sheet);
  const accountCol = fieldMap['公務帳號'];
  const passwordCol = fieldMap['密碼'];
  const nameCol = fieldMap['姓名'];
  const unitCol = fieldMap['單位名稱']; 

  const numRows = sheet.getLastRow() - 1;
  if (numRows <= 0) return [];
  const data = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();

  return data.map(row => ({
    account: row[accountCol],
    password: row[passwordCol],
    name: row[nameCol],
    unit: row[unitCol]    // unit 欄位就會帶到管理單位
  })).filter(e => e.account && e.password);
}

  // * 取得「銷案清單、帳號資訊」表格中的資料欄位
  function getDisposalListData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    
    // 1. 一次性取得所有工作表的資料，減少與 Sheet 的連線次數
    // 使用 getValues() 抓取整塊資料比多次抓取單一單元格快得多
    const sheetNames = ["帳號資訊", "銷案清單", "題庫", "答題紀錄", "編輯紀錄"];
    const sheetsData = {};
    sheetNames.forEach(name => {
      const sh = ss.getSheetByName(name);
      sheetsData[name] = sh ? sh.getDataRange().getValues() : [];
    });

    // ===== 處理帳號資訊 =====
    const accountData = sheetsData["帳號資訊"];
    let filteredAccounts = [];
    if (accountData.length > 0) {
      const accountHeaders = accountData[0];
      const fieldArr = ["公務帳號", "密碼", "姓名", "單位名稱", "訊息通知", "便利貼", "審核", "合格率"];
      const idxArr = fieldArr.map(f => accountHeaders.indexOf(f));
      const idxExam = accountHeaders.indexOf("考試");
      const idxLastLogin = accountHeaders.indexOf("上次登入時間");

      // 預先過濾有效帳號，減少 map 的次數
      for (let i = 1; i < accountData.length; i++) {
        const row = accountData[i];
        if (!row[idxArr[0]] || !row[idxArr[1]]) continue;

        let obj = {};
        fieldArr.forEach((f, j) => obj[f] = row[idxArr[j]]);
        obj["考試"] = idxExam >= 0 ? row[idxExam] : "TRUE";

        let lastLoginVal = idxLastLogin >= 0 ? row[idxLastLogin] : "";
        obj["上次登入時間"] = (lastLoginVal instanceof Date) 
          ? Utilities.formatDate(lastLoginVal, tz, "yyyy-MM-dd HH:mm:ss") 
          : (lastLoginVal || "");
        
        filteredAccounts.push(obj);
      }
    }

    // ===== 處理銷案清單 =====
    const rawData = sheetsData["銷案清單"];
    let all = [];
    if (rawData.length > 0) {
      const headers = rawData[0].map(h => typeof h === 'string' ? h.trim() : h);
      
      // 預先找出日期與時間欄位的索引，避免在迴圈內反覆判斷
      const dateIdxs = [];
      const timeIdxs = [];
      headers.forEach((h, idx) => {
        if (typeof h === 'string') {
          if (h.includes("日") || h === "銷案期限") dateIdxs.push(idx);
          else if (h.includes("時間")) timeIdxs.push(idx);
        }
      });

      const baseCols = [
        "管理單位","輔導員","個案類型","總編號","個案姓名","申請人","申請時間","結束治療日","銷案原因","銷案期限","審核狀態",
        "不可逆缺失項","不可逆缺失項數","不可逆缺失項(文本)","可逆缺失項","可逆缺失項數","可逆缺失項(文本)","建議事項","建議事項(文本)","備註", "缺失項類別數",
        "退件次數","補件次數","前次退件時間","前次補件時間","局抽回次數","局前次抽回時間", "第1次退件時間", "第1次退件人",
        "最近審核者","最近審核時間","審核通過時間","審核通過者","審核花費天數","案件處理可用天數"
      ];
      const allCols = baseCols; 
      const allIdx = allCols.map(col => headers.indexOf(col));
      const deadlineIdx = headers.indexOf("銷案期限");

      // 排序與過濾空白行
      const sorted = rawData.slice(1)
        .filter(row => row && row.some(cell => cell !== ""))
        .sort((a, b) => {
          const d1 = (deadlineIdx >= 0 && a[deadlineIdx] instanceof Date) ? a[deadlineIdx] : new Date("9999-12-31");
          const d2 = (deadlineIdx >= 0 && b[deadlineIdx] instanceof Date) ? b[deadlineIdx] : new Date("9999-12-31");
          return d1 - d2;
        });

      // 轉換資料格式
      const formattedBody = sorted.map(row => allIdx.map(idx => {
        let val = idx >= 0 ? row[idx] : "";
        if (val instanceof Date) {
          if (dateIdxs.includes(idx)) return Utilities.formatDate(val, tz, 'yyyy/MM/dd');
          if (timeIdxs.includes(idx)) return Utilities.formatDate(val, tz, 'yyyy/MM/dd HH:mm:ss');
        }
        return val ?? "";
      }));
      all = [allCols].concat(formattedBody);
    }

    // ===== 處理編輯紀錄 =====
    const editRaw = sheetsData["編輯紀錄"];
    let editRecords = [];

    if (editRaw.length > 0) {
      const headers = editRaw[0];
      const timeIdx = headers.indexOf("編輯時間");

      const body = editRaw.slice(1).map(row => {
        if (timeIdx >= 0 && row[timeIdx] instanceof Date) {
          let newRow = [...row];
          newRow[timeIdx] = Utilities.formatDate(row[timeIdx], tz, "yyyy/MM/dd HH:mm:ss");
          return newRow;
        }
        return row;
      });

      editRecords = [headers].concat(body);
    }

    // ===== 處理答題紀錄 =====
    const examRaw = sheetsData["答題紀錄"];
    let examRecords = [];
    if (examRaw.length > 0) {
      const headers = examRaw[0];
      const tsIdx = headers.indexOf("提交時間");
      const body = examRaw.slice(1).map(row => {
        if (tsIdx >= 0 && row[tsIdx] instanceof Date) {
          let newRow = [...row];
          newRow[tsIdx] = Utilities.formatDate(row[tsIdx], tz, "yyyy/MM/dd HH:mm:ss");
          return newRow;
        }
        return row;
      });
      examRecords = [headers].concat(body);
    }

    return {
      all: all,
      accounts: filteredAccounts,
      questions: sheetsData["題庫"], // 題庫不需處理直接回傳
      examRecords: examRecords,
      editRecords: editRecords
    };
  }

// 儲存便利貼
function saveUserNotepad(name, unit, content) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("帳號資訊");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = headers.indexOf("姓名");
  const unitIdx = headers.indexOf("單位名稱");
  const noteIdx = headers.indexOf("便利貼");

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameIdx] == name && data[i][unitIdx] == unit) {
      sheet.getRange(i + 1, noteIdx + 1).setValue(content);
      return { success: true, message: "便利貼已更新" };
    }
  }
  return { success: false, message: "找不到使用者" };
}

//寫入新的銷案申請資料
function submitDisposalRequest(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("銷案清單");
    const headers = sheet.getDataRange().getValues()[0];
    const newRow = [];

    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const formattedNow = Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm:ss");
    const formattedEndDate = Utilities.formatDate(new Date(data.endDate), tz, "yyyy/MM/dd");

    // 預先算出銷案期限
    const deadline = calcDisposalDeadline(data.caseType, data.endDate, data.reason);

    // 計算案件處理可用天數
    const deadlineDate = toDate(deadline);
    const applyDate = toDate(formattedNow);
    let daysAvailable = "";
    if (!isNaN(deadlineDate) && !isNaN(applyDate)) {
      daysAvailable = Math.floor((deadlineDate - applyDate) / (1000 * 60 * 60 * 24));
      if (daysAvailable < 0) daysAvailable = 0;
    }

    // 1. 各區域對應輔導員設定
    const advisorMap = {
      "板橋區": "林君諭",
      "三重區": "江適揚",
      "中和區": "劉宇倫",
      "永和區": "李權展",
      "新莊區": "黃依婷",
      "新店區": "楊政憲",
      "土城區": "施冠毅",
      "蘆洲區": "陳詠梅",
      "汐止區": "周文懋",
      "樹林區": "陳詠梅",
      "淡水區": "李權展",
      "三峽區": "施冠毅",
      "鶯歌區": "楊政憲",
      "瑞芳區": "楊政憲",
      "五股區": "周文懋",
      "泰山區": "李權展",
      "林口區": "施冠毅",
      "深坑區": "楊政憲",
      "石碇區": "江適揚",
      "坪林區": "江適揚",
      "三芝區": "劉宇倫",
      "石門區": "劉宇倫",
      "八里區": "施冠毅",
      "平溪區": "江適揚",
      "萬里區": "陳詠梅",
      "烏來區": "林君諭",
      "雙溪區": "江適揚",
      "貢寮區": "李權展",
      "金山區": "陳詠梅"
    };
    const advisorName = advisorMap[data.unit] || "";

    headers.forEach(header => {
      switch (header) {
        case "個案類型": newRow.push(data.caseType); break;
        case "總編號": newRow.push(`'${data.tb}`); break;
        case "個案姓名": newRow.push(data.caseName); break;
        case "結束治療日": newRow.push(formattedEndDate); break;
        case "銷案原因": newRow.push(data.reason); break;
        case "銷案期限": newRow.push(deadline); break;
        case "案件處理可用天數": newRow.push(daysAvailable); break;
        case "管理單位": newRow.push(data.unit); break;
        case "申請人": newRow.push(data.userName); break;
        case "申請時間": newRow.push(formattedNow); break;
        case "審核狀態": newRow.push("❓ 尚未處理"); break;
        case "輔導員": newRow.push(advisorName); break;
        case "申請附件(圖片JSON)": newRow.push(data.attachmentsJson || "[]"); break;
        default: newRow.push("");
      }
    });

    sheet.appendRow(newRow);
    const tbIdx = headers.indexOf("總編號");
    const lastRow = sheet.getLastRow();
    if (tbIdx !== -1) {
      sheet.getRange(lastRow, tbIdx + 1).setNumberFormat("@");
    }

    const daysIdx = headers.indexOf("案件處理可用天數");
    if (daysIdx !== -1) {
      sheet.getRange(lastRow, daysIdx + 1).setNumberFormat("0");
    }

    // 訊息通知輔導員
    /*
    if (advisorName) {
      const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("帳號資訊");
      const userData = userSheet.getDataRange().getValues();
      const userHeaders = userData[0].map(h => typeof h === "string" ? h.trim() : "");
      const nameIdx = userHeaders.indexOf("姓名");
      const msgIdx = userHeaders.indexOf("訊息通知");

      if (nameIdx !== -1 && msgIdx !== -1) {
        const userMap = new Map();
        for (let i = 1; i < userData.length; i++) {
          const uname = String(userData[i][nameIdx]).trim();
          userMap.set(uname, i);
        }

        const targetIdx = userMap.get(advisorName);
        if (targetIdx !== undefined) {
          const msgRow = userData[targetIdx];
          const newMsg = `${data.caseType}個案 ${data.tb} 已於 ${formattedNow} 📨 申請`;
          const oldMsg = msgRow[msgIdx] ? String(msgRow[msgIdx]) : "";
          const totalMsg = oldMsg ? newMsg + "\n" + oldMsg : newMsg;
          userSheet.getRange(targetIdx + 1, msgIdx + 1).setValue(totalMsg);
        }
      }
    }
    */

    // 新增編輯紀錄
    writeDisposalLog({
      unit: data.unit,
      userName: data.userName,
      caseType: data.caseType,
      tb: data.tb,
      caseName: data.caseName,
      unreversible: "[]",
      reversible: "[]",
      suggestion: "[]",
      status: "📨 申請銷案",
      note: data.備註
    });

  } catch (err) {
    // 🚨 發生錯誤或逾時 → 寄信通知
    const subject = "【銷案系統異常通知】submitDisposalRequest 發生錯誤";
    const body =
      "時間：" + new Date() +
      "\n錯誤訊息：" + err.message +
      "\n\n堆疊：" + (err.stack || "無") +
      "\n\n可能失敗的申請資料如下：\n" +
      JSON.stringify(data, null, 2);

    GmailApp.sendEmail({
      to: "AF7422@ntpc.gov.tw",
      subject,
      body
    });

    throw err; // 保留原本錯誤供前端捕捉
  }
}


  //提交審核者表單
  function submitReviewDecision(data) {
    const adminEmail = "AF7422@ntpc.gov.tw";  // 📬 通知收件人
    const reviewerName = data.reviewerName || "(未填)";
    const startTime = new Date();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("銷案清單");
    const rawData = sheet.getDataRange().getValues();
    const headers = rawData[0].map(h => typeof h === 'string' ? h.trim() : h);
    const tbNumber = String(data.tb).trim();


    try {
      // 🕐 格式統一化
      function normalizeDateTime(val) {
        if (val instanceof Date) {
          return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
        }
        return String(val).replace(/-/g, "/").trim();
      }

      const tbIdx = headers.indexOf("總編號");
      const applyTimeIdx = headers.indexOf("申請時間");
      
      // === 主要比對：總編號 + 申請時間 ===
      let rowIdx = rawData.findIndex((row, i) =>
        i > 0 &&
        String(row[tbIdx]).trim() === tbNumber &&
        normalizeDateTime(row[applyTimeIdx]) === normalizeDateTime(data.applyTime)
      );

      Logger.log("申請時間：" + applyTimeIdx);

      // === 若找不到，退回用 TB 搜尋 ===
      let matchStatus = "FULL_MATCH";
      if (rowIdx === -1) {
        const fallbackIdx = rawData.findIndex((row, i) =>
          i > 0 && String(row[tbIdx]).trim() === tbNumber
        );
        if (fallbackIdx !== -1) {
          rowIdx = fallbackIdx;
          matchStatus = "TB_ONLY_MATCH";  // ⚠️ 時間不符但找到 TB
        } else {
          matchStatus = "NO_MATCH";       // ❌ 完全找不到
        }
      }

      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

      // === 寄信邏輯 ===
      if (matchStatus === "NO_MATCH") {
        const msg = [
          "🚨【銷案比對失敗】",
          `審核者：${reviewerName}`,
          `TB 總編號：${data.tb}`,
          `申請時間：${data.applyTime || "(無)"}`,
          `發生時間：${nowStr}`,
          "",
          "比對條件：總編號 + 申請時間",
          "比對結果：完全找不到對應資料列。"
        ].join("\n");

        GmailApp.sendEmail({
          to: adminEmail,
          subject: `[銷案審核錯誤] ${reviewerName} 找不到 ${data.tb}`,
          body: msg
        });

        throw new Error(`找不到個案：${data.tb} / ${data.applyTime}`);
      }

      if (matchStatus === "TB_ONLY_MATCH") {
        const msg = [
          "⚠️【銷案時間不符警告】",
          `審核者：${reviewerName}`,
          `TB 總編號：${data.tb}`,
          `申請時間(傳入)：${data.applyTime || "(無)"}`,
          `資料列中的申請時間：${rawData[rowIdx][applyTimeIdx]}`,
          `發生時間：${nowStr}`,
          "",
          "比對條件：總編號 + 申請時間",
          "比對結果：時間不符，但找到相同 TB 資料列。",
          "",
          "（系統已自動繼續以 TB 比對進行寫入）"
        ].join("\n");

        try {
          GmailApp.sendEmail({
            to: adminEmail,
            subject: `[銷案警告] ${reviewerName} 審核 ${data.tb} 時申請時間不符`,
            body: msg
          });
        } catch (mailErr) {
          Logger.log("⚠️ 郵件寄送失敗：" + mailErr.message);
        }
      }

      // === 找到資料列（正常或 fallback）
      let row = rawData[rowIdx];
        const noteIdx = headers.indexOf("備註");
        const spendDaysIdx = headers.indexOf("審核花費天數");
        const reasonIdx = headers.indexOf("銷案原因");
        const endDateIdx = headers.indexOf("結束治療日");
        const statusIdx = headers.indexOf("審核狀態");
        const caseTypeIdx = headers.indexOf("個案類型");
        const hasunreversibleIdx = headers.indexOf("是否有不可逆缺失項");
        const hasReversibleIdx = headers.indexOf("是否有可逆缺失項");
        const hasSuggestionIdx = headers.indexOf("是否有建議事項");
        const unreversibleIdx = headers.indexOf("不可逆缺失項");
        const unreversibleCountIdx = headers.indexOf("不可逆缺失項數");
        const reversibleIdx = headers.indexOf("可逆缺失項");
        const reversibleCountIdx = headers.indexOf("可逆缺失項數");
        const suggestionCountIdx = headers.indexOf("建議事項數");
        const suggestionIdx = headers.indexOf("建議事項");
        const approveTimeIdx = headers.indexOf("審核通過時間");
        const approveUserIdx = headers.indexOf("審核通過者");
        const rejectCountIdx = headers.indexOf("退件次數");
        const lastRejectTimeIdx = headers.indexOf("前次退件時間");
        const cancelTimeIdx = headers.indexOf("取消銷案時間");
        const cancelUserIdx = headers.indexOf("取消銷案者");
        const irrTextIdx = headers.indexOf("不可逆缺失項(文本)");
        const revTextIdx = headers.indexOf("可逆缺失項(文本)");
        const sugTextIdx = headers.indexOf("建議事項(文本)");

          // 銷案期限
          const deadlineIdx = headers.findIndex(h => String(h).trim() === "銷案期限");
          const daysIdx = headers.findIndex(h => String(h).trim() === "案件處理可用天數");

          if (deadlineIdx !== -1) {
            const newDeadline = calcDisposalDeadline(data.caseType, data.endDate, data.reason);
            row[deadlineIdx] = newDeadline;

            // ✅ 計算「案件處理可用天數」
            if (daysIdx !== -1 && applyTimeIdx !== -1) {
              const applyDate = toDate(row[applyTimeIdx]);
              const deadlineDate = toDate(newDeadline);
              if (applyDate && deadlineDate) {
                let daysAvailable = Math.floor((deadlineDate - applyDate) / (1000 * 60 * 60 * 24));
                if (daysAvailable < 0) daysAvailable = 0;
                row[daysIdx] = daysAvailable;
              }
            }
          }


          row[reasonIdx] = data.reason;
          row[unreversibleCountIdx] = data.unreversibleCount;
          row[caseTypeIdx] = data.caseType;
          row[reversibleCountIdx] = data.reversibleCount;
          row[suggestionCountIdx] = data.suggestionCount;
          row[unreversibleIdx] = data.unreversible;
          row[reversibleIdx] = data.reversible;
          row[suggestionIdx] = data.suggestion;
          row[endDateIdx] = data.endDate;

          // 三大區塊JSON → 文本
          function lossItemJsonToText(jsonStr) {
            let arr = [];
            try {
              arr = JSON.parse(jsonStr || '[]');
              if (!Array.isArray(arr)) arr = [];
            } catch(e) {
              arr = [];
            }
            return arr.map(item => {
              let checkMark = item.checked ? "🗹" : "☐";
              let selectText = item.select ? `【${item.select}】` : "";
              let remark = item.remark ? item.remark : "";
              return `${checkMark} ${selectText}${remark}`;
            }).join('\n');
          }
          if (irrTextIdx !== -1) row[irrTextIdx] = lossItemJsonToText(data.unreversible);
          if (revTextIdx !== -1) row[revTextIdx] = lossItemJsonToText(data.reversible);
          if (sugTextIdx !== -1) row[sugTextIdx] = lossItemJsonToText(data.suggestion);
 
          // ===== 計算「缺失項類別數」=====
          let defectCategoryCount = 0;

          try {
            const irr = JSON.parse(data.unreversible || '[]');
            const rev = JSON.parse(data.reversible || '[]');

            const set = new Set();

            [...irr, ...rev].forEach(item => {
              if (item.select && item.select.trim()) {
                set.add(item.select.trim());
              }
            });

            defectCategoryCount = set.size;
          } catch (e) {
            defectCategoryCount = 0;
          }

          //缺失項類別數
          function ensureCol(headerArr, sheet, colName) {
            let idx = headerArr.findIndex(h => String(h).trim() === colName.trim());
            if (idx === -1) {
              sheet.getRange(1, headerArr.length + 1).setValue(colName.trim());
              headerArr.push(colName.trim());
              row.push("");
              idx = headerArr.length - 1;
            }
            return idx;
          }

          const defectCategoryIdx = ensureCol(headers, sheet, "缺失項類別數");
          row[defectCategoryIdx] = defectCategoryCount;

          // ===== 新增／更新「最近審核者」「最近審核時間」=====
          const lastReviewerIdx = ensureCol(headers, sheet, "最近審核者");
          const lastReviewTimeIdx = ensureCol(headers, sheet, "最近審核時間");

          row[lastReviewerIdx] = reviewerName;
          row[lastReviewTimeIdx] = nowStr;

          if (hasunreversibleIdx !== -1) {
            row[hasunreversibleIdx] = Number(row[unreversibleCountIdx]) > 0 ? "Y" : "N";
          }
          if (hasReversibleIdx !== -1) {
            row[hasReversibleIdx] = Number(row[reversibleCountIdx]) > 0 ? "Y" : "N";
          }
          if (hasSuggestionIdx !== -1) {
            row[hasSuggestionIdx] = Number(row[suggestionCountIdx]) > 0 ? "Y" : "N";
          }

          if (noteIdx !== -1) row[noteIdx] = data.備註 || "";

          if (data.status.indexOf("approve") >= 0) {
            row[statusIdx] = "✅ 通過";
            row[approveTimeIdx] = data.time;
            row[approveUserIdx] = data.reviewerName;
            // === 計算「審核花費天數」：無條件捨去整數天 ===
            if (applyTimeIdx !== -1 && spendDaysIdx !== -1 && approveTimeIdx !== -1) {
              const applyAt = toDate(row[applyTimeIdx]);
              // 優先用剛寫入列的值（可能是 Date 或字串）；退而求其次用 data.time
              const approveAt = toDate(row[approveTimeIdx]) || toDate(data.time);

              if (applyAt && approveAt) {
                // 以毫秒差換算天數，無條件捨去
                let days = Math.floor((approveAt.getTime() - applyAt.getTime()) / (1000 * 60 * 60 * 24));
                if (days < 0) days = 0; // 防呆：不讓負數進表
                row[spendDaysIdx] = days;

                // 也可以順便把欄位格式設為整數
                sheet.getRange(rowIdx + 1, spendDaysIdx + 1).setNumberFormat("0");
              }
            }
          } else if (data.status === "⏪ 取消銷案") {
            row[statusIdx] = "⏪ 取消銷案";
            if (cancelTimeIdx !== -1) row[cancelTimeIdx] = data.time || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
            if (cancelUserIdx !== -1) row[cancelUserIdx] = data.reviewerName || "";
          }
          else if (data.status.indexOf("暫存") >= 0 || data.status.indexOf("pending") >= 0 || data.status.indexOf("⌛") >= 0) {
            row[statusIdx] = "⌛ 處理中";
          } else {
            row[statusIdx] = "❌ 退件";
            row[rejectCountIdx] = data.rejectCount;
            row[lastRejectTimeIdx] = data.time;
            const thisReject = data.rejectCount;
            const rejectTimeIdx = headers.indexOf(`第${thisReject}次退件時間`);
            const rejectUserIdx = headers.indexOf(`第${thisReject}次退件人`);
            if (rejectTimeIdx !== -1) row[rejectTimeIdx] = data.time;
            if (rejectUserIdx !== -1) row[rejectUserIdx] = data.reviewerName;
          }

          sheet.getRange(rowIdx+1, 1, 1, headers.length).setValues([row]);
          if (tbIdx !== -1) {
            sheet.getRange(rowIdx+1, tbIdx + 1).setNumberFormat("@");
          }
          if (daysIdx !== -1) {
            sheet.getRange(rowIdx + 1, daysIdx + 1).setNumberFormat("0");
          }

          /*
          // === 通知申請人 ===
          if (["✅ 通過", "❌ 退件", "⏪ 取消銷案"].includes(row[statusIdx])) {
            const applicantIdx = headers.indexOf("申請人");
            const tbNumber = row[tbIdx];
            const applicantName = applicantIdx !== -1 ? String(row[applicantIdx]).trim() : "";
            const caseType = row[headers.indexOf("個案類型")] || "";

            const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("帳號資訊");
            const userData = userSheet.getDataRange().getValues();
            const userHeaders = userData[0].map(h => typeof h === "string" ? h.trim() : h);
            const nameIdx = userHeaders.indexOf("姓名");
            const msgIdx = userHeaders.indexOf("訊息通知");

            if (nameIdx !== -1 && msgIdx !== -1) {
              const userMap = new Map();
              for (let i = 1; i < userData.length; i++) {
                const name = String(userData[i][nameIdx]).trim();
                userMap.set(name, i);
              }

              const userRowIdx = userMap.get(applicantName);
              if (userRowIdx !== undefined) {
                const msgRow = userData[userRowIdx];
                const now = new Date();
                const pad = n => String(n).padStart(2, "0");
                const nowStr = `${now.getFullYear()}/${pad(now.getMonth() + 1)}/${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
                const newMsg = `${caseType}個案 ${tbNumber} 已於 ${nowStr} ${row[statusIdx]}`;
                const oldMsg = msgRow[msgIdx] ? String(msgRow[msgIdx]) : "";
                const totalMsg = oldMsg ? newMsg + "\n" + oldMsg : newMsg;
                userSheet.getRange(userRowIdx + 1, msgIdx + 1).setValue(totalMsg);
              }
            }
          }
          */
          
          // === 新增審核編輯紀錄 ===
          writeDisposalLog({
            unit: data.unit,
            reviewerName: data.reviewerName,
            caseType: row[headers.indexOf("個案類型")],
            tb: row[headers.indexOf("總編號")],
            caseName: row[headers.indexOf("個案姓名")],
            unreversible: data.unreversible,
            reversible: data.reversible,
            suggestion: data.suggestion,
            status: row[statusIdx],
            note: data.備註
          });

      } catch (err) {
        // === 精簡後錯誤通知區塊（只寄嚴重錯誤）===
        const errMsg = [
          "❌ submitReviewDecision() 執行發生未預期錯誤",
          `審核者：${reviewerName}`,
          `TB 總編號：${data.tb}`,
          `申請時間：${data.applyTime || "(無)"}`,
          `錯誤時間：${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss")}`,
          "",
          "錯誤訊息：",
          err.message,
          "",
          "堆疊追蹤：",
          err.stack
        ].join("\n");

        Logger.log(errMsg);
        GmailApp.sendEmail({
          to: adminEmail,
          subject: `[銷案執行錯誤] ${reviewerName} 執行 submitReviewDecision 失敗`,
          body: errMsg
        });

        throw err;

      } finally {
        const endTime = new Date();
        const duration = ((endTime - startTime) / 1000).toFixed(2);
        Logger.log(`🕒 submitReviewDecision 執行完畢（${duration} 秒）`);
      }
  }
  
  //審核者抽回
  function recallReview(data) { 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("銷案清單");
    const rawData = sheet.getDataRange().getValues();
    const headers = rawData[0].map(h => String(h).trim());

    const tbIdx = headers.indexOf("總編號");
    const applyTimeIdx = headers.indexOf("申請時間");
    const statusIdx = headers.indexOf("審核狀態");
    const recallCntIdx = headers.indexOf("局抽回次數");
    const recallTimeIdx = headers.indexOf("局前次抽回時間");

    function normalizeDateTime(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
      }
      return String(val).replace(/-/g, "/").trim();
    }

    // ✅ 改成用「總編號 + 申請時間」比對唯一列
    const rowIdx = rawData.findIndex((row, i) =>
      i > 0 &&
      String(row[tbIdx]).trim() === String(data.tb).trim() &&
      normalizeDateTime(row[applyTimeIdx]) === normalizeDateTime(data.applyTime)
    );

    if (rowIdx === -1) {
      throw new Error(`找不到個案：${data.tb} / ${data.applyTime}`);
    }

    const row = rawData[rowIdx];
    const now = new Date();

    // 更新銷案清單
    row[statusIdx] = "⌛ 處理中";
    row[recallCntIdx] = (parseInt(row[recallCntIdx] || "0", 10) + 1);
    row[recallTimeIdx] = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    sheet.getRange(rowIdx + 1, 1, 1, headers.length).setValues([row]);

    /*
    // ===== 通知承辦人 =====
    const applicantIdx = headers.indexOf("申請人");
    const caseTypeIdx = headers.indexOf("個案類型");
    const caseNameIdx = headers.indexOf("個案姓名");

    const applicantName = applicantIdx !== -1 ? String(row[applicantIdx]).trim() : "";
    const caseType = caseTypeIdx !== -1 ? row[caseTypeIdx] : "";
    const caseName = caseNameIdx !== -1 ? row[caseNameIdx] : "";

    const userSheet = ss.getSheetByName("帳號資訊");
    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData[0].map(h => String(h).trim());
    const nameIdx = userHeaders.indexOf("姓名");
    const msgIdx = userHeaders.indexOf("訊息通知");

    if (nameIdx !== -1 && msgIdx !== -1) {
      const userMap = new Map();
      for (let i = 1; i < userData.length; i++) {
        const name = String(userData[i][nameIdx]).trim();
        userMap.set(name, i);
      }

      const userRowIdx = userMap.get(applicantName);
      if (userRowIdx !== undefined) {
        const msgRow = userData[userRowIdx];
        const pad = n => String(n).padStart(2, "0");
        const nowStr = `${now.getFullYear()}/${pad(now.getMonth() + 1)}/${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

        // ✅ 改這行：使用 data.tb 而非 tbNumber
        const newMsg = `${caseType}個案 ${data.tb} 已於 ${nowStr} ↩️ 抽回`;

        const oldMsg = msgRow[msgIdx] ? String(msgRow[msgIdx]) : "";
        const totalMsg = oldMsg ? newMsg + "\n" + oldMsg : newMsg;
        userSheet.getRange(userRowIdx + 1, msgIdx + 1).setValue(totalMsg);
      }
    }
    */

    // ===== 編輯紀錄 =====
    writeDisposalLog({
      unit: data.unit,
      reviewerName: data.reviewerName,
      caseType: caseType,
      tb: data.tb,   // ✅ 同樣使用 data.tb
      caseName: caseName,
      unreversible: "",
      reversible: "",
      suggestion: "",
      status: "↩️ 抽回",
      note: ""
    });

    return "OK";
  }


  // 小工具：把各種可能的值轉成 Date，沒法轉就回傳 null
  function toDate(val) {
    if (val instanceof Date) return val;
    if (typeof val === 'number') return new Date(val); // Google Sheet 內部日期可能是 serial number
    if (typeof val === 'string') {
      // 嘗試解析 "yyyy/MM/dd HH:mm:ss" 或 "yyyy-MM-dd HH:mm:ss"
      const m = val.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
      if (m) {
        return new Date(
          Number(m[1]), Number(m[2]) - 1, Number(m[3]),
          Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)
        );
      }
      const d = new Date(val);
      if (!isNaN(d)) return d;
    }
    return null;
  }

  /**
   * 補件者送出表單，寫入勾選內容、更新「第N次補件時間」與「第N次補件人」
   */

  //確認審核狀態
  function checkCaseStatus(tb, applyTime) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("銷案清單");
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const tbIdx = headers.indexOf("總編號");
    const applyIdx = headers.indexOf("申請時間");
    const statusIdx = headers.indexOf("審核狀態");

    function normalizeDateTime(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
      }
      return String(val).replace(/-/g, "/").trim();
    }

    const target = data.find((row, i) =>
      i > 0 &&
      String(row[tbIdx]).trim() === tb &&
      normalizeDateTime(row[applyIdx]) === normalizeDateTime(applyTime)
    );

    if (!target) return { found: false, status: null };
    return { found: true, status: target[statusIdx] };
  }


  function submitEditorPatch(data) {
    const adminEmail = "AF7422@ntpc.gov.tw";  // 📬 錯誤/警告通知收件人
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("銷案清單");
    const rawData = sheet.getDataRange().getValues();
    const headers = rawData[0].map(h => typeof h === 'string' ? h.trim() : h);
    const tbIdx = headers.indexOf("總編號");
    const applyTimeIdx = headers.indexOf("申請時間");
    const reviewerName = data.patchName || "(未填)";
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    function normalizeDateTime(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
      }
      return String(val).replace(/-/g, "/").trim();
    }

    // === 主要比對：總編號 + 申請時間 ===
    let rowIdx = rawData.findIndex((row, i) =>
      i > 0 &&
      String(row[tbIdx]).trim() === String(data.tb).trim() &&
      normalizeDateTime(row[applyTimeIdx]) === normalizeDateTime(data.applyTime)
    );

    // === 若找不到，用 TB 比對 fallback ===
    let matchStatus = "FULL_MATCH";
    if (rowIdx === -1) {
      const fallbackIdx = rawData.findIndex((row, i) =>
        i > 0 && String(row[tbIdx]).trim() === String(data.tb).trim()
      );
      if (fallbackIdx !== -1) {
        rowIdx = fallbackIdx;
        matchStatus = "TB_ONLY_MATCH";
      } else {
        matchStatus = "NO_MATCH";
      }
    }

    // === 根據結果寄信通知 ===
    if (matchStatus === "NO_MATCH") {
      const msg = [
        "🚨【補件比對失敗】",
        `補件者：${reviewerName}`,
        `TB 總編號：${data.tb}`,
        `申請時間：${data.applyTime || "(無)"}`,
        `發生時間：${nowStr}`,
        "",
        "比對條件：總編號 + 申請時間",
        "比對結果：完全找不到對應資料列。"
      ].join("\n");

      GmailApp.sendEmail({
        to: adminEmail,
        subject: `[銷案補件錯誤] ${reviewerName} 找不到 ${data.tb}`,
        body: msg
      });
      throw new Error(`找不到個案：${data.tb} / ${data.applyTime}`);
    }

    if (matchStatus === "TB_ONLY_MATCH") {
      const msg = [
        "⚠️【補件時間不符警告】",
        `補件者：${reviewerName}`,
        `TB 總編號：${data.tb}`,
        `申請時間(傳入)：${data.applyTime || "(無)"}`,
        `資料列中的申請時間：${rawData[rowIdx][applyTimeIdx]}`,
        `發生時間：${nowStr}`,
        "",
        "比對條件：總編號 + 申請時間",
        "比對結果：時間不符，但找到相同 TB 資料列。",
        "",
        "（系統已自動以 TB 比對繼續補件寫入）"
      ].join("\n");

      try {
        GmailApp.sendEmail({
          to: adminEmail,
          subject: `[銷案補件警告] ${reviewerName} 補件 ${data.tb} 時申請時間不符`,
          body: msg
        });
      } catch (mailErr) {
        Logger.log("⚠️ 郵件寄送失敗：" + mailErr.message);
      }
    }

    // === 實際更新資料列 ===
    let row = rawData[rowIdx];

    const noteIdx = headers.indexOf("備註");
    const caseTypeIdx = headers.indexOf("個案類型");
    const reasonIdx = headers.indexOf("銷案原因");

    const patchCountIdx = headers.indexOf("補件次數");
    let patchNum = 1;
    if (patchCountIdx !== -1) {
      const prev = row[patchCountIdx];
      patchNum = (typeof prev === "number" ? prev : parseInt(prev, 10) || 0) + 1;
      row[patchCountIdx] = patchNum;
    }

    const deadlineIdx = headers.indexOf("銷案期限");
    const statusIdx = headers.indexOf("審核狀態");
    const lastPatchTimeIdx = headers.indexOf("前次補件時間");
    const unreversibleIdx = headers.indexOf("不可逆缺失項");
    const unreversibleCountIdx = headers.indexOf("不可逆缺失項數");
    const reversibleIdx = headers.indexOf("可逆缺失項");
    const reversibleCountIdx = headers.indexOf("可逆缺失項數");
    const suggestionIdx = headers.indexOf("建議事項");
    const suggestionCountIdx = headers.indexOf("建議事項數");
    const irrTextIdx = headers.indexOf("不可逆缺失項(文本)");
    const revTextIdx = headers.indexOf("可逆缺失項(文本)");
    const sugTextIdx = headers.indexOf("建議事項(文本)");

    const patchTimeIdx = headers.indexOf(`第${patchNum}次補件時間`);
    const patchUserIdx = headers.indexOf(`第${patchNum}次補件人`);

    row[caseTypeIdx]          = data.caseType;
    row[reasonIdx]            = data.reason;
    row[noteIdx]              = data.備註 || "";
    row[statusIdx]            = "🔄 補件";
    row[unreversibleIdx]      = data.unreversible;
    row[unreversibleCountIdx] = data.unreversibleCount;
    row[reversibleIdx]        = data.reversible;
    row[reversibleCountIdx]   = data.reversibleCount;
    row[suggestionIdx]        = data.suggestion;
    row[suggestionCountIdx]   = data.suggestionCount;
    row[patchTimeIdx]         = data.patchTime;
    row[patchUserIdx]         = data.patchName;
    row[lastPatchTimeIdx]     = data.patchTime;
    row[deadlineIdx]          = calcDisposalDeadline(data.caseType, data.endDate, data.reason);

    function lossItemJsonToText(jsonStr) {
      let arr = [];
      try {
        arr = JSON.parse(jsonStr || '[]');
        if (!Array.isArray(arr)) arr = [];
      } catch (e) {
        arr = [];
      }
      return arr.map(item => {
        let checkMark = item.checked ? "🗹" : "☐";
        let selectText = item.select ? `【${item.select}】` : "";
        let remark = item.remark ? item.remark : "";
        return `${checkMark} ${selectText}${remark}`;
      }).join('\n');
    }

    if (irrTextIdx !== -1) row[irrTextIdx] = lossItemJsonToText(data.unreversible);
    if (revTextIdx !== -1) row[revTextIdx] = lossItemJsonToText(data.reversible);
    if (sugTextIdx !== -1) row[sugTextIdx] = lossItemJsonToText(data.suggestion);

    sheet.getRange(rowIdx + 1, 1, 1, headers.length).setValues([row]);
    if (tbIdx !== -1) {
      sheet.getRange(rowIdx + 1, tbIdx + 1).setNumberFormat("@");
    }

    /*
    // === 補件通知輔導員 ===
    if (row[statusIdx] === "🔄 補件") {
      const helperIdx = headers.indexOf("輔導員");
      const tbNumber = row[tbIdx];
      const helperName = helperIdx !== -1 ? String(row[helperIdx]).trim() : "";

      const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("帳號資訊");
      const userData = userSheet.getDataRange().getValues();
      const userHeaders = userData[0].map(h => typeof h === "string" ? h.trim() : h);
      const nameIdx = userHeaders.indexOf("姓名");
      const msgIdx = userHeaders.indexOf("訊息通知");

      if (nameIdx !== -1 && msgIdx !== -1 && helperName) {
        const userMap = new Map();
        for (let i = 1; i < userData.length; i++) {
          const uname = String(userData[i][nameIdx]).trim();
          userMap.set(uname, i);
        }

        const userRowIdx = userMap.get(helperName);
        if (userRowIdx !== undefined) {
          const msgRow = userData[userRowIdx];
          const now = new Date();
          const pad = n => String(n).padStart(2, "0");
          const nowStr = `${now.getFullYear()}/${pad(now.getMonth() + 1)}/${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
          const caseTypeColIdx = headers.indexOf("個案類型");
          const caseType = row[caseTypeColIdx];
          const newMsg = `${caseType}個案 ${tbNumber} 已於 ${nowStr} ${row[statusIdx]}`;
          const oldMsg = msgRow[msgIdx] ? String(msgRow[msgIdx]) : "";
          const totalMsg = oldMsg ? newMsg + "\n" + oldMsg : newMsg;
          userSheet.getRange(userRowIdx + 1, msgIdx + 1).setValue(totalMsg);
        }
      }
    }
    */

    // === 新增補件編輯紀錄 ===
    writeDisposalLog({
      unit: row[headers.indexOf("管理單位")],
      patchName: data.patchName,
      caseType: row[headers.indexOf("個案類型")],
      tb: row[headers.indexOf("總編號")],
      caseName: row[headers.indexOf("個案姓名")],
      unreversible: data.unreversible,
      reversible: data.reversible,
      suggestion: data.suggestion,
      status: "🔄 補件",
      note: data.備註
    });
  }



  // 取得帳號的訊息內容
  function getUserMsgContent(name) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("帳號資訊");
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(String);
    var idxName = headers.indexOf("姓名");
    var idxMsg = headers.indexOf("訊息通知");
    for (var i=1; i<data.length; i++) {
      if (String(data[i][idxName]).trim() === name) {
        return data[i][idxMsg] || "";
      }
    }
    return "";
  }

  // 清除帳號的訊息內容
  function clearUserMsgContent(name) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("帳號資訊");
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(String);
    var idxName = headers.indexOf("姓名");
    var idxMsg = headers.indexOf("訊息通知");
    for (var i=1; i<data.length; i++) {
      if (String(data[i][idxName]).trim() === name) {
        sheet.getRange(i+1, idxMsg+1).setValue("");
        break;
      }
    }
  }

//每日批次提醒銷案期限
function updateDeadlineNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const caseSheet = ss.getSheetByName("銷案清單");
  const accountSheet = ss.getSheetByName("帳號資訊");
  const caseData = caseSheet.getDataRange().getValues();
  const headers = caseData[0].map(h => typeof h === "string" ? h.trim() : h);

  // 欄位位置
  const typeIdx = headers.indexOf("個案類型");
  const tbIdx = headers.indexOf("總編號");
  const deadlineIdx = headers.indexOf("銷案期限");
  const advisorIdx = headers.indexOf("輔導員");
  const applicantIdx = headers.indexOf("申請人");
  const statusIdx = headers.indexOf("審核狀態");  // 新增判斷
  if (typeIdx < 0 || tbIdx < 0 || deadlineIdx < 0 || advisorIdx < 0 || applicantIdx < 0 || statusIdx < 0) return;

  // 準備所有要通知的人
  const notifications = {}; // {名字: [訊息, ...]}

  // 今天日期
  const now = new Date();
  now.setHours(0,0,0,0);

  for (let i = 1; i < caseData.length; i++) {
    const row = caseData[i];
    const status = row[statusIdx];
    if (status === "✅ 通過") continue; // 只通知尚未通過的個案
    const caseType = row[typeIdx];
    const tbNumber = row[tbIdx];
    const deadlineStr = row[deadlineIdx];
    const advisorName = row[advisorIdx];
    const applicantName = row[applicantIdx];
    if (!deadlineStr) continue;

    // 日期解析
    let deadlineDate;
    if (deadlineStr instanceof Date) {
      deadlineDate = new Date(deadlineStr.getFullYear(), deadlineStr.getMonth(), deadlineStr.getDate());
    } else if (typeof deadlineStr === "string" && deadlineStr.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
      const [y, m, d] = deadlineStr.split('/').map(Number);
      deadlineDate = new Date(y, m - 1, d);
    } else {
      continue;
    }

    // 計算相差天數
    const diffDays = Math.floor((deadlineDate - now) / (1000 * 60 * 60 * 24));

    let msg = "";
    if (diffDays <= 7 && diffDays >= 1) {
      msg = `⏱️ ${caseType}個案 ${tbNumber} 將於 ${diffDays} 天後到期`;
    } else if (diffDays === 0) {
      msg = `⏱️ ${caseType}個案 ${tbNumber} 將於今天到期`;
    } else if (diffDays < 0) {
      msg = `⏱️ ${caseType}個案 ${tbNumber} 已逾期 ${Math.abs(diffDays)} 天`;
    }

    if (msg) {
      [advisorName, applicantName].forEach(name => {
        if (!name) return;
        if (!notifications[name]) notifications[name] = [];
        notifications[name].push(msg);
      });
    }
  }

  // 寫入帳號資訊
  const accountData = accountSheet.getDataRange().getValues();
  const accHeaders = accountData[0].map(h => typeof h === "string" ? h.trim() : h);
  const nameIdx = accHeaders.indexOf("姓名");
  const msgIdx = accHeaders.indexOf("訊息通知");
  for (let i = 1; i < accountData.length; i++) {
    const userName = accountData[i][nameIdx];
    const newMsgs = notifications[userName] 
      ? notifications[userName].join('\n') 
      : "";
    accountSheet.getRange(i+1, msgIdx+1).setValue(newMsgs || "");
  }
}

function getFieldIndexMap(sheet) {
  const titles = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < titles.length; i++) {
    map[titles[i]] = i;
  }
  return map;
}

function registerAccount(info) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('帳號資訊');
  const fieldMap = getFieldIndexMap(sheet);
  const accountCol = fieldMap['公務帳號'];
  const accountData = sheet.getRange(2, accountCol + 1, sheet.getLastRow() - 1, 1)
    .getValues().flat().map(e => String(e).toLowerCase());
  if (accountData.includes(info.acc.toLowerCase())) {
    return { success: false };
  }
  const row = [];
  for (let key in fieldMap) {
    switch (key) {
      case '單位名稱':
        row[fieldMap[key]] = info.unit;
        break;
      case '公務帳號':
        row[fieldMap[key]] = info.acc;
        break;
      case '密碼':
        row[fieldMap[key]] = "'" + String(info.pwd);
        break;
      case '姓名':
        row[fieldMap[key]] = info.name;
        break;
      case '分機':
        row[fieldMap[key]] = info.ext;
        break;
      default:
        row[fieldMap[key]] = '';
    }
  }

  GmailApp.sendEmail({
    to: "AF7422@ntpc.gov.tw",
    subject: "[銷案審核助手] 有新帳號申請待審核",
    htmlBody:
      "<b>新帳號申請通知：</b><br><br>" +
      "單位：" + info.unit + "<br>" +
      "姓名：" + info.name + "<br>" +
      "公務帳號：" + info.acc + "<br>" +
      "分機：" + info.ext + "<br><br>" +
      "請盡快登入 Google Sheet 進行審核！"
  });

  sheet.appendRow(row);
  return { success: true };
}

//計算銷案期限
function calcDisposalDeadline(caseType, endDate, reason) {
  var date = new Date(endDate);
  var days = 0;
  if (caseType === "LTBI") {
    days = 14;
  } else if (caseType === "TB") {
    if (["完成管理", "其他完治", "其他（視同結果失落）"].indexOf(reason) >= 0) {
      days = 30;
    } else {
      days = 90;
    }
  }
  date.setDate(date.getDate() + days);
  var pad = function(n){return (n<10?"0":"")+n;};
  return date.getFullYear() + "/" + pad(date.getMonth()+1) + "/" + pad(date.getDate());
}

  // ✅ Part 2: 後端 - 上傳圖片至雲端
  // ➜ 需部署成 GAS 並公開給前端使用（doSubmitEditorPatch 中呼叫）

  /**
   * @param {string} tbNumber TB總編號
   * @param {string} endDate 結束治療日 yyyy-mm-dd
   * @param {Object[]} items 含 imageBase64 欄位
   * @return {Object[]} 附加 imageUrl 後的資料
   */
  function uploadImageBase64ToDrive(base64, folderName, filename) {
    if (!base64 || typeof base64 !== "string" || !base64.includes(',')) {
      throw new Error("無效的 base64 圖片資料");
    }

    const contentType = base64.startsWith('data:image/jpeg') ? 'image/jpeg' : 'image/png';
    const ext = contentType === 'image/jpeg' ? '.jpg' : '.png';
    const name = filename ? filename + ext : 'image' + ext;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64.split(',')[1]),
      contentType,
      name
    );

    // === 取得目前試算表與父資料夾 ===
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();

    // === 確保存在「試算表名稱_上傳圖片」的主資料夾 ===
    const mainFolderName = ss.getName() + "_上傳圖片";
    let mainFolder;
    const mainFolders = parentFolder.getFoldersByName(mainFolderName);
    mainFolder = mainFolders.hasNext() ? mainFolders.next() : parentFolder.createFolder(mainFolderName);

    // === 從 folderName 解析 yyyy-mm ===
    const match = folderName.match(/^(\d{4})-(\d{2})/);
    if (!match) {
      throw new Error("folderName 格式錯誤，應為 yyyy-mm-dd-xxx");
    }

    const year = match[1];   // yyyy
    const month = match[2];  // mm

    // === 年資料夾 ===
    const yearFolderName = `${year}年`;
    let yearFolder;
    const yearFolders = mainFolder.getFoldersByName(yearFolderName);
    yearFolder = yearFolders.hasNext() 
      ? yearFolders.next() 
      : mainFolder.createFolder(yearFolderName);

    // === 月資料夾 ===
    const monthFolderName = `${year}-${month}`;
    let monthFolder;
    const monthFolders = yearFolder.getFoldersByName(monthFolderName);
    monthFolder = monthFolders.hasNext()
      ? monthFolders.next()
      : yearFolder.createFolder(monthFolderName);

    // === 最終子資料夾（案件）===
    let targetFolder;
    const subFolders = monthFolder.getFoldersByName(folderName);
    targetFolder = subFolders.hasNext()
      ? subFolders.next()
      : monthFolder.createFolder(folderName);

    // === 建立檔案並設分享權限 ===
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ✅ 回傳可嵌入的圖片網址
    return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1600`;
  }

  /**
   * 共用函式：寫入編輯紀錄（含 LockService 鎖定機制）
   * 用於 submitDisposalRequest、submitReviewDecision、submitEditorPatch 等流程
   * 自動處理表頭對應與 JSON 轉文字格式，安全防止多人同時寫入衝突
   */
  function writeDisposalLog(data) {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000); // 最多等 5 秒取得鎖
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("編輯紀錄");
      const headers = sheet.getDataRange().getValues()[0];
      const row = [];

      const now = new Date();
      const tz = Session.getScriptTimeZone();
      const formattedNow = Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm:ss");

      function lossItemJsonToText(jsonStr) {
        let arr = [];
        try {
          arr = JSON.parse(jsonStr || '[]');
          if (!Array.isArray(arr)) arr = [];
        } catch (e) {
          arr = [];
        }
        return arr.map(item => {
          let checkMark = item.checked ? "🗹" : "☐";
          let selectText = item.select ? `【${item.select}】` : "";
          let remark = item.remark ? item.remark : "";
          return `${checkMark} ${selectText}${remark}`;
        }).join('\n');
      }

      headers.forEach(header => {
        switch (header) {
          case "編輯時間":
            row.push(formattedNow); break;
          case "單位":
            row.push(data.unit || ""); break;
          case "用戶名稱":
            row.push(data.userName || data.reviewerName || data.patchName || ""); break;
          case "個案類型":
            row.push(data.caseType || ""); break;
          case "總編號":
            row.push("'" + String(data.tb).replace(/^'+/, "").trim());
            break;
          case "個案姓名":
            row.push(data.caseName || ""); break;
          case "不可逆缺失項(文本)":
            row.push(lossItemJsonToText(data.unreversible)); break;
          case "可逆缺失項(文本)":
            row.push(lossItemJsonToText(data.reversible)); break;
          case "建議事項(文本)":
            row.push(lossItemJsonToText(data.suggestion)); break;
          case "審核狀態":
            row.push(data.status || "📨 申請"); break;
          case "備註":              
            row.push(data.note || ""); break;
          default:
            row.push("");
        }
      });

      sheet.appendRow(row);
      const lastRow = sheet.getLastRow();
      const tbIndex = headers.indexOf("總編號");
      if (tbIndex !== -1) {
        sheet.getRange(lastRow, tbIndex + 1).setNumberFormat("@");
      }
    } catch (e) {
      Logger.log("❌ 寫入編輯紀錄失敗：" + e.message);
    } finally {
      lock.releaseLock();
    }
  }

//寄送密碼給使用者
function sendPasswordEmail(acc, name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("帳號資訊");
  const data = sheet.getDataRange().getValues();

  const headers = data[0];
  const accIdx = headers.indexOf("公務帳號");
  const nameIdx = headers.indexOf("姓名");
  const pwdIdx = headers.indexOf("密碼");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    if (
      String(row[accIdx]).toLowerCase() === acc.toLowerCase() &&
      String(row[nameIdx]).trim() === name
    ) {
      const pwd = row[pwdIdx];
      const email = acc + "@ntpc.gov.tw";

      const subject = "銷案審核助手－密碼通知";
      const body =
`您好，

您的密碼資訊如下：

${pwd}

請妥善保管您的密碼。

（本系統為內部工具，請勿外流）`;

      GmailApp.sendEmail(email, subject, body);
      return;
    }
  }
}

