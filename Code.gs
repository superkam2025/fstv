/**
 * =========================================================================
 * 綜合管理平台：後端核心邏輯 (Code.gs) 完整整合修正版
 * =========================================================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('校園電視台管理平台')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Members');
}

// ==========================================
// 1. 用戶登入驗證 (已強化防禦機制)
// ==========================================
function login(studentId, password) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    
    var searchId = studentId ? studentId.toString().trim() : "";
    var searchPwd = password ? password.toString() : "";

    for (var i = 1; i < data.length; i++) {
      var sheetId = data[i][0] !== undefined ? data[i][0].toString().trim() : "";
      var sheetPwd = data[i][1] !== undefined ? data[i][1].toString() : "";
      
      // 支援前端靜態重整機制 REFRESH_DATA
      if (sheetId === searchId && (searchPwd === "REFRESH_DATA" || sheetPwd === searchPwd)) {
        var role = data[i][2] || 'pending'; 
        if (role === 'pending') {
          return { success: false, message: '您的帳號正在等待管理員審批，請稍後再試！' };
        }

        // 安全讀取第 K 欄 (索引 10) 的成長檔案 JSON
        var portfolioRaw = (data[i].length > 10 && data[i][10]) ? data[i][10].toString().trim() : "";
        var portfolioData = { learnHours: 0, activityHours: 0, contests: 0, skills: 0, milestones: [] };
        
        if (portfolioRaw !== "") {
          try {
            portfolioData = JSON.parse(portfolioRaw);
          } catch(e) {
            // 萬一解析出錯，維持預設結構，絕不崩潰
          }
        }

        return {
          success: true,
          role: role, 
          name: data[i][3] || '未命名用戶',
          message: '驗證成功',
          memberData: {
            studentId: sheetId,
            name: data[i][3] || '未命名用戶',
            className: data[i][4] || '',
            photo: data[i][5] || 'https://via.placeholder.com/150',
            phone: data[i][6] || '',
            parentName: data[i][7] || '',
            parentPhone: data[i][8] || '',
            relationship: data[i][9] || '',
            portfolio: portfolioData
          }
        };
      }
    }
    return { success: false, message: '學號或密碼錯誤！' };
  } catch(err) {
    return { success: false, message: '後端驗證核心異常：' + err.toString() };
  }
}

// ==========================================
// 2. 管理員更新成員成長檔案 (彈性索引修正)
// ==========================================
function updateMemberPortfolio(studentId, newPortfolioObj) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var targetId = studentId ? studentId.toString().trim() : "";

    for (var i = 1; i < data.length; i++) {
      var sheetId = data[i][0] !== undefined ? data[i][0].toString().trim() : "";
      if (sheetId === targetId) {
        var jsonString = JSON.stringify(newPortfolioObj);
        sheet.getRange(i + 1, 11).setValue(jsonString); // 寫入第 11 欄 (K欄)
        return { success: true, message: '✅ 成長檔案已更新成功！' };
      }
    }
    return { success: false, message: '❌ 找不到該成員！' };
  } catch(e) {
    return { success: false, message: '更新失敗：' + e.toString() };
  }
}

// 儲存 Base64 圖片至雲端硬碟
function saveImageToDrive(base64Data, fileName) {
  try {
    var folderName = "校園電視台_成員相片";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var blob = Utilities.newBlob(byteCharacters, type, fileName);
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return "https://drive.google.com/thumbnail?sz=w300&id=" + file.getId();
  } catch(e) {
    return 'https://via.placeholder.com/150'; 
  }
}

// ==========================================
// 3. 成員註冊功能
// ==========================================
function registerMember(member) {
  try {
    if (!member) return { success: false, message: '接收到的資料為空值！' };
    
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var newId = member.studentId ? member.studentId.toString().trim() : "";
    
    for (var i = 1; i < data.length; i++) {
      var sheetId = data[i][0] !== undefined ? data[i][0].toString().trim() : "";
      if (sheetId === newId) return { success: false, message: '此帳號已被註冊！' };
    }
    
    var photoUrl = 'https://via.placeholder.com/150'; 
    if (member.photoBase64) {
      var fileName = newId + "_" + (member.name || "member");
      photoUrl = saveImageToDrive(member.photoBase64, fileName);
    }
    
    var initialPortfolio = { learnHours: 0, activityHours: 0, contests: 0, skills: 0, milestones: [] };
    
    sheet.appendRow([
      newId,
      member.password,
      'pending', 
      member.name,
      member.className,
      photoUrl, 
      member.phone,
      member.parentName,
      member.parentPhone,
      member.relationship,
      JSON.stringify(initialPortfolio)
    ]);
    
    return { success: true, message: '註冊成功！請等待管理員審批後再登入。' };
  } catch(err) {
    return { success: false, message: '註冊連線程序失敗：' + err.toString() };
  }
}

// 查詢所有成員
function getAllMembers() {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    for (var i = 1; i < data.length; i++) {
      var portfolioRaw = (data[i].length > 10 && data[i][10]) ? data[i][10].toString().trim() : "";
      var portfolioData = { learnHours: 0, activityHours: 0, contests: 0, skills: 0, milestones: [] };
      if (portfolioRaw) {
        try { portfolioData = JSON.parse(portfolioRaw); } catch(e) {}
      }

      list.push({
        studentId: data[i][0] !== undefined ? data[i][0].toString().trim() : "",
        role: data[i][2] || 'member',
        name: data[i][3] || '未命名',
        className: data[i][4] || '-',
        photo: data[i][5] || 'https://via.placeholder.com/150',
        phone: data[i][6] || '-',
        parentName: data[i][7] || '-',
        parentPhone: data[i][8] || '-',
        relationship: data[i][9] || '-',
        portfolio: portfolioData
      });
    }
    return list;
  } catch (e) {
    return [];
  }
}

function deleteMember(studentId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var targetId = studentId ? studentId.toString().trim() : "";

    for (var i = data.length - 1; i >= 1; i--) {
      var sheetId = data[i][0] !== undefined ? data[i][0].toString().trim() : "";
      if (sheetId === targetId) {
        sheet.deleteRow(i + 1); 
        return { success: true, message: '✅ 成員已成功刪除！' };
      }
    }
    return { success: false, message: '❌ 找不到該成員或刪除失敗！' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function approveMember(studentId) {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var targetId = studentId ? studentId.toString().trim() : "";

    for (var i = 1; i < data.length; i++) {
      var sheetId = data[i][0] !== undefined ? data[i][0].toString().trim() : "";
      if (sheetId === targetId) {
        sheet.getRange(i + 1, 3).setValue('member'); 
        return { success: true, message: '✅ 已成功通過該成員的註冊申請！' };
      }
    }
    return { success: false, message: '❌ 找不到該成員！' };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// 器材借用登記提交
function submitEquipmentBorrow(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "器材借用";
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["登記時間戳記", "身份類別", "借用人姓名", "班級 / 單位", "聯絡電話", "借用器材", "借用日期", "借用開始時間", "預計借用時長"]);
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#fff2cc").setHorizontalAlignment("center");
    }
    
    var timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([timestamp, data.identity, data.name, data.className, data.phone, data.equipment, data.date, data.time, data.duration]);
    return { success: true, message: "🎉 器材借用登記成功！請準時前往電視台領取器材。" };
  } catch (error) {
    return { success: false, message: "❌ 登記失敗，原因：" + error.toString() };
  }
}

function getBorrowRecords() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("器材借用");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    var records = [];
    for (var i = 1; i < data.length; i++) {
      records.push({
        timestamp: data[i][0], identity: data[i][1], name: data[i][2], className: data[i][3],
        phone: data[i][4], equipment: data[i][5], date: data[i][6], time: data[i][7], duration: data[i][8]
      });
    }
    return records.reverse();
  } catch (e) { return []; }
}

function submitActivityDuty(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "活動值日";
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["發佈時間戳記", "活動名稱", "活動日期", "活動時間", "負責組長", "拍攝者1", "拍攝者2", "拍攝者3", "剪輯者"]);
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
    }
    var timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([timestamp, data.activityName, data.date, data.time, data.leader, data.cam1, data.cam2, data.cam3, data.editor]);
    return { success: true, message: "🎉 活動值日工作發佈成功！全體成員已可實時查看最新分工。" };
  } catch (error) { return { success: false, message: "❌ 發佈失敗：" + error.toString() }; }
}

function submitCheckIn(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var checkInSheet = ss.getSheetByName("值日簽到"); 
    if (!checkInSheet) return { success: false, message: "錯誤：找不到 [值日簽到] 工作表！" };
    
    var now = new Date();
    var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
    checkInSheet.appendRow([payload.activityName, payload.date, payload.memberName, timeStr, payload.location]);

    var memberSheet = ss.getSheetByName("Members"); 
    if (!memberSheet) return { success: false, message: "錯誤：找不到 [Members] 工作表！" };
    
    var data = memberSheet.getDataRange().getValues();
    var updated = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === payload.memberName) {
        var portfolioStr = data[i].length > 10 ? data[i][10] : "";
        var portfolioObj = { learnHours: 0, activityHours: 0, contests: 0, skills: 0, milestones: [] };
        if (portfolioStr) {
          try { portfolioObj = JSON.parse(portfolioStr); } catch(e) {}
        }
        
        portfolioObj.activityHours = (portfolioObj.activityHours || 0) + 1;
        if (!portfolioObj.milestones) portfolioObj.milestones = [];
        
        var currentMonth = payload.date ? payload.date.substring(0, 7) : Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");
        if (payload.location === "💻 剪輯已完成") {
          portfolioObj.milestones.unshift({ date: currentMonth, title: "🎬 完成影片剪輯工作", desc: "負責「" + payload.activityName + "」的後期剪輯與回報。", type: "normal" });
        } else {
          portfolioObj.milestones.unshift({ date: currentMonth, title: "🎥 參與活動現場拍攝值日", desc: "於「" + payload.activityName + "」擔任現場工作，地點：" + payload.location, type: "normal" });
        }
        
        memberSheet.getRange(i + 1, 11).setValue(JSON.stringify(portfolioObj));
        updated = true;
        break; 
      }
    }
    return updated ? { success: true, message: "✅ 回報與簽到成功！活動次數已自動增加 1 次！" } : { success: true, message: "⚠️ 簽到成功，但未在 [Members] 中匹配到同名帳號，未能同步累計時數。" };
  } catch (error) { return { success: false, message: "後端處理失敗：" + error.toString() }; }
}

// ==========================================
// 4. 獲取活動值日清單 (🏆 修正安全性與參數對齊崩潰點)
// ==========================================
function getActivityDuties(memberName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone(); 
    
    var dutySheet = ss.getSheetByName("活動值日");
    if (!dutySheet) return [];
    
    var dutyRows = dutySheet.getDataRange().getValues();
    if (dutyRows.length <= 1) return [];
    
    var checkInSheet = ss.getSheetByName("值日簽到");
    var checkInMap = {}; 
    if (checkInSheet) {
      var ciRange = checkInSheet.getDataRange();
      var ciRows = ciRange.getValues();
      var ciDisplays = ciRange.getDisplayValues(); 
      
      for (var j = 1; j < ciRows.length; j++) {
        var ciDate = ciRows[j][1];
        if (ciDate instanceof Date) ciDate = Utilities.formatDate(ciDate, tz, "yyyy-MM-dd");
        var ciKey = ciRows[j][0].toString().trim() + "_" + ciDate + "_" + ciRows[j][2].toString().trim();
        var timeStr = ciDisplays[j][3] ? ciDisplays[j][3].toString().trim() : "";
        checkInMap[ciKey] = { time: timeStr, location: ciRows[j][4] ? ciRows[j][4].toString().trim() : "" };
      }
    }
    
    var todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var duties = [];
    var targetName = memberName ? memberName.toString().trim() : ""; // 安全容錯：避免 undefined 呼叫 .toString() 造成崩潰

    for (var i = 1; i < dutyRows.length; i++) {
      if (!dutyRows[i][1]) continue;
      var activityName = dutyRows[i][1].toString().trim();
      var dateVal = dutyRows[i][2];
      if (dateVal instanceof Date) dateVal = Utilities.formatDate(dateVal, tz, "yyyy-MM-dd");
      
      var cam1 = dutyRows[i][5] ? dutyRows[i][5].toString().trim() : "";
      var cam2 = dutyRows[i][6] ? dutyRows[i][6].toString().trim() : "";
      var cam3 = dutyRows[i][7] ? dutyRows[i][7].toString().trim() : "";
      var editor = dutyRows[i][8] ? dutyRows[i][8].toString().trim() : "";
      
      var assignedPeople = [cam1, cam2, cam3, editor].filter(Boolean);
      if (assignedPeople.length === 0) continue; 
      
      var personnelStatus = [];
      var roles = [
        { name: cam1, label: "拍攝者 1" }, { name: cam2, label: "拍攝者 2" },
        { name: cam3, label: "拍攝者 3" }, { name: editor, label: "剪輯者" }
      ];
      
      roles.forEach(function(r) {
        if (r.name) {
          var key = activityName + "_" + dateVal + "_" + r.name;
          var hasCheckedIn = !!checkInMap[key];
          personnelStatus.push({
            roleLabel: r.label, memberName: r.name, hasCheckedIn: hasCheckedIn,
            time: hasCheckedIn ? checkInMap[key].time : "", location: hasCheckedIn ? checkInMap[key].location : ""
          });
        }
      });
      
      // 💡 移除原先在後端直接 continue 的下架邏輯，改由前端進行動態消失比對，確保看板資料流完整傳遞
      if (targetName !== "") {
        if (cam1 !== targetName && cam2 !== targetName && cam3 !== targetName && editor !== targetName) {
          continue; 
        }
      }
      
      var myKey = activityName + "_" + dateVal + "_" + targetName;
      var myCheckInInfo = checkInMap[myKey] || null;
      var isToday = (dateVal === todayStr);
      
      duties.push({
        activityName: activityName, date: dateVal,
        time: dutyRows[i][3] ? dutyRows[i][3].toString() : "",
        leader: dutyRows[i][4] ? dutyRows[i][4].toString() : "",
        cam1: cam1, cam2: cam2, cam3: cam3, editor: editor,
        isToday: isToday, myCheckIn: myCheckInInfo, personnelStatus: personnelStatus
      });
    }
    return duties.reverse();
  } catch (error) {
    Logger.log("getActivityDuties Error: " + error.toString());
    return []; // 確保發生例外時依然能安全回傳陣列
  }
}

function getAllMemberNames() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Members");
    if (!sheet) return { success: false, memberList: [], message: "找不到 Members 工作表" };
    
    var data = sheet.getDataRange().getValues();
    var names = [];
    for (var i = 1; i < data.length; i++) {
      var name = data[i][3]; var role = data[i][2]; 
      if (name && name.toString().trim() !== "" && role !== "admin") {
        names.push(name.toString().trim());
      }
    }
    return { success: true, memberList: names };
  } catch(e) { return { success: false, memberList: [], message: e.toString() }; }
}

// 管理員更新檔案（升級版：支援違規記錄整合）
function updateMemberPortfolioByAdmin(adminPayload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName("Members");
    if (!memberSheet) return { success: false, message: "錯誤：找不到 [Members] 工作表！" };
    
    var data = memberSheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] === adminPayload.memberName) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) return { success: false, message: "錯誤：找不到該成員！" };
    
    var oldJsonStr = data[targetRow - 1].length > 10 ? data[targetRow - 1][10] : "";
    var portfolioObj = { learnHours: 0, activityHours: 0, contests: 0, skills: 0, milestones: [] };
    if (oldJsonStr) {
      try { portfolioObj = JSON.parse(oldJsonStr); } catch(e) {}
    }
    if (!portfolioObj.milestones) portfolioObj.milestones = [];
    
    // 1. 處理原本的時數變更
    if (adminPayload.learnHours !== undefined && adminPayload.learnHours.trim() !== "") {
      portfolioObj.learnHours = Number(adminPayload.learnHours) || 0;
    }
    if (adminPayload.activityHours !== undefined && adminPayload.activityHours.trim() !== "") {
      portfolioObj.activityHours = Number(adminPayload.activityHours) || 0;
    }
    
    var now = new Date();
    var currentMonthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");
    var changeLogText = "管理員手動調整數據。";
    
    // 2. 處理原本的比賽變更
    if (adminPayload.newContestTitle && adminPayload.newContestTitle.trim() !== "") {
      portfolioObj.contests = (portfolioObj.contests || 0) + 1;
      portfolioObj.milestones.unshift({ date: currentMonthStr, title: "🏆 " + adminPayload.newContestTitle.trim(), desc: adminPayload.newContestDesc ? adminPayload.newContestDesc.trim() : "參加比賽項目獲取紀錄。", type: "award" });
      changeLogText += " 新增比賽：" + adminPayload.newContestTitle;
    }
    
    // 3. 處理原本的技能變更
    if (adminPayload.newSkillTitle && adminPayload.newSkillTitle.trim() !== "") {
      portfolioObj.skills = (portfolioObj.skills || 0) + 1;
      portfolioObj.milestones.unshift({ date: currentMonthStr, title: "⚡ 獲得新技能：" + adminPayload.newSkillTitle.trim(), desc: adminPayload.newSkillDesc ? adminPayload.newSkillDesc.trim() : "經後台考核通過。", type: "normal" });
      changeLogText += " 新增技能：" + adminPayload.newSkillTitle;
    }

    // 🔥 4. 全新功能：處理違規記錄寫入個人時間軸
    if (adminPayload.violationType && adminPayload.violationType.trim() !== "") {
      var vType = adminPayload.violationType.trim(); // 預期傳入："遲到"、"缺席" 或 "其它違規事項"
      var vDesc = adminPayload.violationDesc ? adminPayload.violationDesc.trim() : "未填寫詳細說明。";
      var emoji = "⚠️";
      
      if (vType === "遲到") emoji = "🕒";
      if (vType === "缺席") emoji = "❌";
      
      // 將記錄塞入里程碑陣列的最前方 (時間軸最新事件)
      portfolioObj.milestones.unshift({
        date: currentMonthStr,
        title: emoji + " 違規記錄：" + vType,
        desc: vDesc,
        type: "violation" // 定義新狀態，方便前端調整 CSS 樣式（例如標記為紅色）
      });
      changeLogText += " 登記違規：" + vType + " (" + vDesc + ")";
    }
    
    // 寫回試算表 K 欄
    memberSheet.getRange(targetRow, 11).setValue(JSON.stringify(portfolioObj));
    
    // 寫入系統日誌
    var logSheet = ss.getSheetByName("個人成長里程碑");
    if (logSheet) {
      var fullTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      logSheet.appendRow([fullTimeStr, adminPayload.memberName, "管理員後台變更", changeLogText]);
    }
    return { success: true, message: "✅ [" + adminPayload.memberName + "] 的數據（含違規紀錄）已成功同步！" };
  } catch(error) { return { success: false, message: "後端更新失敗：" + error.toString() }; }
}

