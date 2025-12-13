// LeaveManagement.gs - 請假系統後端邏輯（修正版）

// ==================== 所有常數都在 Constants.gs 中定義 ====================
// 不需要在這裡重複定義 SHEET_LEAVE_RECORDS、SHEET_LEAVE_BALANCE、EMPLOYEE_COL
// 所有 .gs 檔案會共享 Constants.gs 中的常數

/**
 * 計算特休假天數（根據台灣勞基法）
 * @param {Date} hireDate - 到職日期
 * @param {Date} calculateDate - 計算日期（預設為今天）
 * @return {number} 特休假天數
 */
function calculateAnnualLeave_(hireDate, calculateDate = new Date()) {
  const hire = new Date(hireDate);
  const calc = new Date(calculateDate);
  
  // 計算工作月數
  const monthsDiff = (calc.getFullYear() - hire.getFullYear()) * 12 
                   + (calc.getMonth() - hire.getMonth());
  
  // 計算工作年數
  const years = Math.floor(monthsDiff / 12);
  
  // 未滿6個月：0天
  if (monthsDiff < 6) return 0;
  
  // 6個月以上未滿1年：3天
  if (monthsDiff >= 6 && monthsDiff < 12) return 3;
  
  // 1年以上未滿2年：7天
  if (years >= 1 && years < 2) return 7;
  
  // 2年以上未滿3年：10天
  if (years >= 2 && years < 3) return 10;
  
  // 3年以上未滿5年：14天
  if (years >= 3 && years < 5) return 14;
  
  // 5年以上未滿10年：15天
  if (years >= 5 && years < 10) return 15;
  
  // 10年以上：每年15天 + 每多一年加1天（最高30天）
  if (years >= 10) {
    const additionalDays = Math.min(years - 10, 15);
    return 15 + additionalDays;
  }
  
  return 0;
}

/**
 * 初始化或更新員工假期額度
 * @param {string} userId - 員工 ID
 * @param {Date} hireDate - 到職日期
 */
function initializeLeaveBalance_(userId, hireDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEAVE_BALANCE);
  
  // 如果工作表不存在，創建它
  if (!sheet) {
    const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_LEAVE_BALANCE);
    newSheet.appendRow([
      '員工ID', '姓名', '到職日期', '年度',
      '特休假', '病假', '事假', '婚假', '喪假',
      '產假', '陪產假', '家庭照顧假', '生理假',
      '更新時間'
    ]);
    return initializeLeaveBalance_(userId, hireDate);
  }
  
  const values = sheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  
  // 查找現有記錄
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === userId && values[i][3] === currentYear) {
      // 記錄已存在，更新特休假
      const annualLeave = calculateAnnualLeave_(hireDate);
      sheet.getRange(i + 1, 5).setValue(annualLeave);
      sheet.getRange(i + 1, 14).setValue(new Date());
      Logger.log(`✅ 更新員工 ${userId} 的特休假：${annualLeave} 天`);
      return;
    }
  }
  
  // 記錄不存在，新增
  const employee = findEmployeeByLineUserId_(userId);
  const annualLeave = calculateAnnualLeave_(hireDate);
  
  sheet.appendRow([
    userId,                    // 員工ID
    employee.name || '',       // 姓名
    hireDate,                  // 到職日期
    currentYear,               // 年度
    annualLeave,               // 特休假
    30,                        // 病假（每年30天）
    14,                        // 事假（每年14天）
    8,                         // 婚假（8天）
    0,                         // 喪假（依親屬關係）
    56,                        // 產假（8週=56天）
    7,                         // 陪產假（7天）
    7,                         // 家庭照顧假（7天）
    12,                        // 生理假（每月1天，一年12天）
    new Date()                 // 更新時間
  ]);
  
  Logger.log(`✅ 新增員工 ${userId} 的假期額度，特休假：${annualLeave} 天`);
}

/**
 * 取得員工假期餘額
 * @param {string} sessionToken - Session Token
 * @return {Object} 假期餘額資訊
 */
function getLeaveBalance(sessionToken) {
  const session = checkSession_(sessionToken);
  if (!session.ok) return { ok: false, code: session.code };
  
  const userId = session.user.userId;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEAVE_BALANCE);
  
  if (!sheet) {
    return { 
      ok: false, 
      code: 'ERR_LEAVE_SYSTEM_NOT_INITIALIZED' 
    };
  }
  
  const values = sheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  
  // 查找當年度的假期額度
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === userId && values[i][3] === currentYear) {
      return {
        ok: true,
        balance: {
          ANNUAL_LEAVE: values[i][4],        // 特休假
          SICK_LEAVE: values[i][5],          // 病假
          PERSONAL_LEAVE: values[i][6],      // 事假
          MARRIAGE_LEAVE: values[i][7],      // 婚假
          BEREAVEMENT_LEAVE: values[i][8],   // 喪假
          MATERNITY_LEAVE: values[i][9],     // 產假
          PATERNITY_LEAVE: values[i][10],    // 陪產假
          FAMILY_CARE_LEAVE: values[i][11],  // 家庭照顧假
          MENSTRUAL_LEAVE: values[i][12]     // 生理假
        }
      };
    }
  }
  
  // 如果找不到記錄，可能需要初始化
  return { 
    ok: false, 
    code: 'ERR_NO_LEAVE_BALANCE',
    msg: '找不到假期額度，請聯繫管理員初始化系統'
  };
}

/**
 * 取得已核准的請假記錄
 * @param {string} monthParam - 月份，格式 "YYYY-MM"
 * @param {string} userIdParam - 員工ID
 */
// function getApprovedLeaveRecords(monthParam, userIdParam) {
//   try {
//     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請假紀錄');
    
//     if (!sheet) {
//       Logger.log('⚠️ 找不到請假紀錄工作表');
//       return [];
//     }
    
//     const values = sheet.getDataRange().getValues();
    
//     if (values.length <= 1) {
//       return [];
//     }
    
//     const headers = values[0];
//     const employeeIdCol = headers.indexOf('員工ID');
//     const startDateCol = headers.indexOf('開始日期');
//     const endDateCol = headers.indexOf('結束日期');
//     const leaveTypeCol = headers.indexOf('假別');
//     const daysCol = headers.indexOf('天數');
//     const statusCol = headers.indexOf('狀態');
//     const reasonCol = headers.indexOf('原因');
//     const nameCol = headers.indexOf('姓名');
    
//     const leaveRecords = [];
    
//     values.slice(1).forEach(row => {
//       const employeeId = row[employeeIdCol];
//       const startDate = formatDate(row[startDateCol]);
//       const endDate = formatDate(row[endDateCol]);
//       const status = String(row[statusCol]).trim();
//       const leaveType = String(row[leaveTypeCol]).trim();
      
//       // 只取已核准的記錄
//       if (status !== 'APPROVED') return;
      
//       // 檢查月份是否匹配
//       if (!startDate.startsWith(monthParam)) return;
      
//       // 檢查員工ID是否匹配
//       if (userIdParam && employeeId !== userIdParam) return;
      
//       // 生成請假期間的每一天（如果是多天請假）
//       const start = new Date(startDate);
//       const end = new Date(endDate);
      
//       for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
//         const dateStr = formatDate(d);
        
//         // 確保日期在查詢月份內
//         if (dateStr.startsWith(monthParam)) {
//           leaveRecords.push({
//             employeeId: employeeId,
//             employeeName: row[nameCol],
//             date: dateStr,
//             leaveType: leaveType,
//             days: parseFloat(row[daysCol]) || 0,
//             status: status,
//             reason: row[reasonCol] || ''
//           });
//         }
//       }
//     });
    
//     Logger.log(`✅ getApprovedLeaveRecords: 找到 ${leaveRecords.length} 筆已核准請假記錄`);
//     return leaveRecords;
    
//   } catch (error) {
//     Logger.log('❌ getApprovedLeaveRecords 錯誤: ' + error);
//     return [];
//   }
// }

/**
 * 👉 修正：取得已核准的請假記錄（根據實際工作表結構）
 * 
 * 實際欄位順序（根據截圖）：
 * A - 申請時間
 * B - 員工ID
 * C - 姓名
 * D - 部門
 * E - 假別
 * F - 開始日期
 * G - 結束日期
 * H - 天數
 * I - 原因
 * J - 狀態
 */
function getApprovedLeaveRecords(monthParam, userIdParam) {
  try {
    Logger.log('📋 getApprovedLeaveRecords 開始');
    Logger.log(`   monthParam: ${monthParam}`);
    Logger.log(`   userIdParam: ${userIdParam}`);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('請假紀錄');
    
    if (!sheet) {
      Logger.log('⚠️ 找不到請假紀錄工作表');
      return [];
    }
    
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      Logger.log('⚠️ 工作表只有標題，沒有資料');
      return [];
    }
    
    Logger.log(`✅ 工作表有 ${values.length - 1} 筆資料`);
    Logger.log('');
    
    // ✅ 根據實際截圖的欄位順序（固定索引）
    const leaveRecords = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // 固定欄位索引（從 0 開始）
      const employeeId = row[1];           // B 欄 (索引 1)
      const employeeName = row[2];         // C 欄 (索引 2)
      const leaveType = row[4];            // E 欄 (索引 4)
      const startDate = row[5];            // F 欄 (索引 5)
      const endDate = row[6];              // G 欄 (索引 6)
      const days = row[7];                 // H 欄 (索引 7)
      const reason = row[8];               // I 欄 (索引 8)
      const status = row[9];               // J 欄 (索引 9) ← 關鍵修正
      
      // 格式化日期
      const formattedStartDate = formatDate(startDate);
      const formattedEndDate = formatDate(endDate);
      
      Logger.log(`第 ${i + 1} 行:`);
      Logger.log(`   員工ID: ${employeeId}`);
      Logger.log(`   狀態: "${status}"`);
      Logger.log(`   開始日期: ${formattedStartDate}`);
      
      // 檢查狀態（只取已核准的）
      if (String(status).trim() !== 'APPROVED') {
        Logger.log(`   ⏭️ 狀態不是 APPROVED，跳過`);
        continue;
      }
      
      // 檢查月份
      if (!formattedStartDate.startsWith(monthParam)) {
        Logger.log(`   ⏭️ 月份不符，跳過`);
        continue;
      }
      
      // 檢查員工ID
      if (userIdParam && employeeId !== userIdParam) {
        Logger.log(`   ⏭️ 員工ID不符，跳過`);
        continue;
      }
      
      Logger.log(`   ✅ 符合條件，加入結果`);
      
      // 生成請假期間的每一天
      const start = new Date(startDate);
      const end = new Date(endDate);
      
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        const dateStr = formatDate(d);
        
        // 確保日期在查詢月份內
        if (dateStr.startsWith(monthParam)) {
          leaveRecords.push({
            employeeId: employeeId,
            employeeName: employeeName,
            date: dateStr,
            leaveType: leaveType,
            days: parseFloat(days) || 0,
            status: status,
            reason: reason || ''
          });
          
          Logger.log(`      📅 加入日期: ${dateStr}`);
        }
      }
      
      Logger.log('');
    }
    
    Logger.log('═══════════════════════════════════════');
    Logger.log(`✅ getApprovedLeaveRecords 完成: 找到 ${leaveRecords.length} 筆記錄`);
    Logger.log('═══════════════════════════════════════');
    
    return leaveRecords;
    
  } catch (error) {
    Logger.log('❌ getApprovedLeaveRecords 錯誤: ' + error);
    Logger.log('   錯誤堆疊: ' + error.stack);
    return [];
  }
}

/**
 * 🧪 測試修正後的函數
 */
function testFixedGetApprovedLeaveRecords() {
  Logger.log('🧪 測試修正後的 getApprovedLeaveRecords');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  const monthParam = '2025-12';
  const userIdParam = 'U68e0ca9d516e63ed15bf9387fad174ac';
  
  const records = getApprovedLeaveRecords(monthParam, userIdParam);
  
  Logger.log('');
  Logger.log('📤 最終結果:');
  Logger.log(`   找到 ${records.length} 筆請假記錄`);
  Logger.log('');
  
  if (records.length > 0) {
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('');
    Logger.log('📋 記錄詳情:');
    records.forEach((rec, i) => {
      Logger.log(`   ${i + 1}. ${rec.date} - ${rec.leaveType} (${rec.days}天)`);
    });
  } else {
    Logger.log('❌ 沒有找到記錄');
    Logger.log('   請檢查上面的除錯訊息');
  }
  
  Logger.log('');
  Logger.log('═══════════════════════════════════════');
}
/**
 * 提交請假申請
 * @param {string} sessionToken - Session Token
 * @param {string} leaveType - 假別
 * @param {string} startDate - 開始日期
 * @param {string} endDate - 結束日期
 * @param {number} days - 請假天數
 * @param {string} reason - 請假原因
 */
function submitLeaveRequest(sessionToken, leaveType, startDate, endDate, days, reason) {
  const session = checkSession_(sessionToken);
  if (!session.ok) return { ok: false, code: session.code };
  
  const user = session.user;
  
  // 驗證假期餘額
  const balanceResult = getLeaveBalance(sessionToken);
  if (!balanceResult.ok) return balanceResult;
  
  const balance = balanceResult.balance;
  const currentBalance = balance[leaveType];
  
  if (currentBalance === undefined) {
    return { ok: false, code: 'ERR_INVALID_LEAVE_TYPE' };
  }
  
  if (currentBalance < days) {
    return { 
      ok: false, 
      code: 'ERR_INSUFFICIENT_LEAVE_BALANCE',
      params: { 
        leaveType: leaveType,
        required: days,
        available: currentBalance
      }
    };
  }
  
  // 建立請假紀錄工作表（如果不存在）
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEAVE_RECORDS);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_LEAVE_RECORDS);
    sheet.appendRow([
      '申請時間', '員工ID', '姓名', '部門', '假別',
      '開始日期', '結束日期', '天數', '原因',
      '狀態', '審核人', '審核時間', '審核意見'
    ]);
  }
  
  // 新增請假記錄
  sheet.appendRow([
    new Date(),              // 申請時間
    user.userId,             // 員工ID
    user.name,               // 姓名
    user.dept,               // 部門
    leaveType,               // 假別
    startDate,               // 開始日期
    endDate,                 // 結束日期
    days,                    // 天數
    reason || '',            // 原因
    'PENDING',               // 狀態（待審核）
    '',                      // 審核人
    '',                      // 審核時間
    ''                       // 審核意見
  ]);
  
  return { 
    ok: true, 
    code: 'LEAVE_SUBMIT_SUCCESS'
  };
}
/**
 * ✅ 取得員工的請假記錄（完整除錯版）
 */
function getEmployeeLeaveRecords(sessionToken) {
  Logger.log('═══════════════════════════════════════');
  Logger.log('📋 getEmployeeLeaveRecords 開始');
  Logger.log('═══════════════════════════════════════');
  
  // Step 1: 驗證 Session
  const session = checkSession_(sessionToken);
  if (!session.ok) {
    Logger.log('❌ Session 驗證失敗');
    return { 
      ok: false,
      success: false,
      code: session.code,
      records: [],
      data: []
    };
  }
  
  Logger.log('✅ Session 驗證成功');
  
  // Step 2: 取得 userId
  const userId = session.user.userId;
  Logger.log('👤 員工ID: ' + userId);
  Logger.log('   型別: ' + typeof userId);
  Logger.log('   長度: ' + (userId ? userId.length : 0));
  Logger.log('');
  
  // Step 3: 取得工作表
  const SHEET_NAME = '請假紀錄';  // ⭐ 直接寫死工作表名稱
  Logger.log('📄 查找工作表: ' + SHEET_NAME);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ 找不到工作表: ' + SHEET_NAME);
    Logger.log('');
    Logger.log('📋 可用的工作表:');
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach((s, i) => {
      Logger.log(`   ${i + 1}. ${s.getName()}`);
    });
    return { 
      ok: true,
      success: true,
      records: [],
      data: []
    };
  }
  
  Logger.log('✅ 工作表存在');
  Logger.log('');
  
  // Step 4: 讀取資料
  const values = sheet.getDataRange().getValues();
  Logger.log(`📊 工作表總行數: ${values.length}`);
  Logger.log('');
  
  if (values.length <= 1) {
    Logger.log('⚠️ 工作表只有標題，沒有資料');
    return {
      ok: true,
      success: true,
      records: [],
      data: []
    };
  }
  
  // Step 5: 顯示標題列
  Logger.log('📋 標題列:');
  values[0].forEach((header, index) => {
    const column = String.fromCharCode(65 + index);
    Logger.log(`   ${column} (${index}): ${header}`);
  });
  Logger.log('');
  
  // Step 6: 遍歷資料
  const records = [];
  
  Logger.log('🔍 開始比對資料...');
  Logger.log('');
  
  for (let i = 1; i < values.length; i++) {
    const rowEmployeeId = values[i][1];  // B 欄
    
    // 轉換並清理
    const cleanRowId = String(rowEmployeeId).trim();
    const cleanUserId = String(userId).trim();
    
    Logger.log(`第 ${i + 1} 行:`);
    Logger.log(`   原始值: "${rowEmployeeId}"`);
    Logger.log(`   清理後: "${cleanRowId}"`);
    Logger.log(`   比對值: "${cleanUserId}"`);
    Logger.log(`   是否匹配: ${cleanRowId === cleanUserId ? '✅ 是' : '❌ 否'}`);
    
    if (cleanRowId === cleanUserId) {
      Logger.log('   🎯 找到匹配記錄！');
      
      const record = {
        applyTime: formatDateTime(values[i][0]),      // A: 申請時間
        employeeName: values[i][2],                   // C: 姓名
        dept: values[i][3],                           // D: 部門
        leaveType: values[i][4],                      // E: 假別
        startDate: formatDate(values[i][5]),          // F: 開始日期
        endDate: formatDate(values[i][6]),            // G: 結束日期
        days: values[i][7],                           // H: 天數
        reason: values[i][8] || '',                   // I: 原因
        status: values[i][9] || 'PENDING',            // J: 狀態
        reviewer: values[i][10] || '',                // K: 審核人
        reviewTime: values[i][11] ? formatDateTime(values[i][11]) : '', // L: 審核時間
        reviewComment: values[i][12] || ''            // M: 審核意見
      };
      
      records.push(record);
      
      Logger.log(`   📝 記錄詳情:`);
      Logger.log(`      假別: ${record.leaveType}`);
      Logger.log(`      日期: ${record.startDate} ~ ${record.endDate}`);
      Logger.log(`      天數: ${record.days}`);
      Logger.log(`      狀態: ${record.status}`);
    }
    
    Logger.log('');
  }
  
  Logger.log('═══════════════════════════════════════');
  Logger.log(`📤 查詢結果: 找到 ${records.length} 筆請假記錄`);
  Logger.log('═══════════════════════════════════════');
  
  // 按申請時間倒序排列
  records.sort((a, b) => new Date(b.applyTime) - new Date(a.applyTime));
  
  // ⭐⭐⭐ 統一回應格式
  return {
    ok: true,
    success: true,
    records: records,
    data: records
  };
}

/**
 * 🧪 測試取得請假記錄（修正版）
 */
function testGetEmployeeLeaveRecords() {
  Logger.log('🧪 測試 getEmployeeLeaveRecords');
  Logger.log('');
  
  const token = '0ed9f26d-d3ec-48b1-8eaf-40c29e725438';  // ⚠️ 替換成你的有效 token
  
  Logger.log('📡 開始查詢...');
  Logger.log('');
  
  const result = getEmployeeLeaveRecords(token);
  
  Logger.log('📤 完整回應:');
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log('');
  
  // ⭐ 修正：檢查兩種格式
  const records = result.records || result.data || [];
  
  Logger.log('🔍 解析結果:');
  Logger.log('   ok: ' + result.ok);
  Logger.log('   success: ' + result.success);
  Logger.log('   records 數量: ' + records.length);
  Logger.log('');
  
  if (records.length > 0) {
    Logger.log('✅✅✅ 找到請假記錄！');
    Logger.log('');
    Logger.log('📋 記錄詳情:');
    
    records.forEach((rec, i) => {
      Logger.log(`   ${i + 1}. ${rec.leaveType} - ${rec.startDate} ~ ${rec.endDate}`);
      Logger.log(`      狀態: ${rec.status}, 天數: ${rec.days}`);
      if (rec.reason) {
        Logger.log(`      原因: ${rec.reason}`);
      }
      Logger.log('');
    });
  } else {
    Logger.log('ℹ️ 該員工目前沒有請假記錄');
    Logger.log('');
    Logger.log('🔍 可能原因:');
    Logger.log('   1. 員工真的沒有申請過請假');
    Logger.log('   2. 工作表「請假紀錄」中的員工ID不匹配');
    Logger.log('   3. 工作表名稱不正確');
  }
}
/**
 * 取得待審核的請假申請（管理員用）
 * @param {string} sessionToken - Session Token
 */
function getPendingLeaveRequests(sessionToken) {
  const session = checkSession_(sessionToken);
  if (!session.ok) return { ok: false, code: session.code };
  
  // 檢查是否為管理員
  if (session.user.dept !== '管理員') {
    return { ok: false, code: 'ERR_NO_PERMISSION' };
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEAVE_RECORDS);
  
  if (!sheet) {
    return { ok: true, requests: [] };
  }
  
  const values = sheet.getDataRange().getValues();
  const requests = [];
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][9] === 'PENDING') {
      requests.push({
        rowNumber: i + 1,
        applyTime: values[i][0],
        userId: values[i][1],
        employeeName: values[i][2],
        dept: values[i][3],
        leaveType: values[i][4],
        startDate: values[i][5],
        endDate: values[i][6],
        days: values[i][7],
        reason: values[i][8]
      });
    }
  }
  
  return { ok: true, requests };
}

/**
 * 審核請假申請（含 LINE 通知）
 * @param {string} sessionToken - Session Token
 * @param {number} rowNumber - 記錄行號
 * @param {string} reviewAction - 審核動作（approve/reject）
 * @param {string} comment - 審核意見
 */
function reviewLeaveRequest(sessionToken, rowNumber, reviewAction, comment) {
  const session = checkSession_(sessionToken);
  if (!session.ok) return { ok: false, code: session.code };
  
  // 檢查是否為管理員
  if (session.user.dept !== '管理員') {
    return { ok: false, code: 'ERR_NO_PERMISSION' };
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEAVE_RECORDS);
  const balanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LEAVE_BALANCE);
  
  if (!sheet) {
    return { ok: false, code: 'ERR_LEAVE_RECORDS_NOT_FOUND' };
  }
  
  // 👉 取得請假記錄（包含所有欄位用於通知）
  const record = sheet.getRange(rowNumber, 1, 1, 13).getValues()[0];
  const userId = record[1];              // 員工ID
  const employeeName = record[2];        // 員工姓名
  const leaveType = record[4];           // 假別
  const startDate = record[5];           // 開始日期
  const endDate = record[6];             // 結束日期
  const days = record[7];                // 天數
  const currentYear = new Date().getFullYear();
  
  // 👉 格式化日期用於通知
  const formattedStartDate = formatDate(startDate);
  const formattedEndDate = formatDate(endDate);
  
  // 👉 翻譯假別名稱
  const leaveTypeName = getLeaveTypeName(leaveType);
  
  // 更新審核狀態
  const status = reviewAction === 'approve' ? 'APPROVED' : 'REJECTED';
  sheet.getRange(rowNumber, 10).setValue(status);           // 狀態
  sheet.getRange(rowNumber, 11).setValue(session.user.name); // 審核人
  sheet.getRange(rowNumber, 12).setValue(new Date());       // 審核時間
  sheet.getRange(rowNumber, 13).setValue(comment || "");    // 審核意見
  
  // 👉 發送 LINE 通知
  try {
    const isApproved = (reviewAction === 'approve');
    notifyLeaveReview(
      userId,
      employeeName,
      leaveTypeName,
      formattedStartDate,
      formattedEndDate,
      days,
      session.user.name,
      isApproved,
      comment || ""
    );
    Logger.log(`📤 已發送請假審核通知給 ${employeeName} (${userId})`);
  } catch (err) {
    Logger.log(`⚠️ LINE 通知發送失敗: ${err.message}`);
    // 通知失敗不影響審核流程
  }
  
  // 如果核准，扣除假期餘額
  if (reviewAction === 'approve' && balanceSheet) {
    const balanceValues = balanceSheet.getDataRange().getValues();
    
    // 假別對應的欄位索引
    const leaveTypeColumns = {
      'ANNUAL_LEAVE': 5,
      'SICK_LEAVE': 6,
      'PERSONAL_LEAVE': 7,
      'MARRIAGE_LEAVE': 8,
      'BEREAVEMENT_LEAVE': 9,
      'MATERNITY_LEAVE': 10,
      'PATERNITY_LEAVE': 11,
      'FAMILY_CARE_LEAVE': 12,
      'MENSTRUAL_LEAVE': 13
    };
    
    const columnIndex = leaveTypeColumns[leaveType];
    
    if (columnIndex) {
      for (let i = 1; i < balanceValues.length; i++) {
        if (balanceValues[i][0] === userId && balanceValues[i][3] === currentYear) {
          const currentBalance = balanceValues[i][columnIndex - 1];
          const newBalance = currentBalance - days;
          balanceSheet.getRange(i + 1, columnIndex).setValue(newBalance);
          balanceSheet.getRange(i + 1, 14).setValue(new Date()); // 更新時間
          Logger.log(`✅ 扣除 ${userId} 的 ${leaveType}：${days} 天，剩餘 ${newBalance} 天`);
          break;
        }
      }
    }
  }
  
  return { 
    ok: true, 
    code: reviewAction === 'approve' ? 'LEAVE_APPROVED' : 'LEAVE_REJECTED'
  };
}


/**
 * 翻譯假別代碼為中文名稱
 */
function getLeaveTypeName(leaveType) {
  const leaveTypeMap = {
    'ANNUAL_LEAVE': '特休假',
    'SICK_LEAVE': '病假',
    'PERSONAL_LEAVE': '事假',
    'MARRIAGE_LEAVE': '婚假',
    'BEREAVEMENT_LEAVE': '喪假',
    'MATERNITY_LEAVE': '產假',
    'PATERNITY_LEAVE': '陪產假',
    'FAMILY_CARE_LEAVE': '家庭照顧假',
    'MENSTRUAL_LEAVE': '生理假'
  };
  
  return leaveTypeMap[leaveType] || leaveType;
}

// ==================== Handler 函式 ====================

function handleGetLeaveBalance(params) {
  return getLeaveBalance(params.token);
}

function handleSubmitLeave(params) {
  const { token, leaveType, startDate, endDate, days, reason } = params;
  return submitLeaveRequest(
    token,
    leaveType,
    startDate,
    endDate,
    parseFloat(days),
    reason
  );
}

function handleGetEmployeeLeaveRecords(params) {
  return getEmployeeLeaveRecords(params.token);
}

function handleGetPendingLeaveRequests(params) {
  return getPendingLeaveRequests(params.token);
}

function handleReviewLeave(params) {
  const { token, rowNumber, reviewAction, comment } = params;
  
  Logger.log(`📥 handleReviewLeave 收到參數:`);
  Logger.log(`   - rowNumber: ${rowNumber}`);
  Logger.log(`   - reviewAction: "${reviewAction}"`);
  Logger.log(`   - comment: "${comment}"`);
  
  return reviewLeaveRequest(
    token,
    parseInt(rowNumber),
    reviewAction,
    comment || ''
  );
}

/**
 * 管理員手動初始化員工假期額度
 * @param {string} sessionToken - Session Token
 * @param {string} targetUserId - 目標員工ID
 * @param {string} hireDate - 到職日期（YYYY-MM-DD）
 */
function handleInitializeEmployeeLeave(params) {
  const { token, targetUserId, hireDate } = params;
  
  const session = checkSession_(token);
  if (!session.ok) return { ok: false, code: session.code };
  
  // 檢查是否為管理員
  if (session.user.dept !== '管理員') {
    return { ok: false, code: 'ERR_NO_PERMISSION' };
  }
  
  try {
    initializeLeaveBalance_(targetUserId, new Date(hireDate));
    return { ok: true, code: 'LEAVE_BALANCE_INITIALIZED' };
  } catch (err) {
    return { ok: false, code: 'ERR_INITIALIZE_FAILED', msg: err.message };
  }
}


/**
 * 🔍 終極診斷：直接抓 userId 並查詢
 */
function ultimateDiagnosis() {
  Logger.log('═══════════════════════════════════════');
  Logger.log('🔍 終極診斷：直接抓 userId');
  Logger.log('═══════════════════════════════════════');
  Logger.log('');
  
  // Step 1: 從 Session 工作表抓 userId
  Logger.log('📡 Step 1: 讀取 Session 工作表');
  const sessionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Session');
  
  if (!sessionSheet) {
    Logger.log('❌ 找不到 Session 工作表');
    return;
  }
  
  const sessionData = sessionSheet.getDataRange().getValues();
  Logger.log('✅ Session 工作表有 ' + (sessionData.length - 1) + ' 筆資料');
  Logger.log('');
  
  // 找最新的 Session（通常是第 2 行）
  if (sessionData.length < 2) {
    Logger.log('❌ Session 工作表沒有資料');
    return;
  }
  
  const latestSession = sessionData[1];  // 第 2 行（索引 1）
  const userId = latestSession[1];  // B 欄（索引 1）
  
  Logger.log('👤 從 Session 抓到的 User ID:');
  Logger.log('   原始值: "' + userId + '"');
  Logger.log('   型別: ' + typeof userId);
  Logger.log('   長度: ' + (userId ? userId.length : 0));
  Logger.log('');
  
  // Step 2: 列出所有工作表名稱
  Logger.log('📋 所有工作表名稱:');
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  allSheets.forEach((s, i) => {
    Logger.log(`   ${i + 1}. "${s.getName()}"`);
  });
  Logger.log('');
  
  // Step 3: 嘗試各種可能的工作表名稱
  const possibleNames = [
    '請假紀錄',
    '申請時間',  // 從你的截圖看到的
    'Leave Records',
    '請假次錄'
  ];
  
  let leaveSheet = null;
  let actualSheetName = '';
  
  Logger.log('🔍 嘗試找請假工作表...');
  for (const name of possibleNames) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (sheet) {
      leaveSheet = sheet;
      actualSheetName = name;
      Logger.log(`✅ 找到工作表: "${name}"`);
      break;
    } else {
      Logger.log(`❌ 找不到: "${name}"`);
    }
  }
  
  Logger.log('');
  
  if (!leaveSheet) {
    Logger.log('❌ 完全找不到請假工作表！');
    Logger.log('');
    Logger.log('💡 請從上面的「所有工作表名稱」找出正確的名稱');
    return;
  }
  
  Logger.log('📄 實際工作表名稱: "' + actualSheetName + '"');
  Logger.log('');
  
  const leaveData = leaveSheet.getDataRange().getValues();
  Logger.log('✅ 請假記錄有 ' + (leaveData.length - 1) + ' 筆資料');
  Logger.log('');
  
  // Step 4: 顯示標題列
  Logger.log('📋 標題列:');
  leaveData[0].forEach((header, index) => {
    const column = String.fromCharCode(65 + index);
    Logger.log('   ' + column + ' (' + index + '): ' + header);
  });
  Logger.log('');
  
  // Step 5: 逐行比對
  Logger.log('🔍 開始比對資料...');
  Logger.log('');
  
  let foundCount = 0;
  
  for (let i = 1; i < leaveData.length; i++) {
    const rowUserId = leaveData[i][1];  // B 欄
    
    Logger.log('第 ' + (i + 1) + ' 行:');
    Logger.log('   原始值: "' + rowUserId + '"');
    Logger.log('   型別: ' + typeof rowUserId);
    Logger.log('   長度: ' + (rowUserId ? rowUserId.length : 0));
    
    // 清理並比對
    const cleanRowId = String(rowUserId).trim();
    const cleanUserId = String(userId).trim();
    
    Logger.log('   清理後: "' + cleanRowId + '"');
    Logger.log('   比對值: "' + cleanUserId + '"');
    Logger.log('   相等: ' + (cleanRowId === cleanUserId));
    
    // 逐字元比對（如果不相等）
    if (cleanRowId !== cleanUserId) {
      Logger.log('   ⚠️ 不匹配！逐字元檢查:');
      const maxLen = Math.max(cleanRowId.length, cleanUserId.length);
      for (let j = 0; j < maxLen; j++) {
        const char1 = cleanRowId[j] || '(無)';
        const char2 = cleanUserId[j] || '(無)';
        const code1 = cleanRowId.charCodeAt(j) || 0;
        const code2 = cleanUserId.charCodeAt(j) || 0;
        
        if (char1 !== char2) {
          Logger.log('      位置 ' + j + ': "' + char1 + '" (code:' + code1 + ') vs "' + char2 + '" (code:' + code2 + ') ❌');
        }
      }
    } else {
      Logger.log('   ✅✅✅ 找到匹配記錄！');
      Logger.log('      假別: ' + leaveData[i][4]);
      Logger.log('      開始日期: ' + leaveData[i][5]);
      Logger.log('      結束日期: ' + leaveData[i][6]);
      Logger.log('      狀態: ' + leaveData[i][9]);
      foundCount++;
    }
    
    Logger.log('');
  }
  
  Logger.log('═══════════════════════════════════════');
  Logger.log('📊 診斷結果');
  Logger.log('═══════════════════════════════════════');
  Logger.log('   User ID: ' + userId);
  Logger.log('   請假記錄總數: ' + (leaveData.length - 1));
  Logger.log('   找到匹配記錄: ' + foundCount);
  Logger.log('');
  
  if (foundCount === 0) {
    Logger.log('❌ 沒有找到任何匹配記錄！');
    Logger.log('');
    Logger.log('🔍 可能原因:');
    Logger.log('   1. User ID 不匹配（有隱藏字元或編碼問題）');
    Logger.log('   2. 工作表中的員工ID格式錯誤');
    Logger.log('   3. 請檢查上面的逐字元比對結果');
  } else {
    Logger.log('✅✅✅ 找到 ' + foundCount + ' 筆匹配記錄！');
    Logger.log('   問題應該在 getEmployeeLeaveRecords 函數');
  }
  
  Logger.log('═══════════════════════════════════════');
}


function testGetEmployeeLeaveRecords() {
  Logger.log('🧪 測試 getEmployeeLeaveRecords');
  Logger.log('');
  
  const token = '0ed9f26d-d3ec-48b1-8eaf-40c29e725438';  // ⚠️ 替換成你的有效 token
  
  Logger.log('📡 開始查詢...');
  Logger.log('');
  
  const result = getEmployeeLeaveRecords(token);
  
  Logger.log('📤 完整回應:');
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log('');
  
  const records = result.records || result.data || [];
  
  Logger.log('🔍 解析結果:');
  Logger.log('   ok: ' + result.ok);
  Logger.log('   success: ' + result.success);
  Logger.log('   records 數量: ' + records.length);
  Logger.log('');
  
  if (records.length > 0) {
    Logger.log('✅✅✅ 找到請假記錄！');
    Logger.log('');
    Logger.log('📋 記錄詳情:');
    
    records.forEach((rec, i) => {
      Logger.log(`   ${i + 1}. ${rec.leaveType} - ${rec.startDate} ~ ${rec.endDate}`);
      Logger.log(`      狀態: ${rec.status}, 天數: ${rec.days}`);
      if (rec.reason) {
        Logger.log(`      原因: ${rec.reason}`);
      }
      Logger.log('');
    });
  } else {
    Logger.log('❌ 還是沒找到記錄');
  }
}
