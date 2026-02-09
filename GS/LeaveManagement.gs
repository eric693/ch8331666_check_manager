// LeaveManagement.gs - 小時制請假系統（無時段限制版）

/**
 * ✅ 提交請假申請（無時段限制版）
 * 
 * 修改內容：
 * 1. 移除工作時段限制
 * 2. 移除整點時間檢查（如果需要更靈活）
 * 3. 使用新的計算邏輯
 */
function submitLeaveRequest(sessionToken, leaveType, startDateTime, endDateTime, reason) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 開始處理請假申請（無時段限制版）');
    Logger.log('═══════════════════════════════════════');
    
    // 步驟 1：驗證 Session
    Logger.log('📡 驗證 Session...');
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      Logger.log('❌ Session 驗證失敗');
      return { 
        ok: false, 
        code: "ERR_SESSION_INVALID",
        msg: "未授權或 session 已過期" 
      };
    }
    
    const user = employee.user;
    Logger.log('✅ Session 驗證成功');
    Logger.log(`   員工ID: ${user.userId}`);
    Logger.log(`   員工姓名: ${user.name}`);
    Logger.log('');
    
    // 步驟 2：記錄收到的參數
    Logger.log('📥 收到的參數:');
    Logger.log(`   leaveType: ${leaveType}`);
    Logger.log(`   startDateTime: ${startDateTime}`);
    Logger.log(`   endDateTime: ${endDateTime}`);
    Logger.log(`   reason: ${reason}`);
    Logger.log('');
    
    // 步驟 3：解析日期時間
    Logger.log('🔄 解析日期時間...');
    
    let start, end;
    
    try {
      start = new Date(startDateTime);
      end = new Date(endDateTime);
      
      if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        Logger.log('❌ 日期時間格式無效');
        Logger.log(`   startDateTime: ${startDateTime} → ${start}`);
        Logger.log(`   endDateTime: ${endDateTime} → ${end}`);
        return {
          ok: false,
          code: "ERR_INVALID_DATETIME",
          msg: "日期時間格式無效"
        };
      }
      
      Logger.log('✅ 日期時間解析成功');
      Logger.log(`   開始: ${start.toISOString()}`);
      Logger.log(`   結束: ${end.toISOString()}`);
      
    } catch (parseError) {
      Logger.log('❌ 日期時間解析失敗: ' + parseError.message);
      return {
        ok: false,
        code: "ERR_DATETIME_PARSE",
        msg: "無法解析日期時間"
      };
    }
    
    Logger.log('');

    // 步驟 4：驗證時間順序
    if (end <= start) {
      Logger.log('❌ 結束時間必須晚於開始時間');
      return {
        ok: false,
        code: "ERR_INVALID_TIME_RANGE",
        msg: "結束時間必須晚於開始時間"
      };
    }

    // ⭐ 步驟 4.5：檢查是否為整點時間（可選，如果想更靈活可以移除）
    Logger.log('🔍 檢查是否為整點時間...');

    const startMinutes = start.getMinutes();
    const startSeconds = start.getSeconds();
    const endMinutes = end.getMinutes();
    const endSeconds = end.getSeconds();

    if (startMinutes !== 0 || startSeconds !== 0) {
      Logger.log(`❌ 開始時間不是整點: ${start.toISOString()}`);
      Logger.log(`   分鐘: ${startMinutes}, 秒: ${startSeconds}`);
      return {
        ok: false,
        code: "ERR_INVALID_TIME_FORMAT",
        msg: "開始時間必須是整點（例如：09:00, 10:00）"
      };
    }

    if (endMinutes !== 0 || endSeconds !== 0) {
      Logger.log(`❌ 結束時間不是整點: ${end.toISOString()}`);
      Logger.log(`   分鐘: ${endMinutes}, 秒: ${endSeconds}`);
      return {
        ok: false,
        code: "ERR_INVALID_TIME_FORMAT",
        msg: "結束時間必須是整點（例如：09:00, 10:00）"
      };
    }

    Logger.log('✅ 時間格式檢查通過（整點）');
    Logger.log('');

    // ⭐⭐⭐ 步驟 5：計算工作時數和天數（使用新的無限制計算邏輯）
    Logger.log('💡 計算工作時數和天數（無時段限制）...');
    
    const { workHours, days } = calculateWorkHoursAndDays_Unlimited(start, end);
    
    Logger.log(`   工作時數: ${workHours} 小時`);
    Logger.log(`   天數: ${days} 天`);
    Logger.log('');
    
    // 步驟 6：檢查假期餘額
    Logger.log('🔍 檢查假期餘額...');
    const balance = getLeaveBalance(sessionToken);
    
    if (!balance.ok) {
      Logger.log('❌ 無法取得假期餘額');
      return {
        ok: false,
        code: "ERR_BALANCE_CHECK",
        msg: "無法取得假期餘額"
      };
    }
    
    Logger.log('✅ 假期餘額檢查完成');
    Logger.log('');
    
    // 步驟 7：格式化日期時間
    Logger.log('📝 格式化日期時間...');
    
    const formattedStartDateTime = Utilities.formatDate(
      start,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    
    const formattedEndDateTime = Utilities.formatDate(
      end,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    
    Logger.log(`   開始時間（格式化）: ${formattedStartDateTime}`);
    Logger.log(`   結束時間（格式化）: ${formattedEndDateTime}`);
    Logger.log('');
    
    // 步驟 8：取得工作表
    Logger.log('📊 取得工作表...');
    const sheet = getLeaveRecordsSheet();
    
    if (!sheet) {
      Logger.log('❌ 無法取得請假記錄工作表');
      return {
        ok: false,
        code: "ERR_SHEET_ACCESS",
        msg: "無法存取請假記錄工作表"
      };
    }
    
    Logger.log('✅ 工作表已就緒');
    Logger.log('');
    
    // 步驟 9：寫入資料
    Logger.log('💾 準備寫入資料...');
    
    const row = [
      new Date(),                  // A: 申請時間
      user.userId || '',           // B: 員工ID
      user.name || '',             // C: 姓名
      user.dept || '',             // D: 部門
      leaveType || '',             // E: 假別
      formattedStartDateTime,      // F: 開始時間
      formattedEndDateTime,        // G: 結束時間
      workHours,                   // H: 工作時數
      days,                        // I: 天數
      reason || '',                // J: 原因
      'PENDING',                   // K: 狀態
      '',                          // L: 審核人
      '',                          // M: 審核時間
      ''                           // N: 審核意見
    ];
    
    Logger.log('📋 準備寫入的資料:');
    Logger.log(`   A (申請時間): ${row[0]}`);
    Logger.log(`   B (員工ID): ${row[1]}`);
    Logger.log(`   C (姓名): ${row[2]}`);
    Logger.log(`   D (部門): ${row[3]}`);
    Logger.log(`   E (假別): ${row[4]}`);
    Logger.log(`   F (開始時間): ${row[5]}`);
    Logger.log(`   G (結束時間): ${row[6]}`);
    Logger.log(`   H (工作時數): ${row[7]}`);
    Logger.log(`   I (天數): ${row[8]}`);
    Logger.log(`   J (原因): ${row[9]}`);
    Logger.log(`   K (狀態): ${row[10]}`);
    Logger.log('');
    
    try {
      sheet.appendRow(row);
      Logger.log('✅ 資料寫入成功');
    } catch (writeError) {
      Logger.log('❌ 資料寫入失敗: ' + writeError.message);
      return {
        ok: false,
        code: "ERR_WRITE_FAILED",
        msg: "無法寫入請假記錄"
      };
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 請假申請提交成功');
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: true,
      code: "LEAVE_SUBMIT_SUCCESS",
      msg: "請假申請已提交",
      data: {
        leaveType: leaveType,
        startDateTime: formattedStartDateTime,
        endDateTime: formattedEndDateTime,
        workHours: workHours,
        days: days,
        reason: reason
      }
    };
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌❌❌ submitLeaveRequest 發生錯誤');
    Logger.log('錯誤訊息: ' + error.message);
    Logger.log('錯誤堆疊: ' + error.stack);
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: false,
      code: "ERR_INTERNAL_ERROR",
      msg: "系統錯誤：" + error.message
    };
  }
}

/**
 * ✅ 無時段限制版：計算工作時數和天數
 * 
 * 特點：
 * 1. 不限制工作時段（24小時制）
 * 2. 可選是否扣除午休時間
 * 3. 自動處理跨日請假
 * 
 * @param {Date} start - 開始時間
 * @param {Date} end - 結束時間
 * @return {Object} { workHours: 小時數, days: 天數 }
 */
function calculateWorkHoursAndDays_Unlimited(start, end) {
  try {
    Logger.log('💡 計算工作時數（無時段限制版）');
    Logger.log(`   開始: ${start.toISOString()}`);
    Logger.log(`   結束: ${end.toISOString()}`);
    
    // ⭐ 配置選項
    const DEDUCT_LUNCH = true;   // 是否扣除午休時間（true = 扣除，false = 不扣除）
    const LUNCH_START = 12;      // 午休開始（小時）
    const LUNCH_END = 13;        // 午休結束（小時）
    
    // 計算總毫秒數
    const totalMs = end - start;
    
    // 轉換為小時
    let totalHours = totalMs / (1000 * 60 * 60);
    
    Logger.log(`   ⏱️ 原始時數: ${totalHours.toFixed(2)} 小時`);
    
    // ⭐ 扣除午休時間（如果啟用）
    if (DEDUCT_LUNCH) {
      Logger.log('   🍱 開始計算午休扣除...');
      
      // 計算跨越的天數
      const startDate = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endDate = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      
      let lunchHoursToDeduct = 0;
      
      // 遍歷每一天，檢查是否跨越午休時間
      let currentDate = new Date(startDate);
      
      while (currentDate <= endDate) {
        // 當天的午休開始和結束時間
        const lunchStartTime = new Date(currentDate);
        lunchStartTime.setHours(LUNCH_START, 0, 0, 0);
        
        const lunchEndTime = new Date(currentDate);
        lunchEndTime.setHours(LUNCH_END, 0, 0, 0);
        
        // 計算請假時間與午休時間的交集
        const overlapStart = start > lunchStartTime ? start : lunchStartTime;
        const overlapEnd = end < lunchEndTime ? end : lunchEndTime;
        
        // 如果有交集，計算重疊的時間
        if (overlapStart < overlapEnd) {
          const overlapMs = overlapEnd - overlapStart;
          const overlapHours = overlapMs / (1000 * 60 * 60);
          lunchHoursToDeduct += overlapHours;
          
          Logger.log(`      ${Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')} 扣除: ${overlapHours.toFixed(2)} 小時`);
        }
        
        // 移到下一天
        currentDate.setDate(currentDate.getDate() + 1);
      }
      
      totalHours -= lunchHoursToDeduct;
      
      if (lunchHoursToDeduct > 0) {
        Logger.log(`   🍱 總共扣除午休: ${lunchHoursToDeduct.toFixed(2)} 小時`);
      } else {
        Logger.log(`   🍱 無需扣除午休`);
      }
    } else {
      Logger.log('   ℹ️ 不扣除午休時間');
    }
    
    // 確保不會是負數
    totalHours = Math.max(0, totalHours);
    
    // 四捨五入到小數點後 2 位
    const finalHours = Math.round(totalHours * 100) / 100;
    const days = Math.round((finalHours / 8) * 100) / 100;
    
    Logger.log(`   ✅ 最終工時: ${finalHours} 小時 = ${days} 天`);
    
    return {
      workHours: finalHours,
      days: days
    };
    
  } catch (error) {
    Logger.log(`❌ calculateWorkHoursAndDays_Unlimited 錯誤: ${error.message}`);
    return { workHours: 0, days: 0 };
  }
}

/**
 * ✅ 取得或建立請假記錄工作表
 */
function getLeaveRecordsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('請假紀錄');
  
  if (!sheet) {
    Logger.log('📝 請假紀錄工作表不存在，自動建立...');
    
    sheet = ss.insertSheet('請假紀錄');
    
    sheet.appendRow([
      '申請時間', '員工ID', '姓名', '部門', '假別',
      '開始時間', '結束時間', '工作時數', '天數', '原因',
      '狀態', '審核人', '審核時間', '審核意見'
    ]);
    
    const headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 請假紀錄工作表已建立');
  }
  
  return sheet;
}

/**
 * ✅ 取得假期餘額
 */
function getLeaveBalance(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    Logger.log('🔍 查詢員工: ' + user.userId);
    
    const sheet = getLeaveBalanceSheet();
    
    if (!sheet) {
      Logger.log('❌ 工作表不存在，嘗試建立...');
      initializeEmployeeLeave(sessionToken);
      return getLeaveBalance(sessionToken);
    }
    
    const values = sheet.getDataRange().getValues();
    Logger.log('📊 工作表行數: ' + values.length);
    
    for (let i = 1; i < values.length; i++) {
      Logger.log(`   檢查第 ${i} 行: ${values[i][0]}`);
      
      if (values[i][0] === user.userId) {
        Logger.log('✅ 找到員工資料');
        
        const balance = {
          employeeName: values[i][1] || user.name,
          ANNUAL_LEAVE: values[i][2] || 0,
          SICK_LEAVE: values[i][3] || 0,
          PERSONAL_LEAVE: values[i][4] || 0,
          BEREAVEMENT_LEAVE: values[i][5] || 0,
          MARRIAGE_LEAVE: values[i][6] || 0,
          MATERNITY_LEAVE: values[i][7] || 0,
          PATERNITY_LEAVE: values[i][8] || 0,
          HOSPITALIZATION_LEAVE: values[i][9] || 0,
          MENSTRUAL_LEAVE: values[i][10] || 0,
          FAMILY_CARE_LEAVE: values[i][11] || 0,
          OFFICIAL_LEAVE: values[i][12] || 0,
          WORK_INJURY_LEAVE: values[i][13] || 0,
          NATURAL_DISASTER_LEAVE: values[i][14] || 0,
          COMP_TIME_OFF: values[i][15] || 0,
          ABSENCE_WITHOUT_LEAVE: values[i][16] || 0
        };
        
        Logger.log('📋 假期餘額（小時）:');
        Logger.log(JSON.stringify(balance, null, 2));
        
        return {
          ok: true,
          balance: balance
        };
      }
    }
    
    Logger.log('⚠️ 找不到員工資料，嘗試初始化...');
    initializeEmployeeLeave(sessionToken);
    return getLeaveBalance(sessionToken);
    
  } catch (error) {
    Logger.log('❌ getLeaveBalance 錯誤: ' + error);
    Logger.log('錯誤堆疊: ' + error.stack);
    return {
      ok: false,
      code: "ERR_INTERNAL_ERROR",
      msg: error.message
    };
  }
}

/**
 * ✅ 取得或建立假期餘額工作表
 */
function getLeaveBalanceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('假期餘額');
  
  if (!sheet) {
    Logger.log('📝 假期餘額工作表不存在，自動建立...');
    
    sheet = ss.insertSheet('假期餘額');
    
    sheet.appendRow([
      '員工ID',
      '姓名',
      '特休假',
      '未住院病假',
      '事假',
      '喪假',
      '婚假',
      '產假',
      '陪產檢及陪產假',
      '住院病假',
      '生理假',
      '家庭照顧假',
      '公假(含兵役假)',
      '公傷假',
      '天然災害停班',
      '加班補休假',
      '曠工',
      '更新時間'
    ]);
    
    const headerRange = sheet.getRange(1, 1, 1, 18);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    sheet.setFrozenRows(1);
    
    Logger.log('✅ 假期餘額工作表已建立');
  }
  
  return sheet;
}

/**
 * ✅ 初始化假期餘額（小時制）
 */
function initializeEmployeeLeave(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    const sheet = getLeaveBalanceSheet();
    
    // 檢查是否已存在
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === user.userId) {
        Logger.log('ℹ️ 員工 ' + user.name + ' 的假期餘額已存在');
        return {
          ok: true,
          msg: "假期餘額已存在"
        };
      }
    }
    
    // 小時制預設值
    const defaultBalance = [
      user.userId,
      user.name,
      56,   // 特休假（7天 × 8）
      240,  // 未住院病假（30天 × 8）
      112,  // 事假（14天 × 8）
      40,   // 喪假（5天 × 8）
      64,   // 婚假（8天 × 8）
      448,  // 產假（56天 × 8）
      56,   // 陪產檢及陪產假（7天 × 8）
      240,  // 住院病假（30天 × 8）
      96,   // 生理假（12天 × 8）
      56,   // 家庭照顧假（7天 × 8）
      0,    // 公假
      0,    // 公傷假
      0,    // 天然災害停班
      0,    // 加班補休假
      0,    // 曠工
      new Date()
    ];
    
    sheet.appendRow(defaultBalance);
    
    Logger.log('✅ 已為員工 ' + user.name + ' 初始化假期餘額（小時制）');
    
    return {
      ok: true,
      msg: "假期餘額已初始化（小時制）"
    };
    
  } catch (error) {
    Logger.log('❌ initializeEmployeeLeave 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 取得員工請假記錄
 */
function getEmployeeLeaveRecords(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    const user = employee.user;
    const sheet = getLeaveRecordsSheet();
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      return {
        ok: true,
        records: []
      };
    }
    
    const records = [];
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === user.userId) {
        const record = {
          applyTime: formatDateTime(values[i][0]),
          employeeName: values[i][2],
          dept: values[i][3],
          leaveType: values[i][4],
          startDateTime: values[i][5],
          endDateTime: values[i][6],
          workHours: values[i][7],
          days: values[i][8],
          reason: values[i][9] || '',
          status: values[i][10] || 'PENDING',
          reviewer: values[i][11] || '',
          reviewTime: values[i][12] ? formatDateTime(values[i][12]) : '',
          reviewComment: values[i][13] || ''
        };
        
        records.push(record);
      }
    }
    
    records.sort((a, b) => new Date(b.applyTime) - new Date(a.applyTime));
    
    return {
      ok: true,
      records: records
    };
    
  } catch (error) {
    Logger.log('❌ getEmployeeLeaveRecords 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 取得待審核請假申請（使用無限制計算邏輯）
 */
function getPendingLeaveRequests(sessionToken) {
  try {
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    if (employee.user.dept !== '管理員') {
      return {
        ok: false,
        code: "ERR_PERMISSION_DENIED",
        msg: "需要管理員權限"
      };
    }
    
    const sheet = getLeaveRecordsSheet();
    const values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) {
      return {
        ok: true,
        requests: []
      };
    }
    
    const requests = [];
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][10] === 'PENDING') {
        
        const startDateTime = values[i][5];
        const endDateTime = values[i][6];
        
        let correctWorkHours = 0;
        let correctDays = 0;
        
        try {
          const start = new Date(startDateTime);
          const end = new Date(endDateTime);
          
          if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
            const result = calculateWorkHoursAndDays_Unlimited(start, end);
            correctWorkHours = result.workHours;
            correctDays = result.days;
          }
        } catch (err) {
          Logger.log('⚠️ 計算工時失敗:', err);
          correctWorkHours = values[i][7] || 0;
          correctDays = values[i][8] || 0;
        }
        
        const request = {
          rowNumber: i + 1,
          applyTime: formatDateTime(values[i][0]),
          employeeId: values[i][1],
          employeeName: values[i][2],
          dept: values[i][3],
          leaveType: values[i][4],
          startDateTime: startDateTime,
          endDateTime: endDateTime,
          workHours: correctWorkHours,
          days: correctDays,
          reason: values[i][9] || ''
        };
        
        requests.push(request);
      }
    }
    
    return {
      ok: true,
      requests: requests
    };
    
  } catch (error) {
    Logger.log('❌ getPendingLeaveRequests 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 審核請假申請
 */
function reviewLeaveRequest(sessionToken, rowNumber, reviewAction, comment) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 開始審核請假');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   行號: ${rowNumber}`);
    Logger.log(`   動作: ${reviewAction}`);
    Logger.log('');
    
    const employee = checkSession_(sessionToken);
    
    if (!employee.ok || !employee.user) {
      return {
        ok: false,
        code: "ERR_SESSION_INVALID"
      };
    }
    
    if (employee.user.dept !== '管理員') {
      return {
        ok: false,
        code: "ERR_PERMISSION_DENIED",
        msg: "需要管理員權限"
      };
    }
    
    const sheet = getLeaveRecordsSheet();
    const record = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];
    
    const userId = record[1];
    const employeeName = record[2];
    const leaveType = record[4];
    const workHours = record[7];
    const days = record[8];
    
    Logger.log('📋 請假資料:');
    Logger.log(`   員工: ${employeeName} (${userId})`);
    Logger.log(`   假別: ${leaveType}`);
    Logger.log(`   時數: ${workHours} 小時`);
    Logger.log(`   天數: ${days} 天`);
    Logger.log('');
    
    const status = (reviewAction === 'approve') ? 'APPROVED' : 'REJECTED';
    
    sheet.getRange(rowNumber, 11).setValue(status);
    sheet.getRange(rowNumber, 12).setValue(employee.user.name);
    sheet.getRange(rowNumber, 13).setValue(new Date());
    sheet.getRange(rowNumber, 14).setValue(comment || '');
    
    Logger.log(`✅ 審核狀態已更新: ${status}`);
    Logger.log('');
    
    if (reviewAction === 'approve') {
      Logger.log('💰 開始扣除假期餘額...');
      
      const deductResult = deductLeaveBalance(userId, leaveType, workHours);
      
      if (!deductResult.ok) {
        Logger.log('❌ 扣除餘額失敗: ' + deductResult.msg);
        
        sheet.getRange(rowNumber, 11).setValue('PENDING');
        
        return {
          ok: false,
          code: "ERR_DEDUCT_FAILED",
          msg: "扣除餘額失敗: " + deductResult.msg
        };
      }
      
      Logger.log('✅ 假期餘額扣除成功');
      Logger.log(`   ${leaveType}: 扣除 ${workHours} 小時`);
      Logger.log(`   剩餘: ${deductResult.remaining} 小時`);
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅✅✅ 審核完成');
    Logger.log('═══════════════════════════════════════');
    
    return {
      ok: true,
      msg: "審核完成"
    };
    
  } catch (error) {
    Logger.log('❌ reviewLeaveRequest 錯誤: ' + error);
    Logger.log('錯誤堆疊: ' + error.stack);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 扣除假期餘額
 */
function deductLeaveBalance(userId, leaveType, hours) {
  try {
    Logger.log('📊 扣除假期餘額');
    Logger.log(`   員工ID: ${userId}`);
    Logger.log(`   假別: ${leaveType}`);
    Logger.log(`   小時數: ${hours}`);
    Logger.log('');
    
    const sheet = getLeaveBalanceSheet();
    const values = sheet.getDataRange().getValues();
    
    const leaveTypeColumnMap = {
      'ANNUAL_LEAVE': 3,
      'SICK_LEAVE': 4,
      'PERSONAL_LEAVE': 5,
      'BEREAVEMENT_LEAVE': 6,
      'MARRIAGE_LEAVE': 7,
      'MATERNITY_LEAVE': 8,
      'PATERNITY_LEAVE': 9,
      'HOSPITALIZATION_LEAVE': 10,
      'MENSTRUAL_LEAVE': 11,
      'FAMILY_CARE_LEAVE': 12,
      'OFFICIAL_LEAVE': 13,
      'WORK_INJURY_LEAVE': 14,
      'NATURAL_DISASTER_LEAVE': 15,
      'COMP_TIME_OFF': 16,
      'ABSENCE_WITHOUT_LEAVE': 17
    };
    
    const columnIndex = leaveTypeColumnMap[leaveType];
    
    if (!columnIndex) {
      Logger.log('❌ 無效的假別: ' + leaveType);
      return {
        ok: false,
        msg: "無效的假別"
      };
    }
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === userId) {
        Logger.log(`✅ 找到員工記錄（第 ${i + 1} 行）`);
        Logger.log(`   姓名: ${values[i][1]}`);
        
        const currentBalance = values[i][columnIndex - 1];
        
        Logger.log(`   目前餘額: ${currentBalance} 小時`);
        
        if (currentBalance < hours) {
          Logger.log(`   ⚠️ 餘額不足：需要 ${hours} 小時，只剩 ${currentBalance} 小時`);
          return {
            ok: false,
            msg: `${leaveType} 餘額不足（需要 ${hours} 小時，只剩 ${currentBalance} 小時）`
          };
        }
        
        const newBalance = currentBalance - hours;
        
        Logger.log(`   扣除 ${hours} 小時後: ${newBalance} 小時`);
        
        sheet.getRange(i + 1, columnIndex).setValue(newBalance);
        sheet.getRange(i + 1, 18).setValue(new Date());
        
        Logger.log('✅ 餘額已更新');
        
        return {
          ok: true,
          remaining: newBalance
        };
      }
    }
    
    Logger.log('❌ 找不到員工記錄');
    return {
      ok: false,
      msg: "找不到員工記錄"
    };
    
  } catch (error) {
    Logger.log('❌ deductLeaveBalance 錯誤: ' + error);
    return {
      ok: false,
      msg: error.message
    };
  }
}

/**
 * ✅ 取得已核准的請假記錄（使用無限制計算）
 */
function getApprovedLeaveRecords(monthParam, userIdParam) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('📋 getApprovedLeaveRecords 開始（無時段限制版）');
    Logger.log('═══════════════════════════════════════');
    Logger.log(`   monthParam: ${monthParam}`);
    Logger.log(`   userIdParam: ${userIdParam}`);
    Logger.log('');
    
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
    
    const leaveRecords = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      const applyTime = row[0];
      const employeeId = row[1];
      const employeeName = row[2];
      const dept = row[3];
      const leaveType = row[4];
      const startDateTime = row[5];
      const endDateTime = row[6];
      const workHours = row[7];
      const days = row[8];
      const reason = row[9];
      const status = row[10];
      const reviewer = row[11];
      const reviewTime = row[12];
      const reviewComment = row[13];
      
      Logger.log(`═══════════════════════════════════════`);
      Logger.log(`📋 第 ${i + 1} 行:`);
      Logger.log(`   員工ID: ${employeeId}`);
      Logger.log(`   員工姓名: ${employeeName}`);
      Logger.log(`   狀態: "${status}"`);
      Logger.log(`   開始時間: ${startDateTime}`);
      Logger.log(`   結束時間: ${endDateTime}`);
      Logger.log(`   工作時數: ${workHours} 小時`);
      Logger.log(`   天數: ${days} 天`);
      
      if (String(status).trim() !== 'APPROVED') {
        Logger.log(`   ⏭️ 狀態不是 APPROVED，跳過`);
        Logger.log('');
        continue;
      }
      
      let formattedStartDate, formattedEndDate;
      
      try {
        const startDate = new Date(startDateTime);
        const endDate = new Date(endDateTime);
        
        if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
          Logger.log(`   ⚠️ 日期格式無效，跳過`);
          Logger.log('');
          continue;
        }
        
        formattedStartDate = formatDate(startDate);
        formattedEndDate = formatDate(endDate);
        
        Logger.log(`   格式化開始日期: ${formattedStartDate}`);
        Logger.log(`   格式化結束日期: ${formattedEndDate}`);
        
      } catch (dateError) {
        Logger.log(`   ⚠️ 日期解析錯誤: ${dateError.message}`);
        Logger.log('');
        continue;
      }
      
      if (!formattedStartDate.startsWith(monthParam)) {
        Logger.log(`   ⏭️ 月份不符，跳過`);
        Logger.log('');
        continue;
      }
      
      if (userIdParam && employeeId !== userIdParam) {
        Logger.log(`   ⏭️ 員工ID不符，跳過`);
        Logger.log('');
        continue;
      }
      
      Logger.log(`   ✅ 符合所有條件！`);
      
      const start = new Date(startDateTime);
      const end = new Date(endDateTime);
      
      const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endDay = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      const totalDays = Math.floor((endDay - startDay) / (1000 * 60 * 60 * 24)) + 1;
      
      Logger.log(`   📅 請假天數範圍: ${totalDays} 天`);
      
      for (let d = new Date(startDay); d <= endDay; d.setDate(d.getDate() + 1)) {
        const dateStr = formatDate(d);
        
        if (dateStr.startsWith(monthParam)) {
          leaveRecords.push({
            employeeId: employeeId,
            employeeName: employeeName,
            date: dateStr,
            leaveType: leaveType,
            workHours: parseFloat(workHours) || 0,
            days: parseFloat(days) || 0,
            status: status,
            reason: reason || '',
            startDateTime: startDateTime,
            endDateTime: endDateTime,
            reviewer: reviewer || '',
            reviewTime: reviewTime || '',
            reviewComment: reviewComment || ''
          });
          
          Logger.log(`      ➕ 加入日期: ${dateStr}`);
        }
      }
      
      Logger.log('');
    }
    
    Logger.log('═══════════════════════════════════════');
    Logger.log(`✅ getApprovedLeaveRecords 完成`);
    Logger.log(`   共找到 ${leaveRecords.length} 筆已核准的請假記錄`);
    Logger.log('═══════════════════════════════════════');
    
    return leaveRecords;
    
  } catch (error) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('❌ getApprovedLeaveRecords 錯誤');
    Logger.log('   錯誤訊息: ' + error.message);
    Logger.log('   錯誤堆疊: ' + error.stack);
    Logger.log('═══════════════════════════════════════');
    return [];
  }
}

/**
 * ✅ 格式化日期時間
 */
function formatDateTime(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    return String(date);
  }
}

/**
 * ✅ 格式化日期
 */
function formatDate(date) {
  if (!date) return '';
  try {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return String(date);
  }
}

/**
 * 🧪 測試函數：測試無時段限制的請假
 */
function testUnlimitedLeave() {
  Logger.log('🧪 測試無時段限制請假');
  Logger.log('');
  
  const testParams = {
    token: '16568f73-dd16-4dde-958d-1ab2e703cab5',  // ⚠️ 替換成有效 token
    leaveType: 'ANNUAL_LEAVE',
    startDateTime: '2026-02-06T18:00',  // 晚上 18:00
    endDateTime: '2026-02-06T22:00',    // 晚上 22:00
    reason: '測試晚上時段請假'
  };
  
  Logger.log('📥 測試參數:');
  Logger.log(JSON.stringify(testParams, null, 2));
  Logger.log('');
  
  const result = submitLeaveRequest(
    testParams.token,
    testParams.leaveType,
    testParams.startDateTime,
    testParams.endDateTime,
    testParams.reason
  );
  
  Logger.log('');
  Logger.log('📤 測試結果:');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.ok) {
    Logger.log('');
    Logger.log('✅✅✅ 測試成功！');
    Logger.log('應該顯示：4 小時');
    Logger.log('請檢查 Google Sheet 的「請假紀錄」工作表');
  } else {
    Logger.log('');
    Logger.log('❌ 測試失敗');
  }
}