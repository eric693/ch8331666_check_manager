// SalaryManagement-Enhanced.gs - 薪資管理系統（完整版 - 修正版）

// ==================== 常數定義 ====================

const SHEET_SALARY_CONFIG_ENHANCED = "員工薪資設定";
const SHEET_MONTHLY_SALARY_ENHANCED = "月薪資記錄";

// 台灣法定最低薪資（2025）
// const MIN_MONTHLY_SALARY = 28590;  // 月薪
// const MIN_HOURLY_SALARY = 190;     // 時薪

// 加班費率
const OVERTIME_RATES = {
  weekday: 1.34,      // 平日加班（前2小時）
  weekdayExtra: 1.67, // 平日加班（第3小時起）
  restday: 1.34,      // 休息日前2小時
  restdayExtra: 1.67, // 休息日第3小時起
  holiday: 2.0        // 國定假日
};

// ==================== 試算表管理 ====================

/**
 * ✅ 取得或建立員工薪資設定試算表（完整版）
 */
function getEmployeeSalarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_SALARY_CONFIG_ENHANCED);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SALARY_CONFIG_ENHANCED);
    
    const headers = [
      // 基本資訊 (6欄: A-F)
      "員工ID", "員工姓名", "身分證字號", "員工類型", "薪資類型", "基本薪資",
      
      // 固定津貼項目 (6欄: G-L)
      "職務加給", "伙食費", "交通補助", "全勤獎金", "績效獎金", "其他津貼",
      
      // 銀行資訊 (4欄: M-P)
      "銀行代碼", "銀行帳號", "到職日期", "發薪日",
      
      // 法定扣款 (6欄: Q-V)
      "勞退自提率(%)", "勞保費", "健保費", "就業保險費", "勞退自提", "所得稅",
      
      // 其他扣款 (4欄: W-Z)
      "福利金扣款", "宿舍費用", "團保費用", "其他扣款",
      
      // 系統欄位 (3欄: AA-AC)
      "狀態", "備註", "最後更新時間"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#10b981");
    sheet.getRange(1, 1, 1, headers.length).setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    
    Logger.log("✅ 建立員工薪資設定試算表（完整版）");
  }
  
  return sheet;
}

/**
 * ✅ 取得或建立月薪資記錄試算表（完整版）
 */
function getMonthlySalarySheetEnhanced() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MONTHLY_SALARY_ENHANCED);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MONTHLY_SALARY_ENHANCED);
    
    const headers = [
      // 基本資訊
      "薪資單ID", "員工ID", "員工姓名", "年月",
      
      // 應發項目
      "基本薪資", "職務加給", "伙食費", "交通補助", "全勤獎金", "績效獎金", "其他津貼",
      "平日加班費", "休息日加班費", "國定假日加班費",
      
      // 法定扣款
      "勞保費", "健保費", "就業保險費", "勞退自提", "所得稅",
      
      // 其他扣款
      "請假扣款", "福利金扣款", "宿舍費用", "團保費用", "其他扣款",
      
      // 總計
      "應發總額", "實發金額",
      
      // 銀行資訊
      "銀行代碼", "銀行帳號",
      
      // 系統欄位
      "狀態", "備註", "建立時間"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange(1, 1, 1, headers.length).setBackground("#10b981");
    sheet.getRange(1, 1, 1, headers.length).setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    
    Logger.log("✅ 建立月薪資記錄試算表（完整版）");
  }
  
  return sheet;
}

// ==================== 薪資設定功能 ====================

/**
 * ✅ 設定員工薪資資料（完整版 - 修正版）
 */
function setEmployeeSalaryTW(salaryData) {
  try {
    Logger.log('💰 開始設定員工薪資（完整版 - 修正版）');
    Logger.log('📥 收到的資料: ' + JSON.stringify(salaryData, null, 2));
    
    const sheet = getEmployeeSalarySheet();
    const data = sheet.getDataRange().getValues();
    
    // 驗證必填欄位
    if (!salaryData.employeeId || !salaryData.employeeName || !salaryData.baseSalary || salaryData.baseSalary <= 0) {
      return { success: false, message: "缺少必填欄位或基本薪資無效" };
    }
    
    // 檢查是否已存在
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(salaryData.employeeId).trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const now = new Date();
    
    // ⭐ 修正：確保順序與 Sheet 欄位完全一致
    const row = [
      // A-F: 基本資訊 (6欄)
      String(salaryData.employeeId).trim(),              // A: 員工ID
      String(salaryData.employeeName).trim(),            // B: 員工姓名
      String(salaryData.idNumber || "").trim(),          // C: 身分證字號
      String(salaryData.employeeType || "正職").trim(),  // D: 員工類型
      String(salaryData.salaryType || "月薪").trim(),    // E: 薪資類型
      parseFloat(salaryData.baseSalary) || 0,            // F: 基本薪資
      
      // G-L: 固定津貼項目 (6欄)
      parseFloat(salaryData.positionAllowance) || 0,     // G: 職務加給
      parseFloat(salaryData.mealAllowance) || 0,         // H: 伙食費
      parseFloat(salaryData.transportAllowance) || 0,    // I: 交通補助
      parseFloat(salaryData.attendanceBonus) || 0,       // J: 全勤獎金
      parseFloat(salaryData.performanceBonus) || 0,      // K: 績效獎金
      parseFloat(salaryData.otherAllowances) || 0,       // L: 其他津貼
      
      // M-P: 銀行資訊 (4欄)
      String(salaryData.bankCode || "").trim(),          // M: 銀行代碼
      String(salaryData.bankAccount || "").trim(),       // N: 銀行帳號
      salaryData.hireDate || "",                         // O: 到職日期
      String(salaryData.paymentDay || "5").trim(),       // P: 發薪日
      
      // Q-V: 法定扣款 (6欄)
      parseFloat(salaryData.pensionSelfRate) || 0,       // Q: 勞退自提率(%)
      parseFloat(salaryData.laborFee) || 0,              // R: 勞保費
      parseFloat(salaryData.healthFee) || 0,             // S: 健保費
      parseFloat(salaryData.employmentFee) || 0,         // T: 就業保險費
      parseFloat(salaryData.pensionSelf) || 0,           // U: 勞退自提
      parseFloat(salaryData.incomeTax) || 0,             // V: 所得稅
      
      // W-Z: 其他扣款 (4欄)
      parseFloat(salaryData.welfareFee) || 0,            // W: 福利金扣款
      parseFloat(salaryData.dormitoryFee) || 0,          // X: 宿舍費用
      parseFloat(salaryData.groupInsurance) || 0,        // Y: 團保費用
      parseFloat(salaryData.otherDeductions) || 0,       // Z: 其他扣款
      
      // AA-AC: 系統欄位 (3欄)
      "在職",                                             // AA: 狀態
      String(salaryData.note || "").trim(),              // AB: 備註
      now                                                 // AC: 最後更新時間
    ];
    
    Logger.log(`📝 準備寫入的 row 陣列長度: ${row.length}`);
    Logger.log(`📋 Sheet 標題欄位數: ${data[0].length}`);
    
    if (row.length !== data[0].length) {
      Logger.log(`⚠️ 警告：row 長度 (${row.length}) 與 Sheet 欄位數 (${data[0].length}) 不一致`);
    }
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      Logger.log(`✅ 更新員工薪資設定: ${salaryData.employeeName} (列 ${rowIndex})`);
    } else {
      sheet.appendRow(row);
      Logger.log(`✅ 新增員工薪資設定: ${salaryData.employeeName}`);
    }
    
    // 同步到月薪資記錄
    const currentYearMonth = Utilities.formatDate(now, "Asia/Taipei", "yyyy-MM");
    syncSalaryToMonthlyRecord(salaryData.employeeId, currentYearMonth);
    
    return { success: true, message: "薪資設定成功" };
    
  } catch (error) {
    Logger.log("❌ 設定薪資失敗: " + error);
    Logger.log("❌ 錯誤堆疊: " + error.stack);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 取得員工薪資設定（完整版）
 */
function getEmployeeSalaryTW(employeeId) {
  try {
    const sheet = getEmployeeSalarySheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(employeeId).trim()) {
        const salaryConfig = {};
        headers.forEach((header, index) => {
          salaryConfig[header] = data[i][index];
        });
        
        return { success: true, data: salaryConfig };
      }
    }
    
    return { success: false, message: "找不到該員工薪資資料" };
    
  } catch (error) {
    Logger.log("❌ 取得薪資設定失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 同步薪資到月薪資記錄（完整版）
 */
function syncSalaryToMonthlyRecord(employeeId, yearMonth) {
  try {
    const salaryConfig = getEmployeeSalaryTW(employeeId);
    
    if (!salaryConfig.success) {
      return { success: false, message: "找不到員工薪資設定" };
    }
    
    const config = salaryConfig.data;
    const calculatedSalary = calculateMonthlySalary(employeeId, yearMonth);
    
    if (!calculatedSalary.success) {
      // 建立基本薪資記錄
      const totalAllowances = 
        (parseFloat(config['職務加給']) || 0) +
        (parseFloat(config['伙食費']) || 0) +
        (parseFloat(config['交通補助']) || 0) +
        (parseFloat(config['全勤獎金']) || 0) +
        (parseFloat(config['績效獎金']) || 0) +
        (parseFloat(config['其他津貼']) || 0);
      
      const totalDeductions = 
        (parseFloat(config['勞保費']) || 0) +
        (parseFloat(config['健保費']) || 0) +
        (parseFloat(config['就業保險費']) || 0) +
        (parseFloat(config['勞退自提']) || 0) +
        (parseFloat(config['所得稅']) || 0) +
        (parseFloat(config['福利金扣款']) || 0) +
        (parseFloat(config['宿舍費用']) || 0) +
        (parseFloat(config['團保費用']) || 0) +
        (parseFloat(config['其他扣款']) || 0);
      
      const grossSalary = parseFloat(config['基本薪資']) + totalAllowances;
      
      const basicSalary = {
        employeeId: employeeId,
        employeeName: config['員工姓名'],
        yearMonth: yearMonth,
        baseSalary: config['基本薪資'],
        positionAllowance: config['職務加給'] || 0,
        mealAllowance: config['伙食費'] || 0,
        transportAllowance: config['交通補助'] || 0,
        attendanceBonus: config['全勤獎金'] || 0,
        performanceBonus: config['績效獎金'] || 0,
        otherAllowances: config['其他津貼'] || 0,
        weekdayOvertimePay: 0,
        restdayOvertimePay: 0,
        holidayOvertimePay: 0,
        laborFee: config['勞保費'] || 0,
        healthFee: config['健保費'] || 0,
        employmentFee: config['就業保險費'] || 0,
        pensionSelf: config['勞退自提'] || 0,
        incomeTax: config['所得稅'] || 0,
        leaveDeduction: 0,
        welfareFee: config['福利金扣款'] || 0,
        dormitoryFee: config['宿舍費用'] || 0,
        groupInsurance: config['團保費用'] || 0,
        otherDeductions: config['其他扣款'] || 0,
        grossSalary: grossSalary,
        netSalary: grossSalary - totalDeductions,
        bankCode: config['銀行代碼'] || "",
        bankAccount: config['銀行帳號'] || "",
        status: "已設定",
        note: "自動建立"
      };
      
      return saveMonthlySalary(basicSalary);
    }
    
    return saveMonthlySalary(calculatedSalary.data);
    
  } catch (error) {
    Logger.log(`❌ 同步失敗: ${error}`);
    return { success: false, message: error.toString() };
  }
}

// ==================== 薪資計算功能 ====================

/**
 * ✅ 計算月薪資（完整版）
 */
function calculateMonthlySalary(employeeId, yearMonth) {
  try {
    const salaryConfig = getEmployeeSalaryTW(employeeId);
    if (!salaryConfig.success) {
      return { success: false, message: "找不到員工薪資設定" };
    }
    
    const config = salaryConfig.data;
    const overtimeRecords = getEmployeeOvertimeRecords(employeeId, yearMonth);
    const leaveRecords = getEmployeeMonthlySalary(employeeId, yearMonth);
    
    // 基本薪資
    let baseSalary = parseFloat(config['基本薪資']) || 0;
    
    // 固定津貼
    let positionAllowance = parseFloat(config['職務加給']) || 0;
    let mealAllowance = parseFloat(config['伙食費']) || 0;
    let transportAllowance = parseFloat(config['交通補助']) || 0;
    let attendanceBonus = parseFloat(config['全勤獎金']) || 0;
    let performanceBonus = parseFloat(config['績效獎金']) || 0;
    let otherAllowances = parseFloat(config['其他津貼']) || 0;
    
    // 加班費
    let weekdayOvertimePay = 0;
    let restdayOvertimePay = 0;
    let holidayOvertimePay = 0;
    
    if (overtimeRecords.success && overtimeRecords.data) {
      overtimeRecords.data.forEach(record => {
        if (record.reviewStatus === '核准') {
          const hours = record.overtimeHours;
          const hourlyRate = config['薪資類型'] === '月薪' 
            ? Math.round(baseSalary / 30 / 8) 
            : baseSalary;
          
          if (record.overtimeType === '平日加班') {
            weekdayOvertimePay += hours * hourlyRate * OVERTIME_RATES.weekday;
          } else if (record.overtimeType === '休息日加班') {
            restdayOvertimePay += hours * hourlyRate * OVERTIME_RATES.restday;
          } else if (record.overtimeType === '國定假日加班') {
            holidayOvertimePay += hours * hourlyRate * OVERTIME_RATES.holiday;
          }
        }
      });
    }
    
    // 請假扣款
    let leaveDeduction = 0;
    if (leaveRecords.success && leaveRecords.data) {
      leaveRecords.data.forEach(record => {
        if (record.reviewStatus === '核准') {
          if (record.leaveType === 'PERSONAL_LEAVE' || record.leaveType === '事假') {
            const dailyRate = config['薪資類型'] === '月薪' 
              ? Math.round(baseSalary / 30) 
              : baseSalary * 8;
            leaveDeduction += record.leaveDays * dailyRate;
          }
        }
      });
    }
    
    // 如果有請假，取消全勤獎金
    if (leaveDeduction > 0) {
      attendanceBonus = 0;
    }
    
    // 法定扣款
    const laborFee = parseFloat(config['勞保費']) || 0;
    const healthFee = parseFloat(config['健保費']) || 0;
    const employmentFee = parseFloat(config['就業保險費']) || 0;
    const pensionSelf = parseFloat(config['勞退自提']) || 0;
    const incomeTax = parseFloat(config['所得稅']) || 0;
    
    // 其他扣款
    const welfareFee = parseFloat(config['福利金扣款']) || 0;
    const dormitoryFee = parseFloat(config['宿舍費用']) || 0;
    const groupInsurance = parseFloat(config['團保費用']) || 0;
    const otherDeductions = parseFloat(config['其他扣款']) || 0;
    
    // 應發總額
    const grossSalary = baseSalary + positionAllowance + mealAllowance + transportAllowance + 
                       attendanceBonus + performanceBonus + otherAllowances +
                       weekdayOvertimePay + restdayOvertimePay + holidayOvertimePay;
    
    // 扣款總額
    const totalDeductions = laborFee + healthFee + employmentFee + pensionSelf + incomeTax +
                           leaveDeduction + welfareFee + dormitoryFee + groupInsurance + otherDeductions;
    
    // 實發金額
    const netSalary = grossSalary - totalDeductions;
    
    const result = {
      employeeId: employeeId,
      employeeName: config['員工姓名'],
      yearMonth: yearMonth,
      baseSalary: baseSalary,
      positionAllowance: positionAllowance,
      mealAllowance: mealAllowance,
      transportAllowance: transportAllowance,
      attendanceBonus: attendanceBonus,
      performanceBonus: performanceBonus,
      otherAllowances: otherAllowances,
      weekdayOvertimePay: Math.round(weekdayOvertimePay),
      restdayOvertimePay: Math.round(restdayOvertimePay),
      holidayOvertimePay: Math.round(holidayOvertimePay),
      laborFee: laborFee,
      healthFee: healthFee,
      employmentFee: employmentFee,
      pensionSelf: pensionSelf,
      incomeTax: incomeTax,
      leaveDeduction: Math.round(leaveDeduction),
      welfareFee: welfareFee,
      dormitoryFee: dormitoryFee,
      groupInsurance: groupInsurance,
      otherDeductions: otherDeductions,
      grossSalary: Math.round(grossSalary),
      netSalary: Math.round(netSalary),
      bankCode: config['銀行代碼'] || "",
      bankAccount: config['銀行帳號'] || "",
      status: "已計算",
      note: ""
    };
    
    return { success: true, data: result };
    
  } catch (error) {
    Logger.log("❌ 計算薪資失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 儲存月薪資單（完整版）
 */
function saveMonthlySalary(salaryData) {
  try {
    const sheet = getMonthlySalarySheetEnhanced();
    
    let normalizedYearMonth = salaryData.yearMonth;
    
    if (salaryData.yearMonth instanceof Date) {
      normalizedYearMonth = Utilities.formatDate(salaryData.yearMonth, "Asia/Taipei", "yyyy-MM");
    } else if (typeof salaryData.yearMonth === 'string') {
      normalizedYearMonth = salaryData.yearMonth.substring(0, 7);
    }
    
    const salaryId = `SAL-${normalizedYearMonth}-${salaryData.employeeId}`;
    
    const row = [
      // 基本資訊
      salaryId,
      salaryData.employeeId,
      salaryData.employeeName,
      normalizedYearMonth,
      
      // 應發項目
      salaryData.baseSalary || 0,
      salaryData.positionAllowance || 0,
      salaryData.mealAllowance || 0,
      salaryData.transportAllowance || 0,
      salaryData.attendanceBonus || 0,
      salaryData.performanceBonus || 0,
      salaryData.otherAllowances || 0,
      salaryData.weekdayOvertimePay || 0,
      salaryData.restdayOvertimePay || 0,
      salaryData.holidayOvertimePay || 0,
      
      // 法定扣款
      salaryData.laborFee || 0,
      salaryData.healthFee || 0,
      salaryData.employmentFee || 0,
      salaryData.pensionSelf || 0,
      salaryData.incomeTax || 0,
      
      // 其他扣款
      salaryData.leaveDeduction || 0,
      salaryData.welfareFee || 0,
      salaryData.dormitoryFee || 0,
      salaryData.groupInsurance || 0,
      salaryData.otherDeductions || 0,
      
      // 總計
      salaryData.grossSalary || 0,
      salaryData.netSalary || 0,
      
      // 銀行資訊
      salaryData.bankCode || "",
      salaryData.bankAccount || "",
      
      // 系統欄位
      salaryData.status || "已計算",
      salaryData.note || "",
      new Date()
    ];
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === salaryId) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        found = true;
        Logger.log(`✅ 更新薪資單: ${salaryId}`);
        break;
      }
    }
    
    if (!found) {
      sheet.appendRow(row);
      Logger.log(`✅ 新增薪資單: ${salaryId}`);
    }
    
    return { success: true, salaryId: salaryId, message: "薪資單儲存成功" };
    
  } catch (error) {
    Logger.log("❌ 儲存薪資單失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 查詢我的薪資（完整版）
 */
function getMySalary(userId, yearMonth) {
  try {
    const employeeId = userId;
    const sheet = getMonthlySalarySheetEnhanced();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return { success: false, message: "薪資記錄表中沒有資料" };
    }
    
    const headers = data[0];
    const employeeIdIndex = headers.indexOf('員工ID');
    const yearMonthIndex = headers.indexOf('年月');
    
    if (employeeIdIndex === -1 || yearMonthIndex === -1) {
      return { success: false, message: "試算表缺少必要欄位" };
    }
    
    for (let i = 1; i < data.length; i++) {
      const rowEmployeeId = String(data[i][employeeIdIndex]).trim();
      const rawYearMonth = data[i][yearMonthIndex];
      
      let normalizedYearMonth = '';
      
      if (rawYearMonth instanceof Date) {
        normalizedYearMonth = Utilities.formatDate(rawYearMonth, 'Asia/Taipei', 'yyyy-MM');
      } else if (typeof rawYearMonth === 'string') {
        normalizedYearMonth = rawYearMonth.substring(0, 7);
      } else {
        normalizedYearMonth = String(rawYearMonth).substring(0, 7);
      }
      
      if (rowEmployeeId === employeeId && normalizedYearMonth === yearMonth) {
        const salary = {};
        headers.forEach((header, index) => {
          if (header === '年月' && data[i][index] instanceof Date) {
            salary[header] = Utilities.formatDate(data[i][index], 'Asia/Taipei', 'yyyy-MM');
          } else {
            salary[header] = data[i][index];
          }
        });
        
        return { success: true, data: salary };
      }
    }
    
    return { success: false, message: "查無薪資記錄" };
    
  } catch (error) {
    Logger.log('❌ 查詢薪資失敗: ' + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 查詢我的薪資歷史（完整版）
 */
function getMySalaryHistory(userId, limit = 12) {
  try {
    const employeeId = userId;
    const sheet = getMonthlySalarySheetEnhanced();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return { success: true, data: [], total: 0 };
    }
    
    const headers = data[0];
    const employeeIdIndex = headers.indexOf('員工ID');
    
    if (employeeIdIndex === -1) {
      return { success: false, message: "試算表缺少「員工ID」欄位" };
    }
    
    const salaries = [];
    
    for (let i = 1; i < data.length; i++) {
      const rowEmployeeId = String(data[i][employeeIdIndex]).trim();
      
      if (rowEmployeeId === employeeId) {
        const salary = {};
        headers.forEach((header, index) => {
          if (header === '年月' && data[i][index] instanceof Date) {
            salary[header] = Utilities.formatDate(data[i][index], "Asia/Taipei", "yyyy-MM");
          } else {
            salary[header] = data[i][index];
          }
        });
        salaries.push(salary);
      }
    }
    
    salaries.sort((a, b) => {
      const yearMonthA = String(a['年月'] || '');
      const yearMonthB = String(b['年月'] || '');
      return yearMonthB.localeCompare(yearMonthA);
    });
    
    const result = salaries.slice(0, limit);
    
    return { success: true, data: result, total: salaries.length };
    
  } catch (error) {
    Logger.log("❌ 查詢薪資歷史失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ✅ 查詢所有員工的月薪資列表（完整版）
 */
function getAllMonthlySalary(yearMonth) {
  try {
    const sheet = getMonthlySalarySheetEnhanced();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const salaries = [];
    
    for (let i = 1; i < data.length; i++) {
      const rawYearMonth = data[i][3];
      
      let normalizedYearMonth = '';
      
      if (rawYearMonth instanceof Date) {
        normalizedYearMonth = Utilities.formatDate(rawYearMonth, "Asia/Taipei", "yyyy-MM");
      } else if (typeof rawYearMonth === 'string') {
        normalizedYearMonth = rawYearMonth.substring(0, 7);
      }
      
      if (!yearMonth || normalizedYearMonth === yearMonth) {
        const salary = {};
        headers.forEach((header, index) => {
          if (header === '年月') {
            salary[header] = normalizedYearMonth;
          } else {
            salary[header] = data[i][index];
          }
        });
        salaries.push(salary);
      }
    }
    
    return { success: true, data: salaries };
    
  } catch (error) {
    Logger.log("❌ 查詢薪資列表失敗: " + error);
    return { success: false, message: error.toString() };
  }
}

// ==================== 輔助函數 ====================

/**
 * ✅ 取得員工加班記錄
 */
function getEmployeeOvertimeRecords(employeeId, yearMonth) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("加班申請");
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const values = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      if (!row[1] || !row[3]) continue;
      
      const rowEmployeeId = String(row[1]).trim();
      const overtimeDate = row[3];
      
      if (rowEmployeeId !== employeeId) continue;
      
      let dateStr = "";
      if (overtimeDate instanceof Date) {
        dateStr = Utilities.formatDate(overtimeDate, "Asia/Taipei", "yyyy-MM");
      } else if (typeof overtimeDate === "string") {
        dateStr = overtimeDate.substring(0, 7);
      }
      
      if (dateStr !== yearMonth) continue;
      
      const status = String(row[9] || "").trim().toLowerCase();
      if (status !== "approved") continue;
      
      records.push({
        overtimeDate: dateStr,
        overtimeHours: parseFloat(row[6]) || 0,
        overtimeType: "平日加班",
        reviewStatus: "核准"
      });
    }
    
    return { success: true, data: records };
    
  } catch (error) {
    Logger.log("❌ 取得加班記錄失敗: " + error);
    return { success: false, message: error.toString(), data: [] };
  }
}

/**
 * ✅ 取得員工請假記錄
 */
function getEmployeeMonthlySalary(employeeId, yearMonth) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("請假記錄");
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const values = sheet.getDataRange().getValues();
    const records = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      if (!row[1] || !row[5]) continue;
      
      const rowEmployeeId = String(row[1]).trim();
      const startDate = row[5];
      
      if (rowEmployeeId !== employeeId) continue;
      
      let dateStr = "";
      if (startDate instanceof Date) {
        dateStr = Utilities.formatDate(startDate, "Asia/Taipei", "yyyy-MM");
      } else if (typeof startDate === "string") {
        dateStr = startDate.substring(0, 7);
      }
      
      if (dateStr !== yearMonth) continue;
      
      const status = String(row[9] || "").trim().toUpperCase();
      if (status !== "APPROVED") continue;
      
      records.push({
        leaveType: row[4] || "",
        startDate: startDate,
        leaveDays: parseFloat(row[7]) || 0,
        reviewStatus: "核准"
      });
    }
    
    return { success: true, data: records };
    
  } catch (error) {
    Logger.log("❌ 取得請假記錄失敗: " + error);
    return { success: false, message: error.toString(), data: [] };
  }
}

Logger.log('✅ 薪資管理系統（完整版 - 修正版）已載入');