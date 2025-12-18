// leave.js - 請假系統前端邏輯（15種假別完整版）

/**
 * 🆕 格式化日期函數
 * 將任何日期格式轉換為 YYYY-MM-DD
 */
function formatLeaveDate(dateInput) {
    if (!dateInput) return '';
    
    const date = new Date(dateInput);
    
    // 檢查是否為有效日期
    if (isNaN(date.getTime())) return dateInput;
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
}

/**
 * 初始化請假頁籤
 */
async function initLeaveTab() {
    console.log('📋 初始化請假系統（15種假別）...');
    
    // 載入假期餘額
    await loadLeaveBalance();
    
    // 載入請假紀錄
    await loadLeaveRecords();
    
    // 綁定事件監聽器
    bindLeaveEventListeners();
}

/**
 * 綁定請假相關事件
 */
function bindLeaveEventListeners() {
    // 提交請假申請
    const submitBtn = document.getElementById('submit-leave-btn');
    if (submitBtn) {
        submitBtn.replaceWith(submitBtn.cloneNode(true));
        document.getElementById('submit-leave-btn').addEventListener('click', handleSubmitLeave);
    }
    
    // 請假類型改變時的處理
    const leaveTypeSelect = document.getElementById('leave-type');
    if (leaveTypeSelect) {
        leaveTypeSelect.removeEventListener('change', handleLeaveTypeChange);
        leaveTypeSelect.addEventListener('change', handleLeaveTypeChange);
    }
    
    // 👇 修改：改為 datetime-local 輸入框
    const startDateTimeInput = document.getElementById('leave-start-datetime');
    const endDateTimeInput = document.getElementById('leave-end-datetime');
    
    if (startDateTimeInput && endDateTimeInput) {
        startDateTimeInput.removeEventListener('change', calculateLeaveHours);
        endDateTimeInput.removeEventListener('change', calculateLeaveHours);
        startDateTimeInput.addEventListener('change', calculateLeaveHours);
        endDateTimeInput.addEventListener('change', calculateLeaveHours);
    }
}

/**
 * 計算請假時數（扣除午休）
 */
function calculateLeaveHours() {
    const startDateTime = document.getElementById('leave-start-datetime').value;
    const endDateTime = document.getElementById('leave-end-datetime').value;
    const hoursEl = document.getElementById('calculated-hours');
    const daysEl = document.getElementById('calculated-days');
    
    if (!startDateTime || !endDateTime) {
        if (hoursEl) hoursEl.textContent = '0';
        if (daysEl) daysEl.textContent = '0';
        return;
    }
    
    const start = new Date(startDateTime);
    const end = new Date(endDateTime);
    
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        if (hoursEl) hoursEl.textContent = '0';
        if (daysEl) daysEl.textContent = '0';
        return;
    }
    
    if (end <= start) {
        showNotification('結束時間必須晚於開始時間', 'error');
        if (hoursEl) hoursEl.textContent = '0';
        if (daysEl) daysEl.textContent = '0';
        return;
    }
    
    // 計算總毫秒差
    const diffMs = end - start;
    let totalHours = diffMs / (1000 * 60 * 60);
    
    // 計算需要扣除的午休時數
    let lunchDeduction = 0;
    
    // 如果是同一天的請假
    if (start.toDateString() === end.toDateString()) {
        const startHour = start.getHours() + start.getMinutes() / 60;
        const endHour = end.getHours() + end.getMinutes() / 60;
        
        // 如果請假時段包含 12:00-13:00，扣除1小時
        if (startHour < 13 && endHour > 12) {
            lunchDeduction = 1;
        }
    } else {
        // 跨天請假，計算每個完整天數的午休時數
        const startOfDay = new Date(start);
        startOfDay.setHours(0, 0, 0, 0);
        
        const endOfDay = new Date(end);
        endOfDay.setHours(0, 0, 0, 0);
        
        const daysDiff = Math.floor((endOfDay - startOfDay) / (1000 * 60 * 60 * 24));
        
        // 計算第一天的午休扣除
        const firstDayStart = start.getHours() + start.getMinutes() / 60;
        if (firstDayStart < 13) {
            lunchDeduction += 1;
        }
        
        // 計算中間完整天數的午休（每天1小時）
        if (daysDiff > 1) {
            lunchDeduction += (daysDiff - 1);
        }
        
        // 計算最後一天的午休扣除
        const lastDayEnd = end.getHours() + end.getMinutes() / 60;
        if (lastDayEnd > 12) {
            lunchDeduction += 1;
        }
    }
    
    const workHours = Math.max(0, totalHours - lunchDeduction);
    const days = (workHours / 8).toFixed(2);
    
    if (hoursEl) hoursEl.textContent = workHours.toFixed(2);
    if (daysEl) daysEl.textContent = days;
}

/**
 * 載入假期餘額（15種假別）
 */
async function loadLeaveBalance() {
    const balanceContainer = document.getElementById('leave-balance-container');
    const loadingEl = document.getElementById('leave-balance-loading');
    
    if (loadingEl) loadingEl.style.display = 'block';
    
    try {
        const res = await callApifetch('getLeaveBalance');
        
        if (res.ok) {
            renderLeaveBalance(res.balance);
        } else {
            showNotification(t(res.code || 'ERROR_FETCH_LEAVE_BALANCE'), 'error');
        }
    } catch (err) {
        console.error('載入假期餘額失敗:', err);
        showNotification(t('NETWORK_ERROR'), 'error');
    } finally {
        if (loadingEl) loadingEl.style.display = 'none';
    }
}

/**
 * 渲染假期餘額（15種假別）
 */
function renderLeaveBalance(balance) {
    const listEl = document.getElementById('leave-balance-list');
    if (!listEl) return;
    
    listEl.innerHTML = '';
    
    // 定義假別順序（15種）
    const leaveOrder = [
        'ANNUAL_LEAVE',              // 特休假
        'COMP_TIME_OFF',             // 加班補休假
        'PERSONAL_LEAVE',            // 事假
        'SICK_LEAVE',                // 未住院病假
        'HOSPITALIZATION_LEAVE',     // 住院病假
        'BEREAVEMENT_LEAVE',         // 喪假
        'MARRIAGE_LEAVE',            // 婚假
        'PATERNITY_LEAVE',           // 陪產檢及陪產假
        'MATERNITY_LEAVE',           // 產假
        'OFFICIAL_LEAVE',            // 公假（含兵役假）
        'WORK_INJURY_LEAVE',         // 公傷假
        'ABSENCE_WITHOUT_LEAVE',     // 曠工
        'NATURAL_DISASTER_LEAVE',    // 天然災害停班
        'FAMILY_CARE_LEAVE',         // 家庭照顧假
        'MENSTRUAL_LEAVE'            // 生理假
    ];
    
    leaveOrder.forEach(leaveType => {
        if (balance[leaveType] !== undefined) {
            const item = document.createElement('div');
            item.className = 'flex justify-between items-center p-3 bg-gray-50 dark:bg-gray-700 rounded-lg';
            
            const typeSpan = document.createElement('span');
            typeSpan.className = 'font-medium text-gray-800 dark:text-white';
            typeSpan.setAttribute('data-i18n-key', leaveType);
            typeSpan.textContent = t(leaveType);
            
            // 特殊處理：曠工用紅色顯示
            const daysSpan = document.createElement('span');
            if (leaveType === 'ABSENCE_WITHOUT_LEAVE') {
                daysSpan.className = 'text-red-600 dark:text-red-400 font-bold';
                daysSpan.textContent = `${balance[leaveType]} ${t('DAYS')}`;
            } else {
                daysSpan.className = 'text-indigo-600 dark:text-indigo-400 font-bold';
                daysSpan.textContent = `${balance[leaveType]} ${t('DAYS')}`;
            }
            
            item.appendChild(typeSpan);
            item.appendChild(daysSpan);
            listEl.appendChild(item);
            
            renderTranslations(item);
        }
    });
}

/**
 * 載入請假紀錄
 */
async function loadLeaveRecords() {
    const userId = localStorage.getItem('sessionUserId');
    const loadingEl = document.getElementById('leave-records-loading');
    const emptyEl = document.getElementById('leave-records-empty');
    const listEl = document.getElementById('leave-records-list');
    
    try {
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        const res = await callApifetch(`getEmployeeLeaveRecords&employeeId=${userId}`);
        
        console.log('📥 API 回應:', res);
        
        loadingEl.style.display = 'none';
        
        // ✅ 修正：根據後端實際返回格式判斷
        if (res.ok && res.records && res.records.length > 0) {
            emptyEl.style.display = 'none';
            renderLeaveRecords(res.records);
        } else {
            console.log('ℹ️ 沒有請假紀錄');
            emptyEl.style.display = 'block';
        }
        
    } catch (error) {
        console.error('❌ 載入請假紀錄失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}

function formatDateTimeLocal(isoString) {
    const date = new Date(isoString);
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    
    return `${year}-${month}-${day} ${hours}:${minutes}`;
}
/**
 * 渲染請假記錄（只顯示小時數）
 */
function renderLeaveRecords(records) {
    const listEl = document.getElementById('leave-records-list');
    if (!listEl) return;
    
    listEl.innerHTML = '';
    
    if (!records || records.length === 0) {
        listEl.innerHTML = `
            <div class="text-center text-gray-500 dark:text-gray-400 py-8">
                <span data-i18n="NO_LEAVE_RECORDS">目前沒有請假記錄</span>
            </div>
        `;
        renderTranslations(listEl);
        return;
    }
    
    records.forEach(record => {
        const card = document.createElement('div');
        card.className = 'card p-4 hover:shadow-lg transition-shadow';
        
        // 狀態標籤樣式
        let statusClass = 'bg-yellow-100 text-yellow-800 dark:bg-yellow-900 dark:text-yellow-300';
        let statusText = '待審核';
        
        if (record.status === 'APPROVED') {
            statusClass = 'bg-green-100 text-green-800 dark:bg-green-900 dark:text-green-300';
            statusText = '已核准';
        } else if (record.status === 'REJECTED') {
            statusClass = 'bg-red-100 text-red-800 dark:bg-red-900 dark:text-red-300';
            statusText = '已拒絕';
        }
        
        // ⭐ 只顯示小時數，移除天數
        const workHoursDisplay = `${record.workHours} 小時`;
        
        card.innerHTML = `
            <div class="flex justify-between items-start mb-3">
                <h3 class="font-bold text-lg text-gray-800 dark:text-white" 
                    data-i18n-key="${record.leaveType}">
                    ${t(record.leaveType)}
                </h3>
                <span class="px-3 py-1 rounded-full text-sm font-medium ${statusClass}">
                    ${statusText}
                </span>
            </div>
            
            <div class="space-y-2 text-sm text-gray-600 dark:text-gray-400">
                <div class="flex items-center">
                    <svg class="w-4 h-4 mr-2" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M6 2a1 1 0 00-1 1v1H4a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V6a2 2 0 00-2-2h-1V3a1 1 0 10-2 0v1H7V3a1 1 0 00-1-1zm0 5a1 1 0 000 2h8a1 1 0 100-2H6z" clip-rule="evenodd"/>
                    </svg>
                    <span>${record.startTime} ~ ${record.endTime}</span>
                </div>
                
                <div class="flex items-center">
                    <svg class="w-4 h-4 mr-2" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm1-12a1 1 0 10-2 0v4a1 1 0 00.293.707l2.828 2.829a1 1 0 101.415-1.415L11 9.586V6z" clip-rule="evenodd"/>
                    </svg>
                    <span class="font-medium text-indigo-600 dark:text-indigo-400">
                        ${workHoursDisplay}
                    </span>
                </div>
                
                <div class="flex items-start">
                    <svg class="w-4 h-4 mr-2 mt-0.5" fill="currentColor" viewBox="0 0 20 20">
                        <path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clip-rule="evenodd"/>
                    </svg>
                    <span data-i18n="REASON_LABEL">請假原因:</span>
                    <span class="ml-1">${record.reason}</span>
                </div>
                
                ${record.reviewerComment ? `
                    <div class="mt-3 p-3 bg-gray-50 dark:bg-gray-700 rounded-lg">
                        <div class="font-medium text-gray-700 dark:text-gray-300 mb-1">
                            <span data-i18n="REVIEWER_COMMENT">審核意見:</span>
                        </div>
                        <div class="text-gray-600 dark:text-gray-400">${record.reviewerComment}</div>
                    </div>
                ` : ''}
            </div>
        `;
        
        listEl.appendChild(card);
        renderTranslations(card);
    });
}

/**
 * 處理請假類型變更
 */
function handleLeaveTypeChange(e) {
    const leaveType = e.target.value;
    const reasonContainer = document.getElementById('leave-reason-container');
    
    // 需要特別說明的假別
    const requiresReason = [
        'BEREAVEMENT_LEAVE',          // 喪假
        'MATERNITY_LEAVE',            // 產假
        'PATERNITY_LEAVE',            // 陪產檢及陪產假
        'HOSPITALIZATION_LEAVE',      // 住院病假
        'WORK_INJURY_LEAVE',          // 公傷假
        'OFFICIAL_LEAVE'              // 公假（含兵役假）
    ];
    
    if (requiresReason.includes(leaveType)) {
        reasonContainer.classList.remove('hidden');
    } else {
        reasonContainer.classList.remove('hidden'); // 都顯示以便填寫
    }
}

/**
 * 提交請假申請（小時制修正版）
 */
async function handleSubmitLeave() {
    const button = document.getElementById('submit-leave-btn');
    const loadingText = t('LOADING') || '處理中...';
    
    // 取得表單資料
    const leaveType = document.getElementById('leave-type').value;
    const startDateTime = document.getElementById('leave-start-datetime').value;
    const endDateTime = document.getElementById('leave-end-datetime').value;
    const reason = document.getElementById('leave-reason').value;
    
    // 驗證
    if (!leaveType) {
        showNotification('請選擇請假類型', 'error');
        return;
    }
    
    if (!startDateTime || !endDateTime) {
        showNotification('請選擇開始和結束時間', 'error');
        return;
    }
    
    const start = new Date(startDateTime);
    const end = new Date(endDateTime);
    
    if (end <= start) {
        showNotification('結束時間必須晚於開始時間', 'error');
        return;
    }
    
    if (!reason || reason.length < 2) {
        showNotification('請填寫請假原因（至少2個字）', 'error');
        return;
    }
    
    // 進入處理中狀態
    generalButtonState(button, 'processing', loadingText);
    
    try {
        const res = await callApifetch(
            `submitLeave&leaveType=${encodeURIComponent(leaveType)}` +
            `&startDateTime=${encodeURIComponent(startDateTime)}` +
            `&endDateTime=${encodeURIComponent(endDateTime)}` +
            `&reason=${encodeURIComponent(reason)}`
        );
        
        if (res.ok) {
            showNotification(
                `請假申請成功！時數：${res.hours || 0}小時 (${res.days || 0}天)`,
                'success'
            );
            
            // 清空表單
            document.getElementById('leave-type').value = '';
            document.getElementById('leave-start-datetime').value = '';
            document.getElementById('leave-end-datetime').value = '';
            document.getElementById('leave-reason').value = '';
            document.getElementById('calculated-hours').textContent = '0';
            document.getElementById('calculated-days').textContent = '0';
            
            // 重新載入資料
            await loadLeaveBalance();
            await loadLeaveRecords();
        } else {
            showNotification(t(res.code || 'LEAVE_SUBMIT_FAILED', res.params), 'error');
        }
    } catch (err) {
        console.error('提交請假失敗:', err);
        showNotification('提交失敗，請稍後再試', 'error');
    } finally {
        generalButtonState(button, 'idle');
    }
}

/**
 * 載入待審核的請假申請（管理員用）
 */
async function loadPendingLeaveRequests() {
    const loadingEl = document.getElementById('leave-requests-loading');
    const emptyEl = document.getElementById('leave-requests-empty');
    const listEl = document.getElementById('pending-leave-list');
    
    if (loadingEl) loadingEl.style.display = 'block';
    if (emptyEl) emptyEl.style.display = 'none';
    if (listEl) listEl.innerHTML = '';
    
    try {
        const res = await callApifetch('getPendingLeaveRequests');
        
        if (res.ok) {
            if (res.requests && res.requests.length > 0) {
                renderPendingLeaveRequests(res.requests);
            } else {
                if (emptyEl) emptyEl.style.display = 'block';
            }
        } else {
            showNotification(t(res.code || 'ERROR_FETCH_REQUESTS'), 'error');
        }
    } catch (err) {
        console.error('載入待審核請假失敗:', err);
        showNotification(t('NETWORK_ERROR'), 'error');
    } finally {
        if (loadingEl) loadingEl.style.display = 'none';
    }
}

/**
 * 渲染待審核請假列表（支援時數顯示）
 */
function renderPendingLeaveRequests(requests) {
    const listEl = document.getElementById('pending-leave-list');
    if (!listEl) return;
    
    listEl.innerHTML = '';
    
    requests.forEach(req => {
        const li = document.createElement('li');
        li.className = 'p-4 bg-gray-50 dark:bg-gray-700 rounded-lg';
        
        // 格式化時間顯示
        let timeDisplay = '';
        if (req.startDateTime && req.endDateTime) {
            timeDisplay = `${formatDateTimeLocal(req.startDateTime)} ~ ${formatDateTimeLocal(req.endDateTime)}`;
        } else if (req.startDate && req.endDate) {
            timeDisplay = `${formatLeaveDate(req.startDate)} ~ ${formatLeaveDate(req.endDate)}`;
        }
        
        // 顯示時數或天數
        let durationDisplay = '';
        if (req.workHours !== undefined) {
            durationDisplay = `${req.workHours} 小時 (${req.days} 天)`;
        } else if (req.days !== undefined) {
            durationDisplay = `${req.days} 天`;
        }
        
        li.innerHTML = `
            <div class="flex flex-col space-y-2">
                <div class="flex items-center justify-between">
                    <div class="flex-1">
                        <p class="font-semibold text-gray-800 dark:text-white">
                            ${req.employeeName} - <span data-i18n-key="${req.leaveType}">${t(req.leaveType)}</span>
                        </p>
                        <p class="text-sm text-gray-600 dark:text-gray-400">
                            ${timeDisplay}
                        </p>
                        ${durationDisplay ? `
                            <p class="text-xs text-gray-500 dark:text-gray-500 mt-1">
                                ${durationDisplay}
                            </p>
                        ` : ''}
                        ${req.reason ? `
                            <p class="text-sm text-gray-600 dark:text-gray-400 mt-1">
                                <span data-i18n="LEAVE_REASON_LABEL">原因：</span>${req.reason}
                            </p>
                        ` : ''}
                    </div>
                </div>
                
                <div class="flex space-x-2 mt-2">
                    <button 
                        data-i18n="ADMIN_APPROVE_BUTTON" 
                        data-row="${req.rowNumber}" 
                        class="approve-leave-btn flex-1 px-3 py-2 rounded-md text-sm font-bold btn-primary">
                        核准
                    </button>
                    <button 
                        data-i18n="ADMIN_REJECT_BUTTON" 
                        data-row="${req.rowNumber}" 
                        class="reject-leave-btn flex-1 px-3 py-2 rounded-md text-sm font-bold btn-warning">
                        拒絕
                    </button>
                </div>
            </div>
        `;
        
        listEl.appendChild(li);
        renderTranslations(li);
    });
    
    // 綁定審核按鈕事件
    listEl.querySelectorAll('.approve-leave-btn').forEach(btn => {
        btn.addEventListener('click', (e) => handleReviewLeave(e.currentTarget, 'approve'));
    });
    
    listEl.querySelectorAll('.reject-leave-btn').forEach(btn => {
        btn.addEventListener('click', (e) => handleReviewLeave(e.currentTarget, 'reject'));
    });
}

/**
 * 處理請假審核
 */
async function handleReviewLeave(button, action) {
    const rowNumber = button.dataset.row;
    const loadingText = t('LOADING') || '處理中...';
    
    // 詢問審核意見
    const comment = action === 'reject' 
        ? prompt(t('ENTER_REJECT_REASON') || '請輸入拒絕原因：') 
        : '';
    
    if (action === 'reject' && !comment) {
        showNotification(t('ERR_MISSING_REJECT_REASON'), 'warning');
        return;
    }
    
    generalButtonState(button, 'processing', loadingText);
    
    try {
        const res = await callApifetch(
            `reviewLeave&rowNumber=${rowNumber}` +
            `&reviewAction=${action}` +
            `&comment=${encodeURIComponent(comment || '')}`
        );
        
        if (res.ok) {
            const translationKey = action === 'approve' 
                ? 'LEAVE_APPROVED' 
                : 'LEAVE_REJECTED';
            showNotification(t(translationKey), 'success');
            
            // 延遲後重新載入列表
            await new Promise(resolve => setTimeout(resolve, 500));
            loadPendingLeaveRequests();
        } else {
            showNotification(t(res.code || 'REVIEW_FAILED', res.params), 'error');
        }
    } catch (err) {
        console.error('審核請假失敗:', err);
        showNotification(t('NETWORK_ERROR'), 'error');
    } finally {
        generalButtonState(button, 'idle');
    }
}


/**
 * 驗證請假表單（限制整數小時 + 8小時上限）
 */
function validateLeaveForm() {
    const leaveType = document.getElementById('leave-type').value;
    const startTime = document.getElementById('leave-start').value;
    const endTime = document.getElementById('leave-end').value;
    const reason = document.getElementById('leave-reason').value;
    
    // 基本驗證
    if (!leaveType) {
        showNotification('請選擇假別', 'error');
        return false;
    }
    
    if (!startTime) {
        showNotification('請選擇開始時間', 'error');
        return false;
    }
    
    if (!endTime) {
        showNotification('請選擇結束時間', 'error');
        return false;
    }
    
    if (!reason.trim()) {
        showNotification('請填寫請假原因', 'error');
        return false;
    }
    
    // 計算工時
    const workHours = calculateWorkHours(startTime, endTime);
    
    if (workHours <= 0) {
        showNotification('請假時數必須大於 0', 'error');
        return false;
    }
    
    // ⭐ 檢查是否為整數小時
    if (!Number.isInteger(workHours)) {
        showNotification(`請假時數必須是整數小時，目前為 ${workHours} 小時`, 'error');
        return false;
    }
    
    // ⭐ 檢查是否為同一天
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    if (start.toDateString() === end.toDateString()) {
        // 同一天：限制不超過 8 小時
        if (workHours > 8) {
            showNotification('單日請假不能超過 8 小時（扣除午休後）', 'error');
            return false;
        }
    }
    
    return true;
}

/**
 * 計算工作時數（排除午休時間 12:00-13:00）
 * 確保計算精確的整數小時
 */
function calculateWorkHours(startTime, endTime) {
    if (!startTime || !endTime) {
        return 0;
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    // 檢查日期是否有效
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        console.error('❌ 無效的日期格式');
        return 0;
    }
    
    // 檢查結束時間是否早於開始時間
    if (end <= start) {
        console.error('❌ 結束時間必須晚於開始時間');
        return 0;
    }
    
    // 計算總時長（毫秒）
    const totalMs = end - start;
    
    // 轉換為小時
    let totalHours = totalMs / (1000 * 60 * 60);
    
    console.log('📊 初始計算:', {
        start: start.toISOString(),
        end: end.toISOString(),
        totalMs: totalMs,
        totalHours: totalHours
    });
    
    // 如果是同一天，檢查是否跨越午休時間 12:00-13:00
    if (start.toDateString() === end.toDateString()) {
        const startHour = start.getHours() + start.getMinutes() / 60;
        const endHour = end.getHours() + end.getMinutes() / 60;
        
        const lunchStart = 12; // 12:00
        const lunchEnd = 13;   // 13:00
        
        // 判斷是否跨越午休時間
        if (startHour < lunchEnd && endHour > lunchStart) {
            // 計算重疊的午休時間
            const overlapStart = Math.max(startHour, lunchStart);
            const overlapEnd = Math.min(endHour, lunchEnd);
            const lunchOverlap = Math.max(0, overlapEnd - overlapStart);
            
            totalHours -= lunchOverlap;
            
            console.log('🍱 扣除午休時間:', lunchOverlap.toFixed(2), '小時');
        }
    } else {
        // 跨日請假：每天都要扣除 1 小時午休
        const startDate = new Date(start);
        startDate.setHours(0, 0, 0, 0);
        
        const endDate = new Date(end);
        endDate.setHours(0, 0, 0, 0);
        
        const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        
        // 每天扣除 1 小時午休
        totalHours -= daysDiff;
        
        console.log('📅 跨日請假，扣除', daysDiff, '天的午休時間');
    }
    
    // 確保不會是負數
    totalHours = Math.max(0, totalHours);
    
    // 四捨五入到小數點後 2 位
    const finalHours = Math.round(totalHours * 100) / 100;
    
    console.log('✅ 最終工時:', finalHours, '小時');
    
    return finalHours;
}

/**
 * 更新工時預覽（即時顯示）
 */
function updateWorkHoursPreview() {
    console.log('🔄 updateWorkHoursPreview 被觸發');
    
    const startTime = document.getElementById('leave-start-datetime').value;
    const endTime = document.getElementById('leave-end-datetime').value;
    const previewEl = document.getElementById('work-hours-preview');
    const hoursEl = document.getElementById('calculated-hours');
    const warningEl = document.getElementById('work-hours-warning');
    
    console.log('📥 輸入值:', {
        startTime: startTime,
        endTime: endTime
    });
    
    // 如果沒有輸入，隱藏預覽
    if (!startTime || !endTime) {
        console.log('⚠️ 開始或結束時間為空');
        if (previewEl) previewEl.classList.add('hidden');
        return;
    }
    
    // 計算工時
    const workHours = calculateWorkHours(startTime, endTime);
    
    console.log('💡 計算結果:', workHours, '小時');
    
    // 顯示預覽區塊
    if (previewEl) previewEl.classList.remove('hidden');
    if (hoursEl) hoursEl.textContent = workHours;
    
    // 清除之前的警告
    if (warningEl) {
        warningEl.classList.add('hidden');
        warningEl.textContent = '';
    }
    
    // 檢查各種錯誤情況
    let hasError = false;
    let errorMsg = '';
    
    if (workHours <= 0) {
        hasError = true;
        errorMsg = '❌ 結束時間必須晚於開始時間';
    } else if (!Number.isInteger(workHours)) {
        hasError = true;
        errorMsg = `❌ 請假時數必須是整數小時（目前為 ${workHours} 小時）\n請調整時間使其為整數小時`;
    } else {
        const start = new Date(startTime);
        const end = new Date(endTime);
        
        // 檢查單日上限
        if (start.toDateString() === end.toDateString() && workHours > 8) {
            hasError = true;
            errorMsg = '❌ 單日請假不能超過 8 小時（已扣除午休）';
        }
    }
    
    // 顯示警告訊息
    if (hasError) {
        if (warningEl) {
            warningEl.classList.remove('hidden');
            warningEl.textContent = errorMsg;
        }
        if (hoursEl) {
            hoursEl.classList.add('text-red-600', 'dark:text-red-400');
            hoursEl.classList.remove('text-blue-800', 'dark:text-blue-300');
        }
    } else {
        // 顯示成功狀態
        if (hoursEl) {
            hoursEl.classList.remove('text-red-600', 'dark:text-red-400');
            hoursEl.classList.add('text-blue-800', 'dark:text-blue-300');
        }
        
        // 顯示成功提示
        if (warningEl) {
            warningEl.classList.remove('hidden');
            warningEl.classList.remove('text-red-600', 'dark:text-red-400');
            warningEl.classList.add('text-green-600', 'dark:text-green-400');
            warningEl.textContent = '✅ 時數計算正確，可以提交申請';
        }
    }
}

/**
 * 快速選擇時段
 */
function quickSelectTimeRange(type) {
    console.log('🎯 快速選擇:', type);
    
    const now = new Date();
    const today = now.toISOString().split('T')[0]; // YYYY-MM-DD
    
    let startTime, endTime;
    
    switch(type) {
        case '1h':
            // 1小時：09:00 - 10:00
            startTime = `${today}T09:00`;
            endTime = `${today}T10:00`;
            break;
            
        case '2h':
            // 2小時：09:00 - 11:00
            startTime = `${today}T09:00`;
            endTime = `${today}T11:00`;
            break;
            
        case '4h':
            // 半天 4小時：13:00 - 17:00（下午，不跨午休）
            startTime = `${today}T13:00`;
            endTime = `${today}T17:00`;
            break;
            
        case '8h':
            // 全天 8小時：09:00 - 18:00（扣除午休1小時 = 實際8小時）
            startTime = `${today}T09:00`;
            endTime = `${today}T18:00`;
            break;
            
        default:
            console.error('❌ 未知的時段類型:', type);
            return;
    }
    
    console.log('📅 設定時間:', {
        startTime: startTime,
        endTime: endTime
    });
    
    // 設定時間
    const startInput = document.getElementById('leave-start-datetime');
    const endInput = document.getElementById('leave-end-datetime');
    
    if (startInput) startInput.value = startTime;
    if (endInput) endInput.value = endTime;
    
    // 更新工時預覽
    updateWorkHoursPreview();
}

/**
 * 提交請假申請
 */
async function submitLeaveApplication() {
    console.log('📤 開始提交請假申請');
    
    // 驗證表單
    if (!validateLeaveForm()) {
        console.error('❌ 表單驗證失敗');
        return;
    }
    
    const leaveType = document.getElementById('leave-type').value;
    const startTime = document.getElementById('leave-start-datetime').value;
    const endTime = document.getElementById('leave-end-datetime').value;
    const reason = document.getElementById('leave-reason').value;
    const workHours = calculateWorkHours(startTime, endTime);
    
    console.log('📋 提交資料:', {
        leaveType,
        startTime,
        endTime,
        workHours,
        reason
    });
    
    const button = document.getElementById('submit-leave-btn');
    if (button) {
        button.disabled = true;
        button.textContent = '處理中...';
    }
    
    try {
        const response = await callApifetch(
            `submitLeave&leaveType=${encodeURIComponent(leaveType)}` +
            `&startDateTime=${encodeURIComponent(startTime)}` +
            `&endDateTime=${encodeURIComponent(endTime)}` +
            `&reason=${encodeURIComponent(reason)}`
        );
        
        console.log('📥 後端回應:', response);
        
        if (response.ok) {
            showNotification('請假申請已提交', 'success');
            
            // 清空表單
            document.getElementById('leave-type').value = '';
            document.getElementById('leave-start-datetime').value = '';
            document.getElementById('leave-end-datetime').value = '';
            document.getElementById('leave-reason').value = '';
            document.getElementById('work-hours-preview').classList.add('hidden');
            
            // 重新載入請假記錄
            loadLeaveRecords();
            loadLeaveBalance();
        } else {
            showNotification(response.msg || '提交失敗', 'error');
        }
    } catch (error) {
        console.error('❌ 提交請假申請失敗:', error);
        showNotification('網路錯誤，請稍後再試', 'error');
    } finally {
        if (button) {
            button.disabled = false;
            button.textContent = '提交請假申請';
        }
    }
}

/**
 * 驗證請假表單
 */
function validateLeaveForm() {
    const leaveType = document.getElementById('leave-type').value;
    const startTime = document.getElementById('leave-start-datetime').value;
    const endTime = document.getElementById('leave-end-datetime').value;
    const reason = document.getElementById('leave-reason').value;
    
    if (!leaveType) {
        showNotification('請選擇假別', 'error');
        return false;
    }
    
    if (!startTime) {
        showNotification('請選擇開始時間', 'error');
        return false;
    }
    
    if (!endTime) {
        showNotification('請選擇結束時間', 'error');
        return false;
    }
    
    if (!reason.trim() || reason.trim().length < 2) {
        showNotification('請填寫請假原因（至少2個字）', 'error');
        return false;
    }
    
    const workHours = calculateWorkHours(startTime, endTime);
    
    if (workHours <= 0) {
        showNotification('請假時數必須大於 0', 'error');
        return false;
    }
    
    if (!Number.isInteger(workHours)) {
        showNotification(`請假時數必須是整數小時，目前為 ${workHours} 小時`, 'error');
        return false;
    }
    
    const start = new Date(startTime);
    const end = new Date(endTime);
    
    if (start.toDateString() === end.toDateString() && workHours > 8) {
        showNotification('單日請假不能超過 8 小時（已扣除午休）', 'error');
        return false;
    }
    
    return true;
}