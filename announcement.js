/**
 * 佈告欄功能模組
 * 支援管理員新增公告、員工查看公告
 */

// ==================== 📢 員工端：顯示公告 ====================

/**
 * 載入並顯示最新公告（儀表板）
 */
async function loadLatestAnnouncements() {
    const container = document.getElementById('announcements-container');
    const loadingEl = document.getElementById('announcements-loading');
    const emptyEl = document.getElementById('announcements-empty');
    const listEl = document.getElementById('announcements-list');
    
    if (!container || !loadingEl || !emptyEl || !listEl) {
        console.error('❌ 找不到佈告欄相關元素');
        return;
    }
    
    try {
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        // 呼叫 API 取得公告（只取最新 3 則）
        const res = await callApifetch('getAnnouncements&limit=3');
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.data && res.data.length > 0) {
            displayAnnouncements(res.data, listEl);
            emptyEl.style.display = 'none';
        } else {
            emptyEl.style.display = 'block';
            listEl.innerHTML = '';
        }
        
    } catch (error) {
        console.error('❌ 載入公告失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}

/**
 * 顯示公告列表
 */
function displayAnnouncements(announcements, container) {
    container.innerHTML = '';
    
    announcements.forEach(announcement => {
        const item = document.createElement('div');
        item.className = 'announcement-item bg-gradient-to-r from-indigo-50 to-purple-50 dark:from-indigo-900/20 dark:to-purple-900/20 rounded-lg p-4 border border-indigo-200 dark:border-indigo-700 mb-3';
        
        // 判斷重要性樣式
        let priorityClass = '';
        let priorityIcon = '';
        
        switch(announcement.priority) {
            case 'high':
                priorityClass = 'text-red-600 dark:text-red-400';
                priorityIcon = '🔴';
                break;
            case 'medium':
                priorityClass = 'text-yellow-600 dark:text-yellow-400';
                priorityIcon = '🟡';
                break;
            default:
                priorityClass = 'text-blue-600 dark:text-blue-400';
                priorityIcon = '🔵';
        }
        
        item.innerHTML = `
            <div class="flex items-start justify-between mb-2">
                <h3 class="font-bold text-gray-800 dark:text-white flex items-center gap-2">
                    <span>${priorityIcon}</span>
                    <span>${escapeHtml(announcement.title)}</span>
                </h3>
                <span class="text-xs ${priorityClass} font-semibold">
                    ${t('PRIORITY_' + announcement.priority.toUpperCase()) || announcement.priority}
                </span>
            </div>
            <p class="text-sm text-gray-600 dark:text-gray-300 mb-2">
                ${escapeHtml(announcement.content)}
            </p>
            <div class="flex justify-between items-center text-xs text-gray-500 dark:text-gray-400">
                <span>${formatAnnouncementDate(announcement.createdAt)}</span>
                <span>${announcement.authorName || t('SYSTEM')}</span>
            </div>
        `;
        
        container.appendChild(item);
    });
}

/**
 * 格式化公告日期
 */
function formatAnnouncementDate(dateString) {
    const date = new Date(dateString);
    const now = new Date();
    const diffMs = now - date;
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    
    if (diffHours < 1) {
        return t('JUST_NOW') || '剛剛';
    } else if (diffHours < 24) {
        return t('HOURS_AGO', { hours: diffHours }) || `${diffHours} 小時前`;
    } else if (diffDays < 7) {
        return t('DAYS_AGO', { days: diffDays }) || `${diffDays} 天前`;
    } else {
        const month = date.getMonth() + 1;
        const day = date.getDate();
        return `${month}/${day}`;
    }
}

/**
 * HTML 轉義（防止 XSS）
 */
function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}

// ==================== 📢 管理員端：管理公告 ====================

/**
 * 載入所有公告（管理員）
 */
async function loadAllAnnouncements() {
    const loadingEl = document.getElementById('admin-announcements-loading');
    const emptyEl = document.getElementById('admin-announcements-empty');
    const listEl = document.getElementById('admin-announcements-list');
    
    if (!loadingEl || !emptyEl || !listEl) {
        console.error('❌ 找不到管理員公告元素');
        return;
    }
    
    try {
        loadingEl.style.display = 'block';
        emptyEl.style.display = 'none';
        listEl.innerHTML = '';
        
        const res = await callApifetch('getAnnouncements');
        
        loadingEl.style.display = 'none';
        
        if (res.ok && res.data && res.data.length > 0) {
            displayAdminAnnouncements(res.data, listEl);
            emptyEl.style.display = 'none';
        } else {
            emptyEl.style.display = 'block';
            listEl.innerHTML = '';
        }
        
    } catch (error) {
        console.error('❌ 載入公告失敗:', error);
        loadingEl.style.display = 'none';
        emptyEl.style.display = 'block';
    }
}

/**
 * 顯示管理員公告列表
 */
function displayAdminAnnouncements(announcements, container) {
    container.innerHTML = '';
    
    announcements.forEach(announcement => {
        const item = document.createElement('div');
        item.className = 'bg-white dark:bg-gray-800 rounded-lg p-4 border border-gray-200 dark:border-gray-700 mb-3';
        
        // 重要性徽章
        let priorityBadge = '';
        switch(announcement.priority) {
            case 'high':
                priorityBadge = '<span class="px-2 py-1 text-xs font-semibold bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400 rounded">🔴 重要</span>';
                break;
            case 'medium':
                priorityBadge = '<span class="px-2 py-1 text-xs font-semibold bg-yellow-100 dark:bg-yellow-900/30 text-yellow-600 dark:text-yellow-400 rounded">🟡 中等</span>';
                break;
            default:
                priorityBadge = '<span class="px-2 py-1 text-xs font-semibold bg-blue-100 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 rounded">🔵 一般</span>';
        }
        
        item.innerHTML = `
            <div class="flex items-start justify-between mb-3">
                <div class="flex-1">
                    <h3 class="font-bold text-gray-800 dark:text-white mb-1">
                        ${escapeHtml(announcement.title)}
                    </h3>
                    <p class="text-sm text-gray-600 dark:text-gray-300 mb-2">
                        ${escapeHtml(announcement.content)}
                    </p>
                    <div class="flex items-center gap-3 text-xs text-gray-500 dark:text-gray-400">
                        <span>📅 ${formatAnnouncementDate(announcement.createdAt)}</span>
                        <span>👤 ${announcement.authorName || t('SYSTEM')}</span>
                        ${priorityBadge}
                    </div>
                </div>
                <div class="flex flex-col gap-2 ml-4">
                    <button class="edit-announcement-btn px-3 py-1 text-sm font-semibold bg-blue-500 hover:bg-blue-600 text-white rounded"
                            data-id="${announcement.id}"
                            data-title="${escapeHtml(announcement.title)}"
                            data-content="${escapeHtml(announcement.content)}"
                            data-priority="${announcement.priority}">
                        ✏️ ${t('BTN_EDIT') || '編輯'}
                    </button>
                    <button class="delete-announcement-btn px-3 py-1 text-sm font-semibold bg-red-500 hover:bg-red-600 text-white rounded"
                            data-id="${announcement.id}">
                        🗑️ ${t('BTN_DELETE') || '刪除'}
                    </button>
                </div>
            </div>
        `;
        
        container.appendChild(item);
    });
    
    // 綁定編輯和刪除按鈕事件
    bindAnnouncementActions();
}

/**
 * 新增公告
 */
async function createAnnouncement() {
    const titleInput = document.getElementById('announcement-title');
    const contentInput = document.getElementById('announcement-content');
    const prioritySelect = document.getElementById('announcement-priority');
    const submitBtn = document.getElementById('submit-announcement-btn');
    
    const title = titleInput.value.trim();
    const content = contentInput.value.trim();
    const priority = prioritySelect.value;
    
    if (!title || !content) {
        showNotification(t('ANNOUNCEMENT_EMPTY_ERROR') || '標題和內容不可為空', 'error');
        return;
    }
    
    const loadingText = t('LOADING') || '處理中...';
    generalButtonState(submitBtn, 'processing', loadingText);
    
    try {
        const params = new URLSearchParams({
            title: title,
            content: content,
            priority: priority
        });
        
        const res = await callApifetch(`createAnnouncement&${params.toString()}`);
        
        if (res.ok) {
            showNotification(t('ANNOUNCEMENT_CREATED') || '公告已發布！', 'success');
            
            // 清空表單
            titleInput.value = '';
            contentInput.value = '';
            prioritySelect.value = 'normal';
            
            // 重新載入公告列表
            await loadAllAnnouncements();
            
        } else {
            showNotification(t(res.code) || '發布失敗', 'error');
        }
        
    } catch (error) {
        console.error('❌ 發布公告失敗:', error);
        showNotification(t('NETWORK_ERROR') || '網路錯誤', 'error');
        
    } finally {
        generalButtonState(submitBtn, 'idle');
    }
}

/**
 * 編輯公告
 */
async function editAnnouncement(id, title, content, priority) {
    // 填入表單
    document.getElementById('announcement-title').value = title;
    document.getElementById('announcement-content').value = content;
    document.getElementById('announcement-priority').value = priority;
    
    // 修改提交按鈕為「更新」模式
    const submitBtn = document.getElementById('submit-announcement-btn');
    submitBtn.textContent = t('BTN_UPDATE') || '更新公告';
    submitBtn.dataset.mode = 'update';
    submitBtn.dataset.id = id;
    
    // 顯示取消按鈕
    let cancelBtn = document.getElementById('cancel-edit-btn');
    if (!cancelBtn) {
        cancelBtn = document.createElement('button');
        cancelBtn.id = 'cancel-edit-btn';
        cancelBtn.className = 'w-full py-3 px-4 rounded-lg font-bold btn-secondary mt-2';
        cancelBtn.textContent = t('BTN_CANCEL') || '取消編輯';
        submitBtn.parentElement.appendChild(cancelBtn);
        
        cancelBtn.addEventListener('click', () => {
            document.getElementById('announcement-title').value = '';
            document.getElementById('announcement-content').value = '';
            document.getElementById('announcement-priority').value = 'normal';
            submitBtn.textContent = t('BTN_SUBMIT_ANNOUNCEMENT') || '發布公告';
            submitBtn.dataset.mode = 'create';
            delete submitBtn.dataset.id;
            cancelBtn.remove();
        });
    }
}

/**
 * 更新公告
 */
async function updateAnnouncement(id) {
    const titleInput = document.getElementById('announcement-title');
    const contentInput = document.getElementById('announcement-content');
    const prioritySelect = document.getElementById('announcement-priority');
    const submitBtn = document.getElementById('submit-announcement-btn');
    
    const title = titleInput.value.trim();
    const content = contentInput.value.trim();
    const priority = prioritySelect.value;
    
    if (!title || !content) {
        showNotification(t('ANNOUNCEMENT_EMPTY_ERROR') || '標題和內容不可為空', 'error');
        return;
    }
    
    const loadingText = t('LOADING') || '處理中...';
    generalButtonState(submitBtn, 'processing', loadingText);
    
    try {
        const params = new URLSearchParams({
            id: id,
            title: title,
            content: content,
            priority: priority
        });
        
        const res = await callApifetch(`updateAnnouncement&${params.toString()}`);
        
        if (res.ok) {
            showNotification(t('ANNOUNCEMENT_UPDATED') || '公告已更新！', 'success');
            
            // 清空表單並恢復按鈕
            titleInput.value = '';
            contentInput.value = '';
            prioritySelect.value = 'normal';
            submitBtn.textContent = t('BTN_SUBMIT_ANNOUNCEMENT') || '發布公告';
            submitBtn.dataset.mode = 'create';
            delete submitBtn.dataset.id;
            
            const cancelBtn = document.getElementById('cancel-edit-btn');
            if (cancelBtn) cancelBtn.remove();
            
            // 重新載入公告列表
            await loadAllAnnouncements();
            
        } else {
            showNotification(t(res.code) || '更新失敗', 'error');
        }
        
    } catch (error) {
        console.error('❌ 更新公告失敗:', error);
        showNotification(t('NETWORK_ERROR') || '網路錯誤', 'error');
        
    } finally {
        generalButtonState(submitBtn, 'idle');
    }
}

/**
 * 刪除公告
 */
async function deleteAnnouncement(id) {
    if (!confirm(t('CONFIRM_DELETE_ANNOUNCEMENT') || '確定要刪除這則公告嗎？')) {
        return;
    }
    
    try {
        const res = await callApifetch(`deleteAnnouncement&id=${id}`);
        
        if (res.ok) {
            showNotification(t('ANNOUNCEMENT_DELETED') || '公告已刪除', 'success');
            await loadAllAnnouncements();
        } else {
            showNotification(t(res.code) || '刪除失敗', 'error');
        }
        
    } catch (error) {
        console.error('❌ 刪除公告失敗:', error);
        showNotification(t('NETWORK_ERROR') || '網路錯誤', 'error');
    }
}

/**
 * 綁定公告操作按鈕事件
 */
function bindAnnouncementActions() {
    // 編輯按鈕
    document.querySelectorAll('.edit-announcement-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const id = btn.dataset.id;
            const title = btn.dataset.title;
            const content = btn.dataset.content;
            const priority = btn.dataset.priority;
            editAnnouncement(id, title, content, priority);
        });
    });
    
    // 刪除按鈕
    document.querySelectorAll('.delete-announcement-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            deleteAnnouncement(btn.dataset.id);
        });
    });
}

/**
 * 初始化佈告欄功能
 */
function initAnnouncementModule() {
    const submitBtn = document.getElementById('submit-announcement-btn');
    
    if (submitBtn) {
        submitBtn.addEventListener('click', () => {
            const mode = submitBtn.dataset.mode || 'create';
            
            if (mode === 'update') {
                const id = submitBtn.dataset.id;
                updateAnnouncement(id);
            } else {
                createAnnouncement();
            }
        });
    }
}

// 頁面載入時初始化
document.addEventListener('DOMContentLoaded', () => {
    initAnnouncementModule();
});