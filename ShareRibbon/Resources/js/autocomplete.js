/**
 * autocomplete.js - 智能输入框自动补全模块
 * 实现类似Cursor/Qoder的Tab键补全功能
 */

// ========== 状态管理 ==========
window.autocompleteState = {
    enabled: true,                  // 是否启用自动补全
    delayMs: 800,                   // 防抖延迟（毫秒）
    debounceTimer: null,            // 防抖定时器
    currentCompletions: [],         // 当前补全候选列表
    selectedIndex: 0,               // 当前选中的候选索引
    isDropdownVisible: false,       // 下拉列表是否可见
    lastInputText: '',              // 上次输入的文本
    pendingRequest: null,           // 待处理的请求（用于取消）
    contextSnapshot: null           // Office上下文快照
};

// ========== 初始化 ==========

/**
 * 初始化智能输入框
 */
function initSmartInput() {
    const smartInput = document.getElementById('smart-input');
    const chatInput = document.getElementById('chat-input');
    const ghostText = document.getElementById('ghost-text');
    
    if (!smartInput) {
        console.warn('smart-input element not found');
        return;
    }
    
    // 输入事件 - 防抖触发补全
    smartInput.addEventListener('input', handleSmartInputChange);
    
    // 键盘事件 - 处理Tab、方向键、Esc
    smartInput.addEventListener('keydown', handleSmartInputKeydown);
    
    // 失焦时隐藏补全
    smartInput.addEventListener('blur', function(e) {
        // 延迟隐藏，允许点击下拉项
        setTimeout(function() {
            if (!document.activeElement || document.activeElement.id !== 'smart-input') {
                hideAutocompleteDropdown();
            }
        }, 200);
    });
    
    // 聚焦时同步内容
    smartInput.addEventListener('focus', function() {
        // 确保隐藏的textarea与smart-input同步
        syncToHiddenTextarea();
    });
    
    // 粘贴事件 - 只保留纯文本
    smartInput.addEventListener('paste', function(e) {
        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text/plain');
        document.execCommand('insertText', false, text);
    });
    
    console.log('Smart input initialized');
}

/**
 * 处理输入变化（防抖）
 */
function handleSmartInputChange(e) {
    const smartInput = e.target;
    const text = getSmartInputText();
    
    // 同步到隐藏的textarea
    syncToHiddenTextarea();
    
    // 清除ghost text
    clearGhostText();
    
    // 隐藏下拉列表
    hideAutocompleteDropdown();
    
    // 清除之前的防抖定时器
    if (window.autocompleteState.debounceTimer) {
        clearTimeout(window.autocompleteState.debounceTimer);
    }
    
    // 检查是否启用自动补全
    if (!window.autocompleteState.enabled) {
        return;
    }
    
    // 空输入不触发补全
    if (!text || text.trim().length < 2) {
        return;
    }
    
    // 设置防抖定时器
    window.autocompleteState.debounceTimer = setTimeout(function() {
        requestCompletion(text);
    }, window.autocompleteState.delayMs);
}

/**
 * 处理键盘事件
 * 补全快捷键: Tab (在chat输入框) 或 Ctrl+. (通用)
 */
function handleSmartInputKeydown(e) {
    const key = e.key;
    
    // Ctrl+. - 采纳补全 (主要快捷键，适用于Office原生编辑)
    if (key === '.' && e.ctrlKey) {
        if (window.autocompleteState.currentCompletions.length > 0) {
            e.preventDefault();
            acceptCompletion(window.autocompleteState.selectedIndex);
            return;
        }
    }
    
    // Tab键 - 采纳补全 (在chat输入框中可用)
    if (key === 'Tab') {
        if (window.autocompleteState.currentCompletions.length > 0) {
            e.preventDefault();
            acceptCompletion(window.autocompleteState.selectedIndex);
            return;
        }
    }
    
    // Escape键 - 关闭补全
    if (key === 'Escape') {
        if (window.autocompleteState.isDropdownVisible) {
            e.preventDefault();
            hideAutocompleteDropdown();
            clearGhostText();
            return;
        }
    }
    
    // 上下方向键 - 切换候选
    if (key === 'ArrowDown' || key === 'ArrowUp') {
        if (window.autocompleteState.isDropdownVisible && 
            window.autocompleteState.currentCompletions.length > 0) {
            e.preventDefault();
            
            const delta = key === 'ArrowDown' ? 1 : -1;
            const newIndex = window.autocompleteState.selectedIndex + delta;
            const maxIndex = window.autocompleteState.currentCompletions.length - 1;
            
            // 循环选择
            if (newIndex < 0) {
                window.autocompleteState.selectedIndex = maxIndex;
            } else if (newIndex > maxIndex) {
                window.autocompleteState.selectedIndex = 0;
            } else {
                window.autocompleteState.selectedIndex = newIndex;
            }
            
            updateDropdownSelection();
            updateGhostText();
            return;
        }
    }
    
    // Enter键 - 发送消息（如果没有补全显示）
    if (key === 'Enter' && !e.shiftKey && !e.ctrlKey) {
        // 如果下拉列表可见，先采纳补全
        if (window.autocompleteState.isDropdownVisible && 
            window.autocompleteState.currentCompletions.length > 0) {
            e.preventDefault();
            acceptCompletion(window.autocompleteState.selectedIndex);
            return;
        }
        
        // 否则发送消息
        e.preventDefault();
        syncToHiddenTextarea();
        sendChatMessage();
    }
    
    // Shift+Enter - 换行
    if (key === 'Enter' && e.shiftKey) {
        // 允许默认行为（换行）
    }
}

// ========== 补全请求 ==========

/**
 * 请求AI补全
 */
function requestCompletion(inputText) {
    // 取消之前的请求
    if (window.autocompleteState.pendingRequest) {
        window.autocompleteState.pendingRequest = null;
    }
    
    window.autocompleteState.lastInputText = inputText;
    
    // 构建请求数据
    const requestData = {
        type: 'requestCompletion',
        input: inputText,
        context: window.autocompleteState.contextSnapshot || {},
        timestamp: Date.now()
    };
    
    window.autocompleteState.pendingRequest = requestData.timestamp;
    
    // 发送请求到VB后端
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(requestData);
    } else if (window.vsto && typeof window.vsto.postMessage === 'function') {
        window.vsto.postMessage(requestData);
    } else {
        console.warn('No webview communication available for autocomplete');
    }
}

/**
 * 接收补全结果（由VB调用）
 */
function showCompletions(result) {
    // 检查是否是最新请求的响应
    if (result.timestamp && result.timestamp !== window.autocompleteState.pendingRequest) {
        return; // 忽略过期响应
    }
    
    // 检查输入是否已变化
    const currentText = getSmartInputText();
    if (currentText !== window.autocompleteState.lastInputText) {
        return; // 输入已变化，忽略响应
    }
    
    const completions = result.completions || [];
    
    if (completions.length === 0) {
        hideAutocompleteDropdown();
        clearGhostText();
        return;
    }
    
    // 保存补全列表
    window.autocompleteState.currentCompletions = completions;
    window.autocompleteState.selectedIndex = 0;
    
    // 显示ghost text
    updateGhostText();
    
    // 显示下拉列表
    showAutocompleteDropdown(completions);
}

// ========== UI 操作 ==========

/**
 * 获取smart-input的纯文本内容
 */
function getSmartInputText() {
    const smartInput = document.getElementById('smart-input');
    if (!smartInput) return '';
    return smartInput.innerText || smartInput.textContent || '';
}

/**
 * 设置smart-input的内容
 */
function setSmartInputText(text) {
    const smartInput = document.getElementById('smart-input');
    if (!smartInput) return;
    smartInput.innerText = text;
    
    // 移动光标到末尾
    moveCursorToEnd(smartInput);
    
    // 同步到隐藏textarea
    syncToHiddenTextarea();
}

/**
 * 移动光标到元素末尾
 */
function moveCursorToEnd(element) {
    const range = document.createRange();
    const selection = window.getSelection();
    range.selectNodeContents(element);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
}

/**
 * 同步smart-input内容到隐藏的textarea
 */
function syncToHiddenTextarea() {
    const smartInput = document.getElementById('smart-input');
    const chatInput = document.getElementById('chat-input');
    if (smartInput && chatInput) {
        chatInput.value = getSmartInputText();
    }
}

/**
 * 同步隐藏textarea内容到smart-input（用于外部设置）
 */
function syncFromHiddenTextarea() {
    const smartInput = document.getElementById('smart-input');
    const chatInput = document.getElementById('chat-input');
    if (smartInput && chatInput && chatInput.value) {
        smartInput.innerText = chatInput.value;
    }
}

/**
 * 清空输入框
 */
function clearSmartInput() {
    const smartInput = document.getElementById('smart-input');
    const chatInput = document.getElementById('chat-input');
    if (smartInput) smartInput.innerText = '';
    if (chatInput) chatInput.value = '';
    clearGhostText();
    hideAutocompleteDropdown();
}

// ========== Ghost Text ==========

/**
 * 更新ghost text显示
 */
function updateGhostText() {
    const ghostText = document.getElementById('ghost-text');
    if (!ghostText) return;
    
    const completions = window.autocompleteState.currentCompletions;
    const selectedIndex = window.autocompleteState.selectedIndex;
    
    if (completions.length === 0 || selectedIndex >= completions.length) {
        clearGhostText();
        return;
    }
    
    const completion = completions[selectedIndex];
    ghostText.textContent = completion;
    ghostText.style.display = 'inline';
}

/**
 * 清除ghost text
 */
function clearGhostText() {
    const ghostText = document.getElementById('ghost-text');
    if (ghostText) {
        ghostText.textContent = '';
        ghostText.style.display = 'none';
    }
}

// ========== Dropdown ==========

/**
 * 显示自动补全下拉列表
 */
function showAutocompleteDropdown(completions) {
    const dropdown = document.getElementById('autocomplete-dropdown');
    const list = document.getElementById('autocomplete-list');
    
    if (!dropdown || !list) return;
    
    // 清空列表
    list.innerHTML = '';
    
    // 添加候选项
    completions.forEach(function(completion, index) {
        const item = document.createElement('li');
        item.className = 'autocomplete-item' + (index === 0 ? ' selected' : '');
        item.setAttribute('data-index', index);
        
        // 显示补全文本
        item.innerHTML = '<span class="completion-text">' + escapeHtml(completion) + '</span>' +
                        '<span class="completion-hint">Tab</span>';
        
        // 点击选择
        item.addEventListener('click', function() {
            acceptCompletion(index);
        });
        
        // 悬停高亮
        item.addEventListener('mouseenter', function() {
            window.autocompleteState.selectedIndex = index;
            updateDropdownSelection();
            updateGhostText();
        });
        
        list.appendChild(item);
    });
    
    // 显示下拉列表
    dropdown.classList.remove('hidden');
    window.autocompleteState.isDropdownVisible = true;
}

/**
 * 隐藏自动补全下拉列表
 */
function hideAutocompleteDropdown() {
    const dropdown = document.getElementById('autocomplete-dropdown');
    if (dropdown) {
        dropdown.classList.add('hidden');
    }
    window.autocompleteState.isDropdownVisible = false;
    window.autocompleteState.currentCompletions = [];
    window.autocompleteState.selectedIndex = 0;
}

/**
 * 更新下拉列表选中状态
 */
function updateDropdownSelection() {
    const list = document.getElementById('autocomplete-list');
    if (!list) return;
    
    const items = list.querySelectorAll('.autocomplete-item');
    items.forEach(function(item, index) {
        if (index === window.autocompleteState.selectedIndex) {
            item.classList.add('selected');
        } else {
            item.classList.remove('selected');
        }
    });
}

// ========== 采纳补全 ==========

/**
 * 采纳补全
 */
function acceptCompletion(index) {
    const completions = window.autocompleteState.currentCompletions;
    
    if (index >= completions.length) return;
    
    const completion = completions[index];
    const currentText = getSmartInputText();
    
    // 将补全文本追加到当前输入
    const newText = currentText + completion;
    setSmartInputText(newText);
    
    // 记录采纳历史（发送到VB）
    recordCompletionAcceptance(currentText, completion);
    
    // 隐藏补全UI
    hideAutocompleteDropdown();
    clearGhostText();
    
    // 聚焦输入框
    const smartInput = document.getElementById('smart-input');
    if (smartInput) {
        smartInput.focus();
    }
}

/**
 * 记录补全采纳历史
 */
function recordCompletionAcceptance(input, completion) {
    const recordData = {
        type: 'acceptCompletion',
        input: input,
        completion: completion,
        context: window.officeAppType || 'Unknown',
        timestamp: Date.now()
    };
    
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(recordData);
    } else if (window.vsto && typeof window.vsto.postMessage === 'function') {
        window.vsto.postMessage(recordData);
    }
}

// ========== 上下文更新 ==========

/**
 * 更新Office上下文快照（由VB调用）
 */
function updateContextSnapshot(snapshot) {
    window.autocompleteState.contextSnapshot = snapshot;
}

// ========== 设置 ==========

/**
 * 更新自动补全设置
 */
function updateAutocompleteSettings(settings) {
    if (settings.hasOwnProperty('enabled')) {
        window.autocompleteState.enabled = settings.enabled;
    }
    if (settings.hasOwnProperty('delayMs')) {
        window.autocompleteState.delayMs = settings.delayMs;
    }
}

/**
 * 获取自动补全是否启用
 */
function isAutocompleteEnabled() {
    return window.autocompleteState.enabled;
}

// ========== 工具函数 ==========

/**
 * HTML转义
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ========== 页面加载初始化 ==========
document.addEventListener('DOMContentLoaded', function() {
    // 延迟初始化，确保DOM完全加载
    setTimeout(initSmartInput, 100);
});
