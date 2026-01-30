/**
 * message-sender.js - Message Sending Logic
 * Handles sending messages to backend and managing input UI
 */

// Send message payload to server (VB backend)
function sendMessageToServer(messagePayload) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(messagePayload);
    } else if (window.vsto) {
        if (typeof window.vsto.sendMessage === 'function') {
            if (messagePayload.type === 'sendMessage' && typeof messagePayload.value === 'object') {
                window.vsto.sendMessage(JSON.stringify(messagePayload.value));
            } else {
                window.vsto.sendMessage(messagePayload);
            }
        } else if (typeof window.vsto.postMessage === 'function') {
            window.vsto.postMessage(messagePayload);
        }
    } else {
        alert('无法执行代码：未检测到支持的通信接口');
    }
}

// Send chat message
function sendChatMessage() {
    // 优先从smart-input获取内容，兼容隐藏的textarea
    const smartInput = document.getElementById('smart-input');
    const chatInput = document.getElementById('chat-input');
    
    // 从smart-input获取用户输入
    let userTypedText = '';
    if (smartInput && smartInput.innerText) {
        userTypedText = smartInput.innerText.trim();
    } else if (chatInput) {
        userTypedText = chatInput.value.trim();
    }
    
    // 检测 /loop 命令 - 启动 Ralph Loop
    if (userTypedText.startsWith('/loop ')) {
        const loopGoal = userTypedText.substring(6).trim();
        if (loopGoal) {
            // 清空输入
            if (smartInput) smartInput.innerText = '';
            if (chatInput) chatInput.value = '';
            
            // 发送启动循环消息
            sendMessageToServer({
                type: 'startLoop',
                goal: loopGoal
            });
            console.log('[RalphLoop] 启动循环，目标:', loopGoal);
            return;
        }
    }
    
    const attachedFileObjects = window.attachedFiles;
    const selectedSheetContent = window.getAllSelectedContent();

    // 检查是否处于续写模式
    if (window.continuationModeActive) {
        // 续写模式：发送续写请求而不是普通聊天
        sendContinuationMessage(userTypedText);
        
        // 清空输入
        if (typeof clearSmartInput === 'function') {
            clearSmartInput();
        } else {
            if (chatInput) chatInput.value = '';
            if (smartInput) smartInput.innerText = '';
        }
        if (chatInput) chatInput.style.height = 'auto';
        return;
    }

    // Check if there's any content to send
    if (!userTypedText && attachedFileObjects.length === 0 && selectedSheetContent.length === 0) return;

    // Toggle button display
    const sendButton = document.getElementById('send-button');
    const stopButton = document.getElementById('stop-button');

    sendButton.style.setProperty('display', 'none', 'important');
    stopButton.style.setProperty('display', 'flex', 'important');
    
    // 显示右上角等待动画
    showLoadingIndicator();

    // Prepare message payload
    const messagePayloadValue = {
        text: userTypedText,
        filePaths: attachedFileObjects.map(file => (file && typeof file.path === 'string' && file.path) ? file.path : file.name),
        selectedContent: selectedSheetContent
    };

    // 如果处于模板渲染模式，自动注入模板上下文
    if (window.templateModeActive && window.currentTemplateContext) {
        messagePayloadValue.responseMode = 'template_render';
        messagePayloadValue.templateContext = window.currentTemplateContext;
        messagePayloadValue.templateName = window.currentTemplateName || '';
    }

    sendMessageToServer({
        type: 'sendMessage',
        value: messagePayloadValue
    });

    const uuid = generateUUID();
    const now = new Date();
    const timestamp = formatDateTime(now);

    // Create chat section
    createChatSection('Me', timestamp, uuid);

    // Get message content div
    const messageContentDiv = document.getElementById('content-' + uuid);
    if (!messageContentDiv) {
        console.error('Could not find message content div for ' + uuid);
        return;
    }

    // Build message content HTML
    let htmlContent = '';

    // Add user typed text (parsed as markdown)
    if (userTypedText) {
        htmlContent += marked.parse(userTypedText);
    }

    // Add collapsible selected content reference
    if (selectedSheetContent.length > 0) {
        let itemsHtml = selectedSheetContent.map(item => `<div>${item.sheetName}: ${item.address}</div>`).join('');
        htmlContent += `
            <div class="chat-message-references collapsed" id="msg-ref-sel-${uuid}">
                <div class="chat-message-reference-header" onclick="toggleChatMessageReference(this)">
                    <span class="chat-message-reference-arrow">&#9658;</span>
                    <span class="chat-message-reference-label">引用内容 (${selectedSheetContent.length})</span>
                </div>
                <div class="chat-message-reference-content">
                    ${itemsHtml}
                </div>
            </div>`;
    }

    // Add collapsible file reference
    if (attachedFileObjects.length > 0) {
        let displayItemsHtml = attachedFileObjects.map(file => `<div>${escapeHtml(file.name)}</div>`).join('');
        htmlContent += `
            <div class="chat-message-references collapsed" id="msg-ref-file-${uuid}">
                <div class="chat-message-reference-header" onclick="toggleChatMessageReference(this)">
                    <span class="chat-message-reference-arrow">&#9658;</span>
                    <span class="chat-message-reference-label">引用文件 (${attachedFileObjects.length})</span>
                </div>
                <div class="chat-message-reference-content">
                    ${displayItemsHtml}
                </div>
            </div>`;
    }

    messageContentDiv.innerHTML = htmlContent;

    // Apply syntax highlighting to code blocks
    messageContentDiv.querySelectorAll('pre code').forEach((block) => {
        hljs.highlightElement(block);
    });

    // Clear input area references
    window.selectedContentMap = {};
    window.attachedFiles = [];
    renderReferences();

    // 清空输入框（优先使用smart-input）
    if (typeof clearSmartInput === 'function') {
        clearSmartInput();
    } else {
        chatInput.value = '';
        if (smartInput) smartInput.innerText = '';
    }
    chatInput.style.height = 'auto';
    hidePromptSuggestions();
}

// Stop button click handler
function stopButton() {
    sendMessageToServer({
        type: 'stopMessage'
    });
    // 隐藏等待动画
    if (typeof hideLoadingIndicator === 'function') {
        hideLoadingIndicator();
    }
}

// Change send button state
function changeSendButton() {
    const sendButton = document.getElementById('send-button');
    const stopButton = document.getElementById('stop-button');

    sendButton.style.setProperty('display', 'flex', 'important');
    stopButton.style.setProperty('display', 'none', 'important');
    
    // 隐藏等待动画
    if (typeof hideLoadingIndicator === 'function') {
        hideLoadingIndicator();
    }
}

// Initialize input event handlers
(function initMessageSender() {
    const chatInput = document.getElementById('chat-input');
    const smartInput = document.getElementById('smart-input');
    
    // Send button click
    document.getElementById('send-button').onclick = sendChatMessage;

    // 如果有smart-input，键盘事件由autocomplete.js处理
    // 否则使用传统textarea的事件处理
    if (!smartInput) {
        // Enter to send, Shift+Enter for newline (仅当没有smart-input时)
        chatInput.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                sendChatMessage();
            }
        });

        // Auto-resize textarea and prompt suggestions
        chatInput.addEventListener('input', function () {
            this.style.height = 'auto';
            this.style.height = (this.scrollHeight) + 'px';

            const value = this.value;
            if (value === '#') {
                showPromptSuggestions();
            } else if (!value.startsWith('#') || value.length > 1) {
                hidePromptSuggestions();
            }
        });
    } else {
        // smart-input的#提示词功能
        smartInput.addEventListener('input', function () {
            const value = this.innerText || '';
            if (value === '#') {
                showPromptSuggestions();
            } else if (!value.startsWith('#') || value.length > 1) {
                hidePromptSuggestions();
            }
        });
    }

    // Hide suggestions when clicking outside
    document.addEventListener('click', function (event) {
        const promptSuggestionsDiv = document.getElementById('prompt-suggestions');
        const attachFileButton = document.getElementById('attach-file-button');
        const targetInput = smartInput || chatInput;
        if (!targetInput.contains(event.target) && !promptSuggestionsDiv.contains(event.target) && !attachFileButton.contains(event.target)) {
            if (!event.target.closest('.reference-chip-remove')) {
                hidePromptSuggestions();
            }
        }
    });
})();

// Show prompt suggestions
function showPromptSuggestions() {
    const promptSuggestionsDiv = document.getElementById('prompt-suggestions');
    const chatInput = document.getElementById('chat-input');
    const smartInput = document.getElementById('smart-input');
    
    promptSuggestionsDiv.innerHTML = '';
    predefinedPrompts.forEach(promptText => {
        const item = document.createElement('div');
        item.className = 'prompt-suggestion-item';
        item.textContent = promptText;
        item.onclick = function () {
            // 优先更新smart-input
            if (smartInput) {
                smartInput.innerText = promptText;
                if (typeof syncToHiddenTextarea === 'function') {
                    syncToHiddenTextarea();
                }
            } else {
                chatInput.value = promptText;
            }
            hidePromptSuggestions();
            (smartInput || chatInput).focus();
            const event = new Event('input', { bubbles: true, cancelable: true });
            (smartInput || chatInput).dispatchEvent(event);
        };
        promptSuggestionsDiv.appendChild(item);
    });
    promptSuggestionsDiv.style.display = 'block';
}

// Hide prompt suggestions
function hidePromptSuggestions() {
    const promptSuggestionsDiv = document.getElementById('prompt-suggestions');
    promptSuggestionsDiv.style.display = 'none';
}

// Selected content management
window.addSelectedContentItem = function (sheetName, address, ctrlKey) {
    if (!address || address.trim() === '') {
        return;
    }
    const newItemId = generateUUID();
    const newItem = { id: newItemId, address: address.trim() };

    window.selectedContentMap[sheetName] = newItem;
    renderReferences();
};

window.clearSelectedContentBySheetName = function (sheetName) {
    if (window.selectedContentMap && window.selectedContentMap.hasOwnProperty(sheetName)) {
        delete window.selectedContentMap[sheetName];
        renderReferences();
    }
};

window.removeSelectedContentItem = function (itemIdToRemove) {
    for (const sheetName in window.selectedContentMap) {
        if (window.selectedContentMap.hasOwnProperty(sheetName)) {
            if (window.selectedContentMap[sheetName] && window.selectedContentMap[sheetName].id === itemIdToRemove) {
                delete window.selectedContentMap[sheetName];
                break;
            }
        }
    }
    renderReferences();
};

window.getAllSelectedContent = function () {
    const arr = [];
    for (const sheetName in window.selectedContentMap) {
        if (window.selectedContentMap.hasOwnProperty(sheetName)) {
            const selectedItem = window.selectedContentMap[sheetName];
            if (selectedItem) {
                arr.push({ sheetName: sheetName, address: selectedItem.address, id: selectedItem.id });
            }
        }
    }
    return arr;
};

// Render unified references display
function renderReferences() {
    const referencesWrapper = document.getElementById('references-wrapper');
    const referenceChipsList = document.getElementById('reference-chips-list');
    const referencesTitle = document.getElementById('references-title');
    
    if (!referencesWrapper || !referenceChipsList || !referencesTitle) {
        console.error("Reference display elements not found!");
        return;
    }

    referenceChipsList.innerHTML = '';
    let hasAnyReferences = false;

    // Render selected sheet content
    for (const sheetName in window.selectedContentMap) {
        if (window.selectedContentMap.hasOwnProperty(sheetName)) {
            const selectedItem = window.selectedContentMap[sheetName];
            if (!selectedItem) continue;
            hasAnyReferences = true;

            const itemChip = document.createElement('div');
            itemChip.className = 'reference-chip';
            itemChip.title = `${sheetName} [${selectedItem.address}]`;

            const chipContentWrapper = document.createElement('div');
            chipContentWrapper.className = 'reference-chip-content-wrapper';

            const itemNameSpan = document.createElement('span');
            itemNameSpan.className = 'reference-chip-name';
            itemNameSpan.textContent = `${sheetName}: ${selectedItem.address}`;
            chipContentWrapper.appendChild(itemNameSpan);

            const removeBtn = document.createElement('button');
            removeBtn.className = 'reference-chip-remove';
            removeBtn.title = '移除此引用';
            removeBtn.innerHTML = `<svg viewBox="0 0 20 20"><line x1="5" y1="5" x2="15" y2="15" stroke="currentColor" stroke-width="2"/><line x1="15" y1="5" x2="5" y2="15" stroke="currentColor" stroke-width="2"/></svg>`;
            removeBtn.onclick = function () {
                removeSelectedContentItem(selectedItem.id);
            };
            chipContentWrapper.appendChild(removeBtn);
            itemChip.appendChild(chipContentWrapper);
            referenceChipsList.appendChild(itemChip);
        }
    }

    // Render attached files
    window.attachedFiles.forEach((file, index) => {
        hasAnyReferences = true;
        const itemChip = document.createElement('div');
        itemChip.className = 'reference-chip';
        itemChip.title = file.name;

        const chipContentWrapper = document.createElement('div');
        chipContentWrapper.className = 'reference-chip-content-wrapper';

        const fileNameSpan = document.createElement('span');
        fileNameSpan.className = 'reference-chip-name';
        fileNameSpan.textContent = file.name;
        chipContentWrapper.appendChild(fileNameSpan);

        const removeBtn = document.createElement('button');
        removeBtn.className = 'reference-chip-remove';
        removeBtn.title = '移除此文件';
        removeBtn.innerHTML = `<svg viewBox="0 0 20 20"><line x1="5" y1="5" x2="15" y2="15" stroke="currentColor" stroke-width="2"/><line x1="15" y1="5" x2="5" y2="15" stroke="currentColor" stroke-width="2"/></svg>`;
        removeBtn.onclick = function () {
            window.attachedFiles.splice(index, 1);
            renderReferences();
        };
        chipContentWrapper.appendChild(removeBtn);
        itemChip.appendChild(chipContentWrapper);
        referenceChipsList.appendChild(itemChip);
    });

    // Control visibility
    referencesWrapper.style.display = hasAnyReferences ? 'block' : 'none';
}

// File attachment logic
(function initFileAttachment() {
    const attachFileButton = document.getElementById('attach-file-button');
    const fileInput = document.getElementById('file-input');

    // 点击附件按钮时，优先使用VB.NET对话框（支持完整路径）
    attachFileButton.addEventListener('click', () => {
        openFileDialogFromVB();
    });

    // 保留原有的文件输入处理（作为后备）
    fileInput.addEventListener('change', function (event) {
        const files = event.target.files;
        if (!files) return;
        const allowedExtensions = /(\.xls|\.xlsx|\.xlsm|\.xlsb|\.csv|\.doc|\.docx|\.ppt|\.pptx)$/i;
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (!allowedExtensions.exec(file.name)) {
                alert(`文件类型不支持: ${file.name}`);
                continue;
            }
            const isDuplicate = window.attachedFiles.some(
                existingFile => existingFile.name === file.name && existingFile.size === file.size
            );
            if (isDuplicate) {
                console.log(`文件已添加: ${file.name}`);
                continue;
            }
            window.attachedFiles.push({
                name: file.name,
                path: file.path || file.name,
                size: file.size
            });
        }
        renderReferences();
        fileInput.value = '';
    });
})();

// ========== 文件引用增强功能 ==========

/**
 * 打开文件选择对话框（调用VB.NET）
 */
function openFileDialogFromVB() {
    try {
        sendMessageToServer({ type: 'openFileDialog' });
    } catch (err) {
        console.error('openFileDialogFromVB error:', err);
        // 回退到HTML文件选择
        const fileInput = document.getElementById('file-input');
        if (fileInput) {
            fileInput.value = '';
            fileInput.click();
        }
    }
}

/**
 * 接收VB.NET返回的文件列表
 * @param {Array} files - 文件数组 [{name, path}, ...]
 */
function addFilesFromDialog(files) {
    try {
        if (!files || !Array.isArray(files)) return;

        const allowedExtensions = /(\.xls|\.xlsx|\.xlsm|\.xlsb|\.csv|\.doc|\.docx|\.ppt|\.pptx)$/i;

        files.forEach(file => {
            if (!file || !file.name) return;

            // 检查文件类型
            if (!allowedExtensions.exec(file.name)) {
                console.log(`文件类型不支持: ${file.name}`);
                return;
            }

            // 检查重复
            const isDuplicate = window.attachedFiles.some(
                existingFile => existingFile.name === file.name && existingFile.path === file.path
            );
            if (isDuplicate) {
                console.log(`文件已添加: ${file.name}`);
                return;
            }

            // 添加文件
            window.attachedFiles.push({
                name: file.name,
                path: file.path,
                fromDialog: true
            });
        });

        renderReferences();
        console.log(`通过对话框添加了 ${files.length} 个文件`);
    } catch (err) {
        console.error('addFilesFromDialog error:', err);
    }
}

/**
 * 初始化拖拽区域
 */
function initDragDrop() {
    const chatContainer = document.getElementById('chat-container');
    if (!chatContainer) return;

    // 拖拽进入
    chatContainer.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        chatContainer.classList.add('drag-over');
    });

    // 拖拽离开
    chatContainer.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        // 只有当离开的是容器本身时才移除样式
        if (e.target === chatContainer) {
            chatContainer.classList.remove('drag-over');
        }
    });

    // 放下文件
    chatContainer.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        chatContainer.classList.remove('drag-over');

        const files = e.dataTransfer.files;
        if (!files || files.length === 0) return;

        const allowedExtensions = /(\.xls|\.xlsx|\.xlsm|\.xlsb|\.csv|\.doc|\.docx|\.ppt|\.pptx)$/i;

        for (let i = 0; i < files.length; i++) {
            const file = files[i];

            // 检查文件类型
            if (!allowedExtensions.exec(file.name)) {
                console.log(`文件类型不支持: ${file.name}`);
                continue;
            }

            // 检查重复
            const isDuplicate = window.attachedFiles.some(
                existingFile => existingFile.name === file.name
            );
            if (isDuplicate) {
                console.log(`文件已添加: ${file.name}`);
                continue;
            }

            // 添加文件（拖拽的文件在WebView2环境下可能有path属性）
            window.attachedFiles.push({
                name: file.name,
                path: file.path || file.name,
                size: file.size,
                fromDrag: true
            });
        }

        renderReferences();
        console.log(`通过拖拽添加了 ${files.length} 个文件`);
    });

    console.log('拖拽功能已初始化');
}

// 在页面加载后初始化拖拽功能
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initDragDrop);
} else {
    initDragDrop();
}

// 导出函数供全局使用
window.openFileDialogFromVB = openFileDialogFromVB;
window.addFilesFromDialog = addFilesFromDialog;
window.initDragDrop = initDragDrop;

// ========== 意图识别显示功能 ==========

/**
 * 显示检测到的意图
 * @param {string} intentType - 意图类型
 */
function showDetectedIntent(intentType) {
    try {
        // 意图类型到中文标签的映射
        const intentLabels = {
            'DATA_ANALYSIS': '数据分析',
            'FORMULA_CALC': '公式计算',
            'CHART_GEN': '图表生成',
            'DATA_CLEANING': '数据清洗',
            'REPORT_GEN': '报表生成',
            'DATA_TRANSFORMATION': '数据转换',
            'FORMAT_STYLE': '格式调整',
            'GENERAL_QUERY': '通用查询'
        };

        // 意图类型到颜色的映射
        const intentColors = {
            'DATA_ANALYSIS': '#4a6fa5',
            'FORMULA_CALC': '#28a745',
            'CHART_GEN': '#ffc107',
            'DATA_CLEANING': '#17a2b8',
            'REPORT_GEN': '#6f42c1',
            'DATA_TRANSFORMATION': '#fd7e14',
            'FORMAT_STYLE': '#e83e8c',
            'GENERAL_QUERY': '#6c757d'
        };

        const label = intentLabels[intentType] || intentType;
        const color = intentColors[intentType] || '#6c757d';

        // 创建或获取意图指示器
        let indicator = document.getElementById('intent-indicator');
        if (!indicator) {
            indicator = document.createElement('div');
            indicator.id = 'intent-indicator';
            indicator.style.cssText = `
                position: fixed;
                top: 10px;
                right: 10px;
                z-index: 1000;
                padding: 6px 12px;
                border-radius: 16px;
                font-size: 12px;
                font-weight: 500;
                color: white;
                box-shadow: 0 2px 8px rgba(0,0,0,0.15);
                opacity: 0;
                transform: translateY(-10px);
                transition: opacity 0.3s ease, transform 0.3s ease;
            `;
            document.body.appendChild(indicator);
        }

        // 设置内容和颜色
        indicator.textContent = '识别: ' + label;
        indicator.style.backgroundColor = color;

        // 显示动画
        setTimeout(() => {
            indicator.style.opacity = '1';
            indicator.style.transform = 'translateY(0)';
        }, 10);

        // 3秒后淡出
        setTimeout(() => {
            indicator.style.opacity = '0';
            indicator.style.transform = 'translateY(-10px)';
        }, 3000);

        console.log('显示意图: ' + label);
    } catch (err) {
        console.error('showDetectedIntent error:', err);
    }
}
