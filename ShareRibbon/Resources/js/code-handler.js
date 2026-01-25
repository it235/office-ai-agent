/**
 * code-handler.js - Code Block Handling
 * Functions for copying, executing, and editing code blocks
 */

/**
 * 隐藏指定消息中代码块的编辑和执行按钮（校对/排版模式使用）
 * @param {string} uuid - 消息的UUID
 */
function hideCodeActionButtons(uuid) {
    const messageContainer = document.getElementById('content-' + uuid);
    if (!messageContainer) return;
    
    // 隐藏所有编辑和执行按钮，只保留复制按钮
    const editButtons = messageContainer.querySelectorAll('.edit-button');
    const executeButtons = messageContainer.querySelectorAll('.execute-button');
    
    editButtons.forEach(btn => btn.style.display = 'none');
    executeButtons.forEach(btn => btn.style.display = 'none');
}

// Copy code from code block
function copyCode(button) {
    const codeBlock = button.closest('.code-block');
    const codeElement = codeBlock.querySelector('code');
    const code = codeElement.textContent;

    // Create temp textarea for copying
    const textarea = document.createElement('textarea');
    textarea.value = code;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);

    try {
        textarea.select();
        textarea.setSelectionRange(0, 99999);
        document.execCommand('copy');

        const originalText = button.innerHTML;
        button.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="20 6 9 17 4 12"></polyline>
            </svg>
            已复制
        `;
        setTimeout(() => {
            button.innerHTML = originalText;
        }, 2000);
    } catch (err) {
        console.error('copy failure:', err);
        alert('copy failure');
    } finally {
        document.body.removeChild(textarea);
    }
}

// Execute code from code block
function executeCode(button) {
    const codeBlock = button.closest('.code-block');
    const codeElement = codeBlock.querySelector('code');
    const code = codeElement.textContent;
    const language = codeElement.className.replace('language-', '');
    let preview = document.getElementById('settings-executecode-preview').checked;

    try {
        // Find parent chat container for UUID mapping
        const chatContainer = button.closest('.chat-container');
        let responseUuid = null;
        let requestUuid = null;
        if (chatContainer && chatContainer.id && chatContainer.id.startsWith('chat-')) {
            responseUuid = chatContainer.id.replace('chat-', '');
            requestUuid = chatContainer.dataset ? chatContainer.dataset.requestId : null;
        }

        const payload = {
            type: 'executeCode',
            code: code,
            language: language,
            executecodePreview: preview,
            responseUuid: responseUuid,
            requestUuid: requestUuid
        };

        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage(payload);
        } else if (window.vsto) {
            window.vsto.executeCode(code, language, preview);
        } else {
            alert('无法执行代码：未检测到支持的通信接口');
        }

        // UI feedback
        const originalText = button.innerHTML;
        button.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polygon points="5 3 19 12 5 21 5 3"></polygon>
            </svg>
            已执行
        `;
        setTimeout(() => {
            button.innerHTML = originalText;
        }, 2000);
    } catch (err) {
        alert('执行失败：' + err.message);
    }
}

// Edit code in code block
function editCode(button) {
    const codeBlock = button.closest('.code-block');
    const codeElement = codeBlock.querySelector('code');
    const code = codeElement.textContent;
    const language = codeElement.className.replace('language-', '');

    // Create editor container
    const editorContainer = document.createElement('div');
    editorContainer.className = 'editor-container';

    const textarea = document.createElement('textarea');
    textarea.className = 'code-editor';
    textarea.value = code;

    const buttonsDiv = document.createElement('div');
    buttonsDiv.className = 'editor-buttons';

    const saveButton = document.createElement('button');
    saveButton.className = 'code-button';
    saveButton.innerHTML = '保存';
    saveButton.onclick = function () {
        const newCode = textarea.value;
        const newCodeHtml = marked.parse('```' + language + '\n' + newCode + '\n```');

        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = newCodeHtml;
        const newCodeBlock = tempDiv.querySelector('.code-block');

        codeBlock.parentNode.replaceChild(newCodeBlock, codeBlock);

        // Re-apply syntax highlighting
        document.querySelectorAll('pre code').forEach((block) => {
            hljs.highlightElement(block);
        });

        editorContainer.remove();
    };

    const cancelButton = document.createElement('button');
    cancelButton.className = 'code-button';
    cancelButton.style.backgroundColor = '#f44336';
    cancelButton.innerHTML = '取消';
    cancelButton.onclick = function () {
        codeBlock.style.display = 'block';
        editorContainer.remove();
    };

    buttonsDiv.appendChild(cancelButton);
    buttonsDiv.appendChild(saveButton);

    editorContainer.appendChild(textarea);
    editorContainer.appendChild(buttonsDiv);

    // Hide original code block, insert editor
    codeBlock.style.display = 'none';
    codeBlock.parentNode.insertBefore(editorContainer, codeBlock);

    textarea.focus();
    editorContainer.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// Process stream complete - handle UI updates after message completion
function processStreamComplete(uuid, totalTokens) {
    // Add token display
    const footerDiv = document.getElementById('footer-' + uuid);
    if (footerDiv) {
        footerDiv.innerHTML = `<span class="token-count">消耗token：${totalTokens}</span>`;
    }

    // Switch back to send button
    const sendButton = document.getElementById('send-button');
    const stopButton = document.getElementById('stop-button');

    sendButton.style.setProperty('display', 'flex', 'important');
    stopButton.style.setProperty('display', 'none', 'important');

    // Collapse code blocks
    const contentDiv = document.getElementById('content-' + uuid);
    if (contentDiv) {
        const codeBlocks = contentDiv.querySelectorAll('pre code');
        codeBlocks.forEach(codeBlock => {
            const preElement = codeBlock.parentElement;
            if (preElement) {
                if (!preElement.classList.contains('collapsible')) {
                    preElement.classList.add('collapsible', 'collapsed');

                    const toggleLabel = document.createElement('div');
                    toggleLabel.className = 'code-toggle-label';
                    toggleLabel.innerHTML = '点击展开代码';
                    toggleLabel.onclick = function (e) {
                        e.stopPropagation();
                        preElement.classList.toggle('collapsed');
                        toggleLabel.innerHTML = preElement.classList.contains('collapsed') ? '点击展开代码' : '点击折叠代码';
                    };

                    preElement.parentNode.insertBefore(toggleLabel, preElement);
                }
            }
        });
    }

    // Auto-execute in agent mode
    if (document.getElementById("chatMode").value === 'agent') {
        let executeBtns = document.getElementById("content-" + uuid).querySelector(".execute-button");
        if (executeBtns) {
            executeBtns.click();
        }
    }

    // Render accept/reject buttons for AI messages
    try {
        renderAcceptRejectButtons(uuid);
    } catch (err) {
        console.error('renderAcceptRejectButtons error:', err);
    }
}

// Render accept/reject buttons (only for AI messages)
function renderAcceptRejectButtons(uuid) {
    try {
        const chatDiv = document.getElementById('chat-' + uuid);
        if (!chatDiv) return;
        
        const sender = chatDiv.dataset && chatDiv.dataset.sender ? chatDiv.dataset.sender : (chatDiv.querySelector('.sender-name') ? chatDiv.querySelector('.sender-name').textContent : '');

        // Only show buttons for AI messages
        if (!sender || sender === 'Me') return;

        const footer = document.getElementById('footer-' + uuid);
        if (!footer) return;

        // Skip if buttons already exist
        if (footer.querySelector('.accept-btn') || footer.querySelector('.reject-btn')) return;

        const btnAccept = document.createElement('button');
        btnAccept.className = 'code-button accept-btn';
        btnAccept.style.backgroundColor = '#4CAF50';
        btnAccept.style.marginRight = '8px';
        btnAccept.textContent = '接受该答案';
        btnAccept.onclick = function () { acceptAnswer(uuid); };

        const btnReject = document.createElement('button');
        btnReject.className = 'code-button reject-btn';
        btnReject.style.backgroundColor = '#E9525F';
        btnReject.textContent = '不接受，继续改进';
        btnReject.onclick = function () { rejectAnswer(uuid); };

        footer.insertBefore(btnReject, footer.firstChild);
    } catch (err) {
        console.error('renderAcceptRejectButtons error:', err);
    }
}

// Accept answer handler
function acceptAnswer(uuid) {
    try {
        const contentDiv = document.getElementById('content-' + uuid);
        const plainText = contentDiv ? (contentDiv.innerText || contentDiv.textContent || '') : '';

        sendMessageToServer({
            type: 'acceptAnswer',
            uuid: uuid,
            content: plainText
        });

        const footer = document.getElementById('footer-' + uuid);
        if (footer) {
            footer.querySelectorAll('.accept-btn, .reject-btn').forEach(b => b.disabled = true);
            const statusSpan = document.createElement('span');
            statusSpan.className = 'token-count';
            statusSpan.textContent = '已接受';
            footer.appendChild(statusSpan);
        }
    } catch (err) {
        console.error('acceptAnswer error:', err);
    }
}

// Reject answer handler
function rejectAnswer(uuid) {
    try {
        const contentDiv = document.getElementById('content-' + uuid);
        const plainText = contentDiv ? (contentDiv.innerText || contentDiv.textContent || '') : '';

        let reason = '';
        try {
            reason = prompt('请简要说明希望如何改进（可留空）：', '');
            if (reason === null) {
                return;
            }
        } catch (e) {
            reason = '';
        }

        sendMessageToServer({
            type: 'rejectAnswer',
            uuid: uuid,
            content: plainText,
            reason: reason
        });

        const footer = document.getElementById('footer-' + uuid);
        if (footer) {
            footer.querySelectorAll('.accept-btn, .reject-btn').forEach(b => b.disabled = true);
            const statusSpan = document.createElement('span');
            statusSpan.className = 'token-count';
            statusSpan.textContent = '已请求改进，等待新结果…';
            footer.appendChild(statusSpan);
        }

        const reasoning = document.getElementById('reasoning-' + uuid);
        if (reasoning) {
            reasoning.classList.remove('collapsed');
        }
    } catch (err) {
        console.error('rejectAnswer error:', err);
    }
}

// Chat mode changed handler
function chatModeChanged(select) {
    settingsSave();
}

// Batch delete chat function
function showBatchDeleteChat() {
    // Show action buttons
    if (!document.getElementById('delete-chat-actions')) {
        const actionsDiv = document.createElement('div');
        actionsDiv.id = 'delete-chat-actions';
        actionsDiv.style = 'display:block; position:fixed; bottom:80px; left:50%; transform:translateX(-50%); z-index:999;';
        actionsDiv.innerHTML = `
            <button id="confirm-delete-chat" style="background:#e9525f;color:white;border:none;padding:6px 16px;border-radius:6px;margin-right:10px;">确定删除</button>
            <button id="cancel-delete-chat" style="background:#f5f5f5;color:#333;border:none;padding:6px 16px;border-radius:6px;">取消</button>
        `;
        document.body.appendChild(actionsDiv);
    } else {
        document.getElementById('delete-chat-actions').style.display = 'block';
    }
    
    // Insert checkboxes
    document.querySelectorAll('#chat-container .chat-container').forEach(function (chatDiv) {
        if (!chatDiv.querySelector('.chat-select-checkbox')) {
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.className = 'chat-select-checkbox';
            checkbox.style.marginRight = '8px';
            chatDiv.insertBefore(checkbox, chatDiv.firstChild);
        }
    });

    // Bind events (ensure only bound once)
    if (!window._deleteChatEventsBound) {
        document.getElementById('cancel-delete-chat').addEventListener('click', function () {
            document.getElementById('delete-chat-actions').style.display = 'none';
            document.querySelectorAll('#chat-container .chat-container .chat-select-checkbox').forEach(function (cb) {
                cb.parentNode.removeChild(cb);
            });
        });
        document.getElementById('confirm-delete-chat').addEventListener('click', function () {
            document.querySelectorAll('#chat-container .chat-container').forEach(function (chatDiv) {
                const cb = chatDiv.querySelector('.chat-select-checkbox');
                if (cb && cb.checked) {
                    chatDiv.parentNode.removeChild(chatDiv);
                }
            });
            document.getElementById('delete-chat-actions').style.display = 'none';
            document.querySelectorAll('#chat-container .chat-container .chat-select-checkbox').forEach(function (cb) {
                cb.parentNode.removeChild(cb);
            });
        });
        window._deleteChatEventsBound = true;
    }
}

// ========== AI续写功能 ==========

/**
 * 触发AI续写
 */
function triggerContinuation() {
    try {
        window.chrome.webview.postMessage({
            type: 'triggerContinuation'
        });
    } catch (err) {
        console.error('triggerContinuation error:', err);
    }
}

/**
 * 显示续写预览界面 - 在AI响应完成后调用
 * @param {string} uuid - 消息的唯一标识
 */
function showContinuationPreview(uuid) {
    try {
        const chatSection = document.getElementById('chat-' + uuid);
        if (!chatSection) {
            console.error('showContinuationPreview: 找不到 chat section, uuid=' + uuid);
            return;
        }

        // 使用正确的选择器：message-content 或通过 id
        const contentEl = document.getElementById('content-' + uuid) || chatSection.querySelector('.message-content');
        if (!contentEl) {
            console.error('showContinuationPreview: 找不到 content 元素, uuid=' + uuid);
            return;
        }

        // 检查是否已经有续写操作按钮
        if (document.getElementById('continuation-actions-' + uuid)) return;

        // 隐藏常规聊天的 reject-btn（如果存在）
        const footer = document.getElementById('footer-' + uuid);
        if (footer) {
            const rejectBtn = footer.querySelector('.reject-btn');
            if (rejectBtn) rejectBtn.style.display = 'none';
        }

        // 检测应用类型：PPT 或 Word/其他
        const isPPT = window.officeAppType === 'PowerPoint';
        
        // 根据应用类型设置按钮文案
        const insertStartLabel = isPPT ? '插入首页' : '插入开头';
        const insertCurrentLabel = isPPT ? '插入当前页' : '插入文档';
        const insertEndLabel = isPPT ? '插入末页' : '插入结尾';

        // 创建续写操作按钮区域
        const actionsHtml = `
            <div class="continuation-actions" id="continuation-actions-${uuid}" style="margin-top: 8px; padding: 8px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e9ecef;">
                <div style="margin-bottom: 6px; font-size: 12px; color: #666;">续写预览完成：</div>
                <div style="margin-bottom: 6px;">
                    <button class="btn-primary continuation-btn" onclick="handleContinuationInsert('${uuid}', 'start')" style="background: #6c757d; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; margin-right: 4px; font-size: 11px;">
                        ${insertStartLabel}
                    </button>
                    <button class="btn-primary continuation-btn" onclick="handleContinuationInsert('${uuid}', 'current')" style="background: #4a6fa5; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; margin-right: 4px; font-size: 11px;">
                        ${insertCurrentLabel}
                    </button>
                    <button class="btn-primary continuation-btn" onclick="handleContinuationInsert('${uuid}', 'end')" style="background: #6c757d; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        ${insertEndLabel}
                    </button>
                </div>
                <div>
                    <button class="btn-secondary continuation-btn" onclick="handleContinuationRefine('${uuid}')" style="background: #e9ecef; color: #333; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; margin-right: 4px; font-size: 11px;">
                        调整提示词
                    </button>
                    <button class="btn-secondary continuation-btn" onclick="handleContinuationRegenerate()" style="background: #e9ecef; color: #333; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        重新生成
                    </button>
                </div>
            </div>
        `;
        contentEl.insertAdjacentHTML('afterend', actionsHtml);
        console.log('showContinuationPreview: 续写操作按钮已添加, uuid=' + uuid);
        
        // 停止续写按钮的闪烁动画
        stopContinuationHint();
    } catch (err) {
        console.error('showContinuationPreview error:', err);
    }
}

/**
 * 处理续写内容插入
 * @param {string} uuid - 消息的唯一标识
 * @param {string} position - 插入位置：'start'/'current'/'end'
 */
function handleContinuationInsert(uuid, position) {
    try {
        position = position || 'current';
        
        const chatSection = document.getElementById('chat-' + uuid);
        if (!chatSection) return;

        // 使用正确的选择器
        const contentEl = document.getElementById('content-' + uuid) || chatSection.querySelector('.message-content');
        if (!contentEl) return;

        // 获取续写内容（纯文本）
        const content = contentEl.innerText || contentEl.textContent;

        // 发送插入请求到VB，包含位置参数
        window.chrome.webview.postMessage({
            type: 'applyContinuation',
            uuid: uuid,
            content: content,
            position: position
        });

        // 移除操作按钮并显示成功提示
        removeContinuationActions(uuid);
        
        // 添加成功提示
        const successMsg = document.createElement('div');
        successMsg.style = 'margin-top: 8px; padding: 8px 12px; background: #d4edda; color: #155724; border-radius: 6px; font-size: 13px;';
        successMsg.textContent = '续写内容已插入文档';
        contentEl.parentNode.appendChild(successMsg);
        
        // 3秒后移除提示
        setTimeout(() => successMsg.remove(), 3000);
    } catch (err) {
        console.error('handleContinuationInsert error:', err);
    }
}

/**
 * 处理续写方向调整
 * @param {string} uuid - 消息的唯一标识
 */
function handleContinuationRefine(uuid) {
    try {
        const refinement = prompt('请输入调整方向（如：更正式、更简洁、加长、更详细等）：');
        if (refinement && refinement.trim()) {
            window.chrome.webview.postMessage({
                type: 'refineContinuation',
                uuid: uuid,
                refinement: refinement.trim()
            });
            
            // 更新按钮状态
            const actionsDiv = document.getElementById('continuation-actions-' + uuid);
            if (actionsDiv) {
                actionsDiv.innerHTML = '<div style="color: #666; font-size: 13px;">正在根据您的要求调整内容...</div>';
            }
        }
    } catch (err) {
        console.error('handleContinuationRefine error:', err);
    }
}

/**
 * 处理重新生成续写
 */
function handleContinuationRegenerate() {
    try {
        window.chrome.webview.postMessage({
            type: 'triggerContinuation',
            regenerate: true
        });
    } catch (err) {
        console.error('handleContinuationRegenerate error:', err);
    }
}

/**
 * 移除续写操作按钮
 * @param {string} uuid - 消息的唯一标识
 */
function removeContinuationActions(uuid) {
    try {
        const actionsDiv = document.getElementById('continuation-actions-' + uuid);
        if (actionsDiv) {
            actionsDiv.remove();
        }
    } catch (err) {
        console.error('removeContinuationActions error:', err);
    }
}

// ========== 续写模式状态管理 ==========

// 续写模式状态
window.continuationModeActive = false;
window.continuationContext = null; // 保存续写上下文，用于多轮续写

/**
 * 进入续写模式
 */
function enterContinuationMode() {
    window.continuationModeActive = true;
    
    // 更新UI
    updateContinuationModeUI(true);
    
    console.log('已进入续写模式');
}

/**
 * 退出续写模式
 */
function exitContinuationMode() {
    window.continuationModeActive = false;
    window.continuationContext = null;
    
    // 恢复UI
    updateContinuationModeUI(false);
    
    console.log('已退出续写模式');
}

/**
 * 更新续写模式的UI状态
 * @param {boolean} isActive - 是否处于续写模式
 */
function updateContinuationModeUI(isActive) {
    const chatInput = document.getElementById('chat-input');
    const inputCard = document.getElementById('chat-input-card');
    const continuationBtn = document.getElementById('continuation-button');
    
    // 工具栏按钮（续写模式下隐藏）
    const mcpBtn = document.getElementById('mcp-toggle-btn');
    const clearBtn = document.getElementById('clear-context-btn');
    const historyBtn = document.getElementById('history-toggle-btn');
    
    if (isActive) {
        // 续写模式：更改placeholder和样式
        if (chatInput) {
            chatInput.placeholder = '在此输入续写要求（如：更正式、加长、换个角度等），或直接回车继续续写...';
        }
        if (inputCard) {
            inputCard.style.borderColor = '#4a6fa5';
            inputCard.style.boxShadow = '0 0 0 2px rgba(74, 111, 165, 0.2)';
        }
        if (continuationBtn) {
            continuationBtn.style.background = '#4a6fa5';
            continuationBtn.style.borderRadius = '4px';
            continuationBtn.querySelector('svg').style.stroke = 'white';
        }
        
        // 隐藏工具栏按钮
        if (mcpBtn) mcpBtn.style.display = 'none';
        if (clearBtn) clearBtn.style.display = 'none';
        if (historyBtn) historyBtn.style.display = 'none';
        
        // 显示续写模式指示器
        showContinuationModeIndicator();
    } else {
        // 普通模式：恢复默认
        if (chatInput) {
            chatInput.placeholder = '请在此输入您的问题... 按Enter键直接发送，Shift+Enter换行';
        }
        if (inputCard) {
            inputCard.style.borderColor = '';
            inputCard.style.boxShadow = '';
        }
        if (continuationBtn) {
            continuationBtn.style.background = '';
            continuationBtn.querySelector('svg').style.stroke = '';
        }
        
        // 显示工具栏按钮
        if (mcpBtn) mcpBtn.style.display = '';
        if (clearBtn) clearBtn.style.display = '';
        if (historyBtn) historyBtn.style.display = '';
        
        // 隐藏续写模式指示器
        hideContinuationModeIndicator();
    }
}

/**
 * 显示续写模式指示器（吸顶fixed）
 */
function showContinuationModeIndicator() {
    if (document.getElementById('continuation-mode-indicator')) return;
    
    const indicator = document.createElement('div');
    indicator.id = 'continuation-mode-indicator';
    indicator.innerHTML = `
        <div style="background: linear-gradient(135deg, #4a6fa5 0%, #3d5a7c 100%); color: white; 
                    padding: 8px 12px; font-size: 12px; display: flex; align-items: center; justify-content: space-between;
                    position: fixed; top: 0; left: 0; right: 0; z-index: 9999; box-shadow: 0 2px 8px rgba(0,0,0,0.15);">
            <span>📝 续写模式 - 输入框内容将作为续写要求发送</span>
            <button onclick="exitContinuationMode()" style="background: rgba(255,255,255,0.25); border: none; 
                    color: white; padding: 4px 12px; border-radius: 4px; cursor: pointer; font-size: 11px; font-weight: 500;">
                退出续写
            </button>
        </div>
    `;
    
    document.body.appendChild(indicator);
    
    // 给body添加顶部padding以防止内容被遮挡
    document.body.style.paddingTop = '36px';
}

/**
 * 隐藏续写模式指示器
 */
function hideContinuationModeIndicator() {
    const indicator = document.getElementById('continuation-mode-indicator');
    if (indicator) indicator.remove();
    
    // 恢复body的padding
    document.body.style.paddingTop = '';
}

/**
 * 在续写模式下发送消息（由message-sender.js调用）
 * @param {string} text - 用户输入的文本（作为续写要求/风格）
 */
function sendContinuationMessage(text) {
    if (!window.continuationModeActive) return false;
    
    // 发送续写请求，text作为风格/要求
    window.chrome.webview.postMessage({
        type: 'triggerContinuation',
        style: text || '',
        isContinuationMode: true
    });
    
    return true;
}

// ========== 续写按钮动画提示 ==========

let continuationHintInterval = null;

/**
 * 启动续写按钮的闪烁提示动画
 */
function startContinuationHint() {
    const btn = document.getElementById('continuation-button');
    if (!btn) return;
    
    // 添加闪烁动画样式
    btn.style.animation = 'continuation-hint-pulse 1s ease-in-out infinite';
    btn.style.boxShadow = '0 0 8px #4a6fa5';
    btn.title = '点击此处开始AI续写';
    
    // 显示提示气泡
    showContinuationTooltip();
}

/**
 * 停止续写按钮的闪烁提示
 */
function stopContinuationHint() {
    const btn = document.getElementById('continuation-button');
    if (!btn) return;
    
    btn.style.animation = '';
    btn.style.boxShadow = '';
    btn.title = 'AI续写';
    
    // 移除提示气泡
    hideContinuationTooltip();
}

/**
 * 显示续写提示气泡
 */
function showContinuationTooltip() {
    // 移除已有的提示
    hideContinuationTooltip();
    
    const btn = document.getElementById('continuation-button');
    if (!btn) return;
    
    const tooltip = document.createElement('div');
    tooltip.id = 'continuation-tooltip';
    tooltip.innerHTML = `
        <div style="position: absolute; bottom: 45px; left: 50%; transform: translateX(-50%); 
                    background: #4a6fa5; color: white; padding: 8px 12px; border-radius: 6px; 
                    font-size: 12px; white-space: nowrap; z-index: 1000; box-shadow: 0 2px 8px rgba(0,0,0,0.2);">
            点击开始AI续写，可输入风格要求
            <div style="position: absolute; bottom: -6px; left: 50%; transform: translateX(-50%); 
                        border-left: 6px solid transparent; border-right: 6px solid transparent; 
                        border-top: 6px solid #4a6fa5;"></div>
        </div>
    `;
    btn.style.position = 'relative';
    btn.appendChild(tooltip);
    
    // 5秒后自动隐藏
    setTimeout(hideContinuationTooltip, 5000);
}

/**
 * 隐藏续写提示气泡
 */
function hideContinuationTooltip() {
    const tooltip = document.getElementById('continuation-tooltip');
    if (tooltip) tooltip.remove();
}

/**
 * 显示续写风格输入对话框
 * @param {boolean} autoTrigger - 是否自动触发（从Ribbon点击）
 */
function showContinuationDialog(autoTrigger) {
    // 创建对话框
    const dialogHtml = `
        <div id="continuation-dialog-overlay" style="position: fixed; top: 0; left: 0; right: 0; bottom: 0; 
                background: rgba(0,0,0,0.4); z-index: 9998; display: flex; align-items: center; justify-content: center;">
            <div style="background: white; border-radius: 8px; padding: 16px; width: 280px; box-shadow: 0 4px 20px rgba(0,0,0,0.2);">
                <div style="font-size: 14px; font-weight: 500; margin-bottom: 12px; color: #333;">AI续写设置</div>
                <div style="font-size: 12px; color: #666; margin-bottom: 8px;">可选：输入续写风格要求</div>
                <input type="text" id="continuation-style-input" placeholder="如：更正式、更简洁、幽默风格..." 
                       style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 12px; box-sizing: border-box; margin-bottom: 12px;">
                <div style="display: flex; justify-content: flex-end; gap: 8px;">
                    <button onclick="closeContinuationDialog()" 
                            style="padding: 6px 12px; border: 1px solid #ddd; background: white; border-radius: 4px; cursor: pointer; font-size: 12px;">
                        取消
                    </button>
                    <button onclick="submitContinuation()" 
                            style="padding: 6px 12px; border: none; background: #4a6fa5; color: white; border-radius: 4px; cursor: pointer; font-size: 12px;">
                        开始续写
                    </button>
                </div>
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', dialogHtml);
    
    // 聚焦输入框
    setTimeout(() => {
        const input = document.getElementById('continuation-style-input');
        if (input) input.focus();
    }, 100);
    
    // 支持回车提交
    const input = document.getElementById('continuation-style-input');
    if (input) {
        input.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') submitContinuation();
        });
    }
}

/**
 * 关闭续写对话框
 */
function closeContinuationDialog() {
    const overlay = document.getElementById('continuation-dialog-overlay');
    if (overlay) overlay.remove();
    stopContinuationHint();
}

/**
 * 提交续写请求
 */
function submitContinuation() {
    const input = document.getElementById('continuation-style-input');
    const style = input ? input.value.trim() : '';
    
    closeContinuationDialog();
    
    // 进入续写模式
    enterContinuationMode();
    
    // 发送续写请求，带上风格参数
    window.chrome.webview.postMessage({
        type: 'triggerContinuation',
        style: style
    });
}

/**
 * 触发AI续写（支持从Ribbon自动触发）
 * @param {boolean} autoTrigger - 是否自动触发（从Ribbon点击过来）
 */
function triggerContinuation(autoTrigger) {
    try {
        if (window.continuationModeActive) {
            // 已在续写模式，直接续写（不弹框）
            window.chrome.webview.postMessage({
                type: 'triggerContinuation',
                style: '',
                isContinuationMode: true
            });
        } else if (autoTrigger) {
            // 从Ribbon触发，显示风格输入对话框
            showContinuationDialog(true);
        } else {
            // 从侧栏按钮触发，也显示对话框进入续写模式
            showContinuationDialog(false);
        }
    } catch (err) {
        console.error('triggerContinuation error:', err);
    }
}

// ========== 校对/排版模式吸顶提示 ==========

/**
 * 显示校对模式指示器（吸顶fixed）
 */
function showProofreadModeIndicator() {
    // 移除其他模式指示器
    hideAllModeIndicators();
    
    if (document.getElementById('proofread-mode-indicator')) return;
    
    const indicator = document.createElement('div');
    indicator.id = 'proofread-mode-indicator';
    indicator.innerHTML = `
        <div style="background: linear-gradient(135deg, #e67e22 0%, #d35400 100%); color: white; 
                    padding: 8px 12px; font-size: 12px; display: flex; align-items: center; justify-content: center;
                    position: fixed; top: 0; left: 0; right: 0; z-index: 9999; box-shadow: 0 2px 8px rgba(0,0,0,0.15);">
            <span>🔍 校对模式 - AI正在帮您检查语法、拼写和表达问题</span>
        </div>
    `;
    
    document.body.appendChild(indicator);
    document.body.style.paddingTop = '36px';
}

/**
 * 隐藏校对模式指示器
 */
function hideProofreadModeIndicator() {
    const indicator = document.getElementById('proofread-mode-indicator');
    if (indicator) {
        indicator.remove();
        document.body.style.paddingTop = '';
    }
}

/**
 * 显示排版模式指示器（吸顶fixed）
 */
function showReformatModeIndicator() {
    // 移除其他模式指示器
    hideAllModeIndicators();
    
    if (document.getElementById('reformat-mode-indicator')) return;
    
    const indicator = document.createElement('div');
    indicator.id = 'reformat-mode-indicator';
    indicator.innerHTML = `
        <div style="background: linear-gradient(135deg, #9b59b6 0%, #8e44ad 100%); color: white; 
                    padding: 8px 12px; font-size: 12px; display: flex; align-items: center; justify-content: center;
                    position: fixed; top: 0; left: 0; right: 0; z-index: 9999; box-shadow: 0 2px 8px rgba(0,0,0,0.15);">
            <span>📐 排版模式 - AI正在帮您优化文档结构和格式</span>
        </div>
    `;
    
    document.body.appendChild(indicator);
    document.body.style.paddingTop = '36px';
}

/**
 * 隐藏排版模式指示器
 */
function hideReformatModeIndicator() {
    const indicator = document.getElementById('reformat-mode-indicator');
    if (indicator) {
        indicator.remove();
        document.body.style.paddingTop = '';
    }
}

/**
 * 隐藏所有模式指示器
 */
function hideAllModeIndicators() {
    const indicators = [
        'continuation-mode-indicator',
        'proofread-mode-indicator', 
        'reformat-mode-indicator'
    ];
    
    indicators.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.remove();
    });
    
    document.body.style.paddingTop = '';
}

// 添加CSS动画样式
(function() {
    const style = document.createElement('style');
    style.textContent = `
        @keyframes continuation-hint-pulse {
            0%, 100% { transform: scale(1); opacity: 1; }
            50% { transform: scale(1.1); opacity: 0.8; }
        }
        .continuation-btn:hover {
            opacity: 0.85;
        }
    `;
    document.head.appendChild(style);
})();
