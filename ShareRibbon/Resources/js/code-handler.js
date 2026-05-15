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

/**
 * 完全隐藏代码块的操作栏（模板渲染模式使用）
 * @param {string} uuid - 消息的UUID
 */
function hideAllCodeBlockActions(uuid) {
    const messageContainer = document.getElementById('content-' + uuid);
    if (!messageContainer) return;
    
    // 隐藏所有代码块操作按钮（复制、编辑、执行）
    const codeButtons = messageContainer.querySelectorAll('.code-buttons');
    codeButtons.forEach(btn => btn.style.display = 'none');
    
    // 如果需要，也可以将代码块转换为普通文本显示
    const codeBlocks = messageContainer.querySelectorAll('.code-block');
    codeBlocks.forEach(block => {
        block.style.border = 'none';
        block.style.background = 'transparent';
        block.style.padding = '0';
    });
    
    // 隐藏代码折叠标签
    const toggleLabels = messageContainer.querySelectorAll('.code-toggle-label');
    toggleLabels.forEach(label => label.style.display = 'none');
    
    // 移除pre元素的折叠样式
    const preElements = messageContainer.querySelectorAll('pre.collapsible');
    preElements.forEach(pre => {
        pre.classList.remove('collapsible', 'collapsed');
    });
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
    let language = codeElement.className.replace('language-', '').replace(/\s*hljs\s*/g, '').trim();
    
    // 自动检测JSON：如果语言未标识或不明确，检查代码内容是否为JSON格式
    if (!language || language === '' || language === 'plaintext' || language === 'text') {
        const trimmedCode = code.trim();
        if ((trimmedCode.startsWith('{') && trimmedCode.endsWith('}')) ||
            (trimmedCode.startsWith('[') && trimmedCode.endsWith(']'))) {
            try {
                JSON.parse(trimmedCode);
                language = 'json';
                } catch (e) {
                // 不是有效的JSON，保持原语言
            }
        }
    }
    
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

    // 先尝试将 JSON 命令转换为执行步骤展示
    try {
        convertJsonToExecutionPlan(uuid);
    } catch (err) {
        console.error('convertJsonToExecutionPlan error:', err);
    }

    // Collapse code blocks (对于未转换的代码块)
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
}

// Render accept button (only for AI messages) - 移除了reject按钮
function renderAcceptRejectButtons(uuid) {
    try {
        const chatDiv = document.getElementById('chat-' + uuid);
        if (!chatDiv) return;
        
        const sender = chatDiv.dataset && chatDiv.dataset.sender ? chatDiv.dataset.sender : (chatDiv.querySelector('.sender-name') ? chatDiv.querySelector('.sender-name').textContent : '');

        // Only show button for AI messages
        if (!sender || sender === 'Me') return;

        const footer = document.getElementById('footer-' + uuid);
        if (!footer) return;

        // Skip if button already exists
        if (footer.querySelector('.accept-btn')) return;

        const btnAccept = document.createElement('button');
        btnAccept.className = 'code-button accept-btn';
        btnAccept.style.backgroundColor = '#4CAF50';
        btnAccept.onclick = function () { acceptAnswer(uuid); };

        footer.appendChild(btnAccept);
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

// 防抖标志 - 防止重复点击"继续改进"按钮
let rejectInProgress = false;

// Reject answer handler
function rejectAnswer(uuid) {
    // 防抖检查
    if (rejectInProgress) {
        return;
    }
    rejectInProgress = true;

    try {
        // 立即置灰按钮，防止重复点击
        const footer = document.getElementById('footer-' + uuid);
        if (footer) {
            footer.querySelectorAll('.accept-btn, .reject-btn').forEach(b => b.disabled = true);
        }

        const contentDiv = document.getElementById('content-' + uuid);
        const plainText = contentDiv ? (contentDiv.innerText || contentDiv.textContent || '') : '';

        let reason = '';
        try {
            reason = prompt('请简要说明希望如何改进（可留空）：', '');
            if (reason === null) {
                // 用户取消了，恢复按钮
                if (footer) {
                    footer.querySelectorAll('.accept-btn, .reject-btn').forEach(b => b.disabled = false);
                }
                rejectInProgress = false;
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

        // 显示状态提示
        if (footer) {
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
    } finally {
        // 500ms后解除防抖锁定
        setTimeout(() => { rejectInProgress = false; }, 500);
    }
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
    
    }

/**
 * 退出续写模式
 */
function exitContinuationMode() {
    window.continuationModeActive = false;
    window.continuationContext = null;
    
    // 恢复UI
    updateContinuationModeUI(false);
    
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

/**
 * 显示或隐藏AI续写按钮（由Ribbon续写功能调用）
 * @param {boolean} visible - 是否显示
 */
function setContinuationButtonVisible(visible) {
    const btn = document.getElementById('continuation-button');
    if (btn) {
        btn.style.display = visible ? 'inline-flex' : 'none';
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
        'reformat-mode-indicator',
        'template-mode-indicator'
    ];
    
    indicators.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.remove();
    });
    
    document.body.style.paddingTop = '';
}

// ==================== 模板渲染模式相关函数 ====================

/**
 * 进入模板渲染模式
 * @param {string} templateContext - 解析后的模板结构描述
 * @param {string} templateName - 模板文件名
 */
function enterTemplateMode(templateContext, templateName) {
    window.templateModeActive = true;
    window.currentTemplateContext = templateContext;
    window.currentTemplateName = templateName || '未命名模板';
    
    // 显示模式指示器
    showTemplateModeIndicator(window.currentTemplateName);
    
    }

/**
 * 退出模板渲染模式
 */
function exitTemplateMode() {
    window.templateModeActive = false;
    window.currentTemplateContext = null;
    window.currentTemplateName = null;
    
    // 隐藏模式指示器
    hideTemplateModeIndicator();
    
    }

/**
 * 显示模板模式指示器
 * @param {string} templateName - 模板文件名
 */
function showTemplateModeIndicator(templateName) {
    // 先隐藏其他模式指示器
    hideAllModeIndicators();
    
    const indicator = document.createElement('div');
    indicator.id = 'template-mode-indicator';
    indicator.innerHTML = `
        <div style="
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            background: linear-gradient(135deg, #9c27b0, #7b1fa2);
            color: white;
            padding: 8px 12px;
            font-size: 13px;
            z-index: 9999;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        ">
            <span>📋 模板模式 - 正在基于 "${templateName}" 生成内容</span>
            <button onclick="exitTemplateMode()" style="
                background: rgba(255,255,255,0.2);
                border: 1px solid rgba(255,255,255,0.4);
                color: white;
                padding: 4px 12px;
                border-radius: 4px;
                cursor: pointer;
                font-size: 12px;
            ">退出模板模式</button>
        </div>
    `;
    
    document.body.appendChild(indicator);
    document.body.style.paddingTop = '40px';
}

/**
 * 隐藏模板模式指示器
 */
function hideTemplateModeIndicator() {
    const indicator = document.getElementById('template-mode-indicator');
    if (indicator) {
        indicator.remove();
        document.body.style.paddingTop = '';
    }
}

/**
 * 显示模板内容预览界面（AI响应完成后调用）
 * @param {string} uuid - 消息的唯一标识
 */
function showTemplatePreview(uuid) {
    try {
        const chatSection = document.getElementById('chat-' + uuid);
        if (!chatSection) {
            console.error('showTemplatePreview: 找不到 chat section, uuid=' + uuid);
            return;
        }

        const contentEl = document.getElementById('content-' + uuid) || chatSection.querySelector('.message-content');
        if (!contentEl) {
            console.error('showTemplatePreview: 找不到 content 元素, uuid=' + uuid);
            return;
        }

        // 检查是否已经有模板操作按钮
        if (document.getElementById('template-actions-' + uuid)) return;

        // 隐藏常规聊天的 reject-btn（如果存在）
        const footer = document.getElementById('footer-' + uuid);
        if (footer) {
            const rejectBtn = footer.querySelector('.reject-btn');
            if (rejectBtn) rejectBtn.style.display = 'none';
        }

        // 检测应用类型
        const isPPT = window.officeAppType === 'PowerPoint';
        
        // 根据应用类型设置按钮文案
        const insertStartLabel = isPPT ? '插入首页' : '插入开头';
        const insertCurrentLabel = isPPT ? '插入当前页' : '插入当前位置';
        const insertEndLabel = isPPT ? '插入末页' : '插入结尾';

        // 创建模板操作按钮区域（紫色主题）
        const actionsHtml = `
            <div class="template-actions" id="template-actions-${uuid}" style="margin-top: 8px; padding: 8px; background: #f3e5f5; border-radius: 6px; border: 1px solid #ce93d8;">
                <div style="margin-bottom: 6px; font-size: 12px; color: #7b1fa2;">模板内容生成完成，选择插入位置：</div>
                <div style="margin-bottom: 6px;">
                    <button class="btn-primary template-btn" onclick="handleTemplateInsert('${uuid}', 'start')" style="background: #9c27b0; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; margin-right: 4px; font-size: 11px;">
                        ${insertStartLabel}
                    </button>
                    <button class="btn-primary template-btn" onclick="handleTemplateInsert('${uuid}', 'current')" style="background: #7b1fa2; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; margin-right: 4px; font-size: 11px;">
                        ${insertCurrentLabel}
                    </button>
                    <button class="btn-primary template-btn" onclick="handleTemplateInsert('${uuid}', 'end')" style="background: #9c27b0; color: white; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        ${insertEndLabel}
                    </button>
                </div>
                <div>
                    <button class="btn-secondary template-btn" onclick="handleTemplateRefine('${uuid}')" style="background: #e1bee7; color: #4a148c; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; margin-right: 4px; font-size: 11px;">
                        调整需求
                    </button>
                    <button class="btn-secondary template-btn" onclick="handleTemplateRegenerate()" style="background: #e1bee7; color: #4a148c; border: none; padding: 4px 8px; border-radius: 4px; cursor: pointer; font-size: 11px;">
                        重新生成
                    </button>
                </div>
            </div>
        `;
        contentEl.insertAdjacentHTML('afterend', actionsHtml);
        } catch (err) {
        console.error('showTemplatePreview error:', err);
    }
}

/**
 * 处理模板内容插入
 * @param {string} uuid - 消息的唯一标识
 * @param {string} position - 插入位置：'start'/'current'/'end'
 */
function handleTemplateInsert(uuid, position) {
    try {
        position = position || 'current';
        
        const chatSection = document.getElementById('chat-' + uuid);
        if (!chatSection) return;

        const contentEl = document.getElementById('content-' + uuid) || chatSection.querySelector('.message-content');
        if (!contentEl) return;

        // 获取生成的内容（纯文本）
        const content = contentEl.innerText || contentEl.textContent;

        // 发送插入请求到VB
        window.chrome.webview.postMessage({
            type: 'applyTemplateContent',
            uuid: uuid,
            content: content,
            position: position
        });

        // 移除操作按钮并显示成功提示
        removeTemplateActions(uuid);
        
        // 添加成功提示
        const successMsg = document.createElement('div');
        successMsg.style = 'margin-top: 8px; padding: 8px 12px; background: #e8f5e9; color: #2e7d32; border-radius: 6px; font-size: 13px;';
        successMsg.textContent = '模板内容已插入文档';
        contentEl.parentNode.appendChild(successMsg);
        
        // 3秒后移除提示
        setTimeout(() => successMsg.remove(), 3000);
    } catch (err) {
        console.error('handleTemplateInsert error:', err);
    }
}

/**
 * 移除模板操作按钮
 * @param {string} uuid - 消息的唯一标识
 */
function removeTemplateActions(uuid) {
    const actionsDiv = document.getElementById('template-actions-' + uuid);
    if (actionsDiv) {
        actionsDiv.remove();
    }
}

/**
 * 处理模板需求调整
 * @param {string} uuid - 消息的唯一标识
 */
function handleTemplateRefine(uuid) {
    try {
        const refinement = prompt('请输入调整需求（如：更详细、添加示例、换个风格等）：');
        if (refinement && refinement.trim()) {
            window.chrome.webview.postMessage({
                type: 'refineTemplateContent',
                uuid: uuid,
                refinement: refinement.trim()
            });
            
            // 更新按钮状态
            const actionsDiv = document.getElementById('template-actions-' + uuid);
            if (actionsDiv) {
                actionsDiv.innerHTML = '<div style="color: #7b1fa2; font-size: 13px;">正在根据您的要求调整内容...</div>';
            }
        }
    } catch (err) {
        console.error('handleTemplateRefine error:', err);
    }
}

/**
 * 处理重新生成模板内容
 */
function handleTemplateRegenerate() {
    try {
        if (window.templateModeActive && window.currentTemplateContext) {
            const input = document.getElementById('smart-input');
            if (input) {
                input.focus();
                alert('请在输入框中重新描述您的内容需求，然后点击发送。');
            }
        }
    } catch (err) {
        console.error('handleTemplateRegenerate error:', err);
    }
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

/**
 * Agent模式下自动执行执行计划（仍然弹出预览框让用户确认）
 * @param {string} uuid - 消息 UUID
 */
function executePlanFromRendererAutoMode(uuid) {
    try {
        const contentDiv = document.getElementById('content-' + uuid);
        if (!contentDiv) return;

        const container = contentDiv.querySelector('.execution-plan-container');
        if (!container) return;

        const codeElement = container.querySelector('.original-code code');
        if (!codeElement) return;

        const code = codeElement.textContent;

        // Agent模式也弹出预览框让用户确认执行
        const payload = {
            type: 'executeCode',
            code: code,
            language: 'json',
            executecodePreview: true, // 弹出预览框让用户确认
            responseUuid: uuid,
            autoMode: true
        };

        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage(payload);
        } else if (window.vsto) {
            window.vsto.executeCode(code, 'json', true);
        }

        // UI反馈
        const btn = container.querySelector('.execute-plan-btn');
        if (btn) {
            btn.textContent = '等待确认...';
            btn.disabled = true;
            // 保存按钮引用，以便执行结果返回后恢复
            btn.dataset.originalText = '执行此计划';
        }
    } catch (err) {
        console.error('executePlanFromRendererAutoMode error:', err);
    }
}

/**
 * Agent模式下自动执行代码（仍然弹出预览框让用户确认）
 * @param {HTMLElement} button - 执行按钮元素
 */
function executeCodeAutoMode(button) {
    try {
        const codeBlock = button.closest('.code-block');
        if (!codeBlock) return;

        const codeElement = codeBlock.querySelector('code');
        if (!codeElement) return;

        const code = codeElement.textContent;
        let language = codeElement.className.replace('language-', '').replace(/\s*hljs\s*/g, '').trim();
        
        // 自动检测JSON
        if (!language || language === '' || language === 'plaintext' || language === 'text') {
            const trimmedCode = code.trim();
            if ((trimmedCode.startsWith('{') && trimmedCode.endsWith('}')) ||
                (trimmedCode.startsWith('[') && trimmedCode.endsWith(']'))) {
                try {
                    JSON.parse(trimmedCode);
                    language = 'json';
                } catch (e) {}
            }
        }

        // 获取UUID
        const chatContainer = button.closest('.chat-container');
        let responseUuid = null;
        if (chatContainer && chatContainer.id && chatContainer.id.startsWith('chat-')) {
            responseUuid = chatContainer.id.replace('chat-', '');
        }

        // Agent模式也弹出预览框让用户确认
        const payload = {
            type: 'executeCode',
            code: code,
            language: language,
            executecodePreview: true, // 弹出预览框让用户确认
            responseUuid: responseUuid,
            autoMode: true
        };

        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage(payload);
        } else if (window.vsto) {
            window.vsto.executeCode(code, language, true);
        }

        // UI反馈
        const originalText = button.innerHTML;
        button.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polygon points="5 3 19 12 5 21 5 3"></polygon>
            </svg>
            等待确认...
        `;
        button.dataset.originalHtml = originalText;
        button.disabled = true;
    } catch (err) {
        console.error('executeCodeAutoMode error:', err);
    }
}

// 导出函数供全局使用
window.executePlanFromRendererAutoMode = executePlanFromRendererAutoMode;
window.executeCodeAutoMode = executeCodeAutoMode;

// ========== JSON 命令转执行步骤功能 ==========

/**
 * 检测并将 JSON 代码块转换为执行步骤展示
 * @param {string} uuid - 消息的 UUID
 * @returns {boolean} 是否有 JSON 命令被转换
 */
function convertJsonToExecutionPlan(uuid) {
    try {
        const contentDiv = document.getElementById('content-' + uuid);
        if (!contentDiv) return false;

        const codeBlocks = contentDiv.querySelectorAll('pre code');
        let converted = false;

        codeBlocks.forEach(codeBlock => {
            // 检测是否为 JSON
            const language = codeBlock.className.replace('language-', '').replace(/\s*hljs\s*/g, '').trim().toLowerCase();
            const code = codeBlock.textContent.trim();

            // 检测 JSON 格式
            if (isJsonCommand(code, language)) {
                try {
                    const json = JSON.parse(code);
                    if (json.command) {
                        // 是有效的命令 JSON，转换为执行步骤
                        const planHtml = buildExecutionPlanHtml(json, uuid, code);
                        
                        // 替换代码块
                        const codeBlockContainer = codeBlock.closest('.code-block');
                        if (codeBlockContainer) {
                            const planContainer = document.createElement('div');
                            planContainer.innerHTML = planHtml;
                            codeBlockContainer.parentNode.replaceChild(planContainer.firstElementChild, codeBlockContainer);
                            converted = true;
                        }
                    }
                } catch (e) {
                    // JSON 解析失败，保持原样
                }
            }
        });

        return converted;
    } catch (err) {
        console.error('convertJsonToExecutionPlan error:', err);
        return false;
    }
}

/**
 * 检测代码是否为 JSON 命令
 * @param {string} code - 代码内容
 * @param {string} language - 语言标识
 * @returns {boolean}
 */
function isJsonCommand(code, language) {
    if (!code) return false;

    // 语言标识检测
    if (language === 'json') return true;

    // 内容检测
    const trimmed = code.trim();
    if ((trimmed.startsWith('{') && trimmed.endsWith('}')) ||
        (trimmed.startsWith('[') && trimmed.endsWith(']'))) {
        try {
            const parsed = JSON.parse(trimmed);
            // 检查是否有 command 字段
            return parsed && (parsed.command || (Array.isArray(parsed) && parsed[0] && parsed[0].command));
        } catch (e) {
            return false;
        }
    }
    return false;
}

/**
 * 构建执行步骤的 HTML
 * @param {Object} json - JSON 命令对象
 * @param {string} uuid - 消息 UUID
 * @param {string} originalCode - 原始 JSON 代码
 * @returns {string} HTML 字符串
 */
function buildExecutionPlanHtml(json, uuid, originalCode) {
    const plan = parseJsonToPlan(json);
    const planId = uuid + '-plan';

    let stepsHtml = plan.steps.map((step, idx) => {
        const icon = getStepIcon(step.icon);
        const willModify = step.willModify ? `<span class="modify-badge">→ ${escapeHtml(step.willModify)}</span>` : '';
        const estimatedTime = step.estimatedTime ? `<span class="time-badge">⏱️ ${step.estimatedTime}</span>` : '';

        return `
            <div class="plan-step">
                <span class="step-badge">${idx + 1}</span>
                <div class="step-content">
                    <div class="step-title">${icon} ${escapeHtml(step.description)}</div>
                    ${(willModify || estimatedTime) ? `<div class="step-details">${willModify}${estimatedTime}</div>` : ''}
                </div>
            </div>
        `;
    }).join('');

    return `
        <div class="execution-plan-container" data-uuid="${uuid}" data-plan-id="${planId}">
            <div class="plan-header">📋 执行计划</div>
            <div class="plan-steps">
                ${stepsHtml}
            </div>
            <div class="plan-actions">
                <button class="execute-plan-btn" onclick="executePlanFromRenderer('${uuid}', this)">执行此计划</button>
                <button class="show-code-btn" onclick="toggleCodeViewFromRenderer('${planId}')">查看代码</button>
            </div>
            <div class="original-code" id="code-${planId}">
                <pre><code class="language-json">${escapeHtml(originalCode)}</code></pre>
            </div>
        </div>
    `;
}

/**
 * 将 JSON 命令解析为执行步骤
 * @param {Object} json - JSON 命令对象
 * @returns {Object} 包含 steps 数组的对象
 */
function parseJsonToPlan(json) {
    const steps = [];
    const command = json.command || '';
    const params = json.params || {};

    // 命令描述映射
    const commandDescriptions = {
        'ApplyFormula': { desc: '应用公式', icon: 'formula' },
        'WriteData': { desc: '写入数据', icon: 'data' },
        'FormatRange': { desc: '格式化区域', icon: 'format' },
        'CreateChart': { desc: '创建图表', icon: 'chart' },
        'CleanData': { desc: '清洗数据', icon: 'clean' },
        'DataAnalysis': { desc: '数据分析', icon: 'data' },
        'TransformData': { desc: '数据转换', icon: 'data' },
        'GenerateReport': { desc: '生成报表', icon: 'data' }
    };

    const cmdInfo = commandDescriptions[command] || { desc: command, icon: 'default' };

    // 根据命令类型生成步骤
    switch (command.toLowerCase()) {
        case 'applyformula':
        case 'formula':
            steps.push({
                description: `在 ${params.targetRange || '目标区域'} 应用公式`,
                icon: 'formula',
                willModify: params.targetRange,
                estimatedTime: '1秒'
            });
            if (params.formula) {
                steps.push({
                    description: `公式: ${getFormulaDescription(params.formula)}`,
                    icon: 'formula'
                });
            }
            if (params.fillDown) {
                steps.push({
                    description: '自动向下填充',
                    icon: 'formula'
                });
            }
            break;

        case 'createchart':
        case 'chart':
            const chartTypes = { 'Column': '柱状图', 'Line': '折线图', 'Pie': '饼图', 'Bar': '条形图' };
            const chartType = chartTypes[params.type] || params.type || '图表';
            steps.push({
                description: `读取 ${params.dataRange || '数据区域'} 作为图表数据`,
                icon: 'search'
            });
            steps.push({
                description: `创建 ${chartType}`,
                icon: 'chart',
                estimatedTime: '2秒'
            });
            if (params.title) {
                steps.push({
                    description: `设置标题: ${params.title}`,
                    icon: 'chart'
                });
            }
            break;

        case 'formatrange':
        case 'format':
            const range = params.range || params.targetRange || '目标区域';
            steps.push({
                description: `选择 ${range} 区域`,
                icon: 'search'
            });
            let formatDesc = '应用格式设置';
            if (params.style) {
                formatDesc = `应用 ${params.style} 样式`;
            }
            steps.push({
                description: formatDesc,
                icon: 'format',
                willModify: range,
                estimatedTime: '1秒'
            });
            break;

        case 'cleandata':
        case 'clean':
            const operations = {
                'removeDuplicates': '删除重复项',
                'fillEmpty': '填充空值',
                'trim': '去除空格'
            };
            const opDesc = operations[params.operation] || params.operation || '清洗';
            steps.push({
                description: `扫描 ${params.range || '数据区域'}`,
                icon: 'search'
            });
            steps.push({
                description: `执行: ${opDesc}`,
                icon: 'clean',
                willModify: params.range,
                estimatedTime: '2秒'
            });
            break;

        default:
            steps.push({
                description: `执行 ${cmdInfo.desc}`,
                icon: cmdInfo.icon,
                estimatedTime: '1秒'
            });
    }

    return { steps };
}

/**
 * 获取公式的友好描述
 * @param {string} formula - 公式字符串
 * @returns {string}
 */
function getFormulaDescription(formula) {
    if (!formula) return '';
    formula = formula.replace(/^=/, '');
    const upper = formula.toUpperCase();

    if (upper.startsWith('SUM(')) return '求和';
    if (upper.startsWith('AVERAGE(')) return '平均值';
    if (upper.startsWith('COUNT(')) return '计数';
    if (upper.startsWith('MAX(')) return '最大值';
    if (upper.startsWith('MIN(')) return '最小值';
    if (upper.startsWith('VLOOKUP(')) return '垂直查找';
    if (upper.startsWith('IF(')) return '条件判断';
    if (formula.includes('+')) return '加法运算';
    if (formula.includes('-')) return '减法运算';
    if (formula.includes('*')) return '乘法运算';
    if (formula.includes('/')) return '除法运算';

    return formula.length > 25 ? formula.substring(0, 22) + '...' : formula;
}

/**
 * 获取步骤图标
 * @param {string} iconType - 图标类型
 * @returns {string} emoji
 */
function getStepIcon(iconType) {
    const icons = {
        'search': '🔍',
        'data': '📊',
        'formula': '🧮',
        'chart': '📈',
        'format': '🎨',
        'clean': '🧹',
        'default': '⚡'
    };
    return icons[iconType] || icons['default'];
}

/**
 * 执行计划按钮点击处理
 * @param {string} uuid - 消息 UUID
 * @param {HTMLElement} button - 按钮元素
 */
function executePlanFromRenderer(uuid, button) {
    try {
        // 找到原始代码
        const container = button.closest('.execution-plan-container');
        if (!container) return;

        const codeElement = container.querySelector('.original-code code');
        if (!codeElement) return;

        const code = codeElement.textContent;
        const preview = document.getElementById('settings-executecode-preview')?.checked || false;

        // 发送执行请求
        const payload = {
            type: 'executeCode',
            code: code,
            language: 'json',
            executecodePreview: preview,
            responseUuid: uuid
        };

        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage(payload);
        } else if (window.vsto) {
            window.vsto.executeCode(code, 'json', preview);
        }

        // UI 反馈
        button.textContent = '已执行';
        button.disabled = true;
        setTimeout(() => {
            button.textContent = '执行此计划';
            button.disabled = false;
        }, 2000);
    } catch (err) {
        console.error('executePlanFromRenderer error:', err);
        alert('执行失败：' + err.message);
    }
}

/**
 * 切换代码视图显示/隐藏
 * @param {string} planId - 计划 ID
 */
function toggleCodeViewFromRenderer(planId) {
    try {
        const codeDiv = document.getElementById('code-' + planId);
        if (codeDiv) {
            codeDiv.classList.toggle('visible');

            // 更新按钮文字
            const container = codeDiv.closest('.execution-plan-container');
            if (container) {
                const btn = container.querySelector('.show-code-btn');
                if (btn) {
                    btn.textContent = codeDiv.classList.contains('visible') ? '隐藏代码' : '查看代码';
                }
            }

            // 高亮代码
            if (codeDiv.classList.contains('visible')) {
                const codeBlock = codeDiv.querySelector('code');
                if (codeBlock && typeof hljs !== 'undefined') {
                    hljs.highlightElement(codeBlock);
                }
            }
        }
    } catch (err) {
        console.error('toggleCodeViewFromRenderer error:', err);
    }
}

// 导出函数供全局使用
window.convertJsonToExecutionPlan = convertJsonToExecutionPlan;
window.executePlanFromRenderer = executePlanFromRenderer;
window.toggleCodeViewFromRenderer = toggleCodeViewFromRenderer;

// ========== 执行结果处理函数 ==========

/**
 * 处理代码执行成功
 * @param {string} uuid - 消息 UUID
 */
function handleExecutionSuccess(uuid) {
    try {
        // 清空引用区
        clearAllReferences();
        
        // 恢复执行按钮状态
        restoreExecuteButtons(uuid, true);
    } catch (err) {
        console.error('handleExecutionSuccess error:', err);
    }
}

/**
 * 处理代码执行失败
 * @param {string} uuid - 消息 UUID
 * @param {string} errorMsg - 错误信息
 */
function handleExecutionError(uuid, errorMsg) {
    try {
        // 恢复执行按钮状态（可再次点击）
        restoreExecuteButtons(uuid, false);
    } catch (err) {
        console.error('handleExecutionError error:', err);
    }
}

/**
 * 处理用户取消执行（在预览对话框中点击取消）
 * @param {string} uuid - 消息 UUID
 */
function handleExecutionCancelled(uuid) {
    try {
        // 恢复执行按钮状态
        restoreExecuteButtons(uuid, false);
    } catch (err) {
        console.error('handleExecutionCancelled error:', err);
    }
}

/**
 * 恢复执行按钮状态
 * @param {string} uuid - 消息 UUID
 * @param {boolean} success - 是否执行成功
 */
function restoreExecuteButtons(uuid, success) {
    try {
        const contentDiv = document.getElementById('content-' + uuid);
        if (!contentDiv) return;

        // 恢复执行计划按钮
        const planContainer = contentDiv.querySelector('.execution-plan-container');
        if (planContainer) {
            const btn = planContainer.querySelector('.execute-plan-btn');
            if (btn) {
                if (success) {
                    btn.textContent = '已执行';
                    btn.disabled = true;
                    // 5秒后恢复可点击状态，允许重复执行
                    setTimeout(() => {
                        btn.textContent = btn.dataset.originalText || '执行此计划';
                        btn.disabled = false;
                    }, 5000);
                } else {
                    btn.textContent = '重试';
                    btn.disabled = false;
                }
            }
        }

        // 恢复普通执行按钮
        const executeButtons = contentDiv.querySelectorAll('.execute-button');
        executeButtons.forEach(btn => {
            if (success) {
                btn.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <polyline points="20 6 9 17 4 12"></polyline>
                    </svg>
                    已执行
                `;
                btn.disabled = true;
                // 5秒后恢复可点击状态，允许重复执行
                setTimeout(() => {
                    if (btn.dataset.originalHtml) {
                        btn.innerHTML = btn.dataset.originalHtml;
                    } else {
                        btn.innerHTML = `
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                                <polygon points="5 3 19 12 5 21 5 3"></polygon>
                            </svg>
                            执行
                        `;
                    }
                    btn.disabled = false;
                }, 5000);
            } else {
                // 执行失败，显示重试按钮
                btn.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <polygon points="5 3 19 12 5 21 5 3"></polygon>
                    </svg>
                    重试
                `;
                btn.disabled = false;
            }
        });
    } catch (err) {
        console.error('restoreExecuteButtons error:', err);
    }
}

/**
 * 清空所有引用区（文件和选中内容）
 */
function clearAllReferences() {
    try {
        // 清空选中内容
        if (window.selectedContentMap) {
            window.selectedContentMap = {};
        }
        
        // 清空附加文件
        if (window.attachedFiles) {
            window.attachedFiles = [];
        }
        
        // 重新渲染引用区
        if (typeof renderReferences === 'function') {
            renderReferences();
        }
        
        } catch (err) {
        console.error('clearAllReferences error:', err);
    }
}

// 导出执行结果处理函数
window.handleExecutionSuccess = handleExecutionSuccess;
window.handleExecutionError = handleExecutionError;
window.handleExecutionCancelled = handleExecutionCancelled;
window.clearAllReferences = clearAllReferences;

// ========== 文件解析进度显示 ==========

/**
 * 显示/隐藏文件解析进度
 * @param {boolean} show - 是否显示
 */
function showFileParsingProgress(show) {
    try {
        let progressOverlay = document.getElementById('file-parsing-progress');
        
        if (show) {
            if (!progressOverlay) {
                progressOverlay = document.createElement('div');
                progressOverlay.id = 'file-parsing-progress';
                progressOverlay.innerHTML = `
                    <div class="progress-content">
                        <div class="progress-spinner"></div>
                        <div class="progress-text">正在解析文件...</div>
                        <div class="progress-detail" id="file-parsing-detail"></div>
                    </div>
                `;
                progressOverlay.style.cssText = `
                    position: fixed;
                    top: 0;
                    left: 0;
                    right: 0;
                    bottom: 0;
                    background: rgba(0,0,0,0.5);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10000;
                `;
                
                const style = document.createElement('style');
                style.id = 'file-parsing-progress-style';
                style.textContent = `
                    #file-parsing-progress .progress-content {
                        background: white;
                        padding: 24px 32px;
                        border-radius: 8px;
                        text-align: center;
                        box-shadow: 0 4px 20px rgba(0,0,0,0.2);
                        min-width: 280px;
                    }
                    #file-parsing-progress .progress-spinner {
                        width: 40px;
                        height: 40px;
                        border: 4px solid #e0e0e0;
                        border-top-color: #4a6fa5;
                        border-radius: 50%;
                        margin: 0 auto 16px;
                        animation: file-parsing-spin 1s linear infinite;
                    }
                    @keyframes file-parsing-spin {
                        to { transform: rotate(360deg); }
                    }
                    #file-parsing-progress .progress-text {
                        font-size: 14px;
                        color: #333;
                        margin-bottom: 8px;
                    }
                    #file-parsing-progress .progress-detail {
                        font-size: 12px;
                        color: #666;
                    }
                `;
                document.head.appendChild(style);
                document.body.appendChild(progressOverlay);
            } else {
                progressOverlay.style.display = 'flex';
            }
        } else {
            if (progressOverlay) {
                progressOverlay.style.display = 'none';
            }
        }
    } catch (err) {
        console.error('showFileParsingProgress error:', err);
    }
}

/**
 * 更新文件解析进度
 * @param {number} current - 当前进度
 * @param {number} total - 总数
 * @param {string} fileName - 当前文件名
 */
function updateFileParsingProgress(current, total, fileName) {
    try {
        const textEl = document.querySelector('#file-parsing-progress .progress-text');
        const detailEl = document.getElementById('file-parsing-detail');
        
        if (textEl) {
            textEl.textContent = `正在解析文件 (${current}/${total})`;
        }
        if (detailEl) {
            detailEl.textContent = fileName || '';
        }
    } catch (err) {
        console.error('updateFileParsingProgress error:', err);
    }
}

// 导出文件解析进度函数
window.showFileParsingProgress = showFileParsingProgress;
window.updateFileParsingProgress = updateFileParsingProgress;

// ============================================================
// 语义排版结果展示 + 撤销/确认
// ============================================================

/**
 * 显示排版结果卡片（渲染引擎完成后由VB调用）
 * @param {Object} result - {appliedCount, skippedCount, tags: {tagId: count}, success, error}
 */
function showReformatResult(result) {
    try {
        // 隐藏排版模式指示器
        hideReformatModeIndicator();

        const chatContainer = document.getElementById('chat-container');
        if (!chatContainer) return;

        // 处理错误情况（无排版可撤销）
        if (result.success === false) {
            const errorCard = document.createElement('div');
            errorCard.className = 'reformat-result-card';
            errorCard.innerHTML = `
                <div style="background: #fff5f5; border: 1px solid #fed7d7; border-radius: 8px; padding: 12px 16px; margin: 8px 0;">
                    <div style="font-size: 13px; font-weight: 600; color: #c53030; margin-bottom: 6px;">
                        排版失败
                    </div>
                    <div style="font-size: 12px; color: #718096; word-break: break-all;">${result.error || '未知错误'}</div>
                </div>
            `;
            chatContainer.appendChild(errorCard);
            chatContainer.scrollTop = chatContainer.scrollHeight;
            return;
        }

        // 构建标签使用统计
        let tagStats = '';
        if (result.tags) {
            const entries = Object.entries(result.tags);
            tagStats = entries.map(([tag, count]) => `${tag} × ${count}`).join(' | ');
        }

        const card = document.createElement('div');
        card.className = 'reformat-result-card';
        const hasChanges = (result.appliedCount || 0) > 0;
        card.innerHTML = `
            <div style="background: #f8f9fa; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px 16px; margin: 8px 0;">
                <div style="font-size: 13px; font-weight: 600; color: #2d3748; margin-bottom: 6px;">
                    排版完成：处理 ${result.appliedCount || 0} 个段落${result.skippedCount > 0 ? `，跳过 ${result.skippedCount} 个特殊元素` : ''}
                </div>
                ${tagStats ? `<div style="font-size: 11px; color: #718096; margin-bottom: 8px; word-break: break-all;">${tagStats}</div>` : ''}
                <div style="display: flex; gap: 8px;">
                    ${hasChanges ? `<button onclick="undoReformat()" style="padding: 5px 14px; font-size: 12px; border: 1px solid #e53e3e; color: #e53e3e; background: white; border-radius: 4px; cursor: pointer;">
                        撤销排版
                    </button>` : ''}
                    <button onclick="reenterReformatMode()" style="padding: 5px 14px; font-size: 12px; border: 1px solid #667eea; color: #667eea; background: white; border-radius: 4px; cursor: pointer;">
                        继续排版
                    </button>
                    <button onclick="acceptReformat(this)" style="padding: 5px 14px; font-size: 12px; border: 1px solid #38a169; color: white; background: #38a169; border-radius: 4px; cursor: pointer;">
                        确认
                    </button>
                </div>
            </div>
        `;

        chatContainer.appendChild(card);
        chatContainer.scrollTop = chatContainer.scrollHeight;
    } catch (err) {
        console.error('showReformatResult error:', err);
    }
}

/**
 * 撤销排版（发送消息到VB后端）
 */
function undoReformat() {
    try {
        const payload = JSON.stringify({ type: 'undoReformat' });
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage(payload);
        } else if (window.vsto) {
            window.vsto.postMessage(payload);
        }

        // 移除结果卡片
        const cards = document.querySelectorAll('.reformat-result-card');
        cards.forEach(c => c.remove());
    } catch (err) {
        console.error('undoReformat error:', err);
    }
}

/**
 * 确认排版结果
 * @param {HTMLElement} btn - 确认按钮
 */
function acceptReformat(btn) {
    try {
        // 隐藏结果卡片
        const card = btn ? btn.closest('.reformat-result-card') : null;
        if (card) {
            card.style.opacity = '0.5';
            card.querySelector('div > div:last-child').innerHTML = '<span style="color: #38a169; font-size: 12px;">已确认</span>';
        }
    } catch (err) {
        console.error('acceptReformat error:', err);
    }
}

/**
 * 重新进入排版模板选择模式（从结果卡片触发）
 */
function reenterReformatMode() {
    try {
        // 移除结果卡片
        const cards = document.querySelectorAll('.reformat-result-card');
        cards.forEach(c => c.remove());

        // 隐藏排版模式指示器
        hideReformatModeIndicator();

        // 重新进入模板选择模式
        if (typeof enterReformatTemplateMode === 'function') {
            enterReformatTemplateMode();
        }
    } catch (err) {
        console.error('reenterReformatMode error:', err);
    }
}
