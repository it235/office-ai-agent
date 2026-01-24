/**
 * code-handler.js - Code Block Handling
 * Functions for copying, executing, and editing code blocks
 */

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
