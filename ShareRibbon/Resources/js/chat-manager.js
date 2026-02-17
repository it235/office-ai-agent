/**
 * chat-manager.js - Chat Section Management
 * Functions for creating and managing chat message sections
 */

// Create chat section with sender info and content area
window.createChatSection = function (sender, timestamp, uuid) {
    console.log('[createChatSection] 被调用, sender=' + sender + ', uuid=' + uuid);
    
    // Create chat container
    const chatContainer = document.createElement('div');
    chatContainer.className = 'chat-container';
    chatContainer.id = 'chat-' + uuid;

    // Record sender for later reference
    chatContainer.dataset.sender = sender;

    // Create message header
    const messageHeader = document.createElement('div');
    messageHeader.className = 'message-header';

    // Add avatar
    const avatar = document.createElement('div');
    if (sender !== 'Me') {
        avatar.innerHTML = 'AI';
        avatar.className = 'avatar-ai';
        
        // 如果在Ralph Loop模式中，添加步骤标签
        if (window.ralphLoopState && window.ralphLoopState.active && window.ralphLoopState.currentSession) {
            const session = window.ralphLoopState.currentSession;
            const runningStep = session.steps ? session.steps.findIndex(s => s.status === 'running') : -1;
            if (runningStep >= 0) {
                chatContainer.dataset.loopStep = runningStep;
                chatContainer.classList.add('ralph-loop-message');
            }
        }
    } else {
        avatar.innerHTML = 'Me';
        avatar.className = 'avatar-me';
    }
    messageHeader.appendChild(avatar);

    // Add sender name and timestamp
    const senderInfo = document.createElement('div');
    senderInfo.className = 'sender-info';

    const senderName = document.createElement('div');
    senderName.className = 'sender-name';
    
    // 如果是Ralph Loop步骤，添加步骤标记
    if (chatContainer.dataset.loopStep !== undefined && sender !== 'Me') {
        const stepNum = parseInt(chatContainer.dataset.loopStep) + 1;
        senderName.innerHTML = sender + ' <span class="loop-step-badge">步骤 ' + stepNum + '</span>';
    } else {
        senderName.textContent = sender;
    }

    const timestampElem = document.createElement('div');
    timestampElem.className = 'timestamp';
    timestampElem.textContent = timestamp;

    senderInfo.appendChild(senderName);
    senderInfo.appendChild(timestampElem);
    messageHeader.appendChild(senderInfo);

    chatContainer.appendChild(messageHeader);

    // Create content area
    const contentContainer = document.createElement('div');
    contentContainer.className = 'message-content';
    contentContainer.id = 'content-' + uuid;
    chatContainer.appendChild(contentContainer);

    // Add footer area (for token count / accept-reject buttons)
    const footerContainer = document.createElement('div');
    footerContainer.className = 'message-footer';
    footerContainer.id = 'footer-' + uuid;
    chatContainer.appendChild(footerContainer);

    // Add to chat history container
    const chatHistoryContainer = document.getElementById('chat-container');
    if (chatHistoryContainer) {
        chatHistoryContainer.appendChild(chatContainer);
    } else {
        console.error('找不到 chat-container 元素，请确保页面中存在此元素');
        document.body.appendChild(chatContainer);
    }

    // Create renderer - 直接传入 DOM 元素而不是 ID 字符串
    window.rendererMap[uuid] = new MarkdownStreamRenderer(contentContainer);

    return uuid;;
};

// Create reasoning container for a chat section
window.createReasoningContainer = function (uuid) {
    if (!uuid) {
        console.error('UUID不能为空');
        return false;
    }

    // Check if reasoning container already exists
    if (document.getElementById('reasoning-' + uuid)) {
        console.warn('UUID为' + uuid + '的推理容器已存在');
        return true;
    }

    const chatContainer = document.getElementById('chat-' + uuid);
    if (!chatContainer) {
        console.error('找不到UUID为' + uuid + '的聊天容器');
        return false;
    }

    // Create reasoning container
    const reasoningContainer = document.createElement('div');
    reasoningContainer.className = 'reasoning-container';
    reasoningContainer.id = 'reasoning-' + uuid;

    // Create reasoning header
    const reasoningHeader = document.createElement('div');
    reasoningHeader.className = 'reasoning-header';
    reasoningHeader.innerHTML = `
        <span class="reasoning-title">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
                <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13h2v6h-2zm0 8h2v2h-2z"/>
            </svg>
            思考过程
        </span>
        <span class="reasoning-toggle">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
                <path d="M7 10l5 5 5-5z"/>
            </svg>
        </span>
    `;

    // Add click event for collapse/expand
    reasoningHeader.addEventListener('click', function () {
        reasoningContainer.classList.toggle('collapsed');
    });

    // Create reasoning content area
    const reasoningContent = document.createElement('div');
    reasoningContent.className = 'reasoning-content';
    reasoningContent.id = 'reasoning-content-' + uuid;

    // Assemble reasoning container
    reasoningContainer.appendChild(reasoningHeader);
    reasoningContainer.appendChild(reasoningContent);

    // Insert between message header and content
    const contentContainer = document.getElementById('content-' + uuid);
    chatContainer.insertBefore(reasoningContainer, contentContainer);

    // Create reasoning renderer
    window.reasoningRendererMap[uuid] = new MarkdownStreamRenderer(reasoningContent);

    return true;
};

// Append content to renderer
window.appendRenderer = function (uuid, text) {
    console.log('[appendRenderer] 被调用, uuid=' + uuid + ', text长度=' + (text ? text.length : 0));
    
    if (!uuid) {
        console.error('[appendRenderer] uuid为空');
        return false;
    }

    // Auto collapse reasoning container if exists and not collapsed
    const reasoningContainer = document.getElementById('reasoning-' + uuid);
    if (reasoningContainer && !reasoningContainer.classList.contains('collapsed')) {
        reasoningContainer.classList.add('collapsed');
    }

    const renderer = window.rendererMap[uuid];
    console.log('[appendRenderer] rendererMap中是否存在该uuid:', !!renderer);
    console.log('[appendRenderer] 当前rendererMap keys:', Object.keys(window.rendererMap || {}));
    
    if (renderer) {
        renderer.append(text);
        contentScroll();
        return true;
    } else {
        console.error('[appendRenderer] 找不到UUID为' + uuid + '的渲染器');
        return false;
    }
};

// Append reasoning content
window.appendReasoning = function (uuid, text) {
    if (!uuid) {
        console.error('UUID不能为空');
        return false;
    }

    // Create reasoning container if not exists
    if (!document.getElementById('reasoning-' + uuid)) {
        if (!window.createReasoningContainer(uuid)) {
            return false;
        }
    }

    const reasoningRenderer = window.reasoningRendererMap[uuid];
    if (reasoningRenderer) {
        reasoningRenderer.append(text);

        // Ensure reasoning container is visible during reasoning
        const reasoningContainer = document.getElementById('reasoning-' + uuid);
        if (reasoningContainer) {
            reasoningContainer.classList.remove('collapsed');
        }
        
        // Scroll reasoning content to bottom
        const reasoningContent = document.getElementById('reasoning-content-' + uuid);
        if (reasoningContent) {
            reasoningContent.scrollTop = reasoningContent.scrollHeight;
        }
        contentScroll();
        return true;
    } else {
        console.error('找不到UUID为' + uuid + '的推理渲染器');
        return false;
    }
};

// Complete reasoning process, collapse reasoning area
window.completeReasoning = function (uuid) {
    if (!uuid) {
        console.error('UUID不能为空');
        return false;
    }

    const reasoningContainer = document.getElementById('reasoning-' + uuid);
    if (reasoningContainer) {
        // Brief delay to let user see final reasoning result
        setTimeout(() => {
            reasoningContainer.classList.add('collapsed');
        }, 1000);
        return true;
    } else {
        console.error('找不到UUID为' + uuid + '的推理容器');
        return false;
    }
};

// Clear all chat content
window.clearAllChats = function () {
    const chatContainer = document.getElementById('chat-container');
    if (chatContainer) {
        chatContainer.innerHTML = '';
        window.rendererMap = {};
        window.reasoningRendererMap = {};
        window.autoScrollEnabled = true;
        console.log('已清空所有聊天内容');
        return true;
    } else {
        console.error('找不到聊天容器');
        return false;
    }
};

// 供 VB 调用：清空当前聊天区域（新会话）
window.clearChatContent = function () {
    return window.clearAllChats ? window.clearAllChats() : false;
};

// 供 VB 调用：加载历史会话消息并渲染
window.setChatMessages = function (messages) {
    if (!messages || !Array.isArray(messages) || messages.length === 0) {
        if (window.clearChatContent) window.clearChatContent();
        return;
    }
    if (window.clearChatContent) window.clearChatContent();
    const chatContainer = document.getElementById('chat-container');
    if (!chatContainer) return;
    for (let i = 0; i < messages.length; i++) {
        const m = messages[i];
        const role = (m.role || '').toLowerCase();
        const sender = role === 'user' ? 'Me' : 'AI';
        const createTime = m.createTime || new Date().toLocaleString('zh-CN');
        const uuid = 'hist-' + i + '-' + Date.now();
        if (typeof createChatSection === 'function') {
            createChatSection(sender, createTime, uuid);
        }
        const contentEl = document.getElementById('content-' + uuid);
        if (contentEl && typeof marked !== 'undefined') {
            try {
                contentEl.innerHTML = marked.parse(m.content || '');
            } catch (e) {
                contentEl.textContent = m.content || '';
            }
        }
    }
    if (window.autoScrollEnabled !== false && chatContainer.lastElementChild) {
        chatContainer.lastElementChild.scrollIntoView({ behavior: 'smooth' });
    }
};

// Get full chat container HTML
window.getFullChatHTML = function () {
    const chatContainer = document.getElementById('chat-container');
    if (chatContainer) {
        return chatContainer.innerHTML;
    }
    return '';
};

// Set full chat container HTML
window.setFullChatHTML = function (html) {
    const chatContainer = document.getElementById('chat-container');
    if (chatContainer) {
        chatContainer.innerHTML = '';
        chatContainer.innerHTML = html;
        window.rebuildRendererMaps();
        return true;
    }
    return false;
};

// Rebuild renderer maps after setting HTML
window.rebuildRendererMaps = function () {
    window.rendererMap = {};
    window.reasoningRendererMap = {};

    // Find all content and reasoning containers
    const contentContainers = document.querySelectorAll('.message-content');
    const reasoningContainers = document.querySelectorAll('.reasoning-content');

    // Rebuild content renderer map - 直接传入 DOM 元素
    contentContainers.forEach(container => {
        const id = container.id;
        if (id && id.startsWith('content-')) {
            const uuid = id.replace('content-', '');
            window.rendererMap[uuid] = new MarkdownStreamRenderer(container);
        }
    });

    // Rebuild reasoning renderer map - 直接传入 DOM 元素
    reasoningContainers.forEach(container => {
        const id = container.id;
        if (id && id.startsWith('reasoning-content-')) {
            const uuid = id.replace('reasoning-content-', '');
            window.reasoningRendererMap[uuid] = new MarkdownStreamRenderer(container);
        }
    });

    console.log('已重建渲染器映射');
    return true;
};

// Toggle chat message reference visibility
function toggleChatMessageReference(headerElement) {
    const container = headerElement.closest('.chat-message-references');
    if (container) {
        container.classList.toggle('collapsed');
    }
}

// Toggle reference visibility by uuid
function toggleReference(uuid) {
    const ref = document.getElementById('reference-' + uuid);
    if (ref) {
        ref.classList.toggle('collapsed');
    }
}
