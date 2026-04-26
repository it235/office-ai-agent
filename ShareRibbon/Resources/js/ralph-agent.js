/**
 * ralph-agent.js - Ralph Agent 前端控制
 * 类似Cursor的自动化Agent，嵌入聊天流，自动执行步骤
 */

// Agent 状态
window.ralphAgentState = {
    active: false,
    session: null,
    locked: false  // 锁定聊天输入
};

/**
 * 锁定聊天输入（Agent执行期间）
 */
function lockChatInput() {
    window.ralphAgentState.locked = true;
    const smartInput = document.getElementById('smart-input');
    const sendBtn = document.getElementById('send-button');
    const chatInput = document.getElementById('chat-input');
    
    if (smartInput) {
        smartInput.contentEditable = 'false';
        smartInput.classList.add('input-locked');
        smartInput.dataset.placeholder = 'Agent执行中，请等待完成或点击终止...';
    }
    if (chatInput) chatInput.disabled = true;
    if (sendBtn) sendBtn.disabled = true;
}

/**
 * 解锁聊天输入
 */
function unlockChatInput() {
    window.ralphAgentState.locked = false;
    const smartInput = document.getElementById('smart-input');
    const sendBtn = document.getElementById('send-button');
    const chatInput = document.getElementById('chat-input');
    
    if (smartInput) {
        smartInput.contentEditable = 'true';
        smartInput.classList.remove('input-locked');
        smartInput.dataset.placeholder = '请在此输入您的问题... 按Enter键直接发送，Tab采纳补全';
    }
    if (chatInput) chatInput.disabled = false;
    if (sendBtn) sendBtn.disabled = false;
}

/**
 * 检查是否被Agent锁定
 */
function isAgentLocked() {
    return window.ralphAgentState.locked;
}

/**
 * 显示Agent规划卡片（嵌入聊天流）
 * @param {Object} planData - { understanding, steps, summary, sessionId, replaceThinkingUuid }
 */
function showAgentPlanCard(planData) {
    try {
        window.ralphAgentState.active = true;
        window.ralphAgentState.session = planData;
        
        // 锁定输入
        lockChatInput();

        const uuid = planData.sessionId || generateUUID();
        const timestamp = formatDateTime(new Date());

        let chatContainer;
        // 检查是否有需要替换的思考消息
        if (planData.replaceThinkingUuid) {
            const thinkingDiv = document.getElementById('content-' + planData.replaceThinkingUuid);
            const parentContainer = thinkingDiv ? thinkingDiv.closest('.chat-container') : null;
            if (parentContainer) {
                // 找到父容器，直接使用它
                chatContainer = parentContainer;
                chatContainer.className = 'chat-container ralph-agent-container';
                chatContainer.id = 'agent-plan-' + uuid;
                // 清空内容
                chatContainer.innerHTML = '';
            }
        }
        
        // 如果没有找到可替换的容器，创建新的
        if (!chatContainer) {
            chatContainer = document.createElement('div');
            chatContainer.className = 'chat-container ralph-agent-container';
            chatContainer.id = 'agent-plan-' + uuid;
        }

        // 构建步骤HTML
        const stepsHtml = planData.steps ? planData.steps.map((step, idx) => `
            <div class="agent-step" id="agent-step-${uuid}-${idx}" data-status="pending">
                <div class="agent-step-header">
                    <span class="agent-step-icon">⏳</span>
                    <span class="agent-step-num">${idx + 1}</span>
                    <span class="agent-step-desc">${escapeHtml(step.description)}</span>
                </div>
                <div class="agent-step-detail" style="display:none;">
                    <div class="step-detail-text">${escapeHtml(step.detail || '')}</div>
                    <div class="step-code-area" id="step-code-${uuid}-${idx}"></div>
                </div>
            </div>
        `).join('') : '';

        chatContainer.innerHTML = `
            <div class="message-header">
                <div class="avatar-ai">AI</div>
                <div class="sender-info">
                    <div class="sender-name">Ralph Agent <span class="agent-badge">自动执行</span></div>
                    <div class="timestamp">${timestamp}</div>
                </div>
            </div>
            <div class="message-content agent-plan-content">
                <div class="agent-understanding">
                    <strong>📋 理解：</strong>${escapeHtml(planData.understanding || '')}
                </div>
                <div class="agent-steps-container">
                    <div class="agent-steps-header">
                        <span>📝 执行计划</span>
                        <span class="agent-step-count">${planData.steps ? planData.steps.length : 0} 个步骤</span>
                    </div>
                    <div class="agent-steps-list" id="agent-steps-${uuid}">
                        ${stepsHtml}
                    </div>
                </div>
                <div class="agent-summary">
                    <strong>🎯 预期结果：</strong>${escapeHtml(planData.summary || '')}
                </div>
                <div class="agent-actions" id="agent-actions-${uuid}">
                    <button class="agent-btn agent-btn-execute" onclick="confirmAgentExecution('${uuid}')">
                        ▶ 开始执行
                    </button>
                    <button class="agent-btn agent-btn-refine" onclick="refineAgentPlan('${uuid}')">🔄 修改计划</button>
                    <button class="agent-btn agent-btn-abort" onclick="abortAgent('${uuid}')">
                        ✖ 取消
                    </button>
                </div>
            </div>
            <div class="agent-status-bar" id="agent-status-${uuid}">
                <span class="status-icon">⏸</span>
                <span class="status-text">等待确认执行</span>
            </div>
        `;

        // 检查是否已经在容器中（即我们是否替换了思考消息）
        const isAlreadyInContainer = chatContainer.parentElement && chatContainer.parentElement.id === 'chat-container';
        if (!isAlreadyInContainer) {
            // 添加到聊天容器
            const chatHistoryContainer = document.getElementById('chat-container');
            if (chatHistoryContainer) {
                chatHistoryContainer.appendChild(chatContainer);
            }
        }
        
        // 滚动到底部
        chatContainer.scrollIntoView({ behavior: 'smooth', block: 'end' });

        window.ralphAgentState.session.uuid = uuid;
        } catch (err) {
        console.error('showAgentPlanCard error:', err);
        unlockChatInput();
    }
}

/**
 * 确认执行Agent
 */
function confirmAgentExecution(uuid) {
    // Immediately disable both buttons to prevent double-click
    const executeBtn = document.querySelector(`[onclick="confirmAgentExecution('${uuid}')"]`);
    const abortBtn = document.querySelector(`[onclick="abortAgent('${uuid}')"]`);
    if (executeBtn) { executeBtn.disabled = true; executeBtn.textContent = '⏳ 执行中...'; }
    if (abortBtn) { abortBtn.disabled = true; }

    // 隐藏按钮，显示终止按钮
    const actions = document.getElementById('agent-actions-' + uuid);
    if (actions) {
        actions.innerHTML = `
            <button class="agent-btn agent-btn-abort" onclick="abortAgent('${uuid}')">
                ⏹ 终止执行
            </button>
        `;
    }

    updateAgentStatus(uuid, 'running', '正在执行...');

    sendMessageToServer({ type: 'startAgentExecution', sessionId: uuid });
}

/**
 * 终止Agent
 */
function abortAgent(uuid) {
    updateAgentStatus(uuid, 'aborted', '已终止');
    
    // 更新操作按钮区域
    const actions = document.getElementById('agent-actions-' + uuid);
    if (actions) {
        actions.innerHTML = `
            <span class="agent-complete-text">⏹ 已终止</span>
        `;
    }
    
    unlockChatInput();
    window.ralphAgentState.active = false;
    
    // 通知后端
    sendMessageToServer({
        type: 'abortAgent',
        sessionId: uuid
    });
}

/**
 * 请求修改Agent计划
 */
function refineAgentPlan(uuid) {
    const feedback = prompt('请说明对执行计划的修改意见（例如：步骤太多、方式不对、需要先备份等）:');
    if (feedback && feedback.trim()) {
        // Disable all buttons on the plan card
        document.querySelectorAll('#agent-plan-' + uuid + ' .agent-btn').forEach(function(b) { b.disabled = true; });
        sendMessageToServer({ type: 'refineAgentPlan', sessionId: uuid, feedback: feedback.trim() });
    }
}

/**
 * 更新Agent状态
 */
function updateAgentStatus(uuid, status, text) {
    const statusBar = document.getElementById('agent-status-' + uuid);
    if (!statusBar) return;
    
    let icon = '⏸';
    let className = '';
    
    switch(status) {
        case 'running': icon = '▶️'; className = 'status-running'; break;
        case 'completed': icon = '✅'; className = 'status-completed'; break;
        case 'failed': icon = '❌'; className = 'status-failed'; break;
        case 'aborted': icon = '⏹'; className = 'status-aborted'; break;
    }
    
    statusBar.className = 'agent-status-bar ' + className;
    statusBar.innerHTML = `
        <span class="status-icon">${icon}</span>
        <span class="status-text">${escapeHtml(text)}</span>
    `;
}

/**
 * 更新步骤状态
 */
function updateAgentStep(uuid, stepIndex, status, message) {
    const step = document.getElementById('agent-step-' + uuid + '-' + stepIndex);
    if (!step) return;
    
    step.dataset.status = status;
    step.className = 'agent-step agent-step-' + status;
    
    const icon = step.querySelector('.agent-step-icon');
    if (icon) {
        switch(status) {
            case 'running': icon.textContent = '▶️'; break;
            case 'completed': icon.textContent = '✅'; break;
            case 'failed': icon.textContent = '❌'; break;
            case 'skipped': icon.textContent = '⏭️'; break;
            default: icon.textContent = '⏳';
        }
    }
    
    // 展开当前执行的步骤详情
    const detail = step.querySelector('.agent-step-detail');
    if (detail) {
        detail.style.display = (status === 'running') ? 'block' : 'none';
    }
    
    // 如果有消息，显示在详情中
    if (message && status !== 'running') {
        const detailText = step.querySelector('.step-detail-text');
        if (detailText) {
            detailText.innerHTML += `<div class="step-result ${status}">${escapeHtml(message)}</div>`;
        }
    }
    
    // 滚动到当前步骤
    step.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

/**
 * 显示步骤生成的代码
 */
function showStepCode(uuid, stepIndex, code, language) {
    const codeArea = document.getElementById('step-code-' + uuid + '-' + stepIndex);
    if (!codeArea) return;
    
    codeArea.innerHTML = `
        <div class="step-code-header">
            <span>生成的代码 (${language})</span>
        </div>
        <pre><code class="language-${language}">${escapeHtml(code)}</code></pre>
    `;
    
    // 代码高亮
    if (window.hljs) {
        codeArea.querySelectorAll('pre code').forEach((block) => {
            hljs.highlightElement(block);
        });
    }
}

/**
 * Agent完成
 */
function completeAgent(uuid, success, message) {
    updateAgentStatus(uuid, success ? 'completed' : 'failed', message || (success ? '执行完成' : '执行失败'));
    
    // 更新操作按钮
    const actions = document.getElementById('agent-actions-' + uuid);
    if (actions) {
        actions.innerHTML = `
            <span class="agent-complete-text">${success ? '✅ 任务完成' : '❌ 任务失败'}</span>
        `;
    }
    
    // 解锁输入
    unlockChatInput();
    window.ralphAgentState.active = false;
}

/**
 * 显示Agent输入对话框（用于启动Agent）
 */
function showAgentInputDialog() {
    // 如果Agent正在运行，不允许启动新的
    if (window.ralphAgentState.active) {
        alert('Agent正在执行中，请等待完成或终止后再试');
        return;
    }
    
    // 移除已存在的对话框
    hideAgentInputDialog();
    
    const dialog = document.createElement('div');
    dialog.id = 'ralph-agent-dialog';
    dialog.className = 'ralph-agent-dialog';
    dialog.innerHTML = `
        <div class="agent-dialog-overlay" onclick="hideAgentInputDialog()"></div>
        <div class="agent-dialog-content">
            <div class="agent-dialog-header">
                <span class="agent-icon">🤖</span>
                <span>启动 Ralph Agent</span>
                <button class="agent-dialog-close" onclick="hideAgentInputDialog()">×</button>
            </div>
            <div class="agent-dialog-body">
                <p class="agent-dialog-desc">
                    Ralph Agent 会自动获取当前文档/选区内容，分析您的需求，制定执行计划，并自动逐步执行。
                </p>
                <textarea id="agent-request-input" class="agent-request-input" 
                    placeholder="请描述您想要完成的任务...&#10;&#10;例如：将选中的文字改成表格格式，第一列是姓名，第二列是分数" 
                    rows="4"></textarea>
            </div>
            <div class="agent-dialog-actions">
                <button class="agent-btn agent-btn-cancel" onclick="hideAgentInputDialog()">取消</button>
                <button class="agent-btn agent-btn-start" onclick="startAgentFromDialog()">🚀 启动Agent</button>
            </div>
        </div>
    `;
    
    document.body.appendChild(dialog);
    
    // 聚焦输入框
    setTimeout(() => {
        const input = document.getElementById('agent-request-input');
        if (input) input.focus();
    }, 100);
    
    // 回车键启动（Ctrl+Enter）
    const input = document.getElementById('agent-request-input');
    if (input) {
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && e.ctrlKey) {
                e.preventDefault();
                startAgentFromDialog();
            }
        });
    }
}

/**
 * 隐藏Agent输入对话框
 */
function hideAgentInputDialog() {
    const dialog = document.getElementById('ralph-agent-dialog');
    if (dialog) dialog.remove();
}

/**
 * 从对话框启动Agent
 */
function startAgentFromDialog() {
    const input = document.getElementById('agent-request-input');
    if (!input) return;
    
    const request = input.value.trim();
    if (!request) {
        alert('请输入任务描述');
        return;
    }
    
    hideAgentInputDialog();
    
    // 发送启动消息到后端
    sendMessageToServer({
        type: 'startAgent',
        request: request
    });
    }

// 导出函数
window.showAgentInputDialog = showAgentInputDialog;
window.hideAgentInputDialog = hideAgentInputDialog;
window.startAgentFromDialog = startAgentFromDialog;
window.showAgentPlanCard = showAgentPlanCard;
window.confirmAgentExecution = confirmAgentExecution;
window.abortAgent = abortAgent;
window.updateAgentStatus = updateAgentStatus;
window.updateAgentStep = updateAgentStep;
window.showStepCode = showStepCode;
window.completeAgent = completeAgent;
window.lockChatInput = lockChatInput;
window.unlockChatInput = unlockChatInput;
window.isAgentLocked = isAgentLocked;
window.refineAgentPlan = refineAgentPlan;
