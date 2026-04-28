/**
 * agent-card.js - 统一 Agent UI 组件
 * 合并 ralph-agent.js + ralph-loop.js 为单一 AgentCard
 * 支持 ReAct 循环展示：Think → Action → Observation
 */

// 统一 Agent 状态
window.agentCardState = {
    active: false,
    session: null,
    locked: false
};

/**
 * 锁定聊天输入（Agent 执行期间）
 */
function lockChatInput() {
    window.agentCardState.locked = true;
    const smartInput = document.getElementById('smart-input');
    const sendBtn = document.getElementById('send-button');
    const chatInput = document.getElementById('chat-input');

    if (smartInput) {
        smartInput.contentEditable = 'false';
        smartInput.classList.add('input-locked');
        smartInput.dataset.placeholder = 'Agent 执行中，请等待完成或点击终止...';
    }
    if (chatInput) chatInput.disabled = true;
    if (sendBtn) sendBtn.disabled = true;
}

/**
 * 解锁聊天输入
 */
function unlockChatInput() {
    window.agentCardState.locked = false;
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
 * 检查是否被 Agent 锁定
 */
function isAgentLocked() {
    return window.agentCardState.locked;
}

/**
 * 显示 Agent 规划卡片（嵌入聊天流）
 * @param {Object} planData - { sessionId, understanding, steps, summary, replaceThinkingUuid }
 */
function showAgentPlanCard(planData) {
    try {
        window.agentCardState.active = true;
        window.agentCardState.session = planData;

        lockChatInput();

        const uuid = planData.sessionId || generateUUID();
        const timestamp = formatDateTime(new Date());

        let chatContainer;
        if (planData.replaceThinkingUuid) {
            const thinkingDiv = document.getElementById('content-' + planData.replaceThinkingUuid);
            const parentContainer = thinkingDiv ? thinkingDiv.closest('.chat-container') : null;
            if (parentContainer) {
                chatContainer = parentContainer;
                chatContainer.className = 'chat-container agent-card-container';
                chatContainer.id = 'agent-plan-' + uuid;
                chatContainer.innerHTML = '';
            }
        }

        if (!chatContainer) {
            chatContainer = document.createElement('div');
            chatContainer.className = 'chat-container agent-card-container';
            chatContainer.id = 'agent-plan-' + uuid;
        }

        const stepsHtml = planData.steps ? planData.steps.map((step, idx) => `
            <div class="agent-step" id="agent-step-${uuid}-${idx}" data-status="pending">
                <div class="agent-step-header">
                    <span class="agent-step-icon" id="step-icon-${uuid}-${idx}">⏳</span>
                    <span class="agent-step-num">${idx + 1}</span>
                    <span class="agent-step-desc">${escapeHtml(step.description)}</span>
                </div>
                <div class="agent-step-detail" id="step-detail-${uuid}-${idx}" style="display:none;">
                    <div class="step-detail-text">${escapeHtml(step.detail || step.code || '')}</div>
                </div>
                <div class="agent-iteration-area" id="iteration-area-${uuid}-${idx}"></div>
            </div>
        `).join('') : '';

        chatContainer.innerHTML = `
            <div class="message-header">
                <div class="avatar-ai">AI</div>
                <div class="sender-info">
                    <div class="sender-name">Agent <span class="agent-badge">自动执行</span></div>
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

        const isAlreadyInContainer = chatContainer.parentElement && chatContainer.parentElement.id === 'chat-container';
        if (!isAlreadyInContainer) {
            const chatHistoryContainer = document.getElementById('chat-container');
            if (chatHistoryContainer) {
                chatHistoryContainer.appendChild(chatContainer);
            }
        }

        chatContainer.scrollIntoView({ behavior: 'smooth', block: 'end' });

        window.agentCardState.session.uuid = uuid;

        // 如果是简单任务（自动执行），自动点击执行
        if (planData.autoExecute) {
            setTimeout(() => confirmAgentExecution(uuid), 500);
        }
    } catch (err) {
        console.error('showAgentPlanCard error:', err);
        unlockChatInput();
    }
}

/**
 * 更新 Agent 步骤状态
 * @param {string} sessionId - 会话 ID
 * @param {number} stepIndex - 步骤索引
 * @param {string} status - 状态：pending / running / completed / failed / skipped
 * @param {string} message - 可选消息
 */
function updateAgentStep(sessionId, stepIndex, status, message) {
    try {
        const stepEl = document.getElementById(`agent-step-${sessionId}-${stepIndex}`);
        if (!stepEl) return;

        stepEl.setAttribute('data-status', status);

        const iconEl = document.getElementById(`step-icon-${sessionId}-${stepIndex}`);
        if (iconEl) {
            const icons = {
                pending: '⏳',
                running: '🔄',
                completed: '✅',
                failed: '❌',
                skipped: '⏭'
            };
            iconEl.textContent = icons[status] || '⏳';
        }

        if (status === 'running') {
            stepEl.classList.add('step-running');
            const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
            if (detailEl) detailEl.style.display = 'block';
        } else if (status === 'completed') {
            stepEl.classList.remove('step-running');
            stepEl.classList.add('step-completed');
        } else if (status === 'failed') {
            stepEl.classList.remove('step-running');
            stepEl.classList.add('step-failed');
        }

        if (message) {
            const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
            if (detailEl) {
                const msgDiv = document.createElement('div');
                msgDiv.className = `step-message step-msg-${status}`;
                msgDiv.textContent = message;
                detailEl.appendChild(msgDiv);
            }
        }
    } catch (err) {
        console.error('updateAgentStep error:', err);
    }
}

/**
 * 更新 ReAct 迭代（Think → Action → Observation）
 * @param {string} sessionId - 会话 ID
 * @param {Object} iteration - { index, thought, action, observation }
 */
function updateAgentIteration(sessionId, iteration) {
    try {
        if (!iteration) return;

        const areaEl = document.getElementById(`iteration-area-${sessionId}-${iteration.index}`);
        if (!areaEl) {
            // 如果没有精确匹配的步骤区域，放到当前运行的步骤下
            const runningStep = document.querySelector(`#agent-steps-${sessionId} .step-running`);
            if (runningStep) {
                const idx = runningStep.querySelector('.agent-step-num')?.textContent;
                if (idx) {
                    const fallbackArea = document.getElementById(`iteration-area-${sessionId}-${parseInt(idx) - 1}`);
                    if (fallbackArea) fallbackArea.innerHTML = buildIterationHtml(iteration);
                }
            }
            return;
        }

        areaEl.innerHTML = buildIterationHtml(iteration);
    } catch (err) {
        console.error('updateAgentIteration error:', err);
    }
}

/**
 * 构建迭代 HTML
 */
function buildIterationHtml(iteration) {
    const thought = escapeHtml(iteration.thought || '');
    const action = escapeHtml(iteration.action || '');
    const observation = escapeHtml(iteration.observation || '');

    return `
        <div class="react-iteration">
            <div class="iteration-thought">
                <span class="iteration-label">💭 思考</span>
                <div class="iteration-content">${thought}</div>
            </div>
            ${action ? `
            <div class="iteration-action">
                <span class="iteration-label">🔧 行动</span>
                <div class="iteration-content"><code>${action}</code></div>
            </div>` : ''}
            ${observation ? `
            <div class="iteration-observation">
                <span class="iteration-label">👁 观察</span>
                <div class="iteration-content">${observation}</div>
            </div>` : ''}
        </div>
    `;
}

/**
 * 显示审批请求 UI
 * @param {string} sessionId - 会话 ID
 * @param {string} message - 审批提示消息
 */
function showAgentApproval(sessionId, message) {
    try {
        const actionsEl = document.getElementById(`agent-actions-${sessionId}`);
        if (actionsEl) {
            actionsEl.innerHTML = `
                <div class="agent-approval-request">
                    <span class="approval-msg">${escapeHtml(message)}</span>
                    <button class="agent-btn agent-btn-execute" onclick="agentApprove('${sessionId}')">✅ 确认</button>
                    <button class="agent-btn agent-btn-abort" onclick="agentReject('${sessionId}')">❌ 跳过</button>
                </div>
            `;
        }
        updateAgentStatus(sessionId, 'waitingApproval', '等待用户确认...');
    } catch (err) {
        console.error('showAgentApproval error:', err);
    }
}

/**
 * 用户确认执行 Agent
 */
function confirmAgentExecution(uuid) {
    const executeBtn = document.querySelector(`[onclick="confirmAgentExecution('${uuid}')"]`);
    const abortBtn = document.querySelector(`[onclick="abortAgent('${uuid}')"]`);
    if (executeBtn) { executeBtn.disabled = true; executeBtn.textContent = '⏳ 执行中...'; }
    if (abortBtn) { abortBtn.disabled = true; }

    const actions = document.getElementById('agent-actions-' + uuid);
    if (actions) {
        actions.innerHTML = `
            <button class="agent-btn agent-btn-abort" onclick="abortAgent('${uuid}')">
                ⏹ 终止执行
            </button>
        `;
    }

    updateAgentStatus(uuid, 'running', '正在执行...');
    requestApprove(uuid);
}

/**
 * 用户批准当前审批项
 */
function agentApprove(sessionId) {
    requestApprove(sessionId);
}

/**
 * 用户拒绝当前审批项
 */
function agentReject(sessionId) {
    requestReject(sessionId);
}

/**
 * 修改 Agent 计划
 */
function refineAgentPlan(uuid) {
    const feedback = prompt('请输入修改意见：');
    if (feedback && feedback.trim()) {
        requestRefinePlan(uuid, feedback.trim());
    }
}

/**
 * 终止 Agent
 */
function abortAgent(uuid) {
    updateAgentStatus(uuid, 'aborted', '已终止');
    const actions = document.getElementById('agent-actions-' + uuid);
    if (actions) {
        actions.innerHTML = '<span class="agent-terminated">已终止</span>';
    }
    requestAbortAgent();
    unlockChatInput();
}

/**
 * 更新 Agent 状态栏
 * @param {string} uuid - 会话 UUID
 * @param {string} status - 状态
 * @param {string} text - 显示文本
 */
function updateAgentStatus(uuid, status, text) {
    const statusBar = document.getElementById('agent-status-' + uuid);
    if (!statusBar) return;

    const icons = {
        running: '🔄',
        waitingApproval: '⏸',
        completed: '✅',
        failed: '❌',
        aborted: '⏹',
        paused: '⏸'
    };

    statusBar.innerHTML = `
        <span class="status-icon">${icons[status] || '⏳'}</span>
        <span class="status-text">${escapeHtml(text || '')}</span>
    `;
}

/**
 * 完成 Agent（成功或失败）
 * @param {string} uuid - 会话 UUID
 * @param {boolean} success - 是否成功
 * @param {string} message - 完成消息
 */
function completeAgent(uuid, success, message) {
    try {
        const statusBar = document.getElementById('agent-status-' + uuid);
        if (statusBar) {
            const icon = success ? '✅' : '❌';
            const text = success ? '任务完成' : '任务失败';
            statusBar.innerHTML = `
                <span class="status-icon">${icon}</span>
                <span class="status-text">${text}${message ? ': ' + escapeHtml(message) : ''}</span>
            `;
        }

        const actions = document.getElementById('agent-actions-' + uuid);
        if (actions) {
            actions.innerHTML = `<span class="agent-finished">${success ? '✅ 已完成' : '❌ 已失败'}</span>`;
        }

        window.agentCardState.active = false;
        window.agentCardState.session = null;
        unlockChatInput();
    } catch (err) {
        console.error('completeAgent error:', err);
        unlockChatInput();
    }
}

/**
 * 获取步骤状态图标
 */
function getStepIcon(status) {
    const icons = {
        pending: '⏳',
        running: '🔄',
        completed: '✅',
        failed: '❌',
        skipped: '⏭'
    };
    return icons[status] || '⏳';
}

/**
 * 获取状态文本
 */
function getStatusText(status) {
    const texts = {
        planning: '规划中...',
        ready: '准备就绪',
        running: '执行中...',
        paused: '已暂停',
        completed: '已完成',
        failed: '失败',
        aborted: '已终止'
    };
    return texts[status] || status;
}

// 导出到全局
window.agentCardState = window.agentCardState;
window.showAgentPlanCard = showAgentPlanCard;
window.updateAgentStep = updateAgentStep;
window.updateAgentIteration = updateAgentIteration;
window.showAgentApproval = showAgentApproval;
window.confirmAgentExecution = confirmAgentExecution;
window.agentApprove = agentApprove;
window.agentReject = agentReject;
window.refineAgentPlan = refineAgentPlan;
window.abortAgent = abortAgent;
window.updateAgentStatus = updateAgentStatus;
window.completeAgent = completeAgent;
window.lockChatInput = lockChatInput;
window.unlockChatInput = unlockChatInput;
window.isAgentLocked = isAgentLocked;
