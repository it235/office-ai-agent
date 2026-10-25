/**
 * agent-card.js - 统一 Agent UI 组件
 * 统一展示 Agent 计划、执行步骤、审批和执行解释
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
 * 恢复一次 Agent 请求占用的全部输入 UI。
 * 输入锁和发送/终止按钮由不同模块管理，Agent 进入终态时必须同时复位。
 */
function restoreAgentRequestUi() {
    unlockChatInput();

    if (typeof changeSendButton === 'function') {
        changeSendButton();
        return;
    }

    // 脚本异常或加载顺序变化时的兜底，避免终止按钮永久停留。
    const sendButton = document.getElementById('send-button');
    const stopButton = document.getElementById('stop-button');
    if (sendButton) sendButton.style.setProperty('display', 'flex', 'important');
    if (stopButton) stopButton.style.setProperty('display', 'none', 'important');
    if (typeof hideLoadingIndicator === 'function') hideLoadingIndicator();
}

/**
 * 检查是否被 Agent 锁定
 */
function isAgentLocked() {
    return window.agentCardState.locked;
}

/**
 * 显示 Agent 规划卡片（嵌入聊天流）
 * @param {Object} planData - { sessionId, understanding, steps, summary, planVersion, scope, stepCount, toolCount, riskLevel, riskyTools, replaceThinkingUuid }
 */
function showAgentPlanCard(planData) {
    try {
        window.agentCardState.active = true;
        window.agentCardState.session = planData;

        lockChatInput();

        const uuid = planData.sessionId || generateUUID();
        const timestamp = formatDateTime(new Date());
        const planVersion = Number(planData.planVersion) || 1;

        // 重新规划会对同一会话再发一次计划：原位重绘，否则时间线一直停留在已经作废的旧计划上。
        const existingContainer = document.getElementById('agent-plan-' + uuid);
        const isReplan = planVersion > 1 && !!existingContainer;

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
            chatContainer = isReplan ? existingContainer : document.createElement('div');
            if (!isReplan) {
                chatContainer.className = 'chat-container agent-card-container';
                chatContainer.id = 'agent-plan-' + uuid;
            }
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
                ${buildPlanTrustHtml(planData, planVersion)}
                <div class="agent-steps-container">
                    <div class="agent-steps-header">
                        <span>📝 执行计划${planVersion > 1 ? ` <span class="agent-plan-version">第 ${planVersion} 版</span>` : ''}</span>
                        <span class="agent-step-count">${planData.steps ? planData.steps.length : 0} 个步骤</span>
                    </div>
                    <div class="agent-steps-list" id="agent-steps-${uuid}">
                        ${stepsHtml}
                    </div>
                </div>
                <div class="agent-summary">
                    <strong>🎯 预期结果：</strong>${escapeHtml(planData.summary || '')}
                </div>
                <div class="agent-actions" id="agent-actions-${uuid}"></div>
            </div>
            <div class="agent-status-bar" id="agent-status-${uuid}"></div>
        `;

        const actionsEl = document.getElementById('agent-actions-' + uuid);
        if (isReplan) {
            // 新一轮计划已经在执行中，不能把卡片退回“等待确认执行”状态
            if (actionsEl) {
                actionsEl.innerHTML = `
                    <button class="agent-btn agent-btn-abort" onclick="abortAgent('${uuid}')">
                        ⏹ 终止执行
                    </button>
                `;
            }
            updateAgentStatus(uuid, 'running', '按新计划继续执行...');
        } else {
            if (actionsEl) {
                actionsEl.innerHTML = `
                    <button class="agent-btn agent-btn-execute" onclick="confirmAgentExecution('${uuid}')">
                        ▶ 开始执行
                    </button>
                    <button class="agent-btn agent-btn-refine" onclick="refineAgentPlan('${uuid}')">🔄 修改计划</button>
                    <button class="agent-btn agent-btn-abort" onclick="abortAgent('${uuid}')">
                        ✖ 取消
                    </button>
                `;
            }
            updateAgentStatus(uuid, 'paused', '等待确认执行');
        }

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
 * 计划卡的信任信息：范围/数量/风险。三者都由后端按上下文与工具注册表算出，
 * 不让模型自述“我要改哪一段”，避免卡片上的范围和实际写入对不上。
 */
function buildPlanTrustHtml(planData, planVersion) {
    const scope = planData.scope || '';
    const stepCount = Number(planData.stepCount) || (planData.steps ? planData.steps.length : 0);
    const toolCount = Number(planData.toolCount) || 0;
    const riskLevel = planData.riskLevel || 'safe';
    const riskyTools = Array.isArray(planData.riskyTools) ? planData.riskyTools.filter(Boolean) : [];

    let riskText;
    if (riskLevel === 'safe') {
        riskText = '风险：安全操作';
    } else if (riskyTools.length > 0) {
        const shown = riskyTools.slice(0, 3).join('、');
        riskText = `风险：需确认 ${riskyTools.length} 项（${shown}${riskyTools.length > 3 ? ' 等' : ''}）`;
    } else {
        riskText = `风险：${riskLevel}`;
    }

    const chips = [];
    if (scope) chips.push(`<span class="agent-trust-chip agent-trust-scope">范围 ${escapeHtml(scope)}</span>`);
    chips.push(`<span class="agent-trust-chip">数量 ${stepCount} 步 · ${toolCount} 个工具</span>`);
    chips.push(`<span class="agent-trust-chip agent-trust-risk-${escapeHtml(riskLevel)}">${escapeHtml(riskText)}</span>`);
    if (planVersion > 1) chips.push(`<span class="agent-trust-chip">第 ${planVersion} 版计划</span>`);
    return `<div class="agent-plan-trust">${chips.join('')}</div>`;
}

/**
 * 更新 Agent 步骤状态
 * @param {string} sessionId - 会话 ID
 * @param {number} stepIndex - 步骤索引
 * @param {string} status - 状态：pending / running / completed / failed / skipped / aborted
 * @param {string} message - 可选消息
 */
function updateAgentStep(sessionId, stepIndex, status, message) {
    try {
        const stepEl = document.getElementById(`agent-step-${sessionId}-${stepIndex}`);
        if (!stepEl) return;

        stepEl.setAttribute('data-status', status);

        const iconEl = document.getElementById(`step-icon-${sessionId}-${stepIndex}`);
        if (iconEl) iconEl.textContent = getAgentStepIcon(status);

        // 重规划、重试都会让同一步骤换状态；先清掉旧类，否则失败样式会一直挂在已完成的步骤上
        stepEl.classList.remove('step-running', 'step-completed', 'step-failed', 'step-skipped');
        if (status === 'running') {
            stepEl.classList.add('step-running');
            const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
            if (detailEl) detailEl.style.display = 'block';
        } else if (status === 'completed') {
            stepEl.classList.add('step-completed');
        } else if (status === 'failed') {
            stepEl.classList.add('step-failed');
        } else if (status === 'skipped' || status === 'aborted') {
            stepEl.classList.add('step-skipped');
        }

        if (message) {
            const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
            if (detailEl) {
                const msgDiv = document.createElement('div');
                msgDiv.className = `step-message step-msg-${status}`;
                msgDiv.textContent = message;
                detailEl.appendChild(msgDiv);
                detailEl.style.display = 'block';
            }
        }
    } catch (err) {
        console.error('updateAgentStep error:', err);
    }
}

/**
 * 更新 ReAct 迭代（Think → Action → Observation）
 * @param {string} sessionId - 会话 ID
 * @param {Object} iteration - { index, stepIndex, thought, action, observation }
 */
function updateAgentIteration(sessionId, iteration) {
    try {
        if (!iteration) return;

        // 迭代详情属于某个步骤：优先用后端给的步骤索引。全局迭代号和步骤号在重规划后必然错开，
        // 只按 index 找会把详情写进别的步骤，甚至找不到位置。
        const stepIndex = Number.isInteger(iteration.stepIndex) && iteration.stepIndex >= 0
            ? iteration.stepIndex
            : iteration.index;

        const areaEl = document.getElementById(`iteration-area-${sessionId}-${stepIndex}`);
        if (!areaEl) {
            // 兜底：仍找不到位置时写到当前正在执行的步骤下，而不是丢弃这次观察
            const runningStep = document.querySelector(`#agent-steps-${sessionId} .step-running`);
            const stepNum = runningStep ? runningStep.querySelector('.agent-step-num')?.textContent : '';
            const fallbackArea = document.getElementById(`iteration-area-${sessionId}-${parseInt(stepNum, 10) - 1}`);
            if (fallbackArea) fallbackArea.innerHTML = buildIterationHtml(iteration);
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
    const explanation = iteration.explanation || null;
    const explanationText = explanation ? escapeHtml(explanation.ExplanationText || explanation.explanationText || '') : '';

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
            ${explanationText ? `
            <div class="iteration-explanation">
                <span class="iteration-label">执行解释</span>
                <div class="iteration-content">${explanationText}</div>
            </div>` : ''}
        </div>
    `;
}

/**
 * 显示单步执行解释
 * @param {string} sessionId - 会话 ID
 * @param {Object} explanation - ExecutionExplanation
 */
function showAgentExecutionExplanation(sessionId, explanation) {
    try {
        if (!explanation) return;
        const stepIndex = explanation.StepIndex ?? explanation.stepIndex;
        const text = explanation.ExplanationText || explanation.explanationText || explanation.Message || explanation.message || '';
        if (stepIndex === undefined || !text) return;

        const detailEl = document.getElementById(`step-detail-${sessionId}-${stepIndex}`);
        if (!detailEl) return;

        const explanationEl = document.createElement('details');
        explanationEl.className = 'step-explanation';
        explanationEl.open = false;

        const toolId = explanation.ToolId || explanation.toolId || '';
        const category = explanation.ToolCategory || explanation.toolCategory || '';
        const skillName = explanation.SkillName || explanation.skillName || '';
        const scriptFileName = explanation.ScriptFileName || explanation.scriptFileName || '';
        const mcpToolName = explanation.McpToolName || explanation.mcpToolName || '';
        const mcpStatus = explanation.McpStatus || explanation.mcpStatus || '';
        const failureReason = explanation.FailureReason || explanation.failureReason || '';
        const risk = explanation.RiskLevel || explanation.riskLevel || '';
        const paramsJson = explanation.ParametersJson || explanation.parametersJson || '';
        const fixed = explanation.FixAttempts || explanation.fixAttempts || 0;
        const elapsedMs = explanation.ElapsedMs ?? explanation.elapsedMs ?? 0;
        const beforeSummary = explanation.BeforeSummary || explanation.beforeSummary || '';
        const afterSummary = explanation.AfterSummary || explanation.afterSummary || '';
        const undoPointName = explanation.UndoPointName || explanation.undoPointName || '';
        const undoHint = explanation.UndoHint || explanation.undoHint || '';
        const canUndo = explanation.CanUndo ?? explanation.canUndo;
        const autoRepairSummary = explanation.AutoRepairSummary || explanation.autoRepairSummary || '';

        explanationEl.innerHTML = `
            <summary>${escapeHtml(text)}</summary>
            <div class="step-explanation-meta">
                ${toolId ? `<div><strong>工具</strong> <code>${escapeHtml(toolId)}</code></div>` : ''}
                ${category ? `<div><strong>类别</strong> ${escapeHtml(category)}</div>` : ''}
                ${elapsedMs ? `<div><strong>耗时</strong> ${Number(elapsedMs).toLocaleString()} ms</div>` : ''}
                ${skillName ? `<div><strong>Skill</strong> ${escapeHtml(skillName)}</div>` : ''}
                ${scriptFileName ? `<div><strong>脚本</strong> <code>${escapeHtml(scriptFileName)}</code></div>` : ''}
                ${mcpToolName ? `<div><strong>MCP</strong> <code>${escapeHtml(mcpToolName)}</code>${mcpStatus ? ' ' + escapeHtml(mcpStatus) : ''}</div>` : ''}
                ${risk ? `<div><strong>风险</strong> ${escapeHtml(risk)}</div>` : ''}
                ${beforeSummary ? `<div><strong>执行前</strong> ${escapeHtml(beforeSummary)}</div>` : ''}
                ${afterSummary ? `<div><strong>执行后</strong> ${escapeHtml(afterSummary)}</div>` : ''}
                ${fixed ? `<div><strong>自动修复</strong> ${fixed} 次</div>` : ''}
                ${autoRepairSummary ? `<div><strong>修复结果</strong> ${escapeHtml(autoRepairSummary)}</div>` : ''}
                ${undoPointName ? `<div><strong>撤销点</strong> ${escapeHtml(undoPointName)}${canUndo === false ? ' <span>(需手动撤销)</span>' : ''}</div>` : ''}
                ${undoHint ? `<div><strong>撤销提示</strong> ${escapeHtml(undoHint)}</div>` : ''}
                ${failureReason ? `<div><strong>失败原因</strong> ${escapeHtml(failureReason)}</div>` : ''}
                ${paramsJson ? `<pre>${escapeHtml(paramsJson)}</pre>` : ''}
            </div>
        `;
        detailEl.appendChild(explanationEl);
    } catch (err) {
        console.error('showAgentExecutionExplanation error:', err);
    }
}

/**
 * 显示审批请求 UI
 * @param {string} sessionId - 会话 ID
 * @param {string} toolId - 待审批工具 ID
 * @param {string} message - 审批提示消息
 * @param {string} riskLevel - 风险级别
 */
function showAgentApproval(sessionId, toolId, message, riskLevel) {
    try {
        const actionsEl = document.getElementById(`agent-actions-${sessionId}`);
        if (actionsEl) {
            const safeToolId = toolId ? escapeHtml(toolId) : '';
            const safeRisk = riskLevel ? escapeHtml(riskLevel) : '';
            const toolText = safeToolId ? `<span class="approval-tool"><code>${safeToolId}</code></span>` : '';
            const riskText = safeRisk ? `<span class="approval-risk">${safeRisk}</span>` : '';
            actionsEl.innerHTML = `
                <div class="agent-approval-request">
                    <span class="approval-msg">${escapeHtml(message)}${toolText}${riskText}</span>
                    <button class="agent-btn agent-btn-execute" title="只放行这一次操作" onclick="agentApprove('${sessionId}')">允许一次</button>
                    <button class="agent-btn agent-btn-approve-session" title="本轮任务内该工具不再询问，任务结束后失效" onclick="agentApproveSession('${sessionId}', '${toolId || ''}')">本会话允许此工具</button>
                    <button class="agent-btn agent-btn-abort" title="只跳过当前这一步，后续步骤继续" onclick="agentReject('${sessionId}')">拒绝</button>
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
 * 用户允许本会话内该工具的所有调用
 */
function agentApproveSession(sessionId, toolId) {
    if (typeof requestApproveToolSession === 'function') {
        requestApproveToolSession(sessionId, toolId || '');
    } else {
        requestApprove(sessionId);
    }
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
    restoreAgentRequestUi();
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
 * @param {string} thinkingUuid - 思考消息 UUID
 * @param {Object} terminal - { writtenCount, rolledBackCount, changeTruncated, changes[], canUndo, runId, units, undoReason }
 */
function completeAgent(uuid, success, message, thinkingUuid, terminal) {
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

        renderAgentChangeSummary(uuid, terminal);

        const actions = document.getElementById('agent-actions-' + uuid);
        if (actions) {
            actions.innerHTML = '';
            const state = document.createElement('span');
            state.className = 'agent-finished';
            state.textContent = success ? '✅ 已完成' : '❌ 已失败';
            actions.appendChild(state);

            if (terminal && terminal.canUndo && terminal.runId) {
                const undoBtn = document.createElement('button');
                undoBtn.type = 'button';
                undoBtn.className = 'agent-btn agent-btn-undo';
                undoBtn.dataset.runId = terminal.runId;
                const units = Number(terminal.units) || 0;
                const restores = Number(terminal.restoreCount) || 0;
                if (units > 0) {
                    undoBtn.textContent = `撤销本轮（${units} 个撤销单元）`;
                } else if (restores > 0) {
                    // Excel/PPT 没有整栈撤销：按钮按写入前的目标区域快照逐步还原，说清楚还原几处
                    undoBtn.textContent = `撤销本轮（按快照还原 ${restores} 处）`;
                } else {
                    undoBtn.textContent = '撤销本轮';
                }
                undoBtn.addEventListener('click', function () { undoAgentRun(undoBtn); });
                actions.appendChild(undoBtn);
            } else if (terminal && terminal.undoReason) {
                // 不给按钮就必须说清原因，否则用户会以为这些改动随时能一键回退
                const reason = document.createElement('span');
                reason.className = 'agent-undo-reason';
                reason.textContent = terminal.undoReason;
                actions.appendChild(reason);
            }
        }

        // 没有生成 plan card 时，状态仍显示在最初的 thinking 消息中。
        if (thinkingUuid) {
            const thinkingDiv = document.getElementById('content-' + thinkingUuid);
            if (thinkingDiv) {
                const icon = success ? '✅' : '❌';
                const text = success ? '任务完成' : '任务失败';
                thinkingDiv.innerHTML = `<div class="agent-terminal-message ${success ? 'success' : 'failed'}">` +
                    `<span class="status-icon">${icon}</span>` +
                    `<span class="status-text">${text}${message ? ': ' + escapeHtml(message) : ''}</span>` +
                    `</div>`;
            }
        }

        window.agentCardState.active = false;
        window.agentCardState.session = null;
    } catch (err) {
        console.error('completeAgent error:', err);
    } finally {
        restoreAgentRequestUi();
    }
}

/**
 * 终态变更摘要：把这一轮真正留在文档里的改动列成清单，作为“撤销本轮”的事实依据。
 * 全部用 textContent 写入，后端摘要里若带 HTML 也只当纯文本。
 */
function renderAgentChangeSummary(uuid, terminal) {
    if (!terminal) return;
    const actions = document.getElementById('agent-actions-' + uuid);
    let box = document.getElementById('agent-change-summary-' + uuid);
    if (!box && actions && actions.parentNode) {
        box = document.createElement('div');
        box.id = 'agent-change-summary-' + uuid;
        box.className = 'agent-change-summary';
        actions.parentNode.insertBefore(box, actions);
    }
    if (!box) return;

    box.innerHTML = '';
    const written = Number(terminal.writtenCount) || 0;
    const title = document.createElement('div');
    title.className = 'agent-change-title';
    title.textContent = written > 0 ? `🧾 本轮已写入 ${written} 处` : '🧾 本轮没有留在文档里的改动';
    box.appendChild(title);

    const changes = Array.isArray(terminal.changes) ? terminal.changes : [];
    if (changes.length > 0) {
        const list = document.createElement('ol');
        list.className = 'agent-change-list';
        changes.forEach(function (item) {
            const li = document.createElement('li');
            if (item && item.toolId) {
                const tool = document.createElement('code');
                tool.className = 'agent-change-tool';
                tool.textContent = item.toolId;
                li.appendChild(tool);
            }
            const text = document.createElement('span');
            text.textContent = (item && item.summary) ? item.summary : '（宿主未返回变更说明）';
            li.appendChild(text);
            if (item && item.partialApply) {
                const tag = document.createElement('em');
                tag.className = 'agent-change-partial';
                tag.textContent = '部分生效';
                li.appendChild(tag);
            }
            list.appendChild(li);
        });
        box.appendChild(list);
    }

    if (terminal.changeTruncated) {
        const more = document.createElement('div');
        more.className = 'agent-change-more';
        more.textContent = `另有 ${Math.max(0, written - changes.length)} 处改动未列出，展开步骤可查看执行记录。`;
        box.appendChild(more);
    }

    const rolledBack = Number(terminal.rolledBackCount) || 0;
    if (rolledBack > 0) {
        const note = document.createElement('div');
        note.className = 'agent-change-rollback';
        note.textContent = `已自动回滚 ${rolledBack} 处未完成的改动。`;
        box.appendChild(note);
    }

    // E3-04：本轮用量。整句（token 数 / 金额 / 超额度提示）由 UsageCostReporter 拼好，
    // 前端只显示——金额在 JS 里再算一遍就会出现两套口径，未配置单价时还会各自编一个数。
    const usage = terminal.usage;
    if (usage && usage.summary) {
        const usageLine = document.createElement('div');
        usageLine.className = 'agent-usage';
        usageLine.textContent = '📊 ' + usage.summary;
        box.appendChild(usageLine);
    }
}

/**
 * 请求撤销某一轮 Agent 写入
 * @param {HTMLButtonElement} button - 触发撤销的按钮
 */
function undoAgentRun(button) {
    if (!button || button.disabled) return;
    const runId = button.dataset.runId || '';
    if (!runId) return;

    const errorEl = button.parentNode ? button.parentNode.querySelector('.agent-undo-error') : null;
    if (errorEl) errorEl.remove();

    button.disabled = true;
    button.textContent = '正在撤销...';

    if (typeof sendMessageToServer === 'function') {
        sendMessageToServer({ type: 'agent:undoRun', runId: runId });
    } else {
        button.disabled = false;
        button.textContent = '撤销本轮';
        console.error('[AgentUndo] sendMessageToServer 不可用');
    }
}

/**
 * 回写撤销结果；只改动对应 runId 的按钮，不影响其他轮次。
 * @param {string} runId - 运行标识
 * @param {boolean} success - 是否撤销成功
 * @param {string} message - 结果说明
 * @param {boolean} verified - 是否已核对文档指纹
 */
function setAgentUndoResult(runId, success, message, verified) {
    try {
        if (!runId) return;
        const buttons = document.querySelectorAll('.agent-btn-undo');
        let target = null;
        for (let i = 0; i < buttons.length; i++) {
            if (buttons[i].dataset.runId === runId) { target = buttons[i]; break; }
        }
        if (!target) return;

        if (success) {
            const done = document.createElement('span');
            done.className = 'agent-undo-result';
            done.textContent = verified ? '↩ 已撤销本轮' : '↩ 已撤销';
            done.title = message || '';
            target.parentNode.replaceChild(done, target);
            return;
        }

        target.disabled = false;
        target.textContent = '撤销本轮';
        target.title = message || '撤销失败';
        if (message) {
            const errorEl = document.createElement('span');
            errorEl.className = 'agent-undo-error';
            errorEl.textContent = message;
            target.parentNode.insertBefore(errorEl, target.nextSibling);
        }
    } catch (err) {
        console.error('setAgentUndoResult error:', err);
    }
}

/**
 * 获取步骤状态图标
 */
function getAgentStepIcon(status) {
    const icons = {
        pending: '⏳',
        running: '🔄',
        completed: '✅',
        failed: '❌',
        skipped: '⏭',
        aborted: '⏹'
    };
    return icons[status] || '⏳';
}

// 导出到全局
window.agentCardState = window.agentCardState;
window.showAgentPlanCard = showAgentPlanCard;
window.updateAgentStep = updateAgentStep;
window.updateAgentIteration = updateAgentIteration;
window.showAgentExecutionExplanation = showAgentExecutionExplanation;
window.showAgentApproval = showAgentApproval;
window.confirmAgentExecution = confirmAgentExecution;
window.agentApprove = agentApprove;
window.agentApproveSession = agentApproveSession;
window.agentReject = agentReject;
window.refineAgentPlan = refineAgentPlan;
window.abortAgent = abortAgent;
window.updateAgentStatus = updateAgentStatus;
window.completeAgent = completeAgent;
window.undoAgentRun = undoAgentRun;
window.setAgentUndoResult = setAgentUndoResult;
window.lockChatInput = lockChatInput;
window.unlockChatInput = unlockChatInput;
window.restoreAgentRequestUi = restoreAgentRequestUi;
window.isAgentLocked = isAgentLocked;
