/**
 * ralph-loop.js - Ralph Loop 前端控制
 * 管理循环任务的UI展示和用户交互
 */

// Ralph Loop 状态
window.ralphLoopState = {
    active: false,
    currentSession: null
};

/**
 * 显示循环任务输入对话框
 */
function showLoopInputDialog() {
    try {
        // 如果已有对话框，先移除
        hideLoopInputDialog();
        
        const dialog = document.createElement('div');
        dialog.id = 'ralph-loop-dialog';
        dialog.className = 'ralph-loop-dialog';
        dialog.innerHTML = `
            <div class="loop-dialog-overlay" onclick="hideLoopInputDialog()"></div>
            <div class="loop-dialog-content">
                <div class="loop-dialog-header">
                    <span class="loop-icon">🔄</span>
                    <span>启动 Ralph Loop</span>
                    <button class="loop-dialog-close" onclick="hideLoopInputDialog()">×</button>
                </div>
                <div class="loop-dialog-body">
                    <p class="loop-dialog-desc">请描述您想要完成的任务目标，AI将自动规划并分步执行。</p>
                    <textarea id="loop-goal-input" class="loop-goal-input" placeholder="例如：分析当前工作表数据并生成销售趋势图表" rows="3"></textarea>
                </div>
                <div class="loop-dialog-actions">
                    <button class="loop-btn loop-btn-cancel" onclick="hideLoopInputDialog()">取消</button>
                    <button class="loop-btn loop-btn-start" onclick="startLoopFromDialog()">开始规划</button>
                </div>
            </div>
        `;
        
        document.body.appendChild(dialog);
        
        // 聚焦输入框
        setTimeout(() => {
            const input = document.getElementById('loop-goal-input');
            if (input) input.focus();
        }, 100);
        
        // 回车键启动
        const input = document.getElementById('loop-goal-input');
        if (input) {
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    startLoopFromDialog();
                }
            });
        }
    } catch (err) {
        console.error('showLoopInputDialog error:', err);
    }
}

/**
 * 隐藏输入对话框
 */
function hideLoopInputDialog() {
    const dialog = document.getElementById('ralph-loop-dialog');
    if (dialog) dialog.remove();
}

/**
 * 从对话框启动循环
 */
function startLoopFromDialog() {
    const input = document.getElementById('loop-goal-input');
    if (!input) return;
    
    const goal = input.value.trim();
    if (!goal) {
        alert('请输入任务目标');
        return;
    }
    
    hideLoopInputDialog();
    
    // 发送启动消息到后端
    sendMessageToServer({
        type: 'startLoop',
        goal: goal
    });
    console.log('[RalphLoop] 启动循环，目标:', goal);
}

/**
 * 显示循环规划卡片 - 固定在顶部
 * @param {Object} loopData - { goal, steps, status }
 */
function showLoopPlanCard(loopData) {
    try {
        // 移除已存在的卡片
        hideLoopPlanCard();
        
        window.ralphLoopState.active = true;
        window.ralphLoopState.currentSession = loopData;

        const card = document.createElement('div');
        card.id = 'ralph-loop-card';
        card.className = 'ralph-loop-card ralph-loop-fixed';

        const stepsHtml = loopData.steps ? loopData.steps.map((step, idx) => {
            const statusIcon = getStepIcon(step.status);
            const statusClass = step.status === 'running' ? 'step-running' : 
                               step.status === 'completed' ? 'step-completed' : 
                               step.status === 'failed' ? 'step-failed' : '';
            return `
                <div class="loop-step ${statusClass}" id="loop-step-${idx}">
                    <span class="step-icon">${statusIcon}</span>
                    <span class="step-num">${idx + 1}</span>
                    <span class="step-desc">${escapeHtml(step.description)}</span>
                </div>
            `;
        }).join('') : '';

        card.innerHTML = `
            <div class="loop-header">
                <span class="loop-icon">🔄</span>
                <span class="loop-title">Ralph Loop - 任务规划</span>
                <button class="loop-minimize-btn" onclick="toggleLoopCard()" title="最小化">−</button>
                <button class="loop-close-btn" onclick="cancelLoop()" title="关闭">×</button>
            </div>
            <div class="loop-body">
                <div class="loop-goal">
                    <strong>目标:</strong> ${escapeHtml(loopData.goal || '')}
                </div>
                <div class="loop-steps">
                    ${stepsHtml}
                </div>
                <div class="loop-status">
                    状态: <span class="status-text">${getStatusText(loopData.status)}</span>
                </div>
                <div class="loop-actions">
                    <button class="loop-btn loop-btn-continue" onclick="continueLoop()" ${loopData.status !== 'paused' && loopData.status !== 'ready' ? 'disabled' : ''}>
                        ▶ 继续执行
                    </button>
                    <button class="loop-btn loop-btn-cancel" onclick="cancelLoop()">
                        ✖ 取消
                    </button>
                </div>
            </div>
        `;

        // 固定在case-container之后
        const caseContainer = document.getElementById('case-container');
        if (caseContainer && caseContainer.parentNode) {
            caseContainer.parentNode.insertBefore(card, caseContainer.nextSibling);
        } else {
            document.body.insertBefore(card, document.body.firstChild);
        }

    } catch (err) {
        console.error('showLoopPlanCard error:', err);
    }
}

/**
 * 切换卡片最小化状态
 */
function toggleLoopCard() {
    const card = document.getElementById('ralph-loop-card');
    if (!card) return;
    
    card.classList.toggle('minimized');
    const btn = card.querySelector('.loop-minimize-btn');
    if (btn) {
        btn.textContent = card.classList.contains('minimized') ? '+' : '−';
        btn.title = card.classList.contains('minimized') ? '展开' : '最小化';
    }
}

/**
 * 更新循环步骤状态
 * @param {number} stepIndex - 步骤索引
 * @param {string} status - 状态
 */
function updateLoopStep(stepIndex, status) {
    const card = document.getElementById('ralph-loop-card');
    if (!card) return;

    const steps = card.querySelectorAll('.loop-step');
    if (steps[stepIndex]) {
        const step = steps[stepIndex];
        step.className = 'loop-step';
        if (status === 'running') step.classList.add('step-running');
        if (status === 'completed') step.classList.add('step-completed');
        if (status === 'failed') step.classList.add('step-failed');
        
        const iconEl = step.querySelector('.step-icon');
        if (iconEl) iconEl.textContent = getStepIcon(status);
        
        // 滚动到当前步骤
        step.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    if (window.ralphLoopState.currentSession && window.ralphLoopState.currentSession.steps) {
        window.ralphLoopState.currentSession.steps[stepIndex].status = status;
    }
}

/**
 * 更新循环整体状态
 * @param {string} status - 状态
 */
function updateLoopStatus(status) {
    const card = document.getElementById('ralph-loop-card');
    if (!card) return;

    const statusText = card.querySelector('.status-text');
    if (statusText) statusText.textContent = getStatusText(status);

    const continueBtn = card.querySelector('.loop-btn-continue');
    if (continueBtn) {
        continueBtn.disabled = (status !== 'paused' && status !== 'ready');
    }

    if (window.ralphLoopState.currentSession) {
        window.ralphLoopState.currentSession.status = status;
    }

    // 完成时显示完成消息
    if (status === 'completed') {
        setTimeout(() => {
            const header = card.querySelector('.loop-header');
            if (header) {
                header.innerHTML = `
                    <span class="loop-icon">✅</span>
                    <span class="loop-title">Ralph Loop - 任务完成</span>
                    <button class="loop-close-btn" onclick="hideLoopPlanCard()" title="关闭">×</button>
                `;
            }
        }, 500);
    }
}

/**
 * 隐藏循环卡片
 */
function hideLoopPlanCard() {
    const card = document.getElementById('ralph-loop-card');
    if (card) card.remove();
    window.ralphLoopState.active = false;
}

/**
 * 继续执行循环
 */
function continueLoop() {
    sendMessageToServer({
        type: 'continueLoop'
    });
}

/**
 * 取消循环
 */
function cancelLoop() {
    hideLoopPlanCard();
    sendMessageToServer({
        type: 'cancelLoop'
    });
}

/**
 * 获取步骤图标
 */
function getStepIcon(status) {
    switch(status) {
        case 'pending': return '⏳';
        case 'running': return '▶️';
        case 'completed': return '✅';
        case 'failed': return '❌';
        case 'skipped': return '⏭️';
        default: return '❓';
    }
}

/**
 * 获取状态文本
 */
function getStatusText(status) {
    switch(status) {
        case 'planning': return '规划中';
        case 'ready': return '准备执行';
        case 'running': return '执行中';
        case 'paused': return '等待继续';
        case 'completed': return '已完成';
        case 'failed': return '失败';
        default: return '未知';
    }
}

// 导出函数
window.showLoopInputDialog = showLoopInputDialog;
window.hideLoopInputDialog = hideLoopInputDialog;
window.startLoopFromDialog = startLoopFromDialog;
window.showLoopPlanCard = showLoopPlanCard;
window.toggleLoopCard = toggleLoopCard;
window.updateLoopStep = updateLoopStep;
window.updateLoopStatus = updateLoopStatus;
window.hideLoopPlanCard = hideLoopPlanCard;
window.continueLoop = continueLoop;
window.cancelLoop = cancelLoop;
