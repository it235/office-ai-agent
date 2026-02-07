/**
 * intent-preview.js - 意图预览组件
 * 显示"我理解您想要..."的预览卡片，用户确认后再发送
 */

// 意图预览状态
window.intentPreviewState = {
    active: false,
    currentIntent: null,
    pendingMessage: null,
    autoConfirm: false  // Agent模式下自动确认
};

// 图标映射
const stepIcons = {
    'search': '🔍',
    'data': '📊',
    'formula': '🧮',
    'chart': '📈',
    'format': '🎨',
    'clean': '🧹',
    'default': '⚡'
};

/**
 * 显示意图预览卡片
 * @param {Object} intentData - 意图数据 { description, plan, originalInput, autoConfirm }
 */
function showIntentPreview(intentData) {
    try {
        window.intentPreviewState.active = true;
        window.intentPreviewState.currentIntent = intentData;
        window.intentPreviewState.autoConfirm = intentData.autoConfirm || false;

        // 移除已存在的预览卡片
        hideIntentPreview();

        // 创建预览卡片
        const previewCard = createIntentPreviewCard(intentData);
        
        // 插入到输入区域上方
        const chatInputCard = document.getElementById('chat-input-card');
        if (chatInputCard && chatInputCard.parentNode) {
            chatInputCard.parentNode.insertBefore(previewCard, chatInputCard);
        }

        // 滚动到可见区域
        previewCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

        // Agent模式下自动确认
        if (window.intentPreviewState.autoConfirm) {
            setTimeout(function() {
                confirmIntent();
            }, 500);
        }

        console.log('显示意图预览:', intentData.description);
    } catch (err) {
        console.error('showIntentPreview error:', err);
    }
}

/**
 * 创建意图预览卡片
 * @param {Object} intentData - 意图数据
 * @returns {HTMLElement} 预览卡片元素
 */
function createIntentPreviewCard(intentData) {
    const card = document.createElement('div');
    card.id = 'intent-preview-card';
    card.className = 'intent-preview-card intent-preview-compact';

    card.innerHTML = `
        <div class="intent-preview-header">
            <span class="intent-preview-icon">🎯</span>
            <span class="intent-preview-title">我理解您想要：</span>
            <button class="intent-close-btn" onclick="cancelIntent()" title="关闭">×</button>
        </div>
        <div class="intent-preview-description">${escapeHtml(intentData.description || '处理您的请求')}</div>
        <div class="intent-preview-actions">
            <button class="intent-btn intent-btn-confirm" onclick="confirmIntent()">
                ✔ 确认执行
            </button>
            <button class="intent-btn intent-btn-edit" onclick="editIntent()">
                ✏ 修改
            </button>
            <button class="intent-btn intent-btn-cancel" onclick="cancelIntent()">
                ✖ 取消
            </button>
        </div>
    `;

    // 添加按钮事件监听（确保点击有效）
    setTimeout(function() {
        const confirmBtn = card.querySelector('.intent-btn-confirm');
        const editBtn = card.querySelector('.intent-btn-edit');
        const cancelBtn = card.querySelector('.intent-btn-cancel');
        const closeBtn = card.querySelector('.intent-close-btn');
        
        if (confirmBtn) confirmBtn.addEventListener('click', function(e) { e.stopPropagation(); confirmIntent(); });
        if (editBtn) editBtn.addEventListener('click', function(e) { e.stopPropagation(); editIntent(); });
        if (cancelBtn) cancelBtn.addEventListener('click', function(e) { e.stopPropagation(); cancelIntent(); });
        if (closeBtn) closeBtn.addEventListener('click', function(e) { e.stopPropagation(); cancelIntent(); });
    }, 0);

    return card;
}

/**
 * 渲染执行步骤
 * @param {Array} plan - 执行计划数组
 * @returns {string} HTML字符串
 */
function renderExecutionSteps(plan) {
    if (!plan || plan.length === 0) return '';

    return plan.map((step, idx) => {
        const icon = stepIcons[step.icon] || stepIcons['default'];
        const willModify = step.willModify ? `<span class="step-modify">→ ${escapeHtml(step.willModify)}</span>` : '';
        
        return `
            <div class="execution-step">
                <span class="step-number">${step.stepNumber || (idx + 1)}</span>
                <span class="step-icon">${icon}</span>
                <span class="step-description">${escapeHtml(step.description)}</span>
                ${willModify}
            </div>
        `;
    }).join('');
}

/**
 * 隐藏意图预览卡片
 */
function hideIntentPreview() {
    const existingCard = document.getElementById('intent-preview-card');
    if (existingCard) {
        existingCard.remove();
    }
    window.intentPreviewState.active = false;
    window.intentPreviewState.autoConfirm = false;
}

// 防抖标志 - 防止重复点击确认按钮
let confirmInProgress = false;

/**
 * 确认意图 - 发送消息
 */
function confirmIntent() {
    // 防抖检查
    if (confirmInProgress) {
        console.log('确认操作正在进行中，忽略重复点击');
        return;
    }
    confirmInProgress = true;

    try {
        const intentData = window.intentPreviewState.currentIntent;
        
        // 隐藏预览卡片
        hideIntentPreview();

        // 发送确认消息到后端
        sendMessageToServer({
            type: 'confirmIntent',
            intentData: intentData
        });

        console.log('用户确认意图');
    } catch (err) {
        console.error('confirmIntent error:', err);
    } finally {
        // 300ms后解除防抖锁定
        setTimeout(() => { confirmInProgress = false; }, 300);
    }
}

/**
 * 修改意图 - 允许用户编辑需求
 */
function editIntent() {
    try {
        const intentData = window.intentPreviewState.currentIntent;
        
        // 获取输入框
        const smartInput = document.getElementById('smart-input');
        const chatInput = document.getElementById('chat-input');
        
        // 将原始输入放回输入框
        if (intentData && intentData.originalInput) {
            if (smartInput) {
                smartInput.innerText = intentData.originalInput;
                smartInput.focus();
            } else if (chatInput) {
                chatInput.value = intentData.originalInput;
                chatInput.focus();
            }
        }

        // 隐藏预览卡片
        hideIntentPreview();

        console.log('用户选择修改需求');
    } catch (err) {
        console.error('editIntent error:', err);
    }
}

/**
 * 取消意图
 */
function cancelIntent() {
    try {
        // 隐藏预览卡片
        hideIntentPreview();

        // 清空输入框
        const smartInput = document.getElementById('smart-input');
        const chatInput = document.getElementById('chat-input');
        
        if (smartInput) {
            smartInput.innerText = '';
        }
        if (chatInput) {
            chatInput.value = '';
        }

        // 通知后端取消
        sendMessageToServer({
            type: 'cancelIntent'
        });

        // 恢复发送按钮状态
        changeSendButton();

        console.log('用户取消意图');
    } catch (err) {
        console.error('cancelIntent error:', err);
    }
}

/**
 * 检查是否处于意图预览状态
 * @returns {boolean}
 */
function isIntentPreviewActive() {
    return window.intentPreviewState.active;
}

/**
 * 更新意图预览状态指示器
 * @param {boolean} isProcessing - 是否正在处理
 */
function updateIntentPreviewStatus(isProcessing) {
    const card = document.getElementById('intent-preview-card');
    if (!card) return;

    if (isProcessing) {
        card.classList.add('processing');
        const header = card.querySelector('.intent-preview-header');
        if (header) {
            header.innerHTML = `
                <span class="intent-preview-icon spinning">⏳</span>
                <span class="intent-preview-title">正在分析您的意图...</span>
            `;
        }
    } else {
        card.classList.remove('processing');
    }
}

// 导出函数供全局使用
window.showIntentPreview = showIntentPreview;
window.hideIntentPreview = hideIntentPreview;
window.confirmIntent = confirmIntent;
window.editIntent = editIntent;
window.cancelIntent = cancelIntent;
window.isIntentPreviewActive = isIntentPreviewActive;
