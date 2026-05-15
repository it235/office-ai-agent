/**
 * reformat-chat.js - 排版卡片交互模块
 * 在Chat中渲染并交互排版建议卡片，支持应用排版、预览、切换模板、微调等操作。
 * 由VB.NET后端通过 ExecuteJavaScriptAsync 调用 appendFormattingCard 推送卡片HTML，
 * 前端JS通过事件委托处理按钮点击，通过 window.location.href 回传操作命令。
 */

// ====== CSS注入 ======
(function injectStyles() {
    var styleId = 'reformat-chat-styles';
    if (document.getElementById(styleId)) return;

    var style = document.createElement('style');
    style.id = styleId;
    style.textContent =
'/* ====== 排版卡片样式 (reformat-chat.js) ====== */\n' +
'.formatting-card {\n' +
'    background: #ffffff;\n' +
'    border: 1px solid #e2e8f0;\n' +
'    border-radius: 10px;\n' +
'    box-shadow: 0 2px 8px rgba(0,0,0,0.06);\n' +
'    overflow: hidden;\n' +
'    margin: 12px 0;\n' +
'    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;\n' +
'}\n' +
'.formatting-card-header {\n' +
'    display: flex;\n' +
'    align-items: center;\n' +
'    gap: 8px;\n' +
'    padding: 12px 16px;\n' +
'    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);\n' +
'    color: white;\n' +
'    font-weight: 600;\n' +
'    font-size: 14px;\n' +
'}\n' +
'.formatting-card-icon {\n' +
'    font-size: 16px;\n' +
'}\n' +
'.formatting-card-title {\n' +
'    font-size: 14px;\n' +
'    font-weight: 600;\n' +
'}\n' +
'.formatting-card-body {\n' +
'    padding: 14px 16px;\n' +
'}\n' +
'.formatting-info-row {\n' +
'    font-size: 13px;\n' +
'    color: #4a5568;\n' +
'    padding: 4px 0;\n' +
'    line-height: 1.6;\n' +
'}\n' +
'.formatting-info-row strong {\n' +
'    color: #2d3748;\n' +
'}\n' +
'.formatting-changes {\n' +
'    margin: 10px 0;\n' +
'    background: #f7fafc;\n' +
'    border-radius: 8px;\n' +
'    padding: 10px 12px;\n' +
'    border: 1px solid #e2e8f0;\n' +
'}\n' +
'.formatting-changes-title {\n' +
'    font-size: 13px;\n' +
'    font-weight: 600;\n' +
'    color: #2d3748;\n' +
'    margin-bottom: 6px;\n' +
'}\n' +
'.formatting-change-item {\n' +
'    display: flex;\n' +
'    align-items: baseline;\n' +
'    gap: 6px;\n' +
'    padding: 4px 0;\n' +
'    font-size: 13px;\n' +
'    line-height: 1.5;\n' +
'}\n' +
'.formatting-change-section {\n' +
'    color: #2563eb;\n' +
'    font-weight: 500;\n' +
'    white-space: nowrap;\n' +
'}\n' +
'.formatting-change-count {\n' +
'    color: #718096;\n' +
'    font-size: 12px;\n' +
'    white-space: nowrap;\n' +
'}\n' +
'.formatting-change-desc {\n' +
'    color: #4a5568;\n' +
'}\n' +
'.formatting-change-summary {\n' +
'    margin-top: 6px;\n' +
'    padding-top: 6px;\n' +
'    border-top: 1px dashed #e2e8f0;\n' +
'    font-size: 12px;\n' +
'    color: #718096;\n' +
'}\n' +
'.formatting-diff {\n' +
'    margin: 10px 0;\n' +
'}\n' +
'.formatting-diff-item {\n' +
'    display: flex;\n' +
'    align-items: baseline;\n' +
'    gap: 6px;\n' +
'    padding: 6px 8px;\n' +
'    font-size: 13px;\n' +
'    background: #f7fafc;\n' +
'    border-radius: 6px;\n' +
'    margin-bottom: 4px;\n' +
'    border: 1px solid #e2e8f0;\n' +
'}\n' +
'.formatting-diff-section {\n' +
'    color: #2563eb;\n' +
'    font-weight: 500;\n' +
'    white-space: nowrap;\n' +
'}\n' +
'.formatting-diff-old {\n' +
'    color: #e53e3e;\n' +
'    text-decoration: line-through;\n' +
'    font-size: 12px;\n' +
'}\n' +
'.formatting-diff-arrow {\n' +
'    color: #a0aec0;\n' +
'    font-size: 14px;\n' +
'}\n' +
'.formatting-diff-new {\n' +
'    color: #38a169;\n' +
'    font-weight: 500;\n' +
'}\n' +
'.formatting-card-actions {\n' +
'    display: flex;\n' +
'    flex-wrap: wrap;\n' +
'    gap: 8px;\n' +
'    margin-top: 12px;\n' +
'    padding-top: 12px;\n' +
'    border-top: 1px solid #edf2f7;\n' +
'}\n' +
'.formatting-btn {\n' +
'    display: inline-flex;\n' +
'    align-items: center;\n' +
'    gap: 4px;\n' +
'    padding: 7px 16px;\n' +
'    border: 1px solid transparent;\n' +
'    border-radius: 6px;\n' +
'    font-size: 13px;\n' +
'    font-weight: 500;\n' +
'    cursor: pointer;\n' +
'    transition: all 0.2s;\n' +
'    font-family: inherit;\n' +
'}\n' +
'.formatting-btn-primary {\n' +
'    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);\n' +
'    color: white;\n' +
'    border: none;\n' +
'    box-shadow: 0 2px 6px rgba(37,99,235,0.3);\n' +
'}\n' +
'.formatting-btn-primary:hover {\n' +
'    box-shadow: 0 4px 12px rgba(37,99,235,0.4);\n' +
'    transform: translateY(-1px);\n' +
'}\n' +
'.formatting-btn-secondary {\n' +
'    background: #edf2f7;\n' +
'    color: #2d3748;\n' +
'    border: 1px solid #e2e8f0;\n' +
'}\n' +
'.formatting-btn-secondary:hover {\n' +
'    background: #e2e8f0;\n' +
'}\n' +
'.formatting-btn-outline {\n' +
'    background: white;\n' +
'    color: #718096;\n' +
'    border: 1px solid #e2e8f0;\n' +
'}\n' +
'.formatting-btn-outline:hover {\n' +
'    border-color: #a0aec0;\n' +
'    color: #4a5568;\n' +
'}\n' +
'.formatting-btn-ghost {\n' +
'    background: transparent;\n' +
'    color: #718096;\n' +
'    border: none;\n' +
'    padding: 7px 12px;\n' +
'}\n' +
'.formatting-btn-ghost:hover {\n' +
'    color: #4a5568;\n' +
'    background: #f7fafc;\n' +
'}\n' +
'.formatting-refine-area {\n' +
'    margin-top: 10px;\n' +
'    padding: 10px 12px;\n' +
'    background: #fffaf0;\n' +
'    border: 1px solid #fbd38d;\n' +
'    border-radius: 8px;\n' +
'}\n' +
'.formatting-refine-input-row {\n' +
'    display: flex;\n' +
'    gap: 8px;\n' +
'}\n' +
'.formatting-refine-input {\n' +
'    flex: 1;\n' +
'    padding: 8px 12px;\n' +
'    border: 1px solid #e2e8f0;\n' +
'    border-radius: 6px;\n' +
'    font-size: 13px;\n' +
'    font-family: inherit;\n' +
'    outline: none;\n' +
'    transition: border 0.2s;\n' +
'}\n' +
'.formatting-refine-input:focus {\n' +
'    border-color: #2563eb;\n' +
'    box-shadow: 0 0 0 2px rgba(37,99,235,0.15);\n' +
'}\n' +
'.formatting-refine-send-btn {\n' +
'    padding: 8px 16px;\n' +
'    background: #2563eb;\n' +
'    color: white;\n' +
'    border: none;\n' +
'    border-radius: 6px;\n' +
'    font-size: 13px;\n' +
'    cursor: pointer;\n' +
'    font-weight: 500;\n' +
'    white-space: nowrap;\n' +
'}\n' +
'.formatting-refine-send-btn:hover {\n' +
'    background: #1d4ed8;\n' +
'}\n' +
'.quick-reformat-indicator {\n' +
'    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);\n' +
'    border-radius: 8px;\n' +
'    padding: 10px 16px;\n' +
'    margin-bottom: 12px;\n' +
'    color: white;\n' +
'    animation: fadeIn 0.3s ease;\n' +
'}\n' +
'.quick-reformat-indicator-content {\n' +
'    display: flex;\n' +
'    align-items: center;\n' +
'    gap: 8px;\n' +
'}\n' +
'.quick-reformat-icon {\n' +
'    font-size: 18px;\n' +
'}\n' +
'.quick-reformat-text {\n' +
'    flex: 1;\n' +
'    font-size: 14px;\n' +
'    font-weight: 500;\n' +
'}\n' +
'.quick-reformat-dismiss {\n' +
'    background: rgba(255,255,255,0.2);\n' +
'    border: none;\n' +
'    color: white;\n' +
'    width: 24px;\n' +
'    height: 24px;\n' +
'    border-radius: 50%;\n' +
'    cursor: pointer;\n' +
'    display: flex;\n' +
'    align-items: center;\n' +
'    justify-content: center;\n' +
'    font-size: 16px;\n' +
'    transition: background 0.2s;\n' +
'}\n' +
'.quick-reformat-dismiss:hover {\n' +
'    background: rgba(255,255,255,0.35);\n' +
'}\n' +
'\n' +
'/* ====== 智能排版面板（模板Tab内） ====== */\n' +
'.smart-format-panel {\n' +
'    display: flex;\n' +
'    flex-direction: column;\n' +
'    align-items: center;\n' +
'    gap: 24px;\n' +
'    padding: 40px 20px;\n' +
'    max-width: 480px;\n' +
'    margin: 0 auto;\n' +
'}\n' +
'.smart-format-hero {\n' +
'    text-align: center;\n' +
'}\n' +
'.smart-format-quick-btn {\n' +
'    display: flex;\n' +
'    flex-direction: column;\n' +
'    align-items: center;\n' +
'    gap: 10px;\n' +
'    padding: 36px 56px;\n' +
'    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);\n' +
'    color: white;\n' +
'    border: none;\n' +
'    border-radius: 16px;\n' +
'    cursor: pointer;\n' +
'    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);\n' +
'    box-shadow: 0 8px 24px rgba(37,99,235,0.3);\n' +
'    width: 100%;\n' +
'    max-width: 360px;\n' +
'}\n' +
'.smart-format-quick-btn:hover {\n' +
'    transform: translateY(-4px);\n' +
'    box-shadow: 0 12px 36px rgba(37,99,235,0.4);\n' +
'}\n' +
'.smart-format-quick-btn svg {\n' +
'    width: 48px;\n' +
'    height: 48px;\n' +
'    opacity: 0.9;\n' +
'}\n' +
'.smart-format-quick-btn span {\n' +
'    font-size: 20px;\n' +
'    font-weight: 700;\n' +
'}\n' +
'.smart-format-quick-btn small {\n' +
'    font-size: 13px;\n' +
'    opacity: 0.85;\n' +
'    font-weight: 400;\n' +
'}\n' +
'.smart-format-dialog {\n' +
'    width: 100%;\n' +
'    max-width: 420px;\n' +
'    background: white;\n' +
'    border: 1px solid #e2e8f0;\n' +
'    border-radius: 12px;\n' +
'    padding: 20px;\n' +
'    box-shadow: 0 2px 8px rgba(0,0,0,0.04);\n' +
'}\n' +
'.smart-format-dialog-label {\n' +
'    font-size: 14px;\n' +
'    font-weight: 600;\n' +
'    color: #2d3748;\n' +
'    margin-bottom: 10px;\n' +
'    display: flex;\n' +
'    align-items: center;\n' +
'    gap: 6px;\n' +
'}\n' +
'.smart-format-dialog-input-row {\n' +
'    display: flex;\n' +
'    gap: 8px;\n' +
'}\n' +
'.smart-format-dialog-input {\n' +
'    flex: 1;\n' +
'    padding: 10px 14px;\n' +
'    border: 1px solid #e2e8f0;\n' +
'    border-radius: 8px;\n' +
'    font-size: 14px;\n' +
'    font-family: inherit;\n' +
'    outline: none;\n' +
'    transition: border 0.2s;\n' +
'}\n' +
'.smart-format-dialog-input:focus {\n' +
'    border-color: #2563eb;\n' +
'    box-shadow: 0 0 0 2px rgba(37,99,235,0.15);\n' +
'}\n' +
'.smart-format-dialog-send {\n' +
'    padding: 10px 20px;\n' +
'    background: #2563eb;\n' +
'    color: white;\n' +
'    border: none;\n' +
'    border-radius: 8px;\n' +
'    font-size: 14px;\n' +
'    cursor: pointer;\n' +
'    font-weight: 500;\n' +
'    white-space: nowrap;\n' +
'}\n' +
'.smart-format-dialog-send:hover {\n' +
'    background: #1d4ed8;\n' +
'}\n' +
'.smart-format-clone-btn {\n' +
'    display: flex;\n' +
'    flex-direction: column;\n' +
'    align-items: center;\n' +
'    gap: 8px;\n' +
'    padding: 28px 40px;\n' +
'    background: white;\n' +
'    color: #2d3748;\n' +
'    border: 2px dashed #cbd5e0;\n' +
'    border-radius: 12px;\n' +
'    cursor: pointer;\n' +
'    transition: all 0.3s;\n' +
'    width: 100%;\n' +
'    max-width: 420px;\n' +
'}\n' +
'.smart-format-clone-btn:hover {\n' +
'    border-color: #667eea;\n' +
'    background: #f8f9ff;\n' +
'    transform: translateY(-2px);\n' +
'}\n' +
'.smart-format-clone-btn svg {\n' +
'    width: 36px;\n' +
'    height: 36px;\n' +
'    opacity: 0.6;\n' +
'}\n' +
'.smart-format-clone-btn span {\n' +
'    font-size: 16px;\n' +
'    font-weight: 600;\n' +
'}\n' +
'.smart-format-clone-btn small {\n' +
'    font-size: 12px;\n' +
'    color: #718096;\n' +
'}\n';
    document.head.appendChild(style);
})();

// ====== 全局状态 ======
/** @type {string|null} 当前活跃的排版卡片UUID */
var _activeFormattingCardUuid = null;

// ====== 核心函数 ======

/**
 * 追加排版卡片到Chat（由VB.NET通过 ExecuteJavaScriptAsync 调用）。
 * @param {Object|string} payload - { uuid: string, html: string } 或 JSON字符串
 */
window.appendFormattingCard = function(payload) {
    // 解析字符串payload
    if (typeof payload === 'string') {
        try { payload = JSON.parse(payload); } catch (e) {
            console.error('[ReformatChat] 解析payload失败:', e);
            return;
        }
    }
    if (!payload || !payload.uuid || typeof payload.html !== 'string') {
        console.error('[ReformatChat] payload缺少uuid或html字段');
        return;
    }

    var uuid = payload.uuid;
    _activeFormattingCardUuid = uuid;

    // 查找或创建Chat消息Section
    var contentDiv = document.getElementById('content-' + uuid);
    if (!contentDiv) {
        var sender = '排版助手'; // "排版助手"
        var timestamp = typeof formatDateTime === 'function'
            ? formatDateTime(new Date())
            : new Date().toLocaleString('zh-CN');
        window.createChatSection(sender, timestamp, uuid);
        contentDiv = document.getElementById('content-' + uuid);
        if (!contentDiv) {
            console.error('[ReformatChat] createChatSection 未能创建 contentDiv');
            return;
        }
    }

    // 注入卡片HTML
    contentDiv.innerHTML = payload.html;

    // 为卡片标注uuid，移除内联onclick（改用事件委托）
    var card = contentDiv.querySelector('.formatting-card');
    if (card) {
        card.dataset.uuid = uuid;
        card.querySelectorAll('[onclick]').forEach(function(el) {
            el.removeAttribute('onclick');
        });
    }

    // 追加微调输入区域（默认隐藏）
    var cardBody = contentDiv.querySelector('.formatting-card-body');
    if (cardBody && !cardBody.querySelector('.formatting-refine-area')) {
        var refineArea = document.createElement('div');
        refineArea.className = 'formatting-refine-area';
        refineArea.style.display = 'none';
        refineArea.innerHTML =
            '<div class="formatting-refine-input-row">' +
            '  <input type="text" class="formatting-refine-input" placeholder="输入微调指令，如：标题改为黑体居中..." />' +
            '  <button class="formatting-btn formatting-refine-send-btn" data-uuid="' + uuid + '">发送</button>' +
            '</div>';
        cardBody.appendChild(refineArea);
    }

    // 自动滚动到卡片
    setTimeout(function() {
        var el = document.getElementById('content-' + uuid);
        if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }, 100);
};

// ====== 事件委托：处理卡片按钮点击 ======
document.addEventListener('click', function(e) {
    var btn = e.target.closest('.formatting-btn');
    if (!btn) return;

    var card = btn.closest('.formatting-card');
    if (!card) return;

    var uuid = card.dataset.uuid;
    if (!uuid) return;

    // 根据按钮样式类名分派操作
    if (btn.classList.contains('formatting-btn-primary')) {
        // "应用排版"
        applyFormattingCard(uuid);
    } else if (btn.classList.contains('formatting-btn-secondary')) {
        // "预览对比"
        previewFormattingCompare(uuid);
    } else if (btn.classList.contains('formatting-btn-outline')) {
        // "换一种"
        switchFormattingTemplate(uuid);
    } else if (btn.classList.contains('formatting-btn-ghost')) {
        // "微调" / "继续微调"
        toggleRefinementInput(uuid);
    } else if (btn.classList.contains('formatting-refine-send-btn')) {
        // "发送"微调指令
        var input = card.querySelector('.formatting-refine-input');
        if (input && input.value.trim()) {
            sendRefinementCommand(uuid, input.value.trim());
            input.value = '';
        }
    }
});

// ====== 向后兼容：VB生成HTML中的内联onclick ======
// 对于未通过 appendFormattingCard 注入的卡片仍有效
window.applyReformat = function() {
    if (_activeFormattingCardUuid) {
        applyFormattingCard(_activeFormattingCardUuid);
    }
};

window.previewReformat = function() {
    if (_activeFormattingCardUuid) {
        previewFormattingCompare(_activeFormattingCardUuid);
    }
};

window.alternateReformat = function() {
    if (_activeFormattingCardUuid) {
        switchFormattingTemplate(_activeFormattingCardUuid);
    }
};

window.startRefinement = function() {
    if (_activeFormattingCardUuid) {
        toggleRefinementInput(_activeFormattingCardUuid);
    }
};

// ====== UUID基础操作函数（与VB.NET通信） ======

/**
 * 向VB.NET后端发送消息（统一通过chrome.webview.postMessage）
 */
function sendReformatAction(actionType, params) {
    var payload = Object.assign({ type: actionType }, params || {});
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(payload);
    } else if (window.vsto && typeof window.vsto.postMessage === 'function') {
        window.vsto.postMessage(payload);
    } else {
        console.error('[ReformatChat] 无法发送消息，WebView不可用');
    }
}

/**
 * 应用排版。
 * @param {string} uuid - 卡片UUID
 */
window.applyFormattingCard = function(uuid) {
    sendReformatAction('applySmartReformat', { uuid: uuid });
};

/**
 * 切换排版方案（换一种）。
 * @param {string} uuid - 卡片UUID
 */
window.switchFormattingTemplate = function(uuid) {
    sendReformatAction('switchReformatTemplate', { uuid: uuid });
};

/**
 * 预览排版对比效果。
 * @param {string} uuid - 卡片UUID
 */
window.previewFormattingCompare = function(uuid) {
    sendReformatAction('previewReformatCompare', { uuid: uuid });
};

/**
 * 发送微调指令。
 * @param {string} uuid - 卡片UUID
 * @param {string} command - 微调指令文本
 */
window.sendRefinementCommand = function(uuid, command) {
    if (!command || !command.trim()) return;
    sendReformatAction('refineSmartReformat', { uuid: uuid, command: command });
};

/**
 * 从卡片输入框读取内容并发送微调指令。
 * @param {string} uuid - 卡片UUID
 */
window.sendRefinementFromCard = function(uuid) {
    var input = document.querySelector('#content-' + uuid + ' .formatting-refine-input');
    if (input && input.value.trim()) {
        sendRefinementCommand(uuid, input.value.trim());
        input.value = '';
    }
};

/**
 * 切换微调输入区域的显示/隐藏。
 * @param {string} uuid - 卡片UUID
 */
window.toggleRefinementInput = function(uuid) {
    var refineArea = document.querySelector('#content-' + uuid + ' .formatting-refine-area');
    if (!refineArea) return;

    var isHidden = refineArea.style.display === 'none' || refineArea.style.display === '';
    refineArea.style.display = isHidden ? 'block' : 'none';

    if (isHidden) {
        var input = refineArea.querySelector('.formatting-refine-input');
        if (input) setTimeout(function() { input.focus(); }, 50);
    }
};

// ====== 速排指示器 ======

/**
 * 在Chat顶部显示速排模式指示条。
 */
window.showQuickReformatIndicator = function() {
    var existing = document.getElementById('quick-reformat-indicator');
    if (existing) return;

    var chatContainer = document.getElementById('chat-container');
    if (!chatContainer) return;

    var indicator = document.createElement('div');
    indicator.id = 'quick-reformat-indicator';
    indicator.className = 'quick-reformat-indicator';
    indicator.innerHTML =
        '<div class="quick-reformat-indicator-content">' +
        '  <span class="quick-reformat-icon">⚡</span>' +
        '  <span class="quick-reformat-text">速排模式已开启</span>' +
        '  <button class="quick-reformat-dismiss" id="quick-reformat-dismiss-btn">&times;</button>' +
        '</div>';

    chatContainer.insertBefore(indicator, chatContainer.firstChild);

    var dismissBtn = document.getElementById('quick-reformat-dismiss-btn');
    if (dismissBtn) {
        dismissBtn.addEventListener('click', function() {
            window.dismissQuickReformatIndicator();
        });
    }
};

/**
 * 关闭速排指示器。
 */
window.dismissQuickReformatIndicator = function() {
    var indicator = document.getElementById('quick-reformat-indicator');
    if (indicator) indicator.remove();
};

// ====== 智能排版模式切换 ======

/**
 * 进入智能排版模式。
 */
window.enterSmartReformatMode = function() {
    document.body.classList.add('smart-reformat-mode');
    showQuickReformatIndicator();
    sendReformatAction('enterSmartReformat');
};

/**
 * 退出智能排版模式。
 */
window.exitSmartReformatMode = function() {
    document.body.classList.remove('smart-reformat-mode');
    sendReformatAction('exitSmartReformat');
};

/**
 * 显示排版引导提示（当用户未选中文本时由 Ribbon 按钮调用）
 */
window.showQuickReformatGuide = function() {
    var existing = document.getElementById('quick-reformat-guide');
    if (existing) return;

    var chatContainer = document.getElementById('chat-container');
    if (!chatContainer) return;

    var guide = document.createElement('div');
    guide.id = 'quick-reformat-guide';
    guide.className = 'quick-reformat-guide';
    guide.innerHTML =
        '<div class="quick-reformat-guide-content" style="background:#f0f7ff;border:1px solid #b3d4fc;border-radius:8px;padding:16px;margin:12px;">' +
        '  <div style="font-weight:600;color:#1a56db;margin-bottom:8px;">&#x1F4D0; 智能排版使用说明</div>' +
        '  <ul style="margin:0;padding-left:20px;color:#444;font-size:13px;line-height:1.8;">' +
        '    <li>在文档中<span style="font-weight:600;">选中需要排版的内容</span>，然后点击 Ribbon 上的"排版"按钮</li>' +
        '    <li>或在 Chat 中输入排版指令，如：<code style="background:#e8e8e8;padding:2px 6px;border-radius:3px;">按公文标准排版</code></li>' +
        '    <li>支持标准：公文(GB/T 9704)、学术论文、商务报告、合同等</li>' +
        '  </ul>' +
        '  <button onclick="this.parentElement.parentElement.remove()" style="margin-top:8px;padding:4px 12px;border:1px solid #1a56db;background:white;color:#1a56db;border-radius:4px;cursor:pointer;font-size:12px;">知道了</button>' +
        '</div>';

    chatContainer.insertBefore(guide, chatContainer.firstChild);
};
