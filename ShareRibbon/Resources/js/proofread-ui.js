/**
 * proofread-ui.js - 校对UI交互模块
 * 提供WPS风格的校对体验：波浪线标注 + Hover Tooltip + 问题列表
 */

// ========== 校对专注模式 ==========

/**
 * 显示校对侧边面板
 */
function showProofreadSidePanel() {
    // 移除其他面板
    if (typeof hideTemplateEditorPane === 'function') {
        hideTemplateEditorPane();
    }
    
    // 检查是否已存在
    if (document.getElementById('proofread-side-panel')) return;
    
    // 创建侧边面板容器
    var panel = document.createElement('div');
    panel.id = 'proofread-side-panel';
    panel.className = 'proofread-side-panel';
    panel.innerHTML = '<div class="proofread-panel-content" id="proofread-panel-content"></div>';
    
    document.body.appendChild(panel);
    
    // 添加面板样式（如果尚未添加）
    injectProofreadStyles();
}

/**
 * 隐藏校对侧边面板
 */
function hideProofreadSidePanel() {
    var panel = document.getElementById('proofread-side-panel');
    if (panel) panel.remove();
}

/**
 * 显示校对列表
 */
function showProofreadList(html) {
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = html;
        bindProofreadListEvents();
    }
}

/**
 * 绑定校对列表事件
 */
function bindProofreadListEvents() {
    // 接受按钮
    var acceptBtns = document.querySelectorAll('.issue-btn.accept');
    acceptBtns.forEach(function(btn) {
        btn.addEventListener('click', function() {
            var issueId = this.getAttribute('data-issue-id');
            acceptProofreadIssue(issueId);
        });
    });
    
    // 忽略按钮
    var ignoreBtns = document.querySelectorAll('.issue-btn.ignore');
    ignoreBtns.forEach(function(btn) {
        btn.addEventListener('click', function() {
            var issueId = this.getAttribute('data-issue-id');
            ignoreProofreadIssue(issueId);
        });
    });
}

/**
 * 接受校对修正
 */
function acceptProofreadIssue(issueId) {
    var payload = {
        type: 'proofread',
        action: 'accept',
        issueId: issueId
    };
    sendProofreadAction(payload);
    
    // 移除该项
    var item = document.querySelector('.proofread-issue-item[data-issue-id="' + issueId + '"]');
    if (item) {
        item.style.opacity = '0.5';
        item.style.pointerEvents = 'none';
        setTimeout(function() { item.remove(); }, 300);
    }
}

/**
 * 忽略校对问题
 */
function ignoreProofreadIssue(issueId) {
    var payload = {
        type: 'proofread',
        action: 'ignore',
        issueId: issueId
    };
    sendProofreadAction(payload);
    
    // 标记该项为已忽略
    var item = document.querySelector('.proofread-issue-item[data-issue-id="' + issueId + '"]');
    if (item) {
        item.style.opacity = '0.4';
        var actions = item.querySelector('.issue-actions');
        if (actions) actions.remove();
    }
}

/**
 * 接受所有校对修正
 */
function proofreadAcceptAll() {
    var payload = {
        type: 'proofread',
        action: 'acceptAll',
        issueId: ''
    };
    sendProofreadAction(payload);
    
    // 隐藏列表
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = '<div class="proofread-success">' +
            '<span class="success-icon">🎉</span>' +
            '<span class="success-text">所有问题已修正！</span>' +
            '</div>';
    }
}

/**
 * 退出校对模式
 */
function proofreadExit() {
    var payload = {
        type: 'proofread',
        action: 'exit',
        issueId: ''
    };
    sendProofreadAction(payload);
    
    hideProofreadSidePanel();
    hideProofreadModeIndicator();
}

/**
 * 发送校对操作到VB
 */
function sendProofreadAction(payload) {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(payload);
    } else if (window.vsto && typeof window.vsto.postMessage === 'function') {
        window.vsto.postMessage(payload);
    } else {
        console.error('[ProofreadUI] 无法发送消息，WebView不可用');
    }
}

/**
 * 更新校对摘要
 */
function updateProofreadSummary(total, high, medium, low) {
    // 查找现有摘要元素
    var existingSummary = document.getElementById('proofread-summary');
    if (existingSummary) {
        existingSummary.innerHTML = '共 ' + total + ' 处问题（' +
            '<span class="high">' + high + '处必须修改</span>，' +
            '<span class="medium">' + medium + '处建议修改</span>，' +
            '<span class="low">' + low + '处可选优化</span>）';
    }
}

/**
 * 显示无问题消息
 */
function showProofreadNoIssues() {
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = '<div class="proofread-success">' +
            '<span class="success-icon">✅</span>' +
            '<span class="success-text">没有发现问题！</span>' +
            '<p class="success-hint">您的文档没有需要修改的内容。</p>' +
            '</div>';
    }
}

/**
 * 显示全部修正完成消息
 */
function showProofreadAllCorrected() {
    var content = document.getElementById('proofread-panel-content');
    if (content) {
        content.innerHTML = '<div class="proofread-success">' +
            '<span class="success-icon">🎉</span>' +
            '<span class="success-text">所有问题已修正完成！</span>' +
            '<p class="success-hint">文档已全部修正，可以关闭校对面板了。</p>' +
            '</div>';
    }
}

// ========== 注入校对样式 ==========

function injectProofreadStyles() {
    if (document.getElementById('proofread-styles')) return;
    
    var style = document.createElement('style');
    style.id = 'proofread-styles';
    style.textContent = 
/* ========== 校对面板样式 ========== */
'.proofread-side-panel {' +
'    position: fixed;' +
'    right: 0;' +
'    top: 0;' +
'    bottom: 0;' +
'    width: 380px;' +
'    background: #fff;' +
'    box-shadow: -4px 0 20px rgba(0,0,0,0.1);' +
'    z-index: 1000;' +
'    display: flex;' +
'    flex-direction: column;' +
'    font-family: "Microsoft YaHei", "PingFang SC", sans-serif;' +
'}' +
'.proofread-panel-content {' +
'    flex: 1;' +
'    overflow-y: auto;' +
'    padding: 16px;' +
'}' +
/* ========== 校对列表样式 ========== */
'.proofread-list {' +
'    width: 100%;' +
'}' +
'.proofread-list-header {' +
'    display: flex;' +
'    align-items: center;' +
'    gap: 8px;' +
'    padding: 12px 0;' +
'    border-bottom: 1px solid #e5e7eb;' +
'    margin-bottom: 16px;' +
'}' +
'.proofread-list-icon {' +
'    font-size: 20px;' +
'}' +
'.proofread-list-title {' +
'    font-size: 16px;' +
'    font-weight: 600;' +
'    color: #1f2937;' +
'}' +
'.proofread-severity-group {' +
'    margin-bottom: 16px;' +
'}' +
'.severity-header {' +
'    font-size: 13px;' +
'    font-weight: 600;' +
'    padding: 8px 12px;' +
'    border-radius: 6px;' +
'    margin-bottom: 8px;' +
'}' +
'.severity-header.high {' +
'    background: #fef2f2;' +
'    color: #dc2626;' +
'}' +
'.severity-header.medium {' +
'    background: #fef3c7;' +
'    color: #d97706;' +
'}' +
'.severity-header.low {' +
'    background: #f0fdf4;' +
'    color: #16a34a;' +
'}' +
'.proofread-issue-item {' +
'    background: #f9fafb;' +
'    border-radius: 8px;' +
'    padding: 12px;' +
'    margin-bottom: 8px;' +
'    border-left: 3px solid transparent;' +
'    transition: opacity 0.3s;' +
'}' +
'.proofread-issue-item.high {' +
'    border-left-color: #dc2626;' +
'}' +
'.proofread-issue-item.medium {' +
'    border-left-color: #d97706;' +
'}' +
'.proofread-issue-item.low {' +
'    border-left-color: #16a34a;' +
'}' +
'.issue-header {' +
'    display: flex;' +
'    justify-content: space-between;' +
'    margin-bottom: 8px;' +
'    font-size: 12px;' +
'}' +
'.issue-location {' +
'    color: #6b7280;' +
'}' +
'.issue-type {' +
'    background: #e5e7eb;' +
'    color: #4b5563;' +
'    padding: 2px 8px;' +
'    border-radius: 10px;' +
'}' +
'.issue-content {' +
'    margin-bottom: 8px;' +
'}' +
'.issue-original,' +
'.issue-suggestion {' +
'    font-size: 13px;' +
'    margin-bottom: 4px;' +
'}' +
'.issue-original .label,' +
'.issue-suggestion .label {' +
'    color: #6b7280;' +
'    margin-right: 6px;' +
'}' +
'.issue-original .text {' +
'    color: #dc2626;' +
'}' +
'.issue-suggestion .text {' +
'    color: #16a34a;' +
'    font-weight: 500;' +
'}' +
'.issue-explanation {' +
'    font-size: 12px;' +
'    color: #6b7280;' +
'    background: #f3f4f6;' +
'    padding: 6px 10px;' +
'    border-radius: 4px;' +
'    margin-bottom: 8px;' +
'}' +
'.issue-actions {' +
'    display: flex;' +
'    gap: 8px;' +
'}' +
'.issue-btn {' +
'    flex: 1;' +
'    padding: 6px 12px;' +
'    border: none;' +
'    border-radius: 6px;' +
'    font-size: 12px;' +
'    cursor: pointer;' +
'    transition: all 0.2s;' +
'}' +
'.issue-btn.accept {' +
'    background: #2563eb;' +
'    color: white;' +
'}' +
'.issue-btn.accept:hover {' +
'    background: #1d4ed8;' +
'}' +
'.issue-btn.ignore {' +
'    background: #f3f4f6;' +
'    color: #6b7280;' +
'}' +
'.issue-btn.ignore:hover {' +
'    background: #e5e7eb;' +
'}' +
'.proofread-list-actions {' +
'    display: flex;' +
'    gap: 8px;' +
'    margin-top: 16px;' +
'    padding-top: 16px;' +
'    border-top: 1px solid #e5e7eb;' +
'}' +
'.proofread-btn {' +
'    flex: 1;' +
'    padding: 10px 16px;' +
'    border: none;' +
'    border-radius: 8px;' +
'    font-size: 14px;' +
'    font-weight: 500;' +
'    cursor: pointer;' +
'    transition: all 0.2s;' +
'}' +
'.proofread-btn.primary {' +
'    background: #2563eb;' +
'    color: white;' +
'}' +
'.proofread-btn.primary:hover {' +
'    background: #1d4ed8;' +
'}' +
'.proofread-btn.secondary {' +
'    background: #f3f4f6;' +
'    color: #4b5563;' +
'}' +
'.proofread-btn.secondary:hover {' +
'    background: #e5e7eb;' +
'}' +
'.proofread-more {' +
'    text-align: center;' +
'    color: #6b7280;' +
'    font-size: 12px;' +
'    padding: 8px;' +
'}' +
/* ========== 校对成功样式 ========== */
'.proofread-success {' +
'    display: flex;' +
'    flex-direction: column;' +
'    align-items: center;' +
'    justify-content: center;' +
'    padding: 60px 20px;' +
'    text-align: center;' +
'}' +
'.proofread-success .success-icon {' +
'    font-size: 56px;' +
'    margin-bottom: 16px;' +
'}' +
'.proofread-success .success-text {' +
'    font-size: 18px;' +
'    color: #16a34a;' +
'    font-weight: 600;' +
'    margin-bottom: 8px;' +
'}' +
'.proofread-success .success-hint {' +
'    font-size: 14px;' +
'    color: #6b7280;' +
'    margin-top: 8px;' +
'}';
    
    document.head.appendChild(style);
}

// ========== 页面加载时初始化 ==========

// 等待DOM加载完成
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectProofreadStyles);
} else {
    injectProofreadStyles();
}
