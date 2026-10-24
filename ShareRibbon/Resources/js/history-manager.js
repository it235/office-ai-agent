/**
 * history-manager.js - History Sidebar Management
 * Handles chat session list, document scoping, rename and delete
 */

window.historyManager = {
    isOpen: false,
    scope: 'all',
    editingSessionId: null,
    documentName: '',
    hasDocument: false,
    items: [],

    // Initialize history functionality
    init: function () {
        const toggleBtn = document.getElementById('history-toggle-btn');
        const sidebar = document.getElementById('history-sidebar');
        const overlay = document.getElementById('sidebar-overlay');
        const closeBtn = document.getElementById('close-sidebar-btn');
        const newSessionBtn = document.getElementById('new-session-btn');

        // Bind events
        toggleBtn.addEventListener('click', () => this.toggleSidebar());
        closeBtn.addEventListener('click', () => this.closeSidebar());
        overlay.addEventListener('click', () => this.closeSidebar());
        if (newSessionBtn) {
            newSessionBtn.addEventListener('click', () => {
                this.sendMessageToVB({ type: 'newSession' });
                this.closeSidebar();
            });
        }

        this.bindListEvents();

        // Keyboard event
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && this.isOpen) {
                if (this.editingSessionId) {
                    this.cancelRename();
                    return;
                }
                this.closeSidebar();
            }
        });
    },

    // 条目是每次刷新重绘的，动作一律走委托绑定：逐条 addEventListener 会随刷新重复挂同一批监听
    bindListEvents: function () {
        const list = document.getElementById('history-list');
        if (!list || list.dataset.bound === '1') return;
        list.dataset.bound = '1';

        list.addEventListener('click', (e) => {
            const scopeBtn = e.target.closest('[data-scope]');
            if (scopeBtn) {
                this.setScope(scopeBtn.getAttribute('data-scope'));
                return;
            }

            const item = e.target.closest('.history-item');
            if (!item) return;
            const sessionId = item.getAttribute('data-session-id') || '';

            if (e.target.closest('.history-item-rename')) {
                this.startRename(sessionId);
                return;
            }
            if (e.target.closest('.history-item-delete')) {
                this.deleteSession(sessionId);
                return;
            }
            this.loadSession(sessionId);
        });

        list.addEventListener('keydown', (e) => {
            if (!e.target.classList || !e.target.classList.contains('history-rename-input')) return;
            if (e.key === 'Enter') {
                e.preventDefault();
                this.commitRename(e.target.getAttribute('data-session-id'), e.target.value);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                e.stopPropagation();
                this.cancelRename();
            }
        });

        list.addEventListener('focusout', (e) => {
            if (e.target.classList && e.target.classList.contains('history-rename-input')) {
                this.commitRename(e.target.getAttribute('data-session-id'), e.target.value);
            }
        });
    },

    setScope: function (scope) {
        this.scope = scope === 'document' ? 'document' : 'all';
        this.editingSessionId = null;
        this.refresh();
    },

    refresh: function () {
        const historyList = document.getElementById('history-list');
        if (historyList) {
            historyList.innerHTML = '<div class="loading-state">正在加载历史记录...</div>';
        }
        this.sendMessageToVB({ type: 'getSessionList', scope: this.scope });
    },

    // Toggle sidebar visibility
    toggleSidebar: function () {
        if (this.isOpen) {
            this.closeSidebar();
        } else {
            this.openSidebar();
        }
    },

    // Open sidebar
    openSidebar: function () {
        const sidebar = document.getElementById('history-sidebar');
        const overlay = document.getElementById('sidebar-overlay');

        sidebar.classList.remove('sidebar-hidden');
        sidebar.classList.add('sidebar-visible');
        overlay.classList.remove('overlay-hidden');
        overlay.classList.add('overlay-visible');

        this.isOpen = true;

        // Load history files
        this.loadHistoryFiles();
    },

    // Close sidebar
    closeSidebar: function () {
        const sidebar = document.getElementById('history-sidebar');
        const overlay = document.getElementById('sidebar-overlay');

        sidebar.classList.remove('sidebar-visible');
        sidebar.classList.add('sidebar-hidden');
        overlay.classList.remove('overlay-visible');
        overlay.classList.add('overlay-hidden');

        this.isOpen = false;
        this.editingSessionId = null;
    },

    // Load history files list
    loadHistoryFiles: function () {
        // Request session list from backend (conversation_event + session_summary)
        this.refresh();
    },

    // Display session list from backend
    displayHistoryFiles: function (payload) {
        const historyList = document.getElementById('history-list');
        if (!historyList) return;

        const data = (payload && !Array.isArray(payload))
            ? payload
            : { items: Array.isArray(payload) ? payload : [], scope: 'all', documentName: '', hasDocument: false };

        this.items = data.items || [];
        this.scope = data.scope === 'document' ? 'document' : 'all';
        this.documentName = data.documentName || '';
        this.hasDocument = data.hasDocument === true;

        const barHtml = this.renderScopeBar();
        if (this.items.length === 0) {
            historyList.innerHTML = barHtml + this.renderEmptyState();
            return;
        }

        // 顺序由后端决定：scope=all 时当前文档的会话排在最前，前端再按时间重排会把它冲掉
        historyList.innerHTML = barHtml + this.items.map(s => this.renderItem(s)).join('');
        this.focusRenameInput();
    },

    renderScopeBar: function () {
        const docLabel = this.hasDocument ? this.escapeHtml(this.documentName) : '通用会话';
        return `<div class="history-scope-bar">
            <button type="button" class="history-scope-tab${this.scope === 'document' ? ' active' : ''}" data-scope="document">本文档</button>
            <button type="button" class="history-scope-tab${this.scope === 'all' ? ' active' : ''}" data-scope="all">全部</button>
            <span class="history-scope-doc" title="${docLabel}">${docLabel}</span>
        </div>`;
    },

    renderEmptyState: function () {
        if (this.scope === 'document' && this.hasDocument) {
            return `<div class="empty-state">
                <div class="empty-state-text">这篇文档还没有会话</div>
                <button type="button" class="history-empty-action" data-scope="all">查看全部会话</button>
            </div>`;
        }
        return `<div class="empty-state">
            <div class="empty-state-icon">📄</div>
            <div class="empty-state-text">您还没有任何历史会话</div>
        </div>`;
    },

    renderItem: function (s) {
        const sid = s.sessionId || '';
        if (this.editingSessionId && this.editingSessionId === sid) {
            return `<div class="history-item history-item-editing" data-session-id="${this.escapeHtml(sid)}">
                <input type="text" class="history-rename-input" data-session-id="${this.escapeHtml(sid)}" maxlength="60" />
                <div class="history-rename-hint">回车保存 · Esc 取消</div>
            </div>`;
        }
        const title = s.title || '会话';
        return `<div class="history-item" data-session-id="${this.escapeHtml(sid)}" title="${this.escapeHtml(title)}">
            <div class="history-item-main">
                <div class="history-item-title">${this.escapeHtml(title)}</div>
                <div class="history-item-meta">${this.renderItemMeta(s)}</div>
            </div>
            <div class="history-item-actions">
                <button type="button" class="history-item-rename" title="重命名">✎</button>
                <button type="button" class="history-item-delete" title="删除会话">🗑</button>
            </div>
        </div>`;
    },

    renderItemMeta: function (s) {
        const parts = [];
        if (this.scope === 'all') {
            parts.push(s.documentName ? this.escapeHtml(s.documentName) : '通用会话');
        }
        parts.push(this.escapeHtml(this.formatSessionDate(s.createdAt)));
        return parts.map(p => `<span class="history-meta-part">${p}</span>`).join('');
    },

    startRename: function (sessionId) {
        if (!sessionId) return;
        this.editingSessionId = sessionId;
        this.refresh();
    },

    focusRenameInput: function () {
        if (!this.editingSessionId) return;
        const input = document.querySelector('.history-rename-input');
        if (!input) return;
        const item = this.items.find(s => s.sessionId === this.editingSessionId);
        input.value = item ? (item.title || '') : '';
        input.focus();
        input.select();
    },

    cancelRename: function () {
        if (!this.editingSessionId) return;
        this.editingSessionId = null;
        this.refresh();
    },

    commitRename: function (sessionId, title) {
        if (!this.editingSessionId) return;
        this.editingSessionId = null;
        const next = (title || '').trim();
        if (sessionId && next) {
            this.sendMessageToVB({ type: 'renameSession', sessionId: sessionId, title: next, scope: this.scope });
        }
        this.refresh();
    },

    deleteSession: function (sessionId) {
        if (!sessionId) return;
        // 删除会连带清掉该会话的消息、记忆与运行轨迹，且不可撤销，所以必须二次确认
        if (!confirm('删除该会话及其全部历史记录？此操作不可撤销。')) return;
        this.sendMessageToVB({ type: 'deleteSession', sessionId: sessionId, scope: this.scope });
    },

    escapeHtml: function (text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    },

    formatSessionDate: function (createdAt) {
        if (!createdAt) return '未知时间';
        return String(createdAt).replace('T', ' ').substring(0, 19);
    },

    loadSession: function (sessionId) {
        this.sendMessageToVB({
            type: 'loadSession',
            sessionId: sessionId
        });
        this.closeSidebar();
    },

    // Send message to VB backend
    sendMessageToVB: function (message) {
        try {
            if (window.officeAi && typeof window.officeAi.post === 'function') {
                window.officeAi.post(message);
            } else if (window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage(message);
            } else if (window.vsto) {
                if (typeof window.vsto.sendMessage === 'function') {
                    window.vsto.sendMessage(JSON.stringify(message));
                } else if (typeof window.vsto.postMessage === 'function') {
                    window.vsto.postMessage(message);
                }
            } else {
                console.error('无法与后端通信');
            }
        } catch (error) {
            console.error('发送消息到VB后端失败:', error);
        }
    }
};

// Global function for VB backend to call
window.setHistoryFilesList = function (payload) {
    historyManager.displayHistoryFiles(payload);
};
