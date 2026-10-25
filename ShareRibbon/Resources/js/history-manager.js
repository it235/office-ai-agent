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

            const emptyBtn = e.target.closest('[data-empty]');
            if (emptyBtn) {
                this.runEmptyAction(emptyBtn.getAttribute('data-empty'));
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

    // 空状态的引导动作：历史面板只负责查看，起点永远是输入框，
    // 这里不另开一条发送路径，免得又长出第二个入口。
    runEmptyAction: function (action) {
        if (action === 'onboarding') {
            this.closeSidebar();
            if (typeof window.showOnboarding === 'function') window.showOnboarding();
            return;
        }
        if (action === 'start') {
            this.closeSidebar();
            this.focusComposer();
        }
    },

    focusComposer: function () {
        const composer = document.getElementById('smart-input') || document.getElementById('chat-input');
        if (composer) composer.focus();
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

        // 顺序由后端决定：scope=all 时当前文档的会话排在最前，前端再按时间重排会把它冲掉。
        // 分组只是呈现：按「首次出现」切桶，桶内保持后端顺序。
        const groups = this.buildGroups();
        historyList.innerHTML = barHtml + groups.map(g => this.renderGroup(g)).join('');
        this.focusRenameInput();
    },

    // 按文档或时间切桶：scope=all 归到文档/通用会话，scope=document 只剩一篇文档、改按最近活跃分桶
    buildGroups: function () {
        const groups = [];
        const byKey = {};
        this.items.forEach(s => {
            const bucket = this.groupOf(s);
            let group = byKey[bucket.key];
            if (!group) {
                group = { key: bucket.key, label: bucket.label, isCurrent: bucket.isCurrent === true, items: [] };
                byKey[bucket.key] = group;
                groups.push(group);
            }
            group.items.push(s);
        });
        return groups;
    },

    groupOf: function (s) {
        if (this.scope === 'document') {
            const label = this.timeBucket(s);
            return { key: 'time:' + label, label: label };
        }
        // 分组键优先用文档哈希：同名不同文件的会话并成一组会把用户带偏
        const key = s.documentId || s.documentName || '';
        if (s.isCurrentDocument === true) {
            return { key: 'current:' + key, label: this.documentName || s.documentName || '本文档', isCurrent: true };
        }
        if (key) {
            return { key: 'doc:' + key, label: s.documentName || '其他文档' };
        }
        return { key: 'generic', label: '通用会话' };
    },

    // 分桶用排序键 lastActiveAt（老 payload 没这列时回落 createdAt），保证桶序与列表序一致
    timeBucket: function (s) {
        const matched = String(s.lastActiveAt || s.createdAt || '').match(/(\d{4})-(\d{2})-(\d{2})/);
        // 只截 YYYY-MM-DD 自己算天数差：带空格的 "2026-10-24 09:12" 不是标准 ISO，各引擎解析结果不一致
        if (!matched) return '更早';
        const day = Date.UTC(Number(matched[1]), Number(matched[2]) - 1, Number(matched[3]));
        const now = new Date();
        const elapsed = Math.round((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - day) / 86400000);
        if (elapsed <= 0) return '今天';
        if (elapsed === 1) return '昨天';
        if (elapsed <= 7) return '最近 7 天';
        if (elapsed <= 30) return '最近 30 天';
        return '更早';
    },

    renderGroup: function (group) {
        const header = `<div class="history-group-header${group.isCurrent ? ' history-group-current' : ''}">
            <span class="history-group-label" title="${this.escapeHtml(group.label)}">${this.escapeHtml(group.label)}</span>
            <span class="history-group-count">${group.items.length}</span>
        </div>`;
        return `<div class="history-group">${header}${group.items.map(s => this.renderItem(s)).join('')}</div>`;
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
        if (this.scope === 'document') {
            const text = this.hasDocument
                ? `《${this.escapeHtml(this.documentName || '未命名文档')}》还没有会话`
                : '这篇文档还没保存，会话只能先记成通用会话';
            const hint = this.hasDocument
                ? '关闭面板后直接提问，这次会话会自动归到当前文档。'
                : '保存文档后新会话就能按文档归组，已有会话可以在「全部」里找到。';
            return `<div class="empty-state">
                <div class="empty-state-text">${text}</div>
                <div class="empty-state-hint">${hint}</div>
                <div class="empty-state-actions">
                    <button type="button" class="history-empty-action" data-empty="start">直接开始</button>
                    <button type="button" class="history-empty-action" data-scope="all">查看全部会话</button>
                </div>
            </div>`;
        }
        const onboarding = typeof window.showOnboarding === 'function'
            ? '<button type="button" class="history-empty-action" data-empty="onboarding">新手引导</button>'
            : '';
        return `<div class="empty-state">
            <div class="empty-state-icon">📄</div>
            <div class="empty-state-text">您还没有任何历史会话</div>
            <div class="empty-state-hint">在下方输入框说清目标和范围，例如「把销售表按金额排序并高亮前 10 行」。每次任务都会自动记在这里。</div>
            <div class="empty-state-actions">
                <button type="button" class="history-empty-action" data-empty="start">开始第一个任务</button>
                ${onboarding}
            </div>
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
        // 文档名交给分组标题承载，条目只留最近活跃时间，免得同一行把组名再抄一遍
        return `<span class="history-meta-part">${this.escapeHtml(this.formatSessionDate(s.lastActiveAt || s.createdAt))}</span>`;
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
        // 到分钟就够：分组标题已经说明是哪一天，条目再挂秒数只会把行挤满
        return String(createdAt).replace('T', ' ').substring(0, 16);
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
