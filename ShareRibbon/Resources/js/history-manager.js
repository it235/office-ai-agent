/**
 * history-manager.js - History Sidebar and File Management
 * Handles chat history display and navigation
 */

window.historyManager = {
    isOpen: false,

    // Initialize history functionality
    init: function () {
        const toggleBtn = document.getElementById('history-toggle-btn');
        const sidebar = document.getElementById('history-sidebar');
        const overlay = document.getElementById('sidebar-overlay');
        const closeBtn = document.getElementById('close-sidebar-btn');

        // Bind events
        toggleBtn.addEventListener('click', () => this.toggleSidebar());
        closeBtn.addEventListener('click', () => this.closeSidebar());
        overlay.addEventListener('click', () => this.closeSidebar());

        // Keyboard event
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && this.isOpen) {
                this.closeSidebar();
            }
        });
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
    },

    // Load history files list
    loadHistoryFiles: function () {
        const historyList = document.getElementById('history-list');

        // Show loading state
        historyList.innerHTML = '<div class="loading-state">正在加载历史记录...</div>';

        // Request history files from backend
        this.sendMessageToVB({
            type: 'getHistoryFiles'
        });
    },

    // Display history files from backend
    displayHistoryFiles: function (files) {
        const historyList = document.getElementById('history-list');

        if (!files || files.length === 0) {
            historyList.innerHTML = `
                <div class="empty-state">
                    <div class="empty-state-icon">📄</div>
                    <div class="empty-state-text">您还没有任何历史记录</div>
                </div>
            `;
            return;
        }

        // Sort by filename (newest first)
        files.sort((a, b) => b.fileName.localeCompare(a.fileName));

        // Generate file list HTML
        const itemsHtml = files.map(file => `
            <div class="history-item" onclick="historyManager.openHistoryFile('${file.fullPath.replace(/\\/g, '\\\\')}')">
                <div class="history-item-title">${this.formatFileName(file.fileName)}</div>
                <div class="history-item-date">${this.formatFileDate(file.fileName)}</div>
                <div class="history-item-size">${this.formatFileSize(file.size)}</div>
            </div>
        `).join('');

        historyList.innerHTML = itemsHtml;
    },

    // Format filename for display
    formatFileName: function (fileName) {
        // Format: saved_chat_yyyyMMdd_HHmmss_中文内容.html or saved_chat_yyyyMMdd_HHmmss.html
        const match = fileName.match(/saved_chat_\d{8}_\d{6}_(.+)\.html/);
        if (match && match[1]) {
            return match[1];
        } else {
            return fileName.replace('saved_chat_', '').replace('.html', '');
        }
    },

    // Format file date from filename
    formatFileDate: function (fileName) {
        // Extract date time from filename
        const match = fileName.match(/saved_chat_(\d{8})_(\d{6})(?:_.*)?\.html/);
        if (match) {
            const dateStr = match[1]; // yyyyMMdd
            const timeStr = match[2]; // HHmmss

            const year = dateStr.substring(0, 4);
            const month = dateStr.substring(4, 6);
            const day = dateStr.substring(6, 8);
            const hour = timeStr.substring(0, 2);
            const minute = timeStr.substring(2, 4);
            const second = timeStr.substring(4, 6);

            return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
        }
        return '未知时间';
    },

    // Format file size
    formatFileSize: function (bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    },

    // Open history file
    openHistoryFile: function (filePath) {
        // Send to backend to open file
        this.sendMessageToVB({
            type: 'openHistoryFile',
            filePath: filePath
        });

        // Close sidebar
        this.closeSidebar();
    },

    // Send message to VB backend
    sendMessageToVB: function (message) {
        try {
            if (window.chrome && window.chrome.webview) {
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
window.setHistoryFilesList = function (files) {
    historyManager.displayHistoryFiles(files);
};
