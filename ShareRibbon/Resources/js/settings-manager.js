/**
 * settings-manager.js - Settings Dialog Management
 * Handles settings dialog display, save, and cancel operations
 */

// Open settings dialog
function settingsButton() {
    document.getElementById('settings-overlay').style.display = 'block';
    document.getElementById('settings-dialog').style.display = 'block';
}

// Cancel settings dialog
function settingsCancel() {
    document.getElementById('settings-overlay').style.display = 'none';
    document.getElementById('settings-dialog').style.display = 'none';
}

// Save settings
function settingsSave() {
    let topicRandomness = document.getElementById('topic-randomness').value;
    let contextLimit = document.getElementById('context-limit').value;
    let settingsScroll = document.getElementById('settings-scroll-checked').checked;
    let selectedCell = document.getElementById('settings-selected-cell').checked;
    let executeCodePreview = document.getElementById('settings-executecode-preview').checked;
    let chatMode = document.getElementById("chatMode").value;
    
    // 自动补全设置
    let enableAutocomplete = document.getElementById('settings-autocomplete-enable').checked;
    let autocompleteShortcut = document.getElementById('settings-autocomplete-shortcut').value;

    // Save settings to backend
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'saveSettings',
            topicRandomness: topicRandomness,
            contextLimit: contextLimit,
            selectedCell: selectedCell,
            settingsScroll: settingsScroll,
            chatMode: chatMode,
            executeCodePreview: executeCodePreview,
            enableAutocomplete: enableAutocomplete,
            autocompleteShortcut: autocompleteShortcut,
        });
    } else if (window.vsto) {
        window.vsto.saveSettings({
            topicRandomness: topicRandomness,
            contextLimit: contextLimit,
            selectedCell: selectedCell,
            settingsScroll: settingsScroll,
            chatMode: chatMode,
            executeCodePreview: executeCodePreview,
            enableAutocomplete: enableAutocomplete,
            autocompleteShortcut: autocompleteShortcut,
        });
    } else {
        alert('无法执行代码：未检测到支持的通信接口');
    }
    
    // 更新前端自动补全状态
    if (typeof updateAutocompleteSettings === 'function') {
        updateAutocompleteSettings({ enabled: enableAutocomplete, shortcut: autocompleteShortcut });
    }

    // Close dialog
    document.getElementById('settings-overlay').style.display = 'none';
    document.getElementById('settings-dialog').style.display = 'none';
}

// Initialize settings slider event handlers
(function initSettingsSliders() {
    // Topic randomness slider
    document.getElementById('topic-randomness').oninput = function () {
        document.getElementById('topic-randomness-value').textContent =
            Number(this.value).toFixed(1);
    };

    // Context limit slider
    document.getElementById('context-limit').oninput = function () {
        document.getElementById('context-limit-value').textContent = this.value;
    };

    // Overlay click to close dialog
    document.getElementById('settings-overlay').onclick = function () {
        document.getElementById('settings-overlay').style.display = 'none';
        document.getElementById('settings-dialog').style.display = 'none';
    };
})();
