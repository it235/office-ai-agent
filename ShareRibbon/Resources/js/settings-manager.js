/**
 * settings-manager.js - Settings Dialog Management
 * Handles settings dialog display, save, and cancel operations
 */

// Open settings dialog
function settingsButton() {
    document.getElementById('settings-overlay').style.display = 'block';
    document.getElementById('settings-dialog').style.display = 'block';
    // E3-05：同意行的勾选状态只由宿主下发。页面刚装载就打开设置时推送可能还没到，
    // 所以问一次，而不是先显示一个未经证实的默认值。
    if (!window.onboardingState && typeof sendMessageToServer === 'function') {
        sendMessageToServer({ type: 'getOnboardingState' });
    }
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

    // 自动补全设置
    let enableAutocomplete = document.getElementById('settings-autocomplete-enable').checked;
    let autocompleteShortcut = document.getElementById('settings-autocomplete-shortcut').value;

    // E4 原生入口三个开关 + 命令面板快捷键。少带一个键，后端就会把没提交的那一项当成默认值写回文件。
    let payload = {
        topicRandomness: topicRandomness,
        contextLimit: contextLimit,
        selectedCell: selectedCell,
        settingsScroll: settingsScroll,
        executeCodePreview: executeCodePreview,
        enableAutocomplete: enableAutocomplete,
        autocompleteShortcut: autocompleteShortcut,
        enableSelectionToolbar: document.getElementById('settings-enable-selection-toolbar').checked,
        enableContextMenu: document.getElementById('settings-enable-context-menu').checked,
        enableHotkey: document.getElementById('settings-enable-hotkey').checked,
        hotkeyChord: document.getElementById('settings-hotkey-chord').value
    };

    // Save settings to backend
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(Object.assign({ type: 'saveSettings' }, payload));
    } else if (window.vsto) {
        window.vsto.saveSettings(payload);
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
