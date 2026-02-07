/**
 * model-switcher.js - Model Switching Management
 * Handles model display and switching functionality
 */

/**
 * Open model configuration dialog
 * Sends message to VB.NET to open ConfigApiForm
 */
function openModelConfig() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'openApiConfigForm'
        });
    } else if (window.vsto) {
        window.vsto.openApiConfigForm();
    } else {
        alert('Unable to open configuration: communication interface not detected');
    }
}

/**
 * Update the current model display in the header bar
 * Called from VB.NET after model changes
 * @param {string} platform - The platform/provider name
 * @param {string} modelName - The model name
 */
function updateCurrentModelDisplay(platform, modelName) {
    var displayElement = document.getElementById('current-model-display');
    if (displayElement) {
        if (platform && modelName) {
            displayElement.textContent = platform + ' / ' + modelName;
        } else if (modelName) {
            displayElement.textContent = modelName;
        } else {
            displayElement.textContent = 'No model configured';
        }
    }
}

/**
 * Request current model info from VB.NET
 * Used to initialize the display on page load
 */
function requestCurrentModelInfo() {
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({
            type: 'getCurrentModel'
        });
    }
}

/**
 * Initialize model switcher on page load
 */
(function initModelSwitcher() {
    // Add hover effect to switch button
    var switchBtn = document.getElementById('switch-model-btn');
    if (switchBtn) {
        switchBtn.addEventListener('mouseenter', function() {
            this.style.background = 'rgba(255,255,255,0.35)';
        });
        switchBtn.addEventListener('mouseleave', function() {
            this.style.background = 'rgba(255,255,255,0.2)';
        });
    }

    // Request current model info when page loads
    // Small delay to ensure VB.NET communication is ready
    setTimeout(function() {
        requestCurrentModelInfo();
    }, 500);
})();
