/**
 * core.js - Core initialization for OfficeAI Chat
 * Marked.js configuration and renderer setup
 */

// Configure marked.js
marked.setOptions({
    highlight: function (code, lang) {
        if (lang && hljs.getLanguage(lang)) {
            return hljs.highlight(code, { language: lang }).value;
        }
        return hljs.highlightAuto(code).value;
    },
    breaks: true,
    gfm: true
});

// Extend marked renderer for code blocks with action buttons
const renderer = new marked.Renderer();
const originalCodeRenderer = renderer.code;

renderer.code = function (code, language, isEscaped) {
    const codeHtml = originalCodeRenderer.call(this, code, language, isEscaped);
    
    // 检测是否为可执行的代码类型
    const lang = (language || '').toLowerCase().trim();
    const executableLanguages = ['json', 'vba', 'vbnet', 'vbscript', 'javascript', 'js', 'excel', 'formula', 'function'];
    let isExecutable = executableLanguages.some(l => lang.includes(l));
    
    // 自动检测JSON格式（即使没有语言标记）
    if (!isExecutable && (!lang || lang === 'plaintext' || lang === 'text')) {
        const trimmed = code.trim();
        if ((trimmed.startsWith('{') && trimmed.endsWith('}')) ||
            (trimmed.startsWith('[') && trimmed.endsWith(']'))) {
            try {
                const parsed = JSON.parse(trimmed);
                if (parsed && (parsed.command || parsed.commands)) {
                    isExecutable = true;
                }
            } catch (e) {}
        }
    }
    
    // 只为可执行代码显示执行按钮
    const executeButton = isExecutable ? `
                <button class="code-button execute-button" onclick="executeCode(this)">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <polygon points="5 3 19 12 5 21 5 3"></polygon>
                    </svg>
                    执行
                </button>` : '';

    // Add action buttons to code blocks
    return `
        <div class="code-block">
            ${codeHtml}
            <div class="code-buttons">
                <button class="code-button copy-button" onclick="copyCode(this)">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                    </svg>
                    复制
                </button>
                <button class="code-button edit-button" onclick="editCode(this)">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                    </svg>
                    编辑
                </button>${executeButton}
            </div>
        </div>
    `;
};

marked.use({ renderer });

// Global state variables
window.rendererMap = {};
window.reasoningRendererMap = {};
window.userScrollPosition = 0;
window.autoScrollEnabled = true;
window.selectedContentMap = {};
window.attachedFiles = [];

// Predefined prompt suggestions
const predefinedPrompts = [
    "帮我把A列加B列的值写入C列",
    "帮我把Sheet1和Sheet2的表格按名字合并",
    "帮我把Sheet1的数据，按照中文名称拆分成多个xlsx文件",
    "给我将我选中的Word内容格式调整一下",
    "给我生成一个3页的周报PPT文件",
    "什么？没有你想要的，点击此处维护吧",
];

// DOM Content Loaded initialization
document.addEventListener('DOMContentLoaded', function () {
    // Initialize MCP button events
    document.getElementById('mcp-toggle-btn').addEventListener('click', toggleMcpDialog);
    document.getElementById('mcp-close-btn').addEventListener('click', closeMcpDialog);
    document.getElementById('mcp-overlay').addEventListener('click', closeMcpDialog);
    document.getElementById('mcp-save-btn').addEventListener('click', saveMcpSettings);

    // Initialize history manager
    historyManager.init();

    // Initialize clear context button
    document.getElementById('clear-context-btn').addEventListener('click', function () {
        document.getElementById('clear-or-delete-actions').style.display = 'block';
    });

    // Clear context memory action
    document.getElementById('action-clear-context').addEventListener('click', function () {
        document.getElementById('clear-or-delete-actions').style.display = 'none';
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage({ type: 'clearContext' });
        } else if (window.vsto && typeof window.vsto.postMessage === 'function') {
            window.vsto.postMessage({ type: 'clearContext' });
        } else {
            alert('无法清空上下文：未检测到支持的通信接口');
        }
    });

    // Delete chat records action
    document.getElementById('action-delete-chat').addEventListener('click', function () {
        document.getElementById('clear-or-delete-actions').style.display = 'none';
        showBatchDeleteChat();
    });

    // Cancel clear action
    document.getElementById('action-cancel').addEventListener('click', function () {
        document.getElementById('clear-or-delete-actions').style.display = 'none';
    });

    // Case header collapse/expand
    document.getElementById('case-header').addEventListener('click', function () {
        var content = document.getElementById('case-content');
        var arrow = document.getElementById('case-toggle-arrow').querySelector('svg');
        if (content.classList.contains('collapsed')) {
            content.classList.remove('collapsed');
            arrow.style.transform = 'rotate(0deg)';
        } else {
            content.classList.add('collapsed');
            arrow.style.transform = 'rotate(-90deg)';
        }
    });
});

// MCP connection check interval (every 3 seconds)
setInterval(function () {
    console.log('检测');
    requestMcpConnections();
}, 3000);
