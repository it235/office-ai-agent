Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' ShareRibbon\Controls\BaseDoubaoChat.vb
Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms.ListBox
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Markdig
Imports Microsoft.Web.WebView2.WinForms

Public MustInherit Class BaseDoubaoChat
    Inherits BaseChat

    Public Overrides ReadOnly Property ChatUrl As String = "https://www.doubao.com"
    Public Overrides ReadOnly Property SessionFileName As String = "doubao_session.json"

    Protected Async Function InitializeWebView2() As Task
        Try
            ' 使用固定的用户数据目录而不是临时目录，以保持会话持久化
            Dim userDataFolder As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        ConfigSettings.OfficeAiAppDataFolder,
        "DoubaoChatWebView2Data")

            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            ' 创建环境选项，使用持久化的用户数据文件夹
            Dim options As New CoreWebView2EnvironmentOptions()
            options.AdditionalBrowserArguments = "--no-sandbox"

            ' 创建WebView2环境，使用固定目录保持会话
            Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, options)

            ' 初始化WebView2
            Await ChatBrowser.EnsureCoreWebView2Async(env)

            ' 配置WebView2
            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                ' 确保WebView2可以接收焦点
                ChatBrowser.TabStop = True
                ChatBrowser.TabIndex = 1
                ChatBrowser.Visible = True
                ChatBrowser.Enabled = True
                ChatBrowser.BringToFront()

                ' 配置WebView2设置以改善焦点行为
                ChatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                ChatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
                ' 启用开发者工具以便调试可能的焦点问题
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True

                ' 重要：在导航前注册所有事件处理器
                AddHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived

                ' 启用持久化的Cookie管理
                ChatBrowser.CoreWebView2.CookieManager.DeleteAllCookies() ' 可选，仅在需要清理时使用

                ' 导航到目标网站
                ChatBrowser.CoreWebView2.Navigate(ChatUrl)

                Debug.WriteLine($"WebView2初始化完成，开始导航到{ChatUrl}")
            Else
                MessageBox.Show("WebView2初始化失败，CoreWebView2不可用。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"初始化失败: {ex.Message}{Environment.NewLine}类型: {ex.GetType().Name}{Environment.NewLine}堆栈:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

' Doubao特有的焦点处理方法
    Protected Overrides Function GetScriptInjectionDelay() As Integer
        Return 2000 ' Doubao需要更长的延迟
    End Function

    Protected Overrides Async Function InitializeWebView2ScriptAsyncSafe() As Task
        Try
            Dim script As String = "
(function(){
    try {
        // 初始化VSTO接口（不依赖页面环境）
        window.vsto = {
            executeCode: function(code, language, preview) {
                console.log('[VSTO] executeCode被调用:', {code: (code||'').substring(0,50) + '...', language: language, preview: preview});
                try { window.chrome.webview.postMessage({ type: 'executeCode', code: code, language: language, executecodePreview: preview }); } catch(e) { console.error(e); }
                return true;
            },
            checkedChange: function(thisProperty, checked) {
                try { return window.chrome.webview.postMessage({ type: 'checkedChange', isChecked: checked, property: thisProperty }); } catch(e) { console.error(e); }
            },
            sendMessage: function(payload) {
                try {
                    let messageToSend = (typeof payload === 'string') ? { type: 'sendMessage', value: payload } : payload;
                    window.chrome.webview.postMessage(messageToSend);
                } catch(e) { console.error(e); }
                return true;
            },
            saveSettings: function(settingsObject) {
                try {
                    return window.chrome.webview.postMessage({
                        type: 'saveSettings',
                        topicRandomness: settingsObject.topicRandomness,
                        contextLimit: settingsObject.contextLimit,
                        selectedCell: settingsObject.selectedCell,
                        executeCodePreview: settingsObject.executeCodePreview,
                    });
                } catch(e) { console.error(e); }
            }
        };
        console.log('[VSTO] 基础API已初始化');

        // 验证通信接口
        if (window.chrome && window.chrome.webview) {
            console.log('[VSTO] ✓ chrome.webview接口可用');
        } else {
            console.log('[VSTO] ✗ chrome.webview接口不可用');
        }

        // Overlay（蒙层）检测与禁用 — 尝试解除覆盖导致的焦点问题
        function disableOverlays() {
            try {
                const els = Array.from((document.body || document.documentElement).querySelectorAll('*')).filter(el => {
                    const style = window.getComputedStyle(el);
                    if (!(style.position === 'fixed' || style.position === 'absolute')) return false;
                    const rect = el.getBoundingClientRect();
                    // 只处理覆盖大部分视口的元素
                    if (rect.width < window.innerWidth * 0.5 || rect.height < window.innerHeight * 0.5) return false;
                    const z = parseInt(style.zIndex) || 0;
                    // 只处理高 z-index 的覆盖层
                    if (z < 1000) return false;
                    // 排除自身可交互控制，例如真正的侧栏等（根据role/aria-hidden判断）
                    if (el.getAttribute && el.getAttribute('role') === 'dialog' && el.getAttribute('aria-hidden') === 'false') return false;
                    return true;
                });
                els.forEach(el => {
                    try {
                        if (!el.dataset.vstoOverlayDisabled) {
                            el.dataset.vstoOverlayDisabled = '1';
                            el.__vsto_oldPointerEvents = el.style.pointerEvents || '';
                            el.style.pointerEvents = 'none';
                            // 保守处理：仅当明显是透明遮罩时隐藏
                            const bg = window.getComputedStyle(el).backgroundColor || '';
                            const opacity = window.getComputedStyle(el).opacity || '1';
                            if ((bg.indexOf('rgba') === 0 && parseFloat(opacity) < 0.98) || parseFloat(opacity) < 0.5) {
                                el.__vsto_oldDisplay = el.style.display || '';
                                el.style.display = 'none';
                            }
                            console.log('[VSTO] disabled overlay', el);
                        }
                    } catch (inner) { console.error(inner); }
                });
            } catch (e) { console.error('[VSTO] disableOverlays error', e); }
        }

        // 观察 DOM 以在动态添加覆盖层时禁用
        const overlayObserver = new MutationObserver(function(muts) {
            try {
                let shouldRun = false;
                for (const m of muts) {
                    if (m.addedNodes && m.addedNodes.length) { shouldRun = true; break; }
                    if (m.type === 'attributes' && (m.attributeName === 'style' || m.attributeName === 'class' || m.attributeName === 'aria-hidden')) { shouldRun = true; break; }
                }
                if (shouldRun) disableOverlays();
            } catch (e) { console.error(e); }
        });
        try {
            overlayObserver.observe(document.documentElement || document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['style','class','aria-hidden'] });
        } catch (e) { console.error(e); }

        // 尝试聚焦：优先聚焦可编辑/输入区域，其次 body
        function ensureFocus() {
            try {
                disableOverlays();
                const focusable = document.querySelector('[contenteditable], input, textarea, [tabindex]:not([tabindex=\"-1\"]), button, a');
                if (focusable && typeof focusable.focus === 'function') {
                    focusable.focus();
                    console.log('[VSTO] focused', focusable);
                    return;
                }
                // 兜底：确保 body 可聚焦
                if (!document.body.hasAttribute('tabindex')) document.body.setAttribute('tabindex', '0');
                document.body.focus();
                console.log('[VSTO] focused body');
            } catch (e) { console.error('[VSTO] ensureFocus error', e); }
        }

        // DOMContentLoaded + 周期性修复 + 点击修复
        document.addEventListener('DOMContentLoaded', function() {
            try {
                disableOverlays();
                ensureFocus();
                setInterval(function() { ensureFocus(); }, 2000);
                document.addEventListener('click', function() { setTimeout(ensureFocus, 50); }, true);
            } catch(e) { console.error(e); }
        });

        // 立即尝试一次（若 DOMContentLoaded 已经触发）
        disableOverlays();
        ensureFocus();

    } catch (e) {
        console.error('[VSTO] Initialize script top-level error:', e);
    }
});
"
 
            If ChatBrowser.CoreWebView2 IsNot Nothing Then

                Dim escapedInit = JsonConvert.SerializeObject(script)
                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedInit})")
                Debug.WriteLine("InitializeWebView2ScriptAsync执行完成")
            Else
                Debug.WriteLine("InitializeWebView2ScriptAsync: CoreWebView2为空")
        End If
    Catch ex As Exception
        Debug.WriteLine($"InitializeWebView2ScriptAsync出错: {ex.Message}")
    End Try
End Function

    ' 在页面加载完成后，注入脚本 - 修复线程问题
    Private Async Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If e.IsSuccess Then
            Try
                Debug.WriteLine("导航完成，开始注入脚本")

                ' 延迟一些时间，确保页面完全加载
                Await Task.Delay(2000)

                ' 确保在UI线程上执行所有WebView2操作
                If ChatBrowser.InvokeRequired Then
                    ChatBrowser.Invoke(New Action(Async Sub()
                                                      Try
                                                          ' 注入基础辅助脚本
                                                          Await InitializeWebView2ScriptAsyncSafe()

                                                          ' 初始化Doubao执行按钮
                                                          Await InjectExecuteButtonsSafe()

                                                          Debug.WriteLine("所有脚本注入完成")
                                                      Catch ex As Exception
                                                          Debug.WriteLine($"UI线程脚本注入出错: {ex.Message}")
                                                          Debug.WriteLine(ex.StackTrace)
                                                      End Try
                                                  End Sub))
                Else
                    ' 已经在UI线程，直接执行
                    Try
                        ' 注入基础辅助脚本
                        Await InitializeWebView2ScriptAsyncSafe()

                        ' 初始化Doubao执行按钮
                        Await InjectExecuteButtonsSafe()

                        Debug.WriteLine("所有脚本注入完成")
                    Catch ex As Exception
                        Debug.WriteLine($"脚本注入出错: {ex.Message}")
                        Debug.WriteLine(ex.StackTrace)
                    End Try
                End If
            Catch ex As Exception
                Debug.WriteLine($"导航完成事件处理中出错: {ex.Message}")
                Debug.WriteLine(ex.StackTrace)
            End Try
        Else
            Debug.WriteLine($"导航失败: {e.WebErrorStatus}")
        End If
    End Sub

    ' 线程安全的InitializeWebView2ScriptAsync方法
    'Private Async Function InitializeWebView2ScriptAsyncSafe() As Task
    '    Try
    '        Dim script As String = "
    '(function(){
    '    try {
    '        // 初始化VSTO接口（不依赖页面环境）
    '        window.vsto = {
    '            executeCode: function(code, language, preview) {
    '                console.log('[VSTO] executeCode被调用:', {code: (code||'').substring(0,50) + '...', language: language, preview: preview});
    '                try { window.chrome.webview.postMessage({ type: 'executeCode', code: code, language: language, executecodePreview: preview }); } catch(e) { console.error(e); }
    '                return true;
    '            },
    '            checkedChange: function(thisProperty, checked) {
    '                try { return window.chrome.webview.postMessage({ type: 'checkedChange', isChecked: checked, property: thisProperty }); } catch(e) { console.error(e); }
    '            },
    '            sendMessage: function(payload) {
    '                try {
    '                    let messageToSend = (typeof payload === 'string') ? { type: 'sendMessage', value: payload } : payload;
    '                    window.chrome.webview.postMessage(messageToSend);
    '                } catch(e) { console.error(e); }
    '                return true;
    '            },
    '            saveSettings: function(settingsObject) {
    '                try {
    '                    return window.chrome.webview.postMessage({
    '                        type: 'saveSettings',
    '                        topicRandomness: settingsObject.topicRandomness,
    '                        contextLimit: settingsObject.contextLimit,
    '                        selectedCell: settingsObject.selectedCell,
    '                        executeCodePreview: settingsObject.executeCodePreview,
    '                    });
    '                } catch(e) { console.error(e); }
    '            }
    '        };
    '        console.log('[VSTO] 基础API已初始化');

    '        // 验证通信接口
    '        if (window.chrome && window.chrome.webview) {
    '            console.log('[VSTO] ✓ chrome.webview接口可用');
    '        } else {
    '            console.log('[VSTO] ✗ chrome.webview接口不可用');
    '        }

    '        // Overlay（蒙层）检测与禁用 — 尝试解除覆盖导致的焦点问题
    '        function disableOverlays() {
    '            try {
    '                const els = Array.from((document.body || document.documentElement).querySelectorAll('*')).filter(el => {
    '                    const style = window.getComputedStyle(el);
    '                    if (!(style.position === 'fixed' || style.position === 'absolute')) return false;
    '                    const rect = el.getBoundingClientRect();
    '                    // 只处理覆盖大部分视口的元素
    '                    if (rect.width < window.innerWidth * 0.5 || rect.height < window.innerHeight * 0.5) return false;
    '                    const z = parseInt(style.zIndex) || 0;
    '                    // 只处理高 z-index 的覆盖层
    '                    if (z < 1000) return false;
    '                    // 排除自身可交互控制，例如真正的侧栏等（根据role/aria-hidden判断）
    '                    if (el.getAttribute && el.getAttribute('role') === 'dialog' && el.getAttribute('aria-hidden') === 'false') return false;
    '                    return true;
    '                });
    '                els.forEach(el => {
    '                    try {
    '                        if (!el.dataset.vstoOverlayDisabled) {
    '                            el.dataset.vstoOverlayDisabled = '1';
    '                            el.__vsto_oldPointerEvents = el.style.pointerEvents || '';
    '                            el.style.pointerEvents = 'none';
    '                            // 保守处理：仅当明显是透明遮罩时隐藏
    '                            const bg = window.getComputedStyle(el).backgroundColor || '';
    '                            const opacity = window.getComputedStyle(el).opacity || '1';
    '                            if ((bg.indexOf('rgba') === 0 && parseFloat(opacity) < 0.98) || parseFloat(opacity) < 0.5) {
    '                                el.__vsto_oldDisplay = el.style.display || '';
    '                                el.style.display = 'none';
    '                            }
    '                            console.log('[VSTO] disabled overlay', el);
    '                        }
    '                    } catch (inner) { console.error(inner); }
    '                });
    '            } catch (e) { console.error('[VSTO] disableOverlays error', e); }
    '        }

    '        // 观察 DOM 以在动态添加覆盖层时禁用
    '        const overlayObserver = new MutationObserver(function(muts) {
    '            try {
    '                let shouldRun = false;
    '                for (const m of muts) {
    '                    if (m.addedNodes && m.addedNodes.length) { shouldRun = true; break; }
    '                    if (m.type === 'attributes' && (m.attributeName === 'style' || m.attributeName === 'class' || m.attributeName === 'aria-hidden')) { shouldRun = true; break; }
    '                }
    '                if (shouldRun) disableOverlays();
    '            } catch (e) { console.error(e); }
    '        });
    '        try {
    '            overlayObserver.observe(document.documentElement || document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['style','class','aria-hidden'] });
    '        } catch (e) { console.error(e); }

    '        // 尝试聚焦：优先聚焦可编辑/输入区域，其次 body
    '        function ensureFocus() {
    '            try {
    '                disableOverlays();
    '                const focusable = document.querySelector('[contenteditable], input, textarea, [tabindex]:not([tabindex=\"-1\"]), button, a');
    '                if (focusable && typeof focusable.focus === 'function') {
    '                    focusable.focus();
    '                    console.log('[VSTO] focused', focusable);
    '                    return;
    '                }
    '                // 兜底：确保 body 可聚焦
    '                if (!document.body.hasAttribute('tabindex')) document.body.setAttribute('tabindex', '0');
    '                document.body.focus();
    '                console.log('[VSTO] focused body');
    '            } catch (e) { console.error('[VSTO] ensureFocus error', e); }
    '        }

    '        // DOMContentLoaded + 周期性修复 + 点击修复
    '        document.addEventListener('DOMContentLoaded', function() {
    '            try {
    '                disableOverlays();
    '                ensureFocus();
    '                setInterval(function() { ensureFocus(); }, 2000);
    '                document.addEventListener('click', function() { setTimeout(ensureFocus, 50); }, true);
    '            } catch(e) { console.error(e); }
    '        });

    '        // 立即尝试一次（若 DOMContentLoaded 已经触发）
    '        disableOverlays();
    '        ensureFocus();

    '    } catch (e) {
    '        console.error('[VSTO] Initialize script top-level error:', e);
    '    }
    '})();
    '"

    '            If ChatBrowser.CoreWebView2 IsNot Nothing Then

    '                Dim escapedInit = JsonConvert.SerializeObject(script)
    '                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedInit})")
    '                Debug.WriteLine("InitializeWebView2ScriptAsync执行完成")
    '            Else
    '                Debug.WriteLine("InitializeWebView2ScriptAsync: CoreWebView2为空")
    '        End If
    '    Catch ex As Exception
    '        Debug.WriteLine($"InitializeWebView2ScriptAsync出错: {ex.Message}")
    '    End Try
    'End Function

    ' 线程安全的执行按钮注入方法
    Protected Overrides Async Function InjectExecuteButtonsSafe() As Task
        Try
            Await InjectDoubaoExecuteButtonsSafe()
            Debug.WriteLine("InitializeDoubaoButtons执行完成")
        Catch ex As Exception
            Debug.WriteLine($"InitializeDoubaoButtons出错: {ex.Message}")
        End Try
    End Function

    ' Doubao特定的执行按钮注入方法
    Private Async Function InjectDoubaoExecuteButtonsSafe() As Task
        Try
            If ChatBrowser.CoreWebView2 Is Nothing Then
                Debug.WriteLine("InjectDoubaoExecuteButtons: CoreWebView2未初始化")
                Return
            End If

            Dim script As String = "
    (function() {
        console.log('[Doubao Execute Buttons] 注入开始');

        if (window.__doubaoExecuteButtonsInitialized) {
            console.log('[Doubao Execute Buttons] 已初始化，刷新');
            if (window.refreshDoubaoExecuteButtons) window.refreshDoubaoExecuteButtons();
            return;
        }
        window.__doubaoExecuteButtonsInitialized = true;

        function createExecuteSVG() {
            return '<svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" fill=""none"" viewBox=""0 0 24 24""><path fill=""currentColor"" d=""M8 5v14l11-7z""/></svg>';
        }

        // 帮助函数：检查元素是否包含以 prefix 开头的类
        function hasClassWithPrefix(el, prefix) {
            try {
                return Array.from(el.classList).some(c => c.indexOf(prefix) === 0);
            } catch (e) { return false; }
        }

        // 在一个节点内部查找可能的代码容器（优先 class 前缀 code-area-，其次 pre/code）
        function findCodeAreaWithin(root) {
            let el = root.querySelector('[class*=""code-area-""]');
            if (el) return el;
            el = root.querySelector('pre, code, [data-code], [data-language]');
            return el;
        }

        // 查找页面上可能的代码块容器（使用多个宽泛选择器，然后通过相对路径确认）
        function findDoubaoCodeBlocks() {
            const broadSelectors = ['[class*=""code-block""]', '[class*=""custom-code""]', 'pre', 'code', '[data-code]'];
            const candidates = new Set();
            broadSelectors.forEach(sel => {
                Array.from(document.querySelectorAll(sel)).forEach(n => candidates.add(n));
            });

            const codeBlocks = [];
            candidates.forEach(el => {
                // 如果自身或子孙包含代码区域则认为是代码块容器
                const codeArea = findCodeAreaWithin(el);
                if (codeArea && (codeArea.textContent || '').trim().length > 0) {
                    codeBlocks.push(el);
                } else {
                    // 尝试向上查找包含代码区域的祖先（处理嵌套结构）
                    const up = el.closest('[class*=""code-block""], [class*=""custom-code""], section, article, div');
                    if (up && findCodeAreaWithin(up)) codeBlocks.push(up);
                }
            });
            return codeBlocks;
        }

        // 尝试从 codeBlock 中获取代码文本与语言
        function getDoubaoCodeContent(codeBlock) {
    try {
        // 优先寻找真正的代码元素（pre 或 code），避免取到 header/按钮文本
        var codeArea = findCodeAreaWithin(codeBlock);
        var codeElem = null;
        if (codeArea) {
            codeElem = codeArea.querySelector('pre, code') || codeArea.querySelector('.content-y8qlFa pre, .content-y8qlFa code');
        } else {
            codeElem = codeBlock.querySelector('pre, code');
        }

        // 如果找不到 pre/code，退回到对容器文本的保守提取并去掉 UI 标签前缀
        if (!codeElem) {
            var raw = (codeBlock.textContent || '').trim();
            raw = raw.replace(/^\s*执行\s*/i, '');
            raw = raw.replace(/^\s*(?:vba|javascript|js|python|sql|excel|html|css|java|php)\s*/i, '');
            return { code: raw, language: 'unknown' };
        }

        var codeContent = (codeElem.textContent || '').trim();
        var language = 'unknown';

        // 先尝试从 codeElem 的类名中解析 language-xxx
        var m = (codeElem.className || '').match(/language-([a-zA-Z0-9#+-]+)/);
        if (m && m[1]) {
            language = m[1].toLowerCase();
        }

        // 若仍未知，尝试从标题或 data-language 属性获取
        if (language === 'unknown') {
            var titleElement = codeBlock.querySelector('[class*=\""title-\""], [data-language], .title, h1,h2,h3,h4,h5,h6');
            if (titleElement) {
                var langText = (titleElement.getAttribute('data-language') || titleElement.textContent || '').trim().toLowerCase();
                if (langText) {
                    // 只取识别出的语言关键字
                    ['vba','vb','javascript','js','excel','python','sql','typescript','html','css','c#','java','php','formula','function'].forEach(function(k){
                        if (langText.indexOf(k) !== -1) language = k;
                    });
                }
            }
        }

        return { code: codeContent, language: language };
    } catch (e) {
        console.error('[Doubao Execute Buttons] 获取代码失败', e);
        return { code: '', language: 'unknown' };
    }
}

        function createDoubaoExecuteButton() {
    // 使用站点风格的类名和内联样式，尽量与原始 SVG/尺寸一致
    var wrapper = document.createElement('div');
    wrapper.className = 'vsto-doubao-execute-button hoverable-kRHiX2 vsto-exec';
    wrapper.setAttribute('role','button');
    wrapper.setAttribute('aria-disabled','false');
    wrapper.tabIndex = 0;
    // 尽量不破坏页面样式，使用透明背景并微调内边距，使它看起来像页面上的其他按钮
    wrapper.style.cssText = 'display:inline-flex;align-items:center;margin-right:6px;cursor:pointer;user-select:none;padding:6px;border-radius:6px;background:transparent;color:inherit;font-size:12px;border:0;width:30px;';
    // 用与页面相近的 SVG（play 图标），但不强行覆盖页面样式
    wrapper.innerHTML = '<svg xmlns=""http://www.w3.org/2000/svg"" width=""16"" height=""16"" viewBox=""0 0 24 24"" fill=""none"" style=""flex:0 0 16px;""><path fill=""currentColor"" d=""M8 5v14l11-7z""></path></svg><span style=""margin-left:6px;line-height:16px;font-size:12px;"">执行</span>';
    return wrapper;
}


        function attachDoubaoClickHandler(executeBtn, codeBlock) {
            const handler = function(e) {
                e.preventDefault();
                e.stopPropagation();
                console.log('[Doubao Execute Buttons] 执行按钮被点击');
                
                const content = getDoubaoCodeContent(codeBlock);
                if (!content.code || !content.code.trim()) {
                    console.log('[Doubao Execute Buttons] 代码为空，跳过');
                    return;
                }
                
                try {
                    if (window.vsto && typeof window.vsto.executeCode === 'function') {
                        window.vsto.executeCode(content.code, content.language, true);
                        console.log('[Doubao Execute Buttons] 通过 vsto 执行请求发送');
                    } else if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {
                        window.chrome.webview.postMessage({
                            type: 'executeCode',
                            code: content.code,
                            language: content.language,
                            executecodePreview: true
                        });
                        console.log('[Doubao Execute Buttons] 通过 chrome.webview 发送执行请求');
                    } else {
                        console.error('[Doubao Execute Buttons] 无通信接口');
                    }
                } catch (err) {
                    console.error('[Doubao Execute Buttons] 发送执行请求失败', err);
                }
            };
            
            executeBtn.addEventListener('click', handler);
            executeBtn.addEventListener('keydown', function(ev) {
                if (ev.key === 'Enter' || ev.key === ' ') {
                    ev.preventDefault();
                    this.click();
                }
            });
        }

        function findActionArea(codeBlock) {
    if (!codeBlock) return null;

    // helper：在 root 内查找第一个类名以 prefix 开头的元素
    function findDescByPrefix(root, prefix) {
        if (!root) return null;
        var nodes = root.querySelectorAll('[class]');
        for (var i = 0; i < nodes.length; i++) {
            var cl = nodes[i].className || '';
            // 将 class 字符串拆分成 token，再检查每个 token 是否以 prefix 开头
            var tokens = cl.split(/\s+/);
            for (var j = 0; j < tokens.length; j++) {
                if (tokens[j].indexOf(prefix) === 0) return nodes[i];
            }
        }
        return null;
    }

    // 1) 优先直接寻找复制按钮（最可靠）
    var copyBtn = codeBlock.querySelector('[data-testid=""code-block-copy""], [data-testid=""code-block-copy-button""]');
    if (copyBtn) {
        // 尝试定位与 copyBtn 同级的 title-*（通常 title 与 action 同在 header 内）
        var headerAncestor = copyBtn.closest('[class*=""header-""], [class*=""header-wrapper""], [class*=""code-area-""], [class*=""code-block""], section, article, div');
        if (headerAncestor) {
            var titleSib = findDescByPrefix(headerAncestor, 'title-') || headerAncestor.querySelector('[class*=""title-""], .title');
            if (titleSib) {
                // 要把执行按钮放到 title 的第一个位置
                return { container: titleSib, beforeNode: titleSib.firstElementChild || null };
            }
        }

        // 如果没有 title-*，尝试把按钮插入到包含 copy 的 action-* 容器里，位于 copyBtn 之前
        var actionAncestor = copyBtn.closest('[class*=""action-""], [class*=""actions""], [class*=""toolbar""], [class*=""action-""]');
        if (actionAncestor) {
            return { container: actionAncestor, beforeNode: copyBtn };
        }

        // 回退：直接在 copyBtn 前插入
        return { copyBtn: copyBtn };
    }

    // 2) 若找不到复制按钮，则按原有回退策略：优先 title-，其次 header -> action，再兜底
    var titleEl = findDescByPrefix(codeBlock, 'title-') || codeBlock.querySelector('.title');
    if (titleEl) {
        return { container: titleEl, beforeNode: titleEl.firstElementChild || null };
    }

    var headerEl = findDescByPrefix(codeBlock, 'header-') || codeBlock.querySelector('[class*=""header-""], .header, [class*=""header-wrapper""]');
    if (headerEl) {
        var actionInHeader = findDescByPrefix(headerEl, 'action-') || headerEl.querySelector('[class*=""action-""], [class*=""actions""], [role=""toolbar""]');
        if (actionInHeader) {
            var copyBtn2 = actionInHeader.querySelector('[data-testid=""code-block-copy""], [data-testid=""code-block-copy-button""]');
            if (copyBtn2) return { copyBtn: copyBtn2 };
            return { container: actionInHeader, beforeNode: actionInHeader.firstElementChild || null };
        }
        var titleInHeader = findDescByPrefix(headerEl, 'title-') || headerEl.querySelector('.title');
        if (titleInHeader) return { container: titleInHeader, beforeNode: titleInHeader.firstElementChild || null };
    }

    var actionAny = findDescByPrefix(codeBlock, 'action-') || codeBlock.querySelector('[class*=""action-""], [class*=""actions""], [role=""toolbar""]');
    if (actionAny) {
        var copyAny = actionAny.querySelector('[data-testid=""code-block-copy""], [data-testid=""code-block-copy-button""]');
        if (copyAny) return { copyBtn: copyAny };
        return { container: actionAny, beforeNode: actionAny.firstElementChild || null };
    }

    var copyAnywhere = codeBlock.querySelector('[data-testid=""code-block-copy""], [data-testid=""code-block-copy-button""]');
    if (copyAnywhere) return { copyBtn: copyAnywhere };

    var hover = codeBlock.querySelector('.hoverable-kRHiX2, [role=""button""], button');
    if (hover && hover.parentElement) return { container: hover.parentElement, beforeNode: hover };

    return null;
}

        function processDoubaoCodeBlock(codeBlock, index) {
    try {
        if (!codeBlock) return false;
        // 幂等：如果已插入则跳过
        if (codeBlock.querySelector('.vsto-doubao-execute-button')) return false;

        var actionInfo = findActionArea(codeBlock);
        if (!actionInfo) return false;

        console.log('[Doubao Execute Buttons] 创建执行按钮');
        var execBtn = createDoubaoExecuteButton();

        // 精确插入：若找到了 copyBtn，直接在 copyBtn 之前插入（使用 insertAdjacentElement）
        if (actionInfo.copyBtn) {
            try {
                actionInfo.copyBtn.insertAdjacentElement('beforebegin', execBtn);
            } catch (e) {
                if (actionInfo.copyBtn.parentNode) actionInfo.copyBtn.parentNode.insertBefore(execBtn, actionInfo.copyBtn);
            }
        }
        // 若找到 container 与 beforeNode 则使用
        else if (actionInfo.container && actionInfo.beforeNode) {
            actionInfo.container.insertBefore(execBtn, actionInfo.beforeNode);
        }
        // 只有 container 时插入到最前面（更靠近子 header）
        else if (actionInfo.container) {
            actionInfo.container.insertBefore(execBtn, actionInfo.container.firstChild);
        } else {
            // 兜底：尝试插入到 header-wrapper 的最前面
            var headerWrapper = codeBlock.querySelector('.header-wrapper-Mbk8s6') || codeBlock.querySelector('[class*=""header-""], .header, .toolbar');
            if (headerWrapper) headerWrapper.insertBefore(execBtn, headerWrapper.firstChild);
            else return false;
        }

        // 防护：确保不会被页面禁用
        execBtn.style.pointerEvents = 'auto';
        execBtn.removeAttribute('disabled');
        execBtn.setAttribute('aria-disabled','false');

        // 绑定事件
        attachDoubaoClickHandler(execBtn, codeBlock);

        console.log('[Doubao Execute Buttons] 执行按钮插入完成');
        return true;
    } catch (ex) {
        console.error('[Doubao Execute Buttons] 处理代码块失败', ex);
        return false;
    }
}

        function addDoubaoExecuteButtons() {
            const codeBlocks = findDoubaoCodeBlocks();
            if (!codeBlocks || codeBlocks.length === 0) return;
            let count = 0;
            codeBlocks.forEach((block, i) => { if (processDoubaoCodeBlock(block, i)) count++; });
            console.log('[Doubao Execute Buttons] 处理完成: ' + count + '/' + codeBlocks.length);
        }

        const observer = new MutationObserver(function(mutations) {
            let shouldRun = false;
            for (const m of mutations) {
                if (m.addedNodes && m.addedNodes.length > 0) { shouldRun = true; break; }
                if (m.type === 'attributes' && (m.attributeName === 'class' || m.attributeName === 'style' || m.attributeName === 'aria-hidden')) { shouldRun = true; break; }
            }
            if (shouldRun) setTimeout(addDoubaoExecuteButtons, 200);
        });
        
        observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class','style','aria-hidden'] });

        setTimeout(addDoubaoExecuteButtons, 200);
        [1000,2000,3000,5000].forEach((d,i) => setTimeout(addDoubaoExecuteButtons, d));
        
        window.refreshDoubaoExecuteButtons = addDoubaoExecuteButtons;

        console.log('[Doubao Execute Buttons] 注入完成');
    })();
    "
            Dim escapedButtons = JsonConvert.SerializeObject(script)
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedButtons})")
            Debug.WriteLine("Doubao执行按钮注入脚本已执行")
        Catch ex As Exception
            Debug.WriteLine($"注入Doubao执行按钮脚本时出错: {ex.Message}")
        End Try
    End Function

    ' 执行JavaScript代码
    Protected Overrides Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean
        Try
            ' 获取Office应用对象
            Dim appObject As Object = GetOfficeApplicationObject()
            If appObject Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取Office应用程序对象")
                Return False
            End If

            ' 创建脚本控制引擎
            Dim scriptEngine As Object = CreateObject("MSScriptControl.ScriptControl")
            scriptEngine.Language = "JScript"

            ' 将Office应用对象暴露给脚本环境
            scriptEngine.AddObject("app", appObject, True)

            ' 执行JavaScript代码
            Dim result = scriptEngine.Eval(jsCode)

            If result IsNot Nothing Then
                GlobalStatusStrip.ShowInfo("JavaScript执行完成，结果: " & result.ToString())
            Else
                GlobalStatusStrip.ShowInfo("JavaScript执行完成")
            End If

            Return True
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("执行JavaScript代码时出错: " & ex.Message)
            Return False
        End Try
    End Function




    ' JavaScript字符串转义辅助函数
    'Private Function EscapeJavaScriptString(jsString As String) As String
    '    If String.IsNullOrEmpty(jsString) Then Return ""

    '    Return jsString.Replace("\", "\\") _
    '                  .Replace("""", "\""") _
    '                  .Replace("'", "\'") _
    '                  .Replace(vbCrLf, "\n") _
    '                  .Replace(vbCr, "\n") _
    '                  .Replace(vbLf, "\n") _
    '                  .Replace(vbTab, "\t")
    'End Function

    Private Async Sub InjectScript(scriptContent As String)
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim escapedScript = JsonConvert.SerializeObject(scriptContent)
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedScript})")
        Else
            MessageBox.Show("CoreWebView2 未初始化，无法注入脚本。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub



    ' 检查代码是否包含过程声明
    Private Function ContainsProcedureDeclaration(code As String) As Boolean
        ' 使用简单的正则表达式检查是否包含 Sub 或 Function 声明
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function

    ' 查找模块中的第一个过程名
    Private Function FindFirstProcedureName(comp As VBComponent) As String
        Try
            Dim codeModule As CodeModule = comp.CodeModule
            Dim lineCount As Integer = codeModule.CountOfLines
            Dim line As Integer = 1

            While line <= lineCount
                Dim procName As String = codeModule.ProcOfLine(line, vbext_ProcKind.vbext_pk_Proc)
                If Not String.IsNullOrEmpty(procName) Then
                    Return procName
                End If
                line = codeModule.ProcStartLine(procName, vbext_ProcKind.vbext_pk_Proc) + codeModule.ProcCountLines(procName, vbext_ProcKind.vbext_pk_Proc)
            End While

            Return String.Empty
        Catch
            ' 如果出错，尝试使用正则表达式从代码中提取
            Dim code As String = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            Dim match As Match = Regex.Match(code, "^\s*(Sub|Function)\s+(\w+)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)

            If match.Success AndAlso match.Groups.Count > 2 Then
                Return match.Groups(2).Value
            End If

            Return String.Empty
        End Try
    End Function

    Protected Overrides Function GetWebView2DataFolderName() As String
        Return "DoubaoChatWebView2Data"
    End Function
End Class