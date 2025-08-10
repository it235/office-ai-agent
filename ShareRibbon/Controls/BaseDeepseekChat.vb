﻿' ShareRibbon\Controls\BaseChatControl.vb
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseDeepseekChat
    Inherits UserControl

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_PASTE As Integer = &H302
        If m.Msg = WM_PASTE Then
            ' 在此处理粘贴操作，比如：
            If Clipboard.ContainsText() Then
                Dim txt As String = Clipboard.GetText()

                'QuestionTextBox.Text &= txt ' 将粘贴内容直接写入当前光标位置
            End If
            ' 不把消息传递给基类，从而拦截后续处理  
            Return
        End If
        MyBase.WndProc(m)
    End Sub
    ' 添加到类中的新字段
    Private ReadOnly SessionFilePath As String = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    ConfigSettings.OfficeAiAppDataFolder,
    "deepseek_session.json"
)
    Protected Async Function InitializeWebView2() As Task
        Try
            ' 使用固定的用户数据目录而不是临时目录，以保持会话持久化
            Dim userDataFolder As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        ConfigSettings.OfficeAiAppDataFolder,
        "DeepseekChatWebView2Data")

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
                ' 重要：在导航前注册所有事件处理器
                'AddHandler ChatBrowser.CoreWebView2.NavigationStarting, AddressOf OnNavigationStarting
                AddHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived

                ' 启用持久化的Cookie管理
                ChatBrowser.CoreWebView2.CookieManager.DeleteAllCookies() ' 可选，仅在需要清理时使用

                ' 导航到Deepseek网站
                ChatBrowser.CoreWebView2.Navigate("https://chat.deepseek.com")

                SimpleLogger.LogInfo("WebView2初始化完成，开始导航到Deepseek")
            Else
                MessageBox.Show("WebView2初始化失败，CoreWebView2不可用。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"初始化失败: {ex.Message}{Environment.NewLine}类型: {ex.GetType().Name}{Environment.NewLine}堆栈:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    ' 在OnNavigationStarting方法中添加环境检查
    Private Async Sub OnNavigationStarting(sender As Object, e As CoreWebView2NavigationStartingEventArgs)
        Try
            ' 确保WebView2环境已正确初始化
            If ChatBrowser Is Nothing OrElse ChatBrowser.CoreWebView2 Is Nothing Then
                SimpleLogger.LogInfo("WebView2未初始化，无法恢复会话")
                Return
            End If

            ' 只在导航到Deepseek时应用会话
            If e.Uri.StartsWith("https://chat.deepseek.com") Then
                Await RestoreSessionAsync()
            End If
        Catch ex As Exception
            SimpleLogger.LogInfo($"恢复会话时出错: {ex.Message}")
        End Try
    End Sub

    ' 在页面加载完成后，注入脚本 - 修复线程问题
    Private Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If e.IsSuccess Then
            Try
                SimpleLogger.LogInfo("导航完成，开始注入脚本")

                ' 确保在UI线程上执行所有WebView2操作
                If ChatBrowser.InvokeRequired Then
                    ChatBrowser.Invoke(New Action(Async Sub()
                                                      Try
                                                          ' 延迟一些时间，确保页面完全加载
                                                          Await Task.Delay(1000)

                                                          ' 配置Marked和代码高亮
                                                          Await ConfigureMarkedSafe()

                                                          ' 注入基础辅助脚本
                                                          Await InitializeWebView2ScriptAsyncSafe()

                                                          ' 初始化设置和执行按钮
                                                          Await InitializeSettingsSafe()

                                                          ' 注入登录监听器
                                                          'Await InjectLoginObserverSafe()

                                                          SimpleLogger.LogInfo("所有脚本注入完成")
                                                      Catch ex As Exception
                                                          SimpleLogger.LogInfo($"UI线程脚本注入出错: {ex.Message}")
                                                          SimpleLogger.LogInfo(ex.StackTrace)
                                                      End Try
                                                  End Sub))
                Else
                    ' 已经在UI线程，直接执行
                    Task.Run(Async Function()
                                 Try
                                     ' 延迟一些时间，确保页面完全加载
                                     Await Task.Delay(1000)

                                     ' 在UI线程上执行脚本注入
                                     ChatBrowser.Invoke(New Action(Async Sub()
                                                                       Try
                                                                           ' 配置Marked和代码高亮
                                                                           Await ConfigureMarkedSafe()

                                                                           ' 注入基础辅助脚本
                                                                           Await InitializeWebView2ScriptAsyncSafe()

                                                                           ' 初始化设置和执行按钮
                                                                           Await InitializeSettingsSafe()

                                                                           ' 注入登录监听器
                                                                           'Await InjectLoginObserverSafe()

                                                                           SimpleLogger.LogInfo("所有脚本注入完成")
                                                                       Catch ex As Exception
                                                                           SimpleLogger.LogInfo($"脚本注入出错: {ex.Message}")
                                                                           SimpleLogger.LogInfo(ex.StackTrace)
                                                                       End Try
                                                                   End Sub))
                                 Catch ex As Exception
                                     SimpleLogger.LogInfo($"任务执行出错: {ex.Message}")
                                 End Try
                             End Function)
                End If
            Catch ex As Exception
                SimpleLogger.LogInfo($"导航完成事件处理中出错: {ex.Message}")
                SimpleLogger.LogInfo(ex.StackTrace)
            End Try
        Else
            SimpleLogger.LogInfo($"导航失败: {e.WebErrorStatus}")
        End If
    End Sub



    ' 线程安全的ConfigureMarked方法
    Private Async Function ConfigureMarkedSafe() As Task
        Try
            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                Dim script = "
            try {
                if (typeof marked !== 'undefined' && typeof hljs !== 'undefined') {
                    marked.setOptions({
                        highlight: function (code, lang) {
                            if (hljs.getLanguage(lang)) {
                                return hljs.highlight(lang, code).value;
                            } else {
                                return hljs.highlightAuto(code).value;
                            }
                        }
                    });
                    console.log('[VSTO] Marked配置完成');
                } else {
                    console.log('[VSTO] marked或hljs未加载');
                }
            } catch (e) {
                console.log('[VSTO] 配置marked时出错:', e);
            }
        "
                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
                SimpleLogger.LogInfo("ConfigureMarked执行完成")
            Else
                SimpleLogger.LogInfo("ConfigureMarked: CoreWebView2为空")
            End If
        Catch ex As Exception
            SimpleLogger.LogInfo($"ConfigureMarked出错: {ex.Message}")
        End Try
    End Function

    ' 线程安全的InitializeWebView2ScriptAsync方法
    Private Async Function InitializeWebView2ScriptAsyncSafe() As Task
        Try
            Dim script As String = "
    // 初始化VSTO接口
    window.vsto = {
        executeCode: function(code, language, preview) {
            console.log('[VSTO] executeCode被调用:', {code: code.substring(0, 50) + '...', language: language, preview: preview});
            window.chrome.webview.postMessage({
                type: 'executeCode',
                code: code,
                language: language,
                executecodePreview: preview
            });
            return true;
        },
        checkedChange: function(thisProperty, checked) {
            return window.chrome.webview.postMessage({
                type: 'checkedChange',
                isChecked: checked,
                property: thisProperty
            });
        },
        sendMessage: function(payload) {
            let messageToSend;
            if (typeof payload === 'string') {
                messageToSend = { type: 'sendMessage', value: payload };
            } else {
                messageToSend = payload;
            }
            window.chrome.webview.postMessage(messageToSend);
            return true;
        },
        saveSettings: function(settingsObject) {
            return window.chrome.webview.postMessage({
                type: 'saveSettings',
                topicRandomness: settingsObject.topicRandomness,
                contextLimit: settingsObject.contextLimit,
                selectedCell: settingsObject.selectedCell,
                executeCodePreview: settingsObject.executeCodePreview,
            });
        }
    };
    
    console.log('[VSTO] 基础API已初始化');
    
    // 验证通信接口
    if (window.chrome && window.chrome.webview) {
        console.log('[VSTO] ✓ chrome.webview接口可用');
    } else {
        console.log('[VSTO] ✗ chrome.webview接口不可用');
    }
    "

            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
                SimpleLogger.LogInfo("InitializeWebView2ScriptAsync执行完成")
            Else
                SimpleLogger.LogInfo("InitializeWebView2ScriptAsync: CoreWebView2为空")
            End If
        Catch ex As Exception
            SimpleLogger.LogInfo($"InitializeWebView2ScriptAsync出错: {ex.Message}")
        End Try
    End Function

    ' 线程安全的InitializeSettings方法
    Private Async Function InitializeSettingsSafe() As Task
        Try
            Await InjectExecuteButtonsSafe()
            SimpleLogger.LogInfo("InitializeSettings执行完成")
        Catch ex As Exception
            SimpleLogger.LogInfo($"InitializeSettings出错: {ex.Message}")
        End Try
    End Function

    ' 线程安全的InjectLoginObserver方法
    Private Async Function InjectLoginObserverSafe() As Task
        Try
            Dim script = "
            console.log('[VSTO] 开始注入登录观察器');
            
            // 监听登录状态变化
            function observeLoginStatus() {
                let isLoggedIn = false;
                
                // 检测登录状态
                function checkLoginStatus() {
                    // 检查方式1: 检查特定DOM元素存在
                    const hasUserAvatar = !!document.querySelector('.user-avatar') || 
                                         !!document.querySelector('.avatar-img') ||
                                         !!document.querySelector('[data-testid=\""user-dropdown\""]');
                    
                    // 检查方式2: 检查localStorage中的token
                    const hasToken = localStorage.getItem('ds_auth_token') || 
                                    localStorage.getItem('auth_token');
                    
                    // 检查方式3: 检查cookie
                    const hasSessionCookie = document.cookie.includes('ds_session_id');
                    
                    // 整合所有检查结果
                    const newLoginState = hasUserAvatar || !!hasToken || hasSessionCookie;
                    
                    // 如果状态从未登录变为已登录，通知应用
                    if (!isLoggedIn && newLoginState) {
                        console.log('[VSTO] 用户登录状态变化: 已登录');
                        if (window.chrome && window.chrome.webview) {
                            window.chrome.webview.postMessage({
                                type: 'loginStatusChanged',
                                status: 'loggedIn'
                            });
                        }
                    }
                    
                    // 更新状态
                    isLoggedIn = newLoginState;
                }
                
                // 立即检查一次
                checkLoginStatus();
                
                // 监听点击事件 - 用于捕获登录按钮点击后的状态变化
                document.addEventListener('click', function(e) {
                    // 延迟检查以等待登录完成
                    setTimeout(checkLoginStatus, 2000);
                });
                
                // 监听localStorage变化
                const originalSetItem = localStorage.setItem;
                localStorage.setItem = function() {
                    originalSetItem.apply(this, arguments);
                    // 检查是否有令牌被添加
                    if (arguments[0] && 
                        (arguments[0].includes('token') || arguments[0].includes('auth'))) {
                        setTimeout(checkLoginStatus, 500);
                    }
                };
                
                // 定期检查
                setInterval(checkLoginStatus, 5000);
                
                return true;
            }
            
            observeLoginStatus();
            console.log('[VSTO] 登录观察器已设置');
        "

            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
                SimpleLogger.LogInfo("InjectLoginObserver执行完成")
            Else
                SimpleLogger.LogInfo("InjectLoginObserver: CoreWebView2为空")
            End If
        Catch ex As Exception
            SimpleLogger.LogInfo($"注入登录观察器时出错: {ex.Message}")
        End Try
    End Function

    ' 线程安全的执行按钮注入方法 - 分离按钮创建和代码获取逻辑
    Private Async Function InjectExecuteButtonsSafe() As Task
        Try
            If ChatBrowser.CoreWebView2 Is Nothing Then
                SimpleLogger.LogInfo("InjectExecuteButtons: CoreWebView2未初始化")
                Return
            End If

            Dim script As String = "
    (function() {
        console.log('[Execute Buttons] =================== 开始初始化 ===================');
        
        // 避免重复初始化
        if (window.__executeButtonsInitialized) {
            console.log('[Execute Buttons] 已初始化，尝试刷新');
            if (window.refreshExecuteButtons) {
                window.refreshExecuteButtons();
            }
            return;
        }
        window.__executeButtonsInitialized = true;
        
        // 使用多种策略查找复制按钮
        function findCopyButton(container) {
            if (!container) return null;
            
            console.log('[Execute Buttons] 在容器中查找复制按钮');
            
            // 策略1: 直接通过文本内容查找
            const allButtons = container.querySelectorAll('[role=""button""]');
            console.log('[Execute Buttons] 找到按钮数量:', allButtons.length);
            
            for (let i = 0; i < allButtons.length; i++) {
                const btn = allButtons[i];
                const text = btn.textContent || '';
                console.log(`[Execute Buttons] 按钮 ${i+1} 文本: '${text.trim()}'`);
                
                if (text.includes('复制') || text.includes('Copy')) {
                    console.log('[Execute Buttons] ✓ 找到复制按钮!');
                    return btn;
                }
            }
            
            console.log('[Execute Buttons] ✗ 未找到复制按钮');
            return null;
        }
        
        // 动态获取代码内容的函数 - 在点击时调用
        function getCurrentCodeContent(codeBlock) {
            try {
                const preElement = codeBlock.querySelector('pre');
                if (!preElement) {
                    console.log('[Execute Buttons] 点击时未找到pre元素');
                    return { code: '', language: 'unknown' };
                }
                
                // 获取当前完整的代码内容（SSE流式更新后的完整内容）
                const codeContent = preElement.textContent || '';
                console.log(`[Execute Buttons] 动态获取代码长度: ${codeContent.length} 字符`);
                
                // 动态获取语言信息
                let language = 'unknown';
                const allSpans = codeBlock.querySelectorAll('span');
                
                for (const span of allSpans) {
                    const text = span.textContent && span.textContent.trim().toLowerCase();
                    if (text && /^(vba|javascript|js|excel|python|sql|typescript|html|css|c#|java|php|csharp)$/i.test(text)) {
                        language = text;
                        console.log(`[Execute Buttons] 动态获取语言: ${language}`);
                        break;
                    }
                }
                
                return { code: codeContent, language: language };
            } catch (error) {
                console.error('[Execute Buttons] 动态获取代码内容时出错:', error);
                return { code: '', language: 'unknown' };
            }
        }
        
        // 处理单个代码块 - 只创建按钮，不获取代码
        function processCodeBlock(codeBlock, index) {
            try {
                console.log(`[Execute Buttons] ========== 处理代码块 ${index+1} ==========`);
                
                if (!codeBlock.classList.contains('md-code-block')) {
                    console.log('[Execute Buttons] 不是代码块，跳过');
                    return false;
                }
                
                // 查找复制按钮
                const copyButton = findCopyButton(codeBlock);
                if (!copyButton) {
                    console.log('[Execute Buttons] 未找到复制按钮，跳过此代码块');
                    return false;
                }
                
                // 获取按钮容器
                const buttonContainer = copyButton.parentElement;
                if (!buttonContainer) {
                    console.log('[Execute Buttons] 找不到按钮容器');
                    return false;
                }
                
                // 检查是否已添加执行按钮
                if (buttonContainer.querySelector('.execute-code-button')) {
                    console.log('[Execute Buttons] 执行按钮已存在，跳过');
                    return false;
                }
                
                // 验证是否有pre元素（基本检查）
                const preElement = codeBlock.querySelector('pre');
                if (!preElement) {
                    console.log('[Execute Buttons] 未找到pre元素，跳过');
                    return false;
                }
                
                // 创建执行按钮
                console.log('[Execute Buttons] 开始创建执行按钮...');
                
                const executeButton = copyButton.cloneNode(true);
                executeButton.classList.add('execute-code-button');
                
                // 修改按钮内容
                executeButton.innerHTML = '';
                
                // 创建图标容器
                const iconDiv = document.createElement('div');
                iconDiv.className = 'ds-button__icon';
                
                const icon = document.createElement('div');
                icon.className = 'ds-icon';
                icon.style.fontSize = '16px';
                icon.style.width = '16px';
                icon.innerHTML = '▶'; // 播放图标
                
                iconDiv.appendChild(icon);
                executeButton.appendChild(iconDiv);
                
                // 创建文本
                const textSpan = document.createElement('span');
                textSpan.className = 'code-info-button-text';
                textSpan.textContent = '执行';
                executeButton.appendChild(textSpan);
                
                console.log('[Execute Buttons] ✓ 执行按钮创建完成');
                
                // 关键改进：添加点击事件 - 动态获取代码内容
                executeButton.addEventListener('click', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    console.log('[Execute Buttons] =================== 执行按钮被点击 ===================');
                    
                    // 在点击时动态获取当前完整的代码内容
                    const currentContent = getCurrentCodeContent(codeBlock);
                    
                    if (!currentContent.code.trim()) {
                        console.log('[Execute Buttons] ✗ 获取到的代码内容为空');
                        return;
                    }
                    
                    console.log('[Execute Buttons] 当前代码长度:', currentContent.code.length);
                    console.log('[Execute Buttons] 当前语言:', currentContent.language);
                    console.log('[Execute Buttons] 代码预览:', currentContent.code.substring(0, 200) + '...');
                    
                    try {
                        // 发送到VB应用
                        if (window.vsto && typeof window.vsto.executeCode === 'function') {
                            window.vsto.executeCode(currentContent.code, currentContent.language, true);
                            console.log('[Execute Buttons] ✓ 通过vsto接口发送执行请求');
                        } else if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {
                            window.chrome.webview.postMessage({
                                type: 'executeCode',
                                code: currentContent.code,
                                language: currentContent.language,
                                executecodePreview: true
                            });
                            console.log('[Execute Buttons] ✓ 通过chrome.webview发送执行请求');
                        } else {
                            console.error('[Execute Buttons] ✗ 通信接口不可用');
                            console.log('[Execute Buttons] vsto可用:', !!window.vsto);
                            console.log('[Execute Buttons] chrome.webview可用:', !!(window.chrome && window.chrome.webview));
                        }
                    } catch (error) {
                        console.error('[Execute Buttons] ✗ 发送执行请求失败:', error);
                    }
                });
                
                // 插入执行按钮到DOM
                try {
                    buttonContainer.insertBefore(executeButton, copyButton);
                    console.log('[Execute Buttons] ✓ 成功插入执行按钮到DOM');
                    
                    // 验证按钮是否真的添加了
                    const verification = buttonContainer.querySelector('.execute-code-button');
                    if (verification) {
                        console.log('[Execute Buttons] ✓ 验证：执行按钮存在于DOM中');
                        return true;
                    } else {
                        console.log('[Execute Buttons] ✗ 验证失败：执行按钮未在DOM中找到');
                        return false;
                    }
                } catch (e) {
                    console.error('[Execute Buttons] ✗ 插入DOM失败:', e);
                    return false;
                }
                
            } catch (error) {
                console.error(`[Execute Buttons] ✗ 处理代码块 ${index+1} 时出错:`, error);
                return false;
            }
        }
        
        // 扫描并处理所有代码块
        function addExecuteButtons() {
            console.log('[Execute Buttons] ========================================');
            console.log('[Execute Buttons] 开始扫描页面上的代码块');
            
            const codeBlocks = document.querySelectorAll('.md-code-block');
            console.log(`[Execute Buttons] 在页面上找到 ${codeBlocks.length} 个代码块`);
            
            if (codeBlocks.length === 0) {
                console.log('[Execute Buttons] 页面上没有找到任何代码块');
                return;
            }
            
            let processedCount = 0;
            codeBlocks.forEach((block, index) => {
                if (processCodeBlock(block, index)) {
                    processedCount++;
                }
            });
            
            console.log(`[Execute Buttons] 扫描完成！成功处理 ${processedCount}/${codeBlocks.length} 个代码块`);
            console.log('[Execute Buttons] ========================================');
        }
        
        // 设置全局刷新函数
        window.refreshExecuteButtons = addExecuteButtons;
        
        // DOM观察器 - 监听新代码块的添加
        function setupObserver() {
            const observer = new MutationObserver((mutations) => {
                let shouldProcess = false;
                
                for (const mutation of mutations) {
                    if (mutation.addedNodes.length > 0) {
                        for (const node of mutation.addedNodes) {
                            if (node.nodeType === 1) {
                                if ((node.classList && node.classList.contains('md-code-block')) ||
                                    (node.querySelector && node.querySelector('.md-code-block'))) {
                                    shouldProcess = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (shouldProcess) break;
                }
                
                if (shouldProcess) {
                    console.log('[Execute Buttons] DOM变化检测到新代码块');
                    setTimeout(addExecuteButtons, 200);
                }
            });
            
            observer.observe(document.body, {
                childList: true,
                subtree: true
            });
            
            console.log('[Execute Buttons] ✓ DOM观察器已激活');
        }
        
        // 初始化
        console.log('[Execute Buttons] 开始初始化流程');
        
        // 立即尝试
        setTimeout(addExecuteButtons, 100);
        
        // 设置观察器
        setupObserver();
        
        // 多次检查 - 确保捕获所有代码块
        [500, 1000, 2000, 3000, 5000].forEach((delay, index) => {
            setTimeout(() => {
                console.log(`[Execute Buttons] 第${index+2}次检查 - ${delay}ms`);
                addExecuteButtons();
            }, delay);
        });
        
        // 手动刷新快捷键
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.shiftKey && e.key === 'E') {
                console.log('[Execute Buttons] 手动刷新 (Ctrl+Shift+E)');
                e.preventDefault();
                addExecuteButtons();
            }
        });
        
        console.log('[Execute Buttons] =================== 初始化完成 ===================');
    })();
    "

            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            SimpleLogger.LogInfo("执行按钮注入脚本已执行")
        Catch ex As Exception
            SimpleLogger.LogInfo($"注入执行按钮脚本时出错: {ex.Message}")
        End Try
    End Function

    ' 保存会话信息到本地文件
    Private Async Function SaveSessionAsync() As Task
        Try
            ' 确保目录存在
            Dim directoryx As String = Path.GetDirectoryName(SessionFilePath)
            If Not Directory.Exists(directoryx) Then
                Directory.CreateDirectory(directoryx)
            End If

            ' 提取授权令牌
            Dim authToken As String = Await ExtractAuthTokenAsync()

            ' 获取Cookies
            Dim cookies As String = Await GetCookiesAsync("https://chat.deepseek.com")

            ' 如果至少有一项不为空，保存会话信息
            If Not String.IsNullOrEmpty(authToken) OrElse Not String.IsNullOrEmpty(cookies) Then
                Dim sessionInfo As New JObject()
                sessionInfo("authToken") = authToken
                sessionInfo("cookies") = cookies
                sessionInfo("timestamp") = DateTime.Now.ToString("o")

                ' 保存到文件
                File.WriteAllText(SessionFilePath, sessionInfo.ToString())
                SimpleLogger.LogInfo("已保存Deepseek会话信息")
            End If
        Catch ex As Exception
            SimpleLogger.LogInfo($"保存会话信息时出错: {ex.Message}")
        End Try
    End Function

    ' 修复RestoreCookiesAsync方法
    Private Async Function RestoreCookiesAsync(cookieString As String) As Task
        Try
            ' 确保在UI线程上执行
            If ChatBrowser.InvokeRequired Then
                Await Task.Run(Sub()
                                   ChatBrowser.Invoke(New Action(Sub()
                                                                     Try
                                                                         Dim cookieManager = ChatBrowser.CoreWebView2.CookieManager
                                                                         Dim cookiePairs = cookieString.Split(New String() {";"}, StringSplitOptions.RemoveEmptyEntries)

                                                                         For Each pair In cookiePairs
                                                                             Dim parts = pair.Trim().Split(New String() {"="}, 2, StringSplitOptions.None)
                                                                             If parts.Length = 2 Then
                                                                                 Try
                                                                                     ' 正确创建和添加Cookie
                                                                                     Dim cookie = cookieManager.CreateCookie(
                                                                                    parts(0).Trim(),
                                                                                    parts(1).Trim(),
                                                                                    ".deepseek.com",
                                                                                    "/"
                                                                                )
                                                                                     cookie.IsSecure = True
                                                                                     cookieManager.AddOrUpdateCookie(cookie)
                                                                                 Catch cookieEx As Exception
                                                                                     SimpleLogger.LogInfo($"添加Cookie '{parts(0)}'时出错: {cookieEx.Message}")
                                                                                 End Try
                                                                             End If
                                                                         Next
                                                                     Catch ex As Exception
                                                                         SimpleLogger.LogInfo($"在UI线程恢复Cookies时出错: {ex.Message}")
                                                                     End Try
                                                                 End Sub))
                               End Sub)
            Else
                Dim cookieManager = ChatBrowser.CoreWebView2.CookieManager
                Dim cookiePairs = cookieString.Split(New String() {";"}, StringSplitOptions.RemoveEmptyEntries)

                For Each pair In cookiePairs
                    Dim parts = pair.Trim().Split(New String() {"="}, 2, StringSplitOptions.None)
                    If parts.Length = 2 Then
                        Try
                            Dim cookie = cookieManager.CreateCookie(
                            parts(0).Trim(),
                            parts(1).Trim(),
                            ".deepseek.com",
                            "/"
                        )
                            cookie.IsSecure = True
                            cookieManager.AddOrUpdateCookie(cookie)
                        Catch cookieEx As Exception
                            SimpleLogger.LogInfo($"添加Cookie '{parts(0)}'时出错: {cookieEx.Message}")
                        End Try
                    End If
                Next
            End If

            ' 短暂延迟确保Cookie操作完成
            Await Task.Delay(100)
        Catch ex As Exception
            SimpleLogger.LogInfo($"恢复Cookies时出错: {ex.Message}")
        End Try
    End Function


    ' 改进ExtractAuthTokenAsync以获取更准确的令牌
    Private Async Function ExtractAuthTokenAsync() As Task(Of String)
        Try
            Dim script = "
            function getDeepseekAuthToken() {
                try {
                    // 直接查找所有请求头，拦截一个真实请求获取令牌
                    let capturedToken = '';
                    
                    // 检查现有存储
                    for (let i = 0; i < localStorage.length; i++) {
                        const key = localStorage.key(i);
                        if (key && (key.includes('token') || key.includes('auth'))) {
                            const value = localStorage.getItem(key);
                            if (value && value.length > 20) {
                                console.log('找到可能的令牌:', key);
                                return value;
                            }
                        }
                    }
                    
                    // 如果没找到，尝试发送一个请求并捕获令牌
                    const origFetch = window.fetch;
                    window.fetch = function(input, init) {
                        if (init && init.headers) {
                            const headers = new Headers(init.headers);
                            const authHeader = headers.get('Authorization');
                            if (authHeader) {
                                capturedToken = authHeader;
                                console.log('捕获到授权头:', authHeader);
                            }
                        }
                        return origFetch.apply(this, arguments);
                    };
                    
                    // 触发请求 (会在后台执行)
                    setTimeout(() => {
                        fetch('/api/sessions', { 
                            method: 'GET',
                            credentials: 'include'
                        }).catch(() => {});
                    }, 0);
                    
                    // 尝试直接从cookie中提取
                    const cookies = document.cookie.split(';');
                    for (const cookie of cookies) {
                        if (cookie.includes('session')) {
                            console.log('找到会话cookie:', cookie);
                        }
                    }
                    
                    return capturedToken || '';
                } catch (e) {
                    console.error('获取令牌时出错:', e);
                    return '';
                }
            }
            getDeepseekAuthToken();
        "

            Dim result As String = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)

            ' 处理令牌结果
            If Not String.IsNullOrEmpty(result) AndAlso result <> "null" Then
                ' 清理JSON转义和引号
                result = result.Trim("""")

                ' 处理JSON字符串
                If result.StartsWith("{") OrElse result.Contains("\\") Then
                    Try
                        ' 如果是JSON字符串，尝试规范化
                        result = result.Replace("\\\""", """").Replace("\\\\", "\")
                        If result.StartsWith("""") AndAlso result.EndsWith("""") Then
                            result = result.Substring(1, result.Length - 2)
                        End If
                    Catch ex As Exception
                        SimpleLogger.LogInfo("清理令牌格式时出错: " & ex.Message)
                    End Try
                End If

                ' 确保令牌具有正确前缀
                If Not String.IsNullOrEmpty(result) AndAlso Not result.StartsWith("Bearer ") Then
                    result = "Bearer " & result.Trim()
                End If

                Return result
            End If

            Return String.Empty
        Catch ex As Exception
            SimpleLogger.LogInfo($"提取授权令牌时出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 修改Cookie处理，去除重复项
    Private Async Function GetCookiesAsync(url As String) As Task(Of String)
        Try
            If ChatBrowser.InvokeRequired Then
                ' 在UI线程执行
                Dim resultCookies As String = ""
                Dim taskCompletionSource = New TaskCompletionSource(Of String)()

                ChatBrowser.Invoke(Sub()
                                       Try
                                           ' 在UI线程获取Cookies
                                           Dim cookieManager = ChatBrowser.CoreWebView2.CookieManager

                                           ' 使用同步版本避免嵌套异步
                                           Dim task = cookieManager.GetCookiesAsync(url)
                                           task.Wait() ' 同步等待结果
                                           Dim cookies = task.Result

                                           If cookies IsNot Nothing AndAlso cookies.Count > 0 Then
                                               ' 使用字典去重
                                               Dim cookieDict As New Dictionary(Of String, String)

                                               For Each cookie In cookies
                                                   cookieDict(cookie.Name) = cookie.Value
                                               Next

                                               ' 构建Cookie字符串
                                               Dim cookiePairs = New List(Of String)
                                               For Each pair In cookieDict
                                                   cookiePairs.Add($"{pair.Key}={pair.Value}")
                                               Next

                                               resultCookies = String.Join("; ", cookiePairs)
                                           End If

                                           taskCompletionSource.SetResult(resultCookies)
                                       Catch ex As Exception
                                           SimpleLogger.LogInfo($"获取Cookies时出错: {ex.Message}")
                                           taskCompletionSource.SetResult("")
                                       End Try
                                   End Sub)

                Return Await taskCompletionSource.Task
            Else
                ' 在当前线程执行
                Dim cookieManager = ChatBrowser.CoreWebView2.CookieManager
                Dim cookies = Await cookieManager.GetCookiesAsync(url)

                If cookies IsNot Nothing AndAlso cookies.Count > 0 Then
                    ' 使用字典去重
                    Dim cookieDict As New Dictionary(Of String, String)

                    For Each cookie In cookies
                        cookieDict(cookie.Name) = cookie.Value
                    Next

                    ' 构建Cookie字符串
                    Dim cookiePairs = New List(Of String)
                    For Each pair In cookieDict
                        cookiePairs.Add($"{pair.Key}={pair.Value}")
                    Next

                    Return String.Join("; ", cookiePairs)
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            SimpleLogger.LogInfo($"获取Cookies时出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 改进令牌注入方法
    Private Async Function InjectAuthTokenAsync(authToken As String) As Task
        Try
            ' 清理令牌格式
            Dim cleanToken As String = authToken
            If cleanToken.StartsWith("Bearer ") Then
                cleanToken = cleanToken.Substring(7)
            End If

            Dim script = $"
            function injectAuthToken() {{
                try {{
                    // 保存到常用的令牌存储位置
                    localStorage.setItem('ds_auth_token', '{EscapeJavaScriptString(cleanToken)}');
                    localStorage.setItem('auth_token', '{EscapeJavaScriptString(cleanToken)}');
                    
                    // 修补XHR请求
                    const originalXhrOpen = XMLHttpRequest.prototype.open;
                    XMLHttpRequest.prototype.open = function() {{
                        originalXhrOpen.apply(this, arguments);
                        this.setRequestHeader('Authorization', 'Bearer {EscapeJavaScriptString(cleanToken)}');
                    }};
                    
                    // 修补Fetch请求
                    const originalFetch = window.fetch;
                    window.fetch = function(resource, init) {{
                        if (!init) init = {{}};
                        if (!init.headers) init.headers = {{}};
                        
                        // 添加授权头
                        init.headers['Authorization'] = 'Bearer {EscapeJavaScriptString(cleanToken)}';
                        
                        return originalFetch.call(this, resource, init);
                    }};
                    
                    console.log('已成功注入授权令牌');
                    return true;
                }} catch (e) {{
                    console.error('注入令牌失败:', e);
                    return false;
                }}
            }}
            injectAuthToken();
        "

            ' 确保在UI线程执行
            If ChatBrowser.InvokeRequired Then
                Await Task.Run(Sub()
                                   ChatBrowser.Invoke(Sub()
                                                          Try
                                                              Dim task = ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
                                                              task.Wait() ' 同步等待完成
                                                          Catch ex As Exception
                                                              SimpleLogger.LogInfo($"执行脚本时出错: {ex.Message}")
                                                          End Try
                                                      End Sub)
                               End Sub)
            Else
                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            End If
        Catch ex As Exception
            SimpleLogger.LogInfo($"注入授权令牌时出错: {ex.Message}")
        End Try
    End Function

    ' 更高级的会话恢复方法 - 增加错误处理和调试信息
    Private Async Function RestoreSessionAsync() As Task
        Try
            If File.Exists(SessionFilePath) Then
                Dim sessionJson As String = File.ReadAllText(SessionFilePath)
                Dim sessionInfo As JObject = JObject.Parse(sessionJson)

                ' 检查会话是否过期（超过7天）
                Dim timestamp As DateTime
                If DateTime.TryParse(sessionInfo("timestamp")?.ToString(), timestamp) Then
                    If (DateTime.Now - timestamp).TotalDays > 7 Then
                        SimpleLogger.LogInfo("会话已过期，需要重新登录")
                        Return
                    End If
                Else
                    SimpleLogger.LogInfo("无效的会话时间戳")
                    Return
                End If

                ' 恢复Cookie
                If sessionInfo.ContainsKey("cookies") AndAlso
               Not String.IsNullOrEmpty(sessionInfo("cookies")?.ToString()) Then
                    Await RestoreCookiesAsync(sessionInfo("cookies").ToString())
                    SimpleLogger.LogInfo("已恢复Cookies")
                End If

                ' 注入授权令牌
                If sessionInfo.ContainsKey("authToken") AndAlso
               Not String.IsNullOrEmpty(sessionInfo("authToken")?.ToString()) Then
                    Await InjectAuthTokenAsync(sessionInfo("authToken").ToString())
                    SimpleLogger.LogInfo("已注入授权令牌")
                End If

                SimpleLogger.LogInfo("已恢复Deepseek会话信息")
            Else
                SimpleLogger.LogInfo("未找到会话文件，需要重新登录")
            End If
        Catch ex As Exception
            SimpleLogger.LogInfo($"恢复会话信息时出错: {ex.Message}")
        End Try
    End Function

    Private Async Sub InjectScript(scriptContent As String)
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim escapedScript = JsonConvert.SerializeObject(scriptContent)
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedScript})")
        Else
            MessageBox.Show("CoreWebView2 未初始化，无法注入脚本。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Async Function ConfigureMarked() As Task
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim script = "
            marked.setOptions({
                highlight: function (code, lang) {
                    if (hljs.getLanguage(lang)) {
                        return hljs.highlight(lang, code).value;
                    } else {
                        return hljs.highlightAuto(code).value;
                    }
                }
            });
        "
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
        Else
            MessageBox.Show("CoreWebView2 未初始化，无法配置 Marked。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Function

    ' 存储聊天HTML的文件路径
    Protected ReadOnly ChatHtmlFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
    )


    ' 扩展WebView2_WebMessageReceived方法，处理登录状态变化
    Protected Sub WebView2_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim jsonDoc As JObject = JObject.Parse(e.WebMessageAsJson)
            Dim messageType As String = jsonDoc("type").ToString()

            Select Case messageType
                Case "executeCode"
                    HandleExecuteCode(jsonDoc)
                Case Else
                    SimpleLogger.LogInfo($"未知消息类型: {messageType}")
            End Select
        Catch ex As Exception
            SimpleLogger.LogInfo($"处理消息出错: {ex.Message}")
        End Try
    End Sub

    Protected Overridable Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim preview As Boolean = Boolean.Parse(jsonDoc("executecodePreview"))
        Dim language As String = jsonDoc("language").ToString()
        ExecuteCode(code, language, preview)
    End Sub


    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String

    Protected MustOverride Function GetApplication() As ApplicationInfo
    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCodePreview(vbaCode As String, preview As Boolean)
    Protected MustOverride Function RunCode(vbaCode As String)

    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' 执行代码的方法
    Private Sub ExecuteCode(code As String, language As String, preview As Boolean)
        ' 根据语言类型执行不同的操作
        Select Case language.ToLower()
            Case "vba", "vb", "vbscript", "language-vba", "language-vbscript", "language-vba hljs language-vbscript", "vba hljs language-vbscript"
                ' 执行 VBA 代码
                ExecuteVBACode(code, preview)
            Case "js", "javascript", "javascript hljs", "jscript", "language-js", "language-javascript"
                ' 执行 JavaScript 代码
                ExecuteJavaScript(code, preview)
            Case "excel", "formula", "function", "language-excel"
                ' 执行 Excel 函数/公式
                ExecuteExcelFormula(code, preview)
                'Case "sql", "language-sql"
                '    ' 执行 SQL 查询
                '    ExecuteSqlQuery(code, preview)
                'Case "powerquery", "m", "language-powerquery", "language-m"
                '    ' 执行 PowerQuery/M 语言
                '    ExecutePowerQuery(code, preview)
                'Case "python", "py", "language-python"
                '    ' 执行 Python 代码
                '    ExecutePython(code, preview)
            Case Else
                GlobalStatusStrip.ShowWarning("不支持的语言类型: " & language)
        End Select
    End Sub

    ' 执行JavaScript代码 - 专注于操作Office/WPS对象模型，支持Office JS API风格代码
    Protected Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean
        Try
            ' 获取Office应用对象
            Dim appObject As Object = GetOfficeApplicationObject()
            If appObject Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取Office应用程序对象")
                Return False
            End If

            ' 检测是否是Office JS API风格的代码
            Dim isOfficeJsApiStyle As Boolean = jsCode.Contains("getActiveWorksheet") OrElse
                                            jsCode.Contains("getUsedRange") OrElse
                                            jsCode.Contains("getValues") OrElse
                                            jsCode.Contains("setValues")

            ' 创建脚本控制引擎
            Dim scriptEngine As Object = CreateObject("MSScriptControl.ScriptControl")
            scriptEngine.Language = "JScript"

            ' 判断是WPS还是Microsoft Office
            Dim isWPS As Boolean = False
            Try
                Dim appName As String = appObject.Name
                isWPS = appName.Contains("WPS")
            Catch ex As Exception
                isWPS = False
            End Try

            ' 将Office应用对象暴露给脚本环境
            scriptEngine.AddObject("app", appObject, True)

            ' 添加适配层代码
            Dim adapterCode As String = "
        // Office JS API 适配层
        var Office = {
            isWPS: " & isWPS.ToString().ToLower() & ",
            app: app,
            context: {
                workbook: {
                    // 适配 Office JS API 方法到 COM 对象
                    getActiveWorksheet: function() {
                        return {
                            sheet: app.ActiveSheet,
                            getUsedRange: function() {
                                var usedRange = this.sheet.UsedRange;
                                return {
                                    range: usedRange,
                                    getValues: function() {
                                        var values = [];
                                        var rows = this.range.Rows.Count;
                                        var cols = this.range.Columns.Count;
                                        
                                        for(var i = 1; i <= rows; i++) {
                                            var rowValues = [];
                                            for(var j = 1; j <= cols; j++) {
                                                var cellValue = this.range.Cells(i, j).Value;
                                                rowValues.push(cellValue);
                                            }
                                            values.push(rowValues);
                                        }
                                        return values;
                                    },
                                    setValues: function(values) {
                                        if(!values || values.length === 0) return;
                                        
                                        for(var i = 0; i < values.length; i++) {
                                            var row = values[i];
                                            for(var j = 0; j < row.length; j++) {
                                                try {
                                                    this.range.Cells(i+1, j+1).Value = row[j];
                                                } catch(e) {
                                                    // 忽略单元格设置错误
                                                }
                                            }
                                        }
                                    }
                                };
                            }
                        };
                    }
                }
            },
            // 日志函数
            log: function(message) { 
                return '输出: ' + message; 
            }
        };
        
        // Office JS API 主函数适配器
        function executeOfficeJsApi(codeFunc) {
            var workbook = Office.context.workbook;
            if(typeof codeFunc === 'function') {
                try {
                    return codeFunc(workbook);
                } catch(e) {
                    return 'Office JS API 执行错误: ' + e.message;
                }
            }
            return 'Invalid function';
        }
        "

            ' 预执行适配层代码
            scriptEngine.ExecuteStatement(adapterCode)

            ' 构建执行代码，根据代码类型选择不同的执行方式
            Dim wrappedCode As String

            If isOfficeJsApiStyle Then
                ' 如果是Office JS API风格，使用适配层执行
                wrappedCode = "
            try {
                // 将用户代码包装为函数
                var userFunc = function(workbook) {
                    " & jsCode & "
                };
                
                // 使用适配器执行
                executeOfficeJsApi(userFunc);
                return 'Office JS API 代码执行成功';
            } catch(e) {
                return 'Office JS API 执行错误: ' + e.message;
            }
            "
            Else
                ' 普通JavaScript代码
                wrappedCode = "
            try {
                // 用户代码开始
                " & jsCode & "
                // 用户代码结束
                return '代码执行成功';
            } catch(e) {
                return '执行错误: ' + e.message;
            }
            "
            End If

            ' 执行JavaScript代码并获取结果
            Dim result As String = scriptEngine.Eval(wrappedCode)
            GlobalStatusStrip.ShowInfo(result)

            Return True
        Catch ex As Exception
            MessageBox.Show("执行JavaScript代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function



    ' 添加清除特定 sheetName 的方法
    Public Async Sub ClearSelectedContentBySheetName(sheetName As String)
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
        $"clearSelectedContentBySheetName({JsonConvert.SerializeObject(sheetName)})"
    )
    End Sub


    ' 抽象方法 - 获取Office应用程序对象
    Protected MustOverride Function GetOfficeApplicationObject() As Object

    ' 执行Excel公式或函数 - 基类通用实现
    Protected Function ExecuteExcelFormula(formulaCode As String, preview As Boolean) As Boolean
        Try
            ' 获取应用程序信息
            Dim appInfo As ApplicationInfo = GetApplication()

            ' 去除可能的等号前缀
            If formulaCode.StartsWith("=") Then
                formulaCode = formulaCode.Substring(1)
            End If

            ' 根据应用类型处理
            If appInfo.Type = OfficeApplicationType.Excel Then
                ' 对于Excel，使用Evaluate方法
                Dim result As Boolean = EvaluateFormula(formulaCode, preview)
                GlobalStatusStrip.ShowInfo("公式执行结果: " & result.ToString())
                Return True
            Else
                ' 其他应用不支持直接执行Excel公式
                GlobalStatusStrip.ShowWarning("Excel公式执行仅支持Excel环境")
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show("执行Excel公式时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' 虚方法 - 评估Excel公式（只有Excel子类会实现）
    Protected Overridable Function EvaluateFormula(formula As String, preview As Boolean) As Boolean
        ' 默认实现返回Nothing
        Return True
    End Function

    ' 执行前端传来的 VBA 代码片段
    Protected Function ExecuteVBACode(vbaCode As String, preview As Boolean)

        If preview Then
            ' 返回是否需要执行，accept-True，reject-False
            If Not RunCodePreview(vbaCode, preview) Then
                Return True
            End If
            ' 如果预览模式，直接返回
        End If

        ' 获取 VBA 项目
        Dim vbProj As VBProject = GetVBProject()

        ' 添加空值检查
        If vbProj Is Nothing Then
            Return False
        End If

        Dim vbComp As VBComponent = Nothing
        Dim tempModuleName As String = "TempMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

        Try
            ' 创建临时模块
            vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            vbComp.Name = tempModuleName

            ' 检查代码是否已包含 Sub/Function 声明
            If ContainsProcedureDeclaration(vbaCode) Then
                ' 代码已包含过程声明，直接添加
                vbComp.CodeModule.AddFromString(vbaCode)

                ' 查找第一个过程名并执行
                Dim procName As String = FindFirstProcedureName(vbComp)
                If Not String.IsNullOrEmpty(procName) Then
                    RunCode(tempModuleName & "." & procName)
                Else
                    'MessageBox.Show("无法在代码中找到可执行的过程")
                    GlobalStatusStrip.ShowWarning("无法在代码中找到可执行的过程")
                End If
            Else
                ' 代码不包含过程声明，将其包装在 Auto_Run 过程中
                Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                           vbaCode & vbNewLine &
                                           "End Sub"
                vbComp.CodeModule.AddFromString(wrappedCode)

                ' 执行 Auto_Run 过程
                RunCode(tempModuleName & ".Auto_Run")

            End If

        Catch ex As Exception
            MessageBox.Show("执行 VBA 代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 无论成功还是失败，都删除临时模块
            Try
                If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                    vbProj.VBComponents.Remove(vbComp)
                End If
            Catch
                ' 忽略清理错误
            End Try
        End Try
    End Function


    ' 检查代码是否包含过程声明
    Public Function ContainsProcedureDeclaration(code As String) As Boolean
        ' 使用简单的正则表达式检查是否包含 Sub 或 Function 声明
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function


    ' 查找模块中的第一个过程名
    Public Function FindFirstProcedureName(comp As VBComponent) As String
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


    ' 执行js脚本的异步方法
    Private Async Function ExecuteJavaScriptAsyncJS(js As String) As Task
        If ChatBrowser.InvokeRequired Then
            ChatBrowser.Invoke(Sub() ChatBrowser.ExecuteScriptAsync(js))
        Else
            Await ChatBrowser.ExecuteScriptAsync(js)
        End If
    End Function

    Private Function DecodeBase64(base64 As String) As String
        Dim bytes As Byte() = System.Convert.FromBase64String(base64)
        Return System.Text.Encoding.UTF8.GetString(bytes)
    End Function

    Private Function EscapeJavaScriptString(input As String) As String
        Return input _
        .Replace("\", "\\") _
        .Replace("'", "\'") _
        .Replace(vbCr, "") _
        .Replace(vbLf, "\n") _
        .Replace("</script>", "<\/script>")  ' 避免脚本注入
    End Function


    Protected Shared Sub VBAxceptionHandle(ex As Runtime.InteropServices.COMException)
        ' 处理信任中心权限问题
        If ex.Message.Contains("程序访问不被信任") OrElse
       ex.Message.Contains("Programmatic access to Visual Basic Project is not trusted") Then
            VBATrustShowBox()
        Else
            MessageBox.Show("执行 VBA 代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Shared Sub VBATrustShowBox()
        MessageBox.Show(
                        "无法执行 VBA 代码，请按以下步骤设置：" & vbCrLf & vbCrLf &
                        "1. 点击 '文件' -> '选项' -> '信任中心'" & vbCrLf &
                        "2. 点击 '信任中心设置'" & vbCrLf &
                        "3. 选择 '宏设置'" & vbCrLf &
                        "4. 勾选 '信任对 VBA 项目对象模型的访问'",
                        "需要设置信任中心权限",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning)
    End Sub

End Class