' ShareRibbon\Controls\BaseChatControl.vb
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
    Inherits BaseChat


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
' 确保WebView2可以接收焦点
                ChatBrowser.TabStop = True
                ChatBrowser.TabIndex = 1
                ChatBrowser.Visible = True
                
                ' 配置WebView2设置以改善焦点行为
                ChatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                ChatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
                ' 启用开发者工具以便调试可能的焦点问题
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True
                
                ' 重要：在导航前注册所有事件处理器
                'AddHandler ChatBrowser.CoreWebView2.NavigationStarting, AddressOf OnNavigationStarting
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



    ' 在页面加载完成后，注入脚本 - 修复线程问题
    Private Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If e.IsSuccess Then
            Try
                Debug.WriteLine("导航完成，开始注入脚本")

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


                                                          Debug.WriteLine("所有脚本注入完成")
                                                      Catch ex As Exception
                                                          Debug.WriteLine($"UI线程脚本注入出错: {ex.Message}")
                                                          Debug.WriteLine(ex.StackTrace)
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

                                                                           Debug.WriteLine("所有脚本注入完成")
                                                                       Catch ex As Exception
                                                                           Debug.WriteLine($"脚本注入出错: {ex.Message}")
                                                                           Debug.WriteLine(ex.StackTrace)
                                                                       End Try
                                                                   End Sub))
                                 Catch ex As Exception
                                     Debug.WriteLine($"任务执行出错: {ex.Message}")
                                 End Try
                             End Function)
                End If
            Catch ex As Exception
                Debug.WriteLine($"导航完成事件处理中出错: {ex.Message}")
                Debug.WriteLine(ex.StackTrace)
            End Try
        Else
            Debug.WriteLine($"导航失败: {e.WebErrorStatus}")
        End If
    End Sub



    Protected Overrides Async Function ConfigureMarkedSafe() As Task
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
                Debug.WriteLine("ConfigureMarked执行完成")
            Else
                Debug.WriteLine("ConfigureMarked: CoreWebView2为空")
            End If
        Catch ex As Exception
            Debug.WriteLine($"ConfigureMarked出错: {ex.Message}")
        End Try
    End Function

    Protected Overrides Async Function InitializeWebView2ScriptAsyncSafe() As Task
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
                Debug.WriteLine("InitializeWebView2ScriptAsync执行完成")
            Else
                Debug.WriteLine("InitializeWebView2ScriptAsync: CoreWebView2为空")
            End If
        Catch ex As Exception
            Debug.WriteLine($"InitializeWebView2ScriptAsync出错: {ex.Message}")
        End Try
    End Function

    Protected Overrides Async Function InitializeSettingsSafe() As Task
        Try
            Await InjectExecuteButtonsSafe()
            Debug.WriteLine("InitializeSettings执行完成")
        Catch ex As Exception
            Debug.WriteLine($"InitializeSettings出错: {ex.Message}")
        End Try
    End Function

    Protected Overrides Async Function InjectLoginObserverSafe() As Task
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
                Debug.WriteLine("InjectLoginObserver执行完成")
            Else
                Debug.WriteLine("InjectLoginObserver: CoreWebView2为空")
            End If
        Catch ex As Exception
            Debug.WriteLine($"注入登录观察器时出错: {ex.Message}")
        End Try
    End Function

    Protected Overrides Async Function InjectExecuteButtonsSafe() As Task
        Try
            If ChatBrowser.CoreWebView2 Is Nothing Then
                Debug.WriteLine("InjectExecuteButtons: CoreWebView2未初始化")
                Return
            End If

            Dim script As String = "
    (function() {
        console.log('[Execute Buttons] 注入开始（使用自定义可点击元素以避免被页面禁用）');

        if (window.__executeButtonsInitialized) {
            console.log('[Execute Buttons] 已初始化，刷新');
            if (window.refreshExecuteButtons) window.refreshExecuteButtons();
            return;
        }
        window.__executeButtonsInitialized = true;

        function findCopyButton(container) {
            if (!container) return null;
            const allButtons = container.querySelectorAll('[role=""button""]');
            for (let i = 0; i < allButtons.length; i++) {
                const btn = allButtons[i];
                const text = (btn.textContent || '').trim();
                if (text.includes('复制') || text.includes('Copy')) return btn;
            }
            return null;
        }

        function getCurrentCodeContent(codeBlock) {
            try {
                const pre = codeBlock.querySelector('pre');
                if (!pre) return { code: '', language: 'unknown' };
                const codeContent = pre.textContent || '';
                let language = 'unknown';
                const spans = codeBlock.querySelectorAll('span');
                for (const s of spans) {
                    const t = (s.textContent || '').trim().toLowerCase();
                    if (t && /^(vba|javascript|js|excel|python|sql|typescript|html|css|c#|java|php|csharp)$/i.test(t)) {
                        language = t; break;
                    }
                }
                return { code: codeContent, language: language };
            } catch (e) {
                console.error('[Execute Buttons] 获取代码失败', e);
                return { code: '', language: 'unknown' };
            }
        }

        // 创建一个不会被页面禁用的可点击元素（使用 div 而不是 button）
        function createExecuteElement() {
            const el = document.createElement('div');
            el.setAttribute('role','button');
            el.setAttribute('aria-disabled','false');
            el.className = 'vsto-execute-button ds-text-button';
            el.tabIndex = 0;
            el.style.display = 'inline-flex';
            el.style.alignItems = 'center';
            el.style.marginRight = '4px';
            el.style.cursor = 'pointer';
            el.style.userSelect = 'none';
            el.innerHTML = '<div class=""ds-button__icon""><div class=""ds-icon"" style=""font-size:16px;width:16px;"">▶</div></div><span class=""code-info-button-text"">执行</span>';
            return el;
        }

        function attachClickHandler(executeEl, codeBlock) {
            const handler = function(e) {
                e.preventDefault();
                e.stopPropagation();
                console.log('[Execute Buttons] 执行（自定义元素）被点击');
                const content = getCurrentCodeContent(codeBlock);
                if (!content.code || !content.code.trim()) {
                    console.log('[Execute Buttons] 代码为空，跳过');
                    return;
                }
                try {
                    if (window.vsto && typeof window.vsto.executeCode === 'function') {
                        window.vsto.executeCode(content.code, content.language, true);
                        console.log('[Execute Buttons] 通过 vsto 执行请求发送');
                    } else if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {
                        window.chrome.webview.postMessage({
                            type: 'executeCode',
                            code: content.code,
                            language: content.language,
                            executecodePreview: true
                        });
                        console.log('[Execute Buttons] 通过 chrome.webview 发送执行请求');
                    } else {
                        console.error('[Execute Buttons] 无通信接口');
                    }
                } catch (err) {
                    console.error('[Execute Buttons] 发送执行请求失败', err);
                }
            };
            executeEl.addEventListener('click', handler);
            // 键盘可访问性
            executeEl.addEventListener('keydown', function(ev) {
                if (ev.key === 'Enter' || ev.key === ' ') {
                    ev.preventDefault();
                    this.click();
                }
            });
        }

        function processCodeBlock(codeBlock, index) {
            try {
                if (!codeBlock.classList.contains('md-code-block')) return false;
                const copyBtn = findCopyButton(codeBlock);
                if (!copyBtn) return false;
                const container = copyBtn.parentElement;
                if (!container) return false;
                if (container.querySelector('.vsto-execute-button')) return false;
                const pre = codeBlock.querySelector('pre');
                if (!pre) return false;

                console.log('[Execute Buttons] 创建自定义执行元素');
                const execEl = createExecuteElement();

                // 将自定义元素插入到 copyBtn 之前
                container.insertBefore(execEl, copyBtn);

                // 挂载事件
                attachClickHandler(execEl, codeBlock);

                // 防护：确保页面不会把它标记为禁用（定期清理 + 观察）
                execEl.style.pointerEvents = 'auto';
                execEl.removeAttribute('disabled');
                execEl.setAttribute('aria-disabled','false');

                console.log('[Execute Buttons] 自定义执行元素插入完成');
                return true;
            } catch (ex) {
                console.error('[Execute Buttons] 处理代码块失败', ex);
                return false;
            }
        }

        function addExecuteButtons() {
            const codeBlocks = document.querySelectorAll('.md-code-block');
            if (!codeBlocks || codeBlocks.length === 0) return;
            let count = 0;
            codeBlocks.forEach((b,i) => { if (processCodeBlock(b,i)) count++; });
            console.log('[Execute Buttons] 处理完成: ' + count + '/' + codeBlocks.length);
        }

        // 定期修复：移除页面可能添加的 disabled/禁用类
        function keepAlive() {
            document.querySelectorAll('.vsto-execute-button').forEach(b => {
                try {
                    b.removeAttribute('disabled');
                    b.setAttribute('aria-disabled','false');
                    b.classList.remove('ds-atom-button--disabled', 'ds-text-button--disabled', 'execute-code-button');
                    b.style.pointerEvents = 'auto';
                    b.style.opacity = ''; // 如果页面设置了半透明，也恢复
                } catch(e) {}
            });
        }

        // 观察父容器，若页面替换、修改节点则重新注入
        const observer = new MutationObserver(function(mutations) {
            let shouldRun = false;
            for (const m of mutations) {
                if (m.addedNodes && m.addedNodes.length > 0) {
                    shouldRun = true; break;
                }
                if (m.type === 'attributes' && (m.attributeName === 'class' || m.attributeName === 'disabled' || m.attributeName === 'aria-disabled')) {
                    shouldRun = true; break;
                }
            }
            if (shouldRun) {
                setTimeout(addExecuteButtons, 120);
            }
        });
        observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class','disabled','aria-disabled'] });

        // 初始化与周期修复
        setTimeout(addExecuteButtons, 120);
        [500,1000,2000,3000,5000].forEach((d,i) => setTimeout(addExecuteButtons, d));
        setInterval(keepAlive, 1000);

        // 提供外部手动刷新
        window.refreshExecuteButtons = addExecuteButtons;

        console.log('[Execute Buttons] 注入完成（自定义元素策略）');
    })();
    "

            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            Debug.WriteLine("执行按钮注入脚本已执行（自定义元素版本）")
        Catch ex As Exception
            Debug.WriteLine($"注入执行按钮脚本时出错: {ex.Message}")
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
                Debug.WriteLine("已保存Deepseek会话信息")
            End If
        Catch ex As Exception
            Debug.WriteLine($"保存会话信息时出错: {ex.Message}")
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
                                                                                     Debug.WriteLine($"添加Cookie '{parts(0)}'时出错: {cookieEx.Message}")
                                                                                 End Try
                                                                             End If
                                                                         Next
                                                                     Catch ex As Exception
                                                                         Debug.WriteLine($"在UI线程恢复Cookies时出错: {ex.Message}")
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
                            Debug.WriteLine($"添加Cookie '{parts(0)}'时出错: {cookieEx.Message}")
                        End Try
                    End If
                Next
            End If

            ' 短暂延迟确保Cookie操作完成
            Await Task.Delay(100)
        Catch ex As Exception
            Debug.WriteLine($"恢复Cookies时出错: {ex.Message}")
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
                        Debug.WriteLine("清理令牌格式时出错: " & ex.Message)
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
            Debug.WriteLine($"提取授权令牌时出错: {ex.Message}")
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
                                           Debug.WriteLine($"获取Cookies时出错: {ex.Message}")
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
            Debug.WriteLine($"获取Cookies时出错: {ex.Message}")
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
                                                              Debug.WriteLine($"执行脚本时出错: {ex.Message}")
                                                          End Try
                                                      End Sub)
                               End Sub)
            Else
                Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            End If
        Catch ex As Exception
            Debug.WriteLine($"注入授权令牌时出错: {ex.Message}")
        End Try
    End Function

    ' 添加到类中的新字段
    Private ReadOnly SessionFilePath As String = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
    ConfigSettings.OfficeAiAppDataFolder,
    "deepseek_session.json")
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
                        Debug.WriteLine("会话已过期，需要重新登录")
                        Return
                    End If
                Else
                    Debug.WriteLine("无效的会话时间戳")
                    Return
                End If

                ' 恢复Cookie
                If sessionInfo.ContainsKey("cookies") AndAlso
               Not String.IsNullOrEmpty(sessionInfo("cookies")?.ToString()) Then
                    Await RestoreCookiesAsync(sessionInfo("cookies").ToString())
                    Debug.WriteLine("已恢复Cookies")
                End If

                ' 注入授权令牌
                If sessionInfo.ContainsKey("authToken") AndAlso
               Not String.IsNullOrEmpty(sessionInfo("authToken")?.ToString()) Then
                    Await InjectAuthTokenAsync(sessionInfo("authToken").ToString())
                    Debug.WriteLine("已注入授权令牌")
                End If

                Debug.WriteLine("已恢复Deepseek会话信息")
            Else
                Debug.WriteLine("未找到会话文件，需要重新登录")
            End If
        Catch ex As Exception
            Debug.WriteLine($"恢复会话信息时出错: {ex.Message}")
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

    Public Overrides ReadOnly Property ChatUrl As String = "https://chat.deepseek.com"
    Public Overrides ReadOnly Property SessionFileName As String = "deepseek_session.json"

    ' 存储聊天HTML的文件路径
    Protected ReadOnly ChatHtmlFilePath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder,
        $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
    )





    ' 执行JavaScript代码 - 专注于操作Office/WPS对象模型，支持Office JS API风格代码
    Protected Overrides Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean
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

    Protected Overrides Function GetWebView2DataFolderName() As String
        Return "DeepseekChatWebView2Data"
    End Function
End Class