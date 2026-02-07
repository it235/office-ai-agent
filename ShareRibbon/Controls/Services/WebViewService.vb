' ShareRibbon\Controls\Services\WebViewService.vb
' WebView2 初始化、配置和脚本执行服务

Imports System.IO
Imports System.Threading.Tasks
Imports System.Web
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json

''' <summary>
''' WebView2 服务类，负责 WebView2 初始化、配置和脚本执行
''' </summary>
Public Class WebViewService
        Private ReadOnly _chatBrowser As WebView2
        Private ReadOnly _getApplication As Func(Of ApplicationInfo)
        Private _wwwRoot As String

        ''' <summary>
        ''' WebView2 导航完成事件
        ''' </summary>
        Public Event NavigationCompleted As EventHandler(Of CoreWebView2NavigationCompletedEventArgs)

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="chatBrowser">WebView2 控件实例</param>
        ''' <param name="getApplication">获取应用信息的委托</param>
        Public Sub New(chatBrowser As WebView2, getApplication As Func(Of ApplicationInfo))
            _chatBrowser = chatBrowser
            _getApplication = getApplication
        End Sub

        ''' <summary>
        ''' 初始化 WebView2
        ''' </summary>
        Public Async Function InitializeAsync() As Task
            Try
                ' 自定义用户数据目录
                Dim userDataFolder As String = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MyAppWebView2Cache")

                ' 确保目录存在
                If Not Directory.Exists(userDataFolder) Then
                    Directory.CreateDirectory(userDataFolder)
                End If

                ' 释放资源文件到本地
                _wwwRoot = ResourceExtractor.ExtractResources()

                ' 配置 WebView2 的创建属性
                _chatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
                    .UserDataFolder = userDataFolder
                }

                ' 初始化 WebView2
                Await _chatBrowser.EnsureCoreWebView2Async(Nothing)

                ' 确保 CoreWebView2 已初始化
                If _chatBrowser.CoreWebView2 IsNot Nothing Then
                    ConfigureSettings()
                    SetupVirtualHostMapping()
                    LoadHtmlTemplate()
                    ConfigureMarked()
                    AddHandler _chatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnNavigationCompleted
                Else
                    MessageBox.Show("WebView2 初始化失败，CoreWebView2 不可用。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Catch ex As Exception
                Dim errorMessage As String = $"初始化失败: {ex.Message}{Environment.NewLine}类型: {ex.GetType().Name}{Environment.NewLine}堆栈:{ex.StackTrace}"
                MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Function

        ''' <summary>
        ''' 配置 WebView2 安全设置
        ''' </summary>
        Private Sub ConfigureSettings()
            _chatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
            _chatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
            _chatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
            _chatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True
        End Sub

        ''' <summary>
        ''' 设置虚拟主机名映射
        ''' </summary>
        Private Sub SetupVirtualHostMapping()
            _chatBrowser.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "officeai.local",
                _wwwRoot,
                CoreWebView2HostResourceAccessKind.Allow)
        End Sub

        ''' <summary>
        ''' 加载 HTML 模板
        ''' </summary>
        Private Sub LoadHtmlTemplate()
        Dim htmlContent As String = My.Resources.chat_template_refactored
        _chatBrowser.CoreWebView2.NavigateToString(htmlContent)
    End Sub

        ''' <summary>
        ''' 配置 Marked.js
        ''' </summary>
        Private Async Sub ConfigureMarked()
            If _chatBrowser.CoreWebView2 IsNot Nothing Then
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
                Await _chatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            End If
        End Sub

        ''' <summary>
        ''' 导航完成事件处理
        ''' </summary>
        Private Sub OnNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
            If e.IsSuccess Then
                RaiseEvent NavigationCompleted(sender, e)
                ' 移除事件处理器，避免重复触发
                RemoveHandler _chatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnNavigationCompleted
            End If
        End Sub

        ''' <summary>
        ''' 异步执行 JavaScript 脚本
        ''' </summary>
        ''' <param name="js">JavaScript 代码</param>
        Public Async Function ExecuteScriptAsync(js As String) As Task
            If _chatBrowser.InvokeRequired Then
                _chatBrowser.Invoke(Sub() _chatBrowser.ExecuteScriptAsync(js))
            Else
                Await _chatBrowser.ExecuteScriptAsync(js)
            End If
        End Function

        ''' <summary>
        ''' 同步方式在 UI 线程执行操作
        ''' </summary>
        Public Sub InvokeIfRequired(action As Action)
            If _chatBrowser.InvokeRequired Then
                _chatBrowser.Invoke(action)
            Else
                action()
            End If
        End Sub

        ''' <summary>
        ''' 注入 VSTO 辅助脚本
        ''' </summary>
        Public Sub InjectVstoScript()
            Dim script As String = "
            window.vsto = {
                executeCode: function(code, language, preview) {
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
            "
            _chatBrowser.ExecuteScriptAsync(script)
        End Sub

        ''' <summary>
        ''' 初始化前端设置
        ''' </summary>
        Public Sub InitializeSettings(chatSettings As ChatSettings)
            Dim js As String = $"
            document.getElementById('topic-randomness').value = '{chatSettings.topicRandomness}';
            document.getElementById('topic-randomness-value').textContent = '{chatSettings.topicRandomness}';
            document.getElementById('context-limit').value = '{chatSettings.contextLimit}';
            document.getElementById('context-limit-value').textContent = '{chatSettings.contextLimit}';
            document.getElementById('settings-scroll-checked').checked = {chatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('settings-selected-cell').checked = {chatSettings.selectedCellChecked.ToString().ToLower()};
            document.getElementById('settings-executecode-preview').checked = {chatSettings.executecodePreviewChecked.ToString().ToLower()};
            
            var selectElement = document.getElementById('chatMode');
            if (selectElement) {{
                selectElement.value = '{chatSettings.chatMode}';
            }}
            
            document.getElementById('scrollChecked').checked = {chatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('selectedCell').checked = {chatSettings.selectedCellChecked.ToString().ToLower()};
            "
            ExecuteScriptAsync(js)
        End Sub

        ''' <summary>
        ''' 添加选中内容到聊天区
        ''' </summary>
        Public Async Sub AddSelectedContentItem(sheetName As String, address As String)
            Dim ctrlKey As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control
            Await _chatBrowser.CoreWebView2.ExecuteScriptAsync(
                $"addSelectedContentItem({JsonConvert.SerializeObject(sheetName)}, {JsonConvert.SerializeObject(address)}, {ctrlKey.ToString().ToLower()})")
        End Sub

        ''' <summary>
        ''' 清除指定工作表的选中内容
        ''' </summary>
        Public Async Sub ClearSelectedContentBySheetName(sheetName As String)
            Await _chatBrowser.CoreWebView2.ExecuteScriptAsync(
                $"clearSelectedContentBySheetName({JsonConvert.SerializeObject(sheetName)})")
        End Sub

        ''' <summary>
        ''' 获取 CoreWebView2 实例
        ''' </summary>
        Public ReadOnly Property CoreWebView2 As CoreWebView2
            Get
                Return _chatBrowser.CoreWebView2
            End Get
        End Property

        ''' <summary>
        ''' 获取 WebView2 控件
        ''' </summary>
        Public ReadOnly Property Browser As WebView2
            Get
                Return _chatBrowser
            End Get
        End Property
    End Class
