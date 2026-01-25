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
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseChat
    Inherits UserControl

    ' 公共字段 - 子类可以覆盖
    Public MustOverride ReadOnly Property ChatUrl As String
    Public MustOverride ReadOnly Property SessionFileName As String

    ' 公共属性
    Protected Property ChatBrowser As Microsoft.Web.WebView2.WinForms.WebView2

    ' 构造函数
    Protected Sub New()
        ' 确保WebView2控件可以正常交互
    End Sub

    ' 粘贴处理
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_PASTE As Integer = &H302
        If m.Msg = WM_PASTE Then
            ' 在此处理粘贴操作，比如：
            If Clipboard.ContainsText() Then
                Dim txt As String = Clipboard.GetText()
                ' 可以在这里添加自定义的粘贴逻辑
            End If
            ' 不把消息传递给基类，从而拦截后续处理  
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    ' 焦点管理
    Protected Overrides Sub OnGotFocus(e As EventArgs)
        MyBase.OnGotFocus(e)
        ' 当控件获得焦点时，确保WebView2也能接收焦点
        If ChatBrowser IsNot Nothing Then
            ChatBrowser.Focus()
        End If
    End Sub

    Protected Overrides Sub OnClick(e As EventArgs)
        MyBase.OnClick(e)
        ' 确保点击时WebView2获得焦点
        If ChatBrowser IsNot Nothing Then
            ChatBrowser.Focus()
        End If
    End Sub

    ' WebView2初始化（子类可以覆盖以添加特定配置）
    Protected Async Function InitializeWebView2() As Task
        Try
            ' 使用固定的用户数据目录而不是临时目录，以保持会话持久化
            Dim userDataFolder As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ConfigSettings.OfficeAiAppDataFolder,
                GetWebView2DataFolderName())

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

                ' 导航到目标网站（由子类定义）
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

    ' 子类必须实现的WebView2数据文件夹名称
    Protected MustOverride Function GetWebView2DataFolderName() As String

    ' 在页面加载完成后，注入脚本 - 修复线程问题
    Private Async Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If e.IsSuccess Then
            Try
                Debug.WriteLine("导航完成，开始注入脚本")

                ' 延迟一些时间，确保页面完全加载
                Await Task.Delay(GetScriptInjectionDelay())

                ' 确保在UI线程上执行所有WebView2操作
                If ChatBrowser.InvokeRequired Then
                    ChatBrowser.Invoke(New Action(Async Sub()
                                                      Try
                                                          ' 配置Marked和代码高亮（如果支持）
                                                          Await ConfigureMarkedSafe()

                                                          ' 注入基础辅助脚本
                                                          Await InitializeWebView2ScriptAsyncSafe()

                                                          ' 初始化设置和执行按钮
                                                          Await InitializeSettingsSafe()

                                                          ' 注入登录观察器（如果支持）
                                                          Await InjectLoginObserverSafe()

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
                                     Await Task.Delay(GetScriptInjectionDelay())

                                     ' 在UI线程上执行脚本注入
                                     ChatBrowser.Invoke(New Action(Async Sub()
                                                                       Try
                                                                           ' 配置Marked和代码高亮（如果支持）
                                                                           Await ConfigureMarkedSafe()

                                                                           ' 注入基础辅助脚本
                                                                           Await InitializeWebView2ScriptAsyncSafe()

                                                                           ' 初始化设置和执行按钮
                                                                           Await InitializeSettingsSafe()

                                                                           ' 注入登录观察器（如果支持）
                                                                           Await InjectLoginObserverSafe()

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

    ' 子类可以覆盖以自定义脚本注入延迟
    Protected Overridable Function GetScriptInjectionDelay() As Integer
        Return 1000 ' 默认1秒
    End Function

    ' 线程安全的ConfigureMarked方法 - 子类可以覆盖
    Protected Overridable Async Function ConfigureMarkedSafe() As Task
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

    ' 线程安全的InitializeWebView2ScriptAsync方法 - 子类可以覆盖以添加特定功能
    Protected Overridable Async Function InitializeWebView2ScriptAsyncSafe() As Task
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

    ' 线程安全的InitializeSettings方法 - 子类可以覆盖
    Protected Overridable Async Function InitializeSettingsSafe() As Task
        Try
            ' 默认情况下调用执行按钮注入
            Await InjectExecuteButtonsSafe()
            Debug.WriteLine("InitializeSettings执行完成")
        Catch ex As Exception
            Debug.WriteLine($"InitializeSettings出错: {ex.Message}")
        End Try
    End Function

    ' 线程安全的执行按钮注入方法 - 子类必须实现
    Protected MustOverride Async Function InjectExecuteButtonsSafe() As Task

    ' 线程安全的登录观察器注入方法 - 子类可以覆盖
    Protected Overridable Async Function InjectLoginObserverSafe() As Task
        ' 默认实现为空，子类可以覆盖
    End Function

    ' WebView2消息接收处理
    Protected Sub WebView2_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim jsonDoc As JObject = JObject.Parse(e.WebMessageAsJson)
            Dim messageType As String = jsonDoc("type").ToString()

            Select Case messageType
                Case "executeCode"
                    HandleExecuteCode(jsonDoc)
                Case Else
                    Debug.WriteLine($"未知消息类型: {messageType}")
            End Select
        Catch ex As Exception
            Debug.WriteLine($"处理消息出错: {ex.Message}")
        End Try
    End Sub

    ' 执行代码处理
    Protected Overridable Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim preview As Boolean = Boolean.Parse(jsonDoc("executecodePreview"))
        Dim language As String = jsonDoc("language").ToString()
        ExecuteCode(code, language, preview)
    End Sub

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
            Case Else
                GlobalStatusStrip.ShowWarning("不支持的语言类型: " & language)
        End Select
    End Sub

    ' 执行VBA代码 - 子类必须实现' 执行前端传来的 VBA 代码片段
    Protected Function ExecuteVBACode(vbaCode As String, preview As Boolean) As Boolean

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

    ' 执行JavaScript代码 - 子类必须实现  
    Protected MustOverride Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean

    ' 执行Excel公式 - 子类必须实现

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

    ' 获取Office应用对象 - 需要在具体子类中实现
    Protected Overridable Function GetOfficeApplicationObject() As Object
        Try
            Select Case GetApplication().Type
                Case OfficeApplicationType.Excel
                    ' 由子类实现具体的Application访问
                    Return Nothing
                Case OfficeApplicationType.Word
                    ' 由子类实现具体的Application访问
                    Return Nothing
                Case OfficeApplicationType.PowerPoint
                    ' 由子类实现具体的Application访问
                    Return Nothing
                Case Else
                    Return Nothing
            End Select
        Catch ex As Exception
            Debug.WriteLine("获取Office应用对象失败: " & ex.Message)
            Return Nothing
        End Try
    End Function


    ' 抽象方法 - 子类必须实现
    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String
    Protected MustOverride Function GetApplication() As ApplicationInfo
    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCodePreview(vbaCode As String, preview As Boolean)
    Protected MustOverride Function RunCode(vbaCode As String)
    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)

    ' 注入脚本
    Private Async Sub InjectScript(scriptContent As String)
        If ChatBrowser.CoreWebView2 IsNot Nothing Then
            Dim escapedScript = JsonConvert.SerializeObject(scriptContent)
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync($"eval({escapedScript})")
        Else
            MessageBox.Show("CoreWebView2 未初始化，无法注入脚本。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

End Class