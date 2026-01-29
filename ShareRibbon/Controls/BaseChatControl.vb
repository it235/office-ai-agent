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
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseChatControl
    Inherits UserControl

    ' 服务类实例
    Private _fileParserService As New FileParserService()
    Protected _chatStateService As New ChatStateService()
    Private _historyService As HistoryService = Nothing
    Private _mcpService As McpService = Nothing
    
    ' Ralph Loop 控制器
    Protected _ralphLoopController As New RalphLoopController()
    
    ' Ralph Agent 控制器
    Protected _ralphAgentController As RalphAgentController

    ' 延迟初始化的历史服务
    Protected ReadOnly Property HistoryService As HistoryService
        Get
            If _historyService Is Nothing Then
                _historyService = New HistoryService(AddressOf ExecuteJavaScriptAsyncJS)
            End If
            Return _historyService
        End Get
    End Property

    ' 延迟初始化的 MCP 服务
    Protected ReadOnly Property McpService As McpService
        Get
            If _mcpService Is Nothing Then
                _mcpService = New McpService(AddressOf ExecuteJavaScriptAsyncJS, AddressOf GetApplication)
            End If
            Return _mcpService
        End Get
    End Property

    ' 延迟初始化的代码执行服务
    Private _codeExecutionService As CodeExecutionService = Nothing
    Protected ReadOnly Property CodeExecutionService As CodeExecutionService
        Get
            If _codeExecutionService Is Nothing Then
                _codeExecutionService = New CodeExecutionService(
                    AddressOf GetVBProject,
                    AddressOf GetOfficeApplicationObject,
                    AddressOf GetApplication,
                    AddressOf RunCode,
                    AddressOf RunCodePreview,
                    AddressOf EvaluateFormula)
                ' 设置JSON命令执行器（由子类提供）
                _codeExecutionService.JsonCommandExecutor = AddressOf ExecuteJsonCommand
            End If
            Return _codeExecutionService
        End Get
    End Property

    ''' <summary>
    ''' 执行JSON命令（由子类重写以提供具体实现）
    ''' </summary>
    Protected Overridable Function ExecuteJsonCommand(jsonCode As String, preview As Boolean) As Boolean
        ' 默认实现：不支持JSON命令
        GlobalStatusStrip.ShowWarning("当前应用不支持JSON命令执行")
        Return False
    End Function

    ' 延迟初始化的意图识别服务
    Private _intentService As IntentRecognitionService = Nothing
    Protected ReadOnly Property IntentService As IntentRecognitionService
        Get
            If _intentService Is Nothing Then
                ' 根据当前Office应用类型初始化意图识别服务
                Dim appInfo = GetApplication()
                If appInfo IsNot Nothing Then
                    _intentService = New IntentRecognitionService(appInfo.Type)
                Else
                    _intentService = New IntentRecognitionService()
                End If
            End If
            Return _intentService
        End Get
    End Property

    ' 当前意图结果（用于子类访问）
    Protected CurrentIntentResult As IntentResult = Nothing

    ' 意图预览相关字段
    Private _pendingIntentMessage As String = Nothing
    Private _pendingIntentResult As IntentResult = Nothing
    Private _pendingFilePaths As List(Of String) = Nothing

    'settings
    Protected topicRandomness As Double
    Protected contextLimit As Integer
    Protected selectedCellChecked As Boolean = False
    Protected settingsScrollChecked As Boolean = False

    Protected stopReaderStream As Boolean = False


    ' ai的历史回复
    Protected systemHistoryMessageData As New List(Of HistoryMessage)

    Protected loadingPictureBox As PictureBox

    ' 选区对比相关字段
    Protected PendingSelectionInfo As SelectionInfo = Nothing
    Protected _selectionPendingMap As New Dictionary(Of String, SelectionInfo)()
    Private allPlainMarkdownBuffer As New StringBuilder()

    Protected _responseToRequestMap As New Dictionary(Of String, String)()
    Protected _revisionsMap As New Dictionary(Of String, JArray)()

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_PASTE As Integer = &H302
        If m.Msg = WM_PASTE Then
            If Clipboard.ContainsText() Then
                Dim txt As String = Clipboard.GetText()
            End If
            Return
        End If
        MyBase.WndProc(m)
    End Sub

    Protected Async Function InitializeWebView2() As Task
        Try
            Dim userDataFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyAppWebView2Cache")
            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            Dim wwwRoot As String = ResourceExtractor.ExtractResources()
            ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
                .UserDataFolder = userDataFolder
            }
            Await ChatBrowser.EnsureCoreWebView2Async(Nothing)

            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                ChatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                ChatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True

                ChatBrowser.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "officeai.local",
                    wwwRoot,
                    CoreWebView2HostResourceAccessKind.Allow
                )

                Dim htmlContent As String = My.Resources.chat_template_refactored
                ChatBrowser.CoreWebView2.NavigateToString(htmlContent)
                ConfigureMarked()
                AddHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted

            Else
                MessageBox.Show("WebView2 初始化失败，CoreWebView2 不可用。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"初始化失败: {ex.Message}{Environment.NewLine}类型: {ex.GetType().Name}{Environment.NewLine}堆栈:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
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


    ' 动态ChatHtmlFilePath属性
    Protected ReadOnly Property ChatHtmlFilePath As String
        Get
            ' 如果已经生成过文件路径，直接返回缓存的路径
            If Not String.IsNullOrEmpty(_chatHtmlFilePath) Then
                Return _chatHtmlFilePath
            End If

            Dim baseDir As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder
        )

            Dim fileName As String
            If Not String.IsNullOrEmpty(firstQuestion) Then
                ' 简单地取前10个字符
                Dim questionPrefix As String = GetFirst10Characters(firstQuestion)
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}_{questionPrefix}.html"
            Else
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
            End If

            _chatHtmlFilePath = Path.Combine(baseDir, fileName)
            Return _chatHtmlFilePath
        End Get
    End Property

    Private Function GetFirst10Characters(text As String) As String
        Return UtilsService.GetFirst10Characters(text)
    End Function

    Private Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles ChatBrowser.NavigationCompleted
        If e.IsSuccess Then
            Try
                If ChatBrowser.InvokeRequired Then
                    ' 使用同步的 Invoke 而不是异步的
                    ChatBrowser.Invoke(Sub()
                                           Task.Delay(100).Wait() ' 同步等待
                                           InitializeSettings()
                                           InitializeMcpSettings() ' 添加MCP初始化

                                           ' 直接在UI线程移除事件处理器
                                           If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                                               RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                                           End If
                                       End Sub)
                Else
                    Task.Delay(100).Wait() ' 同步等待
                    InitializeSettings()
                    InitializeMcpSettings() ' 添加MCP初始化

                    ' 直接在UI线程移除事件处理器
                    If ChatBrowser IsNot Nothing AndAlso ChatBrowser.CoreWebView2 IsNot Nothing Then
                        RemoveHandler ChatBrowser.CoreWebView2.NavigationCompleted, AddressOf OnWebViewNavigationCompleted
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"导航完成事件处理中出错: {ex.Message}")
                Debug.WriteLine(ex.StackTrace)
            End Try
        End If
    End Sub

    Protected Sub InitializeSettings()
        Try
            ' 加载设置
            Dim chatSettings As New ChatSettings(GetApplication())
            selectedCellChecked = ChatSettings.selectedCellChecked
            contextLimit = ChatSettings.contextLimit
            topicRandomness = ChatSettings.topicRandomness
            settingsScrollChecked = ChatSettings.settingsScrollChecked

            ' 设置Office应用类型（用于前端区分Word/PPT/Excel）
            Dim appType = GetOfficeAppType()

            ' 将设置发送到前端
            Dim js As String = $"
            window.officeAppType = '{appType}';
            document.getElementById('topic-randomness').value = '{ChatSettings.topicRandomness}';
            document.getElementById('topic-randomness-value').textContent = '{ChatSettings.topicRandomness}';
            document.getElementById('context-limit').value = '{ChatSettings.contextLimit}';
            document.getElementById('context-limit-value').textContent = '{ChatSettings.contextLimit}';
            document.getElementById('settings-scroll-checked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('settings-selected-cell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
            document.getElementById('settings-executecode-preview').checked = {ChatSettings.executecodePreviewChecked.ToString().ToLower()};
            
            // 初始化自动补全设置
            var autocompleteCheckbox = document.getElementById('settings-autocomplete-enable');
            if (autocompleteCheckbox) {{
                autocompleteCheckbox.checked = {ChatSettings.EnableAutocomplete.ToString().ToLower()};
            }}
            var shortcutSelect = document.getElementById('settings-autocomplete-shortcut');
            if (shortcutSelect) {{
                shortcutSelect.value = '{ChatSettings.AutocompleteShortcut}';
            }}
            if (typeof updateAutocompleteSettings === 'function') {{
                updateAutocompleteSettings({{ enabled: {ChatSettings.EnableAutocomplete.ToString().ToLower()}, delayMs: {ChatSettings.AutocompleteDelayMs}, shortcut: '{ChatSettings.AutocompleteShortcut}' }});
            }}
            
            var selectElement = document.getElementById('chatMode');
            if (selectElement) {{
                selectElement.value = '{ChatSettings.chatMode}';
            }}
            
            // 同步到主界面的checkbox
            document.getElementById('scrollChecked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('selectedCell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
        "
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"初始化设置失败: {ex.Message}")
        End Try
    End Sub

    Protected Sub WebView2_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim jsonDoc As JObject = JObject.Parse(e.WebMessageAsJson)
            Dim messageType As String = jsonDoc("type").ToString()

            Select Case messageType
                Case "checkedChange"
                    HandleCheckedChange(jsonDoc)
                Case "sendMessage"
                    HandleSendMessage(jsonDoc)
                Case "stopMessage"
                    stopReaderStream = True
                Case "executeCode"
                    HandleExecuteCode(jsonDoc)
                Case "saveSettings"
                    HandleSaveSettings(jsonDoc)
                Case "getHistoryFiles"
                    HandleGetHistoryFiles()
                Case "openHistoryFile"
                    HandleOpenHistoryFile(jsonDoc)
                Case "getMcpConnections"
                    HandleGetMcpConnections()
                Case "saveMcpSettings"
                    HandleSaveMcpSettings(jsonDoc)
                Case "clearContext"
                    ClearChatContext()
                Case "acceptAnswer"
                    HandleAcceptAnswer(jsonDoc)
                Case "rejectAnswer"
                    HandleRejectAnswer(jsonDoc)
                Case "applyRevisionAll"
                    HandleApplyRevisionAll(jsonDoc)
                Case "applyRevisionSegment"
                    HandleApplyRevisionSegment(jsonDoc)
                Case "applyDocumentPlanItem"
                    HandleApplyDocumentPlanItem(jsonDoc)
                Case "rejectShowComparison"
                    ' 排版答案内容格式有误，重试

                Case "applyRevisionAccept" ' 前端请求接受单个 Revision
                    HandleApplyRevisionAccept(jsonDoc)
                Case "applyRevisionReject" ' 前端请求拒绝单个 Revision
                    HandleApplyRevisionReject(jsonDoc)

                ' 续写功能消息处理
                Case "triggerContinuation"
                    HandleTriggerContinuation(jsonDoc)
                Case "applyContinuation"
                    HandleApplyContinuation(jsonDoc)
                Case "refineContinuation"
                    HandleRefineContinuation(jsonDoc)

                ' 模板渲染功能消息处理
                Case "applyTemplateContent"
                    HandleApplyTemplateContent(jsonDoc)
                Case "refineTemplateContent"
                    HandleRefineTemplateContent(jsonDoc)

                ' 自动补全功能消息处理
                Case "requestCompletion"
                    HandleRequestCompletion(jsonDoc)
                Case "acceptCompletion"
                    HandleAcceptCompletion(jsonDoc)

                ' 意图预览功能消息处理
                Case "confirmIntent"
                    HandleConfirmIntent(jsonDoc)
                Case "cancelIntent"
                    HandleCancelIntent()

                ' Ralph Loop 循环功能消息处理
                Case "continueLoop"
                    HandleContinueLoop()
                Case "cancelLoop"
                    HandleCancelLoop()
                Case "startLoop"
                    HandleStartLoop(jsonDoc)

                ' Ralph Agent 智能助手消息处理
                Case "startAgent"
                    HandleStartAgent(jsonDoc)
                Case "startAgentExecution"
                    HandleStartAgentExecution(jsonDoc)
                Case "abortAgent"
                    HandleAbortAgent()

                ' 文件选择对话框
                Case "openFileDialog"
                    HandleOpenFileDialog()

                Case Else
                    Debug.WriteLine($"未知消息类型: {messageType}")
            End Select
        Catch ex As Exception
            Debug.WriteLine($"处理消息出错: {ex.Message}")
        End Try
    End Sub

    ' 添加：在基类提供可覆盖的 CaptureCurrentSelectionInfo（默认返回 Nothing，Word 子类会覆写）
    Protected Overridable Function CaptureCurrentSelectionInfo(mode As String) As SelectionInfo
        Return Nothing
    End Function


    ' 在基类提供默认的 applyRevision 处理（子类可覆盖）
    Protected Overridable Sub HandleApplyRevisionAll(jsonDoc As JObject)
    End Sub

    Protected Overridable Sub HandleApplyRevisionSegment(jsonDoc As JObject)
    End Sub


    Protected Overridable Sub HandleApplyRevisionReject(jsonDoc As JObject)
        Debug.WriteLine("收到 applyRevisionReject 请求（基类默认不做写回）")
        GlobalStatusStrip.ShowInfo("用户拒绝了该修订（未在基类执行写回）")
    End Sub

    Protected Overridable Sub HandleApplyRevisionAccept(jsonDoc As JObject)
    End Sub


    ' ========== 续写功能相关方法 ==========

    ''' <summary>
    ''' 触发续写（由子类实现具体逻辑）
    ''' </summary>
    ''' <param name="jsonDoc">包含style参数的JSON对象</param>
    Protected Overridable Sub HandleTriggerContinuation(jsonDoc As JObject)
        Debug.WriteLine("HandleTriggerContinuation 被调用（基类默认不执行）")
        GlobalStatusStrip.ShowWarning("当前应用不支持续写功能")
    End Sub

    ''' <summary>
    ''' 应用续写结果到文档（由子类实现）
    ''' </summary>
    Protected Overridable Sub HandleApplyContinuation(jsonDoc As JObject)
        Debug.WriteLine("HandleApplyContinuation 被调用（基类默认不执行）")
    End Sub

    ''' <summary>
    ''' 调整续写（多轮对话）
    ''' </summary>
    Protected Overridable Sub HandleRefineContinuation(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim refinement As String = If(jsonDoc("refinement") IsNot Nothing, jsonDoc("refinement").ToString(), String.Empty)

            If String.IsNullOrWhiteSpace(refinement) Then
                GlobalStatusStrip.ShowWarning("请输入调整方向")
                Return
            End If

            ' 构建调整提示
            Dim refinementPrompt As New StringBuilder()
            refinementPrompt.AppendLine("请根据以下要求调整之前的续写内容：")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine($"【调整要求】{refinement}")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("请直接输出调整后的续写内容，不要添加任何解释：")

            ' 保持 responseMode = "continuation"，发送调整请求（不使用历史记录）
            Task.Run(Async Function()
                         Await Send(refinementPrompt.ToString(), GetContinuationSystemPrompt(), False, "continuation")
                     End Function)

            GlobalStatusStrip.ShowInfo("正在调整续写内容...")
        Catch ex As Exception
            Debug.WriteLine($"HandleRefineContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("调整续写时出错")
        End Try
    End Sub

    ''' <summary>
    ''' 发送续写请求（不使用聊天历史记录）
    ''' </summary>
    Protected Sub SendContinuationRequest(context As ContinuationContext, Optional style As String = "")
        Dim systemPrompt = GetContinuationSystemPrompt()
        Dim userPrompt = BuildContinuationUserPrompt(context, style)

        Task.Run(Async Function()
                     Await Send(userPrompt, systemPrompt, False, "continuation")
                 End Function)
    End Sub

    ''' <summary>
    ''' 获取续写的系统提示词
    ''' </summary>
    Protected Function GetContinuationSystemPrompt() As String
        Return "你是一个专业的写作助手。根据提供的上下文，自然地续写内容。要求：
1. 保持与原文一致的语言风格、语气和术语
2. 内容要连贯自然，不要重复上文已有内容
3. 只输出续写内容，不要添加任何解释、前缀或标记
4. 如果上下文不足，可以合理推断但保持谨慎
5. 续写长度适中，约100-300字，除非用户另有要求"
    End Function

    ''' <summary>
    ''' 构建续写请求的用户提示
    ''' </summary>
    Protected Function BuildContinuationUserPrompt(context As ContinuationContext, Optional style As String = "") As String
        Dim sb As New StringBuilder()

        sb.AppendLine("请根据以下上下文续写内容：")
        sb.AppendLine()
        sb.Append(context.BuildPrompt())

        If Not String.IsNullOrWhiteSpace(style) Then
            sb.AppendLine()
            sb.AppendLine($"【风格要求】{style}")
        End If

        sb.AppendLine()
        sb.AppendLine("请直接输出续写内容，不要添加任何前缀或说明：")

        Return sb.ToString()
    End Function

    ' ========== 续写功能相关方法结束 ==========

    ' ========== 模板渲染功能相关方法 ==========

    ''' <summary>
    ''' 应用模板渲染结果到文档（由子类实现）
    ''' </summary>
    Protected Overridable Sub HandleApplyTemplateContent(jsonDoc As JObject)
        Debug.WriteLine("HandleApplyTemplateContent 被调用（基类默认不执行）")
        GlobalStatusStrip.ShowWarning("当前应用不支持模板渲染功能")
    End Sub

    ''' <summary>
    ''' 调整模板渲染需求（多轮对话）
    ''' </summary>
    Protected Overridable Sub HandleRefineTemplateContent(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim refinement As String = If(jsonDoc("refinement") IsNot Nothing, jsonDoc("refinement").ToString(), String.Empty)

            If String.IsNullOrWhiteSpace(refinement) Then
                GlobalStatusStrip.ShowWarning("请输入调整需求")
                Return
            End If

            ' 构建调整提示
            Dim refinementPrompt As New StringBuilder()
            refinementPrompt.AppendLine("请根据以下要求调整之前生成的模板内容：")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine($"【调整需求】{refinement}")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("请直接输出调整后的内容，不要添加任何解释：")

            ' 保持 responseMode = "template_render"，发送调整请求（不使用历史记录）
            Task.Run(Async Function()
                         Await Send(refinementPrompt.ToString(), GetTemplateRenderSystemPrompt(""), False, "template_render")
                     End Function)

            GlobalStatusStrip.ShowInfo("正在调整模板内容...")
        Catch ex As Exception
            Debug.WriteLine($"HandleRefineTemplateContent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("调整模板内容时出错")
        End Try
    End Sub

    ''' <summary>
    ''' 获取模板渲染的系统提示词
    ''' </summary>
    Protected Function GetTemplateRenderSystemPrompt(templateContext As String) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("你是一个专业的文档内容生成助手。你需要根据用户提供的模板结构（JSON格式）和风格来生成新的内容。")
        sb.AppendLine()
        sb.AppendLine("【重要格式要求】")
        sb.AppendLine("- 严禁使用Markdown代码块格式（禁止使用```符号）")
        sb.AppendLine("- 严禁使用任何Markdown格式标记（如#、**、-、>等）")
        sb.AppendLine("- 直接输出纯文本内容，不要包装在任何代码块中")
        sb.AppendLine("- 不要添加任何前缀、后缀、解释或说明文字")
        sb.AppendLine("- 不要输出JSON格式，直接输出可以插入文档的纯文本")
        sb.AppendLine()
        sb.AppendLine("【模板JSON结构说明】")
        sb.AppendLine("模板以JSON格式提供，包含以下信息：")
        sb.AppendLine("- elements: 文档元素数组，每个元素包含type(类型)、text(文本)、styleName(样式名)、formatting(格式详情)")
        sb.AppendLine("- formatting包含: fontName(字体)、fontSize(字号)、bold(加粗)、italic(斜体)、alignment(对齐)等")
        sb.AppendLine("- 对于PPT模板：slides数组包含每张幻灯片的布局和元素信息")
        sb.AppendLine()
        sb.AppendLine("【内容生成要求】")
        sb.AppendLine("1. 严格遵循模板的层级结构（如：标题、副标题、正文的层次关系）")
        sb.AppendLine("2. 保持与模板一致的语气、术语规范和风格")
        sb.AppendLine("3. 参考模板中的字号来判断内容的重要程度（大字号=标题，小字号=正文）")
        sb.AppendLine("4. 内容要专业、连贯、符合实际使用场景")
        sb.AppendLine("5. 按照模板中元素的顺序来组织输出内容")
        sb.AppendLine("6. 每个段落或幻灯片内容之间用空行分隔")

        If Not String.IsNullOrWhiteSpace(templateContext) Then
            sb.AppendLine()
            sb.AppendLine("【参考模板结构】")
            sb.AppendLine("```json")
            sb.AppendLine(templateContext)
            sb.AppendLine("```")
            sb.AppendLine()
            sb.AppendLine("请根据以上模板结构，按照用户的内容需求生成相应格式的文档内容。直接输出纯文本，不要使用任何Markdown格式。")
        End If

        Return sb.ToString()
    End Function

    ' ========== 模板渲染功能相关方法结束 ==========

    ' ========== 自动补全功能相关方法 ==========

    ''' <summary>
    ''' 处理自动补全请求
    ''' </summary>
    Protected Overridable Async Sub HandleRequestCompletion(jsonDoc As JObject)
        Try
            ' 检查设置是否启用自动补全
            If Not ChatSettings.EnableAutocomplete Then
                Return
            End If

            Dim inputText As String = If(jsonDoc("input")?.ToString(), "")
            Dim timestamp As Long = If(jsonDoc("timestamp")?.Value(Of Long)(), 0)

            If String.IsNullOrWhiteSpace(inputText) OrElse inputText.Length < 2 Then
                Return
            End If

            ' 获取上下文快照
            Dim contextSnapshot = GetContextSnapshot()

            ' 构建补全请求的prompt
            Dim completions = Await RequestCompletionsFromLLM(inputText, contextSnapshot)

            ' 返回结果到前端
            Dim resultJson As New JObject()
            resultJson("completions") = JArray.FromObject(completions)
            resultJson("timestamp") = timestamp

            ExecuteJavaScriptAsyncJS($"showCompletions({resultJson.ToString(Newtonsoft.Json.Formatting.None)});")

        Catch ex As Exception
            Debug.WriteLine($"HandleRequestCompletion 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理补全采纳记录
    ''' </summary>
    Protected Overridable Sub HandleAcceptCompletion(jsonDoc As JObject)
        Try
            Dim inputText As String = If(jsonDoc("input")?.ToString(), "")
            Dim completion As String = If(jsonDoc("completion")?.ToString(), "")
            Dim context As String = If(jsonDoc("context")?.ToString(), "")

            ' 记录补全历史
            RecordCompletionHistory(inputText, completion, context)

        Catch ex As Exception
            Debug.WriteLine($"HandleAcceptCompletion 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取当前Office上下文快照（由子类重写提供具体实现）
    ''' </summary>
    Protected Overridable Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = GetOfficeAppType()
        snapshot("selection") = ""
        Return snapshot
    End Function

    ''' <summary>
    ''' 调用大模型获取补全建议
    ''' </summary>
    Private Async Function RequestCompletionsFromLLM(inputText As String, contextSnapshot As JObject) As Task(Of List(Of String))
        Dim completions As New List(Of String)()

        Try
            ' 获取翻译配置（用于补全）
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
                Return completions
            End If

            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)
            
            Dim modelName = selectedModel.modelName
            Dim apiUrl = cfg.url
            Dim apiKey = cfg.key
            
            ' 检查是否支持FIM模式
            Dim useFimMode = selectedModel.fimSupported AndAlso Not String.IsNullOrEmpty(selectedModel.fimUrl)
            
            If useFimMode Then
                ' 使用FIM API
                completions = Await RequestCompletionsWithFIM(inputText, contextSnapshot, selectedModel, apiKey)
            Else
                ' 使用Chat Completion API
                completions = Await RequestCompletionsWithChat(inputText, contextSnapshot, cfg, selectedModel, apiKey)
            End If

        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsFromLLM 出错: {ex.Message}")
        End Try

        Return completions
    End Function

    ''' <summary>
    ''' 使用FIM (Fill-In-the-Middle) API获取补全
    ''' </summary>
    Private Async Function RequestCompletionsWithFIM(inputText As String, contextSnapshot As JObject, 
                                                      model As ConfigManager.ConfigItemModel, apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            Dim fimUrl = model.fimUrl
            
            ' 构建FIM请求
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("prompt") = inputText
            requestObj("suffix") = "" ' 光标后无内容
            requestObj("max_tokens") = 50
            requestObj("temperature") = 0.3
            requestObj("stream") = False
            
            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)
            
            Using client As New Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(10)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New Net.Http.StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(fimUrl, content)
                response.EnsureSuccessStatusCode()
                
                Dim responseBody = Await response.Content.ReadAsStringAsync()
                Dim jObj = JObject.Parse(responseBody)
                
                ' FIM API返回格式: {"choices": [{"text": "补全内容"}]}
                Dim text = jObj("choices")?(0)?("text")?.ToString()
                If Not String.IsNullOrWhiteSpace(text) Then
                    ' 清理换行和多余空白
                    text = text.Trim().Split({vbCr, vbLf, vbCrLf}, StringSplitOptions.RemoveEmptyEntries)(0)
                    If text.Length <= 50 Then
                        completions.Add(text)
                    End If
                End If
            End Using
            
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsWithFIM 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function

    ''' <summary>
    ''' 使用Chat Completion API获取补全
    ''' </summary>
    Private Async Function RequestCompletionsWithChat(inputText As String, contextSnapshot As JObject,
                                                       cfg As ConfigManager.ConfigItem, model As ConfigManager.ConfigItemModel, 
                                                       apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            Dim apiUrl = cfg.url
            Dim modelName = model.modelName

            ' 获取上下文信息
            Dim appType = If(contextSnapshot("appType")?.ToString(), "Office")
            Dim selectionText = If(contextSnapshot("selection")?.ToString(), "")

            ' 根据Office类型构建场景化系统提示词
            Dim systemPrompt = GetCompletionSystemPrompt(appType)

            ' 构建用户消息
            Dim userContent As New StringBuilder()
            userContent.AppendLine($"当前应用: {appType}")
            userContent.AppendLine($"用户已输入: ""{inputText}""")
            If Not String.IsNullOrWhiteSpace(selectionText) Then
                userContent.AppendLine($"选中内容: ""{selectionText.Substring(0, Math.Min(200, selectionText.Length))}""")
            End If
            
            ' 添加额外上下文信息
            If contextSnapshot("sheetName") IsNot Nothing Then
                userContent.AppendLine($"当前工作表: {contextSnapshot("sheetName")}")
            End If
            If contextSnapshot("slideIndex") IsNot Nothing Then
                userContent.AppendLine($"当前幻灯片: 第{contextSnapshot("slideIndex")}页")
            End If
            
            userContent.AppendLine()
            userContent.AppendLine("请给出补全建议（JSON格式）。")

            ' 构建请求
            Dim requestObj As New JObject()
            requestObj("model") = modelName
            requestObj("stream") = False
            requestObj("temperature") = 0.3

            Dim messages As New JArray()
            messages.Add(New JObject() From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject() From {{"role", "user"}, {"content", userContent.ToString()}})
            requestObj("messages") = messages

            ' 发送请求
            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)

            Using client As New Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(10) ' 补全请求超时短一些
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New Net.Http.StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()

                Dim responseBody = Await response.Content.ReadAsStringAsync()
                
                ' 解析API响应
                Dim jObj As JObject = Nothing
                Try
                    jObj = JObject.Parse(responseBody)
                Catch apiParseEx As Exception
                    Debug.WriteLine($"解析API响应失败: {apiParseEx.Message}")
                    Return completions
                End Try

                Dim msg As String = Nothing
                Try
                    msg = jObj("choices")?(0)?("message")?("content")?.ToString()
                Catch
                    ' 尝试其他格式（例如某些API使用不同的响应结构）
                    msg = jObj("message")?.ToString()
                End Try

                If Not String.IsNullOrEmpty(msg) Then
                    ' 尝试解析JSON响应
                    Try
                        ' 清理可能的markdown代码块标记
                        Dim cleanedMsg = msg.Trim()
                        If cleanedMsg.StartsWith("```") Then
                            ' 去除开头的```json或```
                            Dim firstNewLine = cleanedMsg.IndexOf(vbLf)
                            If firstNewLine > 0 Then
                                cleanedMsg = cleanedMsg.Substring(firstNewLine + 1)
                            End If
                        End If
                        If cleanedMsg.EndsWith("```") Then
                            cleanedMsg = cleanedMsg.Substring(0, cleanedMsg.Length - 3)
                        End If
                        cleanedMsg = cleanedMsg.Trim()

                        ' 尝试找到JSON对象的起始位置
                        Dim jsonStart = cleanedMsg.IndexOf("{")
                        Dim jsonEnd = cleanedMsg.LastIndexOf("}")
                        If jsonStart >= 0 AndAlso jsonEnd > jsonStart Then
                            cleanedMsg = cleanedMsg.Substring(jsonStart, jsonEnd - jsonStart + 1)
                        End If

                        Dim resultObj = JObject.Parse(cleanedMsg)
                        Dim completionsArray = resultObj("completions")
                        If completionsArray IsNot Nothing Then
                            For Each item In completionsArray
                                Dim c = item.ToString().Trim()
                                If Not String.IsNullOrWhiteSpace(c) Then
                                    completions.Add(c)
                                End If
                            Next
                        End If
                    Catch parseEx As Exception
                        Debug.WriteLine($"解析补全JSON失败: {parseEx.Message}, 原始内容: {msg}")
                        ' 如果不是有效JSON，尝试直接使用返回内容
                        If Not String.IsNullOrWhiteSpace(msg) AndAlso msg.Length < 50 Then
                            completions.Add(msg.Trim())
                        End If
                    End Try
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsFromLLM 出错: {ex.Message}")
        End Try

        Return completions
    End Function

    ''' <summary>
    ''' 记录补全历史
    ''' </summary>
    Private Sub RecordCompletionHistory(inputText As String, completion As String, context As String)
        Try
            Dim historyPath = IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                ConfigSettings.OfficeAiAppDataFolder,
                "autocomplete_history.json")

            Dim history As JObject
            If IO.File.Exists(historyPath) Then
                Dim json = IO.File.ReadAllText(historyPath)
                history = JObject.Parse(json)
            Else
                history = New JObject()
                history("version") = 1
                history("history") = New JArray()
            End If

            Dim historyArray = CType(history("history"), JArray)

            ' 查找是否已有相同的记录
            Dim existingItem = historyArray.FirstOrDefault(Function(item)
                                                               Return item("input")?.ToString() = inputText AndAlso
                                                                      item("completion")?.ToString() = completion
                                                           End Function)

            If existingItem IsNot Nothing Then
                ' 更新计数和时间
                existingItem("count") = existingItem("count").Value(Of Integer)() + 1
                existingItem("lastUsed") = DateTime.UtcNow.ToString("o")
            Else
                ' 添加新记录
                Dim newItem As New JObject()
                newItem("input") = inputText
                newItem("completion") = completion
                newItem("context") = context
                newItem("count") = 1
                newItem("lastUsed") = DateTime.UtcNow.ToString("o")
                historyArray.Add(newItem)

                ' 限制历史记录数量（最多保留100条）
                While historyArray.Count > 100
                    historyArray.RemoveAt(0)
                End While
            End If

            ' 保存
            Dim dir = IO.Path.GetDirectoryName(historyPath)
            If Not IO.Directory.Exists(dir) Then
                IO.Directory.CreateDirectory(dir)
            End If
            IO.File.WriteAllText(historyPath, history.ToString(Newtonsoft.Json.Formatting.Indented))

        Catch ex As Exception
            Debug.WriteLine($"RecordCompletionHistory 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 根据Office应用类型获取场景化的补全系统提示词
    ''' </summary>
    Private Function GetCompletionSystemPrompt(appType As String) As String
        Dim baseRules = "
规则：
1. 只返回补全的剩余部分，不要重复用户已输入的内容
2. 返回JSON格式: {""completions"": [""补全1"", ""补全2"", ""补全3""]}
3. 最多返回3个候选
4. 补全应简洁，通常不超过20个字"

        Select Case appType.ToLower()
            Case "excel"
                Return $"你是Excel AI助手的输入补全引擎。根据用户当前输入和Excel上下文，预测用户想要的操作。

常见Excel场景补全示例：
- ""帮我"" → ""计算这列的总和"", ""筛选重复数据"", ""生成数据透视表""
- ""把"" → ""选中区域转换为表格"", ""这列数据去重"", ""A列和B列合并""
- ""统计"" → ""每个类别的数量"", ""销售额的平均值"", ""各月份的增长率""
- ""公式"" → ""计算两列的差值"", ""查找匹配的数据"", ""条件求和""
- ""格式"" → ""设置为货币格式"", ""添加条件格式"", ""调整列宽""
- ""图表"" → ""创建柱状图"", ""生成趋势线"", ""添加数据标签""
{baseRules}"

            Case "word"
                Return $"你是Word AI助手的输入补全引擎。根据用户当前输入和Word上下文，预测用户想要的操作。

常见Word场景补全示例：
- ""帮我"" → ""润色这段文字"", ""翻译选中内容"", ""生成文章大纲""
- ""把"" → ""这段改成正式语气"", ""标题设为一级标题"", ""段落缩进调整""
- ""总结"" → ""这篇文章的要点"", ""会议纪要"", ""核心观点""
- ""扩写"" → ""这个段落"", ""详细说明这个观点"", ""增加案例论证""
- ""格式"" → ""统一段落间距"", ""添加页眉页脚"", ""设置目录样式""
- ""检查"" → ""语法错误"", ""错别字"", ""标点符号""
{baseRules}"

            Case "powerpoint"
                Return $"你是PowerPoint AI助手的输入补全引擎。根据用户当前输入和PPT上下文，预测用户想要的操作。

常见PPT场景补全示例：
- ""帮我"" → ""美化这张幻灯片"", ""生成演讲稿"", ""添加过渡动画""
- ""把"" → ""文字转换为SmartArt"", ""图片裁剪为圆形"", ""背景改为渐变色""
- ""生成"" → ""项目汇报PPT"", ""产品介绍页"", ""团队介绍页""
- ""添加"" → ""图表展示数据"", ""时间线"", ""流程图""
- ""设计"" → ""统一字体样式"", ""配色方案"", ""母版布局""
- ""总结"" → ""演示要点"", ""关键数据"", ""结论页内容""
{baseRules}"

            Case Else
                Return $"你是Office AI助手的输入补全引擎。根据用户当前输入和Office上下文，预测用户想要输入的内容。
{baseRules}
5. 考虑Office上下文（选中内容、文档类型）"
        End Select
    End Function

    ' ========== 自动补全功能相关方法结束 ==========

    ' 新增：处理用户接受答案
    Protected Sub HandleAcceptAnswer(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)

            ' 简单记录与提示（可扩展为在历史中设置标记或持久化）
            Debug.WriteLine($"用户接受回答: UUID={uuid}")
            Debug.WriteLine(content)

            ' 你也可以向前端反馈已接受（可选）
            GlobalStatusStrip.ShowInfo("用户已接受 AI 回答")
        Catch ex As Exception
            Debug.WriteLine($"HandleAcceptAnswer 出错: {ex.Message}")
        End Try
    End Sub

    ' 新增：处理用户拒绝答案并发起改进
    Protected Sub HandleRejectAnswer(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim rejectedContent As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim reason As String = If(jsonDoc("reason") IsNot Nothing, jsonDoc("reason").ToString(), String.Empty)

            Debug.WriteLine($"用户拒绝回答: UUID={uuid}; reason={reason}")


            ' 构建用于改进的大模型提示（包含用户理由）
            Dim refinementPrompt As New StringBuilder()
            refinementPrompt.AppendLine("用户标记之前的回答为不接受，请基于当前会话历史与以下被拒绝的回答进行改进：")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("【用户改进诉求】")
            If Not String.IsNullOrWhiteSpace(reason) Then
                refinementPrompt.AppendLine(reason)
            Else
                refinementPrompt.AppendLine("[无具体改进诉求，用户仅标记为不接受]")
            End If
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("请按以下格式返回：")
            refinementPrompt.AppendLine("1) 改进点（1-3 行），说明要如何修正；")
            refinementPrompt.AppendLine("2) Plan：简短列出修正步骤（要点式，最多6条）；")
            refinementPrompt.AppendLine("3) Answer：给出修正后的、尽可能准确的答案（使用 Markdown，必要时给出示例/代码）；")
            refinementPrompt.AppendLine("4) Clarifying Questions：如需更多信息，请在最后以简短问题列出并暂停执行；")
            refinementPrompt.AppendLine()
            refinementPrompt.AppendLine("[注意]：回答要简洁、可验证，优先给出可直接执行的结论与验证方法，不要重复冗长的背景说明。")

            ' 管理历史大小，保证不会无限增长
            ManageHistoryMessageSize()

            ' 将该改进请求当作新的用户问题发起（会走你已有的 SendChatMessage 流程）
            SendChatMessage(refinementPrompt.ToString())

            GlobalStatusStrip.ShowInfo("已触发改进请求，正在向模型发起新一轮改进")
        Catch ex As Exception
            Debug.WriteLine($"HandleRejectAnswer 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("触发改进请求时出错")
        End Try
    End Sub

    Private Sub ClearChatContext()
        systemHistoryMessageData.Clear()
        Debug.WriteLine("已清空聊天记忆（上下文）")
    End Sub

    ' 处理获取MCP连接列表请求 - 委托给 McpService
    Protected Sub HandleGetMcpConnections()
        McpService.GetMcpConnections()
    End Sub

    ' 处理保存MCP设置请求 - 委托给 McpService
    Protected Sub HandleSaveMcpSettings(jsonDoc As JObject)
        McpService.SaveMcpSettings(jsonDoc)
    End Sub

    ' MCP初始化方法 - 委托给 McpService
    Protected Sub InitializeMcpSettings()
        McpService.InitializeMcpSettings()
    End Sub

    ' 处理获取历史文件列表请求 - 委托给 HistoryService
    Protected Sub HandleGetHistoryFiles()
        HistoryService.GetHistoryFiles()
    End Sub

    ' 处理打开历史文件请求 - 委托给 HistoryService
    Protected Sub HandleOpenHistoryFile(jsonDoc As JObject)
        HistoryService.OpenHistoryFile(jsonDoc)
    End Sub

    Protected Overridable Sub HandleCheckedChange(jsonDoc As JObject)
        Dim prop As String = jsonDoc("property").ToString()
        Dim isChecked As Boolean = Boolean.Parse(jsonDoc("isChecked").ToString())
        If prop = "selectedCell" Then
            selectedCellChecked = isChecked
        End If
    End Sub

    Protected Overridable Sub HandleSaveSettings(jsonDoc As JObject)
        topicRandomness = jsonDoc("topicRandomness")
        contextLimit = jsonDoc("contextLimit")
        selectedCellChecked = jsonDoc("selectedCell")
        settingsScrollChecked = jsonDoc("settingsScroll")
        Dim chatMode As String = jsonDoc("chatMode")
        Dim executeCodePreview As Boolean = jsonDoc("executeCodePreview")
        Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
        Dim autocompleteShortcut As String = If(jsonDoc("autocompleteShortcut")?.Value(Of String)(), "Ctrl+.")
        Dim chatSettings As New ChatSettings(GetApplication())
        ' 保存设置到配置文件
        chatSettings.SaveSettings(topicRandomness, contextLimit, selectedCellChecked,
                                  settingsScrollChecked, executeCodePreview, chatMode,
                                  enableAutocomplete, 800, autocompleteShortcut)
    End Sub

    Public Class SendMessageReferenceContentItem
        Public Property id As String
        Public Property sheetName As String
        Public Property address As String
    End Class

    ' FileContentResult 类已移至 Controls/Models/HistoryMessage.vb


    ' 添加存储第一个问题的变量
    Protected firstQuestion As String = String.Empty
    Protected isFirstMessage As Boolean = True
    Private _chatHtmlFilePath As String = String.Empty ' 缓存文件路径



    ' 在 HandleSendMessage 方法中添加文件内容解析逻辑
    Protected Overridable Sub HandleSendMessage(jsonDoc As JObject)
        Dim messageValue As JToken = jsonDoc("value")
        Dim question As String
        Dim filePaths As List(Of String) = New List(Of String)()
        Dim selectedContents As List(Of SendMessageReferenceContentItem) = New List(Of SendMessageReferenceContentItem)()

        If messageValue.Type = JTokenType.Object Then
            ' New format with text, potentially filePaths, and selectedContent
            question = messageValue("text")?.ToString()

            If messageValue("filePaths") IsNot Nothing AndAlso messageValue("filePaths").Type = JTokenType.Array Then
                filePaths = messageValue("filePaths").ToObject(Of List(Of String))()
            End If

            ' 解析 selectedContent
            If messageValue("selectedContent") IsNot Nothing AndAlso messageValue("selectedContent").Type = JTokenType.Array Then
                Try
                    selectedContents = messageValue("selectedContent").ToObject(Of List(Of SendMessageReferenceContentItem))()
                Catch ex As Exception
                    Debug.WriteLine($"Error deserializing selectedContent: {ex.Message}")
                End Try
            End If
        Else
            Debug.WriteLine("HandleSendMessage: Invalid message format for 'value'.")
            Return
        End If

        If String.IsNullOrEmpty(question) AndAlso
       (filePaths Is Nothing OrElse filePaths.Count = 0) AndAlso
       (selectedContents Is Nothing OrElse selectedContents.Count = 0) Then
            Debug.WriteLine("HandleSendMessage: Empty question, no files, and no selected content.")
            Return ' Nothing to send
        End If

        ' 保存原始用户输入（用于意图识别）
        Dim originalQuestion As String = question

        ' 保存第一个问题（仅保存一次）
        If isFirstMessage AndAlso Not String.IsNullOrEmpty(question) Then
            firstQuestion = question
            isFirstMessage = False
            ' 清空缓存的文件路径，强制重新生成
            _chatHtmlFilePath = String.Empty
            Debug.WriteLine($"保存第一个问题: {firstQuestion}")
            Debug.WriteLine($"将生成文件路径: {ChatHtmlFilePath}")
        End If

        ' --- 处理选中的内容 ---
        question = AppendCurrentSelectedContent("--- 我此次的问题：" & question & " ---")

        ' 检查是否有文件需要解析
        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
            ' 异步处理文件解析，避免卡死UI
            HandleSendMessageWithFilesAsync(question, originalQuestion, filePaths, selectedContents, messageValue)
        Else
            ' 没有文件，直接处理消息
            HandleSendMessageCore(question, originalQuestion, filePaths, selectedContents, messageValue, "")
        End If
    End Sub

    ''' <summary>
    ''' 异步解析文件并发送消息
    ''' </summary>
    Private Sub HandleSendMessageWithFilesAsync(question As String, originalQuestion As String, 
                                                 filePaths As List(Of String), 
                                                 selectedContents As List(Of SendMessageReferenceContentItem),
                                                 messageValue As JToken)
        ' 显示进度提示
        GlobalStatusStrip.ShowInfo($"正在解析 {filePaths.Count} 个文件...")
        ExecuteJavaScriptAsyncJS("showFileParsingProgress(true)")

        Task.Run(Sub()
            Try
                Dim fileContentBuilder As New StringBuilder()
                Dim parsedFiles As New List(Of FileContentResult)()
                Dim totalFiles = filePaths.Count
                Dim processedFiles = 0

                fileContentBuilder.AppendLine(vbCrLf & "--- 以下是用户引用的其他文件内容 ---")

                ' 获取当前工作目录（需要在主线程调用）
                Dim currentWorkingDir As String = ""
                Me.Invoke(Sub()
                    currentWorkingDir = GetCurrentWorkingDirectory()
                End Sub)

                If String.IsNullOrEmpty(currentWorkingDir) Then
                    Me.Invoke(Sub()
                        GlobalStatusStrip.ShowWarning("请保存当前文件，并且把引用文件和当前文件放在同一目录下后重试")
                        ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")
                    End Sub)
                    Return
                End If

                For Each filePath As String In filePaths
                    Try
                        processedFiles += 1
                        
                        ' 更新进度
                        Me.Invoke(Sub()
                            GlobalStatusStrip.ShowInfo($"正在解析文件 ({processedFiles}/{totalFiles}): {Path.GetFileName(filePath)}")
                            ExecuteJavaScriptAsyncJS($"updateFileParsingProgress({processedFiles}, {totalFiles}, '{EscapeJsString(Path.GetFileName(filePath))}')")
                        End Sub)

                        ' 检查文件是否为绝对路径
                        Dim fullFilePath As String = filePath

                        ' 如果不是绝对路径，则尝试在当前工作目录下查找
                        If Not Path.IsPathRooted(filePath) AndAlso Not String.IsNullOrEmpty(currentWorkingDir) Then
                            fullFilePath = Path.Combine(currentWorkingDir, Path.GetFileName(filePath))
                            Debug.WriteLine($"尝试在工作目录查找文件: {fullFilePath}")
                        End If

                        If File.Exists(fullFilePath) Then
                            ' 根据文件扩展名选择合适的解析方法
                            Dim fileExtension As String = Path.GetExtension(fullFilePath).ToLower()
                            Dim fileContentResult As FileContentResult = Nothing

                            Select Case fileExtension
                                Case ".xlsx", ".xls", ".xlsm", ".xlsb"
                                    ' Excel文件解析需要在主线程
                                    Me.Invoke(Sub()
                                        fileContentResult = ParseFile(fullFilePath)
                                    End Sub)
                                Case ".docx", ".doc"
                                    Me.Invoke(Sub()
                                        fileContentResult = ParseFile(fullFilePath)
                                    End Sub)
                                Case ".csv", ".txt"
                                    fileContentResult = _fileParserService.ParseTextFile(fullFilePath)
                                Case Else
                                    fileContentResult = New FileContentResult With {
                                        .FileName = Path.GetFileName(fullFilePath),
                                        .FileType = "Unknown",
                                        .ParsedContent = $"[不支持的文件类型: {fileExtension}]"
                                    }
                            End Select

                            If fileContentResult IsNot Nothing Then
                                parsedFiles.Add(fileContentResult)
                                fileContentBuilder.AppendLine($"文件名: {fileContentResult.FileName}")
                                fileContentBuilder.AppendLine($"文件内容:")
                                fileContentBuilder.AppendLine(fileContentResult.ParsedContent)
                                fileContentBuilder.AppendLine("---")
                            End If
                        Else
                            fileContentBuilder.AppendLine($"文件 '{Path.GetFileName(filePath)}' 未找到或路径无效")
                            Debug.WriteLine($"文件未找到: {fullFilePath}")
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                        fileContentBuilder.AppendLine($"处理文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}")
                        fileContentBuilder.AppendLine("---")
                    End Try
                Next

                fileContentBuilder.AppendLine("--- 文件内容结束 ---" & vbCrLf)

                ' 文件解析完成，在主线程继续处理消息
                Me.Invoke(Sub()
                    GlobalStatusStrip.ShowInfo($"文件解析完成，共解析 {parsedFiles.Count} 个文件")
                    ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")
                    
                    Dim questionWithFiles = question & " 用户提问结束，后续引用的文件都在同一目录下所以可以放心读取。 ---"
                    HandleSendMessageCore(questionWithFiles, originalQuestion, filePaths, selectedContents, messageValue, fileContentBuilder.ToString())
                End Sub)

            Catch ex As Exception
                Debug.WriteLine($"HandleSendMessageWithFilesAsync 出错: {ex.Message}")
                Me.Invoke(Sub()
                    GlobalStatusStrip.ShowWarning($"文件解析失败: {ex.Message}")
                    ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")
                End Sub)
            End Try
        End Sub)
    End Sub

    ''' <summary>
    ''' 转义JS字符串中的特殊字符
    ''' </summary>
    Private Function EscapeJsString(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Return s.Replace("\", "\\").Replace("'", "\'").Replace("""", "\""").Replace(vbCrLf, "\n").Replace(vbCr, "\n").Replace(vbLf, "\n")
    End Function

    ''' <summary>
    ''' 处理消息发送的核心逻辑（文件解析完成后调用）
    ''' </summary>
    Private Sub HandleSendMessageCore(question As String, originalQuestion As String,
                                       filePaths As List(Of String),
                                       selectedContents As List(Of SendMessageReferenceContentItem),
                                       messageValue As JToken,
                                       fileContent As String)
        
        ' 构建最终发送给 LLM 的消息
        Dim finalMessageToLLM As String = question

        ' 然后添加文件内容（如果有）
        If Not String.IsNullOrEmpty(fileContent) Then
            finalMessageToLLM &= fileContent
        End If

        stopReaderStream = False ' Reset stop flag before sending new message

        ' 检查是否为模板渲染模式
        Dim responseMode As String = If(messageValue("responseMode")?.ToString(), "")
        Dim templateContext As String = If(messageValue("templateContext")?.ToString(), "")

        If responseMode = "template_render" AndAlso Not String.IsNullOrWhiteSpace(templateContext) Then
            ' 模板渲染模式：使用专用systemPrompt，不使用历史记录
            Dim templateSystemPrompt = GetTemplateRenderSystemPrompt(templateContext)
            Task.Run(Async Function()
                         Await Send(finalMessageToLLM, templateSystemPrompt, False, "template_render")
                     End Function)
        Else
            ' 获取当前聊天模式
            Dim currentChatMode As String = ChatSettings.chatMode

            ' 普通消息模式：先检查是否为追问，再决定是否进行意图识别
            Task.Run(Async Function()
                Try
                    ' 检查是否有引用内容（文件或选中内容）
                    Dim hasReferences As Boolean = (filePaths IsNot Nothing AndAlso filePaths.Count > 0) OrElse
                                                    (selectedContents IsNot Nothing AndAlso selectedContents.Count > 0)

                    ' 检查是否有历史对话记录，如果有则判断新问题是否为追问
                    Dim isFollowUp As Boolean = False
                    If systemHistoryMessageData.Count >= 2 AndAlso Not String.IsNullOrWhiteSpace(originalQuestion) Then
                        ' 有历史记录，检查新问题是否与之前对话相关
                        isFollowUp = Await IntentService.IsFollowUpQuestionAsync(originalQuestion, systemHistoryMessageData)
                        Debug.WriteLine($"追问检查结果: isFollowUp={isFollowUp}")
                    End If

                    ' 如果是追问（与之前对话相关），直接发送给大模型，不弹出意图确认框
                    If isFollowUp Then
                        Debug.WriteLine($"检测到追问，直接发送给大模型处理")
                        SendChatMessage(finalMessageToLLM)
                        Return
                    End If

                    ' 不是追问，进行意图识别
                    ' 获取上下文快照
                    Dim contextSnapshot = GetContextSnapshot()

                    ' 使用异步方法进行意图识别（调用大模型）
                    CurrentIntentResult = Await IntentService.IdentifyIntentAsync(originalQuestion, contextSnapshot)
                    CurrentIntentResult.OriginalInput = originalQuestion

                    ' 如果LLM没有生成描述，使用默认生成
                    If String.IsNullOrEmpty(CurrentIntentResult.UserFriendlyDescription) Then
                        IntentService.GenerateUserFriendlyDescription(CurrentIntentResult)
                    End If
                    IntentService.BuildExecutionPlanPreview(CurrentIntentResult)

                    ' 决定是否需要询问用户确认
                    Dim needConfirmation As Boolean = False
                    Dim autoConfirmCountdown As Boolean = False  ' Agent模式下倒计时后自动确认

                    ' 情况1：用户只引用了内容但没有输入问题
                    If hasReferences AndAlso String.IsNullOrWhiteSpace(originalQuestion) Then
                        CurrentIntentResult.UserFriendlyDescription = "您引用了内容，请问您想要做什么？"
                        needConfirmation = True
                    ' 情况2：置信度太低（<0.4），让大模型来询问用户澄清
                    ElseIf CurrentIntentResult.Confidence < 0.4 Then
                        ' 不弹出意图预览卡片，直接发送给大模型，由大模型来询问用户
                        needConfirmation = False
                    ' 情况3：Agent模式下也需要确认，但会倒计时自动执行
                    ElseIf currentChatMode = "agent" Then
                        needConfirmation = True
                        autoConfirmCountdown = True  ' Agent模式：倒计时后自动确认
                    ' 情况4：普通模式下，置信度较高时也需要确认（仅第一次对话）
                    ElseIf CurrentIntentResult.Confidence >= 0.4 AndAlso systemHistoryMessageData.Count < 2 Then
                        needConfirmation = True
                        autoConfirmCountdown = False  ' Chat模式：不自动确认
                    Else
                        needConfirmation = False ' 有历史记录时默认不弹出确认框
                    End If

                    If needConfirmation Then
                        ' 需要用户确认，保存待发送的消息
                        _pendingIntentMessage = finalMessageToLLM
                        _pendingIntentResult = CurrentIntentResult
                        _pendingFilePaths = filePaths

                        ' 构建意图预览数据并发送给前端
                        Dim clarification = IntentService.GenerateIntentClarification(originalQuestion, contextSnapshot)
                        
                        ' 使用LLM生成的描述
                        If Not String.IsNullOrEmpty(CurrentIntentResult.UserFriendlyDescription) Then
                            clarification.Description = CurrentIntentResult.UserFriendlyDescription
                        End If
                        
                        Dim previewJson = IntentService.IntentClarificationToJson(clarification)
                        
                        ' 添加倒计时相关参数
                        previewJson("autoConfirm") = autoConfirmCountdown
                        previewJson("countdownSeconds") = If(autoConfirmCountdown, 5, 10)  ' Agent模式5秒，Chat模式10秒
                        
                        ' 通知前端显示意图预览卡片（带倒计时）
                        ExecuteJavaScriptAsyncJS($"showIntentPreview({previewJson.ToString(Formatting.None)})")
                        Debug.WriteLine($"显示意图预览（需确认）: {CurrentIntentResult.UserFriendlyDescription}, 自动确认: {autoConfirmCountdown}")
                    Else
                        ' 不需要确认，直接发送
                        Debug.WriteLine($"直接发送消息（置信度:{CurrentIntentResult.Confidence:F2}）")
                        SendChatMessageWithIntent(finalMessageToLLM, CurrentIntentResult)
                    End If

                Catch ex As Exception
                    Debug.WriteLine($"意图识别失败，直接发送: {ex.Message}")
                    ' 回退到直接发送模式
                    SendChatMessage(finalMessageToLLM)
                End Try
            End Function)
        End If
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息
    ''' </summary>
    ''' <param name="message">消息内容</param>
    ''' <param name="intent">意图识别结果</param>
    Protected Overridable Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        ' 默认实现：如果有有效意图，使用优化的systemPrompt
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"使用意图优化提示词: {intent.IntentType}")
            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "")
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

    ''' <summary>
    ''' 处理用户确认意图
    ''' </summary>
    Protected Sub HandleConfirmIntent(jsonDoc As JObject)
        Try
            If String.IsNullOrEmpty(_pendingIntentMessage) Then
                Debug.WriteLine("HandleConfirmIntent: 没有待确认的消息")
                Return
            End If

            Debug.WriteLine("用户确认意图，开始发送消息")

            ' 显示意图类型指示器
            If _pendingIntentResult IsNot Nothing AndAlso _pendingIntentResult.Confidence > 0.2 Then
                ExecuteJavaScriptAsyncJS($"showDetectedIntent('{_pendingIntentResult.IntentType}')")
            End If

            ' 使用保存的意图结果发送消息
            SendChatMessageWithIntent(_pendingIntentMessage, _pendingIntentResult)

            ' 清空待确认状态
            _pendingIntentMessage = Nothing
            _pendingIntentResult = Nothing
            _pendingFilePaths = Nothing

            GlobalStatusStrip.ShowInfo("已确认意图，正在处理...")
        Catch ex As Exception
            Debug.WriteLine($"HandleConfirmIntent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("确认意图时出错")
        End Try
    End Sub

    ''' <summary>
    ''' 处理用户取消意图
    ''' </summary>
    Protected Sub HandleCancelIntent()
        Try
            Debug.WriteLine("用户取消意图")

            ' 清空待确认状态
            _pendingIntentMessage = Nothing
            _pendingIntentResult = Nothing
            _pendingFilePaths = Nothing

            ' 恢复发送按钮状态
            ExecuteJavaScriptAsyncJS("changeSendButton()")

            GlobalStatusStrip.ShowInfo("已取消")
        Catch ex As Exception
            Debug.WriteLine($"HandleCancelIntent 出错: {ex.Message}")
        End Try
    End Sub

#Region "Ralph Loop 循环功能"

    ''' <summary>
    ''' 启动Ralph Loop - 用户输入目标后调用
    ''' </summary>
    Public Async Function StartRalphLoop(userGoal As String) As Task
        Try
            Debug.WriteLine($"[RalphLoop] 启动循环，目标: {userGoal}")
            
            ' 获取应用类型
            Dim appType = GetApplicationType()
            
            ' 创建新的循环会话
            Dim loopSession = Await _ralphLoopController.StartNewLoop(userGoal, appType)
            
            ' 显示规划中状态
            Dim loopDataJson = $"{{""goal"":""{EscapeJavaScriptString(userGoal)}"",""steps"":[],""status"":""planning""}}"
            ExecuteJavaScriptAsyncJS($"showLoopPlanCard({loopDataJson})")
            
            GlobalStatusStrip.ShowInfo("正在规划任务...")
            
            ' 调用AI进行任务规划
            Dim planningPrompt = _ralphLoopController.GetPlanningPrompt(userGoal)
            
            ' 发送规划请求（使用特殊模式标记）
            _isRalphLoopPlanning = True
            Await Send(planningPrompt, "", False, "")
            
        Catch ex As Exception
            Debug.WriteLine($"[RalphLoop] 启动失败: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"启动循环失败: {ex.Message}")
        End Try
    End Function
    
    ' Ralph Loop 规划模式标记
    Private _isRalphLoopPlanning As Boolean = False
    
    ''' <summary>
    ''' 处理前端startLoop消息
    ''' </summary>
    Protected Sub HandleStartLoop(jsonDoc As JObject)
        Try
            Dim goal = jsonDoc("goal")?.ToString()
            If Not String.IsNullOrEmpty(goal) Then
                StartRalphLoop(goal)
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleStartLoop 出错: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 处理继续执行循环
    ''' </summary>
    Protected Async Sub HandleContinueLoop()
        Try
            Debug.WriteLine("[RalphLoop] 用户点击继续执行")
            
            Dim nextStep = _ralphLoopController.ExecuteNextStep()
            If nextStep Is Nothing Then
                Debug.WriteLine("[RalphLoop] 没有更多步骤")
                ExecuteJavaScriptAsyncJS("updateLoopStatus('completed')")
                GlobalStatusStrip.ShowInfo("所有步骤已完成")
                Return
            End If
            
            ' 更新UI显示当前步骤
            ExecuteJavaScriptAsyncJS($"updateLoopStep({nextStep.StepNumber - 1}, 'running')")
            ExecuteJavaScriptAsyncJS("updateLoopStatus('running')")
            GlobalStatusStrip.ShowInfo($"正在执行步骤 {nextStep.StepNumber}: {nextStep.Description}")
            
            ' 执行当前步骤
            _currentRalphLoopStep = nextStep
            Await Send(nextStep.Description, "", True, "")
            
        Catch ex As Exception
            Debug.WriteLine($"HandleContinueLoop 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"执行步骤失败: {ex.Message}")
        End Try
    End Sub
    
    ' 当前执行的步骤
    Private _currentRalphLoopStep As RalphLoopStep = Nothing
    
    ''' <summary>
    ''' 处理取消循环
    ''' </summary>
    Protected Sub HandleCancelLoop()
        Try
            Debug.WriteLine("[RalphLoop] 用户取消循环")
            
            _ralphLoopController.ClearAndEndLoop()
            _isRalphLoopPlanning = False
            _currentRalphLoopStep = Nothing
            
            ExecuteJavaScriptAsyncJS("hideLoopPlanCard()")
            GlobalStatusStrip.ShowInfo("已取消循环任务")
            
        Catch ex As Exception
            Debug.WriteLine($"HandleCancelLoop 出错: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 在流完成后检查是否需要处理Ralph Loop
    ''' </summary>
    Protected Sub CheckRalphLoopCompletion(responseContent As String)
        Try
            ' 检查是否在规划模式
            If _isRalphLoopPlanning Then
                _isRalphLoopPlanning = False
                
                ' 解析规划结果
                If _ralphLoopController.ParsePlanningResult(responseContent) Then
                    Dim loopSession = _ralphLoopController.GetActiveLoop()
                    If loopSession IsNot Nothing Then
                        ' 更新前端显示规划结果
                        Dim stepsJson = BuildStepsJson(loopSession.Steps)
                        Dim loopDataJson = $"{{""goal"":""{EscapeJavaScriptString(loopSession.OriginalGoal)}"",""steps"":{stepsJson},""status"":""ready""}}"
                        ExecuteJavaScriptAsyncJS($"showLoopPlanCard({loopDataJson})")
                        GlobalStatusStrip.ShowInfo("规划完成，点击[继续执行]开始")
                    End If
                Else
                    GlobalStatusStrip.ShowWarning("规划结果解析失败")
                    ExecuteJavaScriptAsyncJS("hideLoopPlanCard()")
                End If
                Return
            End If
            
            ' 检查是否在执行步骤
            If _currentRalphLoopStep IsNot Nothing Then
                Dim stepNum = _currentRalphLoopStep.StepNumber
                _ralphLoopController.CompleteCurrentStep(responseContent, True)
                _currentRalphLoopStep = Nothing
                
                ' 更新UI
                ExecuteJavaScriptAsyncJS($"updateLoopStep({stepNum - 1}, 'completed')")
                
                ' 检查是否还有更多步骤
                Dim loopSession = _ralphLoopController.GetActiveLoop()
                If loopSession IsNot Nothing Then
                    If loopSession.Status = RalphLoopStatus.Paused Then
                        ExecuteJavaScriptAsyncJS("updateLoopStatus('paused')")
                        GlobalStatusStrip.ShowInfo($"步骤 {stepNum} 完成，点击继续执行下一步")
                    ElseIf loopSession.Status = RalphLoopStatus.Completed Then
                        ExecuteJavaScriptAsyncJS("updateLoopStatus('completed')")
                        GlobalStatusStrip.ShowInfo("所有步骤已完成！")
                    End If
                End If
            End If
            
        Catch ex As Exception
            Debug.WriteLine($"CheckRalphLoopCompletion 出错: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 构建步骤JSON
    ''' </summary>
    Private Function BuildStepsJson(steps As List(Of RalphLoopStep)) As String
        Dim sb As New StringBuilder()
        sb.Append("[")
        For i = 0 To steps.Count - 1
            If i > 0 Then sb.Append(",")
            Dim s = steps(i)
            Dim statusStr = s.Status.ToString().ToLower()
            sb.Append($"{{""description"":""{EscapeJavaScriptString(s.Description)}"",""status"":""{statusStr}""}}")
        Next
        sb.Append("]")
        Return sb.ToString()
    End Function
    
    ''' <summary>
    ''' 获取应用类型（子类重写）
    ''' </summary>
    Protected Overridable Function GetApplicationType() As String
        Return "Office"
    End Function

#End Region

#Region "Ralph Agent 智能助手"

    ''' <summary>
    ''' 初始化Agent控制器
    ''' </summary>
    Protected Sub InitializeAgentController()
        If _ralphAgentController Is Nothing Then
            _ralphAgentController = New RalphAgentController()
            
            ' 设置回调
            _ralphAgentController.OnStatusChanged = Sub(status)
                ExecuteJavaScriptAsyncJS($"updateAgentStatus('{_ralphAgentController.GetCurrentSession()?.Id}', 'running', '{EscapeJavaScriptString(status)}')")
            End Sub
            
            _ralphAgentController.OnStepStarted = Sub(stepIndex, desc)
                ExecuteJavaScriptAsyncJS($"updateAgentStep('{_ralphAgentController.GetCurrentSession()?.Id}', {stepIndex}, 'running', '')")
            End Sub
            
            _ralphAgentController.OnStepCompleted = Sub(stepIndex, success, msg)
                Dim status = If(success, "completed", "failed")
                ExecuteJavaScriptAsyncJS($"updateAgentStep('{_ralphAgentController.GetCurrentSession()?.Id}', {stepIndex}, '{status}', '{EscapeJavaScriptString(msg)}')")
            End Sub
            
            _ralphAgentController.OnAgentCompleted = Sub(success)
                ExecuteJavaScriptAsyncJS($"completeAgent('{_ralphAgentController.GetCurrentSession()?.Id}', {success.ToString().ToLower()}, '')")
            End Sub
            
            ' 设置AI请求委托
            _ralphAgentController.SendAIRequest = Async Function(prompt, sysPrompt)
                Return Await SendAndGetResponse(prompt, sysPrompt)
            End Function
            
            ' 设置代码执行委托
            _ralphAgentController.ExecuteCode = Sub(code, lang, preview)
                _codeExecutionService?.ExecuteCode(code, lang, preview)
            End Sub
        End If
    End Sub

    ''' <summary>
    ''' 处理启动Agent请求
    ''' </summary>
    Protected Async Sub HandleStartAgent(jsonDoc As JObject)
        Try
            Dim request = jsonDoc("request")?.ToString()
            If String.IsNullOrEmpty(request) Then
                GlobalStatusStrip.ShowWarning("请输入任务描述")
                Return
            End If
            
            Debug.WriteLine($"[RalphAgent] 启动Agent，需求: {request}")
            
            ' 初始化控制器
            InitializeAgentController()
            
            ' 获取当前Office内容
            Dim appType = GetApplicationType()
            Dim currentContent = GetCurrentOfficeContent()
            
            GlobalStatusStrip.ShowInfo("正在分析您的需求...")
            
            ' 启动Agent规划
            Dim success = Await _ralphAgentController.StartAgent(request, appType, currentContent)
            
            If success Then
                ' 显示规划卡片
                Dim session = _ralphAgentController.GetCurrentSession()
                If session IsNot Nothing Then
                    ShowAgentPlanCard(session)
                End If
            Else
                GlobalStatusStrip.ShowWarning("无法分析您的需求，请重试")
            End If
            
        Catch ex As Exception
            Debug.WriteLine($"HandleStartAgent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"启动Agent失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 显示Agent规划卡片
    ''' </summary>
    Private Sub ShowAgentPlanCard(session As RalphAgentSession)
        Try
            Dim stepsJson As New StringBuilder()
            stepsJson.Append("[")
            For i = 0 To session.Steps.Count - 1
                If i > 0 Then stepsJson.Append(",")
                Dim s = session.Steps(i)
                stepsJson.Append($"{{""description"":""{EscapeJavaScriptString(s.Description)}"",""detail"":""{EscapeJavaScriptString(s.Detail)}"",""status"":""pending""}}")
            Next
            stepsJson.Append("]")
            
            Dim planJson = $"{{""sessionId"":""{session.Id}"",""understanding"":""{EscapeJavaScriptString(session.Understanding)}"",""steps"":{stepsJson.ToString()},""summary"":""{EscapeJavaScriptString(session.Summary)}""}}"
            
            ExecuteJavaScriptAsyncJS($"showAgentPlanCard({planJson})")
            
        Catch ex As Exception
            Debug.WriteLine($"ShowAgentPlanCard 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理开始执行Agent
    ''' </summary>
    Protected Async Sub HandleStartAgentExecution(jsonDoc As JObject)
        Try
            Debug.WriteLine("[RalphAgent] 用户确认执行")
            
            If _ralphAgentController IsNot Nothing Then
                Await _ralphAgentController.StartExecution()
            End If
            
        Catch ex As Exception
            Debug.WriteLine($"HandleStartAgentExecution 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"执行失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理终止Agent
    ''' </summary>
    Protected Sub HandleAbortAgent()
        Try
            Debug.WriteLine("[RalphAgent] 用户终止Agent")
            
            If _ralphAgentController IsNot Nothing Then
                _ralphAgentController.AbortAgent()
            End If
            
            GlobalStatusStrip.ShowInfo("已终止Agent")
            
        Catch ex As Exception
            Debug.WriteLine($"HandleAbortAgent 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取当前Office内容（子类重写以提供具体实现）
    ''' </summary>
    Protected Overridable Function GetCurrentOfficeContent() As String
        ' 基类默认实现：尝试获取选区内容
        Try
            Dim selInfo = CaptureCurrentSelectionInfo("")
            If selInfo IsNot Nothing AndAlso Not String.IsNullOrEmpty(selInfo.SelectedText) Then
                Return selInfo.SelectedText
            End If
        Catch
        End Try
        Return "(无选中内容)"
    End Function

    ''' <summary>
    ''' 发送AI请求并获取完整响应（用于Agent规划）
    ''' </summary>
    Private Async Function SendAndGetResponse(prompt As String, systemPrompt As String) As Task(Of String)
        ' 这里需要实现一个同步等待AI响应的方法
        ' 简单实现：使用Send方法但收集完整响应
        Try
            Dim responseBuilder As New StringBuilder()
            Dim completed As Boolean = False
            Dim uuid = Guid.NewGuid().ToString()
            
            ' 创建临时的响应收集器
            _agentResponseBuffer = New StringBuilder()
            _agentResponseUuid = uuid
            _agentResponseCompleted = False
            
            ' 发送请求
            Await Send(prompt, systemPrompt, False, "agent_planning")
            
            ' 等待响应完成（最多60秒）
            Dim timeout = 60000
            Dim waited = 0
            While Not _agentResponseCompleted AndAlso waited < timeout
                Await Task.Delay(100)
                waited += 100
            End While
            
            Dim result = _agentResponseBuffer.ToString()
            _agentResponseBuffer = Nothing
            _agentResponseUuid = Nothing
            
            Return result
        Catch ex As Exception
            Debug.WriteLine($"SendAndGetResponse 出错: {ex.Message}")
            Return ""
        End Try
    End Function
    
    ' Agent响应收集
    Private _agentResponseBuffer As StringBuilder
    Private _agentResponseUuid As String
    Private _agentResponseCompleted As Boolean

#End Region

    ''' <summary>
    ''' 处理打开文件对话框请求
    ''' </summary>
    Protected Sub HandleOpenFileDialog()
        Try
            ' 需要在UI线程上执行
            If Me.InvokeRequired Then
                Me.Invoke(Sub() HandleOpenFileDialog())
                Return
            End If

            Using dialog As New OpenFileDialog()
                dialog.Title = "选择要引用的文件"
                dialog.Filter = "Excel文件|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv|" &
                               "Word文件|*.doc;*.docx|" &
                               "PowerPoint文件|*.ppt;*.pptx|" &
                               "所有支持的文件|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv;*.doc;*.docx;*.ppt;*.pptx"
                dialog.FilterIndex = 4 ' 默认显示所有支持的文件
                dialog.Multiselect = True

                If dialog.ShowDialog() = DialogResult.OK Then
                    ' 构建文件列表JSON
                    Dim filesArray As New JArray()
                    For Each filePath In dialog.FileNames
                        Dim fileObj As New JObject()
                        fileObj("name") = Path.GetFileName(filePath)
                        fileObj("path") = filePath
                        filesArray.Add(fileObj)
                    Next

                    ' 发送给前端
                    ExecuteJavaScriptAsyncJS($"addFilesFromDialog({filesArray.ToString(Formatting.None)})")
                    Debug.WriteLine($"选择了 {dialog.FileNames.Length} 个文件")
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine($"HandleOpenFileDialog 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("打开文件对话框时出错")
        End Try
    End Sub

    Protected Overridable Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim preview As Boolean = Boolean.Parse(jsonDoc("executecodePreview"))
        Dim language As String = jsonDoc("language").ToString()
        Dim responseUuid As String = If(jsonDoc("responseUuid")?.ToString(), "")
        
        Try
            ' 执行代码
            ExecuteCode(code, language, preview)
            
            ' 执行成功后通知前端（清空引用区、更新按钮状态）
            If Not String.IsNullOrEmpty(responseUuid) Then
                ExecuteJavaScriptAsyncJS($"handleExecutionSuccess('{responseUuid}')")
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleExecuteCode 执行出错: {ex.Message}")
            ' 执行失败后通知前端（恢复按钮可点击）
            If Not String.IsNullOrEmpty(responseUuid) Then
                Dim escapedMsg = ex.Message.Replace("'", "\'").Replace(vbCrLf, " ")
                ExecuteJavaScriptAsyncJS($"handleExecutionError('{responseUuid}', '{escapedMsg}')")
            End If
        End Try
    End Sub


    ' 抽象方法，由子类实现
    Protected MustOverride Function ParseFile(filePath As String) As FileContentResult
    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String

    ' 文本/CSV 解析已委托给 FileParserService，请使用 _fileParserService.ParseTextFile()

    Protected MustOverride Function GetApplication() As ApplicationInfo
    
    ''' <summary>
    ''' 获取Office应用类型，用于前端区分Word/PowerPoint/Excel
    ''' </summary>
    Protected Overridable Function GetOfficeAppType() As String
        Return "Unknown"
    End Function
    
    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
    Protected MustOverride Function RunCode(vbaCode As String)

    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' 执行代码的方法 - 委托给 CodeExecutionService
    Public Sub ExecuteCode(code As String, language As String, preview As Boolean)
        CodeExecutionService.ExecuteCode(code, language, preview)
    End Sub

    ' ExecuteJavaScript 已委托给 CodeExecutionService
    ' 添加清除特定 sheetName 的方法
    Public Async Sub ClearSelectedContentBySheetName(sheetName As String)
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
        $"clearSelectedContentBySheetName({JsonConvert.SerializeObject(sheetName)})"
    )
    End Sub


    ' 抽象方法 - 获取Office应用程序对象
    Protected MustOverride Function GetOfficeApplicationObject() As Object

    ' ExecuteExcelFormula, ExecuteVBACode, ContainsProcedureDeclaration, FindFirstProcedureName 已委托给 CodeExecutionService

    ' 虚方法 - 评估Excel公式（只有Excel子类会实现）
    Protected Overridable Function EvaluateFormula(formula As String, preview As Boolean) As Boolean
        ' 默认实现返回Nothing
        Return True
    End Function

    ' 在类字段区：新增 response mode 映射
    Protected _responseModeMap As New Dictionary(Of String, String)() ' responseUuid -> mode (e.g. "reformat","proofread","revisions_only","comparison_only")

    ' 测试方法已移除，如需调试请使用单独的测试类

    Private Function TryExtractJsonArrayFromText(text As String) As JArray
        Return UtilsService.TryExtractJsonArrayFromText(text)
    End Function

    ' 存储调用Send时的请求参数（requestUuid/responseUuid -> JObject）
    Protected _savedRequestParams As New Dictionary(Of String, JObject)()

    Public Async Function Send(question As String, systemPrompt As String, addHistory As Boolean, responseMode As String) As Task
        Dim apiUrl As String = ConfigSettings.ApiUrl
        Dim apiKey As String = ConfigSettings.ApiKey

        If String.IsNullOrWhiteSpace(apiKey) Then
            GlobalStatusStrip.ShowWarning("请先配置大模型ApiKey！")
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        If String.IsNullOrWhiteSpace(apiUrl) Then
            GlobalStatusStrip.ShowWarning("请先配置大模型Api！")
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        If String.IsNullOrWhiteSpace(question) Then
            GlobalStatusStrip.ShowWarning("请输入问题！")
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        Dim uuid As String = Guid.NewGuid().ToString()
        ' 这里生成 requestUuid（用于绑定选区）
        Dim requestUuid As String = Guid.NewGuid().ToString()


        ' 将 PendingSelectionInfo 绑定到 requestUuid
        Try
            If PendingSelectionInfo Is Nothing Then
                Dim captured As SelectionInfo = Nothing
                Try
                    captured = CaptureCurrentSelectionInfo(responseMode)
                Catch ex As Exception
                    Debug.WriteLine("CaptureCurrentSelectionInfo 异常: " & ex.Message)
                End Try
                If captured IsNot Nothing Then
                    PendingSelectionInfo = captured
                End If
            End If

            ' 将 PendingSelectionInfo 绑定到 requestUuid（原有逻辑）
            If PendingSelectionInfo IsNot Nothing Then
                Try
                    _selectionPendingMap(requestUuid) = PendingSelectionInfo
                Catch ex As Exception
                    Debug.WriteLine($"绑定 PendingSelectionInfo 到 requestUuid 失败: {ex.Message}")
                End Try
                ' 清空 PendingSelectionInfo，避免被下一个请求误用
                PendingSelectionInfo = Nothing
            End If
        Catch
        End Try

        Try
            If String.IsNullOrWhiteSpace(systemPrompt) Then
                ' 使用PromptManager生成组合后的提示词
                Dim appInfo = GetApplication()
                Dim appType = If(appInfo IsNot Nothing, appInfo.Type.ToString(), "Excel")
                
                Dim context As New PromptContext With {
                    .ApplicationType = appType,
                    .IntentResult = CurrentIntentResult,
                    .FunctionMode = responseMode
                }
                
                systemPrompt = PromptManager.Instance.GetCombinedPrompt(context)
                
                ' 如果PromptManager返回空（没有配置），使用基础提示词
                If String.IsNullOrWhiteSpace(systemPrompt) Then
                    systemPrompt =
                    "系统指令（必读）：" & vbCrLf & ConfigSettings.propmtContent & vbCrLf & vbCrLf &
                    "1) 首先输出一个名为 'Plan' 的简短计划，按步骤列出解决路径（要点式，最多6条）。" & vbCrLf &
                    "2) 然后输出名为 'Answer' 的部分，给出最终可执行的解决方案或操作步骤，使用 Markdown，必要时给出代码/示例或差异说明。" & vbCrLf &
                    "3) 如果信息不足，请不要猜测；在最后输出名为 'Clarifying Questions' 的部分，列出需要用户回答的问题并暂停执行。" & vbCrLf &
                    "4) 对于用户请求的改进（用户标记当前回答为不接受），在回复开头先写明 '改进点'（1-3 行），然后给出修正的 Plan 与 Answer。" & vbCrLf &
                    "5) 保持回答简洁、有条理，优先提供可直接执行的结论和示例。"
                End If
            End If


            Dim requestBody As String = CreateRequestBody(requestUuid, question, systemPrompt, addHistory)
            Await SendHttpRequestStream(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody, StripQuestion(question), requestUuid, addHistory, responseMode)
            Await SaveFullWebPageAsync2()
        Catch ex As Exception
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try

    End Function

    Private Sub ManageHistoryMessageSize()
        ' 如果历史消息数超过限制，有一条system和assistant，所以+2
        While systemHistoryMessageData.Count > contextLimit + 2
            ' 保留系统消息（第一条消息）
            If systemHistoryMessageData.Count > 2 Then
                ' 移除第二条消息（最早的非系统消息）
                systemHistoryMessageData.RemoveAt(2)
            End If
        End While
    End Sub


    Private Function StripQuestion(question As String) As String
        Return UtilsService.StripQuestion(question)
    End Function

    Private Function CreateRequestBody(uuid As String, question As String, systemPrompt As String, addHistory As Boolean) As String
        Dim result As String = StripQuestion(question)

        ' 构建 messages 数组（使用 JArray 确保正确的 JSON 序列化）
        Dim messagesArray As New JArray()

        Dim systemMessage = New HistoryMessage() With {
            .role = "system",
            .content = systemPrompt
        }
        Dim q = New HistoryMessage() With {
                .role = "user",
                .content = result
            }

        If addHistory Then
            ' 添加或替换 system 消息（保证只有一条 system）
            Dim existingSystem = systemHistoryMessageData.FirstOrDefault(Function(m) m.role = "system")
            If existingSystem IsNot Nothing Then
                systemHistoryMessageData.Remove(existingSystem)
            End If
            systemHistoryMessageData.Insert(0, systemMessage)
            systemHistoryMessageData.Add(q)

            ' 管理历史消息大小
            ManageHistoryMessageSize()

            ' 将历史消息转换为 JObject（自动处理所有特殊字符转义）
            For Each message In systemHistoryMessageData
                Dim msgObj = New JObject()
                msgObj("role") = message.role
                msgObj("content") = If(message.content, String.Empty)
                messagesArray.Add(msgObj)
            Next
        Else
            ' 仅使用当前消息，不添加历史
            Dim tempMessageData As New List(Of HistoryMessage)
            tempMessageData.Insert(0, systemMessage)
            tempMessageData.Add(q)
            For Each message In tempMessageData
                Dim msgObj = New JObject()
                msgObj("role") = message.role
                msgObj("content") = If(message.content, String.Empty)
                messagesArray.Add(msgObj)
            Next
        End If



        ' 添加MCP工具信息（如果有）
        Dim toolsArray As JArray = Nothing
        Dim chatSettings As New ChatSettings(GetApplication())

        ' 如果有启用的MCP连接
        If chatSettings.EnabledMcpList IsNot Nothing AndAlso chatSettings.EnabledMcpList.Count > 0 Then
            toolsArray = New JArray()

            ' 加载所有MCP连接
            Dim connections = MCPConnectionManager.LoadConnections()

            ' 找到启用的连接
            For Each mcpName In chatSettings.EnabledMcpList
                ' 使用IsActive替代Enabled
                Dim connection = connections.FirstOrDefault(Function(c) c.Name = mcpName AndAlso c.IsActive)
                If connection IsNot Nothing Then
                    ' 从连接配置中获取已保存的工具列表
                    If connection.Tools IsNot Nothing AndAlso connection.Tools.Count > 0 Then
                        ' 将所有工具添加到工具数组
                        For Each toolObj In connection.Tools
                            toolsArray.Add(toolObj)
                        Next
                        Debug.WriteLine($"从连接 '{connection.Name}' 加载了 {connection.Tools.Count} 个工具")
                    Else
                        ' 如果连接中没有保存工具信息，则使用通用的mcp_call工具
                        Dim toolObj = New JObject()
                        toolObj("type") = "function"
                        toolObj("function") = New JObject()
                        toolObj("function")("name") = "mcp_call"
                        toolObj("function")("description") = $"Call MCP tool through {connection.Name} connection"

                        ' 添加参数架构
                        toolObj("function")("parameters") = New JObject()
                        toolObj("function")("parameters")("type") = "object"
                        toolObj("function")("parameters")("properties") = New JObject()

                        ' 工具名称参数
                        toolObj("function")("parameters")("properties")("tool_name") = New JObject()
                        toolObj("function")("parameters")("properties")("tool_name")("type") = "string"
                        toolObj("function")("parameters")("properties")("tool_name")("description") = "The name of the MCP tool to call"

                        ' 工具参数
                        toolObj("function")("parameters")("properties")("arguments") = New JObject()
                        toolObj("function")("parameters")("properties")("arguments")("type") = "object"
                        toolObj("function")("parameters")("properties")("arguments")("description") = "The arguments to pass to the MCP tool"

                        ' 添加必需参数
                        toolObj("function")("parameters")("required") = New JArray({"tool_name", "arguments"})

                        ' 添加到工具数组
                        toolsArray.Add(toolObj)
                        Debug.WriteLine($"连接 '{connection.Name}' 没有保存工具信息，使用通用mcp_call工具")
                    End If
                End If
            Next
        End If

        ' 构建 JSON 请求体（使用 JObject 确保正确序列化）
        Dim requestObj = New JObject()
        requestObj("model") = ConfigSettings.ModelName
        requestObj("messages") = messagesArray
        requestObj("stream") = True

        ' 如果有工具，添加到请求中
        If toolsArray IsNot Nothing AndAlso toolsArray.Count > 0 Then
            requestObj("tools") = toolsArray
        End If

        Return requestObj.ToString(Newtonsoft.Json.Formatting.None)

    End Function


    ' 添加处理MCP工具调用的方法
    Private Async Function HandleMcpToolCall(toolName As String, arguments As JObject, mcpConnectionName As String) As Task(Of JObject)
        Try
            Debug.WriteLine($"开始处理MCP工具调用: 工具={toolName}, 连接={mcpConnectionName}")

            ' 加载MCP连接
            Dim connections = MCPConnectionManager.LoadConnections()
            ' 注意这里使用isActive而不是Enabled
            Dim connection = connections.FirstOrDefault(Function(c) c.Name = mcpConnectionName AndAlso c.IsActive)

            If connection Is Nothing Then
                Return CreateErrorResponse($"MCP连接 '{mcpConnectionName}' 未找到或未启用。可用连接: {String.Join(", ", connections.Where(Function(c) c.IsActive).Select(Function(c) c.Name))}")
            End If

            Debug.WriteLine($"找到MCP连接: {connection.Name}, URL: {connection.Url}")

            ' 创建MCP客户端
            Using client As New StreamJsonRpcMCPClient()
                Try
                    ' 配置客户端
                    Await client.ConfigureAsync(connection.Url)
                    Debug.WriteLine("MCP客户端配置完成")

                    ' 初始化连接
                    Dim initResult = Await client.InitializeAsync()
                    If Not initResult.Success Then
                        Return CreateErrorResponse($"初始化MCP连接失败: {initResult.ErrorMessage}。连接URL: {connection.Url}")
                    End If

                    Debug.WriteLine("MCP连接初始化成功")

                    ' 调用工具
                    Debug.WriteLine($"开始调用工具: {toolName}, 参数: {arguments.ToString()}")
                    Dim result = Await client.CallToolAsync(toolName, arguments)

                    ' 处理结果
                    If result.IsError Then
                        Return CreateErrorResponse($"调用MCP工具 '{toolName}' 失败: {result.ErrorMessage}")
                    End If

                    Debug.WriteLine($"工具调用成功，返回内容数量: {result.Content?.Count}")

                    ' 创建成功响应
                    Dim responseObj = New JObject()

                    ' 添加内容数组
                    Dim contentArray = New JArray()
                    If result.Content IsNot Nothing Then
                        For Each content In result.Content
                            Dim contentObj = New JObject()
                            contentObj("type") = content.Type

                            If Not String.IsNullOrEmpty(content.Text) Then
                                contentObj("text") = content.Text
                            End If

                            If Not String.IsNullOrEmpty(content.Data) Then
                                contentObj("data") = content.Data
                            End If

                            If Not String.IsNullOrEmpty(content.MimeType) Then
                                contentObj("mimeType") = content.MimeType
                            End If

                            contentArray.Add(contentObj)
                        Next
                    End If

                    responseObj("content") = contentArray
                    Return responseObj

                Catch clientEx As Exception
                    Debug.WriteLine($"MCP客户端操作失败: {clientEx.Message}")
                    Return CreateErrorResponse($"MCP客户端操作失败: {clientEx.Message}。详细信息: {clientEx.ToString()}")
                End Try
            End Using
        Catch ex As Exception
            Debug.WriteLine($"HandleMcpToolCall整体异常: {ex.Message}")
            Return CreateErrorResponse($"MCP工具调用出现异常: {ex.Message}。工具: {toolName}, 连接: {mcpConnectionName}。堆栈跟踪: {ex.StackTrace}")
        End Try
    End Function

    ' 创建错误响应
    Private Function CreateErrorResponse(errorMessage As String) As JObject
        Dim responseObj = New JObject()
        responseObj("isError") = True
        responseObj("errorMessage") = errorMessage
        responseObj("timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        Debug.WriteLine($"创建错误响应: {errorMessage}")
        Return responseObj
    End Function
    ' 添加一个结构来存储token信息
    Private Structure TokenInfo
        Public PromptTokens As Integer
        Public CompletionTokens As Integer
        Public TotalTokens As Integer
    End Structure

    Private totalTokens As Integer = 0
    Private lastTokenInfo As Nullable(Of TokenInfo)

    ' 用于累加当前会话中所有API调用的token消耗（mcp多次消耗的情况）
    Private currentSessionTotalTokens As Integer = 0

    ' 用于跟踪待处理的异步任务
    Private _pendingMcpTasks As Integer = 0
    Private _mainStreamCompleted As Boolean = False
    Private _finalUuid As String = String.Empty


    ' 现在接收 requestUuid，内部生成 responseUuid（用于前端展示），并建立 response->request 映射
    Private Async Function SendHttpRequestStream(apiUrl As String, apiKey As String, requestBody As String, originQuestion As String, requestUuid As String, addHistory As Boolean, responseMode As String) As Task

        ' responseUuid 用于前端显示（与 requestUuid 分离）
        Dim responseUuid As String = Guid.NewGuid().ToString()

        ' 保存映射：response -> request
        Try
            _responseToRequestMap(responseUuid) = requestUuid
            ' 保存 response -> mode 映射（用于决定 showComparison/showRevisions 行为）
            If Not String.IsNullOrEmpty(responseMode) Then
                _responseModeMap(responseUuid) = responseMode
            End If

            ' 如果之前在 request 级别有选区信息（旧逻辑可能把选区存到 _selectionPendingMap(requestUuid)），
            ' 则立即把选区迁移到以 responseUuid 为键的映射，后续完成阶段直接用 responseUuid 查找。
            If Not String.IsNullOrEmpty(requestUuid) AndAlso _selectionPendingMap.ContainsKey(requestUuid) Then
                Try
                    _responseSelectionMap(responseUuid) = _selectionPendingMap(requestUuid)
                    ' 可选地从 request map 中移除，避免内存泄露
                    _selectionPendingMap.Remove(requestUuid)
                Catch ex As Exception
                    Debug.WriteLine("迁移选区信息到 responseSelectionMap 失败: " & ex.Message)
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine($"保存 response->request/response->mode 映射失败: {ex.Message}")
        End Try

        ' 保持以前使用的 _finalUuid 用于现有完成逻辑（注意：这是 responseUuid）
        _finalUuid = responseUuid
        _mainStreamCompleted = False
        _pendingMcpTasks = 0

        ' 重置当前会话的token累加器
        currentSessionTotalTokens = 0

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = Timeout.InfiniteTimeSpan

                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Debug.WriteLine("[HTTP] 开始发送流式请求...")
                Debug.WriteLine($"[HTTP] Request Body (for requestUuid={requestUuid}): {requestBody}")

                Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName

                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()
                    Debug.WriteLine($"[HTTP] 响应状态码: {response.StatusCode}")

                    ' 创建前端聊天节（使用 responseUuid 作为显示 id）
                    Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
                    Await ExecuteJavaScriptAndWaitAsync(jsCreate)

                    ' 等待确认 rendererMap 已创建
                    Await WaitForRendererMapAsync(responseUuid)

                    ' 在前端 DOM 的 chat 节上设置 dataset.requestId，以便前端后续执行时可以把 requestUuid 发回
                    Dim jsSetMapping As String = $"(function(){{ var el = document.getElementById('chat-{responseUuid}'); if(el) el.dataset.requestId = '{requestUuid}'; }})();"
                    Await ExecuteJavaScriptAndWaitAsync(jsSetMapping)

                    ' 处理流（后续逻辑不变，但使用 responseUuid 进行 flush 等操作）
                    Dim stringBuilder As New StringBuilder()
                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Dim buffer(102300) As Char
                            Dim readCount As Integer
                            Dim chunkCount As Integer = 0
                            Do
                                If stopReaderStream Then
                                    Debug.WriteLine("[Stream] 用户手动停止流读取")
                                    _currentMarkdownBuffer.Clear()
                                    allMarkdownBuffer.Clear()
                                    Exit Do
                                End If
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do
                                chunkCount += 1
                                Dim chunk As String = New String(buffer, 0, readCount)
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)

                                ' 调试：记录每次读取的数据
                                If chunkCount <= 3 Then
                                    Debug.WriteLine($"[Stream] chunk#{chunkCount} 长度={readCount}, 内容前100字符: {chunk.Substring(0, Math.Min(100, chunk.Length))}")
                                End If

                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    ProcessStreamChunk(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), responseUuid, originQuestion)
                                    stringBuilder.Clear()
                                End If
                            Loop

                            ' 调试：如果循环结束但stringBuilder不为空，说明有未处理的数据
                            If stringBuilder.Length > 0 Then
                                Debug.WriteLine($"[Stream] 警告：循环结束但stringBuilder还有未处理数据，长度={stringBuilder.Length}")
                                Debug.WriteLine($"[Stream] 未处理数据内容: {stringBuilder.ToString().Substring(0, Math.Min(200, stringBuilder.Length))}")
                                ' 尝试处理剩余数据
                                ProcessStreamChunk(stringBuilder.ToString().Trim(), responseUuid, originQuestion)
                            End If

                            Debug.WriteLine($"[Stream] 流接收完成，共处理了 {chunkCount} 个chunk")
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] 请求过程中出错: {ex.ToString()}")
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _mainStreamCompleted = True

            Dim finalTokens As Integer = currentSessionTotalTokens
            If lastTokenInfo.HasValue Then
                finalTokens += lastTokenInfo.Value.TotalTokens
                currentSessionTotalTokens += lastTokenInfo.Value.TotalTokens
            End If

            Debug.WriteLine($"finally 当前会话总tokens: {currentSessionTotalTokens}")

            ' Check 完成：会使用 _finalUuid（即 responseUuid）
            CheckAndCompleteProcessing()

            Dim answer = New HistoryMessage() With {
            .role = "assistant",
            .content = $"这是大模型基于用户问题的答复作为历史参考：{allMarkdownBuffer.ToString()}"
        }

            If addHistory Then
                systemHistoryMessageData.Add(answer)
                ManageHistoryMessageSize()
            End If

            allMarkdownBuffer.Clear()
            lastTokenInfo = Nothing
        End Try
    End Function

    ' 在类字段区：新增 response -> selection 映射（用于在 responseUuid 可用时快速查找选区）
    Protected _responseSelectionMap As New Dictionary(Of String, SelectionInfo)() ' responseUuid -> SelectionInfo

    ' 检查并完成处理
    Private Sub CheckAndCompleteProcessing()
        Debug.WriteLine($"CheckAndCompleteProcessing: 主流完成={_mainStreamCompleted}, 待处理MCP任务={_pendingMcpTasks}")

        ' 只有在主流完成且没有待处理的MCP任务时才调用完成函数
        If _mainStreamCompleted AndAlso _pendingMcpTasks = 0 Then
            Debug.WriteLine("所有处理完成，调用 processStreamComplete")
            ExecuteJavaScriptAsyncJS($"processStreamComplete('{_finalUuid}',{currentSessionTotalTokens});")
            CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
        End If
    End Sub


    ' 会话完成的钩子，可自行实现
    Protected Overridable Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)
        ' 处理续写模式的完成 - 显示续写预览界面
        If _responseModeMap.ContainsKey(_finalUuid) AndAlso _responseModeMap(_finalUuid) = "continuation" Then
            ExecuteJavaScriptAsyncJS($"showContinuationPreview('{_finalUuid}');")
        End If

        ' 处理模板渲染模式的完成 - 显示模板预览界面并完全隐藏代码块
        If _responseModeMap.ContainsKey(_finalUuid) AndAlso _responseModeMap(_finalUuid) = "template_render" Then
            ExecuteJavaScriptAsyncJS($"showTemplatePreview('{_finalUuid}');")
            ExecuteJavaScriptAsyncJS($"hideAllCodeBlockActions('{_finalUuid}');") ' 完全隐藏代码块操作栏
        End If

        ' 校对/排版模式 - 隐藏代码块的编辑和执行按钮（只保留复制）
        If _responseModeMap.ContainsKey(_finalUuid) Then
            Dim mode = _responseModeMap(_finalUuid)
            If mode = "proofread" OrElse mode = "reformat" Then
                ExecuteJavaScriptAsyncJS($"hideCodeActionButtons('{_finalUuid}');")
            End If
        End If

        ' Ralph Loop 完成检查
        CheckRalphLoopCompletion(allPlainMarkdownBuffer.ToString())
    End Sub


    Private ReadOnly markdownPipeline As MarkdownPipeline = New MarkdownPipelineBuilder() _
    .UseAdvancedExtensions() _      ' 启用表格、代码块等扩展
    .Build()                        ' 构建不可变管道

    Private _currentMarkdownBuffer As New StringBuilder()
    Private allMarkdownBuffer As New StringBuilder()

    ' 用于收集工具调用参数的变量
    Private _pendingToolCalls As New Dictionary(Of String, JObject) ' 按ID存储未完成的工具调用
    Private _completedToolCalls As New List(Of JObject) ' 存储已完成的工具调用


    Private Sub ProcessStreamChunk(rawChunk As String, uuid As String, originQuestion As String)
        Try
            Debug.WriteLine($"[ProcessStreamChunk] 收到原始数据长度: {rawChunk.Length}")
            Dim lines As String() = rawChunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            Debug.WriteLine($"[ProcessStreamChunk] 分割后行数: {lines.Length}")

            For Each line In lines
                line = line.Trim()
                If line = "[DONE]" Then
                    Debug.WriteLine("[ProcessStreamChunk] 收到 [DONE] 标记")
                    ' 在流结束时处理所有完成的工具调用
                    If _pendingToolCalls.Count > 0 Then
                        Debug.WriteLine("[DONE] 时发现未处理的工具调用，开始处理")
                        ProcessCompletedToolCalls(uuid, originQuestion)
                    End If
                    FlushBuffer("content", uuid) ' 最后刷新缓冲区

                    ' 标记Agent响应完成
                    _agentResponseCompleted = True
                    Return
                End If
                If line = "" Then
                    Continue For
                End If

                Dim jsonObj As JObject = JObject.Parse(line)

                ' 获取token信息 - 只保存最后一个响应块的usage信息
                Dim usage = jsonObj("usage")
                If usage IsNot Nothing AndAlso usage.Type = JTokenType.Object Then
                    lastTokenInfo = New TokenInfo With {
                    .PromptTokens = CInt(usage("prompt_tokens")),
                    .CompletionTokens = CInt(usage("completion_tokens")),
                    .TotalTokens = CInt(usage("total_tokens"))
                }
                End If

                Dim reasoning_content As String = jsonObj("choices")(0)("delta")("reasoning_content")?.ToString()
                If Not String.IsNullOrEmpty(reasoning_content) Then
                    _currentMarkdownBuffer.Append(reasoning_content)
                    FlushBuffer("reasoning", uuid)
                End If

                Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()
                If Not String.IsNullOrEmpty(content) Then
                    Debug.WriteLine($"[ProcessStreamChunk] 解析到content: {content.Substring(0, Math.Min(50, content.Length))}...")
                    _currentMarkdownBuffer.Append(content)
                    FlushBuffer("content", uuid)

                    ' 如果是Agent规划请求，同时收集到缓冲区
                    If _agentResponseBuffer IsNot Nothing Then
                        _agentResponseBuffer.Append(content)
                    End If
                End If

                ' 检查是否有工具调用
                Dim choices = jsonObj("choices")
                If choices IsNot Nothing AndAlso choices.Count > 0 Then
                    Dim choice = choices(0)
                    Dim delta = choice("delta")
                    Dim finishReason = choice("finish_reason")?.ToString()

                    ' 收集工具调用数据
                    If delta IsNot Nothing Then
                        Dim toolCalls = delta("tool_calls")
                        If toolCalls IsNot Nothing AndAlso toolCalls.Count > 0 Then
                            CollectToolCallData(toolCalls, originQuestion)
                        End If
                    End If

                    ' 当finish_reason为tool_calls时，说明所有工具调用数据已接收完毕
                    If finishReason = "tool_calls" Then
                        Debug.WriteLine("检测到 finish_reason = tool_calls，开始处理工具调用")
                        ProcessCompletedToolCalls(uuid, originQuestion)
                    End If
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"[ERROR] 数据处理失败: {ex.Message}")
            MessageBox.Show("请求失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 收集工具调用数据
    Private Sub CollectToolCallData(toolCalls As JArray, originQuestion As String)
        Try
            For Each toolCall In toolCalls
                Dim toolIndex = toolCall("index")?.Value(Of Integer)()
                Dim toolId = toolCall("id")?.ToString()

                ' 统一使用index作为主键，因为index是唯一且连续的
                Dim toolKey As String = $"tool_{toolIndex}"

                ' 如果是新的工具调用，创建新的条目
                If Not _pendingToolCalls.ContainsKey(toolKey) Then
                    _pendingToolCalls(toolKey) = New JObject()
                    ' 保存真实的ID，但使用index作为内部键
                    _pendingToolCalls(toolKey)("realId") = If(String.IsNullOrEmpty(toolId), toolKey, toolId)
                    _pendingToolCalls(toolKey)("index") = toolIndex
                    _pendingToolCalls(toolKey)("type") = toolCall("type")?.ToString()
                    _pendingToolCalls(toolKey)("function") = New JObject()
                    _pendingToolCalls(toolKey)("function")("name") = ""
                    _pendingToolCalls(toolKey)("function")("arguments") = ""
                    _pendingToolCalls(toolKey)("processed") = False
                End If

                Dim currentTool = _pendingToolCalls(toolKey)

                ' 累积函数名称
                Dim functionName = toolCall("function")("name")?.ToString()
                If Not String.IsNullOrEmpty(functionName) Then
                    currentTool("function")("name") = functionName
                    Debug.WriteLine($"设置工具名称: Key={toolKey}, Name={functionName}")
                End If

                ' 累积参数
                Dim arguments = toolCall("function")("arguments")?.ToString()
                If Not String.IsNullOrEmpty(arguments) Then
                    Dim currentArgs = currentTool("function")("arguments").ToString()
                    currentTool("function")("arguments") = currentArgs & arguments
                    Debug.WriteLine($"收集工具调用数据: Key={toolKey}, 本次参数片段='{arguments}', 累积后参数长度={currentTool("function")("arguments").ToString().Length}")
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"收集工具调用数据时出错: {ex.Message}")
        End Try
    End Sub

    ' 处理所有已完成的工具调用
    Private Async Sub ProcessCompletedToolCalls(uuid As String, originQuestion As String)
        Try
            If _pendingToolCalls.Count = 0 Then Return

            Debug.WriteLine($"开始处理 {_pendingToolCalls.Count} 个工具调用")

            For Each kvp In _pendingToolCalls
                Dim toolCall = kvp.Value
                Dim toolKey = kvp.Key

                ' 检查是否已经处理过
                If CBool(toolCall("processed")) Then
                    Debug.WriteLine($"工具调用 {toolKey} 已处理，跳过")
                    Continue For
                End If

                Dim toolName = toolCall("function")("name").ToString()
                Dim argumentsStr = toolCall("function")("arguments").ToString()

                ' 验证工具调用是否完整 - 必须同时有名称和参数
                If String.IsNullOrEmpty(toolName) Then
                    Debug.WriteLine($"工具调用 {toolKey} 缺少名称，跳过处理")
                    Continue For
                End If

                ' 如果参数为空，也跳过（除非某些工具真的不需要参数）
                If String.IsNullOrEmpty(argumentsStr) Then
                    Debug.WriteLine($"工具调用 {toolKey} 参数为空，使用空对象")
                End If

                Debug.WriteLine($"处理工具调用: Key={toolKey}, Name={toolName}, Arguments={argumentsStr}")

                ' 标记为已处理，防止重复执行
                toolCall("processed") = True

                ' 验证参数是否为有效JSON
                Dim argumentsObj As JObject = Nothing
                Try
                    If Not String.IsNullOrEmpty(argumentsStr) Then
                        argumentsObj = JObject.Parse(argumentsStr)
                        Debug.WriteLine($"成功解析参数JSON: {argumentsObj.ToString()}")
                    Else
                        argumentsObj = New JObject()
                        Debug.WriteLine("参数为空，使用空对象")
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"工具 {toolName} 的参数格式错误: {ex.Message}, 原始参数: {argumentsStr}")

                    ' 通过FlushBuffer向前端显示详细错误
                    Dim errorMessage = $"<br/>**工具调用参数解析错误：**<br/>" &
                                     $"工具名称: {toolName}<br/>" &
                                     $"错误详情: {ex.Message}<br/>" &
                                     $"原始参数: `{argumentsStr}`<br/>"
                    _currentMarkdownBuffer.Append(errorMessage)
                    FlushBuffer("content", uuid)

                    Continue For ' 跳过这个有问题的工具调用
                End Try

                ' 添加消息到界面，说明正在调用工具
                Dim toolCallMessage = $"<br/>**正在调用工具: {toolName}**<br/>参数: `{argumentsObj.ToString(Newtonsoft.Json.Formatting.None)}`<br/>"
                _currentMarkdownBuffer.Append(toolCallMessage)
                FlushBuffer("content", uuid)

                ' 从设置中获取启用的MCP连接
                Dim chatSettings As New ChatSettings(GetApplication())
                Dim enabledMcpList = chatSettings.EnabledMcpList

                If enabledMcpList IsNot Nothing AndAlso enabledMcpList.Count > 0 Then
                    ' 使用第一个启用的MCP连接
                    Dim mcpConnectionName = enabledMcpList(0)

                    ' 调用工具
                    Dim result = Await HandleMcpToolCall(toolName, argumentsObj, mcpConnectionName)

                    ' 处理结果
                    If result("isError") IsNot Nothing AndAlso CBool(result("isError")) Then
                        ' 通过FlushBuffer显示详细错误信息
                        Dim detailedError = result("content")?.ToString()
                        Dim errorMessage = $"<br/>**工具调用失败：**<br/>" &
                                         $"**工具名称:** {toolName}<br/>" &
                                         $"**连接名称:** {mcpConnectionName}<br/>" &
                                         $"**错误详情:** {detailedError}<br/>" &
                                         $"**调用参数:**<br/>```json{vbCrLf}{argumentsObj.ToString(Newtonsoft.Json.Formatting.Indented)}{vbCrLf}```<br/>"

                        _currentMarkdownBuffer.Append(errorMessage)
                        FlushBuffer("content", uuid)
                    Else
                        ' 增加待处理任务计数
                        _pendingMcpTasks += 1
                        Debug.WriteLine($"增加MCP任务，当前待处理任务数: {_pendingMcpTasks}")

                        ' 不直接显示结果，而是发送给大模型进行润色
                        Await SendToolResultForFormatting(toolName, argumentsObj, result, uuid, originQuestion)
                    End If
                Else
                    ' 没有启用的MCP连接
                    Dim errorMessage = "<br/>**配置错误：**<br/>没有启用的MCP连接，无法调用工具。请在设置中启用MCP连接。<br/>"
                    _currentMarkdownBuffer.Append(errorMessage)
                    FlushBuffer("content", uuid)
                End If
            Next

            ' 清空已处理的工具调用
            _pendingToolCalls.Clear()
            _completedToolCalls.Clear()

        Catch ex As Exception
            Debug.WriteLine($"处理完成的工具调用时出错: {ex.Message}")

            ' 向前端显示处理错误
            Dim errorMessage = $"<br/>**工具调用处理异常：**<br/>" &
                             $"**错误详情:** {ex.Message}<br/>" &
                             $"**堆栈跟踪:**<br/>```{vbCrLf}{ex.StackTrace}{vbCrLf}```<br/>"
            _currentMarkdownBuffer.Append(errorMessage)
            FlushBuffer("content", uuid)
        End Try
    End Sub

    ' 新增方法：发送工具结果给大模型进行润色
    Private Async Function SendToolResultForFormatting(toolName As String, arguments As JObject, result As JObject, uuid As String, originQuestion As String) As Task
        Try
            ' 准备发送给大模型的消息内容
            Dim promptBuilder As New StringBuilder()
            promptBuilder.AppendLine($"用户的原始问题：'{originQuestion}' ,但用户使用了 MCP 工具 '{toolName}'，参数为：")
            promptBuilder.AppendLine("```json")
            promptBuilder.AppendLine(arguments.ToString(Newtonsoft.Json.Formatting.Indented))
            promptBuilder.AppendLine("```")
            promptBuilder.AppendLine()
            promptBuilder.AppendLine("工具执行结果为：")
            promptBuilder.AppendLine("```json")
            promptBuilder.AppendLine(result.ToString(Newtonsoft.Json.Formatting.Indented))
            promptBuilder.AppendLine("```")
            promptBuilder.AppendLine()
            promptBuilder.AppendLine("请将上述结果结合用户的原始问题整理成易于理解的格式，并使用合适的Markdown格式化呈现，突出重要信息。不需要解释工具调用过程，只需要呈现结果。不要重复用户的请求内容。")

            ' 构建请求体
            Dim messagesArray = New JArray()
            Dim systemMessage = New JObject()
            systemMessage("role") = "system"
            systemMessage("content") = "你是一个帮助解释API调用结果的助手。你的任务是将MCP工具返回的JSON结果转换为人类易读的格式，可适当根据用户原始问题作出取舍，并用Markdown呈现，且没有任何一句废话。"

            Dim userMessage = New JObject()
            userMessage("role") = "user"
            userMessage("content") = promptBuilder.ToString()

            messagesArray.Add(systemMessage)
            messagesArray.Add(userMessage)

            Dim requestObj = New JObject()
            requestObj("model") = ConfigSettings.ModelName
            requestObj("messages") = messagesArray
            requestObj("stream") = True

            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)

            ' 用于存储当前MCP润色调用的token信息
            Dim mcpTokenInfo As Nullable(Of TokenInfo) = Nothing

            ' 发送请求
            Using client As New HttpClient()
                client.Timeout = Timeout.InfiniteTimeSpan

                Dim request As New HttpRequestMessage(HttpMethod.Post, ConfigSettings.ApiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", ConfigSettings.ApiKey)
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()

                    ' 处理流响应
                    Dim formattedBuilder As New StringBuilder()
                    formattedBuilder.AppendLine("<br/>**工具调用结果：**<br/>")

                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Dim stringBuilder As New StringBuilder()
                            Dim buffer(1023) As Char
                            Dim readCount As Integer

                            Do
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do

                                Dim chunk As String = New String(buffer, 0, readCount)
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)

                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    Dim lines As String() = stringBuilder.ToString().Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                                    For Each line In lines
                                        line = line.Trim()
                                        If line = "[DONE]" Then
                                            Continue For
                                        End If
                                        If line = "" Then
                                            Continue For
                                        End If

                                        Try
                                            Dim jsonObj As JObject = JObject.Parse(line)
                                            ' 收集token信息
                                            Dim usage = jsonObj("usage")
                                            If usage IsNot Nothing Then
                                                mcpTokenInfo = New TokenInfo With {
                                                    .PromptTokens = CInt(usage("prompt_tokens")),
                                                    .CompletionTokens = CInt(usage("completion_tokens")),
                                                    .TotalTokens = CInt(usage("total_tokens"))
                                                }
                                                'Debug.WriteLine($"MCP润色调用tokens: {mcpTokenInfo.Value.TotalTokens}")
                                            End If

                                            Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()

                                            If Not String.IsNullOrEmpty(content) Then
                                                formattedBuilder.Append(content)
                                            End If
                                        Catch ex As Exception
                                            ' 忽略解析错误
                                            Debug.WriteLine($"解析工具结果润色响应出错: {ex.Message}")
                                        End Try
                                    Next

                                    stringBuilder.Clear()
                                End If
                            Loop
                        End Using
                    End Using

                    ' 显示格式化后的结果
                    _currentMarkdownBuffer.Append(formattedBuilder.ToString())
                    FlushBuffer("content", uuid)

                    ' 累加MCP润色调用的token消耗
                    If mcpTokenInfo.HasValue Then
                        currentSessionTotalTokens += mcpTokenInfo.Value.TotalTokens
                        Debug.WriteLine($"累加MCP润色tokens: {mcpTokenInfo.Value.TotalTokens}, 当前总tokens: {currentSessionTotalTokens}")
                    End If
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"格式化工具结果时出错: {ex.Message}")

            ' 如果格式化失败，直接显示原始结果
            _currentMarkdownBuffer.Append("\n\n**工具调用结果：**\n\n```json\n")
            _currentMarkdownBuffer.Append(result.ToString(Newtonsoft.Json.Formatting.Indented))
            _currentMarkdownBuffer.Append("\n```\n")
            FlushBuffer("content", uuid)
        Finally
            ' 减少待处理任务计数
            _pendingMcpTasks -= 1
            Debug.WriteLine($"MCP任务完成，当前待处理任务数: {_pendingMcpTasks}")

            ' 检查是否可以完成处理
            CheckAndCompleteProcessing()
        End Try
    End Function

    Private Async Sub FlushBuffer(contentType As String, uuid As String)
        If _currentMarkdownBuffer.Length = 0 Then Return
        Dim plainContent As String = _currentMarkdownBuffer.ToString()

        Dim escapedContent = HttpUtility.JavaScriptStringEncode(_currentMarkdownBuffer.ToString())
        _currentMarkdownBuffer.Clear()
        Dim js As String
        If contentType = "reasoning" Then
            js = $"appendReasoning('{uuid}','{escapedContent}');"
        Else
            js = $"appendRenderer('{uuid}','{escapedContent}');"
            allMarkdownBuffer.Append(escapedContent)
            allPlainMarkdownBuffer.Append(plainContent)
        End If
        Debug.Print(js)
        Await ExecuteJavaScriptAsyncJS(js)
    End Sub


    ' 执行js脚本的异步方法
    Public Async Function ExecuteJavaScriptAsyncJS(js As String) As Task
        If ChatBrowser.InvokeRequired Then
            ChatBrowser.Invoke(Sub() ChatBrowser.ExecuteScriptAsync(js))
        Else
            Await ChatBrowser.ExecuteScriptAsync(js)
        End If
    End Function

    ' 执行JS脚本并确保等待完成（解决跨线程调用时的时序问题）
    Private Async Function ExecuteJavaScriptAndWaitAsync(js As String) As Task
        Try
            If ChatBrowser.InvokeRequired Then
                ' 使用 TaskCompletionSource 确保等待完成
                Dim tcs As New TaskCompletionSource(Of Boolean)()
                ChatBrowser.Invoke(Sub()
                                       Try
                                           ChatBrowser.ExecuteScriptAsync(js).ContinueWith(Sub(t)
                                                                                               If t.IsFaulted Then
                                                                                                   tcs.SetException(t.Exception)
                                                                                               Else
                                                                                                   tcs.SetResult(True)
                                                                                               End If
                                                                                           End Sub)
                                       Catch ex As Exception
                                           tcs.SetException(ex)
                                       End Try
                                   End Sub)
                Await tcs.Task
            Else
                Await ChatBrowser.ExecuteScriptAsync(js)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[ExecuteJavaScriptAndWaitAsync] 执行JS出错: {ex.Message}")
        End Try
    End Function

    ' 等待前端 rendererMap 创建完成
    Private Async Function WaitForRendererMapAsync(uuid As String) As Task
        Dim maxRetries As Integer = 10
        Dim delayMs As Integer = 50

        For i As Integer = 0 To maxRetries - 1
            Try
                Dim checkJs As String = $"(window.rendererMap && window.rendererMap['{uuid}']) ? 'true' : 'false'"
                Dim result As String = Nothing

                If ChatBrowser.InvokeRequired Then
                    Dim tcs As New TaskCompletionSource(Of String)()
                    ChatBrowser.Invoke(Sub()
                                           ChatBrowser.ExecuteScriptAsync(checkJs).ContinueWith(Sub(t)
                                                                                                    If t.IsFaulted Then
                                                                                                        tcs.SetResult("false")
                                                                                                    Else
                                                                                                        tcs.SetResult(t.Result)
                                                                                                    End If
                                                                                                End Sub)
                                       End Sub)
                    result = Await tcs.Task
                Else
                    result = Await ChatBrowser.ExecuteScriptAsync(checkJs)
                End If

                ' 结果可能包含引号，清理后判断
                result = result?.Trim(""""c)
                If result = "true" Then
                    Debug.WriteLine($"[WaitForRendererMapAsync] rendererMap[{uuid}] 已就绪，重试次数={i}")
                    Return
                End If
            Catch ex As Exception
                Debug.WriteLine($"[WaitForRendererMapAsync] 检查时出错: {ex.Message}")
            End Try

            Await Task.Delay(delayMs)
        Next

        Debug.WriteLine($"[WaitForRendererMapAsync] 警告：等待超时，rendererMap[{uuid}] 可能未创建")
    End Function

    Private Function DecodeBase64(base64 As String) As String
        Return UtilsService.DecodeBase64(base64)
    End Function

    Private Function EscapeJavaScriptString(input As String) As String
        Return UtilsService.EscapeJavaScriptString(input)
    End Function



    ' 共用的HTTP请求方法 - 委托给 UtilsService
    Protected Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Return Await UtilsService.SendHttpRequestAsync(apiUrl, apiKey, requestBody)
    End Function

    ' 加载本地HTML文件
    Public Async Function LoadLocalHtmlFile() As Task
        Try
            ' 检查HTML文件是否存在
            Dim htmlFilePath As String = ChatHtmlFilePath
            If File.Exists(htmlFilePath) Then

                Await Task.Run(Sub()
                                   Dim htmlContent As String = File.ReadAllText(htmlFilePath, System.Text.Encoding.UTF8)
                                   htmlContent = htmlContent.TrimStart("""").TrimEnd("""")
                                   ' 直接导航到本地HTML文件
                                   If ChatBrowser.InvokeRequired Then
                                       ChatBrowser.Invoke(Sub() ChatBrowser.CoreWebView2.NavigateToString(htmlContent))
                                   Else
                                       ChatBrowser.CoreWebView2.NavigateToString(htmlContent)
                                   End If
                               End Sub)

            End If
        Catch ex As Exception
            Debug.WriteLine($"加载本地HTML文件时出错：{ex.Message}")
        End Try
    End Function

    Public Async Function SaveFullWebPageAsync2() As Task
        Try
            ' 1. 创建目录（同步操作，无需异步）

            Dim dir = Path.GetDirectoryName(ChatHtmlFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' 2. 获取 HTML（异步无阻塞）
            Dim htmlContent As String = Await GetFullHtmlContentAsync()

            ' 3. 保存文件（异步后台线程）
            Await Task.Run(Sub()
                               Dim fullHtml As String = "<!DOCTYPE html>" & Environment.NewLine & htmlContent
                               File.WriteAllText(
                $"{ChatHtmlFilePath}",
                HttpUtility.HtmlDecode(fullHtml),
                System.Text.Encoding.UTF8
            )
                           End Sub)

            Debug.WriteLine("保存成功")
        Catch ex As Exception
            Debug.WriteLine($"保存失败: {ex.Message}")
        End Try
    End Function

    Private Async Function GetFullHtmlContentAsync() As Task(Of String)
        Dim tcs As New TaskCompletionSource(Of String)()

        ' 强制切换到 WebView2 的 UI 线程操作
        ChatBrowser.BeginInvoke(Async Sub()
                                    Try
                                        Await EnsureWebView2InitializedAsync()

                                        Dim js As String = "
                (function(){
                    const serializer = new XMLSerializer();
                    return serializer.serializeToString(document.documentElement);
                })();"

                                        Dim rawResult As String = Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(js)
                                        Dim decodedHtml As String = UnescapeHtmlContent(rawResult)
                                        decodedHtml = decodedHtml.TrimStart("""").TrimEnd("""")

                                        ' 2. 使用正则表达式移除底部输入栏
                                        Dim bottomBarPattern As String = "<div[^>]*id=[""']chat-bottom-bar[""'][^>]*>.*?</div>\s*</div>\s*</div>"
                                        decodedHtml = Regex.Replace(decodedHtml, bottomBarPattern, "", RegexOptions.Singleline)

                                        ' 移除 <script> 标签及其内容
                                        Dim scriptPattern As String = "<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>"
                                        decodedHtml = Regex.Replace(decodedHtml, scriptPattern, String.Empty, RegexOptions.IgnoreCase)

                                        ' 内联本地 CSS 资源（用于离线查看）
                                        decodedHtml = UtilsService.InlineCssResources(decodedHtml)


                                        ' 重新注入必要的JavaScript代码
                                        Dim essentialScript As String = GetEssentialJavaScript()

                                        ' 在 </body> 标签前插入必要的脚本
                                        If decodedHtml.Contains("</body>") Then
                                            decodedHtml = decodedHtml.Replace("</body>", essentialScript & Environment.NewLine & "</body>")
                                        Else
                                            ' 如果没有 </body> 标签，在末尾添加
                                            decodedHtml &= essentialScript
                                        End If

                                        tcs.SetResult(decodedHtml)
                                    Catch ex As Exception
                                        tcs.SetException(ex)
                                    End Try
                                End Sub)

        Return Await tcs.Task
    End Function

    Private Function GetEssentialJavaScript() As String
        Return UtilsService.GetEssentialJavaScript()
    End Function

    Private Async Function EnsureWebView2InitializedAsync() As Task
        If ChatBrowser.CoreWebView2 Is Nothing Then
            Await ChatBrowser.EnsureCoreWebView2Async()
        End If
    End Function

    Private Function UnescapeHtmlContent(htmlContent As String) As String
        ' 处理转义字符（直接从 JSON 字符串中提取）
        Return System.Text.RegularExpressions.Regex.Unescape(
        htmlContent
    )
    End Function

    ' HistoryMessage 类已移至 Controls/Models/HistoryMessage.vb

    ' 注入辅助脚本
    Protected Sub InitializeWebView2Script()
        ' 设置 Web 消息处理器
        AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived
        ' 注入 VSTO 桥接脚本
        ChatBrowser.ExecuteScriptAsync(UtilsService.GetVstoBridgeScript())
    End Sub

    ' 选中内容发送到聊天区
    Public Async Sub AddSelectedContentItem(sheetName As String, address As String)
        Dim ctrlKey As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
    $"addSelectedContentItem({JsonConvert.SerializeObject(sheetName)}, {JsonConvert.SerializeObject(address)}, {ctrlKey.ToString().ToLower()})"
)
    End Sub


    ' VBA 异常处理 - 委托给 UtilsService
    Protected Shared Sub VBAxceptionHandle(ex As Runtime.InteropServices.COMException)
        UtilsService.HandleVbaException(ex)
    End Sub


    Protected Overridable Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
    End Sub
End Class