' ShareRibbon\Controls\BaseChatControl.vb
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.JSON
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public MustInherit Class BaseChatControl
    Inherits UserControl
    Implements IChatRoutingHost

    ' 服务类实例
    Private _fileParserService As New FileParserService()
    Protected _chatStateService As New ChatStateService()
    Private _mcpService As McpService = Nothing
    Private _selfCheckLoopController As SelfCheckLoopController = Nothing
    Private _conversationRuntime As IConversationRuntime = Nothing
    Private ReadOnly _webViewInitSemaphore As New SemaphoreSlim(1, 1)
    ' E4-05：页面装载只做一次。这份文档注入设置、MCP 状态和命令面板状态各一份就够，
    ' 面板里之后的任何导航都不应该再重跑一遍（重跑会把开关值重复下发、把监听重复注册）。
    ' 真正的重新初始化（CoreWebView2 重建并重新 NavigateToString）在装载前把它复位。
    Private _pageBootstrapDone As Boolean
    Private _commandRouter As ChatCommandRouter = Nothing
    Private _webViewBridge As WebViewBridge = Nothing
    Private _systemPromptResolver As ChatSystemPromptResolver = Nothing
    Private _sendValidator As ChatSendValidator = Nothing
    Private _modelConfigGate As ModelConfigGate = Nothing
    Private _memoryTurnRecorder As MemoryTurnRecorder = Nothing
    Private _aiNativeRuntime As Agent.IAiNativeRuntime = Nothing
    Private _chatRequestOrchestrator As ChatRequestOrchestrator = Nothing
    Private _chatRoutingOrchestrator As ChatRoutingOrchestrator = Nothing

    Private ReadOnly Property CommandRouter As ChatCommandRouter
        Get
            If _commandRouter Is Nothing Then
                _commandRouter = New ChatCommandRouter(Sub(messageType) Debug.WriteLine($"未知消息类型: {messageType}"))
                RegisterWebViewCommandHandlers(_commandRouter)
            End If
            Return _commandRouter
        End Get
    End Property

    Private ReadOnly Property WebViewBridge As WebViewBridge
        Get
            If _webViewBridge Is Nothing Then
                _webViewBridge = New WebViewBridge(ChatBrowser, AddressOf InitializeWebView2)
            End If
            Return _webViewBridge
        End Get
    End Property

    Private ReadOnly Property SystemPromptResolver As ChatSystemPromptResolver
        Get
            If _systemPromptResolver Is Nothing Then
                _systemPromptResolver = New ChatSystemPromptResolver()
            End If
            Return _systemPromptResolver
        End Get
    End Property

    Private ReadOnly Property SendValidator As ChatSendValidator
        Get
            If _sendValidator Is Nothing Then
                _sendValidator = New ChatSendValidator()
            End If
            Return _sendValidator
        End Get
    End Property

    Private ReadOnly Property MemoryTurnRecorder As MemoryTurnRecorder
        Get
            If _memoryTurnRecorder Is Nothing Then
                _memoryTurnRecorder = New MemoryTurnRecorder()
            End If
            Return _memoryTurnRecorder
        End Get
    End Property

    Private ReadOnly Property ConfigGate As ModelConfigGate
        Get
            If _modelConfigGate Is Nothing Then
                _modelConfigGate = New ModelConfigGate()
            End If
            Return _modelConfigGate
        End Get
    End Property

    ''' <summary>
    ''' E3-03：配置缺失时阻断，并在面板内给出可点的 CTA，而不是只在状态栏报错。
    ''' 聊天和 Agent 两条入口共用，返回 True 表示本轮已经拦下。
    ''' </summary>
    Protected Function IsBlockedByModelConfig(blockedAction As String, Optional validation As ChatSendValidationResult = Nothing) As Boolean
        Dim prompt = If(validation Is Nothing,
                        ConfigGate.Evaluate(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, blockedAction),
                        ConfigGate.Evaluate(validation, blockedAction))
        If prompt Is Nothing Then Return False

        ExecuteJavaScriptAsyncJS(ModelConfigGate.BuildShowScript(prompt))
        Return True
    End Function

    ''' <summary>
    ''' E3-05：第一次云调用（以及撤回之后的第一次）之前必须先有同意。撤回复用同一张卡片，
    ''' 所以拦下的位置和给用户看的承诺不可能各说一套。返回 True 表示本轮已经拦下。
    ''' </summary>
    Protected Function IsBlockedByPrivacy(blockedAction As String) As Boolean
        Dim prompt = PrivacyConsentGate.Evaluate(PrivacyConsentGate.IsAccepted(), blockedAction)
        If prompt Is Nothing Then Return False
        ExecuteJavaScriptAsyncJS(ModelConfigGate.BuildShowScript(prompt))
        Return True
    End Function

    Protected ReadOnly Property AiNativeRuntime As Agent.IAiNativeRuntime
        Get
            If _aiNativeRuntime Is Nothing Then
                Dim toolRegistry As New Agent.ToolRegistry()
                toolRegistry.LoadFromRuntimeDirectories()
                toolRegistry.LoadSkillScriptsAsTools()
                _aiNativeRuntime = New Agent.AiNativeRuntime(toolRegistry)
            End If
            Return _aiNativeRuntime
        End Get
    End Property

    Protected ReadOnly Property ConversationRuntime As IConversationRuntime
        Get
            If _conversationRuntime Is Nothing Then
                _conversationRuntime = New DefaultConversationRuntime(
                    New DefaultContextComposer(),
                    New McpToolBroker())
            End If
            Return _conversationRuntime
        End Get
    End Property

    Private ReadOnly Property ChatRequestOrchestrator As ChatRequestOrchestrator
        Get
            If _chatRequestOrchestrator Is Nothing Then
                _chatRequestOrchestrator = New ChatRequestOrchestrator(
                    ConversationRuntime,
                    systemHistoryMessageData,
                    _selectionPendingMap,
                    Function() RunUiFuncSync(Function() GetApplication()),
                    AddressOf ManageHistoryMessageSize)
            End If
            Return _chatRequestOrchestrator
        End Get
    End Property

    Private ReadOnly Property ChatRoutingOrchestrator As ChatRoutingOrchestrator
        Get
            If _chatRoutingOrchestrator Is Nothing Then
                _chatRoutingOrchestrator = New ChatRoutingOrchestrator(Me)
            End If
            Return _chatRoutingOrchestrator
        End Get
    End Property

    ' 延迟初始化的排版服务
    Private _reformatService As ReformatService = Nothing
    Protected ReadOnly Property ReformatSvc As ReformatService
        Get
            If _reformatService Is Nothing Then
                _reformatService = New ReformatService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf RunUiAction,
                    AddressOf ShowTemplateEditorPane,
                    AddressOf GetStylePreviewCallback,
                    AddressOf HandleUploadDocxTemplateFromPath,
                    Function(mode, prompt) SendAndGetResponseAsync(prompt, "", New List(Of HistoryMessage)()))
            End If
            Return _reformatService
        End Get
    End Property

    ' 延迟初始化的统一 AgentKernel 服务（新架构）
    Private _agentKernelService As AgentKernelService = Nothing
    Protected ReadOnly Property AgentKernelSvc As AgentKernelService
        Get
            If _agentKernelService Is Nothing Then
                _agentKernelService = New AgentKernelService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf SendAndGetResponseAsync,
                    Function(c, l, p) CodeExecutionService.ExecuteCodeWithToolResult(c, l, p),
                    _chatStateService,
                    systemHistoryMessageData,
                    AddressOf ManageHistoryMessageSize,
                    AddressOf GetOfficeAppType)
            End If
            Return _agentKernelService
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

    ' 延迟初始化的自动补全服务
    Private _autocompleteService As AutocompleteService = Nothing
    Protected ReadOnly Property AutocompleteSvc As AutocompleteService
        Get
            If _autocompleteService Is Nothing Then
                _autocompleteService = New AutocompleteService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf GetContextSnapshot,
                    AddressOf GetOfficeAppType)
            End If
            Return _autocompleteService
        End Get
    End Property

    ' 延迟初始化的历史会话服务
    Private _historySessionService As HistorySessionService = Nothing
    Protected ReadOnly Property HistorySessionSvc As HistorySessionService
        Get
            If _historySessionService Is Nothing Then
                _historySessionService = New HistorySessionService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    _chatStateService,
                    AddressOf GetOfficeAppType,
                    AddressOf ResolveSessionDocument,
                    AddressOf RunUiAction)
            End If
            Return _historySessionService
        End Get
    End Property

    ' 延迟初始化的首次引导服务（E3-01）
    Private _onboardingService As OnboardingService = Nothing
    Protected ReadOnly Property OnboardingSvc As OnboardingService
        Get
            If _onboardingService Is Nothing Then
                _onboardingService = New OnboardingService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf GetOfficeAppType,
                    AddressOf UpdateModelDisplayInUI)
            End If
            Return _onboardingService
        End Get
    End Property

    ' 延迟初始化的命令面板服务（E4-03）：快捷键监听跟着面板一起生灭，
    ' 面板销毁后还留着轮询就等于快捷键在用户以为已经关掉插件的情况下继续生效。
    Private _commandPaletteService As CommandPaletteService = Nothing
    Protected ReadOnly Property CommandPaletteSvc As CommandPaletteService
        Get
            If _commandPaletteService Is Nothing Then
                _commandPaletteService = New CommandPaletteService(
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf GetOfficeAppType,
                    AddressOf EnsureHostTaskPaneVisible)
                AddHandler Me.Disposed, Sub(sender As Object, args As EventArgs) _commandPaletteService.Shutdown()
            End If
            Return _commandPaletteService
        End Get
    End Property

    ''' <summary>
    ''' 快捷键很可能在任务窗格收起时被按下，这时动作清单没有落点，宿主必须先把窗格展开。
    ''' 基类拿不到 CustomTaskPane，所以由三端各自的 ChatControl 重写；默认什么都不做表示面板已经可见。
    ''' </summary>
    Protected Overridable Sub EnsureHostTaskPaneVisible()
    End Sub

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
                    Function(vbaCode As String, preview As Boolean) As Boolean
                        Dim approved = RunCodePreview(vbaCode, preview)
                        If Not approved Then NotifyPreviewCancelled()
                        Return approved
                    End Function,
                    AddressOf EvaluateFormula)
                ' 脚本预览拒绝走的是 Boolean 合同（取消≠错误），ToolResult 侧必须能读到同一标记
                _codeExecutionService.PreviewCancelledReader = AddressOf ConsumePreviewCancelled
                ' Agent/Loop 可在后台线程运行。所有 native JSON 命令在唯一宿主入口
                ' 统一切回 UI/STA 线程，避免各 Office Executor 自行分散 Invoke。
                _codeExecutionService.JsonCommandExecutorWithResult =
                    Function(jsonCode As String, preview As Boolean) As Agent.ToolResult
                        Return UiDispatcher.InvokeSync(Of Agent.ToolResult)(
                            Me,
                            Function() ExecuteJsonCommandWithToolResult(jsonCode, preview))
                    End Function
            End If
            Return _codeExecutionService
        End Get
    End Property

    ' 延迟初始化的自检Loop控制器
    Protected ReadOnly Property SelfCheckLoopController As SelfCheckLoopController
        Get
            If _selfCheckLoopController Is Nothing Then
                _selfCheckLoopController = New SelfCheckLoopController(
                    New PreSendChecker(),
                    New PostFlushValidator())
            End If
            Return _selfCheckLoopController
        End Get
    End Property

    ' 延迟初始化的 HTTP 流服务
    Private _httpStreamService As HttpStreamService = Nothing
    Protected ReadOnly Property HttpStreamSvc As HttpStreamService
        Get
            If _httpStreamService Is Nothing Then
                _httpStreamService = New HttpStreamService(
                    _chatStateService,
                    AddressOf GetApplication,
                    AddressOf ExecuteJavaScriptAsyncJS,
                    AddressOf WaitForRendererMapAsync,
                    Sub(uuid, plainBuf) CheckAndCompleteProcessingHook(uuid, plainBuf),
                    Nothing,
                    Nothing)
            End If
            Return _httpStreamService
        End Get
    End Property

    ''' <summary>
    ''' 当前流的最终响应 UUID（代理到 HttpStreamSvc，供子类访问）
    ''' </summary>
    Protected ReadOnly Property _finalUuid As String
        Get
            Return HttpStreamSvc.FinalUuid
        End Get
    End Property

    ''' <summary>
    ''' Host JSON command backend used by CodeExecutionService / Agent tools (P0-2).
    ''' Not a product entry for natural-language routing; NL goes ChatRouting → AgentKernel.
    ''' </summary>
    Protected Overridable Function ExecuteJsonCommandWithToolResult(jsonCode As String, preview As Boolean) As Agent.ToolResult
        Return Agent.ToolResult.Failed("",
                                       "当前应用不支持 JSON 命令执行",
                                       errorCode:=ExceptionClassifier.CodeHostUnsupported,
                                       userMessage:="当前应用不支持 JSON 命令执行",
                                       recoverable:=False)
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

    'settings
    Protected topicRandomness As Double
    Protected contextLimit As Integer
    ' E4-04：选区自动引用默认开启。开关在面板设置里可见可关，缺键时保持这个默认值。
    Protected selectedCellChecked As Boolean = True
    Protected settingsScrollChecked As Boolean = False

    Protected stopReaderStream As Boolean = False


    ' ai的历史回复
    Protected systemHistoryMessageData As New List(Of HistoryMessage)

    ' 输入框入口登记的"本回合用户可见文本"，由执行漏斗认领后落库并置空，保证每回合只写一条 user 历史
    Private _pendingUserVisibleText As String

    ' E4-04：本轮真的带进提示的选区引用摘要（Nothing = 没带）。只用于在对话里显示，不落历史、不进日志。
    Private _autoSelectionReference As String

    Protected loadingPictureBox As PictureBox

    ' 选区对比相关字段
    Protected PendingSelectionInfo As SelectionInfo = Nothing

    ' 以下属性代理到 _chatStateService 的内部字典，保持子类兼容性
    Protected ReadOnly Property _selectionPendingMap As Dictionary(Of String, SelectionInfo)
        Get
            Return _chatStateService.SelectionPendingMap
        End Get
    End Property

    Private ReadOnly Property allPlainMarkdownBuffer As StringBuilder
        Get
            Return _chatStateService.PlainMarkdownBuffer
        End Get
    End Property

    Protected ReadOnly Property _responseToRequestMap As Dictionary(Of String, String)
        Get
            Return _chatStateService.ResponseToRequestMap
        End Get
    End Property

    Protected _revisionsMap As New Dictionary(Of String, JArray)()

    Private Async Sub RunUiAction(action As Action)
        Try
            Await UiDispatcher.InvokeAsync(Me, action)
        Catch ex As Exception
            Debug.WriteLine($"[BaseChatControl] UI action dispatch failed: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Background → UI marshal without async GetResult (P0-3). Prefer RunUiAction/async when possible.
    ''' </summary>
    Protected Sub RunUiActionSync(action As Action)
        If action Is Nothing Then Return
        UiDispatcher.InvokeSync(Me, action)
    End Sub

    ''' <summary>
    ''' 读取宿主对象模型必须在 UI/STA 线程：Office COM 是套间绑定的，线程池直读要跨套间排队，
    ''' 既慢又可能死锁。这里只允许包住纯读取，网络与记忆检索留在后台线程。
    ''' </summary>
    Protected Function RunUiFuncSync(Of T)(reader As Func(Of T)) As T
        If reader Is Nothing Then Return Nothing
        Return UiDispatcher.InvokeSync(Me, reader)
    End Function

    ''' <summary>
    ''' 预览被用户拒绝时登记。旧的 Boolean 执行链把取消当“不是错误”返回 True（UI 只需恢复按钮），
    ''' 但 ToolResult 合同不能据此回执成功，所以宿主在布尔取消点登记、在 ToolResult 入口消费一次。
    ''' </summary>
    Private _previewCancelledByUser As Boolean

    Protected Sub NotifyPreviewCancelled()
        _previewCancelledByUser = True
    End Sub

    Protected Function ConsumePreviewCancelled() As Boolean
        Dim cancelled = _previewCancelledByUser
        _previewCancelledByUser = False
        Return cancelled
    End Function

    Protected Async Function InitializeWebView2() As Task
        ' 确保在 UI 线程上执行
        If ChatBrowser.InvokeRequired Then
            Await UiDispatcher.InvokeAsync(ChatBrowser,
                Async Function()
                    Await InitializeWebView2()
                End Function)
            Return
        End If

        ' 会话持久化需要宿主类型；此时子类已完成构造，取虚方法安全
        If _chatStateService.AppTypeProvider Is Nothing Then
            _chatStateService.AppTypeProvider = AddressOf GetOfficeAppType
        End If
        RefreshSessionDocumentContext()

        Await _webViewInitSemaphore.WaitAsync()
        Try
            ' 安全检查 CoreWebView2（必须在 UI 线程）
            Dim isInitialized As Boolean = False
            Try
                isInitialized = (ChatBrowser.CoreWebView2 IsNot Nothing)
            Catch ex As InvalidOperationException
                ' CoreWebView2 只能在 UI 线程访问，如果出现此异常，重新调度到 UI 线程
                Debug.WriteLine("[InitializeWebView2] 检测到线程问题，重新调度到 UI 线程")
                _webViewInitSemaphore.Release()

                ' 使用 Control.Invoke 强制在 UI 线程执行
                If ChatBrowser.InvokeRequired Then
                    ChatBrowser.Invoke(New Action(Async Sub()
                        Await InitializeWebView2()
                    End Sub))
                    Return
                End If
                Throw
            End Try

            If isInitialized Then
                Return
            End If

            Dim userDataFolder As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyAppWebView2Cache")
            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            Dim wwwRoot As String = ResourceExtractor.ExtractResources()
            
            ' 检查资源提取是否成功
            If String.IsNullOrEmpty(wwwRoot) Then
                Dim errMsg As String = "资源提取失败，无法初始化聊天界面。" & Environment.NewLine
                If Not String.IsNullOrEmpty(ResourceExtractor.LastError) Then
                    errMsg &= "错误详情: " & ResourceExtractor.LastError
                End If
                MessageBox.Show(errMsg, "初始化错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            
            ' 检查目录是否存在
            If Not Directory.Exists(wwwRoot) Then
                MessageBox.Show($"资源目录不存在: {wwwRoot}", "初始化错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
            
            ' 诊断：显示资源目录
            Debug.WriteLine($"[WebView2] wwwRoot = {wwwRoot}")

            ChatBrowser.CreationProperties = New CoreWebView2CreationProperties With {
                .UserDataFolder = userDataFolder
            }
            Dim env = Await WebView2EnvironmentCache.GetOrCreateAsync(userDataFolder)
            Await ChatBrowser.EnsureCoreWebView2Async(env)

            If ChatBrowser.CoreWebView2 IsNot Nothing Then
                ChatBrowser.CoreWebView2.Settings.IsScriptEnabled = True
                ChatBrowser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                ChatBrowser.CoreWebView2.Settings.IsWebMessageEnabled = True
#If DEBUG Then
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = True
#Else
                ChatBrowser.CoreWebView2.Settings.AreDevToolsEnabled = False
#End If

                ' 设置虚拟主机映射
                Try
                    ChatBrowser.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "officeai.local",
                        wwwRoot,
                        CoreWebView2HostResourceAccessKind.Allow
                    )
                    Debug.WriteLine($"[WebView2] 虚拟主机映射成功: officeai.local -> {wwwRoot}")
                Catch ex As Exception
                    Debug.WriteLine($"[WebView2] 虚拟主机映射失败: {ex.Message}")
                    MessageBox.Show($"虚拟主机映射失败: {ex.Message}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Try

                ' 添加 WebResourceRequested 事件处理作为备用方案
                AddHandler ChatBrowser.CoreWebView2.WebResourceRequested, AddressOf OnWebResourceRequested
                ChatBrowser.CoreWebView2.AddWebResourceRequestedFilter("http://officeai.local/*", CoreWebView2WebResourceContext.Script)
                ChatBrowser.CoreWebView2.AddWebResourceRequestedFilter("http://officeai.local/*", CoreWebView2WebResourceContext.Stylesheet)
                ChatBrowser.CoreWebView2.AddWebResourceRequestedFilter("http://officeai.local/*", CoreWebView2WebResourceContext.Image)
                
                Dim htmlContent As String = My.Resources.chat_template_refactored
                ' E4-05：这是一次新的文档装载，页面侧的初始化要重做一遍；除此之外不重复注入。
                _pageBootstrapDone = False
                ChatBrowser.CoreWebView2.NavigateToString(htmlContent)

                ' E4-05：这里原来还 AddHandler 了一次 CoreWebView2.NavigationCompleted，
                ' 而 OnWebViewNavigationCompleted 已经用 Handles ChatBrowser.NavigationCompleted 绑过一次，
                ' 同一次导航因此跑两遍；下面的 RemoveHandler 只摘得掉这一份，Handles 那份永远留着。
                ' 现在只保留声明式那一份绑定，配合 _pageBootstrapDone 做到「一次装载一次初始化」。
                
                ' 配置 Markdown 解析器
                Await ConfigureMarked()
                
                ' 设置应用名称的延迟调用
                Await Task.Delay(500) ' 等待一点时间确保页面加载
                Await SetCurrentOfficeAppName()

            Else
                MessageBox.Show("WebView2 初始化失败，CoreWebView2 不可用。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            Dim errorMessage As String = $"初始化失败: {ex.Message}{Environment.NewLine}类型: {ex.GetType().Name}{Environment.NewLine}堆栈:{ex.StackTrace}"
            MessageBox.Show(errorMessage, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _webViewInitSemaphore.Release()
        End Try
    End Function
    
    ''' <summary>
    ''' WebResourceRequested 事件处理 - 作为虚拟主机映射的备用方案
    ''' </summary>
    Private Sub OnWebResourceRequested(sender As Object, e As CoreWebView2WebResourceRequestedEventArgs)
        Try
            Dim requestUri As Uri = New Uri(e.Request.Uri)
            Dim localPath As String = requestUri.AbsolutePath.TrimStart("/"c)
            
            ' 映射到本地资源目录
            Dim wwwRoot As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OfficeAI", "www"
            )
            Dim filePath As String = Path.Combine(wwwRoot, localPath.Replace("/"c, Path.DirectorySeparatorChar))
            
            If File.Exists(filePath) Then
                Dim mimeType As String = GetMimeType(filePath)
                Dim fileBytes As Byte() = File.ReadAllBytes(filePath)
                
                Dim response = ChatBrowser.CoreWebView2.Environment.CreateWebResourceResponse(
                    New MemoryStream(fileBytes),
                    200,
                    "OK",
                    $"Content-Type: {mimeType}{Environment.NewLine}Access-Control-Allow-Origin: *"
                )
                e.Response = response
                Debug.WriteLine($"[WebView2] 资源加载成功: {localPath}")
            Else
                Debug.WriteLine($"[WebView2] 资源未找到: {filePath}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[WebView2] WebResourceRequested 错误: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 获取文件的 MIME 类型
    ''' </summary>
    Private Function GetMimeType(filePath As String) As String
        Dim ext As String = Path.GetExtension(filePath).ToLower()
        Select Case ext
            Case ".js" : Return "application/javascript"
            Case ".css" : Return "text/css"
            Case ".html" : Return "text/html"
            Case ".png" : Return "image/png"
            Case ".jpg", ".jpeg" : Return "image/jpeg"
            Case ".gif" : Return "image/gif"
            Case ".svg" : Return "image/svg+xml"
            Case ".json" : Return "application/json"
            Case Else : Return "application/octet-stream"
        End Select
    End Function

    ''' <summary>
    ''' 设置当前 Office 应用名称
    ''' </summary>
    Private Async Function SetCurrentOfficeAppName() As Task
        Try
            ' 获取当前应用名称
            Dim appName As String = GetOfficeApplicationName()
            
            ' 向网页注入应用名称
            Dim script As String = $"window.currentOfficeAppName = '{appName}';"
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
            
        Catch ex As Exception
            Debug.WriteLine($"设置应用名称失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 获取当前 Office 应用程序名称
    ''' </summary>
    Protected Overridable Function GetOfficeApplicationName() As String
        ' 默认返回 "当前应用"，子类应重写此方法
        Return "当前应用"
    End Function
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
                Dim questionPrefix As String = UtilsService.GetFirst10Characters(firstQuestion)
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}_{questionPrefix}.html"
            Else
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
            End If

            _chatHtmlFilePath = Path.Combine(baseDir, fileName)
            Return _chatHtmlFilePath
        End Get
    End Property

    Private Async Sub OnWebViewNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs) Handles ChatBrowser.NavigationCompleted
        If Not e.IsSuccess Then Return
        If _pageBootstrapDone Then Return
        _pageBootstrapDone = True
        Try
            ' Async Sub中Await后VSTO同步上下文可能丢失，
            ' 统一通过Invoke确保CoreWebView2访问在UI线程
            Await Task.Delay(100)
            Await UiDispatcher.InvokeAsync(ChatBrowser,
                Sub()
                    InitializeSettings()
                    InitializeMcpSettings()
                End Sub)
        Catch ex As Exception
            Debug.WriteLine($"导航完成事件处理中出错: {ex.Message}")
            Debug.WriteLine(ex.StackTrace)
        End Try
    End Sub

    Protected Sub InitializeSettings()
        Try
            ' 确保记忆功能配置是开启的（默认）
            Debug.WriteLine($"[InitializeSettings] 初始化记忆配置...")
            Debug.WriteLine($"[InitializeSettings] UseContextBuilder: {MemoryConfig.UseContextBuilder}")
            Debug.WriteLine($"[InitializeSettings] RagTopN: {MemoryConfig.RagTopN}")
            Debug.WriteLine($"[InitializeSettings] EnableUserProfile: {MemoryConfig.EnableUserProfile}")
            
            ' 如果 UseContextBuilder 是关闭的，我们强制开启它
            If Not MemoryConfig.UseContextBuilder Then
                Debug.WriteLine($"[InitializeSettings] UseContextBuilder 未开启，强制开启...")
                MemoryConfig.UseContextBuilder = True
            End If
            
            ' 确保 RagTopN 至少是 5
            If MemoryConfig.RagTopN < 3 Then
                MemoryConfig.RagTopN = 5
            End If
            
            Debug.WriteLine($"[InitializeSettings] 记忆配置已确保开启: UseContextBuilder={MemoryConfig.UseContextBuilder}, RagTopN={MemoryConfig.RagTopN}")

            ' 加载设置
            Dim chatSettings As New ChatSettings(GetApplication())
            selectedCellChecked = ChatSettings.selectedCellChecked
            contextLimit = ChatSettings.contextLimit
            topicRandomness = ChatSettings.topicRandomness
            settingsScrollChecked = ChatSettings.settingsScrollChecked

            ' 设置Office应用类型（用于前端区分Word/PPT/Excel）
            Dim appType = GetOfficeAppType()
            Dim paletteTitle As String = If(ChatSettings.EnableHotkey,
                                            "命令面板（" & ChatSettings.HotkeyChord & "）", "命令面板")

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
            
            // E4-01/E4-02/E4-03：三个原生入口开关在设置面板里可见，用户不改代码就能关掉它们
            setNativeEntrySetting('settings-enable-hotkey', {ChatSettings.EnableHotkey.ToString().ToLower()});
            setNativeEntrySetting('settings-enable-selection-toolbar', {ChatSettings.EnableSelectionToolbar.ToString().ToLower()});
            setNativeEntrySetting('settings-enable-context-menu', {ChatSettings.EnableContextMenu.ToString().ToLower()});
            setNativeEntrySetting('settings-hotkey-chord', {Newtonsoft.Json.JsonConvert.SerializeObject(ChatSettings.HotkeyChord)});
            setNativeEntrySetting('command-palette-button', null, {Newtonsoft.Json.JsonConvert.SerializeObject(paletteTitle)});
        "
            ExecuteJavaScriptAsyncJS(js)

            ' E4-03：页面这时候才真的有 window.syncCommandPalette，清单和快捷键状态在这里落地一次；
            ' 设置里改过开关的话，ApplyHotkey 会在下一次保存时再同步，不需要重启宿主。
            CommandPaletteSvc.PushState()
            CommandPaletteSvc.ApplyHotkey()
        Catch ex As Exception
            Debug.WriteLine($"初始化设置失败: {ex.Message}")
        End Try
    End Sub

    Private Sub RegisterWebViewCommandHandlers(router As ChatCommandRouter)
        router.Register("sendMessage", Sub(jsonDoc) HandleSendMessage(jsonDoc))
        router.Register("stopMessage", Sub(jsonDoc) HandleStopMessage(jsonDoc))
        router.Register("executeCode", Sub(jsonDoc) HandleExecuteCode(jsonDoc))
        router.Register("saveSettings", Sub(jsonDoc) HandleSaveSettings(jsonDoc))
        router.Register("getSessionList", Sub(jsonDoc) HandleGetSessionList(ReadSessionScope(jsonDoc)))
        router.Register("loadSession", Sub(jsonDoc) HandleLoadSession(jsonDoc))
        router.Register("newSession", Sub(jsonDoc) HandleNewSession())
        router.Register("deleteSession", Sub(jsonDoc) HandleDeleteSession(jsonDoc, ReadSessionScope(jsonDoc)))
        router.Register("renameSession", Sub(jsonDoc) HandleRenameSession(jsonDoc, ReadSessionScope(jsonDoc)))
        router.Register("getMcpConnections", Sub(jsonDoc) HandleGetMcpConnections())
        router.Register("saveMcpSettings", Sub(jsonDoc) HandleSaveMcpSettings(jsonDoc))
        router.Register("clearContext", Sub(jsonDoc) ClearChatContext())
        router.Register("acceptAnswer", Sub(jsonDoc) HandleAcceptAnswer(jsonDoc))
        router.Register("rejectAnswer", Sub(jsonDoc) HandleRejectAnswer(jsonDoc))
        router.Register("applyRevisionSegment", Sub(jsonDoc) HandleApplyRevisionSegment(jsonDoc))
        router.Register("applyDocumentPlanItem", Sub(jsonDoc) HandleApplyDocumentPlanItem(jsonDoc))
        router.Register("retryReformat", Sub(jsonDoc) HandleRetryReformat(jsonDoc))
        router.Register("applyRevisionAccept", Sub(jsonDoc) HandleApplyRevisionAccept(jsonDoc))
        router.Register("triggerContinuation", Sub(jsonDoc) HandleTriggerContinuation(jsonDoc))
        router.Register("applyContinuation", Sub(jsonDoc) HandleApplyContinuation(jsonDoc))
        router.Register("refineContinuation", Sub(jsonDoc) HandleRefineContinuation(jsonDoc))
        router.Register("applyTemplateContent", Sub(jsonDoc) HandleApplyTemplateContent(jsonDoc))
        router.Register("refineTemplateContent", Sub(jsonDoc) HandleRefineTemplateContent(jsonDoc))
        router.Register("requestCompletion", Sub(jsonDoc) HandleRequestCompletion(jsonDoc))
        router.Register("acceptCompletion", Sub(jsonDoc) HandleAcceptCompletion(jsonDoc))
        router.Register("startAgent", Sub(jsonDoc) HandleStartAgent(jsonDoc))
        router.Register("abortAgent", Sub(jsonDoc) HandleAbortAgent())
        ' E5-04 删掉了 7 条没有任何发送者的注册：startLoop/continueLoop/replanLoop/cancelLoop（Ralph 别名，
        ' /loop 已在 message-sender.js 改写成 startAgent）、startAgentExecution/refineAgentPlan（被 agent: 前缀
        ' 那套取代）、checkedChange（旧版 BaseChat 面板有自己的 Select Case，不经过这个 router）。
        ' 保留一条注册等于保留一条走不到、也验证不了的路径；由 scripts/check-entry-reachability.js 钉住。
        router.Register("agent:approvePlan", Sub(jsonDoc) HandleAgentApprovePlan(jsonDoc))
        router.Register("agent:rejectPlan", Sub(jsonDoc) HandleAgentRejectPlan(jsonDoc))
        router.Register("agent:approveStep", Sub(jsonDoc) HandleAgentApproveStep(jsonDoc))
        router.Register("agent:approveToolSession", Sub(jsonDoc) HandleAgentApproveToolSession(jsonDoc))
        router.Register("agent:refinePlan", Sub(jsonDoc) HandleAgentRefinePlan(jsonDoc))
        router.Register("agent:undoRun", Sub(jsonDoc) HandleAgentUndoRun(jsonDoc))
        router.Register("openFileDialog", Sub(jsonDoc) HandleOpenFileDialog())
        router.Register("openApiConfigForm", Sub(jsonDoc) HandleOpenApiConfigForm())
        ' E3-01 首次引导：只有"隐私确认 + 示例任务"两步，配置动作走 openApiConfigForm
        router.Register("getOnboardingState", Sub(jsonDoc) OnboardingSvc.HandleGetState())
        router.Register("onboardingSetPrivacy", Sub(jsonDoc) OnboardingSvc.HandleSetPrivacy(jsonDoc))
        router.Register("onboardingFinish", Sub(jsonDoc) OnboardingSvc.HandleFinish())
        router.Register("getCurrentAppInfo", Sub(jsonDoc) HandleGetCurrentAppInfo())
        router.Register("getCurrentModel", Sub(jsonDoc) HandleGetCurrentModel())
        router.Register("getReformatTemplates", Sub(jsonDoc) HandleGetReformatTemplates())
        router.Register("useReformatTemplate", Sub(jsonDoc) HandleUseReformatTemplate(jsonDoc))
        router.Register("previewTemplateInWord", Sub(jsonDoc) HandlePreviewTemplateInWord(jsonDoc))
        router.Register("saveCurrentDocumentAsTemplate", Sub(jsonDoc) HandleSaveCurrentDocumentAsTemplate())
        router.Register("importTemplate", Sub(jsonDoc) HandleImportTemplate())
        router.Register("exportTemplate", Sub(jsonDoc) HandleExportTemplate(jsonDoc))
        router.Register("duplicateTemplate", Sub(jsonDoc) HandleDuplicateTemplate(jsonDoc))
        router.Register("deleteTemplate", Sub(jsonDoc) HandleDeleteTemplate(jsonDoc))
        router.Register("openTemplateEditor", Sub(jsonDoc) HandleOpenTemplateEditor(jsonDoc))
        router.Register("getStyleGuides", Sub(jsonDoc) HandleGetStyleGuides())
        router.Register("useStyleGuide", Sub(jsonDoc) HandleUseStyleGuide(jsonDoc))
        router.Register("uploadStyleGuideDocument", Sub(jsonDoc) HandleUploadStyleGuideDocument())
        router.Register("deleteStyleGuide", Sub(jsonDoc) HandleDeleteStyleGuide(jsonDoc))
        router.Register("updateStyleGuide", Sub(jsonDoc) HandleUpdateStyleGuide(jsonDoc))
        router.Register("duplicateStyleGuide", Sub(jsonDoc) HandleDuplicateStyleGuide(jsonDoc))
        router.Register("exportStyleGuide", Sub(jsonDoc) HandleExportStyleGuide(jsonDoc))
        router.Register("uploadDocxTemplate", Sub(jsonDoc) HandleUploadDocxTemplate())
        router.Register("deleteDocxMapping", Sub(jsonDoc) HandleDeleteDocxMapping(jsonDoc))
        router.Register("saveAiTemplate", Sub(jsonDoc) HandleSaveAiTemplate(jsonDoc))
        router.Register("previewAiTemplate", Sub(jsonDoc) HandlePreviewAiTemplate(jsonDoc))
        router.Register("applySmartReformat", Sub(jsonDoc) HandleApplySmartReformat(jsonDoc))
        router.Register("undoReformat", Sub(jsonDoc) HandleUndoReformat(jsonDoc))
        router.Register("refineSmartReformat", Sub(jsonDoc) HandleRefineSmartReformat(jsonDoc))
        router.Register("switchReformatTemplate", Sub(jsonDoc) HandleSwitchReformatTemplate(jsonDoc))
        router.Register("previewReformatCompare", Sub(jsonDoc) HandlePreviewReformatCompare(jsonDoc))
        router.Register("proofread", Sub(jsonDoc) HandleProofreadFocusMode(jsonDoc))
    End Sub

    Private Sub HandleStopMessage(jsonDoc As JObject)
        stopReaderStream = True

        Dim requestUuid As String = ""
        Try
            requestUuid = If(jsonDoc("requestUuid")?.ToString(), "")
        Catch
            requestUuid = ""
        End Try

        HttpStreamSvc.CancelRequest(requestUuid)
    End Sub

    Protected Sub WebView2_WebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim rawJson As String = e.WebMessageAsJson
            Dim jsonDoc As JObject = JObject.Parse(rawJson)
            CommandRouter.Dispatch(jsonDoc)
            Return

        Catch ex As Exception
            Debug.WriteLine($"[DEBUG WebMessageReceived] StackTrace: {ex.StackTrace}")
        End Try
    End Sub

    ' 添加：在基类提供可覆盖的 CaptureCurrentSelectionInfo（默认返回 Nothing，Word 子类会覆写）
    Protected Overridable Function CaptureCurrentSelectionInfo(mode As String) As SelectionInfo
        Return Nothing
    End Function


    Protected Overridable Sub HandleApplyRevisionSegment(jsonDoc As JObject)
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

    ''' <summary>
    ''' 构建校对模式追问的系统提示词
    ''' 当用户在校对模式下发送消息时，让AI理解当前是校对上下文
    ''' </summary>
    Private Function BuildProofreadFollowUpPrompt(selectedText As String, issueCount As String) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("你是专业的中文文档校对专家。用户当前正在使用校对模式，已选中一段文本进行校对。")
        sb.AppendLine()
        sb.AppendLine("【当前校对上下文】")

        If Not String.IsNullOrWhiteSpace(selectedText) Then
            sb.AppendLine("用户选中的文本内容（节选）：")
            sb.AppendLine(selectedText)
            sb.AppendLine()
        End If

        Dim count As Integer
        If Integer.TryParse(issueCount, count) AndAlso count > 0 Then
            sb.AppendLine($"AI已检测出 {count} 处校对问题，问题列表显示在右侧校对面板中。")
            sb.AppendLine()
        End If

        sb.AppendLine("【交互规则】")
        sb.AppendLine("1. 用户的消息是在校对模式下发送的，可能是对校对结果的追问、请求解释、或要求重新检查")
        sb.AppendLine("2. 如果用户问的是关于某个具体校对问题，请基于选中内容的上下文给出专业解释")
        sb.AppendLine("3. 如果用户要求补充检查某些方面（如专有名词、数据一致性），请给出针对性建议")
        sb.AppendLine("4. 回答要简洁专业，不要重复列出校对面板已显示的问题")
        sb.AppendLine("5. 严禁使用Markdown代码块格式（禁止使用```符号），直接输出纯文本")

        Return sb.ToString()
    End Function

    ' ========== 模板渲染功能相关方法结束 ==========

    ' ========== 自动补全功能相关方法 ==========

    ''' <summary>
    ''' 处理自动补全请求
    ''' </summary>
    Protected Overridable Async Sub HandleRequestCompletion(jsonDoc As JObject)
        Await AutocompleteSvc.HandleRequestCompletion(jsonDoc)
    End Sub

    ''' <summary>
    ''' 处理补全采纳记录
    ''' </summary>
    Protected Overridable Sub HandleAcceptCompletion(jsonDoc As JObject)
        AutocompleteSvc.HandleAcceptCompletion(jsonDoc)
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
    ''' 为意图识别/规划阶段丰富上下文：内容区引用摘要 + RAG 相关记忆
    ''' </summary>
    Protected Sub EnrichContextForIntent(snapshot As JObject,
                                        question As String,
                                        filePaths As List(Of String),
                                        selectedContents As List(Of SendMessageReferenceContentItem))
        AutocompleteSvc.EnrichContextForIntent(snapshot, question, filePaths, selectedContents)
    End Sub

    ' ========== 已移入 AutocompleteService ==========
    ' RequestCompletionsFromLLM, RequestCompletionsWithFIM, RequestCompletionsWithChat
    ' RecordCompletionHistory, GetCompletionSystemPrompt


    ' ========== 自动补全功能相关方法结束 ==========

    ' 新增：处理用户接受答案（收藏回答时更新 conversation.is_collected）
    Protected Sub HandleAcceptAnswer(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)

            Debug.WriteLine($"用户接受回答: UUID={uuid}")
            GlobalStatusStrip.ShowInfo("用户已接受 AI 回答")

            ' 更新 conversation 表收藏状态
            Dim sid = _chatStateService.CurrentSessionId
            If Not String.IsNullOrEmpty(sid) Then
                Try
                    ConversationRepository.SetLastAssistantCollected(sid, True)
                Catch ex As Exception
                    Debug.WriteLine($"SetLastAssistantCollected 失败: {ex.Message}")
                End Try
            End If
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
        _chatStateService.StartNewSession()
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

    Protected Sub HandleGetSessionList(Optional scope As String = ConversationRepository.SessionScopeAll)
        HistorySessionSvc.HandleGetSessionList(scope)
    End Sub

    ''' <summary>
    ''' 历史侧边栏的可见范围由前端带上；缺省按"全部（本文档优先）"处理。
    ''' </summary>
    Private Shared Function ReadSessionScope(jsonDoc As JObject) As String
        Dim scope = jsonDoc?("scope")?.ToString()
        If String.Equals((If(scope, "")).Trim(), ConversationRepository.SessionScopeDocument, StringComparison.OrdinalIgnoreCase) Then
            Return ConversationRepository.SessionScopeDocument
        End If
        Return ConversationRepository.SessionScopeAll
    End Function

    Protected Sub HandleLoadSession(jsonDoc As JObject)
        HistorySessionSvc.HandleLoadSession(jsonDoc)
    End Sub

    Protected Sub HandleNewSession()
        HistorySessionSvc.HandleNewSession()
    End Sub

    Protected Sub HandleDeleteSession(jsonDoc As JObject, scope As String)
        HistorySessionSvc.HandleDeleteSession(jsonDoc, scope)
    End Sub

    Protected Sub HandleRenameSession(jsonDoc As JObject, scope As String)
        HistorySessionSvc.HandleRenameSession(jsonDoc, scope)
    End Sub

    Protected Overridable Sub HandleSaveSettings(jsonDoc As JObject)
        topicRandomness = jsonDoc("topicRandomness")
        contextLimit = jsonDoc("contextLimit")
        selectedCellChecked = jsonDoc("selectedCell")
        settingsScrollChecked = jsonDoc("settingsScroll")
        Dim executeCodePreview As Boolean = jsonDoc("executeCodePreview")
        Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
        Dim autocompleteShortcut As String = If(jsonDoc("autocompleteShortcut")?.Value(Of String)(), "Ctrl+.")
        ' 原生入口的三个开关由面板自己带上来。这里以前回写的是静态值，等于面板上的开关存不住；
        ' 缺键时退回当前值，避免旧版前端一保存就把用户已经存好的设置抹平。
        Dim enableSelectionToolbar As Boolean = ReadFlagSetting(jsonDoc, "enableSelectionToolbar", ChatSettings.EnableSelectionToolbar)
        Dim enableContextMenu As Boolean = ReadFlagSetting(jsonDoc, "enableContextMenu", ChatSettings.EnableContextMenu)
        Dim enableHotkey As Boolean = ReadFlagSetting(jsonDoc, "enableHotkey", ChatSettings.EnableHotkey)
        Dim hotkeyChord As String = If(jsonDoc("hotkeyChord")?.Value(Of String)(), ChatSettings.HotkeyChord)

        ' 快捷键是独占资源：写法无效还照存，下次启动就是一个按了没反应、又没人知道为什么的开关
        Dim parsedChord = HostHotkeyChord.Parse(hotkeyChord)
        If enableHotkey AndAlso (parsedChord Is Nothing OrElse Not parsedChord.IsSafe) Then
            hotkeyChord = ChatSettings.HotkeyChord
            enableHotkey = False
            GlobalStatusStrip.ShowWarning("命令面板快捷键写法无效，已保持关闭：请使用带 Ctrl 或 Alt 的组合键，例如 Ctrl+Alt+K")
        End If

        Dim chatSettingsStore As New ChatSettings(GetApplication())
        ' 保存设置到配置文件（chatMode 已弃用，统一为智能模式）
        chatSettingsStore.SaveSettings(topicRandomness, contextLimit, selectedCellChecked,
                                  settingsScrollChecked, executeCodePreview,
                                  enableAutocomplete, 800, autocompleteShortcut,
                                  enableSelectionToolbar, enableContextMenu,
                                  enableHotkey, hotkeyChord)

        ' 保存完立刻生效：快捷键不该要求用户重启宿主才能试出它到底注册上了没有
        CommandPaletteSvc.ApplyHotkey()
    End Sub

    ''' <summary>设置面板传来的开关缺键时退回当前值，不当成「用户想关掉」。</summary>
    Private Shared Function ReadFlagSetting(jsonDoc As JObject, key As String, fallback As Boolean) As Boolean
        Dim token = jsonDoc(key)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return fallback
        Return token.Value(Of Boolean)()
    End Function

    ' SendMessageReferenceContentItem 已移至 Controls/Models/HistoryMessage.vb

    ' FileContentResult 类已移至 Controls/Models/HistoryMessage.vb


    ' 添加存储第一个问题的变量
    Protected firstQuestion As String = String.Empty
    Protected isFirstMessage As Boolean = True
    Private _chatHtmlFilePath As String = String.Empty ' 缓存文件路径



    ' 在 HandleSendMessage 方法中添加文件内容解析逻辑
    Protected Overridable Sub HandleSendMessage(jsonDoc As JObject)
        Debug.WriteLine($"[DEBUG BaseChatControl.HandleSendMessage] called, jsonDoc keys={String.Join(",", jsonDoc.Properties().Select(Function(p) p.Name))}")
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
        question = BuildMessageWithAutoSelection("--- 我此次的问题：" & question & " ---")

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
                         RunUiActionSync(Sub()
                                       currentWorkingDir = GetCurrentWorkingDirectory()
                                   End Sub)

                         For Each filePath As String In filePaths
                             Try
                                 processedFiles += 1

                                 ' 更新进度
                                 RunUiActionSync(Sub()
                                               GlobalStatusStrip.ShowInfo($"正在解析文件 ({processedFiles}/{totalFiles}): {Path.GetFileName(filePath)}")
                                               ExecuteJavaScriptAsyncJS($"updateFileParsingProgress({processedFiles}, {totalFiles}, {UtilsService.ToJsStringLiteral(Path.GetFileName(filePath))})")
                                           End Sub)

                                 ' 确定完整文件路径
                                 Dim fullFilePath As String = filePath

                                 ' 如果是绝对路径且文件存在，直接使用
                                 If Path.IsPathRooted(filePath) AndAlso File.Exists(filePath) Then
                                     fullFilePath = filePath
                                     Debug.WriteLine($"使用绝对路径: {fullFilePath}")
                                 ElseIf Not String.IsNullOrEmpty(currentWorkingDir) Then
                                     ' 尝试在当前工作目录下查找
                                     Dim tryPath = Path.Combine(currentWorkingDir, Path.GetFileName(filePath))
                                     If File.Exists(tryPath) Then
                                         fullFilePath = tryPath
                                         Debug.WriteLine($"在工作目录找到文件: {fullFilePath}")
                                     End If
                                 End If

                                 If File.Exists(fullFilePath) Then
                                     ' 根据文件扩展名选择合适的解析方法
                                     Dim fileExtension As String = Path.GetExtension(fullFilePath).ToLower()
                                     Dim fileContentResult As FileContentResult = Nothing

                                     Select Case fileExtension
                                         Case ".xlsx", ".xls", ".xlsm", ".xlsb"
                                             ' Excel文件解析需要在主线程
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".docx", ".doc", ".wps"
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".pptx", ".ppt"
                                             RunUiActionSync(Sub()
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
                                     fileContentBuilder.AppendLine($"文件 '{Path.GetFileName(filePath)}' 未找到，尝试路径: {fullFilePath}")
                                     Debug.WriteLine($"文件未找到: {fullFilePath}")
                                 End If
                             Catch ex As Exception
                                 Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                                 fileContentBuilder.AppendLine($"处理文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}")
                                 fileContentBuilder.AppendLine("---")
                             End Try
                         Next

                         fileContentBuilder.AppendLine("--- 文件内容结束 ---" & vbCrLf)

                         ' 文件解析完成，先保存到记忆（同步保存确保立即可检索），再在主线程继续处理消息
                         Dim appTypeForMemory = GetOfficeAppType()
                         Dim sessionIdForMemory = If(_chatStateService?.CurrentSessionId, Guid.NewGuid().ToString())
                         MemoryService.SaveFileContentToMemory(originalQuestion, fileContentBuilder.ToString(), sessionIdForMemory, appTypeForMemory)

                         ' 文件解析完成，在主线程继续处理消息
                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowInfo($"文件解析完成，共解析 {parsedFiles.Count} 个文件")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")

                                       Dim questionWithFiles = question & " 用户提问结束，后续引用的文件都在同一目录下所以可以放心读取。 ---"
                                       HandleSendMessageCore(questionWithFiles, originalQuestion, filePaths, selectedContents, messageValue, fileContentBuilder.ToString())
                                   End Sub)

                     Catch ex As Exception
                         Debug.WriteLine($"HandleSendMessageWithFilesAsync 出错: {ex.Message}")
                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowWarning($"文件解析失败: {ex.Message}")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")
                                       ' 重置发送按钮状态
                                       ExecuteJavaScriptAsyncJS("changeSendButton()")
                                   End Sub)
                     End Try
                 End Sub)
    End Sub

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

        ' 首先保存用户消息到历史记录（确保即便是意图预览模式，用户消息也会被保存）
        If Not String.IsNullOrWhiteSpace(originalQuestion) Then
            ' 可见文本交给执行漏斗单点落库，避免入口和 Send/Agent 各写一条 user
            _pendingUserVisibleText = originalQuestion
            Debug.WriteLine($"已登记本回合用户可见文本: {originalQuestion}")
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
        ElseIf responseMode = "proofread" Then
            ' 校对模式：用户在校对模式下发送消息，注入校对上下文
            Dim proofreadSelectedText = If(messageValue("proofreadSelectedText")?.ToString(), "")
            Dim proofreadIssueCount = If(messageValue("proofreadIssueCount")?.ToString(), "0")
            Dim proofreadSystemPrompt = BuildProofreadFollowUpPrompt(proofreadSelectedText, proofreadIssueCount)
            Task.Run(Async Function()
                         Await Send(finalMessageToLLM, proofreadSystemPrompt, False, "proofread_followup")
                     End Function)
        Else
            ' 智能模式：分析 → 预检 → AgentKernel / 普通聊天（逻辑在 ChatRoutingOrchestrator）
            Task.Run(Async Function()
                         Await ChatRoutingOrchestrator.RouteSmartModeAsync(
                             finalMessageToLLM,
                             originalQuestion,
                             filePaths,
                             selectedContents)
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
            Dim intentDesc As String = If(intent.UserFriendlyDescription, "")
            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "", intentDesc)
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

#Region "IChatRoutingHost"

    Private Function IChatRoutingHost_GetHistoryMessages() As List(Of HistoryMessage) Implements IChatRoutingHost.GetHistoryMessages
        Return systemHistoryMessageData
    End Function

    Private Function IChatRoutingHost_IsFollowUpQuestionAsync(question As String, history As List(Of HistoryMessage)) As Task(Of Boolean) Implements IChatRoutingHost.IsFollowUpQuestionAsync
        Return IntentService.IsFollowUpQuestionAsync(question, history)
    End Function

    Private Function IChatRoutingHost_GetContextSnapshot() As JObject Implements IChatRoutingHost.GetContextSnapshot
        ' 路由编排跑在后台线程，快照是纯 COM 读取：在宿主适配边界收口到 UI/STA
        Return RunUiFuncSync(Function() GetContextSnapshot())
    End Function

    Private Sub IChatRoutingHost_EnrichContextForIntent(snapshot As JObject,
                                                        originalQuestion As String,
                                                        filePaths As List(Of String),
                                                        selectedContents As List(Of SendMessageReferenceContentItem)) Implements IChatRoutingHost.EnrichContextForIntent
        ' 这里只做 RAG 检索，不读宿主对象模型；留在后台线程，别把 5s 超时压到 UI 上
        EnrichContextForIntent(snapshot, originalQuestion, filePaths, selectedContents)
    End Sub

    Private Function IChatRoutingHost_GetApplicationType() As String Implements IChatRoutingHost.GetApplicationType
        Return RunUiFuncSync(Function() GetApplicationType())
    End Function

    Private Function IChatRoutingHost_CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext Implements IChatRoutingHost.CaptureOfficeContext
        Return RunUiFuncSync(Function() CaptureOfficeContext(appType))
    End Function

    Private Function IChatRoutingHost_AnalyzeAiNativeAsync(request As Agent.AiNativeRequest) As Task(Of Agent.AiNativeRuntimeResult) Implements IChatRoutingHost.AnalyzeAiNativeAsync
        Return AiNativeRuntime.AnalyzeAsync(request)
    End Function

    Private Function IChatRoutingHost_BuildExecutionContextAsync(originalQuestion As String,
                                                                 filePaths As List(Of String),
                                                                 selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ExecutionContext) Implements IChatRoutingHost.BuildExecutionContextAsync
        Return BuildExecutionContextAsync(originalQuestion, filePaths, selectedContents)
    End Function

    Private Function IChatRoutingHost_PreSendCheckAsync(execContext As ExecutionContext) As Task(Of ContextCheckResult) Implements IChatRoutingHost.PreSendCheckAsync
        Return SelfCheckLoopController.PreSendCheckAsync(execContext)
    End Function

    Private Sub IChatRoutingHost_ShowContextHints(ragCount As Integer, intentDescription As String, contextTrace As ChatContextTrace) Implements IChatRoutingHost.ShowContextHints
        Dim intentText = If(intentDescription, "")
        Dim traceJson = If(contextTrace Is Nothing, "{}", UtilsService.SerializeForScript(JObject.FromObject(contextTrace)))
        ExecuteJavaScriptAsyncJS($"showContextHints({{ ragCount: {ragCount}, intent: {UtilsService.ToJsStringLiteral(intentText)}, trace: {traceJson} }});")
    End Sub

    Private Sub IChatRoutingHost_ShowWarning(message As String) Implements IChatRoutingHost.ShowWarning
        GlobalStatusStrip.ShowWarning(message)
    End Sub

    Private Sub IChatRoutingHost_SendChatMessage(message As String) Implements IChatRoutingHost.SendChatMessage
        SendChatMessage(message)
    End Sub

    Private Sub IChatRoutingHost_SendChatMessageWithIntent(message As String, intent As IntentResult) Implements IChatRoutingHost.SendChatMessageWithIntent
        SendChatMessageWithIntent(message, intent)
    End Sub

    Private Sub IChatRoutingHost_StartAgentPlanningFlow(message As String,
                                                        intent As IntentResult,
                                                        analysis As Agent.AiNativeRuntimeResult) Implements IChatRoutingHost.StartAgentPlanningFlow
        StartAgentPlanningFlow(message, intent, analysis)
    End Sub

    Private Sub IChatRoutingHost_SetCurrentIntentResult(intent As IntentResult) Implements IChatRoutingHost.SetCurrentIntentResult
        CurrentIntentResult = intent
    End Sub

#End Region

    ''' <summary>
    ''' Agent模式下直接启动规划流程（不显示意图预览卡片）
    ''' Primary product path: AgentKernel planning/execution.
    ''' 这条路径给的是空入口身份（E6-05）：输入框里是自由文本，没有入口动作，也就没有任务包合同可核对。
    ''' </summary>
    Private Sub StartAgentPlanningFlow(message As String,
                                       intent As IntentResult,
                                       Optional analysis As Agent.AiNativeRuntimeResult = Nothing)
        ' 会话档案要在还在 UI 线程时就把当前文档读走，后台规划线程不能再碰 COM
        RefreshSessionDocumentContext()
        AgentKernelSvc.AgentFullUserMessage = message
        AgentKernelSvc.AgentOriginalUserRequest = If(intent?.OriginalInput, message)
        ' Agent 回合的用户事件由 AgentKernelService 落库，这里丢弃待写文本，避免后续 Send 误用
        _pendingUserVisibleText = Nothing

        Task.Run(Async Function()
                     Try
                         Dim goal = If(String.IsNullOrWhiteSpace(intent.OriginalInput), intent.UserFriendlyDescription, intent.OriginalInput)
                         ' 宿主对象模型只能在 UI/STA 读取，先取快照再把耗时的规划/检索留在后台线程
                         Dim appType = RunUiFuncSync(Function() GetApplicationType())

                         GlobalStatusStrip.ShowInfo("正在规划任务...")

                         Dim currentContent = RunUiFuncSync(Function() GetCurrentOfficeContent())
                         Dim officeContext = RunUiFuncSync(Function() CaptureOfficeContext(appType))
                         Dim historyMessages As New List(Of Tuple(Of String, String))()
                         For Each msg In systemHistoryMessageData
                             If msg.role = "user" OrElse msg.role = "assistant" Then
                                 historyMessages.Add(New Tuple(Of String, String)(msg.role, msg.content))
                             End If
                         Next

                         Dim success = Await AgentKernelSvc.StartAgentAsync(
                             message,
                             appType,
                             currentContent,
                             historyMessages,
                             officeContext,
                             analysis?.TaskSpec,
                             analysis?.SelectedSkills,
                             "")

                         If Not success Then
                             ' StartAgentAsync 已通过 Agent 完成事件展示结构化失败。执行型请求在
                             ' 失败后再发起普通聊天会产生第二份、且可能宣称成功的回答。
                             Debug.WriteLine("[StartAgentPlanningFlow] Agent 执行失败，保留失败结果，不回退普通聊天")
                             GlobalStatusStrip.ShowWarning("任务执行失败，请查看失败步骤")
                         End If
                     Catch ex As Exception
                         Debug.WriteLine($"[StartAgentPlanningFlow] {ex.Message}")
                         GlobalStatusStrip.ShowWarning($"规划启动失败: {ex.Message}")
                     End Try
                 End Function)
    End Sub

    Protected Overridable Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
        Return New Agent.Context.OfficeContext With {.AppType = appType}
    End Function

    ''' <summary>
    ''' 获取应用类型（子类重写）
    ''' </summary>
    Protected Overridable Function GetApplicationType() As String
        Dim appInfo = GetApplication()
        Dim appType = If(appInfo IsNot Nothing, appInfo.Type.ToString(), "Excel")
        Return appType
    End Function

#Region "Ralph Agent 智能助手"

    ''' <summary>
    ''' 宿主原生入口（Word 选区工具条 E4-01、E4-02 右键菜单、E4-03 快捷键、三端 Ribbon）发起一次 Agent 回合。
    ''' 直接进 HandleStartAgent，与面板输入同一个闸门、同一个上下文采集、同一条主链；
    ''' 不再让宿主拼 sendMessageToServer 的 JS 字符串，避免依赖 WebView 加载时序和一处以上话术副本。
    ''' 形参收 QuickAction 而不是 String：E6-05 的断言要靠入口身份绑任务包，
    ''' 话术和 id 必须来自同一个声明处，少传 id 要在这里编译不过，而不是等运行时猜。
    ''' </summary>
    Public Sub StartAgentFromHostEntry(action As Agent.QuickAction)
        If action Is Nothing OrElse String.IsNullOrWhiteSpace(action.RequestText) Then Return
        Dim payload As New JObject()
        payload("request") = action.RequestText
        payload("entryActionId") = If(action.Id, "")
        HandleStartAgent(payload)
    End Sub

    ''' <summary>
    ''' 本轮 Agent 是否还在跑。宿主原生入口据此不再叠加浮层，免得打断计划卡与风险审批。
    ''' 只读已实例化的服务，不触发懒加载。
    ''' </summary>
    Public ReadOnly Property IsAgentRunActive As Boolean
        Get
            If _agentKernelService Is Nothing Then Return False
            Return Not String.IsNullOrWhiteSpace(_agentKernelService.AgentThinkingUuid) OrElse
                   Not String.IsNullOrWhiteSpace(_agentKernelService.CurrentAgentSessionId)
        End Get
    End Property

    ''' <summary>
    ''' 处理启动Agent请求（AgentKernel 主路径）
    ''' </summary>
    Protected Sub HandleStartAgent(jsonDoc As JObject)
        Try
            ' Agent 主链没有配置时就启动，只会在半路失败并留下一个空思考气泡；这里用同一个 CTA 拦住
            If IsBlockedByModelConfig(ModelConfigGate.ActionAgent) Then
                ExecuteJavaScriptAsyncJS("changeSendButton()")
                Return
            End If

            ' 同一条漏斗上的第二把门：还没有同意过数据边界，就还不能把文档发出去（E3-05）
            If IsBlockedByPrivacy(ModelConfigGate.ActionAgent) Then
                ExecuteJavaScriptAsyncJS("changeSendButton()")
                Return
            End If

            ' Agent 回合的用户事件由后台路径落库，文档标识必须此刻在 UI 线程取好
            RefreshSessionDocumentContext()
            Dim request = jsonDoc("request")?.ToString()
            Dim filePathsToken = jsonDoc("filePaths")
            Dim selectedContentToken = jsonDoc("selectedContent")
            ' 入口身份随载荷走：命令面板/宿主入口带 id，自由输入没有。空 id 表示本轮没有包合同可核对。
            Dim entryActionId = If(jsonDoc("entryActionId")?.ToString(), "")

            If String.IsNullOrEmpty(request) AndAlso (filePathsToken Is Nothing OrElse filePathsToken.Type <> JTokenType.Array) AndAlso (selectedContentToken Is Nothing OrElse selectedContentToken.Type <> JTokenType.Array) Then
                GlobalStatusStrip.ShowWarning("请输入任务描述或添加文件引用")
                Return
            End If

            Debug.WriteLine($"[AgentKernel] 启动Agent，需求: {request}")
            AgentKernelSvc.AgentOriginalUserRequest = request
            _pendingUserVisibleText = Nothing

            ' 解析文件路径和选中内容
            Dim filePaths As New List(Of String)()
            Dim selectedContents As New List(Of SendMessageReferenceContentItem)()

            If filePathsToken IsNot Nothing AndAlso filePathsToken.Type = JTokenType.Array Then
                filePaths = filePathsToken.ToObject(Of List(Of String))()
            End If
            If selectedContentToken IsNot Nothing AndAlso selectedContentToken.Type = JTokenType.Array Then
                Try
                    selectedContents = selectedContentToken.ToObject(Of List(Of SendMessageReferenceContentItem))()
                Catch ex As Exception
                    Debug.WriteLine($"Error deserializing selectedContent: {ex.Message}")
                End Try
            End If

            ' 显示思考状态
            AgentKernelSvc.AgentThinkingUuid = Guid.NewGuid().ToString()
            Dim now = DateTime.Now
            Dim timestamp = now.ToString("yyyy-MM-dd HH:mm:ss")
            ExecuteJavaScriptAsyncJS($"createChatSection({UtilsService.ToJsStringLiteral("AI")}, {UtilsService.ToJsStringLiteral(timestamp)}, {UtilsService.ToJsStringLiteral(AgentKernelSvc.AgentThinkingUuid)})")
            ExecuteJavaScriptAsyncJS($"showThinkingIndicator({UtilsService.ToJsStringLiteral(AgentKernelSvc.AgentThinkingUuid)})")

            Dim originalQuestion As String = request
            Dim finalMessageToLLM As String = request
            finalMessageToLLM = BuildMessageWithAutoSelection("--- 我此次的问题：" & finalMessageToLLM & " ---")
            RenderAutoSelectionReference(AgentKernelSvc.AgentThinkingUuid)

            If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
                HandleStartAgentWithFilesAsync(finalMessageToLLM, originalQuestion, filePaths, selectedContents, entryActionId)
            Else
                HandleStartAgentCore(finalMessageToLLM, originalQuestion, "", entryActionId)
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleStartAgent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"启动Agent失败: {ex.Message}")
        End Try
    End Sub

    Private Sub HandleStartAgentWithFilesAsync(question As String, originalQuestion As String,
                                              filePaths As List(Of String),
                                              selectedContents As List(Of SendMessageReferenceContentItem),
                                              entryActionId As String)
        ' 显示进度提示
        GlobalStatusStrip.ShowInfo($"正在解析 {filePaths.Count} 个文件...")
        ExecuteJavaScriptAsyncJS("showFileParsingProgress(true)")

        Task.Run(Sub()
                     Try
                         Dim fileContentBuilder As New StringBuilder()
                         Dim totalFiles = filePaths.Count
                         Dim processedFiles = 0

                         fileContentBuilder.AppendLine(vbCrLf & "--- 以下是用户引用的其他文件内容 ---")

                         ' 获取当前工作目录（需要在主线程调用）
                         Dim currentWorkingDir As String = ""
                         RunUiActionSync(Sub()
                                       currentWorkingDir = GetCurrentWorkingDirectory()
                                   End Sub)

                         For Each filePath As String In filePaths
                             Try
                                 processedFiles += 1

                                 ' 更新进度
                                 RunUiActionSync(Sub()
                                               GlobalStatusStrip.ShowInfo($"正在解析文件 ({processedFiles}/{totalFiles}): {Path.GetFileName(filePath)}")
                                               ExecuteJavaScriptAsyncJS($"updateFileParsingProgress({processedFiles}, {totalFiles}, {UtilsService.ToJsStringLiteral(Path.GetFileName(filePath))})")
                                           End Sub)

                                 ' 确定完整文件路径
                                 Dim fullFilePath As String = filePath

                                 ' 如果是绝对路径且文件存在，直接使用
                                 If Path.IsPathRooted(filePath) AndAlso File.Exists(filePath) Then
                                     fullFilePath = filePath
                                     Debug.WriteLine($"使用绝对路径: {fullFilePath}")
                                 ElseIf Not String.IsNullOrEmpty(currentWorkingDir) Then
                                     ' 尝试在当前工作目录下查找
                                     Dim tryPath = Path.Combine(currentWorkingDir, Path.GetFileName(filePath))
                                     If File.Exists(tryPath) Then
                                         fullFilePath = tryPath
                                         Debug.WriteLine($"在工作目录找到文件: {fullFilePath}")
                                     End If
                                 End If

                                 If File.Exists(fullFilePath) Then
                                     ' 根据文件扩展名选择合适的解析方法
                                     Dim fileExtension As String = Path.GetExtension(fullFilePath).ToLower()
                                     Dim fileContentResult As FileContentResult = Nothing

                                     Select Case fileExtension
                                         Case ".xlsx", ".xls", ".xlsm", ".xlsb"
                                             ' Excel文件解析需要在主线程
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".docx", ".doc", ".wps"
                                             RunUiActionSync(Sub()
                                                           fileContentResult = ParseFile(fullFilePath)
                                                       End Sub)
                                         Case ".pptx", ".ppt"
                                             RunUiActionSync(Sub()
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
                                         fileContentBuilder.AppendLine($"文件名: {fileContentResult.FileName}")
                                         fileContentBuilder.AppendLine($"文件内容:")
                                         fileContentBuilder.AppendLine(fileContentResult.ParsedContent)
                                         fileContentBuilder.AppendLine("---")
                                     End If
                                 Else
                                     fileContentBuilder.AppendLine($"文件 '{Path.GetFileName(filePath)}' 未找到，尝试路径: {fullFilePath}")
                                     Debug.WriteLine($"文件未找到: {fullFilePath}")
                                 End If
                             Catch ex As Exception
                                 Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                                 fileContentBuilder.AppendLine($"处理文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}")
                                 fileContentBuilder.AppendLine("---")
                             End Try
                         Next

                         fileContentBuilder.AppendLine("--- 文件内容结束 ---" & vbCrLf)

                         ' 文件解析完成，先保存到记忆（同步保存确保立即可检索），再在主线程继续处理消息
                         Dim appTypeForMemory = GetOfficeAppType()
                         Dim sessionIdForMemory = If(_chatStateService?.CurrentSessionId, Guid.NewGuid().ToString())
                         MemoryService.SaveFileContentToMemory(originalQuestion, fileContentBuilder.ToString(), sessionIdForMemory, appTypeForMemory)

                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowInfo($"文件解析完成，共解析 {processedFiles} 个文件")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")

                                       Dim questionWithFiles = question & " 用户提问结束，后续引用的文件都在同一目录下所以可以放心读取。 ---"
                                       HandleStartAgentCore(questionWithFiles, originalQuestion, fileContentBuilder.ToString(), entryActionId)
                                   End Sub)

                     Catch ex As Exception
                         Debug.WriteLine($"HandleStartAgentWithFilesAsync 出错: {ex.Message}")
                         RunUiActionSync(Sub()
                                       GlobalStatusStrip.ShowWarning($"文件解析失败: {ex.Message}")
                                       ExecuteJavaScriptAsyncJS("showFileParsingProgress(false)")
                                   End Sub)
                     End Try
                 End Sub)
    End Sub

    ''' <summary>
    ''' 处理Agent启动的核心逻辑（文件解析完成后调用）
    ''' AgentKernel 启动核心（始终注入 CaptureOfficeContext）
    ''' entryActionId 是必填项：每条路径都要说清自己是哪个入口动作（自由文本置为空串），
    ''' 这样任务包断言要么真的核对，要么明确没有合同，不会靠"忘了传"来跳过验证。
    ''' </summary>
    Private Sub HandleStartAgentCore(question As String, originalQuestion As String, fileContent As String,
                                     entryActionId As String)
        ' 构建最终发送给 LLM 的消息
        Dim finalMessageToLLM As String = question

        ' 然后添加文件内容（如果有）
        If Not String.IsNullOrEmpty(fileContent) Then
            finalMessageToLLM &= fileContent
        End If

        AgentKernelSvc.AgentFullUserMessage = finalMessageToLLM
        AgentKernelSvc.AgentOriginalUserRequest = originalQuestion
        _pendingUserVisibleText = Nothing

        Task.Run(Async Function()
                     Try
                         ' 同上：COM 快照取回 UI/STA，模型请求与检索继续跑在后台线程
                         Dim appType = RunUiFuncSync(Function() GetApplicationType())
                         Dim currentContent = RunUiFuncSync(Function() GetCurrentOfficeContent())
                         Dim officeContext = RunUiFuncSync(Function() CaptureOfficeContext(appType))

                         Dim historyMessages As New List(Of Tuple(Of String, String))()
                         For Each msg In systemHistoryMessageData
                             If msg.role = "user" OrElse msg.role = "assistant" Then
                                 historyMessages.Add(New Tuple(Of String, String)(msg.role, msg.content))
                             End If
                         Next

                         GlobalStatusStrip.ShowInfo("正在分析您的需求...")
                         Dim success = Await AgentKernelSvc.StartAgentAsync(finalMessageToLLM, appType, currentContent, historyMessages, officeContext, Nothing, Nothing, entryActionId)

                         If Not success Then
                             GlobalStatusStrip.ShowWarning("无法分析您的需求，请重试")
                             AgentKernelSvc.AgentThinkingUuid = Nothing
                         End If
                     Catch ex As Exception
                         Debug.WriteLine($"[AgentKernel] HandleStartAgentCore 出错: {ex.Message}")
                         GlobalStatusStrip.ShowWarning($"分析需求失败: {ex.Message}")
                         AgentKernelSvc.AgentThinkingUuid = Nothing
                     End Try
                 End Function)
    End Sub

    ''' <summary>
    ''' 处理终止Agent
    ''' </summary>
    Protected Sub HandleAbortAgent()
        AgentKernelSvc.AbortAgent()
    End Sub

#Region "AgentKernel 统一消息处理"

    ''' <summary>
    ''' 处理用户批准计划
    ''' </summary>
    Protected Sub HandleAgentApprovePlan(jsonDoc As JObject)
        Try
            AgentKernelSvc.Approve()
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentApprovePlan 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理用户拒绝计划
    ''' </summary>
    Protected Sub HandleAgentRejectPlan(jsonDoc As JObject)
        Try
            AgentKernelSvc.Reject()
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentRejectPlan 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理用户批准步骤（高风险操作）
    ''' </summary>
    Protected Sub HandleAgentApproveStep(jsonDoc As JObject)
        Try
            AgentKernelSvc.Approve()
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentApproveStep 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理“本会话允许该工具”的审批裁决
    ''' </summary>
    Protected Sub HandleAgentApproveToolSession(jsonDoc As JObject)
        Try
            AgentKernelSvc.ApproveForSession()
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentApproveToolSession 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理用户修改计划
    ''' </summary>
    Protected Sub HandleAgentRefinePlan(jsonDoc As JObject)
        Try
            Dim feedback = jsonDoc("payload")?("feedback")?.ToString()
            If String.IsNullOrEmpty(feedback) Then
                feedback = jsonDoc("feedback")?.ToString()
            End If
            If Not String.IsNullOrEmpty(feedback) Then
                Debug.WriteLine($"[AgentKernel] 用户请求修改计划: {feedback}")
                Dim request = AgentKernelSvc.AgentOriginalUserRequest
                If Not String.IsNullOrEmpty(request) Then
                    Dim refinedRequest = request & vbCrLf & "[用户修改意见] " & feedback
                    AgentKernelSvc.AgentThinkingUuid = Guid.NewGuid().ToString()
                    AgentKernelSvc.AgentThinkingLabel = "正在根据您的意见重新规划..."
                    ' 改计划仍然是同一件事：入口身份沿用本轮在跑的，包断言才不会因为一次改写而失效
                    HandleStartAgentCore(refinedRequest, request, "", AgentKernelSvc.AgentEntryActionId)
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentRefinePlan 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 撤销最近一轮 Agent 写入。每个宿主自行判定能否安全撤销，基类不做虚假承诺。
    ''' </summary>
    Protected Sub HandleAgentUndoRun(jsonDoc As JObject)
        Dim runId = If(jsonDoc?("runId")?.ToString(), jsonDoc?("payload")?("runId")?.ToString())
        Dim outcome As Agent.AgentRunUndoResult
        If String.IsNullOrWhiteSpace(runId) Then
            outcome = Agent.AgentRunUndoResult.Failed("缺少本轮运行标识，无法撤销")
        Else
            Try
                outcome = PerformUndoAgentRun(runId)
            Catch ex As Exception
                outcome = Agent.AgentRunUndoResult.Failed($"撤销失败: {ex.Message}")
            End Try
        End If

        If outcome.Success Then
            GlobalStatusStrip.ShowInfo(outcome.Message)
        Else
            GlobalStatusStrip.ShowWarning(outcome.Message)
        End If

        Try
            ExecuteJavaScriptAsyncJS($"setAgentUndoResult({UtilsService.ToJsStringLiteral(If(runId, ""))}, " &
                                     $"{outcome.Success.ToString().ToLowerInvariant()}, " &
                                     $"{UtilsService.ToJsStringLiteral(outcome.Message)}, " &
                                     $"{outcome.Verified.ToString().ToLowerInvariant()})")
        Catch ex As Exception
            Debug.WriteLine($"HandleAgentUndoRun UI 回写失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 撤销某轮 Agent 写入。默认不支持；宿主具备安全撤销单元时覆写。
    ''' </summary>
    Protected Overridable Function PerformUndoAgentRun(runId As String) As Agent.AgentRunUndoResult
        Return Agent.AgentRunUndoResult.Failed("当前应用不支持撤销本轮写入")
    End Function

#End Region

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
    ''' 构建执行上下文（供自检Loop使用）
    ''' </summary>
    Protected Overridable Async Function BuildExecutionContextAsync(
        originalQuestion As String,
        filePaths As List(Of String),
        selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ExecutionContext)

        Dim context As New ExecutionContext()
        context.UserMessage = originalQuestion
        context.OriginalQuestion = originalQuestion
        context.FilePaths = If(filePaths, New List(Of String)())

        ' 根据意图设置请求操作类型
        If CurrentIntentResult IsNot Nothing Then
            Select Case CurrentIntentResult.OfficeIntent
                Case OfficeIntentType.FORMAT_STYLE, OfficeIntentType.TEXT_FORMAT
                    context.RequestedOperation = RequestedOperation.Reformat
                    context.ExpectedFormat = InstructionFormat.DslJson
                Case OfficeIntentType.DOCUMENT_EDIT
                    ' 文档编辑可能包含校对需求，根据描述判断
                    If CurrentIntentResult.UserFriendlyDescription IsNot Nothing AndAlso
                       CurrentIntentResult.UserFriendlyDescription.ToLower().Contains("校对") Then
                        context.RequestedOperation = RequestedOperation.Proofread
                        context.ExpectedFormat = InstructionFormat.ProofreadJson
                    Else
                        context.RequestedOperation = RequestedOperation.Reformat
                        context.ExpectedFormat = InstructionFormat.DslJson
                    End If
                Case Else
                    context.RequestedOperation = RequestedOperation.GeneralQuery
            End Select
        End If

        ' 获取Office内容信息
        Try
            ' 这段是纯 COM 读取，调用方可能在后台线程；RAG/网络不在这里，不会把等待压到 UI 上
            RunUiActionSync(
                Sub()
                    Dim appInfo = GetApplication()
                    If appInfo IsNot Nothing Then
                        Select Case appInfo.Type.ToString().ToLower()
                            Case "word"
                                context.OfficeAppType = OfficeAppType.Word
                            Case "excel"
                                context.OfficeAppType = OfficeAppType.Excel
                            Case "powerpoint"
                                context.OfficeAppType = OfficeAppType.PowerPoint
                        End Select
                    End If

                    Dim contentInfo As New OfficeContentInfo()
                    Dim selInfo = CaptureCurrentSelectionInfo("")
                    If selInfo IsNot Nothing Then
                        contentInfo.SelectedParagraphs = New List(Of String) From {selInfo.SelectedText}
                        context.SelectionInfo = selInfo
                        context.RequiresSelection = Not String.IsNullOrEmpty(selInfo.SelectedText)
                    End If
                    context.OfficeContent = contentInfo
                End Sub)
        Catch ex As Exception
            Debug.WriteLine($"[BuildExecutionContextAsync] 获取Office内容失败: {ex.Message}")
        End Try

        ' 历史上下文
        If systemHistoryMessageData IsNot Nothing Then
            context.ConversationHistory = systemHistoryMessageData.ToList()
        End If

        Return context
    End Function

    ''' <summary>
    ''' 非流式 AI 请求，直接获取完整响应（供 AgentKernel 使用）
    ''' </summary>
    Private Async Function SendAndGetResponseAsync(prompt As String, systemPrompt As String, historyMessages As List(Of HistoryMessage)) As Task(Of String)
        ' E3-05：排版、审阅与 AgentKernel 的非流式请求都从这一个发送器出去，它们不经过 Send，
        ' 所以同意要在这里再问一次；返回值沿用本函数既有的失败约定（空串），不发明第三种终态。
        If IsBlockedByPrivacy(ModelConfigGate.ActionChat) Then Return ""
        ' Agent 回合的"停止"要能掐断在途请求；不在回合作用域内时为 None，其它调用方行为不变。
        Dim ambientCancellation = Agent.AgentCancellationScope.Current
        Try
            Dim messagesArray As New JArray()
            If Not String.IsNullOrEmpty(systemPrompt) Then
                messagesArray.Add(New JObject From {{"role", "system"}, {"content", systemPrompt}})
            End If
            If historyMessages IsNot Nothing Then
                For Each msg In historyMessages
                    If Not String.IsNullOrEmpty(msg.content) Then
                        messagesArray.Add(New JObject From {{"role", msg.role}, {"content", msg.content}})
                    End If
                Next
            End If
            messagesArray.Add(New JObject From {{"role", "user"}, {"content", prompt}})

            Dim requestObj As New JObject()
            requestObj("model") = ConfigSettings.ModelName
            requestObj("messages") = messagesArray
            requestObj("stream") = False
            ReasoningRequestHelper.ApplyReasoningOptions(requestObj, ConfigSettings.ReasoningMode, ConfigSettings.ModelName, ConfigSettings.platform, ConfigSettings.ApiUrl)

            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)
            Dim apiUrl = ConfigSettings.ApiUrl
            Dim apiKey = ConfigSettings.ApiKey

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim client = HttpClientPool.GetClient(apiUrl)
            Using request As New HttpRequestMessage(HttpMethod.Post, apiUrl)

                ' Anthropic 兼容
                If apiUrl.Contains("anthropic.com") Then
                    request.Headers.Add("x-api-key", apiKey)
                    request.Headers.Add("anthropic-version", "2023-06-01")
                Else
                    request.Headers.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)
                End If

                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Dim timeoutCts = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(ambientCancellation)
                timeoutCts.CancelAfter(TimeSpan.FromMinutes(2))
                Using timeoutCts
                    Using response As HttpResponseMessage = Await client.SendAsync(request, timeoutCts.Token)
                        response.EnsureSuccessStatusCode()
                        Dim jsonContent As String = Await response.Content.ReadAsStringAsync()

                    ' 提取 content
                    Dim json = JObject.Parse(jsonContent)
                    Dim content As String = ""

                    ' OpenAI 格式
                    If json("choices") IsNot Nothing Then
                        Dim choices = CType(json("choices"), JArray)
                        If choices.Count > 0 Then
                            content = If(choices(0)("message")?("content")?.ToString(), "")
                        End If
                    End If

                    ' Anthropic 格式
                    If String.IsNullOrEmpty(content) AndAlso json("content") IsNot Nothing Then
                        Dim contentArr = CType(json("content"), JArray)
                        If contentArr IsNot Nothing AndAlso contentArr.Count > 0 Then
                            content = If(contentArr(0)("text")?.ToString(), "")
                        End If
                    End If

                        Return content
                    End Using
                End Using
            End Using
        Catch ex As OperationCanceledException When ambientCancellation.IsCancellationRequested
            ' 用户点了停止：取消必须继续向上传播，让循环按"已停止"收敛，而不是被当成请求失败吞掉。
            Throw
        Catch ex As Exception
            Debug.WriteLine($"[SendAndGetResponseAsync] 请求失败: {ex.Message}")
            Return ""
        End Try
    End Function

#End Region

    ''' <summary>
    ''' 处理打开文件对话框请求
    ''' </summary>
    Protected Sub HandleOpenFileDialog()
        Try
            ' 需要在UI线程上执行
            If Me.InvokeRequired Then
                RunUiActionSync(Sub() HandleOpenFileDialog())
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
                    ExecuteJavaScriptAsyncJS($"addFilesFromDialog({UtilsService.SerializeForScript(filesArray)})")
                    Debug.WriteLine($"选择了 {dialog.FileNames.Length} 个文件")
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine($"HandleOpenFileDialog 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("打开文件对话框时出错")
        End Try
    End Sub

    ''' <summary>
    ''' 处理打开API配置窗口请求
    ''' </summary>
    Protected Sub HandleOpenApiConfigForm()
        Try
            ' 需要在UI线程上执行
            If Me.InvokeRequired Then
                RunUiActionSync(Sub() HandleOpenApiConfigForm())
                Return
            End If

            Dim configForm As New ConfigApiForm()
            If configForm.ShowDialog() = DialogResult.OK Then
                ' 配置已更新，刷新前端显示；面板内的"请先配置"卡片同时撤掉，否则用户会以为还没保存
                ExecuteJavaScriptAsyncJS(ModelConfigGate.BuildHideScript())
                UpdateModelDisplayInUI()
                ' 引导卡片可能正停在"还没配置"这一步，不重推状态就会一直显示过期的第 1 步
                OnboardingSvc.HandleGetState()
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleOpenApiConfigForm 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("打开配置窗口时出错")
        End Try
    End Sub

    ''' <summary>
    ''' 处理获取当前模型信息请求
    ''' </summary>
    Protected Sub HandleGetCurrentModel()
        Try
            UpdateModelDisplayInUI()
        Catch ex As Exception
            Debug.WriteLine($"HandleGetCurrentModel 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 更新前端的模型显示
    ''' </summary>
    Protected Sub HandleGetCurrentAppInfo()
        Try
            ExecuteJavaScriptAsyncJS($"window.currentOfficeAppName = {UtilsService.ToJsStringLiteral(GetOfficeApplicationName())};")
        Catch ex As Exception
            Debug.WriteLine($"HandleGetCurrentAppInfo failed: {ex.Message}")
        End Try
    End Sub

    Protected Sub UpdateModelDisplayInUI()
        Try
            Dim platform = ConfigSettings.platform
            Dim modelName = ConfigSettings.ModelName

            ' 转义特殊字符
            platform = If(platform, "").Replace("'", "\'").Replace("""", "\""")
            modelName = If(modelName, "").Replace("'", "\'").Replace("""", "\""")

            Dim js = $"updateCurrentModelDisplay('{platform}', '{modelName}');"
            ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine($"UpdateModelDisplayInUI 出错: {ex.Message}")
        End Try
    End Sub

#Region "排版模板功能消息处理"

    ''' <summary>
    ''' 获取排版模板列表（含docx解析出的语义映射卡片）
    ''' </summary>
    Protected Sub HandleGetReformatTemplates()
        ReformatSvc.HandleGetReformatTemplates()
    End Sub

    ''' <summary>
    ''' 刷新排版模板列表（Public，供外部调用）
    ''' </summary>
    Public Sub RefreshReformatTemplates()
        ReformatSvc.HandleGetReformatTemplates()
    End Sub

    ''' <summary>
    ''' 使用排版模板（含docx映射识别）
    ''' </summary>
    Protected Overridable Sub HandleUseReformatTemplate(jsonDoc As JObject)
        Try
            Dim templateId = jsonDoc("templateId")?.ToString()

            ' 识别docx映射卡片（ID前缀 "docx_"）
            If templateId IsNot Nothing AndAlso templateId.StartsWith("docx_") Then
                Dim mappingId = templateId.Substring(5)
                Dim mapping = SemanticMappingManager.Instance.GetMappingById(mappingId)
                If mapping IsNot Nothing Then
                    ApplyReformatWithMapping(mapping)
                    Return
                Else
                    GlobalStatusStrip.ShowWarning("语义映射不存在")
                    Return
                End If
            End If

            ' 常规模板
            Dim template = ReformatTemplateManager.Instance.GetTemplateById(templateId)
            If template Is Nothing Then
                GlobalStatusStrip.ShowWarning("模板不存在")
                Return
            End If

            ApplyReformatWithTemplate(template)

        Catch ex As Exception
            Debug.WriteLine($"HandleUseReformatTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"使用模板失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用模板进行排版（由子类实现）
    ''' </summary>
    Protected Overridable Sub ApplyReformatWithTemplate(template As ReformatTemplate)
        GlobalStatusStrip.ShowWarning("当前应用不支持模板排版")
    End Sub

    ''' <summary>
    ''' 使用SemanticStyleMapping直接排版（由子类实现，用于docx解析的映射）
    ''' </summary>
    Protected Overridable Sub ApplyReformatWithMapping(mapping As SemanticStyleMapping)
        GlobalStatusStrip.ShowWarning("当前应用不支持文档映射排版")
    End Sub

#Region "排版规范处理方法"

    ''' <summary>
    ''' 获取排版规范列表
    ''' </summary>
    Protected Sub HandleGetStyleGuides()
        ReformatSvc.HandleGetStyleGuides()
    End Sub

    ''' <summary>
    ''' 使用排版规范
    ''' </summary>
    Protected Overridable Sub HandleUseStyleGuide(jsonDoc As JObject)
        Try
            Dim guideId = jsonDoc("guideId")?.ToString()
            Dim guide = StyleGuideManager.Instance.GetStyleGuideById(guideId)

            If guide Is Nothing Then
                GlobalStatusStrip.ShowWarning("规范不存在")
                Return
            End If

            ' 由子类实现具体的排版逻辑
            ApplyReformatWithStyleGuide(guide)

        Catch ex As Exception
            Debug.WriteLine($"HandleUseStyleGuide 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"使用规范失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用规范进行排版（由子类实现）
    ''' </summary>
    Protected Overridable Sub ApplyReformatWithStyleGuide(guide As StyleGuideResource)
        GlobalStatusStrip.ShowWarning("当前应用不支持规范排版")
    End Sub

    ''' <summary>
    ''' 上传规范文档
    ''' </summary>
    Protected Sub HandleUploadStyleGuideDocument()
        ReformatSvc.HandleUploadStyleGuideDocument()
    End Sub

    ''' <summary>
    ''' 删除规范
    ''' </summary>
    Protected Sub HandleDeleteStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleDeleteStyleGuide(jsonDoc)
    End Sub

    ''' <summary>
    ''' 更新规范内容（编辑保存）
    ''' </summary>
    Protected Sub HandleUpdateStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleUpdateStyleGuide(jsonDoc)
    End Sub

    ''' <summary>
    ''' 复制规范
    ''' </summary>
    Protected Sub HandleDuplicateStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleDuplicateStyleGuide(jsonDoc)
    End Sub

    ''' <summary>
    ''' 导出规范
    ''' </summary>
    Protected Sub HandleExportStyleGuide(jsonDoc As JObject)
        ReformatSvc.HandleExportStyleGuide(jsonDoc)
    End Sub

#End Region

    ''' <summary>
    ''' 在Word中预览模板
    ''' </summary>
    Protected Overridable Sub HandlePreviewTemplateInWord(jsonDoc As JObject)
        GlobalStatusStrip.ShowWarning("当前应用不支持模板预览")
    End Sub

    ''' <summary>
    ''' 保存当前文档为模板
    ''' </summary>
    Protected Overridable Sub HandleSaveCurrentDocumentAsTemplate()
        GlobalStatusStrip.ShowWarning("当前应用不支持保存文档为模板")
    End Sub

    ''' <summary>
    ''' 导入模板
    ''' </summary>
    Protected Sub HandleImportTemplate()
        ReformatSvc.HandleImportTemplate()
    End Sub

    ''' <summary>
    ''' 导出模板
    ''' </summary>
    Protected Sub HandleExportTemplate(jsonDoc As JObject)
        ReformatSvc.HandleExportTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 复制模板
    ''' </summary>
    Protected Sub HandleDuplicateTemplate(jsonDoc As JObject)
        ReformatSvc.HandleDuplicateTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 删除模板
    ''' </summary>
    Protected Sub HandleDeleteTemplate(jsonDoc As JObject)
        ReformatSvc.HandleDeleteTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 打开模板编辑器
    ''' </summary>
    Protected Sub HandleOpenTemplateEditor(jsonDoc As JObject)
        ReformatSvc.HandleOpenTemplateEditor(jsonDoc)
    End Sub

    ''' <summary>
    ''' 显示模板编辑器面板（子类可重写以使用 CustomTaskPane）
    ''' </summary>
    ''' <param name="template">要编辑的模板，为空则新建</param>
    ''' <returns>如果成功显示返回 True，否则返回 False 以使用回退的 WinForm</returns>
    Protected Overridable Function ShowTemplateEditorPane(template As ReformatTemplate) As Boolean
        Return False ' 默认不支持 CustomTaskPane
    End Function

    ''' <summary>
    ''' 获取样式预览回调（子类可重写以提供实时预览功能）
    ''' </summary>
    Protected Overridable Function GetStylePreviewCallback() As PreviewStyleCallback
        Return Nothing
    End Function

    ''' <summary>
    ''' 进入模板选择模式（供Ribbon调用）
    ''' </summary>
    Public Async Sub EnterReformatTemplateMode()
        Await ReformatSvc.EnterReformatTemplateMode()
    End Sub

    ''' <summary>
    ''' 退出模板选择模式
    ''' </summary>
    Public Sub ExitReformatTemplateMode()
        ReformatSvc.ExitReformatTemplateMode()
    End Sub

    ' ========== AI模板编辑器功能 ==========

    ''' <summary>
    ''' 处理保存AI模板
    ''' </summary>
    Protected Sub HandleSaveAiTemplate(jsonDoc As JObject)
        ReformatSvc.HandleSaveAiTemplate(jsonDoc)
    End Sub

    ''' <summary>
    ''' 处理预览AI模板
    ''' </summary>
    Protected Overridable Sub HandlePreviewAiTemplate(jsonDoc As JObject)
        Try
            If Me.InvokeRequired Then
                RunUiActionSync(Sub() HandlePreviewAiTemplate(jsonDoc))
                Return
            End If

            Dim templateJson As String = jsonDoc("templateJson")?.ToString()
            If String.IsNullOrWhiteSpace(templateJson) Then
                GlobalStatusStrip.ShowWarning("没有可预览的模板数据")
                Return
            End If

            Dim template = JsonConvert.DeserializeObject(Of ReformatTemplate)(templateJson)

            ' 调用子类实现的预览方法
            PreviewTemplateInDocument(template)

        Catch ex As Exception
            Debug.WriteLine($"HandlePreviewAiTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"预览模板失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 在文档中预览模板效果（由子类实现）
    ''' </summary>
    Protected Overridable Sub PreviewTemplateInDocument(template As ReformatTemplate)
        GlobalStatusStrip.ShowWarning("当前应用不支持模板预览")
    End Sub

#End Region

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
                ExecuteJavaScriptAsyncJS($"handleExecutionError({UtilsService.ToJsStringLiteral(responseUuid)}, {UtilsService.ToJsStringLiteral(ex.Message)})")
            End If
        End Try
    End Sub


    ' 抽象方法，由子类实现
    Protected MustOverride Function ParseFile(filePath As String) As FileContentResult
    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String

    ''' <summary>
    ''' E4-04：选区自动引用只有一个决策点。关掉开关就必须真的一个字都不附加——
    ''' 三端的 AppendCurrentSelectedContent 只负责「怎么把选区写成提示」，不负责「要不要写」。
    ''' </summary>
    Protected Function BuildMessageWithAutoSelection(message As String) As String
        _autoSelectionReference = Nothing
        If Not selectedCellChecked Then Return message

        Dim appended = AppendCurrentSelectedContent(message)
        If appended Is Nothing OrElse appended.Length <= message.Length Then Return If(appended, message)

        _autoSelectionReference = DescribeAutoSelection(appended.Substring(message.Length))
        Return appended
    End Function

    ''' <summary>引用摘要压成一行：正文已经进提示，这里只让用户看清带出去的是哪一段、多大。</summary>
    Private Shared Function DescribeAutoSelection(referenceText As String) As String
        Dim flat As New System.Text.StringBuilder()
        Dim pendingSpace As Boolean = False
        For Each c In referenceText
            If Char.IsWhiteSpace(c) Then
                pendingSpace = True
            Else
                If flat.Length > 0 AndAlso pendingSpace Then flat.Append(" "c)
                pendingSpace = False
                flat.Append(c)
                If flat.Length >= 160 Then Exit For
            End If
        Next

        Dim preview = flat.ToString()
        Dim tail = If(referenceText.Length > preview.Length, "…", "")
        Return $"当前选区 · 约 {referenceText.Length} 字：{preview}{tail}"
    End Function

    ''' <summary>
    ''' 把本轮自动引用显示到对话里。Agent 回合不渲染用户气泡，不显示就等于「没打招呼把选区发出去了」。
    ''' </summary>
    Private Sub RenderAutoSelectionReference(turnUuid As String)
        If String.IsNullOrWhiteSpace(_autoSelectionReference) Then Return
        ExecuteJavaScriptAsyncJS($"appendTurnSelectionReference({UtilsService.ToJsStringLiteral(turnUuid)}, {UtilsService.ToJsStringLiteral(_autoSelectionReference)})")
    End Sub

    ' 文本/CSV 解析已委托给 FileParserService，请使用 _fileParserService.ParseTextFile()

    Protected MustOverride Function GetApplication() As ApplicationInfo

    ''' <summary>
    ''' 宿主类型，用于前端区分 Word/PowerPoint/Excel，也是 app_type 落库前的归一输入。
    ''' E5-04 把它从「默认 Unknown 的 Overridable」改成 MustOverride：漏实现不再在运行时
    ''' 变成一个既不属于任何宿主、也拿不到空串豁免的值（见 AppTypeContract 与迁移 019），
    ''' 而是在编译期就报错。
    ''' </summary>
    Protected MustOverride Function GetOfficeAppType() As String

    ''' <summary>
    ''' 宿主提供当前文档标识（一般取活动文档的完整路径）。共享层不碰 COM，
    ''' 未覆写的宿主返回 Nothing，会话按"通用会话"处理而不是猜一个名字。
    ''' </summary>
    Protected Overridable Function GetSessionDocument() As SessionDocument
        Return Nothing
    End Function

    ''' <summary>
    ''' 在宿主 UI 线程读取当前文档并缓存进会话状态：消息落库跑在后台线程，
    ''' 那时再回头读 COM 就是跨套间调用。
    ''' </summary>
    Protected Sub RefreshSessionDocumentContext()
        Try
            If _chatStateService IsNot Nothing Then
                _chatStateService.SetCurrentDocument(GetSessionDocument())
            End If
        Catch ex As Exception
            AppLogger.Warn("BaseChatControl", "刷新会话文档绑定失败，本轮按通用会话处理", ex)
        End Try
    End Sub

    Private Function ResolveSessionDocument() As SessionDocument
        RefreshSessionDocumentContext()
        If _chatStateService Is Nothing Then Return Nothing
        Return _chatStateService.CurrentDocument
    End Function

    Protected MustOverride Function GetVBProject() As VBProject
    Protected MustOverride Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
    Protected MustOverride Function RunCode(vbaCode As String)

    Protected MustOverride Sub SendChatMessage(message As String)
    Protected MustOverride Sub GetSelectionContent(target As Object)


    ' 执行代码的方法 - 委托给 CodeExecutionService
    Public Sub ExecuteCode(code As String, language As String, preview As Boolean)
        Dim result = CodeExecutionService.ExecuteCodeWithToolResult(code, language, preview)
        If result IsNot Nothing AndAlso Not result.Success Then
            GlobalStatusStrip.ShowWarning(If(result.UserMessage, result.Message))
        End If
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

    ' _responseModeMap 代理到 _chatStateService，子类通过此属性访问
    Protected ReadOnly Property _responseModeMap As Dictionary(Of String, String)
        Get
            Return _chatStateService.ResponseModeMap
        End Get
    End Property

    ''' <summary>
    ''' 获取当前响应的模式（用于子类检查是否应该跳过某些操作）
    ''' </summary>
    Protected Function GetCurrentResponseMode() As String
        If String.IsNullOrEmpty(_finalUuid) Then Return ""
        If _responseModeMap.ContainsKey(_finalUuid) Then
            Return _responseModeMap(_finalUuid)
        End If
        Return ""
    End Function

    ''' <summary>
    ''' 检查当前是否处于排版模式（排版模式下JSON不应自动执行命令）
    ''' </summary>
    Protected Function IsInReformatMode() As Boolean
        Return GetCurrentResponseMode() = "reformat"
    End Function

    ''' <summary>
    ''' 检查当前是否处于预览模式（预览模式下JSON用于前端显示，不应自动执行命令）
    ''' 包括：排版(reformat)、校对(proofread)
    ''' </summary>
    Protected Function IsInPreviewMode() As Boolean
        Dim mode = GetCurrentResponseMode()
        Return mode = "reformat" OrElse mode = "proofread"
    End Function

    ' 测试方法已移除，如需调试请使用单独的测试类

    ' 存储调用Send时的请求参数（requestUuid/responseUuid -> JObject）
    Protected _savedRequestParams As New Dictionary(Of String, JObject)()

    Private Function GetSendValidationWarning(failure As ChatSendValidationFailure) As String
        Select Case failure
            Case ChatSendValidationFailure.MissingApiKey
                Return "请先配置大模型ApiKey！"
            Case ChatSendValidationFailure.MissingApiUrl
                Return "请先配置大模型Api！"
            Case ChatSendValidationFailure.MissingQuestion
                Return "请输入问题！"
            Case Else
                Return "请求参数不完整"
        End Select
    End Function

    Public Async Function Send(question As String, systemPrompt As String, addHistory As Boolean, responseMode As String, Optional intentDescription As String = Nothing, Optional responseUuid As String = Nothing) As Task
        Dim apiUrl As String = ConfigSettings.ApiUrl
        Dim apiKey As String = ConfigSettings.ApiKey

        Dim validation = SendValidator.Validate(apiUrl, apiKey, question)
        If Not validation.IsValid Then
            ' 配置缺失只给状态栏红字，用户看不到"下一步做什么"；这类失败走面板内 CTA
            If Not IsBlockedByModelConfig(ModelConfigGate.ActionChat, validation) Then
                GlobalStatusStrip.ShowWarning(GetSendValidationWarning(validation.Failure))
            End If
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        ' 配置齐了不代表可以把文档发出去：还没有同意过数据边界时，这条同样拦在发送之前（E3-05）
        If IsBlockedByPrivacy(ModelConfigGate.ActionChat) Then
            ExecuteJavaScriptAsyncJS($"changeSendButton()")
            Return
        End If

        ' 会话历史的 user 回合在本函数单点落库：输入框路径已登记用户可见原文，宿主 Ribbon 路径用本次问题原文
        ' 只在 addHistory 的成对回合落库：False 路径（续写/模板/校对）传的是合成提示词，答复也不入库，
        ' 写进来就留下一条没有答复的 user 行，既污染会话列表，又会被下一轮当作用户原话读回。
        Dim userTurnText = If(String.IsNullOrWhiteSpace(_pendingUserVisibleText), question, _pendingUserVisibleText)
        _pendingUserVisibleText = Nothing
        If addHistory AndAlso Not String.IsNullOrWhiteSpace(userTurnText) Then
            ' 落库在后台线程读 COM 会跨套间，文档标识只能在 UI 线程这里取
            RefreshSessionDocumentContext()
            _chatStateService?.AddMessage("user", userTurnText)
        End If

        Dim uuid As String = If(responseUuid, Guid.NewGuid().ToString())
        ' 这里生成 requestUuid（用于绑定选区）
        Dim requestUuid As String = Guid.NewGuid().ToString()


        ' 将 PendingSelectionInfo 绑定到 requestUuid
        Try
            If PendingSelectionInfo Is Nothing Then
                Dim captured As SelectionInfo = Nothing
                Try
                    ' Send 可能由后台线程发起（意图路由/流式路径），选区读取要回到 UI/STA
                    captured = RunUiFuncSync(Function() CaptureCurrentSelectionInfo(responseMode))
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
            ' Send 会从线程池发起（意图路由/流式），宿主信息读取回到 UI/STA
            Dim appInfoForPrompt = RunUiFuncSync(Function() GetApplication())
            systemPrompt = SystemPromptResolver.ResolveSystemPrompt(systemPrompt, appInfoForPrompt, CurrentIntentResult, responseMode)

            ' ragCount 由 CreateRequestBody 通过 ByRef 返回，无需再次查询记忆
            Dim ragCount As Integer = 0
            Dim contextTrace As ChatContextTrace = Nothing
            Dim requestBody As String = CreateRequestBody(requestUuid, question, systemPrompt, addHistory, ragCount, contextTrace)
            If ragCount > 0 OrElse Not String.IsNullOrEmpty(intentDescription) OrElse contextTrace IsNot Nothing Then
                Dim intentEscaped As String = If(intentDescription, "").Replace("\", "\\").Replace("'", "\'").Replace(vbCr, " ").Replace(vbLf, " ")
                Dim traceJson = If(contextTrace Is Nothing, "null", JObject.FromObject(contextTrace).ToString(Formatting.None))
                Dim js As String = $"showContextHints({{ ragCount: {ragCount}, intent: '{intentEscaped}', trace: {traceJson} }});"
                ExecuteJavaScriptAsyncJS(js)
            End If
            ' 设置历史保存回调（在 FinalizeStream ClearBuffers 前执行）
            Dim capturedQuestion = question
            Dim capturedAddHistory = addHistory
            HttpStreamSvc.FinalizeCallback = Sub(ah, oq)
                                                 Dim answerContent = _chatStateService.PlainMarkdownBuffer.ToString()
                                                 Dim answer = New HistoryMessage() With {
                                                     .role = "assistant",
                                                     .content = answerContent
                                                 }
                                                 If ah Then
                                                     systemHistoryMessageData.Add(answer)
                                                     ManageHistoryMessageSize()
                                                     _chatStateService.AddMessage("assistant", answer.content)
                                                     MemoryService.SaveConversationTurnAsync(oq, answer.content, _chatStateService.CurrentSessionId, GetOfficeAppType())
                                                     MemoryTurnRecorder.RecordConversationTurn(oq, answer.content, _chatStateService.CurrentSessionId, responseMode, ah, GetOfficeAppType())
                                                 End If
                                             End Sub

            Await HttpStreamSvc.SendStreamRequestAsync(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody, question, requestUuid, addHistory, responseMode, responseUuid)
            Await SaveFullWebPageAsync()
        Catch ex As Exception
            Debug.WriteLine("Send 请求失败: " & ex.Message & vbCrLf & ex.StackTrace)
            GlobalStatusStrip.ShowWarning("请求失败: " & ex.Message)
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


    Private Function CreateRequestBody(uuid As String, question As String, systemPrompt As String, addHistory As Boolean, ByRef ragCountOut As Integer, ByRef contextTraceOut As ChatContextTrace) As String
        Return ChatRequestOrchestrator.CreateRequestBody(uuid, question, systemPrompt, addHistory, ragCountOut, contextTraceOut)
    End Function


    ' MCP工具调用和流处理逻辑已移至 HttpStreamService

    ' _responseSelectionMap 代理到 _chatStateService
    Protected ReadOnly Property _responseSelectionMap As Dictionary(Of String, SelectionInfo)
        Get
            Return _chatStateService.ResponseSelectionMap
        End Get
    End Property

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

        ' === 自检Loop: Flush后校验（排版/校对场景的AI响应格式校验）===
        Try
            If _responseModeMap.ContainsKey(_finalUuid) Then
                Dim mode = _responseModeMap(_finalUuid)
                ' semantic_reformat 使用不同的验证逻辑，在 WordAi 的 ApplySemanticTaggingResult 中处理
                If mode = "reformat" OrElse mode = "proofread" Then
                    Dim aiResponse = allPlainMarkdownBuffer.ToString()
                    If Not String.IsNullOrWhiteSpace(aiResponse) Then
                        Dim expectedFormat = If(mode = "proofread", InstructionFormat.ProofreadJson, InstructionFormat.DslJson)
                        Dim execContext = New ExecutionContext()
                        Dim ignored = RunPostFlushValidationAsync(aiResponse, expectedFormat, execContext)
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SelfCheck] PostFlush校验异常: {ex.Message}")
        End Try

    End Sub

    Private Async Function RunPostFlushValidationAsync(aiResponse As String, expectedFormat As InstructionFormat, execContext As ExecutionContext) As Task
        Try
            Dim validation = Await SelfCheckLoopController.PostFlushValidateAsync(aiResponse, expectedFormat, execContext)
            If Not validation.IsValid Then
                Debug.WriteLine($"[SelfCheck] PostFlush校验失败: {String.Join(";", validation.Errors.Select(Function(e) e.Message))}")
                If validation.Errors.Any(Function(e) e.Level = ErrorLevel.Critical) Then
                    GlobalStatusStrip.ShowWarning($"AI响应格式校验未通过，可能存在指令错误")
                End If
            Else
                Debug.WriteLine($"[SelfCheck] PostFlush校验通过，解析到 {validation.ParsedInstructions.Count} 条指令")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[SelfCheck] PostFlush校验异常: {ex.Message}")
        End Try
    End Function


    ' 执行js脚本的异步方法
    Public Async Function ExecuteJavaScriptAsyncJS(js As String) As Task
        Await WebViewBridge.ExecuteScriptAsync(js)
    End Function

    Private Async Function WaitForRendererMapAsync(uuid As String) As Task
        Await WebViewBridge.WaitForRendererMapAsync(uuid)
    End Function



    ' 共用的HTTP请求方法 - 委托给 UtilsService
    Protected Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Return Await UtilsService.SendHttpRequestAsync(apiUrl, apiKey, requestBody)
    End Function

    Public Async Function SaveFullWebPageAsync() As Task
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
        Return Await WebViewBridge.GetFullHtmlContentAsync()
    End Function
    ' HistoryMessage 类已移至 Controls/Models/HistoryMessage.vb

    ' 注入辅助脚本
    Protected Sub InitializeWebView2Script()
        Debug.WriteLine("[DEBUG InitializeWebView2Script] Registering WebMessageReceived handler")
        ' 设置 Web 消息处理器
        AddHandler ChatBrowser.WebMessageReceived, AddressOf WebView2_WebMessageReceived
        ' 注入 VSTO 桥接脚本
        ChatBrowser.ExecuteScriptAsync(UtilsService.GetVstoBridgeScript())
        ' 注入快捷问题配置
        InjectQuickQuestionsConfig()
    End Sub

    ' 注入快捷问题配置到前端
    Private Async Sub InjectQuickQuestionsConfig()
        Try
            Dim questions = ConfigPromptForm.GetQuickQuestionsList()
            Dim questionsJson = JsonConvert.SerializeObject(questions)
            Dim script = $"if(typeof updateQuickQuestions === 'function') {{ updateQuickQuestions({questionsJson}); }} else {{ window.predefinedPrompts = {questionsJson}; }}"
            Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(script)
        Catch ex As Exception
            Debug.WriteLine($"注入快捷问题失败: {ex.Message}")
        End Try
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

    ''' <summary>
    ''' 处理排版JSON解析失败的重试请求
    ''' </summary>
    Protected Overridable Async Sub HandleRetryReformat(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid")?.ToString(), "")
            Dim errorMsg As String = If(jsonDoc("error")?.ToString(), "格式不符合规范")

            If String.IsNullOrEmpty(uuid) Then
                GlobalStatusStrip.ShowWarning("重试失败：缺少uuid")
                Return
            End If

            Dim retryCount As Integer = ReformatSvc.GetRetryCount(uuid)

            If retryCount >= 1 Then
                GlobalStatusStrip.ShowWarning("排版重试次数已达上限")
                Return
            End If

            ReformatSvc.IncrementRetryCount(uuid)

            ' 构建重试提示
            Dim retryPrompt As New System.Text.StringBuilder()
            retryPrompt.AppendLine("你上次返回的JSON格式有错误，请修正后重新返回。")
            retryPrompt.AppendLine()
            retryPrompt.AppendLine($"错误信息：{errorMsg}")
            retryPrompt.AppendLine()
            retryPrompt.AppendLine("请注意以下JSON格式要求：")
            retryPrompt.AppendLine("1. 所有字符串必须使用英文双引号("")")
            retryPrompt.AppendLine("2. 不要在数组或对象的最后一个元素后加逗号")
            retryPrompt.AppendLine("3. 属性名必须用双引号包裹")
            retryPrompt.AppendLine("4. 不要在JSON中包含注释")
            retryPrompt.AppendLine("5. 确保所有括号正确匹配")
            retryPrompt.AppendLine()
            retryPrompt.AppendLine("请只返回修正后的纯JSON，不要包含任何解释文字或代码块标记。")

            GlobalStatusStrip.ShowWarning("JSON解析失败，正在重试...")

            ' 发送重试请求
            Await Send(retryPrompt.ToString(), "", False, "reformat")

        Catch ex As Exception
            Debug.WriteLine("HandleRetryReformat 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("重试失败: " & ex.Message)
        End Try
    End Sub


#Region "语义排版Handler"

    ''' <summary>
    ''' 上传.docx模板文件并解析为SemanticStyleMapping
    ''' </summary>
    Protected Sub HandleUploadDocxTemplate()
        ReformatSvc.HandleUploadDocxTemplate()
    End Sub

    ''' <summary>
    ''' 从指定路径解析.docx模板
    ''' </summary>
    Protected Overridable Sub HandleUploadDocxTemplateFromPath(filePath As String)
        ' 默认不支持，由WordAi子类覆盖实现
        GlobalStatusStrip.ShowWarning("当前应用不支持解析Word模板")
    End Sub

    ''' <summary>
    ''' 删除docx语义映射
    ''' </summary>
    Protected Sub HandleDeleteDocxMapping(jsonDoc As JObject)
        ReformatSvc.HandleDeleteDocxMapping(jsonDoc)
    End Sub

    ''' <summary>
    ''' 撤销排版（根据不同的Office应用使用正确的撤销方式）
    ''' </summary>
    Protected Overridable Sub HandleUndoReformat()
        Try
            Dim appInfo = GetApplication()
            If appInfo Is Nothing Then Return

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try

            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取Office应用对象，请尝试按 Ctrl+Z 撤销")
                Return
            End If

            Dim appName = GetOfficeApplicationName()
            Debug.WriteLine($"HandleUndoReformat: appName={appName}")

            Select Case appName
                Case "Word"
                    ' Word: 支持 UndoRecord，撤销入口在 Document.Undo
                    Try
                        officeApp.ActiveDocument.Undo(1)
                        GlobalStatusStrip.ShowInfo("已撤销排版操作")
                    Catch ex As Exception
                        Debug.WriteLine($"Word Document.Undo 失败，尝试 CommandBars Undo: {ex.Message}")
                        Try
                            officeApp.CommandBars.ExecuteMso("Undo")
                            GlobalStatusStrip.ShowInfo("已撤销排版操作")
                        Catch ex2 As Exception
                            Debug.WriteLine($"Word CommandBars 撤销也失败: {ex2.Message}")
                            GlobalStatusStrip.ShowWarning("撤销排版失败，请手动按 Ctrl+Z")
                        End Try
                    End Try

                Case "PowerPoint"
                    ' PowerPoint: 不支持 UndoRecord，只能撤销最近一次操作
                    Try
                        ' PowerPoint Application 没有 ActiveDocument 属性，使用 ActivePresentation
                        officeApp.ActivePresentation.Undo()
                        GlobalStatusStrip.ShowWarning("已撤销排版操作（PPT撤销受限，如果未完全恢复请多次按 Ctrl+Z）")
                    Catch ex As Exception
                        Debug.WriteLine($"PowerPoint 撤销失败: {ex.Message}")
                        Try
                            ' 备选：通过 CommandBars 触发标准撤销
                            officeApp.CommandBars.ExecuteMso("Undo")
                            GlobalStatusStrip.ShowWarning("已撤销排版操作（PPT撤销受限，如果未完全恢复请多次按 Ctrl+Z）")
                        Catch ex2 As Exception
                            Debug.WriteLine($"PowerPoint CommandBars 撤销也失败: {ex2.Message}")
                            GlobalStatusStrip.ShowWarning("PPT撤销排版失败，请手动按 Ctrl+Z 多次撤销")
                        End Try
                    End Try

                Case "Excel"
                    ' Excel: 排版功能尚在开发中，但保留撤销入口
                    Try
                        ' Excel 使用 ActiveWorkbook，没有 ActiveDocument
                        officeApp.ActiveWorkbook.Undo()
                        GlobalStatusStrip.ShowInfo("已撤销排版操作")
                    Catch ex As Exception
                        Debug.WriteLine($"Excel 撤销失败: {ex.Message}")
                        Try
                            officeApp.CommandBars.ExecuteMso("Undo")
                            GlobalStatusStrip.ShowInfo("已撤销排版操作")
                        Catch ex2 As Exception
                            Debug.WriteLine($"Excel CommandBars 撤销也失败: {ex2.Message}")
                            GlobalStatusStrip.ShowWarning("撤销排版失败，请手动按 Ctrl+Z")
                        End Try
                    End Try

                Case Else
                    ' 未知应用类型，尝试通用方式
                    Try
                        officeApp.Undo()
                        GlobalStatusStrip.ShowInfo("已撤销排版操作")
                    Catch ex As Exception
                        Debug.WriteLine($"通用撤销失败: {ex.Message}")
                        GlobalStatusStrip.ShowWarning("撤销排版失败，请尝试按 Ctrl+Z")
                    End Try
            End Select

        Catch ex As Exception
            Debug.WriteLine($"HandleUndoReformat 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"撤销排版失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 智能排版 v2：应用排版方案（由派生类覆写实现具体Office应用逻辑）
    ''' </summary>
    Protected Overridable Sub HandleApplySmartReformat(jsonDoc As JObject)
        Debug.WriteLine("HandleApplySmartReformat not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：撤销排版（从快照恢复，由派生类覆写实现具体Office应用逻辑）
    ''' </summary>
    Protected Overridable Sub HandleUndoReformat(jsonDoc As JObject)
        Debug.WriteLine("HandleUndoReformat not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：微调排版方案（由派生类覆写实现具体Office应用逻辑）
    ''' </summary>
    Protected Overridable Sub HandleRefineSmartReformat(jsonDoc As JObject)
        Debug.WriteLine("HandleRefineSmartReformat not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：切换排版模板/标准（由派生类覆写实现具体逻辑）
    ''' </summary>
    Protected Overridable Sub HandleSwitchReformatTemplate(jsonDoc As JObject)
        Debug.WriteLine("HandleSwitchReformatTemplate not overridden in derived class")
    End Sub

    ''' <summary>
    ''' 智能排版 v2：显示排版前后对比（由派生类覆写实现具体逻辑）
    ''' </summary>
    Protected Overridable Sub HandlePreviewReformatCompare(jsonDoc As JObject)
        Debug.WriteLine("HandlePreviewReformatCompare not overridden in derived class")
    End Sub


    ''' <summary>
    ''' 校对专注模式消息处理（由派生类覆写实现具体逻辑）
    ''' </summary>
    Protected Overridable Sub HandleProofreadFocusMode(jsonDoc As JObject)
        Debug.WriteLine("HandleProofreadFocusMode not overridden in derived class")
    End Sub

#End Region

End Class
