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
            End If
            Return _codeExecutionService
        End Get
    End Property

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

            ' 将设置发送到前端
            Dim js As String = $"
            document.getElementById('topic-randomness').value = '{ChatSettings.topicRandomness}';
            document.getElementById('topic-randomness-value').textContent = '{ChatSettings.topicRandomness}';
            document.getElementById('context-limit').value = '{ChatSettings.contextLimit}';
            document.getElementById('context-limit-value').textContent = '{ChatSettings.contextLimit}';
            document.getElementById('settings-scroll-checked').checked = {ChatSettings.settingsScrollChecked.ToString().ToLower()};
            document.getElementById('settings-selected-cell').checked = {ChatSettings.selectedCellChecked.ToString().ToLower()};
            document.getElementById('settings-executecode-preview').checked = {ChatSettings.executecodePreviewChecked.ToString().ToLower()};
            
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
        Dim chatSettings As New ChatSettings(GetApplication())
        ' 保存设置到配置文件
        chatSettings.SaveSettings(topicRandomness, contextLimit, selectedCellChecked,
                                  settingsScrollChecked, executeCodePreview, chatMode)
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

        ' --- 文件内容解析逻辑 ---
        Dim fileContentBuilder As New StringBuilder()
        Dim parsedFiles As New List(Of FileContentResult)()

        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
            question = question & " 用户提问结束，后续引用的文件都在同一目录下所以可以放心读取。 ---"

            fileContentBuilder.AppendLine(vbCrLf & "--- 以下是用户引用的其他文件内容 ---")

            ' 获取当前工作目录
            Dim currentWorkingDir As String = GetCurrentWorkingDirectory()
            If String.IsNullOrEmpty(currentWorkingDir) Then
                GlobalStatusStrip.ShowWarning("请保存当前文件，并且把引用文件和当前文件放在同一目录下后重试: ")
                Return
            End If

            For Each filePath As String In filePaths
                Try
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
                                fileContentResult = ParseFile(fullFilePath)
                            Case ".docx", ".doc"
                                fileContentResult = ParseFile(fullFilePath)
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
                        ' 尝试列出当前目录中的文件，用于调试
                        If Directory.Exists(currentWorkingDir) Then
                            Dim filesInDir = Directory.GetFiles(currentWorkingDir)
                            Debug.WriteLine($"当前目录中的文件: {String.Join(", ", filesInDir.Select(Function(f) Path.GetFileName(f)))}")
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error processing file '{filePath}': {ex.Message}")
                    fileContentBuilder.AppendLine($"处理文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}")
                    fileContentBuilder.AppendLine("---")
                End Try
            Next

            fileContentBuilder.AppendLine("--- 文件内容结束 ---" & vbCrLf)
        End If

        ' 构建最终发送给 LLM 的消息
        Dim finalMessageToLLM As String = question

        ' 然后添加文件内容（如果有）
        If fileContentBuilder.Length > 0 Then
            finalMessageToLLM &= fileContentBuilder.ToString()
        End If

        stopReaderStream = False ' Reset stop flag before sending new message
        SendChatMessage(finalMessageToLLM)
    End Sub

    Protected Overridable Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim preview As Boolean = Boolean.Parse(jsonDoc("executecodePreview"))
        Dim language As String = jsonDoc("language").ToString()
        ExecuteCode(code, language, preview)
    End Sub


    ' 抽象方法，由子类实现
    Protected MustOverride Function ParseFile(filePath As String) As FileContentResult
    Protected MustOverride Function GetCurrentWorkingDirectory() As String
    Protected MustOverride Function AppendCurrentSelectedContent(message As String) As String

    ' 文本/CSV 解析已委托给 FileParserService，请使用 _fileParserService.ParseTextFile()

    Protected MustOverride Function GetApplication() As ApplicationInfo
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
                ' 组合更强的 system 提示，要求先给 Plan 再给 Answer，并在信息不足时提出澄清问题
                systemPrompt =
                "系统指令（必读）：" & vbCrLf & ConfigSettings.propmtContent & vbCrLf & vbCrLf &
                "1) 首先输出一个名为 'Plan' 的简短计划，按步骤列出解决路径（要点式，最多6条）。" & vbCrLf &
                "2) 然后输出名为 'Answer' 的部分，给出最终可执行的解决方案或操作步骤，使用 Markdown，必要时给出代码/示例或差异说明。" & vbCrLf &
                "3) 如果信息不足，请不要猜测；在最后输出名为 'Clarifying Questions' 的部分，列出需要用户回答的问题并暂停执行。" & vbCrLf &
                "4) 对于用户请求的改进（用户标记当前回答为不接受），在回复开头先写明 '改进点'（1-3 行），然后给出修正的 Plan 与 Answer。" & vbCrLf &
                "5) 保持回答简洁、有条理，优先提供可直接执行的结论和示例。"
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

        ' 构建 messages 数组
        Dim messages As New List(Of String)()

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

            ' 将历史消息转换为 JSON messages（对内容做基础转义，避免字符串破坏 JSON）
            For Each message In systemHistoryMessageData
                Dim safeContent As String = If(message.content, String.Empty)
                safeContent = safeContent.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n")
                messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{safeContent}""}}")
            Next
        Else
            ' 仅使用当前消息，不添加历史
            Dim tempMessageData As New List(Of HistoryMessage)
            tempMessageData.Insert(0, systemMessage)
            tempMessageData.Add(q)
            For Each message In tempMessageData
                Dim safeContent As String = If(message.content, String.Empty)
                safeContent = safeContent.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n")
                messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{safeContent}""}}")
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

        ' 构建 JSON 请求体
        Dim messagesJson = String.Join(",", messages)

        ' 如果有工具，添加到请求中
        If toolsArray IsNot Nothing AndAlso toolsArray.Count > 0 Then
            Dim toolsJson = toolsArray.ToString(Formatting.None)
            ' 注意这里没有给tools加额外的引号
            Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"", ""tools"": {toolsJson}, ""messages"": [{messagesJson}], ""stream"": true}}"
            Return requestBody
        Else
            ' 直接使用JSON数组符号
            Dim requestBody = $"{{""model"": ""{ConfigSettings.ModelName}"",  ""messages"": [{messagesJson}], ""stream"": true}}"
            Return requestBody
        End If

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
                    If ChatBrowser.InvokeRequired Then
                        ChatBrowser.Invoke(Sub() ChatBrowser.ExecuteScriptAsync(jsCreate))
                    Else
                        Await ChatBrowser.ExecuteScriptAsync(jsCreate)
                    End If

                    ' 在前端 DOM 的 chat 节上设置 dataset.requestId，以便前端后续执行时可以把 requestUuid 发回
                    Dim jsSetMapping As String = $"(function(){{ var el = document.getElementById('chat-{responseUuid}'); if(el) el.dataset.requestId = '{requestUuid}'; }})();"
                    If ChatBrowser.InvokeRequired Then
                        ChatBrowser.Invoke(Sub() ChatBrowser.ExecuteScriptAsync(jsSetMapping))
                    Else
                        Await ChatBrowser.ExecuteScriptAsync(jsSetMapping)
                    End If

                    ' 处理流（后续逻辑不变，但使用 responseUuid 进行 flush 等操作）
                    Dim stringBuilder As New StringBuilder()
                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Dim buffer(102300) As Char
                            Dim readCount As Integer
                            Do
                                If stopReaderStream Then
                                    Debug.WriteLine("[Stream] 用户手动停止流读取")
                                    _currentMarkdownBuffer.Clear()
                                    allMarkdownBuffer.Clear()
                                    Exit Do
                                End If
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do
                                Dim chunk As String = New String(buffer, 0, readCount)
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)
                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    ProcessStreamChunk(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), responseUuid, originQuestion)
                                    stringBuilder.Clear()
                                End If
                            Loop
                            Debug.WriteLine("[Stream] 流接收完成")
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
            Dim lines As String() = rawChunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                line = line.Trim()
                If line = "[DONE]" Then
                    ' 在流结束时处理所有完成的工具调用
                    If _pendingToolCalls.Count > 0 Then
                        Debug.WriteLine("[DONE] 时发现未处理的工具调用，开始处理")
                        ProcessCompletedToolCalls(uuid, originQuestion)
                    End If
                    FlushBuffer("content", uuid) ' 最后刷新缓冲区
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
                    _currentMarkdownBuffer.Append(content)
                    FlushBuffer("content", uuid)
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
                Dim toolCallMessage = $"<br/>**正在调用工具: {toolName}**<br/>参数: `{argumentsObj.ToString(Formatting.None)}`<br/>"
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
                                         $"**调用参数:**<br/>```json{vbCrLf}{argumentsObj.ToString(Formatting.Indented)}{vbCrLf}```<br/>"

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
            promptBuilder.AppendLine(arguments.ToString(Formatting.Indented))
            promptBuilder.AppendLine("```")
            promptBuilder.AppendLine()
            promptBuilder.AppendLine("工具执行结果为：")
            promptBuilder.AppendLine("```json")
            promptBuilder.AppendLine(result.ToString(Formatting.Indented))
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

            Dim requestBody = requestObj.ToString(Formatting.None)

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
            _currentMarkdownBuffer.Append(result.ToString(Formatting.Indented))
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