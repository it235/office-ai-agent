' ShareRibbon\Controls\Services\MessageService.vb
' 消息服务：处理 WebView2 消息路由和各类消息处理

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 消息服务，负责处理 WebView2 消息路由和各类消息处理
''' </summary>
Public Class MessageService
    Private ReadOnly _stateService As ChatStateService
    Private ReadOnly _fileParserService As FileParserService
    Private ReadOnly _codeExecutionService As CodeExecutionService
    Private ReadOnly _getApplication As Func(Of ApplicationInfo)
    Private ReadOnly _executeScript As Func(Of String, System.Threading.Tasks.Task)

    ' 事件委托
    Public Event SendMessageRequested As EventHandler(Of SendMessageEventArgs)
    Public Event CheckedChanged As EventHandler(Of CheckedChangeEventArgs)
    Public Event SettingsSaveRequested As EventHandler(Of SettingsEventArgs)
    Public Event RevisionApplyRequested As EventHandler(Of RevisionEventArgs)
    Public Event StopStreamRequested As EventHandler

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New(
            stateService As ChatStateService,
            fileParserService As FileParserService,
            codeExecutionService As CodeExecutionService,
            getApplication As Func(Of ApplicationInfo),
            executeScript As Func(Of String, System.Threading.Tasks.Task))

        _stateService = stateService
        _fileParserService = fileParserService
        _codeExecutionService = codeExecutionService
        _getApplication = getApplication
        _executeScript = executeScript
    End Sub

#Region "消息路由"

    ''' <summary>
    ''' 处理 WebView2 消息
    ''' </summary>
    Public Sub HandleWebMessage(jsonMessage As String)
        Try
            Dim jsonDoc As JObject = JObject.Parse(jsonMessage)
            Dim messageType As String = jsonDoc("type").ToString()

            Select Case messageType
                Case "checkedChange"
                    HandleCheckedChange(jsonDoc)
                Case "sendMessage"
                    HandleSendMessage(jsonDoc)
                Case "stopMessage"
                    RaiseEvent StopStreamRequested(Me, EventArgs.Empty)
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
                    _stateService.ClearHistory()
                Case "acceptAnswer"
                    HandleAcceptAnswer(jsonDoc)
                Case "rejectAnswer"
                    HandleRejectAnswer(jsonDoc)
                Case "applyRevisionAll"
                    RaiseEvent RevisionApplyRequested(Me, New RevisionEventArgs With {
                            .Type = "all",
                            .JsonDoc = jsonDoc
                        })
                Case "applyRevisionSegment"
                    RaiseEvent RevisionApplyRequested(Me, New RevisionEventArgs With {
                            .Type = "segment",
                            .JsonDoc = jsonDoc
                        })
                Case "applyDocumentPlanItem"
                    RaiseEvent RevisionApplyRequested(Me, New RevisionEventArgs With {
                            .Type = "documentPlan",
                            .JsonDoc = jsonDoc
                        })
                Case "applyRevisionAccept"
                    RaiseEvent RevisionApplyRequested(Me, New RevisionEventArgs With {
                            .Type = "accept",
                            .JsonDoc = jsonDoc
                        })
                Case "applyRevisionReject"
                    RaiseEvent RevisionApplyRequested(Me, New RevisionEventArgs With {
                            .Type = "reject",
                            .JsonDoc = jsonDoc
                        })
                Case Else
                    Debug.WriteLine($"未知消息类型: {messageType}")
            End Select
        Catch ex As Exception
            Debug.WriteLine($"处理消息出错: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "消息处理"

    ''' <summary>
    ''' 处理复选框变更
    ''' </summary>
    Private Sub HandleCheckedChange(jsonDoc As JObject)
        Dim prop As String = jsonDoc("property").ToString()
        Dim isChecked As Boolean = Boolean.Parse(jsonDoc("isChecked").ToString())

        RaiseEvent CheckedChanged(Me, New CheckedChangeEventArgs With {
                .PropertyName = prop,
                .IsChecked = isChecked
            })
    End Sub

    ''' <summary>
    ''' 处理发送消息
    ''' </summary>
    Private Sub HandleSendMessage(jsonDoc As JObject)
        Dim messageValue As JToken = jsonDoc("value")
        Dim question As String
        Dim filePaths As New List(Of String)()
        Dim selectedContents As New List(Of SendMessageReferenceContentItem)()

        If messageValue.Type = JTokenType.Object Then
            question = messageValue("text")?.ToString()

            If messageValue("filePaths") IsNot Nothing AndAlso messageValue("filePaths").Type = JTokenType.Array Then
                filePaths = messageValue("filePaths").ToObject(Of List(Of String))()
            End If

            If messageValue("selectedContent") IsNot Nothing AndAlso messageValue("selectedContent").Type = JTokenType.Array Then
                Try
                    selectedContents = messageValue("selectedContent").ToObject(Of List(Of SendMessageReferenceContentItem))()
                Catch ex As Exception
                    Debug.WriteLine($"Error deserializing selectedContent: {ex.Message}")
                End Try
            End If
        Else
            Debug.WriteLine("HandleSendMessage: Invalid message format")
            Return
        End If

        If String.IsNullOrEmpty(question) AndAlso
               (filePaths Is Nothing OrElse filePaths.Count = 0) AndAlso
               (selectedContents Is Nothing OrElse selectedContents.Count = 0) Then
            Return
        End If

        ' 记录第一个问题
        _stateService.RecordFirstQuestion(question)

        RaiseEvent SendMessageRequested(Me, New SendMessageEventArgs With {
                .Question = question,
                .FilePaths = filePaths,
                .SelectedContents = selectedContents
            })
    End Sub

    ''' <summary>
    ''' 处理执行代码
    ''' </summary>
    Private Sub HandleExecuteCode(jsonDoc As JObject)
        Dim code As String = jsonDoc("code").ToString()
        Dim preview As Boolean = Boolean.Parse(jsonDoc("executecodePreview"))
        Dim language As String = jsonDoc("language").ToString()
        _codeExecutionService.ExecuteCode(code, language, preview)
    End Sub

    ''' <summary>
    ''' 处理保存设置
    ''' </summary>
    Private Sub HandleSaveSettings(jsonDoc As JObject)
        RaiseEvent SettingsSaveRequested(Me, New SettingsEventArgs With {
                .TopicRandomness = CDbl(jsonDoc("topicRandomness")),
                .ContextLimit = CInt(jsonDoc("contextLimit")),
                .SelectedCellChecked = CBool(jsonDoc("selectedCell")),
                .SettingsScrollChecked = CBool(jsonDoc("settingsScroll")),
                .ChatMode = jsonDoc("chatMode").ToString(),
                .ExecuteCodePreview = CBool(jsonDoc("executeCodePreview"))
            })
    End Sub

#End Region

#Region "历史记录"

    ''' <summary>
    ''' 处理获取历史文件列表
    ''' </summary>
    Private Sub HandleGetHistoryFiles()
        Try
            Dim historyDir As String = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    ConfigSettings.OfficeAiAppDataFolder)

            Dim historyFiles As New List(Of Object)()

            If Directory.Exists(historyDir) Then
                Dim files As String() = Directory.GetFiles(historyDir, "saved_chat_*.html")

                For Each filePath As String In files
                    Try
                        Dim fileInfo As New FileInfo(filePath)
                        historyFiles.Add(New With {
                                .fileName = fileInfo.Name,
                                .fullPath = fileInfo.FullName,
                                .size = fileInfo.Length,
                                .lastModified = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                            })
                    Catch
                    End Try
                Next
            End If

            Dim jsonResult As String = JsonConvert.SerializeObject(historyFiles)
            _executeScript($"setHistoryFilesList({jsonResult});")
        Catch ex As Exception
            _executeScript("setHistoryFilesList([]);")
        End Try
    End Sub

    ''' <summary>
    ''' 处理打开历史文件
    ''' </summary>
    Private Sub HandleOpenHistoryFile(jsonDoc As JObject)
        Try
            Dim filePath As String = jsonDoc("filePath").ToString()

            If File.Exists(filePath) Then
                Process.Start(New ProcessStartInfo() With {
                        .FileName = filePath,
                        .UseShellExecute = True
                    })
                GlobalStatusStrip.ShowInfo("已在浏览器中打开历史记录")
            Else
                GlobalStatusStrip.ShowWarning("历史记录文件不存在")
            End If
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("打开历史记录失败: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "MCP 设置"

    ''' <summary>
    ''' 处理获取 MCP 连接
    ''' </summary>
    Private Sub HandleGetMcpConnections()
        Try
            Dim connections = MCPConnectionManager.LoadConnections()
            Dim enabledConnections = connections.Where(Function(c) c.IsActive).ToList()

            Dim chatSettings As New ChatSettings(_getApplication())
            Dim enabledMcpList = chatSettings.EnabledMcpList

            Dim connectionsJson = JsonConvert.SerializeObject(enabledConnections)
            Dim enabledListJson = JsonConvert.SerializeObject(enabledMcpList)

            Dim mcpSupported As Boolean = ConfigSettings.mcpable

            _executeScript($"renderMcpConnections({connectionsJson}, {enabledListJson},{mcpSupported.ToString().ToLower()});")
        Catch ex As Exception
            Debug.WriteLine($"获取MCP连接列表失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理保存 MCP 设置
    ''' </summary>
    Private Sub HandleSaveMcpSettings(jsonDoc As JObject)
        Try
            Dim enabledList As List(Of String) = jsonDoc("enabledList").ToObject(Of List(Of String))()

            Dim chatSettings As New ChatSettings(_getApplication())
            chatSettings.SaveEnabledMcpList(enabledList)

            GlobalStatusStrip.ShowInfo("MCP设置已保存")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("保存MCP设置失败")
        End Try
    End Sub

#End Region

#Region "接受/拒绝答案"

    ''' <summary>
    ''' 处理接受答案
    ''' </summary>
    Private Sub HandleAcceptAnswer(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid")?.ToString(), String.Empty)
            Dim content As String = If(jsonDoc("content")?.ToString(), String.Empty)
            Debug.WriteLine($"用户接受回答: UUID={uuid}")
            GlobalStatusStrip.ShowInfo("用户已接受 AI 回答")
        Catch ex As Exception
            Debug.WriteLine($"HandleAcceptAnswer 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理拒绝答案
    ''' </summary>
    Private Sub HandleRejectAnswer(jsonDoc As JObject)
        Try
            Dim uuid As String = If(jsonDoc("uuid")?.ToString(), String.Empty)
            Dim reason As String = If(jsonDoc("reason")?.ToString(), String.Empty)

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
            refinementPrompt.AppendLine("2) Plan：简短列出修正步骤；")
            refinementPrompt.AppendLine("3) Answer：给出修正后的答案；")
            refinementPrompt.AppendLine("4) Clarifying Questions：如需更多信息，请列出问题。")

            _stateService.ManageHistorySize()

            ' 触发发送改进请求
            RaiseEvent SendMessageRequested(Me, New SendMessageEventArgs With {
                    .Question = refinementPrompt.ToString(),
                    .FilePaths = New List(Of String)(),
                    .SelectedContents = New List(Of SendMessageReferenceContentItem)()
                })

            GlobalStatusStrip.ShowInfo("已触发改进请求")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("触发改进请求时出错")
        End Try
    End Sub

#End Region

End Class

#Region "事件参数类"

''' <summary>
''' 发送消息事件参数
''' </summary>
Public Class SendMessageEventArgs
    Inherits EventArgs

    Public Property Question As String
    Public Property FilePaths As List(Of String)
    Public Property SelectedContents As List(Of SendMessageReferenceContentItem)
End Class

''' <summary>
''' 复选框变更事件参数
''' </summary>
Public Class CheckedChangeEventArgs
    Inherits EventArgs

    Public Property PropertyName As String
    Public Property IsChecked As Boolean
End Class

''' <summary>
''' 设置事件参数
''' </summary>
Public Class SettingsEventArgs
    Inherits EventArgs

    Public Property TopicRandomness As Double
    Public Property ContextLimit As Integer
    Public Property SelectedCellChecked As Boolean
    Public Property SettingsScrollChecked As Boolean
    Public Property ChatMode As String
    Public Property ExecuteCodePreview As Boolean
End Class

''' <summary>
''' 修订事件参数
''' </summary>
Public Class RevisionEventArgs
    Inherits EventArgs

    Public Property Type As String
    Public Property JsonDoc As JObject
End Class

#End Region
