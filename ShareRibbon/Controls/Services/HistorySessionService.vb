Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 历史会话服务：会话列表、加载、新建、重命名和删除。
''' </summary>
Public Class HistorySessionService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _chatStateService As ChatStateService
    Private ReadOnly _getAppType As Func(Of String)
    Private ReadOnly _getDocument As Func(Of SessionDocument)
    Private ReadOnly _invokeOnUiThread As Action(Of Action)

    Public Sub New(
        executeScript As Func(Of String, Task),
        chatStateService As ChatStateService,
        getAppType As Func(Of String),
        getDocument As Func(Of SessionDocument),
        invokeOnUiThread As Action(Of Action))

        _executeScript = executeScript
        _chatStateService = chatStateService
        _getAppType = getAppType
        _getDocument = getDocument
        _invokeOnUiThread = invokeOnUiThread
    End Sub

    Private Function ResolveDocument() As SessionDocument
        If _getDocument Is Nothing Then Return Nothing
        Try
            Return _getDocument()
        Catch ex As Exception
            AppLogger.Debug("HistorySessionService", "读取当前文档标识失败，列表按通用会话展示", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 获取近期会话列表（由权威历史表 conversation_event 聚合），供历史侧边栏展示。
    ''' scope=document 只看当前文档；默认的 scope=all 也把当前文档排在最前，
    ''' 因为"我正在写的这篇文档之前聊过什么"才是用户打开面板想先看到的东西。
    ''' </summary>
    Public Async Sub HandleGetSessionList(Optional scope As String = ConversationRepository.SessionScopeAll)
        Dim errorScript As String = Nothing
        Try
            Dim limit As Integer = 50
            Dim doc = ResolveDocument()
            Dim documentId = ""
            If doc IsNot Nothing Then documentId = If(doc.Id, "")
            Dim documentName = SessionDocument.DisplayName(doc)
            Dim onlyDocument = String.Equals(scope, ConversationRepository.SessionScopeDocument, StringComparison.OrdinalIgnoreCase)
            Dim summaries = ConversationRepository.GetRecentSessions(
                limit, _getAppType(), documentId,
                If(onlyDocument, ConversationRepository.SessionScopeDocument, ConversationRepository.SessionScopeAll))

            Dim list As New List(Of Object)()
            For Each s In summaries
                list.Add(New With {
                    .sessionId = s.SessionId,
                    .title = If(String.IsNullOrEmpty(s.Title), "会话", s.Title),
                    .snippet = If(String.IsNullOrEmpty(s.Snippet), "", s.Snippet),
                    .createdAt = s.CreatedAt,
                    .documentName = If(String.IsNullOrEmpty(s.DocumentName), "", s.DocumentName),
                    .isCurrentDocument = Not String.IsNullOrEmpty(documentId) AndAlso s.DocumentId = documentId,
                    .isRenamed = s.IsRenamed
                })
            Next

            Dim payload = New With {
                .items = list,
                .scope = If(onlyDocument, ConversationRepository.SessionScopeDocument, ConversationRepository.SessionScopeAll),
                .documentName = documentName,
                .hasDocument = Not String.IsNullOrEmpty(documentId)
            }
            Dim jsonResult As String = JsonConvert.SerializeObject(payload)
            Await _executeScript($"setHistoryFilesList({jsonResult});")
        Catch ex As Exception
            AppLogger.Warn("HistorySessionService", "HandleGetSessionList 失败", ex)
            errorScript = "setHistoryFilesList({""items"":[],""scope"":""all"",""documentName"":"""",""hasDocument"":false});"
        End Try
        If errorScript IsNot Nothing Then
            Await _executeScript(errorScript)
        End If
    End Sub

    ''' <summary>
    ''' 重命名会话：只改自定义标题，不动消息内容。改完立刻回推列表，用户不需要重开面板。
    ''' </summary>
    Public Async Sub HandleRenameSession(jsonDoc As JObject, scope As String)
        Try
            Dim sessionId As String = jsonDoc?("sessionId")?.ToString()
            If String.IsNullOrEmpty(sessionId) Then Return
            Dim title As String = jsonDoc?("title")?.ToString()

            Dim renamed = Await Task.Run(Function() ConversationRepository.RenameSession(sessionId, title))
            If Not renamed Then
                GlobalStatusStrip.ShowWarning("重命名失败：找不到该会话")
                Return
            End If
            GlobalStatusStrip.ShowInfo("已重命名会话")
            HandleGetSessionList(scope)
        Catch ex As Exception
            AppLogger.Warn("HistorySessionService", "HandleRenameSession 失败", ex)
            GlobalStatusStrip.ShowWarning("重命名会话失败")
        End Try
    End Sub

    ''' <summary>
    ''' 加载指定会话到当前 Chat 并渲染消息
    ''' </summary>
    Public Async Sub HandleLoadSession(jsonDoc As JObject)
        Try
            Dim sessionId As String = jsonDoc("sessionId")?.ToString()
            If String.IsNullOrEmpty(sessionId) Then Return
            _chatStateService.SwitchToSession(sessionId)
            Dim messages As New List(Of Object)()
            For Each m In _chatStateService.HistoryMessages
                If m.role = "user" OrElse m.role = "assistant" Then
                    messages.Add(New With {.role = m.role, .content = m.content, .createTime = m.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")})
                End If
            Next
            Dim jsonResult As String = JsonConvert.SerializeObject(messages)
            Await _executeScript($"setChatMessages({jsonResult});")
        Catch ex As Exception
            Debug.WriteLine("HandleLoadSession 失败: " & ex.Message)
            GlobalStatusStrip.ShowWarning("加载会话失败")
        End Try
    End Sub

    ''' <summary>
    ''' 新建会话：清空状态并清空聊天区域
    ''' </summary>
    Public Async Sub HandleNewSession()
        Try
            _chatStateService.StartNewSession()
            Await _executeScript("if(typeof clearChatContent==='function')clearChatContent();")
            GlobalStatusStrip.ShowInfo("已新建会话")
        Catch ex As Exception
            Debug.WriteLine("HandleNewSession 失败: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 删除指定会话：级联清理其持久化行（事件/历史/摘要/记忆/队列/运行轨迹），
    ''' 若删除的是当前会话则切到新会话并清空聊天区，最后刷新历史列表。
    ''' </summary>
    Public Async Sub HandleDeleteSession(jsonDoc As JObject, scope As String)
        Try
            Dim sessionId As String = jsonDoc?("sessionId")?.ToString()
            If String.IsNullOrEmpty(sessionId) Then Return
            Await Task.Run(Sub() ConversationRepository.DeleteSessionCascade(sessionId))

            If String.Equals(sessionId, _chatStateService.CurrentSessionId, StringComparison.OrdinalIgnoreCase) Then
                _chatStateService.StartNewSession()
                Await _executeScript("if(typeof clearChatContent==='function')clearChatContent();")
            End If

            GlobalStatusStrip.ShowInfo("已删除会话")
            HandleGetSessionList(scope)
        Catch ex As Exception
            AppLogger.Warn("HistorySessionService", "HandleDeleteSession 失败", ex)
            GlobalStatusStrip.ShowWarning("删除会话失败")
        End Try
    End Sub

End Class
