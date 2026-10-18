Imports System.Diagnostics

''' <summary>
''' Coordinates chat request context construction and the pre-send history update.
''' </summary>
Public Class ChatRequestOrchestrator
    Private ReadOnly _conversationRuntime As IConversationRuntime
    Private ReadOnly _historyMessages As List(Of HistoryMessage)
    Private ReadOnly _selectionPendingMap As Dictionary(Of String, SelectionInfo)
    Private ReadOnly _getApplication As Func(Of ApplicationInfo)
    Private ReadOnly _manageHistoryMessageSize As Action

    Public Sub New(
        conversationRuntime As IConversationRuntime,
        historyMessages As List(Of HistoryMessage),
        selectionPendingMap As Dictionary(Of String, SelectionInfo),
        getApplication As Func(Of ApplicationInfo),
        manageHistoryMessageSize As Action)

        If conversationRuntime Is Nothing Then Throw New ArgumentNullException(NameOf(conversationRuntime))
        If historyMessages Is Nothing Then Throw New ArgumentNullException(NameOf(historyMessages))
        If selectionPendingMap Is Nothing Then Throw New ArgumentNullException(NameOf(selectionPendingMap))
        If getApplication Is Nothing Then Throw New ArgumentNullException(NameOf(getApplication))
        If manageHistoryMessageSize Is Nothing Then Throw New ArgumentNullException(NameOf(manageHistoryMessageSize))

        _conversationRuntime = conversationRuntime
        _historyMessages = historyMessages
        _selectionPendingMap = selectionPendingMap
        _getApplication = getApplication
        _manageHistoryMessageSize = manageHistoryMessageSize
    End Sub

    Public Function CreateRequestBody(
        uuid As String,
        question As String,
        systemPrompt As String,
        addHistory As Boolean,
        ByRef ragCountOut As Integer,
        ByRef contextTraceOut As ChatContextTrace) As String

        Dim context = BuildContext(uuid, question, systemPrompt, addHistory)
        Dim buildResult = _conversationRuntime.BuildRequest(context)

        ragCountOut = buildResult.RagCount
        contextTraceOut = buildResult.Trace

        If addHistory Then
            AddUserTurnToHistory(systemPrompt, question)
        End If

        Debug.WriteLine($"[ChatRequestOrchestrator] Runtime build complete, UsedContextBuilder={buildResult.UsedContextBuilder}, RagCount={ragCountOut}")
        Return buildResult.RequestBody
    End Function

    Private Function BuildContext(uuid As String, question As String, systemPrompt As String, addHistory As Boolean) As ChatRequestContext
        Return New ChatRequestContext With {
            .RequestUuid = uuid,
            .Question = question,
            .SystemPrompt = systemPrompt,
            .AddHistory = addHistory,
            .ModelName = ConfigSettings.ModelName,
            .Platform = ConfigSettings.platform,
            .ApiUrl = ConfigSettings.ApiUrl,
            .ReasoningMode = ConfigSettings.ReasoningMode,
            .Stream = True,
            .AppInfo = _getApplication(),
            .HistoryMessages = _historyMessages,
            .SelectionPendingMap = _selectionPendingMap,
            .UseContextBuilder = MemoryConfig.UseContextBuilder,
            .EnableMemory = MemoryConfig.EnableUserProfile OrElse MemoryConfig.RagTopN > 0
        }
    End Function

    Private Sub AddUserTurnToHistory(systemPrompt As String, question As String)
        Dim existingSystem = _historyMessages.FirstOrDefault(Function(m) m.role = "system")
        If existingSystem IsNot Nothing Then _historyMessages.Remove(existingSystem)

        _historyMessages.Insert(0, New HistoryMessage With {
            .role = "system",
            .content = systemPrompt
        })
        _historyMessages.Add(New HistoryMessage With {
            .role = "user",
            .content = question
        })

        ' 这里只维护发给模型的上下文；会话历史的 user 事件由 BaseChatControl.Send 单点落库
        _manageHistoryMessageSize()
    End Sub
End Class
