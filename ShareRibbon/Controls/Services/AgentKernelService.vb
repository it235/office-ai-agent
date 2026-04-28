Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' AgentKernel 服务：统一智能体服务，替代 RalphAgentService + RalphLoopService
''' 封装 AgentKernel 的初始化和事件处理，提供与 BaseChatControl 兼容的接口
''' </summary>
Public Class AgentKernelService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _escapeJs As Func(Of String, String)
    Private ReadOnly _sendAiRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
    Private ReadOnly _executeCode As Action(Of String, String, Boolean)
    Private ReadOnly _chatStateService As ChatStateService
    Private ReadOnly _historyMessages As List(Of HistoryMessage)
    Private ReadOnly _manageHistorySize As Action
    Private ReadOnly _getOfficeAppType As Func(Of String)

    ' 统一的 AgentKernel 实例
    Private _agentKernel As Agent.AgentKernel

    ' Agent 状态字段（供 BaseChatControl 访问）
    Public Property AgentThinkingUuid As String = Nothing
    Public Property AgentOriginalUserRequest As String = Nothing
    Public Property AgentFullUserMessage As String = Nothing
    Public Property CurrentAgentSessionId As String = Nothing

    ' 审批等待器（用于 LoopEngine 的 OnRequestApproval 事件）
    Private _approvalTcs As TaskCompletionSource(Of Boolean)

    ''' <summary>暴露底层 AgentKernel 供外部直接调用</summary>
    Public ReadOnly Property Kernel As Agent.AgentKernel
        Get
            EnsureInitialized()
            Return _agentKernel
        End Get
    End Property

    Public Sub New(
        executeScript As Func(Of String, Task),
        escapeJs As Func(Of String, String),
        sendAiRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String)),
        executeCode As Action(Of String, String, Boolean),
        chatStateService As ChatStateService,
        historyMessages As List(Of HistoryMessage),
        manageHistorySize As Action,
        getOfficeAppType As Func(Of String))

        _executeScript = executeScript
        _escapeJs = escapeJs
        _sendAiRequest = sendAiRequest
        _executeCode = executeCode
        _chatStateService = chatStateService
        _historyMessages = historyMessages
        _manageHistorySize = manageHistorySize
        _getOfficeAppType = getOfficeAppType
    End Sub

    ''' <summary>
    ''' 确保 AgentKernel 已初始化
    ''' </summary>
    Private Sub EnsureInitialized()
        If _agentKernel IsNot Nothing Then Return

        _agentKernel = New Agent.AgentKernel()

        ' 绑定 AI 请求委托
        _agentKernel.SendAIRequest = Async Function(prompt, system, history)
                                          Return Await _sendAiRequest(prompt, system, history)
                                      End Function

        ' 绑定代码执行委托
        _agentKernel.ExecuteCode = Sub(code, lang, preview)
                                       _executeCode(code, lang, preview)
                                   End Sub

        ' 绑定事件
        AddHandler _agentKernel.OnStatusChanged, AddressOf OnKernelStatusChanged
        AddHandler _agentKernel.OnIterationUpdate, AddressOf OnKernelIterationUpdate
        AddHandler _agentKernel.OnStepCompleted, AddressOf OnKernelStepCompleted
        AddHandler _agentKernel.OnRequestApproval, AddressOf OnKernelRequestApproval
        AddHandler _agentKernel.OnPlanGenerated, AddressOf OnKernelPlanGenerated
        AddHandler _agentKernel.OnCompleted, AddressOf OnKernelCompleted

        ' 加载工具和技能
        _agentKernel.Initialize()
    End Sub

#Region "Public Methods"

    ''' <summary>
    ''' 启动统一 Agent 任务（替代 StartAgent 和 StartLoop）
    ''' </summary>
    Public Async Function StartAgentAsync(userRequest As String, appType As String, currentContent As String,
                                           historyMessages As List(Of Tuple(Of String, String))) As Task(Of Boolean)
        Try
            EnsureInitialized()

            ' 保存原始请求
            AgentOriginalUserRequest = userRequest
            AgentFullUserMessage = userRequest
            CurrentAgentSessionId = Guid.NewGuid().ToString()

            ' 显示思考状态
            ShowThinkingStatus()

            ' 注入历史消息到 AgentMemory（预加载会话上下文）
            If historyMessages IsNot Nothing AndAlso historyMessages.Count > 0 Then
                For Each msg In historyMessages
                    _agentKernel.AddHistoryMessage(msg.Item1, msg.Item2)
                Next
            End If

            ' 执行 Agent 任务
            Dim result = Await _agentKernel.ExecuteAsync(userRequest, appType, currentContent)

            Return result.Success
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] StartAgentAsync 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 终止当前 Agent 任务
    ''' </summary>
    Public Sub AbortAgent()
        Try
            ' 如果有正在等待的审批，设置为 false
            If _approvalTcs IsNot Nothing AndAlso Not _approvalTcs.Task.IsCompleted Then
                _approvalTcs.TrySetResult(False)
            End If

            ' 清除状态
            AgentThinkingUuid = Nothing
            AgentOriginalUserRequest = Nothing
            AgentFullUserMessage = Nothing
            CurrentAgentSessionId = Nothing

            _executeScript($"completeAgent('{CurrentAgentSessionId}', false, '已终止')")

            GlobalStatusStrip.ShowInfo("已终止Agent")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] AbortAgent 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 用户批准当前计划或步骤
    ''' </summary>
    Public Sub Approve()
        If _approvalTcs IsNot Nothing AndAlso Not _approvalTcs.Task.IsCompleted Then
            _approvalTcs.TrySetResult(True)
        End If
    End Sub

    ''' <summary>
    ''' 用户拒绝当前计划或步骤
    ''' </summary>
    Public Sub Reject()
        If _approvalTcs IsNot Nothing AndAlso Not _approvalTcs.Task.IsCompleted Then
            _approvalTcs.TrySetResult(False)
        End If
    End Sub

#End Region

#Region "Event Handlers"

    ''' <summary>
    ''' 处理状态变更事件
    ''' </summary>
    Private Sub OnKernelStatusChanged(status As String)
        Try
            GlobalStatusStrip.ShowInfo(status)

            ' 更新思考状态 div
            If Not String.IsNullOrEmpty(AgentThinkingUuid) Then
                _executeScript($"var thinkingDiv = document.getElementById('content-{AgentThinkingUuid}'); if(thinkingDiv) thinkingDiv.innerHTML = '<div style=""padding: 8px 0; color: #2563eb;"">{_escapeJs(status)}</div>';")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnStatusChanged 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理 ReAct 迭代更新事件
    ''' </summary>
    Private Sub OnKernelIterationUpdate(iteration As Agent.ReActIteration)
        Try
            If iteration Is Nothing Then Return

            Dim iterationJson = $"{{""index"":{iteration.Index},""thought"":""{_escapeJs(iteration.Thought)}"",""action"":""{_escapeJs(If(iteration.Action?.ToolId, ""))}"",""observation"":""{_escapeJs(iteration.Observation)}""}}"
            _executeScript($"updateAgentIteration('{CurrentAgentSessionId}', {iterationJson})")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnIterationUpdate 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理步骤完成事件
    ''' </summary>
    Private Sub OnKernelStepCompleted(stepIndex As Integer, success As Boolean, message As String)
        Try
            Dim stepStatus = If(success, "completed", "failed")
            _executeScript($"updateAgentStep('{CurrentAgentSessionId}', {stepIndex}, '{stepStatus}', '{_escapeJs(message)}')")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnStepCompleted 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理审批请求事件
    ''' </summary>
    Private Sub OnKernelRequestApproval(message As String, callback As Action(Of Boolean))
        Try
            _approvalTcs = New TaskCompletionSource(Of Boolean)()

            ' 显示审批 UI
            _executeScript($"showAgentApproval('{CurrentAgentSessionId}', '{_escapeJs(message)}')")

            ' 等待用户决策
            Task.Run(Async Function()
                         Dim approved = Await _approvalTcs.Task
                         callback(approved)
                     End Function)
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnRequestApproval 出错: {ex.Message}")
            callback(False)
        End Try
    End Sub

    ''' <summary>
    ''' 处理计划生成事件
    ''' </summary>
    Private Sub OnKernelPlanGenerated(plan As Agent.ExecutionPlan)
        Try
            If plan Is Nothing Then Return

            ' 构建步骤 JSON
            Dim stepsJson As New StringBuilder()
            stepsJson.Append("[")
            For i = 0 To plan.Steps.Count - 1
                If i > 0 Then stepsJson.Append(",")
                Dim s = plan.Steps(i)
                stepsJson.Append($"{{""description"":""{_escapeJs(s.Description)}"",""code"":""{_escapeJs(If(s.Code, ""))}"",""language"":""{s.Language}"",""status"":""pending""}}")
            Next
            stepsJson.Append("]")

            Dim planJson = $"{{""sessionId"":""{CurrentAgentSessionId}"",""understanding"":""{_escapeJs(If(plan.Understanding, ""))}"",""steps"":{stepsJson.ToString()},""summary"":""{_escapeJs(If(plan.Summary, ""))}"",""replaceThinkingUuid"":""{AgentThinkingUuid}""}}"

            _executeScript($"showAgentPlanCard({planJson})")
            _executeScript("var planningCard = document.getElementById('planning-status-card'); if(planningCard) planningCard.remove();")
            AgentThinkingUuid = Nothing
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnPlanGenerated 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理 Agent 完成事件
    ''' </summary>
    Private Sub OnKernelCompleted(result As Agent.AgentResult)
        Try
            Dim userMsgForHistory = If(Not String.IsNullOrWhiteSpace(AgentFullUserMessage), AgentFullUserMessage, AgentOriginalUserRequest)

            If Not String.IsNullOrWhiteSpace(userMsgForHistory) Then
                _historyMessages.Add(New HistoryMessage With {
                    .role = "user",
                    .content = userMsgForHistory
                })
                _manageHistorySize()
                _chatStateService?.AddMessage("user", userMsgForHistory)
            End If

            Dim assistantReply = If(String.IsNullOrEmpty(result.Message), "任务完成", result.Message)
            _historyMessages.Add(New HistoryMessage With {
                .role = "assistant",
                .content = assistantReply
            })
            _manageHistorySize()
            _chatStateService?.AddMessage("assistant", assistantReply)

            MemoryService.SaveConversationTurnAsync(userMsgForHistory, assistantReply, _chatStateService?.CurrentSessionId, _getOfficeAppType())

            AgentOriginalUserRequest = Nothing
            AgentFullUserMessage = Nothing

            _executeScript($"completeAgent('{CurrentAgentSessionId}', {result.Success.ToString().ToLower()}, '{_escapeJs(result.Message)}')")
            CurrentAgentSessionId = Nothing
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnCompleted 出错: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "Private Helpers"

    ''' <summary>
    ''' 在聊天界面显示思考状态
    ''' </summary>
    Private Sub ShowThinkingStatus()
        Try
            If String.IsNullOrEmpty(AgentThinkingUuid) Then
                AgentThinkingUuid = Guid.NewGuid().ToString()
            End If

            Dim timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            _executeScript($"createChatSection('AI', '{timestamp}', '{AgentThinkingUuid}')")
            _executeScript($"var thinkingDiv = document.getElementById('content-{AgentThinkingUuid}'); if(thinkingDiv) thinkingDiv.innerHTML = '<div class=""thinking-indicator""><div class=""thinking-dots""><span></span><span></span><span></span></div><span style=""margin-left: 12px; color: #6c757d;"">正在分析您的需求...</span></div>';")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] ShowThinkingStatus 出错: {ex.Message}")
        End Try
    End Sub

#End Region

End Class
