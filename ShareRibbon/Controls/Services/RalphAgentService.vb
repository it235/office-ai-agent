Imports System.Collections.Generic
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' Ralph Agent 服务：智能助手规划、执行、中止
''' </summary>
Public Class RalphAgentService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _escapeJs As Func(Of String, String)
    Private ReadOnly _sendAiRequest As Func(Of String, String, List(Of HistoryMessage), String, Task(Of String))
    Private ReadOnly _executeCode As Action(Of String, String, Boolean)
    Private ReadOnly _chatStateService As ChatStateService
    Private ReadOnly _historyMessages As List(Of HistoryMessage)
    Private ReadOnly _manageHistorySize As Action
    Private ReadOnly _getOfficeAppType As Func(Of String)

    ' 内部 Agent 控制器
    Private _ralphAgentController As RalphAgentController

    ' Agent 状态字段（供 BaseChatControl 中留下的方法访问）
    Public Property AgentThinkingUuid As String = Nothing
    Public Property AgentOriginalUserRequest As String = Nothing
    Public Property AgentFullUserMessage As String = Nothing
    Public Property CurrentAgentSessionId As String = Nothing

    ' Agent 响应收集（供 BaseChatControl 的 SendAndGetResponse 访问）
    Public Property AgentResponseBuffer As StringBuilder = Nothing
    Public Property AgentResponseUuid As String = Nothing
    Public Property AgentResponseCompleted As Boolean = False

    ''' <summary>暴露底层控制器供 BaseChatControl 的 HandleStartAgentCore 使用</summary>
    Public ReadOnly Property Controller As RalphAgentController
        Get
            Return _ralphAgentController
        End Get
    End Property

    Public Sub New(
        executeScript As Func(Of String, Task),
        escapeJs As Func(Of String, String),
        sendAiRequest As Func(Of String, String, List(Of HistoryMessage), String, Task(Of String)),
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
    ''' 初始化 Agent 控制器并设置所有回调
    ''' </summary>
    Public Sub InitializeAgentController(agentThinkingUuid As String)
        If _ralphAgentController Is Nothing Then
            _ralphAgentController = New RalphAgentController()

            _ralphAgentController.OnStatusChanged = Sub(status)
                                                        Dim currentChatMode As String = ChatSettings.chatMode
                                                        If currentChatMode = "agent" AndAlso Not String.IsNullOrEmpty(AgentThinkingUuid) Then
                                                            _executeScript($"var thinkingDiv = document.getElementById('content-{AgentThinkingUuid}'); if(thinkingDiv) thinkingDiv.innerHTML = '<div style=""padding: 8px 0; color: #2563eb;"">⚡ {_escapeJs(status)}</div>';")
                                                        End If
                                                    End Sub

            _ralphAgentController.OnStepStarted = Sub(stepIndex, desc)
                                                      _executeScript($"updateAgentStep('{CurrentAgentSessionId}', {stepIndex}, 'running', '')")
                                                  End Sub

            _ralphAgentController.OnStepCompleted = Sub(stepIndex, success, msg)
                                                        Dim stepStatus = If(success, "completed", "failed")
                                                        _executeScript($"updateAgentStep('{CurrentAgentSessionId}', {stepIndex}, '{stepStatus}', '{_escapeJs(msg)}')")
                                                    End Sub

            _ralphAgentController.OnAgentCompleted = Sub(success)
                                                         Dim userMsgForHistory = If(Not String.IsNullOrWhiteSpace(AgentFullUserMessage), AgentFullUserMessage, AgentOriginalUserRequest)

                                                         If Not String.IsNullOrWhiteSpace(userMsgForHistory) Then
                                                             _historyMessages.Add(New HistoryMessage With {
                                                                 .role = "user",
                                                                 .content = userMsgForHistory
                                                             })
                                                             _manageHistorySize()
                                                             _chatStateService?.AddMessage("user", userMsgForHistory)
                                                             Debug.WriteLine($"[Agent] 已保存完整用户消息到历史，长度: {userMsgForHistory.Length}")
                                                         End If

                                                         Dim session = _ralphAgentController.GetCurrentSession()
                                                         If session IsNot Nothing Then
                                                             Dim assistantReply = If(String.IsNullOrEmpty(session.Summary), "任务完成", session.Summary)
                                                             If Not String.IsNullOrEmpty(session.Understanding) Then
                                                                 assistantReply = session.Understanding & vbCrLf & vbCrLf & assistantReply
                                                             End If
                                                             _historyMessages.Add(New HistoryMessage With {
                                                                 .role = "assistant",
                                                                 .content = assistantReply
                                                             })
                                                             _manageHistorySize()
                                                             _chatStateService?.AddMessage("assistant", assistantReply)
                                                             Debug.WriteLine($"[Agent] 已保存Assistant回复到历史")

                                                             MemoryService.SaveConversationTurnAsync(userMsgForHistory, assistantReply, _chatStateService.CurrentSessionId, _getOfficeAppType())
                                                         End If

                                                         AgentOriginalUserRequest = Nothing
                                                         AgentFullUserMessage = Nothing

                                                         ' 使用保存的 sessionId 确保一致性
                                                         _executeScript($"completeAgent('{CurrentAgentSessionId}', {success.ToString().ToLower()}, '')")
                                                         CurrentAgentSessionId = Nothing
                                                     End Sub

            _ralphAgentController.SendAIRequest = Async Function(prompt, sysPrompt, historyMsgs)
                                                      Return Await _sendAiRequest(prompt, sysPrompt, historyMsgs, agentThinkingUuid)
                                                  End Function

            _ralphAgentController.ExecuteCode = Sub(code, lang, preview)
                                                    _executeCode(code, lang, preview)
                                                End Sub
        End If
    End Sub

    ''' <summary>
    ''' 显示 Agent 规划卡片（替换思考消息）
    ''' </summary>
    Public Sub ShowAgentPlanCard(session As RalphAgentSession)
        Try
            ' 保存 sessionId 以供后续回调使用
            CurrentAgentSessionId = session.Id

            Dim stepsJson As New StringBuilder()
            stepsJson.Append("[")
            For i = 0 To session.Steps.Count - 1
                If i > 0 Then stepsJson.Append(",")
                Dim s = session.Steps(i)
                stepsJson.Append($"{{""description"":""{_escapeJs(s.Description)}"",""detail"":""{_escapeJs(s.Detail)}"",""status"":""pending""}}")
            Next
            stepsJson.Append("]")

            Dim planJson = $"{{""sessionId"":""{session.Id}"",""understanding"":""{_escapeJs(session.Understanding)}"",""steps"":{stepsJson.ToString()},""summary"":""{_escapeJs(session.Summary)}"",""replaceThinkingUuid"":""{AgentThinkingUuid}""}}"

            _executeScript($"showAgentPlanCard({planJson})")

            AgentThinkingUuid = Nothing

        Catch ex As Exception
            Debug.WriteLine($"ShowAgentPlanCard 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理开始执行 Agent
    ''' </summary>
    Public Async Function HandleStartAgentExecution() As Task
        Try
            Debug.WriteLine("[RalphAgent] 用户确认执行")

            If _ralphAgentController IsNot Nothing Then
                Await _ralphAgentController.StartExecution()
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleStartAgentExecution 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"执行失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 处理终止 Agent
    ''' </summary>
    Public Sub HandleAbortAgent()
        Try
            Debug.WriteLine("[RalphAgent] 用户终止Agent")

            If _ralphAgentController IsNot Nothing Then
                _ralphAgentController.AbortAgent()
            End If

            AgentThinkingUuid = Nothing

            GlobalStatusStrip.ShowInfo("已终止Agent")

        Catch ex As Exception
            Debug.WriteLine($"HandleAbortAgent 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理修改 Agent 计划
    ''' </summary>
    Public Sub HandleRefineAgentPlan(jsonDoc As JObject)
        Dim sessionId = jsonDoc("sessionId")?.ToString()
        Dim feedback = jsonDoc("feedback")?.ToString()
        If String.IsNullOrEmpty(feedback) OrElse _ralphAgentController Is Nothing Then Return

        Debug.WriteLine($"[RalphAgentSvc] 用户请求修改计划，sessionId={sessionId}, 反馈={feedback}")
        _executeScript("addThinkingMessage('正在根据您的意见重新规划...')")

        Dim capturedFeedback = feedback
        Task.Run(Async Function()
            Await _ralphAgentController.RefinePlanAsync(capturedFeedback)
        End Function)
    End Sub

End Class
