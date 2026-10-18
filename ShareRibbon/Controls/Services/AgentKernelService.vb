Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' AgentKernel 服务：统一智能体服务，替代 RalphAgentService + RalphLoopService
''' 封装 AgentKernel 的初始化和事件处理，提供与 BaseChatControl 兼容的接口
''' </summary>
Public Partial Class AgentKernelService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _sendAiRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
    Private ReadOnly _executeCodeWithToolResult As Func(Of String, String, Boolean, Agent.ToolResult)
    Private ReadOnly _chatStateService As ChatStateService
    Private ReadOnly _historyMessages As List(Of HistoryMessage)
    Private ReadOnly _manageHistorySize As Action
    Private ReadOnly _getOfficeAppType As Func(Of String)

    ' 统一的 AgentKernel 实例
    Private _agentKernel As Agent.AgentKernel
    Private _officeHarness As Agent.Harness.IOfficeHarness

    ' Agent 状态字段（供 BaseChatControl 访问）
    Public Property AgentThinkingUuid As String = Nothing
    Public Property AgentThinkingLabel As String = Nothing
    Public Property AgentOriginalUserRequest As String = Nothing
    Public Property AgentFullUserMessage As String = Nothing
    Public Property CurrentAgentSessionId As String = Nothing
    Public Property CurrentHarnessRunId As String = Nothing

    ' 协作式取消：每次运行一个 CTS，AbortAgent 触发 Cancel，Loop 在步骤边界收敛。
    Private _runCts As Threading.CancellationTokenSource

    Public Sub New(
        executeScript As Func(Of String, Task),
        sendAiRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String)),
        executeCodeWithToolResult As Func(Of String, String, Boolean, Agent.ToolResult),
        chatStateService As ChatStateService,
        historyMessages As List(Of HistoryMessage),
        manageHistorySize As Action,
        getOfficeAppType As Func(Of String))

        _executeScript = executeScript
        _sendAiRequest = sendAiRequest
        If executeCodeWithToolResult Is Nothing Then Throw New ArgumentNullException(NameOf(executeCodeWithToolResult))
        _executeCodeWithToolResult = executeCodeWithToolResult
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
        _agentKernel.SendAIRequestWithMessages = AddressOf SendAiRequestWithMessagesAsync

        ' 绑定代码执行委托：Agent 主路径强制 ToolResult 回执。
        _agentKernel.ExecuteCodeWithToolResult = Function(code, lang, preview)
                                                     Return _executeCodeWithToolResult(code, lang, preview)
                                                 End Function

        ' 绑定事件
        AddHandler _agentKernel.OnStatusChanged, AddressOf OnKernelStatusChanged
        AddHandler _agentKernel.OnIterationUpdate, AddressOf OnKernelIterationUpdate
        AddHandler _agentKernel.OnStepCompleted, AddressOf OnKernelStepCompleted
        AddHandler _agentKernel.OnExecutionExplained, AddressOf OnKernelExecutionExplained
        AddHandler _agentKernel.OnRequestApproval, AddressOf OnKernelRequestApproval
        AddHandler _agentKernel.OnPlanGenerated, AddressOf OnKernelPlanGenerated
        AddHandler _agentKernel.OnCompleted, AddressOf OnKernelCompleted

        ' 加载工具和技能
        _agentKernel.Initialize()
        _officeHarness = New Agent.Harness.OfficeHarness(_agentKernel, New Agent.Harness.SqliteRunTraceStore())
        AddHandler _officeHarness.ContextReady, AddressOf OnHarnessContextReady
    End Sub

#Region "Public Methods"

    ''' <summary>
    ''' 启动统一 Agent 任务（替代 StartAgent 和 StartLoop）
    ''' </summary>
    Public Async Function StartAgentAsync(userRequest As String, appType As String, currentContent As String,
                                           historyMessages As List(Of Tuple(Of String, String)),
                                           Optional officeContext As Agent.Context.OfficeContext = Nothing,
                                           Optional taskSpec As Agent.AgentTaskSpec = Nothing,
                                           Optional selectedSkills As List(Of SkillFileDefinition) = Nothing) As Task(Of Boolean)
        Try
            EnsureInitialized()

            ' 本轮可能被 harness 拒绝（上一轮仍在途），先留底，拒绝时原样交还，
            ' 避免把在途 run 的会话标识、思考卡和取消令牌清掉。
            Dim carriedSessionId = CurrentAgentSessionId
            Dim carriedThinkingUuid = AgentThinkingUuid
            Dim carriedOriginalRequest = AgentOriginalUserRequest
            Dim carriedFullMessage = AgentFullUserMessage

            ' 保存原始请求：调用方已登记用户可见原文时不得覆盖，否则重放与历史会带上注入后的膨胀上下文
            If String.IsNullOrWhiteSpace(AgentOriginalUserRequest) Then AgentOriginalUserRequest = userRequest
            If String.IsNullOrWhiteSpace(AgentFullUserMessage) Then AgentFullUserMessage = userRequest
            CurrentAgentSessionId = Guid.NewGuid().ToString()

            ' Agent 回合的 user 事件在此单点落库，任务失败或被取消时也留有用户原文
            If Not String.IsNullOrWhiteSpace(AgentOriginalUserRequest) Then
                _chatStateService?.AddMessage("user", AgentOriginalUserRequest)
            End If

            ' 显示思考状态
            ShowThinkingStatus()

            ' 注入历史消息到 AgentMemory（预加载会话上下文）
            If historyMessages IsNot Nothing AndAlso historyMessages.Count > 0 Then
                For Each msg In historyMessages
                    _agentKernel.AddHistoryMessage(msg.Item1, msg.Item2)
                Next
            End If

            ' 执行 Agent 任务（H0: 通过 Harness adapter 收口，内部仍复用 AgentKernel）
            Dim turn As New Agent.Harness.UserTurn With {
                .SessionId = CurrentAgentSessionId,
                .AppType = appType,
                .Text = userRequest,
                .Mode = "agent",
                .HostContextText = currentContent,
                .OfficeContext = officeContext,
                .TaskSpec = taskSpec,
                .SelectedSkills = If(selectedSkills, New List(Of SkillFileDefinition)())
            }
            ' 本轮可能因上一轮仍在途而被拒绝：先留住旧令牌，被拒绝时原样交还，
            ' 否则上一轮任务就再也无法通过"停止"取消。
            Dim carriedCts = _runCts
            Dim runCts As New Threading.CancellationTokenSource()
            _runCts = runCts
            Dim result As Agent.Harness.HarnessRunResult
            ' 回合作用域只服务于这一轮真人在等的模型请求；"停止"要掐断在途请求，而不是等步骤边界。
            Using Agent.AgentCancellationScope.Enter(runCts.Token)
                result = Await _officeHarness.RunAsync(turn, runCts.Token)
            End Using

            If result Is Nothing Then
                AppLogger.Warn("AgentKernelService", "StartAgentAsync returned null result")
                Return RefuseUnstartedRun(carriedSessionId, carriedThinkingUuid, carriedOriginalRequest,
                                          carriedFullMessage, carriedCts, runCts, "Agent 启动失败，请重试")
            End If
            If result.Status = Agent.Harness.HarnessRunStatus.AwaitingApproval Then
                ' 只有真正挂起等待裁决的 run 才登记审批句柄；未启动的 run 不能覆盖它。
                CurrentHarnessRunId = result.RunId
                AppLogger.Info("AgentKernelService", $"Harness run awaiting approval: {result.RunId}")
                Return True
            End If
            If result.Status = Agent.Harness.HarnessRunStatus.Cancelled Then
                AppLogger.Info("AgentKernelService", $"Harness run cancelled: {result.UserMessage}")
                CurrentHarnessRunId = Nothing
                Return False
            End If
            If result.DebugMessage = "RUN_ALREADY_ACTIVE" Then
                ' 上一轮仍在途：本轮从未启动，交还它的会话与取消令牌，只提示不接管。
                AppLogger.Info("AgentKernelService", $"Turn refused, previous run still active: {result.UserMessage}")
                Return RefuseUnstartedRun(carriedSessionId, carriedThinkingUuid, carriedOriginalRequest,
                                          carriedFullMessage, carriedCts, runCts, result.UserMessage)
            End If
            If result.Status <> Agent.Harness.HarnessRunStatus.Succeeded Then
                AppLogger.Warn("AgentKernelService", $"StartAgentAsync agent failed: {result.UserMessage}")
                If Not String.IsNullOrWhiteSpace(CurrentAgentSessionId) Then
                    FinalizeAgentUi(False, result.UserMessage)
                End If
            End If
            CurrentHarnessRunId = Nothing
            Return result.Status = Agent.Harness.HarnessRunStatus.Succeeded
        Catch ex As Exception
            AppLogger.Error("AgentKernelService", "StartAgentAsync exception", ex)
            Dim userMessage = ExceptionClassifier.ToUserMessage(ex, "Agent 启动失败，请重试")
            GlobalStatusStrip.ShowWarning(userMessage)
            FinalizeAgentUi(False, userMessage)
            Return False
        End Try
    End Function

#End Region

#Region "Event Handlers"

    Private Sub OnHarnessContextReady(sender As Object, e As Agent.Harness.HarnessContextEventArgs)
        If e Is Nothing Then Return
        ' runId 需在运行期间可得，供 AbortAgent 在等待审批时定位挂起的运行。
        If Not String.IsNullOrWhiteSpace(e.RunId) Then CurrentHarnessRunId = e.RunId
        If String.IsNullOrWhiteSpace(e.ContextPackJson) Then Return
        Try
            Dim literal = UtilsService.ToJsStringLiteral(e.ContextPackJson)
            _executeScript($"window.officeAiContextPack = JSON.parse({literal}); if (typeof updateContextPackTrace === 'function') updateContextPackTrace(window.officeAiContextPack);")
        Catch ex As Exception
            AppLogger.Warn("AgentKernelService", $"ContextPack UI trace failed: {AppLogger.Redact(ex.Message)}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理状态变更事件
    ''' </summary>
    Private Sub OnKernelStatusChanged(status As String)
        Try
            GlobalStatusStrip.ShowInfo(status)

            ' 更新思考状态 div
            If Not String.IsNullOrEmpty(AgentThinkingUuid) Then
                _executeScript($"setThinkingStatusText({UtilsService.ToJsStringLiteral(AgentThinkingUuid)}, {UtilsService.ToJsStringLiteral(status)})")
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

            Dim iterationJson = UtilsService.SerializeForScript(JObject.FromObject(New With {
                .index = iteration.Index,
                .thought = If(iteration.Thought, ""),
                .action = If(iteration.Action?.ToolId, ""),
                .observation = If(iteration.Observation, ""),
                .explanation = iteration.Explanation
            }))
            _executeScript($"updateAgentIteration({UtilsService.ToJsStringLiteral(CurrentAgentSessionId)}, {iterationJson})")
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
            _executeScript($"updateAgentStep({UtilsService.ToJsStringLiteral(CurrentAgentSessionId)}, {stepIndex}, {UtilsService.ToJsStringLiteral(stepStatus)}, {UtilsService.ToJsStringLiteral(message)})")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnStepCompleted 出错: {ex.Message}")
        End Try
    End Sub

    Private Sub OnKernelExecutionExplained(explanation As Agent.ExecutionExplanation)
        Try
            If explanation Is Nothing OrElse String.IsNullOrWhiteSpace(CurrentAgentSessionId) Then Return
            Dim explanationJson = UtilsService.SerializeForScript(JObject.FromObject(explanation))
            _executeScript($"showAgentExecutionExplanation({UtilsService.ToJsStringLiteral(CurrentAgentSessionId)}, {explanationJson})")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnExecutionExplained 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理审批请求事件
    ''' </summary>
    Private Sub OnKernelRequestApproval(message As String, callback As Action(Of Boolean))
        Try
            ' 显示审批 UI
            _executeScript($"showAgentApproval({UtilsService.ToJsStringLiteral(CurrentAgentSessionId)}, {UtilsService.ToJsStringLiteral(message)})")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnRequestApproval 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理计划生成事件
    ''' </summary>
    Private Sub OnKernelPlanGenerated(plan As Agent.ExecutionPlan)
        Try
            If plan Is Nothing Then Return

            ' 构建步骤 JSON：一律交给序列化器，避免把模型文本手拼进 JS 源码
            Dim steps As New JArray()
            For Each s In plan.Steps
                steps.Add(New JObject From {
                    {"description", If(s.Description, "")},
                    {"code", If(s.Code, "")},
                    {"language", If(s.Language, "")},
                    {"status", "pending"}
                })
            Next

            Dim planJson = UtilsService.SerializeForScript(New JObject From {
                {"sessionId", If(CurrentAgentSessionId, "")},
                {"understanding", If(plan.Understanding, "")},
                {"steps", steps},
                {"summary", If(plan.Summary, "")},
                {"replaceThinkingUuid", If(AgentThinkingUuid, "")}
            })

            _executeScript($"showAgentPlanCard({planJson})")
            Dim planTrace As New ChatContextTrace With {
                .ExecutionPlan = New ChatContextPlanTrace With {
                    .Summary = If(plan.Summary, ""),
                    .Understanding = If(plan.Understanding, "")
                }
            }
            For Each stepItem In plan.Steps
                planTrace.ExecutionPlan.Steps.Add(New ChatContextPlanStepTrace With {
                    .StepNumber = stepItem.StepNumber,
                    .Description = If(stepItem.Description, ""),
                    .ToolOrCode = If(stepItem.Code, ""),
                    .Language = If(stepItem.Language, "")
                })
            Next
            Dim traceJson = UtilsService.SerializeForScript(JObject.FromObject(planTrace))
            _executeScript($"showContextHints({{ trace: {traceJson} }})")
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
        Dim terminalSuccess = result IsNot Nothing AndAlso result.Success
        Dim terminalMessage = If(result?.Message, If(terminalSuccess, "任务完成", "任务失败"))
        Try
            Dim userMsgForHistory = If(Not String.IsNullOrWhiteSpace(AgentFullUserMessage), AgentFullUserMessage, AgentOriginalUserRequest)

            If Not String.IsNullOrWhiteSpace(userMsgForHistory) Then
                ' 只补充发给模型的完整上下文用于后续追问；会话历史的 user 事件已在 StartAgentAsync 落库
                _historyMessages.Add(New HistoryMessage With {
                    .role = "user",
                    .content = userMsgForHistory
                })
                _manageHistorySize()
            End If

            Dim assistantReply = terminalMessage
            _historyMessages.Add(New HistoryMessage With {
                .role = "assistant",
                .content = assistantReply
            })
            _manageHistorySize()
            _chatStateService?.AddMessage("assistant", assistantReply)

            MemoryService.SaveConversationTurnAsync(userMsgForHistory, assistantReply, _chatStateService?.CurrentSessionId, _getOfficeAppType())

        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnCompleted 出错: {ex.Message}")
        Finally
            FinalizeAgentUi(terminalSuccess, terminalMessage)
        End Try
    End Sub

#End Region

#Region "Private Helpers"

    ''' <summary>
    ''' Shared non-streaming gateway used only when the Loop has ephemeral multimodal
    ''' observation evidence. Request bodies are not logged or added to chat history.
    ''' </summary>
    Private Async Function SendAiRequestWithMessagesAsync(messages As JArray) As Task(Of String)
        If messages Is Nothing OrElse messages.Count = 0 Then Return Nothing

        Dim response = Await AiGateway.SendChatAsync(New AiRequestOptions With {
            .ApiUrl = ConfigSettings.ApiUrl,
            .ApiKey = ConfigSettings.ApiKey,
            .ModelName = ConfigSettings.ModelName,
            .Platform = ConfigSettings.platform,
            .ReasoningMode = ConfigSettings.ReasoningMode,
            .Messages = messages,
            .TimeoutSeconds = 90,
            .Cancellation = Agent.AgentCancellationScope.Current
        })
        If response Is Nothing OrElse Not response.Success Then
            Throw New InvalidOperationException(If(response?.ErrorMessage, "Multimodal AI request failed"))
        End If
        Return response.Content
    End Function

    Private Sub FinalizeAgentUi(success As Boolean, message As String)
        Dim sessionId = If(CurrentAgentSessionId, "")
        Dim thinkingUuid = If(AgentThinkingUuid, "")
        Try
            _executeScript($"completeAgent({UtilsService.ToJsStringLiteral(sessionId)}, {success.ToString().ToLowerInvariant()}, {UtilsService.ToJsStringLiteral(message)}, {UtilsService.ToJsStringLiteral(thinkingUuid)})")
        Catch ex As Exception
            AppLogger.Warn("AgentKernelService", $"FinalizeAgentUi failed: {AppLogger.Redact(ex.Message)}")
            ' 即使主终态函数不存在，也直接执行最小按钮/规划卡复位。
            _executeScript("var p=document.getElementById('planning-status-card');if(p)p.remove();if(typeof restoreAgentRequestUi==='function')restoreAgentRequestUi();else if(typeof changeSendButton==='function')changeSendButton();")
        Finally
            AgentOriginalUserRequest = Nothing
            AgentFullUserMessage = Nothing
            AgentThinkingUuid = Nothing
            CurrentAgentSessionId = Nothing
            CurrentHarnessRunId = Nothing
            Dim cts = _runCts
            _runCts = Nothing
            If cts IsNot Nothing Then
                Try
                    cts.Dispose()
                Catch
                End Try
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 在聊天界面显示思考状态
    ''' </summary>
    Private Sub ShowThinkingStatus()
        Try
            If String.IsNullOrEmpty(AgentThinkingUuid) Then
                AgentThinkingUuid = Guid.NewGuid().ToString()
            End If

            Dim timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            _executeScript($"createChatSection({UtilsService.ToJsStringLiteral("AI")}, {UtilsService.ToJsStringLiteral(timestamp)}, {UtilsService.ToJsStringLiteral(AgentThinkingUuid)})")
            _executeScript($"showThinkingIndicator({UtilsService.ToJsStringLiteral(AgentThinkingUuid)}, {UtilsService.ToJsStringLiteral(AgentThinkingLabel)})")
            AgentThinkingLabel = Nothing
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] ShowThinkingStatus 出错: {ex.Message}")
        End Try
    End Sub

#End Region

End Class
