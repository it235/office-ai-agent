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
    ''' <summary>
    ''' 本轮的入口动作 id（E6-05）：任务包断言唯一的绑定依据。发起方每条路径都必须显式赋值，
    ''' 自由文本置为空串——残留上一轮的 id 会把某个包的断言套到不相干的任务上，比没有断言更糟。
    ''' </summary>
    Public Property AgentEntryActionId As String = Nothing
    Public Property AgentFullUserMessage As String = Nothing
    Public Property CurrentAgentSessionId As String = Nothing
    Public Property CurrentHarnessRunId As String = Nothing

    ' 协作式取消：每次运行一个 CTS，AbortAgent 触发 Cancel，Loop 在步骤边界收敛。
    Private _runCts As Threading.CancellationTokenSource

    ' 终态变更摘要最多列出的条目数：再多用户也读不完，超出部分由 changeTruncated 说明。
    Private Const MaxTerminalChanges As Integer = 8

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
        AddHandler _agentKernel.OnStepStatusChanged, AddressOf OnKernelStepStatusChanged
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
                                           Optional selectedSkills As List(Of SkillFileDefinition) = Nothing,
                                           Optional entryActionId As String = Nothing) As Task(Of Boolean)
        Try
            EnsureInitialized()

            ' 本轮可能被 harness 拒绝（上一轮仍在途），先留底，拒绝时原样交还，
            ' 避免把在途 run 的会话标识、思考卡和取消令牌清掉。
            Dim carriedSessionId = CurrentAgentSessionId
            Dim carriedThinkingUuid = AgentThinkingUuid
            Dim carriedOriginalRequest = AgentOriginalUserRequest
            Dim carriedFullMessage = AgentFullUserMessage
            Dim carriedEntryActionId = AgentEntryActionId

            ' 入口身份按调用参数登记，不读调用方"记得写进去"的共享值：
            ' 面板发起与后台真正开跑的之间有线程间隙，靠属性传会导致后一轮抢走前一轮的包。
            If entryActionId Is Nothing Then
                ' 没登记身份=本轮没有任何包断言会被核对。漏接一个入口不会报错，只会表现为"验证器没生效"，
                ' 所以这里必须留痕；自由文本要显式传空串，表示"确实没有包"，不是"忘了传"。
                AppLogger.Warn("AgentKernelService", "StartAgentAsync called without entry identity; task package assertions are skipped")
            End If
            AgentEntryActionId = If(entryActionId, "")

            ' 保存原始请求：调用方已登记用户可见原文时不得覆盖，否则重放与历史会带上注入后的膨胀上下文
            If String.IsNullOrWhiteSpace(AgentOriginalUserRequest) Then AgentOriginalUserRequest = userRequest
            If String.IsNullOrWhiteSpace(AgentFullUserMessage) Then AgentFullUserMessage = userRequest
            CurrentAgentSessionId = Guid.NewGuid().ToString()

            ' Agent 回合的 user 事件在此单点落库，任务失败或被取消时也留有用户原文
            If Not String.IsNullOrWhiteSpace(AgentOriginalUserRequest) Then
                _chatStateService?.AddMessage("user", AgentOriginalUserRequest)
            End If

            ' E3-04：本轮的用量窗口在开跑时复位，终态卡片报的是"这一轮花了多少"，
            ' 会话累计仍记在同一本账上，不另开第二本。
            _chatStateService?.BeginRunUsage()

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
                .SelectedSkills = If(selectedSkills, New List(Of SkillFileDefinition)()),
                .EntryActionId = AgentEntryActionId
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
                                          carriedFullMessage, carriedEntryActionId,
                                          carriedCts, runCts, "Agent 启动失败，请重试")
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
                                          carriedFullMessage, carriedEntryActionId,
                                          carriedCts, runCts, result.UserMessage)
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
            ' OnCompleted 可能已经渲染终态（含撤销入口），此处只在尚未收尾时补一次。
            If Not String.IsNullOrWhiteSpace(CurrentAgentSessionId) Then
                FinalizeAgentUi(False, userMessage)
            End If
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
                .stepIndex = If(iteration.Explanation?.StepIndex, -1),
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
    ''' 处理步骤状态事件；状态令牌由 LoopEngine 定义（pending/running/completed/failed/skipped/aborted），
    ''' 这里原样透传，收成一个布尔会把"用户拒绝"和"执行失败"画成同一个红叉。
    ''' </summary>
    Private Sub OnKernelStepStatusChanged(stepIndex As Integer, status As String, message As String)
        Try
            Dim stepStatus = If(String.IsNullOrWhiteSpace(status), "failed", status)
            _executeScript($"updateAgentStep({UtilsService.ToJsStringLiteral(CurrentAgentSessionId)}, {stepIndex}, {UtilsService.ToJsStringLiteral(stepStatus)}, {UtilsService.ToJsStringLiteral(message)})")
        Catch ex As Exception
            Debug.WriteLine($"[AgentKernelService] OnStepStatusChanged 出错: {ex.Message}")
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
    Private Sub OnKernelRequestApproval(request As Agent.ApprovalRequest, callback As Action(Of Agent.ApprovalDecision))
        Try
            ' 显示审批 UI：带上工具与风险，用户可以判断是否只放行一次或放到本会话
            Dim toolId = If(request?.ToolId, "")
            Dim riskLevel = If(request?.RiskLevel, "")
            Dim message = If(request?.Message, "该操作需要用户确认")
            _executeScript($"showAgentApproval({UtilsService.ToJsStringLiteral(CurrentAgentSessionId)}, " &
                           $"{UtilsService.ToJsStringLiteral(toolId)}, " &
                           $"{UtilsService.ToJsStringLiteral(message)}, " &
                           $"{UtilsService.ToJsStringLiteral(riskLevel)})")
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

            Dim riskyTools As New JArray()
            If plan.Trust IsNot Nothing Then
                For Each risky In plan.Trust.RiskyTools
                    riskyTools.Add(risky)
                Next
            End If

            Dim planJson = UtilsService.SerializeForScript(New JObject From {
                {"sessionId", If(CurrentAgentSessionId, "")},
                {"understanding", If(plan.Understanding, "")},
                {"planVersion", If(plan.Trust?.PlanVersion, 1)},
                {"scope", If(plan.Trust?.ScopeText, "")},
                {"stepCount", plan.Steps.Count},
                {"toolCount", If(plan.Trust?.ToolCount, 0)},
                {"riskLevel", If(plan.Trust?.RiskLevel, "safe")},
                {"riskyTools", riskyTools},
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
            PersistChangeLog(result)
            FinalizeAgentUi(terminalSuccess, terminalMessage, BuildTerminalDescriptor(result))
        End Try
    End Sub

#End Region

#Region "Private Helpers"

    ''' <summary>
    ''' 在渲染终态前先落库本轮变更，保证“撤销本轮”按钮出现时数据库已有可撤销依据。
    ''' Harness 随后仍会按同一 change_id 覆盖写入，内容一致，天然幂等。
    ''' </summary>
    Private Sub PersistChangeLog(result As Agent.AgentResult)
        If result Is Nothing OrElse result.ChangeLog Is Nothing Then Return
        Dim runId = CurrentHarnessRunId
        If String.IsNullOrWhiteSpace(runId) Then Return
        Try
            AgentChangeLogRepository.SaveRun(runId,
                                             If(String.IsNullOrWhiteSpace(result.SessionId), CurrentAgentSessionId, result.SessionId),
                                             _getOfficeAppType(),
                                             result.ChangeLog)
        Catch ex As Exception
            AppLogger.Warn("AgentKernelService", $"Persist change log failed: {AppLogger.Redact(ex.Message)}")
        End Try
    End Sub

    ''' <summary>
    ''' 终态卡片要回答两个问题：本轮在文档里实际留下了什么，以及能不能一键撤销、为什么不能。
    ''' 变更条目只取宿主观察里的短摘要并限长，不带文档原文、参数或截图。
    ''' 撤销按钮只看残留改动是否全部可还原——Word 走宿主原生撤销单元，Excel/PPT 走可重放的还原快照；
    ''' 不出按钮时必须给出原因，否则用户会以为这些改动随时能回退。
    ''' </summary>
    Private Function BuildTerminalDescriptor(result As Agent.AgentResult) As JObject
        Dim changeLog = result?.ChangeLog
        If changeLog Is Nothing Then Return Nothing

        Dim changes As New JArray()
        Dim rolledBackCount As Integer = 0
        For Each entry In changeLog.Entries
            If entry.RolledBack Then rolledBackCount += 1
            If Not entry.Changed OrElse entry.RolledBack Then Continue For
            If changes.Count < MaxTerminalChanges Then
                changes.Add(New JObject From {
                    {"seq", entry.Seq},
                    {"toolId", entry.ToolId},
                    {"summary", BoundText(If(entry.Summary, entry.StepDescription), 120)},
                    {"partialApply", entry.PartialApply}
                })
            End If
        Next

        Dim descriptor As New JObject From {
            {"canUndo", False},
            {"runId", ""},
            {"units", 0},
            {"restoreCount", 0},
            {"writtenCount", changeLog.WrittenCount},
            {"rolledBackCount", rolledBackCount},
            {"changeTruncated", changeLog.WrittenCount > MaxTerminalChanges},
            {"changes", changes},
            {"undoReason", ""},
            {"usage", BuildUsageBlock()}
        }

        If changeLog.WrittenCount = 0 Then
            descriptor("undoReason") = "本轮没有留在文档里的改动，不需要撤销。"
            Return descriptor
        End If
        If Not changeLog.FullyUndoable Then
            descriptor("undoReason") = "有改动缺少可靠的撤销依据（既没有宿主撤销单元，也没有可重放的还原快照），一键回退可能只撤掉一半，因此不提供撤销按钮。"
            Return descriptor
        End If

        Dim runId = If(CurrentHarnessRunId, "")
        If String.IsNullOrWhiteSpace(runId) Then
            descriptor("undoReason") = "本轮没有运行标识，无法定位需要回退的改动记录。"
            Return descriptor
        End If

        descriptor("canUndo") = True
        descriptor("runId") = runId
        descriptor("units") = changeLog.UndoableUnits
        descriptor("restoreCount") = changeLog.RestoreCount
        descriptor("undoReason") = ""
        Return descriptor
    End Function

    ''' <summary>
    ''' 本轮用量：卡片报「这一轮花了多少」，费用按设置里的单价估算。
    ''' 一句话文案由 UsageCostReporter 持有，这里只搬数字，JS 不再自己拼金额。
    ''' 没有单价时只有 tokens，没有 tokens 时只有「没有用量数据」——三者都不许写成 0 元。
    ''' </summary>
    Private Function BuildUsageBlock() As JObject
        Dim run = If(_chatStateService?.RunUsage, New TokenUsage())
        Dim sessionTokens = If(_chatStateService Is Nothing, 0, _chatStateService.CurrentSessionTotalTokens)
        Dim summary = UsageCostReporter.SummaryText(run, sessionTokens)
        Return New JObject From {
            {"promptTokens", run.PromptTokens},
            {"completionTokens", run.CompletionTokens},
            {"totalTokens", run.TotalTokens},
            {"sessionTokens", sessionTokens},
            {"summary", If(String.IsNullOrEmpty(summary), UsageCostReporter.NoUsageText, summary)}
        }
    End Function

    ''' <summary>摘要来自宿主观察，正常已经很短；这里只是兜底，防止个别宿主把正文塞进 summary。</summary>
    Private Shared Function BoundText(value As String, maxLength As Integer) As String
        Dim text = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        If text.Length <= maxLength Then Return text
        Return text.Substring(0, maxLength) & "…"
    End Function

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
        ' 视觉 Repair 走的也是这一条请求，用量同样要记进本轮账本，否则卡片上的数字比实际花的少。
        If response.Usage IsNot Nothing Then _chatStateService?.RecordUsage(response.Usage)
        Return response.Content
    End Function

    Private Sub FinalizeAgentUi(success As Boolean, message As String, Optional terminal As JObject = Nothing)
        Dim sessionId = If(CurrentAgentSessionId, "")
        Dim thinkingUuid = If(AgentThinkingUuid, "")
        Dim terminalJson = If(terminal Is Nothing, "null", UtilsService.SerializeForScript(terminal))
        Try
            _executeScript($"completeAgent({UtilsService.ToJsStringLiteral(sessionId)}, {success.ToString().ToLowerInvariant()}, {UtilsService.ToJsStringLiteral(message)}, {UtilsService.ToJsStringLiteral(thinkingUuid)}, {terminalJson})")
        Catch ex As Exception
            AppLogger.Warn("AgentKernelService", $"FinalizeAgentUi failed: {AppLogger.Redact(ex.Message)}")
            ' 即使主终态函数不存在，也直接执行最小输入锁/按钮复位（规划卡由 agent-card.js 自己管）。
            _executeScript("if(typeof restoreAgentRequestUi==='function')restoreAgentRequestUi();else if(typeof changeSendButton==='function')changeSendButton();")
        Finally
            AgentOriginalUserRequest = Nothing
            AgentFullUserMessage = Nothing
            AgentEntryActionId = Nothing
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
