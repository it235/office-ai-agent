Imports System.Text
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' ReAct 循环引擎 - 核心执行逻辑
    ''' Think -> Plan -> Act -> Observe -> Reflect
    ''' </summary>
    Public Partial Class LoopEngine
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _promptManager As PromptManager
        Private ReadOnly _undoManager As Core.UndoManager
        Private _multimodalRepairDisabledForRun As Boolean

        ' 循环限制
        Private Const MaxIterations As Integer = 15
        Private Const MaxNoProgress As Integer = 3
        Private Const MaxReplanAttempts As Integer = 2

        ' 步骤状态令牌：UI 需要区分失败、用户拒绝（跳过）、审批放弃（终止）和正在执行，
        ' 只给布尔会把跳过和终止都画成红叉，用户看不出哪些是自己拒绝的。
        Public Const StepStatusPending As String = "pending"
        Public Const StepStatusRunning As String = "running"
        Public Const StepStatusCompleted As String = "completed"
        Public Const StepStatusFailed As String = "failed"
        Public Const StepStatusSkipped As String = "skipped"
        Public Const StepStatusAborted As String = "aborted"

        ' 回调
        Public Property OnStatusChanged As Action(Of String)
        Public Property OnIterationUpdate As Action(Of ReActIteration)
        Public Property OnStepStatusChanged As Action(Of Integer, String, String)
        Public Property OnExecutionExplained As Action(Of ExecutionExplanation)
        Public Property OnRequestApproval As Func(Of ApprovalRequest, Task(Of ApprovalDecision))
        Public Property OnPlanGenerated As Action(Of ExecutionPlan)
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))
        Public Property SendAIRequestWithMessages As Func(Of JArray, Task(Of String))

        Public Sub New(toolRegistry As ToolRegistry, memory As AgentMemory, promptManager As PromptManager)
            _toolRegistry = toolRegistry
            _memory = memory
            _promptManager = promptManager
            _undoManager = New Core.UndoManager()
        End Sub

        ''' <summary>
        ''' 执行 ReAct 循环
        ''' </summary>
        Public Async Function RunAsync(session As AgentSession,
                                        systemPrompt As String,
                                        Optional skill As AgentSkill = Nothing,
                                        Optional cancellationToken As CancellationToken = Nothing) As Task(Of AgentResult)
            Dim noProgressCount As Integer = 0
            Dim replanAttempts As Integer = 0
            _multimodalRepairDisabledForRun = False

            Try
                ' Phase 1: 生成 Spec
                OnStatusChanged?.Invoke("正在分析任务...")
                If session.Spec Is Nothing Then
                    session.Spec = Await GenerateSpecAsync(session)
                Else
                    AppLogger.Info("LoopEngine", $"Use precomputed TaskSpec goal={AppLogger.Redact(session.Spec.Goal)}")
                End If
                ' 分析/规划各是一次模型请求：用户已停止就到此收敛，不再把预算花在还没有执行者的计划上。
                If cancellationToken.IsCancellationRequested Then Return FinishCancelled(session)

                ' Phase 2: 生成计划
                OnStatusChanged?.Invoke("正在制定执行计划...")
                session.Plan = Await GeneratePlanAsync(session, systemPrompt, skill)
                If session.Plan Is Nothing OrElse session.Plan.Steps.Count = 0 Then
                    Dim capabilityGap = If(session.Plan?.CapabilityGap, "")
                    Dim planFailure = If(String.IsNullOrWhiteSpace(capabilityGap),
                                         "规划失败：模型未生成可执行计划",
                                         "当前能力无法完整执行：" & capabilityGap)
                    OnStatusChanged?.Invoke(planFailure)
                    Return AgentResult.Failed(session.Id, planFailure)
                End If

                Dim coverageError = ValidatePlanCoverage(session.Spec, session.Plan)
                If Not String.IsNullOrWhiteSpace(coverageError) Then
                    OnStatusChanged?.Invoke(coverageError)
                    Return AgentResult.Failed(session.Id, coverageError)
                End If

                ' 通知计划已生成
                RaisePlanGenerated(session)

                session.Status = AgentStatus.Executing
                OnStatusChanged?.Invoke($"规划完成（共 {session.Plan.Steps.Count} 步），进入自治执行...")
                Dim executionContext As ToolExecutionContext = ToolExecutionContext.FromSession(session, skill)

                ' Phase 3: ReAct Loop
                Dim stepIndex As Integer = 0
                ' 计划跑完时让循环再进一圈，那一圈只走 run 末合同闸门；每接受一次修复重规划就重新武装它。
                ' 第三个条件同样是为了"再进一圈"：用户对某一步按了重试或跳过，而我们已经跑到它后面了。
                Dim runEndGated As Boolean = False
                While (stepIndex < session.Plan.Steps.Count OrElse Not runEndGated OrElse
                       session.Interventions.PendingCount > 0) AndAlso
                      session.CurrentIteration < MaxIterations
                    ' 协作式停止：在步骤边界检查，不打断进行中的宿主操作，但不启动下一步。
                    If cancellationToken.IsCancellationRequested Then
                        Return FinishCancelled(session)
                    End If

                    ' --- USER STEP RECOVERY (E10-01) ---
                    ' 用户对失败步骤的纠偏只在这里落成步骤状态：面板或宿主自己去改 PlanStep 会造出
                    ' 第二套执行真相，而只有循环知道当前步号、能安全回退到更早的那一步。
                    If session.Interventions.PendingCount > 0 Then
                        ApplyUserInterventions(session, stepIndex, noProgressCount)
                    End If

                    ' 已经被判定过、又没人要求重跑的步骤直接前进，这里一条状态都不发：每个状态都在被写下的
                    ' 那一刻发过一次，再发一次只是让卡片重复出现同一句。纠偏回退到一个已完成的位置、或用户
                    ' 跳过的是更靠后的步骤时，循环必须只落在那一步——否则成功过的写入会被再执行一遍。
                    If stepIndex < session.Plan.Steps.Count AndAlso
                       session.Plan.Steps(stepIndex).Status <> StepStatus.Pending Then
                        stepIndex += 1
                        Continue While
                    End If

                    If stepIndex >= session.Plan.Steps.Count Then
                        ' --- RUN-END CONTRACT GATE (E6-05b) ---
                        ' 步骤都落地了才轮到包的成功标准裁决。未达成且包声明 onFail=repair 时，把未达成的
                        ' 断言当作观察结果交给既有 reflect/replan 路由再带一轮——预算与失败修复共用 replanAttempts，
                        ' 终态仍只有循环后面那一条出口；拿不到可用新计划就照旧报未达成，不猜也不装成功。
                        runEndGated = True
                        Dim repairPlan As ExecutionPlan = Nothing
                        If session.CurrentIteration > 0 AndAlso IncompletePlanSteps(session).Count = 0 AndAlso
                           PackageWantsRunEndRepair(session) AndAlso replanAttempts < MaxReplanAttempts Then
                            Dim unmetReasons = UnmetTaskPackageReasons(session)
                            If unmetReasons.Count > 0 Then
                                repairPlan = Await RequestRunEndRepairPlanAsync(session, unmetReasons, systemPrompt)
                            End If
                        End If
                        If repairPlan Is Nothing Then Exit While

                        replanAttempts += 1
                        session.Plan = repairPlan
                        session.Writes.BeginNextPlan()
                        ' 新计划里同号不是同号：旧步号上的用户纠偏必须作废，否则会落到一个语义无关的步骤上。
                        session.Interventions.Clear()
                        stepIndex = 0
                        runEndGated = False
                        RaisePlanGenerated(session)
                        OnStatusChanged?.Invoke($"成功标准未达成，已按任务包合同重新规划（共 {repairPlan.Steps.Count} 步），已写入 {session.ChangeLog.WrittenCount} 处不会被重复执行...")
                        Continue While
                    End If

                    Dim planStep = session.Plan.Steps(stepIndex)
                    planStep.Status = StepStatus.Running
                    ' 不发这条，计划卡里永远没有 .step-running，迭代详情只能靠全局迭代号猜位置
                    OnStepStatusChanged?.Invoke(stepIndex, StepStatusRunning, Nothing)

                    ' --- THINK ---
                    ' 用户改写过这一步的说明/参数时，工具与参数只认卡片上保存的 Code：再问一次模型，
                    ' 等于让模型按自己的理解决定要不要采纳用户的修改，那次编辑就只是一条建议（E10-02）。
                    session.Status = AgentStatus.Thinking
                    OnStatusChanged?.Invoke($"步骤 {stepIndex + 1}/{session.Plan.Steps.Count}: {planStep.Description}")
                    Dim thought As String = Nothing
                    Dim toolCall As ToolCall = Nothing
                    If planStep.PinnedByUser Then
                        thought = "用户已改写这一步的说明与参数，按改后的内容执行"
                        PlanStepCode.TryParse(planStep.Code, toolCall)
                        AppLogger.Info("LoopEngine", $"Step {stepIndex + 1} runs from user-edited definition without a model round")
                    Else
                        thought = Await ThinkAsync(session, planStep, systemPrompt)
                        toolCall = ParseToolCall(thought)
                    End If
                    ' 停止必须拦在 ACT 之前：即便这一步的参数已由用户定稿，也不能越过取消。
                    If cancellationToken.IsCancellationRequested Then Return FinishCancelled(session)

                    ' --- PARSE ACTION ---
                    If toolCall Is Nothing Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = "无法解析工具调用"
                        OnStepStatusChanged?.Invoke(stepIndex, StepStatusFailed, "解析失败")

                        If noProgressCount >= MaxNoProgress Then Exit While
                        stepIndex += 1
                        Continue While
                    End If

                    Dim normalizeMessage As String = ""
                    If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, executionContext, normalizeMessage) Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = normalizeMessage
                        OnStepStatusChanged?.Invoke(stepIndex, StepStatusFailed, normalizeMessage)

                        If noProgressCount >= MaxNoProgress Then Exit While
                        stepIndex += 1
                        Continue While
                    End If
                    If Not String.IsNullOrWhiteSpace(normalizeMessage) Then
                        Debug.WriteLine($"[LoopEngine] {normalizeMessage}")
                    End If

                    ' --- IDEMPOTENCE CHECK (E1-06) ---
                    ' 重新规划出的新计划不知道文档已被改过，会把上一版已写成的步骤重新排进来。
                    ' 跨计划版本命中写入台账时不再执行，否则同一处内容会被写第二遍。
                    Dim priorWrite = session.Writes.FindAppliedInPriorPlan(toolCall)

                    ' --- RISK NOTICE (autonomous mode) ---
                    Dim tool = _toolRegistry.GetTool(toolCall.ToolId)
                    If priorWrite Is Nothing AndAlso tool IsNot Nothing AndAlso tool.RiskLevel = "risky" Then
                        Debug.WriteLine($"[LoopEngine] 自治模式执行高风险工具: {toolCall.ToolId}")
                        OnStatusChanged?.Invoke($"步骤 {stepIndex + 1} 使用高风险工具 {toolCall.ToolId}，已记录风险并继续执行")
                    End If

                    ' --- ACT (增强版 - 多轮自修复 + 撤销点) ---
                    session.Status = AgentStatus.Executing

                    ' 创建撤销点（执行前保存状态）；被判定为重复写入的步骤不碰文档，不必申请撤销点
                    Dim undoPoint As Core.UndoManager.UndoPoint = Nothing
                    If _undoManager IsNot Nothing AndAlso priorWrite Is Nothing Then
                        undoPoint = _undoManager.CreateUndoPoint(
                            If(session.AppType, "Unknown"),
                            $"步骤 {stepIndex + 1}: {toolCall.ToolId}",
                            planStep.Description)
                        If undoPoint IsNot Nothing Then
                            Debug.WriteLine($"[LoopEngine] 创建撤销点: {undoPoint.Name}")
                        End If
                    End If

                    Const MaxFixAttempts As Integer = 3
                    Dim fixAttempt As Integer = 0
                    Dim toolResult As ToolResult = Nothing
                    Dim stepStartedAt = DateTime.Now

                    ' 多轮修复循环
                    While fixAttempt < MaxFixAttempts
                        If priorWrite IsNot Nothing Then
                            toolResult = BuildAlreadyAppliedResult(toolCall, priorWrite)
                            AppLogger.Info("LoopEngine",
                                $"Duplicate write skipped after replan toolId={toolCall.ToolId} appliedByPlan={priorWrite.PlanGeneration}")
                            Exit While
                        End If

                        Dim normalized As String = ""
                        If Not _toolRegistry.TryNormalizeToolCall(session.AppType, toolCall, executionContext, normalized) Then
                            toolResult = ToolResult.Failed(toolCall.ToolId,
                                                           normalized,
                                                           New With {.availableTools = BuildAvailableToolHint(session.AppType, executionContext)},
                                                           ExceptionClassifier.CodeNotFound,
                                                           normalized,
                                                           normalized,
                                                           recoverable:=True)
                        Else
                            If Not String.IsNullOrWhiteSpace(normalized) Then Debug.WriteLine($"[LoopEngine] {normalized}")
                            ' 执行工具
                            toolResult = ValidateObservedOutcome(
                                Await _toolRegistry.ExecuteToolAsync(executionContext, toolCall.ToolId, toolCall.Parameters))
                        End If

                        If Not toolResult.Success AndAlso
                           String.Equals(toolResult.ErrorCode, ExceptionClassifier.CodeSafetyNeedsApproval, StringComparison.OrdinalIgnoreCase) AndAlso
                           OnRequestApproval IsNot Nothing Then
                            Dim approvalRequest As New ApprovalRequest With {
                                .ToolId = If(toolCall.ToolId, ""),
                                .Message = If(toolResult.UserMessage, toolResult.Message),
                                .RiskLevel = If(tool IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(tool.RiskLevel), tool.RiskLevel, "risky")
                            }
                            OnStatusChanged?.Invoke($"工具 {toolCall.ToolId} 正在等待用户确认...")
                            Dim decision = Await OnRequestApproval(approvalRequest)
                            If decision = ApprovalDecision.Session Then
                                executionContext.ApproveToolForSession(toolCall.ToolId)
                                OnStatusChanged?.Invoke($"本会话已允许工具 {toolCall.ToolId}，继续执行...")
                            ElseIf decision = ApprovalDecision.Once Then
                                executionContext.ApproveTool(toolCall.ToolId, toolCall.Parameters)
                                OnStatusChanged?.Invoke($"用户已批准工具 {toolCall.ToolId}，继续执行...")
                            End If

                            If decision = ApprovalDecision.Once OrElse decision = ApprovalDecision.Session Then
                                toolResult = ValidateObservedOutcome(
                                    Await _toolRegistry.ExecuteToolAsync(executionContext, toolCall.ToolId, toolCall.Parameters))
                            ElseIf decision = ApprovalDecision.Deny Then
                                toolResult = ToolResult.Failed(
                                    toolCall.ToolId,
                                    "用户拒绝高风险操作",
                                    errorCode:=ExceptionClassifier.CodeSafetyBlocked,
                                    userMessage:="已跳过这个需要确认的操作",
                                    recoverable:=False,
                                    observation:=New JObject From {
                                        {"kind", "approval"},
                                        {"summary", "用户拒绝高风险操作"},
                                        {"changed", False},
                                        {"decision", "rejected"},
                                        {"warnings", New JArray("approval_rejected")}
                                    })
                            Else
                                toolResult = ToolResult.Failed(
                                    toolCall.ToolId,
                                    "审批已放弃",
                                    errorCode:=ExceptionClassifier.CodeSafetyBlocked,
                                    userMessage:="审批已放弃，本轮停止执行",
                                    recoverable:=False,
                                    observation:=New JObject From {
                                        {"kind", "approval"},
                                        {"summary", "审批已放弃"},
                                        {"changed", False},
                                        {"decision", "aborted"},
                                        {"warnings", New JArray("approval_aborted")}
                                    })
                            End If
                        End If

                        If toolResult.Success Then
                            ' 成功，跳出循环
                            Exit While
                        End If

                        ' 失败：尝试自动修复
                        fixAttempt += 1
                        If fixAttempt < MaxFixAttempts Then
                            ' Non-recoverable tool failures skip AI repair and go straight to observe/reflect.
                            If Not toolResult.Recoverable Then
                                AppLogger.Warn("LoopEngine", $"Skip repair for non-recoverable tool failure: {toolResult.ToObserveSummary()}")
                                Exit While
                            End If

                            OnStatusChanged?.Invoke($"代码执行失败，AI 正在修复（尝试 {fixAttempt}/{MaxFixAttempts}）...")
                            AppLogger.Info("LoopEngine", $"Repair attempt {fixAttempt}/{MaxFixAttempts}: {toolResult.ToObserveSummary()}")

                            ' 构建修复提示词（含结构化错误契约）
                            Dim fixPrompt = $"上一次执行失败：

错误码: {If(toolResult.ErrorCode, ExceptionClassifier.CodeUnknown)}
用户可见说明: {If(toolResult.UserMessage, toolResult.Message)}
调试细节: {If(toolResult.DebugDetail, toolResult.Message)}
可自动修复: {toolResult.Recoverable}

原工具调用: {toolCall.ToolId}
原参数: {Newtonsoft.Json.JsonConvert.SerializeObject(toolCall.Parameters)}

当前 {If(session.AppType, "Office")} 可用工具（必须使用原样工具 ID，不要自创 snake_case 或未注册命令）:
{BuildAvailableToolHint(session.AppType, executionContext)}

请分析错误原因并返回修正后的工具调用。只返回 JSON，格式：
```json
{{
  ""toolId"": ""..."",
  ""parameters"": {{...}}
}}
```"

                            Try
                                ' 请求 AI 修复
                                Dim fixedResponse = Await SendRepairRequestAsync(fixPrompt, systemPrompt, toolResult)
                                Dim fixedJson = ExtractJson(fixedResponse)

                                If Not String.IsNullOrEmpty(fixedJson) Then
                                    Dim fixedObj = JObject.Parse(fixedJson)
                                    Dim fixedToolCall = ParseFixedToolCall(fixedObj, toolCall)

                                    ' 使用修复后的工具调用
                                    toolCall = fixedToolCall
                                    AppLogger.Info("LoopEngine", "AI generated repair plan")
                                Else
                                    AppLogger.Warn("LoopEngine", "Unable to parse repair response; stop repair")
                                    Exit While
                                End If
                            Catch ex As Exception
                                AppLogger.Error("LoopEngine", "Repair loop exception", ex)
                                Exit While
                            End Try
                        End If
                    End While

                    ' --- OBSERVE ---
                    session.Status = AgentStatus.Observing
                    Dim observation = FormatObservation(toolResult)

                    ' 如果失败且已达最大修复次数，追加提示
                    If Not toolResult.Success AndAlso fixAttempt >= MaxFixAttempts Then
                        observation &= $" (AI 已尝试自动修复 {fixAttempt} 次，仍然失败)"
                    ElseIf fixAttempt > 0 AndAlso toolResult.Success Then
                        observation &= $" (AI 第 {fixAttempt} 次修复成功)"
                    End If

                    _memory.SetWorking("lastObservation", observation)
                    Dim stepFinishedAt = DateTime.Now
                    Dim explanation = BuildExecutionExplanation(stepIndex, planStep, toolCall, toolResult, fixAttempt, undoPoint, observation, stepStartedAt, stepFinishedAt)
                    planStep.LastExplanation = explanation

                    ' 单轮变更日志：只由宿主观察对象驱动，失败/部分应用同样记录，
                    ' 这样取消或失败后仍能如实告诉用户“已写入哪些地方”。
                    ' 返回的记录同时喂给写入台账，replan 才知道哪些步骤已经完成。
                    Dim changeRecord = session.ChangeLog.AddFromStep(
                        toolCall.ToolId,
                        stepIndex,
                        planStep.Description,
                        toolResult,
                        If(String.IsNullOrWhiteSpace(toolResult.UndoPointId), explanation.UndoPointName, toolResult.UndoPointId),
                        explanation.CanUndo)
                    session.Writes.Record(toolCall, changeRecord, planStep.Description)

                    ' 记录迭代
                    Dim iteration = New ReActIteration With {
                        .Index = session.CurrentIteration,
                        .Thought = thought,
                        .Action = toolCall,
                        .Observation = observation,
                        .Explanation = explanation
                    }
                    session.Iterations.Add(iteration)
                    session.CurrentIteration += 1
                    OnExecutionExplained?.Invoke(explanation)
                    OnIterationUpdate?.Invoke(iteration)

                    ' 更新步骤状态
                    If toolResult.Success Then
                        planStep.Status = StepStatus.Completed
                        noProgressCount = 0
                        OnStepStatusChanged?.Invoke(stepIndex, StepStatusCompleted, toolResult.Message)
                    ElseIf IsRejectedApproval(toolResult) Then
                        ' 拒绝只作用于当前步骤：后续步骤继续执行，终态会说明跳过了几处
                        planStep.Status = StepStatus.Skipped
                        planStep.ErrorMessage = "用户拒绝当前步骤"
                        noProgressCount = 0
                        OnStepStatusChanged?.Invoke(stepIndex, StepStatusSkipped, If(toolResult.UserMessage, toolResult.Message))
                        AppLogger.Info("LoopEngine", $"Step skipped by user approval rejection: {toolResult.ToObserveSummary()}")
                    ElseIf IsAbortedApproval(toolResult) Then
                        ' 超时、取消或新一轮对话取代：审批通道已经关闭，本轮必须整轮停下
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = "审批已放弃"
                        OnStepStatusChanged?.Invoke(stepIndex, StepStatusAborted, If(toolResult.UserMessage, toolResult.Message))
                        session.Status = AgentStatus.Failed
                        Dim abortMsg = AppendWriteSummary(session, "审批已放弃，本轮执行已停止")
                        OnStatusChanged?.Invoke(abortMsg)
                        AppLogger.Warn("LoopEngine", abortMsg)
                        Return AgentResult.Failed(session.Id, abortMsg)
                    Else
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = toolResult.ToObserveSummary()
                        noProgressCount += 1
                        OnStepStatusChanged?.Invoke(stepIndex, StepStatusFailed, If(toolResult.UserMessage, toolResult.Message))
                        AppLogger.Warn("LoopEngine", $"Step failed: {toolResult.ToObserveSummary()}")

                        If Not toolResult.Recoverable Then
                            ' 不可恢复是执行器的判断，不是"用户也只能看着"。已经排在收件箱里的纠偏必须先被
                            ' 走到步骤边界落地，否则用户在卡片上按了重试却当场收到一句终态错误。
                            If session.Interventions.PendingCount > 0 Then
                                AppLogger.Info("LoopEngine",
                                    "Non-recoverable failure deferred to user step recovery at the next step boundary")
                                Continue While
                            End If
                            Dim terminalFailure = $"任务因不可恢复错误停止: {toolResult.ToObserveSummary()}"
                            session.Status = AgentStatus.Failed
                            OnStatusChanged?.Invoke(terminalFailure)
                            AppLogger.Warn("LoopEngine", terminalFailure)
                            Return AgentResult.Failed(session.Id, AppendWriteSummary(session, terminalFailure))
                        End If

                        ' 失败且多轮修复失败，提示可以撤销
                        If fixAttempt >= MaxFixAttempts AndAlso undoPoint IsNot Nothing Then
                            Dim undoHint = _undoManager.GetUndoHint(If(session.AppType, "Unknown"))
                            AppLogger.Info("LoopEngine", $"Execution failed; {undoHint}")
                        End If

                        ' --- REFLECT (连续失败) ---
                        If noProgressCount >= MaxNoProgress Then
                            If replanAttempts >= MaxReplanAttempts Then
                                ' 重规划预算耗尽是"AI 想不出修法"，不是"用户也无话可说"：排在收件箱里的
                                ' 纠偏先走到步骤边界落地，拿不到任何纠偏才按失败收敛。
                                If session.Interventions.PendingCount > 0 Then
                                    AppLogger.Info("LoopEngine",
                                        "Replan budget exhausted but user step recovery is queued; applying at the next boundary")
                                    Continue While
                                End If
                                Dim failMsg = $"步骤多次失败，已达最大重规划次数: {toolResult.ToObserveSummary()}"
                                AppLogger.Error("LoopEngine", failMsg)
                                Return AgentResult.Failed(session.Id, AppendWriteSummary(session, failMsg))
                            End If

                            session.Status = AgentStatus.Reflecting
                            OnStatusChanged?.Invoke("正在分析失败原因并重新规划...")
                            replanAttempts += 1
                            AppLogger.Info("LoopEngine", $"Reflect/replan attempt {replanAttempts}: {toolResult.ToObserveSummary()}")

                            Dim reflection = Await ReflectAndReplanAsync(session, toolResult.ToObserveSummary(), systemPrompt)
                            Select Case reflection?.Strategy
                                Case "retry"
                                    ' 重试同一步骤：不前移索引，避免重跑已成功步骤造成重复写入
                                    planStep.Status = StepStatus.Pending
                                    planStep.ErrorMessage = Nothing
                                    noProgressCount = 0
                                    OnStepStatusChanged?.Invoke(stepIndex, StepStatusPending, "重新尝试这一步")
                                    Continue While
                                Case "skip"
                                    ' 跳过当前失败步骤：标记后前进，已完成步骤保持不动
                                    planStep.Status = StepStatus.Skipped
                                    noProgressCount = 0
                                    OnStepStatusChanged?.Invoke(stepIndex, StepStatusSkipped, "重新规划后跳过这一步")
                                    stepIndex += 1
                                    Continue While
                                Case "replan"
                                    ' 新计划要和"上一版已成功的写入"合起来才算覆盖完整，否则会误判漏项
                                    Dim replanCoverage = ValidateReplannedCoverage(session, reflection.Plan)
                                    If Not String.IsNullOrWhiteSpace(replanCoverage) Then
                                        session.Status = AgentStatus.Failed
                                        OnStatusChanged?.Invoke(replanCoverage)
                                        AppLogger.Warn("LoopEngine", replanCoverage)
                                        Return AgentResult.Failed(session.Id, AppendWriteSummary(session, replanCoverage))
                                    End If

                                    ' 只有全新计划才从第 0 步重跑；台账会挡住已被上一版写成的步骤
                                    session.Plan = reflection.Plan
                                    session.Writes.BeginNextPlan()
                                    session.Interventions.Clear()
                                    stepIndex = 0
                                    noProgressCount = 0
                                    RaisePlanGenerated(session)
                                    OnStatusChanged?.Invoke($"已重新规划（共 {session.Plan.Steps.Count} 步），已写入 {session.ChangeLog.WrittenCount} 处不会被重复执行...")
                                    Continue While
                                Case Else
                                    Dim replanFail = $"重新规划未能给出可执行修复: {toolResult.ToObserveSummary()}"
                                    AppLogger.Error("LoopEngine", replanFail)
                                    Return AgentResult.Failed(session.Id, AppendWriteSummary(session, replanFail))
                            End Select
                        End If
                    End If

                    stepIndex += 1
                End While

                If session.CurrentIteration = 0 Then
                    session.Status = AgentStatus.Failed
                    Dim failMsg = "任务未执行任何工具调用。可能是计划步骤没有生成可解析的 action，或当前宿主工具未加载。"
                    OnStatusChanged?.Invoke(failMsg)
                    AppLogger.Warn("LoopEngine", failMsg)
                    Return AgentResult.Failed(session.Id, AppendWriteSummary(session, failMsg))
                End If

                Dim incompleteSteps = IncompletePlanSteps(session)
                If incompleteSteps.Count > 0 Then
                    session.Status = AgentStatus.Failed
                    Dim failMsg = $"任务未完成，失败/未执行步骤 {incompleteSteps.Count} 个: {String.Join("; ", incompleteSteps.Select(Function(s) s.ErrorMessage).Where(Function(m) Not String.IsNullOrWhiteSpace(m)).Take(3))}"
                    OnStatusChanged?.Invoke(failMsg)
                    AppLogger.Warn("LoopEngine", failMsg)
                    Return AgentResult.Failed(session.Id, AppendWriteSummary(session, failMsg))
                End If

                Dim outcomeError = ValidateExecutionOutcome(session)
                If Not String.IsNullOrWhiteSpace(outcomeError) Then
                    session.Status = AgentStatus.Failed
                    OnStatusChanged?.Invoke(outcomeError)
                    AppLogger.Warn("LoopEngine", outcomeError)
                    Return AgentResult.Failed(session.Id, AppendWriteSummary(session, outcomeError))
                End If

                ' 完成
                session.Status = AgentStatus.Completed
                Dim finalMsg = BuildCompletionMessage(session)
                OnStatusChanged?.Invoke(finalMsg)
                AppLogger.Info("LoopEngine", finalMsg)
                Return AgentResult.SuccessResult(session.Id, finalMsg)

            Catch ex As OperationCanceledException When cancellationToken.IsCancellationRequested
                ' 在途模型请求被"停止"掐断：这是用户主动终止，不能按执行失败收敛。
                Return FinishCancelled(session)
            Catch ex As Exception
                session.Status = AgentStatus.Failed
                Dim classified = ExceptionClassifier.Classify(ex)
                OnStatusChanged?.Invoke($"执行出错: {classified.UserMessage}")
                AppLogger.Error("LoopEngine", "RunAsync unhandled exception", ex)
                Return AgentResult.Failed(session.Id, AppendWriteSummary(session, $"执行异常: [{classified.ErrorCode}] {classified.UserMessage}"))
            End Try
        End Function


        ''' <summary>
        ''' 用户纠偏的唯一落地点（E10-01）：把收件箱里的决定写成步骤状态，需要时把步号退回去。
        ''' 「只重跑受影响步骤」不是注释里的承诺：回退只把目标那一步复位 Pending，其余已判定的步骤
        ''' 由循环顶部的"已判定步骤直接前进"落掉，跨计划版本的重复写入另有 E1-06 台账兜着。
        ''' </summary>
        Private Sub ApplyUserInterventions(session As AgentSession,
                                           ByRef stepIndex As Integer,
                                           ByRef noProgressCount As Integer)
            Dim steps = session.Plan.Steps
            For Each idx In session.Interventions.PendingIndices()
                Dim request = session.Interventions.Take(idx)
                If request Is Nothing OrElse request.Action = StepIntervention.None OrElse
                   idx < 0 OrElse idx >= steps.Count Then Continue For

                Dim target = steps(idx)
                If request.Action = StepIntervention.Skip Then
                    target.Status = StepStatus.Skipped
                    target.ErrorMessage = "按你的要求跳过这一步"
                    OnStepStatusChanged?.Invoke(idx, StepStatusSkipped, "按你的要求跳过这一步")
                    AppLogger.Info("LoopEngine", $"Step {idx + 1} skipped by user intervention")
                    noProgressCount = 0
                    Continue For
                End If

                Dim retryMessage As String = "按你的要求重试这一步"
                If request.Action = StepIntervention.Edit Then
                    ' 说明与参数换成用户的版本，这一步的工具与参数从此归用户（见循环里的 PinnedByUser 分支）。
                    ' 还没跑过的步骤别说"重跑"：那是这一轮第一次执行它。
                    Dim neverRan = target.Status = StepStatus.Pending
                    target.Description = request.Description
                    target.Code = request.Code
                    target.PinnedByUser = True
                    retryMessage = If(neverRan, "按你改过的内容执行这一步", "按你改过的内容重跑这一步")
                End If

                target.Status = StepStatus.Pending
                target.ErrorMessage = Nothing
                OnStepStatusChanged?.Invoke(idx, StepStatusPending, retryMessage)
                AppLogger.Info("LoopEngine", $"Step {idx + 1} queued for {request.Action} by user intervention")
                stepIndex = Math.Min(stepIndex, idx)
                noProgressCount = 0
            Next
        End Sub

        ''' <summary>
        ''' 终稿正文由已落地步骤的观察拼出来。迭代计数是内部记账，直接当回复会让用户
        ''' 只看到"任务完成，共执行 1 个迭代"这种无法核对结果的句子（真机第一轮即此形态）。
        ''' </summary>
        Private Shared Function BuildCompletionMessage(session As AgentSession) As String
            Dim lines As New List(Of String)()
            Dim skippedCount = 0
            If session.Plan IsNot Nothing Then
                For Each planStep In session.Plan.Steps
                    If planStep.Status = StepStatus.Skipped Then skippedCount += 1
                    If planStep.Status <> StepStatus.Completed Then Continue For
                    Dim explanation = planStep.LastExplanation
                    Dim detail = explanation?.Message
                    If String.IsNullOrWhiteSpace(detail) Then detail = explanation?.AfterSummary
                    If String.IsNullOrWhiteSpace(detail) Then detail = planStep.Description
                    If String.IsNullOrWhiteSpace(detail) Then Continue For
                    detail = detail.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
                    If detail.Length > 120 Then detail = detail.Substring(0, 120) & "…"
                    Dim label = explanation?.ToolId
                    If String.IsNullOrWhiteSpace(label) Then label = "步骤" & planStep.StepNumber.ToString()
                    lines.Add($"{lines.Count + 1}. {label}: {detail}")
                    If lines.Count >= 8 Then Exit For
                Next
            End If

            Dim skippedText = If(skippedCount > 0, vbCrLf & $"已跳过 {skippedCount} 个你拒绝的操作。", "")
            If lines.Count = 0 Then
                Return "任务完成，本轮没有修改文档的操作。" & skippedText
            End If
            Return "已完成 " & lines.Count & " 个操作：" & vbCrLf & String.Join(vbCrLf, lines) & skippedText
        End Function

        ''' <summary>
        ''' 协作式停止的收敛出口：只在不打断进行中宿主操作的安全点调用。
        ''' </summary>
        Private Function FinishCancelled(session As AgentSession) As AgentResult
            session.Status = AgentStatus.Aborted
            Dim executed = session.CurrentIteration
            Dim written = If(session.ChangeLog Is Nothing, 0, session.ChangeLog.WrittenCount)
            Dim incomplete = CountIncompleteSteps(session)
            Dim cancelMsg As String
            If written > 0 Then
                cancelMsg = $"已按你的请求停止：已写入 {written} 处，未完成 {incomplete} 步。"
            ElseIf executed > 0 Then
                cancelMsg = $"已按你的请求停止，尚未修改文档；未完成 {incomplete} 步。"
            Else
                cancelMsg = "已按你的请求停止，尚未执行任何操作。"
            End If
            OnStatusChanged?.Invoke(cancelMsg)
            AppLogger.Info("LoopEngine", $"Run cancelled at step boundary, iterations={executed}, written={written}")
            Return AgentResult.CancelledResult(session.Id, cancelMsg)
        End Function

        Private Shared Function CountIncompleteSteps(session As AgentSession) As Integer
            If session Is Nothing OrElse session.Plan Is Nothing Then Return 0
            Return session.Plan.Steps.
                Where(Function(s) s.Status <> StepStatus.Completed AndAlso s.Status <> StepStatus.Skipped).
                Count()
        End Function

        ''' <summary>
        ''' 判断失败是否来自用户在审批里点的“拒绝”，用于把终止范围收窄到当前步骤。
        ''' </summary>
        Private Shared Function IsRejectedApproval(toolResult As ToolResult) As Boolean
            Return String.Equals(GetApprovalDecision(toolResult), "rejected", StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' 判断失败是否来自审批通道被放弃（超时、取消、新一轮取代），此时不能继续执行后续步骤。
        ''' </summary>
        Private Shared Function IsAbortedApproval(toolResult As ToolResult) As Boolean
            Return String.Equals(GetApprovalDecision(toolResult), "aborted", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function GetApprovalDecision(toolResult As ToolResult) As String
            If toolResult Is Nothing OrElse
               Not String.Equals(toolResult.ErrorCode, ExceptionClassifier.CodeSafetyBlocked, StringComparison.OrdinalIgnoreCase) Then
                Return ""
            End If
            If toolResult.Observation Is Nothing Then Return ""
            Try
                Dim observation = JToken.FromObject(toolResult.Observation)
                If observation.Type <> JTokenType.Object Then Return ""
                Dim obj = CType(observation, JObject)
                If Not String.Equals(obj("kind")?.ToString(), "approval", StringComparison.OrdinalIgnoreCase) Then Return ""
                Return If(obj("decision")?.ToString(), "")
            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 失败信息只补事实：回滚成功的步骤不进入“已写入”，避免取消/失败提示和文档实际残留对不上。
        ''' </summary>
        Private Shared Function AppendWriteSummary(session As AgentSession, message As String) As String
            If session Is Nothing OrElse session.ChangeLog Is Nothing Then Return message
            Dim written = session.ChangeLog.WrittenCount
            If written <= 0 Then Return message
            Return $"{message}；本轮已写入 {written} 处，未完成 {CountIncompleteSteps(session)} 步。"
        End Function

        ''' <summary>
        ''' 获取撤销管理器（供外部访问）
        ''' </summary>
        Public ReadOnly Property UndoManager As Core.UndoManager
            Get
                Return _undoManager
            End Get
        End Property

    End Class

End Namespace
