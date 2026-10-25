Imports System.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Partial Class LoopEngine

        #Region "Plan And Execution Coverage Validation"

        ''' <summary>
        ''' 计划是否覆盖用户明确要求产出的东西。重规划时把上一版已经成功执行的写入
        ''' 一起算进来（doneToolIds / doneHasCreateSlidesImage），否则已完成的部分会被误判成漏项。
        ''' </summary>
        Private Function ValidatePlanCoverage(spec As AgentTaskSpec,
                                              plan As ExecutionPlan,
                                              Optional doneToolIds As HashSet(Of String) = Nothing,
                                              Optional doneHasCreateSlidesImage As Boolean = False) As String
            If spec Is Nothing OrElse plan Is Nothing Then Return ""

            Dim toolIds = GetPlannedToolIds(plan)
            If doneToolIds IsNot Nothing Then
                For Each doneId In doneToolIds
                    toolIds.Add(doneId)
                Next
            End If
            If spec.ExpectedOutputs.Contains("images") AndAlso
               Not toolIds.Contains("InsertImage") AndAlso
               Not PlanContainsCreateSlidesImage(plan) AndAlso
               Not doneHasCreateSlidesImage Then
                Return "规划未覆盖用户要求的真实图片插入；没有可访问图片来源时必须明确报告 capability gap。"
            End If
            If spec.ExpectedOutputs.Contains("images") Then
                Dim imagePathError = ValidatePlannedImagePaths(plan)
                If Not String.IsNullOrWhiteSpace(imagePathError) Then Return imagePathError
            End If
            If spec.ExpectedSlideCount > 0 AndAlso
               Not toolIds.Contains("CreateSlides") AndAlso
               Not toolIds.Contains("InsertSlide") Then
                Return $"规划未覆盖创建 {spec.ExpectedSlideCount} 张幻灯片的要求。"
            End If
            Return ""
        End Function

        Private Function ValidatePlannedImagePaths(plan As ExecutionPlan) As String
            If plan?.Steps Is Nothing Then Return "图片计划为空。"

            For Each stepItem In plan.Steps
                Try
                    Dim envelope = JObject.Parse(If(stepItem.Code, ""))
                    Dim commands As New List(Of JObject)()
                    If envelope("command") IsNot Nothing Then commands.Add(envelope)
                    Dim commandArray = TryCast(envelope("commands"), JArray)
                    If commandArray IsNot Nothing Then commands.AddRange(commandArray.OfType(Of JObject)())

                    For Each commandObj In commands
                        Dim commandName = commandObj("command")?.ToString()
                        If String.Equals(commandName, "InsertImage", StringComparison.OrdinalIgnoreCase) Then
                            Dim params = TryCast(commandObj("params"), JObject)
                            Dim imagePath = params?("imagePath")?.ToString()
                            If String.IsNullOrWhiteSpace(imagePath) Then
                                Return "图片计划缺少 imagePath，不能开始部分修改。"
                            End If
                            If Not IO.File.Exists(imagePath) Then
                                Return $"图片文件不可访问：{imagePath}。为避免只创建文字页，任务尚未执行。"
                            End If
                        ElseIf String.Equals(commandName, "CreateSlides", StringComparison.OrdinalIgnoreCase) Then
                            Dim params = TryCast(commandObj("params"), JObject)
                            Dim slides = TryCast(params?("slides"), JArray)
                            If slides Is Nothing Then Continue For
                            For Each slide In slides.OfType(Of JObject)()
                                Dim imagePath = slide("imagePath")?.ToString()
                                If Not String.IsNullOrWhiteSpace(imagePath) AndAlso Not IO.File.Exists(imagePath) Then
                                    Return $"图片文件不可访问：{imagePath}。为避免只创建文字页，任务尚未执行。"
                                End If
                            Next
                        End If
                    Next
                Catch
                    ' Invalid plan JSON is handled by the normal tool-call parsing path.
                End Try
            Next
            Return ""
        End Function

        Private Function PlanContainsCreateSlidesImage(plan As ExecutionPlan) As Boolean
            If plan?.Steps Is Nothing Then Return False

            For Each stepItem In plan.Steps
                Try
                    Dim envelope = JObject.Parse(If(stepItem.Code, ""))
                    Dim commands As New List(Of JObject)()
                    If envelope("command") IsNot Nothing Then commands.Add(envelope)
                    Dim commandArray = TryCast(envelope("commands"), JArray)
                    If commandArray IsNot Nothing Then commands.AddRange(commandArray.OfType(Of JObject)())

                    For Each commandObj In commands
                        If String.Equals(commandObj("command")?.ToString(), "CreateSlides", StringComparison.OrdinalIgnoreCase) AndAlso
                           CreateSlidesParametersContainImage(TryCast(commandObj("params"), JObject)) Then
                            Return True
                        End If
                    Next
                Catch
                    ' Invalid plan JSON is handled by the normal tool-call parsing path.
                End Try
            Next
            Return False
        End Function

        Private Shared Function CreateSlidesParametersContainImage(parameters As JObject) As Boolean
            Dim slides = TryCast(parameters?("slides"), JArray)
            If slides Is Nothing Then Return False
            Return slides.OfType(Of JObject)().Any(
                Function(slide) Not String.IsNullOrWhiteSpace(slide("imagePath")?.ToString()))
        End Function

        ''' <summary>
        ''' 本轮成功执行过的工具调用。以迭代记录为唯一依据，跨计划版本累计，
        ''' 因此重规划后"上一版已做成"的动作在这里仍然算数。
        ''' </summary>
        Private Shared Function GetSuccessfulActions(session As AgentSession) As List(Of ToolCall)
            If session?.Iterations Is Nothing Then Return New List(Of ToolCall)()
            Return session.Iterations.
                Where(Function(item) item IsNot Nothing AndAlso
                                     item.Explanation IsNot Nothing AndAlso
                                     item.Explanation.Success AndAlso
                                     item.Action IsNot Nothing).
                Select(Function(item) item.Action).
                ToList()
        End Function

        Private Function ValidateReplannedCoverage(session As AgentSession, plan As ExecutionPlan) As String
            If session?.Spec Is Nothing OrElse plan Is Nothing Then Return ""

            Dim doneToolIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim doneHasCreateSlidesImage As Boolean = False
            For Each action In GetSuccessfulActions(session)
                doneToolIds.Add(action.ToolId)
                If String.Equals(action.ToolId, "CreateSlides", StringComparison.OrdinalIgnoreCase) AndAlso
                   CreateSlidesParametersContainImage(action.Parameters) Then
                    doneHasCreateSlidesImage = True
                End If
            Next
            Return ValidatePlanCoverage(session.Spec, plan, doneToolIds, doneHasCreateSlidesImage)
        End Function

        ''' <summary>
        ''' 命中写入台账时给出的成功结果：不执行、不申请撤销点，观察 kind 不是 write，
        ''' 因此变更日志不会多出重复行，但时间线和终态校验仍能看出这一步已被视为完成。
        ''' </summary>
        Private Shared Function BuildAlreadyAppliedResult(toolCall As ToolCall, priorWrite As AppliedWrite) As ToolResult
            Dim summary = $"该写入已在计划版本 {priorWrite.PlanGeneration} 完成，跳过以避免重复写入"
            Dim observation As New JObject From {
                {"kind", "idempotence"},
                {"summary", summary},
                {"changed", False},
                {"warnings", New JArray("duplicate_write_skipped")}
            }
            Return ToolResult.Succeed(If(toolCall?.ToolId, ""), summary, Nothing, observation)
        End Function

        Private Function ValidateExecutionOutcome(session As AgentSession) As String
            If session Is Nothing OrElse session.Spec Is Nothing Then Return ""

            Dim successfulActions = GetSuccessfulActions(session)
            If session.Spec.ExpectedOutputs.Contains("images") AndAlso
               Not successfulActions.Any(
                   Function(toolCall) String.Equals(toolCall.ToolId, "InsertImage", StringComparison.OrdinalIgnoreCase) OrElse
                                      (String.Equals(toolCall.ToolId, "CreateSlides", StringComparison.OrdinalIgnoreCase) AndAlso
                                       CreateSlidesParametersContainImage(toolCall.Parameters))) Then
                Return "任务未完成：用户要求插入图片，但执行记录中没有成功产生真实图片的操作。"
            End If

            If session.Spec.ExpectedSlideCount > 0 Then
                Dim created As Integer = 0
                For Each action In successfulActions
                    If String.Equals(action.ToolId, "CreateSlides", StringComparison.OrdinalIgnoreCase) Then
                        Dim slides = TryCast(action.Parameters?("slides"), JArray)
                        If slides IsNot Nothing Then created += slides.Count
                    ElseIf String.Equals(action.ToolId, "InsertSlide", StringComparison.OrdinalIgnoreCase) Then
                        created += 1
                    End If
                Next
                If created < session.Spec.ExpectedSlideCount Then
                    Return $"任务未完成：要求创建 {session.Spec.ExpectedSlideCount} 张幻灯片，执行记录仅确认 {created} 张。"
                End If
            End If

            Dim packageReason = ValidateTaskPackageOutcome(session)
            If Not String.IsNullOrWhiteSpace(packageReason) Then Return packageReason
            Return ""
        End Function

        ''' <summary>
        ''' 入口动作所属任务包的成功标准，在 run 收尾时真的对着本轮变更台账核一遍（E6-05）。
        ''' 没有入口身份就不核对：自由文本不能靠猜绑到某个包上，猜错一次，用户看到的"未完成"
        ''' 就不是他自己要求的那件事。
        ''' </summary>
        Private Function ValidateTaskPackageOutcome(session As AgentSession) As String
            Dim pkg = MatchedTaskPackage(session)
            If pkg Is Nothing Then Return ""

            Dim reasons = TaskPackageOutcomeVerifier.UnmetReasons(pkg, session.ChangeLog)
            If reasons.Count = 0 Then Return ""

            If String.Equals(TaskPackageOutcomeVerifier.OnFail(pkg), "rollback", StringComparison.OrdinalIgnoreCase) Then
                ' rollback 需要 run 末把本轮改动交回宿主还原，这条路由尚未接通；如实记一笔再按未达成报告。
                AppLogger.Warn("LoopEngine", $"task package {pkg.Id} declared onFail=rollback but no run-end restore route exists; reported unmet")
            End If
            Return $"任务未完成（{pkg.Id} 的成功标准未达成）：{String.Join("；", reasons)}"
        End Function

        ''' <summary>入口身份 → 任务包；自由文本或没有登记的包都返回 Nothing（不猜包）。</summary>
        Private Shared Function MatchedTaskPackage(session As AgentSession) As TaskPackage
            If session Is Nothing OrElse String.IsNullOrWhiteSpace(session.EntryActionId) Then Return Nothing
            Return TaskPackageCatalog.FindByAction(session.EntryActionId)
        End Function

        ''' <summary>本轮命中包的未达成断言；空表表示「无需修复」，与「修复失败」是两件事。</summary>
        Private Function UnmetTaskPackageReasons(session As AgentSession) As List(Of String)
            Dim pkg = MatchedTaskPackage(session)
            If pkg Is Nothing Then Return New List(Of String)()
            Return TaskPackageOutcomeVerifier.UnmetReasons(pkg, session.ChangeLog)
        End Function

        ''' <summary>包声明的失败策略是否要求「去修」。repair 之外（report/rollback/空）都只报告。</summary>
        Private Shared Function PackageWantsRunEndRepair(session As AgentSession) As Boolean
            Dim pkg = MatchedTaskPackage(session)
            Return pkg IsNot Nothing AndAlso
                   String.Equals(TaskPackageOutcomeVerifier.OnFail(pkg), "repair", StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' run 末的合同修复（E6-05b）：把未达成的断言当作观察结果，交给既有的 reflect/replan 路由，
        ''' 而不是新建第二套 Repair Loop。只接受 replan——run 收尾没有"当前失败步骤"，
        ''' retry/skip 语义无处落地；任何拿不到可用新计划的情况都返回 Nothing，由调用方按未达成报告。
        ''' </summary>
        Private Async Function RequestRunEndRepairPlanAsync(session As AgentSession,
                                                             reasons As List(Of String),
                                                             systemPrompt As String) As Task(Of ExecutionPlan)
            Dim pkg = MatchedTaskPackage(session)
            ' 前缀是给日志和探针看的机器可查标记：观察正文是中文，探针不能靠匹配中文判断这条走的是哪条路由。
            Dim observation = $"RUN_END_CONTRACT_UNMET {pkg.Id}：成功标准未达成 - {String.Join("；", reasons)}。" &
                              $"本轮已落地写入 {session.ChangeLog.WrittenCount} 处，请只规划还缺的那部分，不要重复已经写入的内容。"
            session.Status = AgentStatus.Reflecting
            OnStatusChanged?.Invoke("成功标准未达成，正在按任务包合同重新规划...")
            AppLogger.Info("LoopEngine", $"run-end package repair requested for {pkg.Id}: {AppLogger.Redact(observation)}")

            Dim decision = Await ReflectAndReplanAsync(session, observation, systemPrompt)
            If decision?.Strategy <> "replan" Then
                AppLogger.Warn("LoopEngine",
                    $"run-end repair for {pkg.Id} got strategy '{If(decision?.Strategy, "none")}', reported unmet instead")
                Return Nothing
            End If

            Dim coverage = ValidateReplannedCoverage(session, decision.Plan)
            If Not String.IsNullOrWhiteSpace(coverage) Then
                AppLogger.Warn("LoopEngine", $"run-end repair plan for {pkg.Id} lost coverage: {coverage}")
                Return Nothing
            End If
            Return decision.Plan
        End Function

        ''' <summary>还没落地也没被跳过的步骤。闸门和终态判定共用这一份口径。</summary>
        Private Shared Function IncompletePlanSteps(session As AgentSession) As List(Of PlanStep)
            If session?.Plan?.Steps Is Nothing Then Return New List(Of PlanStep)()
            Return session.Plan.Steps.
                Where(Function(s) s.Status <> StepStatus.Completed AndAlso s.Status <> StepStatus.Skipped).
                ToList()
        End Function

        Private Function GetPlannedToolIds(plan As ExecutionPlan) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If plan?.Steps Is Nothing Then Return result

            For Each stepItem In plan.Steps
                Try
                    Dim obj = JObject.Parse(If(stepItem.Code, ""))
                    Dim command = obj("command")?.ToString()
                    If Not String.IsNullOrWhiteSpace(command) Then result.Add(command)
                    Dim commands = TryCast(obj("commands"), JArray)
                    If commands Is Nothing Then Continue For
                    For Each item In commands.OfType(Of JObject)()
                        command = item("command")?.ToString()
                        If Not String.IsNullOrWhiteSpace(command) Then result.Add(command)
                    Next
                Catch
                    ' Invalid plan JSON is handled by the normal tool-call parsing path.
                End Try
            Next
            Return result
        End Function

        #End Region

        #Region "Plan Trust Card"

        ''' <summary>
        ''' 计划卡只在第一版发出去的话，重规划后用户看到的时间线仍是被作废的旧计划。
        ''' 这里连同按事实算出的范围/数量/风险一起发出，前端按同一会话原位重绘并标出版本号。
        ''' </summary>
        Private Sub RaisePlanGenerated(session As AgentSession)
            If OnPlanGenerated Is Nothing OrElse session?.Plan Is Nothing Then Return
            session.Plan.Trust = BuildPlanTrust(session, session.Plan)
            OnPlanGenerated.Invoke(session.Plan)
        End Sub

        Private Function BuildPlanTrust(session As AgentSession, plan As ExecutionPlan) As PlanTrust
            Dim trust As New PlanTrust With {
                .PlanVersion = Math.Max(1, If(session?.Writes?.PlanGeneration, 1)),
                .StepCount = If(plan?.Steps?.Count, 0),
                .ScopeText = BuildScopeText(session)
            }

            Dim toolIds = GetPlannedToolIds(plan)
            Dim resolved As Integer = 0
            For Each toolId In toolIds
                Dim descriptor = _toolRegistry.GetTool(toolId)
                If descriptor Is Nothing Then Continue For
                resolved += 1
                Dim level = NormalizeRiskLevel(descriptor.RiskLevel)
                If RiskRank(level) > RiskRank(trust.RiskLevel) Then trust.RiskLevel = level
                If Not String.Equals(level, "safe", StringComparison.OrdinalIgnoreCase) Then
                    trust.RiskyTools.Add(If(String.IsNullOrWhiteSpace(descriptor.Name), descriptor.Id, descriptor.Name))
                End If
            Next
            trust.ToolCount = resolved
            Return trust
        End Function

        ''' <summary>
        ''' 范围来自当前上下文包，不让模型复述一遍“我要改哪一段”：模型说的范围不一定属实。
        ''' </summary>
        Private Function BuildScopeText(session As AgentSession) As String
            Dim pack = TryCast(_memory?.GetWorking("lastContextPack"), Context.ContextPack)
            Dim appType = If(session?.AppType, If(pack?.AppType, "当前文档"))
            If pack Is Nothing Then Return $"{appType} 当前文档"

            Dim documentName = CompactText(pack.Document?.Name, 40)
            Dim target = If(documentName.Length = 0, appType, $"{appType}《{documentName}》")
            Dim selection = pack.Selection
            If selection IsNot Nothing AndAlso
               (selection.ItemCount > 0 OrElse Not String.IsNullOrWhiteSpace(selection.Address)) Then
                Dim countText = If(selection.ItemCount > 0, $"（{selection.ItemCount} 项）", "")
                Return $"{target} 当前选区{countText}"
            End If
            Return $"{target} 整篇文档"
        End Function

        ''' <summary>注册表里出现过的未知风险值按最严格处理，和 SafetyGate 缺省一致。</summary>
        Private Shared Function NormalizeRiskLevel(raw As String) As String
            Dim value = If(raw, "").Trim().ToLowerInvariant()
            Select Case value
                Case "safe", "medium", "risky"
                    Return value
                Case Else
                    Return "risky"
            End Select
        End Function

        Private Shared Function RiskRank(level As String) As Integer
            Select Case level
                Case "risky"
                    Return 2
                Case "medium"
                    Return 1
                Case Else
                    Return 0
            End Select
        End Function

        Private Shared Function CompactText(value As String, maxLength As Integer) As String
            Dim text = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If text.Length <= maxLength Then Return text
            Return text.Substring(0, maxLength) & "…"
        End Function

        #End Region

    End Class

End Namespace
