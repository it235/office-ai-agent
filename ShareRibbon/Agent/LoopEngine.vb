Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' ReAct 循环引擎 - 核心执行逻辑
    ''' Think -> Plan -> Act -> Observe -> Reflect
    ''' </summary>
    Public Class LoopEngine
        Private ReadOnly _toolRegistry As ToolRegistry
        Private ReadOnly _memory As AgentMemory
        Private ReadOnly _promptManager As PromptManager

        ' 循环限制
        Private Const MaxIterations As Integer = 15
        Private Const MaxNoProgress As Integer = 3
        Private Const MaxReplanAttempts As Integer = 2

        ' 回调
        Public Property OnStatusChanged As Action(Of String)
        Public Property OnIterationUpdate As Action(Of ReActIteration)
        Public Property OnStepCompleted As Action(Of Integer, Boolean, String)
        Public Property OnRequestApproval As Func(Of String, Task(Of Boolean))
        Public Property OnPlanGenerated As Action(Of ExecutionPlan)
        Public Property SendAIRequest As Func(Of String, String, List(Of HistoryMessage), Task(Of String))

        Public Sub New(toolRegistry As ToolRegistry, memory As AgentMemory, promptManager As PromptManager)
            _toolRegistry = toolRegistry
            _memory = memory
            _promptManager = promptManager
        End Sub

        ''' <summary>
        ''' 执行 ReAct 循环
        ''' </summary>
        Public Async Function RunAsync(session As AgentSession,
                                        systemPrompt As String,
                                        Optional skill As AgentSkill = Nothing) As Task(Of AgentResult)
            Dim noProgressCount As Integer = 0
            Dim replanAttempts As Integer = 0

            Try
                ' Phase 1: 生成 Spec
                OnStatusChanged?.Invoke("正在分析任务...")
                session.Spec = Await GenerateSpecAsync(session)

                ' Phase 2: 生成计划
                OnStatusChanged?.Invoke("正在制定执行计划...")
                session.Plan = Await GeneratePlanAsync(session, systemPrompt, skill)
                If session.Plan Is Nothing OrElse session.Plan.Steps.Count = 0 Then
                    Return AgentResult.Failed(session.Id, "规划失败：无法生成执行计划")
                End If

                ' 通知计划已生成
                OnPlanGenerated?.Invoke(session.Plan)

                ' 简单任务自动执行
                If session.Spec.IsSimple Then
                    OnStatusChanged?.Invoke($"规划完成（共 {session.Plan.Steps.Count} 步），自动执行中...")
                    session.Status = AgentStatus.Executing
                Else
                    ' 等待用户确认
                    session.Status = AgentStatus.WaitingApproval
                    OnStatusChanged?.Invoke($"规划完成，共 {session.Plan.Steps.Count} 个步骤，等待确认")
                    Dim approved = Await WaitForApprovalAsync($"是否执行计划？共 {session.Plan.Steps.Count} 步")
                    If Not approved Then
                        Return AgentResult.Failed(session.Id, "用户取消了执行")
                    End If
                    session.Status = AgentStatus.Executing
                End If

                ' Phase 3: ReAct Loop
                Dim stepIndex As Integer = 0
                While stepIndex < session.Plan.Steps.Count AndAlso session.CurrentIteration < MaxIterations
                    Dim planStep = session.Plan.Steps(stepIndex)
                    planStep.Status = StepStatus.Running

                    ' --- THINK ---
                    session.Status = AgentStatus.Thinking
                    OnStatusChanged?.Invoke($"步骤 {stepIndex + 1}/{session.Plan.Steps.Count}: {planStep.Description}")
                    Dim thought = Await ThinkAsync(session, planStep, systemPrompt)

                    ' --- PARSE ACTION ---
                    Dim toolCall = ParseToolCall(thought)
                    If toolCall Is Nothing Then
                        noProgressCount += 1
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = "无法解析工具调用"
                        OnStepCompleted?.Invoke(stepIndex, False, "解析失败")

                        If noProgressCount >= MaxNoProgress Then Exit While
                        stepIndex += 1
                        Continue While
                    End If

                    ' --- CHECK APPROVAL (risky tools) ---
                    Dim tool = _toolRegistry.GetTool(toolCall.ToolId)
                    If tool IsNot Nothing AndAlso tool.RiskLevel = "risky" Then
                        session.Status = AgentStatus.WaitingApproval
                        Dim approved = Await WaitForApprovalAsync($"步骤 {stepIndex + 1} 包含高风险操作 [{toolCall.ToolId}]，是否继续？")
                        If Not approved Then
                            planStep.Status = StepStatus.Skipped
                            stepIndex += 1
                            session.Status = AgentStatus.Executing
                            Continue While
                        End If
                        session.Status = AgentStatus.Executing
                    End If

                    ' --- ACT ---
                    session.Status = AgentStatus.Executing
                    Dim toolResult = Await _toolRegistry.ExecuteToolAsync(toolCall.ToolId, toolCall.Parameters)

                    ' --- OBSERVE ---
                    session.Status = AgentStatus.Observing
                    Dim observation = FormatObservation(toolResult)
                    _memory.SetWorking("lastObservation", observation)

                    ' 记录迭代
                    Dim iteration = New ReActIteration With {
                        .Index = session.CurrentIteration,
                        .Thought = thought,
                        .Action = toolCall,
                        .Observation = observation
                    }
                    session.Iterations.Add(iteration)
                    session.CurrentIteration += 1
                    OnIterationUpdate?.Invoke(iteration)

                    ' 更新步骤状态
                    If toolResult.Success Then
                        planStep.Status = StepStatus.Completed
                        noProgressCount = 0
                        OnStepCompleted?.Invoke(stepIndex, True, toolResult.Message)
                    Else
                        planStep.Status = StepStatus.Failed
                        planStep.ErrorMessage = toolResult.Message
                        noProgressCount += 1
                        OnStepCompleted?.Invoke(stepIndex, False, toolResult.Message)

                        ' --- REFLECT (连续失败) ---
                        If noProgressCount >= MaxNoProgress Then
                            If replanAttempts >= MaxReplanAttempts Then
                                Return AgentResult.Failed(session.Id, $"步骤多次失败，已达最大重规划次数: {toolResult.Message}")
                            End If

                            session.Status = AgentStatus.Reflecting
                            OnStatusChanged?.Invoke("正在分析失败原因并重新规划...")
                            replanAttempts += 1

                            Dim newPlan = Await ReflectAndReplanAsync(session, toolResult.Message, systemPrompt)
                            If newPlan IsNot Nothing AndAlso newPlan.Steps.Count > 0 Then
                                session.Plan = newPlan
                                stepIndex = 0
                                noProgressCount = 0
                                Continue While
                            Else
                                Return AgentResult.Failed(session.Id, $"重新规划失败: {toolResult.Message}")
                            End If
                        End If
                    End If

                    stepIndex += 1
                End While

                ' 完成
                session.Status = AgentStatus.Completed
                Dim finalMsg = $"任务完成，共执行 {session.CurrentIteration} 个迭代"
                OnStatusChanged?.Invoke(finalMsg)
                Return AgentResult.SuccessResult(session.Id, finalMsg)

            Catch ex As Exception
                session.Status = AgentStatus.Failed
                OnStatusChanged?.Invoke($"执行出错: {ex.Message}")
                Return AgentResult.Failed(session.Id, $"执行异常: {ex.Message}")
            End Try
        End Function

#Region "Private Methods"

        ''' <summary>
        ''' 生成任务 Spec
        ''' </summary>
        Private Async Function GenerateSpecAsync(session As AgentSession) As Task(Of AgentTaskSpec)
            Dim spec As New AgentTaskSpec()
            Try
                Dim prompt = $"分析以下需求，提取结构化任务规格：

需求: {session.UserRequest}

返回 JSON：
```json
{{
  ""goal"": ""一句话描述核心目标"",
  ""constraints"": [""约束1""],
  ""success_criteria"": [""成功标准1""],
  ""complexity"": ""simple|medium|complex""
}}
```

complexity 规则：
- simple：单一操作，步骤数 <= 2，无需用户确认
- medium：2-5个步骤，建议用户确认
- complex：步骤多或逻辑复杂，必须用户确认"

                Dim response = Await SendAIRequest(prompt,
                    "你是一个任务分析专家。只返回JSON，不要解释。", Nothing)

                Dim jsonStr = ExtractJson(response)
                If Not String.IsNullOrEmpty(jsonStr) Then
                    Dim obj = JObject.Parse(jsonStr)
                    spec.Goal = If(obj("goal")?.ToString(), session.UserRequest)
                    spec.Complexity = If(obj("complexity")?.ToString(), "medium")

                    Dim constraints = TryCast(obj("constraints"), JArray)
                    If constraints IsNot Nothing Then
                        For Each c In constraints
                            spec.Constraints.Add(c.ToString())
                        Next
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] Spec生成失败: {ex.Message}")
            End Try
            Return spec
        End Function

        ''' <summary>
        ''' 生成执行计划
        ''' </summary>
        Private Async Function GeneratePlanAsync(session As AgentSession,
                                                  systemPrompt As String,
                                                  skill As AgentSkill) As Task(Of ExecutionPlan)
            Dim plan As New ExecutionPlan()
            Try
                Dim prompt = _promptManager.BuildPlanningPrompt(session, systemPrompt, skill)
                Dim response = Await SendAIRequest(prompt, systemPrompt, _memory.GetRecentMessages(5))

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim obj = JObject.Parse(jsonStr)
                plan.Understanding = obj("understanding")?.ToString()
                plan.Summary = obj("summary")?.ToString()

                Dim stepsArray = TryCast(obj("steps"), JArray)
                If stepsArray IsNot Nothing Then
                    Dim stepNum = 1
                    For Each item In stepsArray
                        plan.Steps.Add(New PlanStep With {
                            .StepNumber = stepNum,
                            .Description = item("description")?.ToString(),
                            .Code = item("code")?.ToString(),
                            .Language = If(item("language")?.ToString(), "json")
                        })
                        stepNum += 1
                    Next
                End If
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 规划失败: {ex.Message}")
                Return Nothing
            End Try
            Return plan
        End Function

        ''' <summary>
        ''' Think：调用LLM生成思考+行动
        ''' </summary>
        Private Async Function ThinkAsync(session As AgentSession,
                                           planStep As PlanStep,
                                           systemPrompt As String) As Task(Of String)
            Dim lastObservation = _memory.GetWorkingString("lastObservation")
            Dim prompt = _promptManager.BuildReactPrompt(planStep, _memory, lastObservation)
            Dim history = _memory.GetRecentMessages(10)
            Return Await SendAIRequest(prompt, systemPrompt, history)
        End Function

        ''' <summary>
        ''' 反思并重新规划
        ''' </summary>
        Private Async Function ReflectAndReplanAsync(session As AgentSession,
                                                      observation As String,
                                                      systemPrompt As String) As Task(Of ExecutionPlan)
            Try
                Dim prompt = _promptManager.BuildReflectionPrompt(session, observation)
                Dim response = Await SendAIRequest(prompt, systemPrompt, Nothing)

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim decision = JObject.Parse(jsonStr)
                Dim strategy = decision("strategy")?.ToString()?.ToLower()

                Select Case strategy
                    Case "retry"
                        ' 重试当前计划
                        Return session.Plan
                    Case "skip"
                        ' 跳过当前步骤继续
                        Return session.Plan
                    Case "replan"
                        ' 重新生成计划
                        Return Await GeneratePlanAsync(session, systemPrompt, session.Skill)
                    Case Else
                        Return Nothing
                End Select
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 反思失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 解析工具调用
        ''' </summary>
        Private Function ParseToolCall(response As String) As ToolCall
            Try
                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return Nothing

                Dim obj = JObject.Parse(jsonStr)
                Dim action = obj("action")
                If action Is Nothing Then Return Nothing

                Dim toolId = action("tool")?.ToString()
                Dim params = TryCast(action("params"), JObject)
                If String.IsNullOrEmpty(toolId) Then Return Nothing
                If params Is Nothing Then params = New JObject()

                Return New ToolCall With {
                    .ToolId = toolId,
                    .Parameters = params
                }
            Catch ex As Exception
                Debug.WriteLine($"[LoopEngine] 解析工具调用失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 格式化观察结果
        ''' </summary>
        Private Function FormatObservation(result As ToolResult) As String
            If result.Success Then
                Return $"✅ [{result.ToolId}] 执行成功: {result.Message}"
            Else
                Return $"❌ [{result.ToolId}] 执行失败: {result.Message}"
            End If
        End Function

        ''' <summary>
        ''' 从响应中提取JSON
        ''' </summary>
        Private Function ExtractJson(response As String) As String
            If String.IsNullOrWhiteSpace(response) Then Return Nothing

            ' 查找 ```json 代码块
            Dim start = response.IndexOf("```json")
            If start >= 0 Then
                start = response.IndexOf("{"c, start)
                If start >= 0 Then
                    Dim endIdx = response.LastIndexOf("}"c)
                    If endIdx > start Then
                        Return response.Substring(start, endIdx - start + 1)
                    End If
                End If
            End If

            ' 查找纯 JSON
            start = response.IndexOf("{"c)
            If start >= 0 Then
                Dim endIdx = response.LastIndexOf("}"c)
                If endIdx > start Then
                    Return response.Substring(start, endIdx - start + 1)
                End If
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 等待用户确认
        ''' </summary>
        Private Async Function WaitForApprovalAsync(message As String) As Task(Of Boolean)
            If OnRequestApproval IsNot Nothing Then
                Return Await OnRequestApproval(message)
            End If
            ' 无回调时默认批准
            Return True
        End Function

#End Region

    End Class

End Namespace
