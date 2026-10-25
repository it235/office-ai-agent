Imports System.Text
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    Public Partial Class LoopEngine

        #Region "Planning, Repair, And Explanation Helpers"

        Private Function BuildExecutionExplanation(stepIndex As Integer,
                                                   planStep As PlanStep,
                                                   toolCall As ToolCall,
                                                   toolResult As ToolResult,
                                                   fixAttempts As Integer,
                                                   undoPoint As Core.UndoManager.UndoPoint,
                                                   observation As String,
                                                   startedAt As DateTime,
                                                   finishedAt As DateTime) As ExecutionExplanation
            Dim tool = If(toolCall Is Nothing, Nothing, _toolRegistry.GetTool(toolCall.ToolId))
            Dim toolId = If(toolCall?.ToolId, "")
            Dim toolName = If(tool?.Name, toolId)
            Dim category = If(tool?.Category, "")
            Dim risk = If(tool?.RiskLevel, "unknown")
            Dim paramsJson = If(toolCall?.Parameters Is Nothing, "{}", toolCall.Parameters.ToString(Formatting.None))
            Dim success = toolResult IsNot Nothing AndAlso toolResult.Success
            Dim message = If(toolResult?.Message, "")
            Dim skillName = ExtractResultDataValue(toolResult, "skillName")
            Dim scriptFileName = ExtractResultDataValue(toolResult, "scriptFileName")
            Dim mcpToolName = ExtractResultDataValue(toolResult, "mcpToolName")
            Dim mcpStatus = ExtractResultDataValue(toolResult, "mcpStatus")
            Dim failureReason = If(success, "", ExtractResultDataValue(toolResult, "failureReason"))
            If String.IsNullOrWhiteSpace(failureReason) AndAlso Not success Then failureReason = message
            Dim verb = If(success, "已完成", "未完成")
            Dim fixText = If(fixAttempts > 0, $"，期间自动修复 {fixAttempts} 次", "")
            Dim skillText = If(String.IsNullOrWhiteSpace(skillName), "", $"，Skill: {skillName}")
            Dim scriptText = If(String.IsNullOrWhiteSpace(scriptFileName), "", $"，脚本: {scriptFileName}")
            Dim mcpText = If(String.IsNullOrWhiteSpace(mcpToolName), "", $"，MCP: {mcpToolName}")
            Dim elapsedMs = CLng(Math.Max(0, (finishedAt - startedAt).TotalMilliseconds))
            If toolResult IsNot Nothing AndAlso toolResult.ElapsedMs > 0 Then elapsedMs = toolResult.ElapsedMs
            Dim undoHint = If(undoPoint Is Nothing, "", _undoManager.GetUndoHint(If(undoPoint.AppType, "")))
            Dim undoPointName = If(undoPoint?.Name, "")
            Dim canUndo = undoPoint IsNot Nothing AndAlso undoPoint.CanUndo
            Dim beforeSummary = $"准备执行步骤 {stepIndex + 1}: {If(planStep?.Description, "")}；工具: {toolId}；参数: {paramsJson}"
            Dim afterSummary = If(observation, message)
            Dim observationJson = SerializeCompactJson(If(toolResult?.Observation, Nothing), 8192)
            Dim dataSummaryJson = SerializeCompactJson(BuildDataSummary(toolResult), 4096)
            Dim repairSummary = If(fixAttempts > 0,
                                   If(success, $"AI 自动修复 {fixAttempts} 次后成功", $"AI 自动修复 {fixAttempts} 次后仍失败"),
                                   "")
            Dim text = $"步骤 {stepIndex + 1} {verb}：调用 {toolName}（{toolId}）{skillText}{scriptText}{mcpText}{fixText}。{message}"

            Return New ExecutionExplanation With {
                .StepIndex = stepIndex,
                .StepDescription = If(planStep?.Description, ""),
                .ToolId = toolId,
                .ToolName = toolName,
                .ToolCategory = category,
                .RiskLevel = risk,
                .ParametersJson = paramsJson,
                .StartedAt = startedAt,
                .FinishedAt = finishedAt,
                .ElapsedMs = elapsedMs,
                .BeforeSummary = beforeSummary,
                .AfterSummary = afterSummary,
                .ObservationJson = observationJson,
                .DataSummaryJson = dataSummaryJson,
                .Success = success,
                .Message = message,
                .SkillName = skillName,
                .ScriptFileName = scriptFileName,
                .McpToolName = mcpToolName,
                .McpStatus = mcpStatus,
                .FailureReason = failureReason,
                .UndoPointName = undoPointName,
                .UndoHint = undoHint,
                .CanUndo = canUndo,
                .AutoRepairSummary = repairSummary,
                .FixAttempts = fixAttempts,
                .ExplanationText = text
            }
        End Function

        Private Function BuildDataSummary(toolResult As ToolResult) As Object
            If toolResult Is Nothing OrElse toolResult.Data Is Nothing Then Return Nothing

            Try
                Dim token = JToken.FromObject(toolResult.Data)
                If token.Type = JTokenType.Array Then
                    Dim arr = CType(token, JArray)
                    Return New JObject From {
                        {"type", "array"},
                        {"count", arr.Count},
                        {"preview", CloneFirstItems(arr, 5)}
                    }
                End If

                If token.Type = JTokenType.Object Then
                    Dim obj = CType(token, JObject)
                    Dim summary As New JObject From {
                        {"type", "object"},
                        {"keys", BuildStringArray(obj.Properties().Select(Function(p) p.Name).Take(20))}
                    }

                    If obj("total") IsNot Nothing Then summary("total") = obj("total")
                    If obj("returned") IsNot Nothing Then summary("returned") = obj("returned")
                    If obj("truncated") IsNot Nothing Then summary("truncated") = obj("truncated")
                    If obj("items") IsNot Nothing AndAlso obj("items").Type = JTokenType.Array Then
                        summary("itemsPreview") = CloneFirstItems(CType(obj("items"), JArray), 5)
                    End If

                    Return summary
                End If

                Return token
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", "BuildDataSummary failed", ex)
                Return Nothing
            End Try
        End Function

        Private Function CloneFirstItems(items As JArray, maxItems As Integer) As JArray
            Dim result As New JArray()
            If items Is Nothing Then Return result

            For Each item In items.Take(Math.Max(0, maxItems))
                result.Add(item.DeepClone())
            Next

            Return result
        End Function

        Private Function BuildStringArray(items As IEnumerable(Of String)) As JArray
            Dim result As New JArray()
            If items Is Nothing Then Return result

            For Each item In items
                result.Add(If(item, ""))
            Next

            Return result
        End Function

        Private Function SerializeCompactJson(value As Object, maxLength As Integer) As String
            If value Is Nothing Then Return ""

            Try
                Dim json = JsonConvert.SerializeObject(value, Formatting.None)
                If maxLength > 0 AndAlso json.Length > maxLength Then
                    Return json.Substring(0, maxLength) & "...(truncated)"
                End If
                Return json
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", "SerializeCompactJson failed", ex)
                Return ""
            End Try
        End Function

        Private Function ExtractResultDataValue(toolResult As ToolResult, key As String) As String
            If toolResult Is Nothing OrElse toolResult.Data Is Nothing OrElse String.IsNullOrWhiteSpace(key) Then Return ""

            Try
                Dim obj = JObject.FromObject(toolResult.Data)
                Dim token = obj.SelectToken(key)
                If token Is Nothing Then Return ""
                Return token.ToString()
            Catch ex As Exception
                AppLogger.Debug("LoopEngine", $"ExtractResultDataValue key={key} failed", ex)
                Return ""
            End Try
        End Function

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
                                                  skill As AgentSkill,
                                                  Optional attempt As Integer = 0) As Task(Of ExecutionPlan)
            Dim plan As New ExecutionPlan()
            Dim failureReason As String = ""
            Try
                Dim prompt = _promptManager.BuildPlanningPrompt(session, systemPrompt, skill)
                Dim response = Await SendAIRequest(prompt, systemPrompt, _memory.GetRecentMessages(5))

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then
                    failureReason = "响应中没有 JSON"
                Else
                    Dim obj = JObject.Parse(jsonStr)
                    plan.Understanding = obj("understanding")?.ToString()
                    plan.Summary = obj("summary")?.ToString()
                    plan.CapabilityGap = obj("capabilityGap")?.ToString()

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
                End If
            Catch ex As Exception
                failureReason = ex.Message
            End Try

            If plan.Steps.Count = 0 AndAlso String.IsNullOrWhiteSpace(plan.CapabilityGap) Then
                If String.IsNullOrWhiteSpace(failureReason) Then failureReason = "响应没有 steps 或 capabilityGap"
                AppLogger.Warn("LoopEngine", $"规划无效 attempt={attempt}: {AppLogger.Redact(failureReason)}")
                If attempt = 0 Then
                    Dim correctedPrompt = systemPrompt & vbCrLf &
                        "【规划纠错】上次响应不是有效计划。必须返回严格 JSON，并且提供非空 steps 或明确 capabilityGap。"
                    Return Await GeneratePlanAsync(session, correctedPrompt, skill, attempt + 1)
                End If
                Return Nothing
            End If
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
        ''' 反思决策：携带策略与（可选）新计划，供调用方按语义处理步骤索引，
        ''' 避免 retry/skip 一律重置到第 0 步而重跑已成功步骤造成重复写入。
        ''' </summary>
        Private Class ReflectionDecision
            Public Property Strategy As String = "none"
            Public Property Plan As ExecutionPlan
            Public Shared ReadOnly Property NoAction As ReflectionDecision
                Get
                    Return New ReflectionDecision With {.Strategy = "none"}
                End Get
            End Property
        End Class

        ''' <summary>
        ''' 反思并给出修复策略。提示词必须带上本轮已落盘的写入清单，
        ''' 否则模型会重新规划出已经执行过的步骤，新计划从头执行时造成重复写入（E1-06）。
        ''' </summary>
        Private Async Function ReflectAndReplanAsync(session As AgentSession,
                                                      observation As String,
                                                      systemPrompt As String) As Task(Of ReflectionDecision)
            Try
                ' 已落盘写入是这一轮的客观事实，作为系统级约束下发，反思和重新规划共用同一份。
                Dim appliedConstraint = session.Writes.BuildAppliedWritesText()
                Dim prompt = _promptManager.BuildReflectionPrompt(session, observation)
                Dim response = Await SendAIRequest(prompt, systemPrompt & appliedConstraint, Nothing)

                Dim jsonStr = ExtractJson(response)
                If String.IsNullOrEmpty(jsonStr) Then Return ReflectionDecision.NoAction

                Dim decision = JObject.Parse(jsonStr)
                Dim strategy = decision("strategy")?.ToString()?.Trim()?.ToLowerInvariant()

                Select Case strategy
                    Case "retry"
                        ' 重试当前步骤：沿用现有计划，由调用方保持在同一步骤，不重置索引
                        Return New ReflectionDecision With {.Strategy = "retry", .Plan = session.Plan}
                    Case "skip"
                        ' 跳过当前失败步骤：由调用方前进一步，已完成步骤保持不动
                        Return New ReflectionDecision With {.Strategy = "skip", .Plan = session.Plan}
                    Case "replan"
                        ' 生成全新计划：调用方会推进计划版本，已落盘的写入由台账挡住不再执行
                        Dim newPlan = Await GeneratePlanAsync(session, systemPrompt & appliedConstraint, session.Skill)
                        If newPlan IsNot Nothing AndAlso newPlan.Steps.Count > 0 Then
                            Return New ReflectionDecision With {.Strategy = "replan", .Plan = newPlan}
                        End If
                        Return ReflectionDecision.NoAction
                    Case Else
                        Return ReflectionDecision.NoAction
                End Select
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Reflect/replan failed: {AppLogger.Redact(ex.Message)}")
                Return ReflectionDecision.NoAction
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

        Private Function ParseFixedToolCall(fixedObj As JObject, fallback As ToolCall) As ToolCall
            If fixedObj Is Nothing Then Return fallback

            Dim toolId = fixedObj("toolId")?.ToString()
            Dim params = TryCast(fixedObj("parameters"), JObject)

            If String.IsNullOrWhiteSpace(toolId) Then
                Dim action = TryCast(fixedObj("action"), JObject)
                If action IsNot Nothing Then
                    toolId = action("tool")?.ToString()
                    If params Is Nothing Then params = TryCast(action("params"), JObject)
                End If
            End If

            If String.IsNullOrWhiteSpace(toolId) Then toolId = fallback.ToolId
            If params Is Nothing Then params = If(fallback.Parameters, New JObject())

            Return New ToolCall With {
                .ToolId = toolId,
                .Parameters = params
            }
        End Function

        Private Function BuildAvailableToolHint(appType As String,
                                                Optional executionContext As ToolExecutionContext = Nothing) As String
            Dim tools = _toolRegistry.GetVisibleTools(appType, executionContext).
                OrderBy(Function(t) t.Category).
                ThenBy(Function(t) t.Id).
                Take(40).
                Select(Function(t) $"- {t.Id}: {t.Name}")
            Return String.Join(vbCrLf, tools)
        End Function

        ''' <summary>
        ''' 格式化观察结果
        ''' </summary>
        Private Function FormatObservation(result As ToolResult) As String
            If result Is Nothing Then
                Return "❌ [unknown] 无工具结果"
            End If
            If result.Success Then
                Dim summary = result.ToObserveSummary()
                Dim dataSummary = FormatResultData(result.Data)
                If Not String.IsNullOrWhiteSpace(dataSummary) Then
                    Return $"✅ [{result.ToolId}] 执行成功: {summary}{vbCrLf}data={dataSummary}"
                End If
                Return $"✅ [{result.ToolId}] 执行成功: {summary}"
            End If
            ' Structured observe payload for repair/reflect (P0-4).
            Return $"❌ [{result.ToolId}] {result.ToObserveSummary()}"
        End Function

        ''' <summary>
        ''' Sends one repair request with ephemeral visual evidence when the host supplied it.
        ''' Evidence is never copied into the textual prompt, history, memory, or trace. If the
        ''' configured provider/model rejects image input, the existing text-only path remains
        ''' the deterministic fallback for the same repair attempt.
        ''' </summary>
        Private Async Function SendRepairRequestAsync(fixPrompt As String,
                                                      systemPrompt As String,
                                                      result As ToolResult) As Task(Of String)
            If _multimodalRepairDisabledForRun OrElse
               result Is Nothing OrElse
               result.VisualEvidence Is Nothing OrElse
               result.VisualEvidence.Count = 0 OrElse
               SendAIRequestWithMessages Is Nothing Then
                Return Await SendAIRequest(fixPrompt, systemPrompt, Nothing)
            End If

            Dim content As New JArray From {
                New JObject From {
                    {"type", "text"},
                    {"text", fixPrompt & vbCrLf & vbCrLf &
                        "请结合附带的实际 PowerPoint 渲染截图诊断视觉问题；只返回约定的修正后工具调用 JSON。"}
                }
            }
            Dim evidenceCount As Integer = 0
            For Each item In result.VisualEvidence
                If item Is Nothing OrElse evidenceCount >= 2 Then Exit For
                If item.ByteLength > 0 AndAlso
                   item.ByteLength <= 2 * 1024 * 1024 AndAlso
                   Not String.IsNullOrWhiteSpace(item.DataUrl) AndAlso
                   item.DataUrl.Length <= 3 * 1024 * 1024 AndAlso
                   item.DataUrl.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase) Then
                    content.Add(New JObject From {
                        {"type", "image_url"},
                        {"image_url", New JObject From {
                            {"url", item.DataUrl},
                            {"detail", "low"}
                        }}
                    })
                    evidenceCount += 1
                End If
            Next
            If evidenceCount = 0 Then Return Await SendAIRequest(fixPrompt, systemPrompt, Nothing)

            Dim messages As New JArray()
            If Not String.IsNullOrWhiteSpace(systemPrompt) Then
                messages.Add(New JObject From {
                    {"role", "system"},
                    {"content", systemPrompt}
                })
            End If
            messages.Add(New JObject From {
                {"role", "user"},
                {"content", content}
            })

            Try
                Dim response = Await SendAIRequestWithMessages(messages)
                If Not String.IsNullOrWhiteSpace(response) Then Return response
                _multimodalRepairDisabledForRun = True
                AppLogger.Warn("LoopEngine", "Multimodal repair returned an empty response; falling back to text-only repair")
            Catch ex As Exception
                _multimodalRepairDisabledForRun = True
                AppLogger.Warn("LoopEngine", $"Multimodal repair unavailable; falling back to text-only repair: {AppLogger.Redact(ex.Message)}")
            End Try
            Return Await SendAIRequest(fixPrompt, systemPrompt, Nothing)
        End Function

        ''' <summary>
        ''' A host call returning Success=True is not sufficient for a write operation.
        ''' Structured write observations must confirm an actual change, and operation
        ''' batches must not contain unhandled failed/partial steps.
        ''' </summary>
        Private Function ValidateObservedOutcome(result As ToolResult) As ToolResult
            If result Is Nothing OrElse Not result.Success OrElse result.Observation Is Nothing Then Return result

            Try
                Dim observation = TryCast(result.Observation, JToken)
                If observation Is Nothing Then observation = JToken.FromObject(result.Observation)
                If observation Is Nothing OrElse observation.Type <> JTokenType.Object Then Return result

                Dim kind = If(observation("kind")?.ToString(), "").Trim().ToLowerInvariant()
                Dim writeExpectedToken = observation("writeExpected")
                Dim writeExpected = writeExpectedToken IsNot Nothing AndAlso
                                    writeExpectedToken.Type = JTokenType.Boolean AndAlso
                                    writeExpectedToken.Value(Of Boolean)()
                If kind <> "write" AndAlso (kind <> "office_operation_batch" OrElse Not writeExpected) Then
                    Return result
                End If

                Dim operations = TryCast(observation("operations"), JArray)
                If operations IsNot Nothing Then
                    Dim failedCount = operations.OfType(Of JObject)().Count(
                        Function(item) Not String.Equals(item("status")?.ToString(), "succeeded", StringComparison.OrdinalIgnoreCase))
                    If failedCount > 0 Then
                        Dim degraded = ToolResult.Failed(
                            result.ToolId,
                            $"操作批次存在 {failedCount} 个未成功步骤",
                            data:=result.Data,
                            errorCode:=ExceptionClassifier.CodePartialApply,
                            userMessage:="部分 Office 操作未完成，正在尝试修复",
                            recoverable:=True,
                            observation:=result.Observation,
                            artifacts:=result.Artifacts)
                        ' 降级成失败时还原快照必须跟着走：部分应用的这一步恰恰最需要「撤销本轮」
                        degraded.RestoreSnapshot = result.RestoreSnapshot
                        Return degraded
                    End If
                End If

                Dim changedToken = observation("changed")
                If changedToken IsNot Nothing AndAlso
                   changedToken.Type = JTokenType.Boolean AndAlso
                   Not changedToken.Value(Of Boolean)() Then
                    Return ToolResult.Failed(
                        result.ToolId,
                        "宿主返回成功，但观察结果未检测到实际变化",
                        data:=result.Data,
                        errorCode:=ExceptionClassifier.CodeVerifyFailed,
                        userMessage:="Office 未产生预期变化，正在重新验证或修复",
                        recoverable:=True,
                        observation:=result.Observation,
                        artifacts:=result.Artifacts)
                End If
            Catch ex As Exception
                AppLogger.Warn("LoopEngine", $"Validate observation failed open: {AppLogger.Redact(ex.Message)}")
            End Try
            Return result
        End Function

        Private Function FormatResultData(data As Object) As String
            If data Is Nothing Then Return ""

            Try
                Dim text As String
                If TypeOf data Is JToken Then
                    text = DirectCast(data, JToken).ToString(Formatting.None)
                Else
                    text = JsonConvert.SerializeObject(data, Formatting.None)
                End If

                Const maxLen As Integer = 1800
                If text.Length > maxLen Then
                    Return text.Substring(0, maxLen) & "...(truncated)"
                End If
                Return text
            Catch ex As Exception
                Return ""
            End Try
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

        #End Region

    End Class

End Namespace
