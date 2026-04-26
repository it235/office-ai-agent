Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 增强版 Ralph Loop 控制器 - 智能规划、执行、循环流程
''' </summary>
Public Class RalphLoopController
    Private ReadOnly _memory As RalphLoopMemory
    Private _currentChatControl As BaseChatControl

    ' 增强版规划提示词模板 - 更智能，支持记忆上下文
    Private Const ENHANCED_PLANNING_PROMPT As String = "你是一个智能任务规划专家。深度分析用户目标、相关记忆和Office上下文，制定详细、可执行的计划。

【输入信息】
- 用户目标：{0}
- 当前应用：{1}
{2}

【规划原则】
1. **目标分解**：将复杂目标分解为可独立执行的步骤
2. **智能步骤**：每个步骤应该具体、可验证、可回滚
3. **上下文感知**：结合选中区域、当前工作表、相关记忆来规划
4. **风险评估**：对每个步骤评估风险级别（safe/medium/risky）
5. **依赖关系**：标注步骤之间的依赖关系

【返回JSON格式】
```json
{
  ""goal_summary"": ""用户目标的简短总结"",
  ""estimated_complexity"": ""low/medium/high"",
  ""risk_level"": ""safe/medium/risky"",
  ""steps"": [
    {
      ""step_number"": 1,
      ""description"": ""步骤的详细描述"",
      ""intent"": ""意图类型(data_query/format/chart/formula/clean/transform)"",
      ""risk_level"": ""safe/medium/risky"",
      ""estimated_time"": ""预估时间(如10秒)"",
      ""depends_on"": [],
      ""rollback_hint"": ""回滚建议""
    }
  ],
  ""success_criteria"": ""任务完成的判断标准"",
  ""notes"": ""注意事项和提示""
}
```

【步骤数量建议】
- low复杂度：1-2步
- medium复杂度：3-4步
- high复杂度：5-8步

【示例】
用户目标：""计算销售总额并生成图表""
返回步骤：
1. 识别销售数据范围
2. 计算销售总额
3. 创建柱状图展示数据
4. 调整图表样式"

    ' 兼容旧版的规划提示词模板
    Private Const PLANNING_PROMPT As String = "你是一个任务规划专家。用户有一个目标需要完成，请分析这个目标并制定执行计划。

用户目标：{0}

请以JSON格式返回执行计划，格式如下：
{{
  ""goal_summary"": ""目标简述"",
  ""steps"": [
    {{
      ""step_number"": 1,
      ""description"": ""步骤描述"",
      ""intent"": ""意图类型(如: data_query, format, chart, formula等)""
    }}
  ],
  ""estimated_complexity"": ""low/medium/high""
}}

注意：
1. 步骤数量根据任务复杂度决定，简单任务1-2步，复杂任务3-5步
2. 每个步骤应该是可独立执行的操作
3. 只返回JSON，不要其他解释"

    Public Sub New()
        _memory = RalphLoopMemory.Instance
    End Sub

    ''' <summary>
    ''' 设置当前聊天控件引用
    ''' </summary>
    Public Sub SetChatControl(chatControl As BaseChatControl)
        _currentChatControl = chatControl
    End Sub

    ''' <summary>
    ''' 开始新的循环 - 规划阶段
    ''' </summary>
    Public Function StartNewLoop(userGoal As String, applicationType As String) As Task(Of RalphLoopSession)
        Dim session As New RalphLoopSession() With {
            .OriginalGoal = userGoal,
            .ApplicationType = applicationType,
            .Status = RalphLoopStatus.Planning
        }
        _memory.SetActiveLoop(session)
        Return Task.FromResult(session)
    End Function

    ''' <summary>
    ''' 解析规划结果（增强版，支持更多字段）
    ''' </summary>
    Public Function ParsePlanningResult(jsonResult As String) As Boolean
        Try
            Dim loopSession = _memory.GetActiveLoop()
            If loopSession Is Nothing Then Return False

            ' 尝试提取JSON
            Dim jsonStart = jsonResult.IndexOf("{")
            Dim jsonEnd = jsonResult.LastIndexOf("}")
            If jsonStart >= 0 AndAlso jsonEnd > jsonStart Then
                jsonResult = jsonResult.Substring(jsonStart, jsonEnd - jsonStart + 1)
            End If

            Dim planObj = JObject.Parse(jsonResult)
            Dim stepsArray = planObj("steps")

            If stepsArray IsNot Nothing Then
                loopSession.Steps.Clear()
                Dim stepNum = 1
                For Each stepItem In stepsArray
                    Dim loopStep As New RalphLoopStep() With {
                        .StepNumber = stepNum,
                        .Description = stepItem("description")?.ToString(),
                        .Intent = stepItem("intent")?.ToString(),
                        .Status = RalphStepStatus.Pending
                    }

                    ' 解析增强字段
                    If stepItem("risk_level") IsNot Nothing Then
                        loopStep.RiskLevel = stepItem("risk_level").ToString()
                    End If
                    If stepItem("estimated_time") IsNot Nothing Then
                        loopStep.EstimatedTime = stepItem("estimated_time").ToString()
                    End If
                    If stepItem("rollback_hint") IsNot Nothing Then
                        loopStep.RollbackHint = stepItem("rollback_hint").ToString()
                    End If
                    If stepItem("depends_on") IsNot Nothing AndAlso stepItem("depends_on").Type = JTokenType.Array Then
                        For Each dep In stepItem("depends_on")
                            loopStep.DependsOn.Add(dep.ToObject(Of Integer))
                        Next
                    End If

                    loopSession.Steps.Add(loopStep)
                    stepNum += 1
                Next
                loopSession.TotalSteps = loopSession.Steps.Count
                loopSession.Status = RalphLoopStatus.Ready
                _memory.Save()
                Return True
            End If
        Catch ex As Exception
            Debug.WriteLine($"[RalphLoop] 解析规划结果失败: {ex.Message}")
        End Try
        Return False
    End Function

    ''' <summary>
    ''' 执行下一步（支持依赖检查）
    ''' </summary>
    Public Function ExecuteNextStep() As RalphLoopStep
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing OrElse loopSession.Status = RalphLoopStatus.Completed Then
            Return Nothing
        End If

        ' 找到下一个待执行且依赖已满足的步骤
        Dim nextStep = FindNextRunnableStep(loopSession)
        If nextStep Is Nothing Then
            ' 检查是否所有步骤都完成了
            Dim allCompleted = loopSession.Steps.All(Function(s) s.Status = RalphStepStatus.Completed)
            If allCompleted Then
                loopSession.Status = RalphLoopStatus.Completed
                _memory.Save()
            Else
                ' 可能有步骤失败了
                Dim hasFailed = loopSession.Steps.Any(Function(s) s.Status = RalphStepStatus.Failed)
                If hasFailed Then
                    loopSession.Status = RalphLoopStatus.Paused
                End If
            End If
            Return Nothing
        End If

        ' 标记为执行中
        nextStep.Status = RalphStepStatus.Running
        nextStep.ExecutedAt = DateTime.Now
        loopSession.CurrentStep = nextStep.StepNumber
        loopSession.Status = RalphLoopStatus.Running
        _memory.Save()

        Return nextStep
    End Function

    ''' <summary>
    ''' 查找下一个可执行的步骤（检查依赖关系）
    ''' </summary>
    Private Function FindNextRunnableStep(session As RalphLoopSession) As RalphLoopStep
        For Each loopStep In session.Steps.OrderBy(Function(s) s.StepNumber)
            If loopStep.Status = RalphStepStatus.Pending Then
                ' 检查依赖
                If loopStep.DependsOn Is Nothing OrElse loopStep.DependsOn.Count = 0 Then
                    Return loopStep
                End If

                ' 检查所有依赖步骤是否已完成
                Dim allDepsCompleted = loopStep.DependsOn.All(Function(depNo)
                                                              Dim depStep = session.Steps.FirstOrDefault(Function(s) s.StepNumber = depNo)
                                                              Return depStep IsNot Nothing AndAlso depStep.Status = RalphStepStatus.Completed
                                                          End Function)
                If allDepsCompleted Then
                    Return loopStep
                End If
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' 获取可以并行执行的步骤
    ''' </summary>
    Public Function GetParallelSteps() As List(Of RalphLoopStep)
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then
            Return New List(Of RalphLoopStep)()
        End If

        Dim parallelSteps As New List(Of RalphLoopStep)()

        ' 找到所有没有依赖且待执行的步骤
        For Each loopStep In loopSession.Steps
            If loopStep.Status = RalphStepStatus.Pending Then
                If loopStep.DependsOn Is Nothing OrElse loopStep.DependsOn.Count = 0 Then
                    parallelSteps.Add(loopStep)
                Else
                    ' 检查依赖是否都完成
                    Dim allDepsCompleted = loopStep.DependsOn.All(Function(depNo)
                                                                  Dim depStep = loopSession.Steps.FirstOrDefault(Function(s) s.StepNumber = depNo)
                                                                  Return depStep IsNot Nothing AndAlso depStep.Status = RalphStepStatus.Completed
                                                              End Function)
                    If allDepsCompleted Then
                        parallelSteps.Add(loopStep)
                    End If
                End If
            End If
        Next

        Return parallelSteps
    End Function

    ''' <summary>
    ''' 标记当前步骤完成（增强版，支持错误处理和重试）
    ''' </summary>
    Public Function CompleteCurrentStep(result As String, success As Boolean, Optional errorMessage As String = "") As Boolean
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then Return False

        Dim currentStep = loopSession.Steps.FirstOrDefault(Function(s) s.Status = RalphStepStatus.Running)
        If currentStep IsNot Nothing Then
            If success Then
                currentStep.Status = RalphStepStatus.Completed
                currentStep.Result = result
                currentStep.CompletedAt = DateTime.Now
                currentStep.ErrorMessage = ""
            Else
                currentStep.RetryCount += 1
                currentStep.ErrorMessage = errorMessage

                ' 检查是否可以重试
                If currentStep.RetryCount < currentStep.MaxRetries Then
                    currentStep.Status = RalphStepStatus.Pending
                    Debug.WriteLine($"[RalphLoop] 步骤 {currentStep.StepNumber} 失败，准备第 {currentStep.RetryCount} 次重试")
                Else
                    currentStep.Status = RalphStepStatus.Failed
                    Debug.WriteLine($"[RalphLoop] 步骤 {currentStep.StepNumber} 失败，已达到最大重试次数")
                End If
            End If

            ' 检查是否还有待执行步骤
            Dim hasMoreSteps = loopSession.Steps.Any(Function(s) s.Status = RalphStepStatus.Pending)
            Dim hasFailedSteps = loopSession.Steps.Any(Function(s) s.Status = RalphStepStatus.Failed)

            If hasFailedSteps Then
                loopSession.Status = RalphLoopStatus.Paused
            ElseIf hasMoreSteps Then
                loopSession.Status = RalphLoopStatus.Paused ' 暂停等待用户确认继续
            Else
                loopSession.Status = RalphLoopStatus.Completed
                ' 记录到任务历史
                _memory.AddTaskRecord(New RalphTaskRecord() With {
                    .UserInput = loopSession.OriginalGoal,
                    .Intent = "multi_step_task",
                    .Plan = String.Join(" -> ", loopSession.Steps.Select(Function(s) s.Description)),
                    .Result = result,
                    .Success = True,
                    .ApplicationType = loopSession.ApplicationType
                })
            End If
            _memory.Save()

            Return success OrElse currentStep.Status = RalphStepStatus.Pending
        End If

        Return False
    End Function

    ''' <summary>
    ''' 回滚到指定步骤
    ''' </summary>
    Public Function RollbackToStep(stepNumber As Integer) As Boolean
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then Return False

        loopSession.Status = RalphLoopStatus.RollingBack

        Try
            ' 回滚该步骤之后的所有步骤
            For Each loopStep In loopSession.Steps.Where(Function(s) s.StepNumber >= stepNumber).OrderByDescending(Function(s) s.StepNumber)
                If loopStep.Status = RalphStepStatus.Completed Then
                    loopStep.Status = RalphStepStatus.RolledBack
                    Debug.WriteLine($"[RalphLoop] 已回滚步骤 {loopStep.StepNumber}")
                End If
            Next

            loopSession.Status = RalphLoopStatus.Paused
            _memory.Save()
            Return True
        Catch ex As Exception
            Debug.WriteLine($"[RalphLoop] 回滚失败: {ex.Message}")
            loopSession.Status = RalphLoopStatus.Paused
            _memory.Save()
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 重试失败的步骤
    ''' </summary>
    Public Function RetryFailedStep(stepNumber As Integer) As RalphLoopStep
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then Return Nothing

        Dim loopStep = loopSession.Steps.FirstOrDefault(Function(s) s.StepNumber = stepNumber AndAlso s.Status = RalphStepStatus.Failed)
        If loopStep IsNot Nothing Then
            loopStep.Status = RalphStepStatus.Pending
            loopStep.ErrorMessage = ""
            _memory.Save()
            Return loopStep
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' 跳过某个步骤
    ''' </summary>
    Public Sub SkipStep(stepNumber As Integer)
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then Return

        Dim loopStep = loopSession.Steps.FirstOrDefault(Function(s) s.StepNumber = stepNumber)
        If loopStep IsNot Nothing Then
            loopStep.Status = RalphStepStatus.Skipped
            _memory.Save()
        End If
    End Sub

    ''' <summary>
    ''' 获取当前循环状态摘要（增强版）
    ''' </summary>
    Public Function GetLoopStatusSummary() As String
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then
            Return "当前没有活动的循环任务"
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine($"目标：{loopSession.OriginalGoal}")
        sb.AppendLine($"状态：{GetStatusText(loopSession.Status)}")

        Dim completedCount = loopSession.Steps.Where(Function(s) s.Status = RalphStepStatus.Completed).Count
        sb.AppendLine($"进度：{completedCount}/{loopSession.TotalSteps}")
        sb.AppendLine("步骤：")

        For Each loopStep In loopSession.Steps
            Dim statusIcon = GetStepStatusIcon(loopStep.Status)
            Dim riskInfo = If(loopStep.RiskLevel = "medium", " ⚠️", If(loopStep.RiskLevel = "risky", " 🚨", ""))
            sb.AppendLine($"  {statusIcon} {loopStep.StepNumber}. {loopStep.Description}{riskInfo}")

            If loopStep.Status = RalphStepStatus.Failed AndAlso Not String.IsNullOrWhiteSpace(loopStep.ErrorMessage) Then
                sb.AppendLine($"      ❌ 错误：{loopStep.ErrorMessage}")
            End If

            If Not String.IsNullOrWhiteSpace(loopStep.RollbackHint) AndAlso loopStep.Status = RalphStepStatus.Completed Then
                sb.AppendLine($"      💡 回滚提示：{loopStep.RollbackHint}")
            End If

            If loopStep.DependsOn IsNot Nothing AndAlso loopStep.DependsOn.Count > 0 Then
                sb.AppendLine($"      🔗 依赖：{String.Join(", ", loopStep.DependsOn)}")
            End If
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 清空上下文并结束循环
    ''' </summary>
    Public Sub ClearAndEndLoop()
        _memory.ClearActiveLoop()
    End Sub

    ''' <summary>
    ''' 保存当前会话（断点续传）
    ''' </summary>
    Public Sub SaveCurrentSession()
        _memory.SaveSession()
    End Sub

    ''' <summary>
    ''' 获取增强版规划提示词 - 支持记忆上下文和意图信息
    ''' </summary>
    Public Function GetEnhancedPlanningPrompt(userGoal As String, applicationType As String, Optional intent As IntentResult = Nothing, Optional memoryContext As String = Nothing, Optional officeContext As String = Nothing) As String
        Dim sb As New StringBuilder()

        ' 添加相似历史任务
        Dim similarTasks = _memory.FindSimilarTasks(userGoal, applicationType, 3)
        If similarTasks.Count > 0 Then
            sb.AppendLine("【历史参考】")
            sb.AppendLine("以下是之前完成的类似任务，供参考：")
            For Each task In similarTasks
                sb.AppendLine($"- {task.Timestamp:yyyy-MM-dd}: {task.UserInput}")
                If task.Success Then
                    sb.AppendLine($"  ✅ 成功：{Left(task.Result, 50)}...")
                Else
                    sb.AppendLine($"  ❌ 失败")
                End If
            Next
            sb.AppendLine()
        End If

        ' 添加相关知识
        Dim knowledge = _memory.GetRelevantKnowledge(userGoal)
        If knowledge.Count > 0 Then
            sb.AppendLine("【相关知识】")
            For Each k In knowledge
                sb.AppendLine($"- {k}")
            Next
            sb.AppendLine()
        End If

        Dim memoryInfo As String = ""
        If Not String.IsNullOrWhiteSpace(memoryContext) Then
            memoryInfo = "- 相关记忆：" & vbCrLf & memoryContext
        End If

        Dim officeInfo As String = ""
        If Not String.IsNullOrWhiteSpace(officeContext) Then
            officeInfo = "- Office上下文：" & vbCrLf & officeContext
        End If

        Dim basePrompt = String.Format(ENHANCED_PLANNING_PROMPT, userGoal, applicationType, sb.ToString() & memoryInfo & vbCrLf & officeInfo)

        If intent IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(intent.UserFriendlyDescription) Then
            basePrompt &= vbCrLf & vbCrLf & "【已识别意图】" & intent.UserFriendlyDescription & "（类型：" & intent.IntentType.ToString() & "）。请基于此制定执行计划。"
        End If

        Return basePrompt
    End Function

    ''' <summary>
    ''' 获取规划提示词（可选带入意图识别结果，供阶段四统一智能体使用）
    ''' </summary>
    Public Function GetPlanningPrompt(userGoal As String, Optional intent As IntentResult = Nothing) As String
        Dim basePrompt = String.Format(PLANNING_PROMPT, userGoal)
        If intent IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(intent.UserFriendlyDescription) Then
            basePrompt &= vbCrLf & vbCrLf & "已识别意图：" & intent.UserFriendlyDescription & "（类型：" & intent.IntentType.ToString() & "）。请基于此制定执行计划。"
        End If
        Return basePrompt
    End Function

    ''' <summary>
    ''' 检查是否有活动循环
    ''' </summary>
    Public Function HasActiveLoop() As Boolean
        Return _memory.GetActiveLoop() IsNot Nothing
    End Function

    ''' <summary>
    ''' 获取活动循环
    ''' </summary>
    Public Function GetActiveLoop() As RalphLoopSession
        Return _memory.GetActiveLoop()
    End Function

    ''' <summary>
    ''' 获取历史任务记录
    ''' </summary>
    Public Function GetTaskHistory(Optional appType As String = "", Optional maxCount As Integer = 10) As List(Of RalphTaskRecord)
        Dim tasks = _memory.MemoryData.TaskHistory.AsEnumerable()
        If Not String.IsNullOrWhiteSpace(appType) Then
            tasks = tasks.Where(Function(t) t.ApplicationType = appType)
        End If
        Return tasks.OrderByDescending(Function(t) t.Timestamp).Take(maxCount).ToList()
    End Function

    Private Function GetStatusText(status As RalphLoopStatus) As String
        Select Case status
            Case RalphLoopStatus.Planning : Return "规划中"
            Case RalphLoopStatus.Ready : Return "准备执行"
            Case RalphLoopStatus.Running : Return "执行中"
            Case RalphLoopStatus.Paused : Return "等待继续"
            Case RalphLoopStatus.Completed : Return "已完成"
            Case RalphLoopStatus.Failed : Return "失败"
            Case RalphLoopStatus.RollingBack : Return "回滚中"
            Case Else : Return "未知"
        End Select
    End Function

    Private Function GetStepStatusIcon(status As RalphStepStatus) As String
        Select Case status
            Case RalphStepStatus.Pending : Return "⏳"
            Case RalphStepStatus.Running : Return "▶️"
            Case RalphStepStatus.Completed : Return "✅"
            Case RalphStepStatus.Failed : Return "❌"
            Case RalphStepStatus.Skipped : Return "⏭️"
            Case RalphStepStatus.RolledBack : Return "↩️"
            Case Else : Return "❓"
        End Select
    End Function
End Class
