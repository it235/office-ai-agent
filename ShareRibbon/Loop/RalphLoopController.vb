Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Ralph Loop 控制器 - 管理规划、执行、循环流程
''' </summary>
Public Class RalphLoopController
    Private ReadOnly _memory As RalphLoopMemory
    Private _currentChatControl As BaseChatControl

    ' 规划提示词模板
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
    Public Async Function StartNewLoop(userGoal As String, applicationType As String) As Task(Of RalphLoopSession)
        ' 创建新的循环会话
        Dim loopSession As New RalphLoopSession() With {
            .OriginalGoal = userGoal,
            .ApplicationType = applicationType,
            .Status = RalphLoopStatus.Planning
        }

        ' 保存到记忆
        _memory.SetActiveLoop(loopSession)

        Return loopSession
    End Function

    ''' <summary>
    ''' 解析规划结果
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
                    loopSession.Steps.Add(New RalphLoopStep() With {
                        .StepNumber = stepNum,
                        .Description = stepItem("description")?.ToString(),
                        .Intent = stepItem("intent")?.ToString(),
                        .Status = RalphStepStatus.Pending
                    })
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
    ''' 执行下一步
    ''' </summary>
    Public Function ExecuteNextStep() As RalphLoopStep
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing OrElse loopSession.Status = RalphLoopStatus.Completed Then
            Return Nothing
        End If

        ' 找到下一个待执行的步骤
        Dim nextStep = loopSession.Steps.FirstOrDefault(Function(s) s.Status = RalphStepStatus.Pending)
        If nextStep Is Nothing Then
            ' 所有步骤已完成
            loopSession.Status = RalphLoopStatus.Completed
            _memory.Save()
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
    ''' 标记当前步骤完成
    ''' </summary>
    Public Sub CompleteCurrentStep(result As String, success As Boolean)
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then Return

        Dim currentStep = loopSession.Steps.FirstOrDefault(Function(s) s.Status = RalphStepStatus.Running)
        If currentStep IsNot Nothing Then
            currentStep.Status = If(success, RalphStepStatus.Completed, RalphStepStatus.Failed)
            currentStep.Result = result

            ' 检查是否还有待执行步骤
            Dim hasMoreSteps = loopSession.Steps.Any(Function(s) s.Status = RalphStepStatus.Pending)
            If hasMoreSteps Then
                loopSession.Status = RalphLoopStatus.Paused ' 暂停等待用户确认继续
            Else
                loopSession.Status = RalphLoopStatus.Completed
                ' 记录到任务历史
                _memory.AddTaskRecord(New RalphTaskRecord() With {
                    .UserInput = loopSession.OriginalGoal,
                    .Intent = "multi_step_task",
                    .Plan = String.Join(" -> ", loopSession.Steps.Select(Function(s) s.Description)),
                    .Result = result,
                    .Success = success,
                    .ApplicationType = loopSession.ApplicationType
                })
            End If
            _memory.Save()
        End If
    End Sub

    ''' <summary>
    ''' 获取当前循环状态摘要
    ''' </summary>
    Public Function GetLoopStatusSummary() As String
        Dim loopSession = _memory.GetActiveLoop()
        If loopSession Is Nothing Then
            Return "当前没有活动的循环任务"
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine($"目标: {loopSession.OriginalGoal}")
        sb.AppendLine($"状态: {GetStatusText(loopSession.Status)}")
        sb.AppendLine($"进度: {loopSession.CurrentStep}/{loopSession.TotalSteps}")
        sb.AppendLine("步骤:")
        
        For Each loopStep In loopSession.Steps
            Dim statusIcon = GetStepStatusIcon(loopStep.Status)
            sb.AppendLine($"  {statusIcon} {loopStep.StepNumber}. {loopStep.Description}")
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
    ''' 获取规划提示词
    ''' </summary>
    Public Function GetPlanningPrompt(userGoal As String) As String
        Return String.Format(PLANNING_PROMPT, userGoal)
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

    Private Function GetStatusText(status As RalphLoopStatus) As String
        Select Case status
            Case RalphLoopStatus.Planning : Return "规划中"
            Case RalphLoopStatus.Ready : Return "准备执行"
            Case RalphLoopStatus.Running : Return "执行中"
            Case RalphLoopStatus.Paused : Return "等待继续"
            Case RalphLoopStatus.Completed : Return "已完成"
            Case RalphLoopStatus.Failed : Return "失败"
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
            Case Else : Return "❓"
        End Select
    End Function
End Class
