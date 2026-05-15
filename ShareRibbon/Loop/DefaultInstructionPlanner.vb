' ShareRibbon\Loop\DefaultInstructionPlanner.vb
' 默认指令规划器 - 基于用户意图生成排版规划

Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' 默认指令规划器
''' </summary>
Public Class DefaultInstructionPlanner
    Implements IInstructionPlanner

    ''' <summary>
    ''' 基于执行上下文规划指令
    ''' </summary>
    Public Async Function PlanAsync(context As ExecutionContext) As Task(Of PlanningResult) Implements IInstructionPlanner.PlanAsync
        Dim result As New PlanningResult With {.IsSuccess = True}

        Try
            ' 这里可以根据用户意图和上下文生成规划
            ' 简化实现：根据文档类型和用户请求生成基础规划
            Dim userRequest = context.UserMessage
            Dim docType = If(context.OfficeContent IsNot Nothing, context.OfficeContent.DocumentType, String.Empty)

            ' 构建规划提示词
            Dim prompt = BuildPlanningPrompt(userRequest, docType, context)

            ' 调用AI生成规划（简化处理）
            Dim planningResult = Await GeneratePlanningAsync(prompt, context)
            result.PlanData = planningResult
            result.IsSuccess = planningResult IsNot Nothing
            If Not result.IsSuccess Then
                result.ErrorMessage = "生成规划失败"
            End If

        Catch ex As Exception
            result.IsSuccess = False
            result.ErrorMessage = $"规划失败: {ex.Message}"
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 构建规划提示词
    ''' </summary>
    Private Function BuildPlanningPrompt(userRequest As String, docType As String, context As ExecutionContext) As String
        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("你是Office文档排版专家。请分析用户的排版需求，并制定详细的执行计划。")
        sb.AppendLine()
        sb.AppendLine($"【用户需求】: {userRequest}")
        sb.AppendLine()
        sb.AppendLine($"【文档类型】: {docType}")
        sb.AppendLine()

        If context.Paragraphs IsNot Nothing AndAlso context.Paragraphs.Count > 0 Then
            sb.AppendLine("【文档内容预览】:")
            For i = 0 To Math.Min(context.Paragraphs.Count - 1, 5)
                sb.AppendLine($"段落{i + 1}: {context.Paragraphs(i)}")
            Next
            If context.Paragraphs.Count > 6 Then
                sb.AppendLine($"... 还有 {context.Paragraphs.Count - 6} 个段落")
            End If
        End If

        sb.AppendLine()
        sb.AppendLine("【规划要求】:")
        sb.AppendLine("1. 确定需要执行的操作类型")
        sb.AppendLine("2. 分析目标段落和目标位置")
        sb.AppendLine("3. 评估操作的破坏性（是否会影响其他内容）")
        sb.AppendLine("4. 确定操作的执行顺序")
        sb.AppendLine("5. 估计操作的执行时间")
        sb.AppendLine()
        sb.AppendLine("【输出格式】:")
        sb.AppendLine("返回JSON格式的规划信息，包含以下字段:")
        sb.AppendLine("- operations: 操作列表")
        sb.AppendLine("- order: 执行顺序")
        sb.AppendLine("- estimatedTime: 估计时间（毫秒）")
        sb.AppendLine("- hasDestructiveOps: 是否包含破坏性操作")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成规划（模拟AI调用）
    ''' </summary>
    Private Async Function GeneratePlanningAsync(prompt As String, context As ExecutionContext) As Task(Of JObject)
        ' 这里可以实现真实的AI调用
        ' 简化实现：返回默认规划
        Await Task.Delay(500)

        Dim plan As New JObject()
        plan("operations") = New JArray() From {"analyze", "reformat", "verify"}
        plan("order") = New JArray() From {"analyze", "reformat", "verify"}
        plan("estimatedTime") = 5000
        plan("hasDestructiveOps") = False

        Return plan
    End Function
End Class
