Imports System.Text
Imports System.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' 本轮已经真实落到文档里的一次写入。重新规划生成的新计划并不知道文档已被改过，
    ''' 会把已成功执行的步骤重新排进来，再执行一遍就是重复写入（E1-06）。
    ''' </summary>
    Public Class AppliedWrite
        Public Property Fingerprint As String = ""
        Public Property ToolId As String = ""
        Public Property StepDescription As String = ""
        Public Property Summary As String = ""
        Public Property PlanGeneration As Integer

        Public Shared Function BuildFingerprint(toolId As String, parameters As JObject) As String
            Return If(toolId, "").Trim().ToLowerInvariant() & "|" & CanonicalJson(parameters)
        End Function

        ''' <summary>
        ''' 对象键排序后再序列化：模型重规划时同样的写入会换键顺序，语义却相同。
        ''' 数组保留原顺序，因为命令序列的先后本身就是语义。
        ''' </summary>
        Private Shared Function CanonicalJson(token As JToken) As String
            If token Is Nothing Then Return "null"

            Select Case token.Type
                Case JTokenType.Object
                    Dim parts = CType(token, JObject).Properties().
                        OrderBy(Function(p) p.Name, StringComparer.Ordinal).
                        Select(Function(p) JsonConvert.SerializeObject(p.Name) & ":" & CanonicalJson(p.Value))
                    Return "{" & String.Join(",", parts) & "}"
                Case JTokenType.Array
                    Return "[" & String.Join(",", token.Children().Select(Function(c) CanonicalJson(c))) & "]"
                Case Else
                    Return token.ToString(Formatting.None)
            End Select
        End Function
    End Class

    ''' <summary>
    ''' 单轮写入台账。只登记成功、确有变化、未被回滚的写入；读取和半成功不登记，
    ''' 否则重新规划会把"其实没写成"的步骤当成已完成而永久跳过。
    ''' </summary>
    Public Class AgentWriteLedger

        Private ReadOnly _items As New List(Of AppliedWrite)()
        Private Const PromptListLimit As Integer = 20

        ''' <summary>当前计划版本，首版为 1，每次 replan 递增。</summary>
        Public Property PlanGeneration As Integer = 1

        Public ReadOnly Property Items As List(Of AppliedWrite)
            Get
                Return _items
            End Get
        End Property

        Public Sub BeginNextPlan()
            PlanGeneration += 1
        End Sub

        Public Function Record(toolCall As ToolCall,
                               change As AgentChangeRecord,
                               stepDescription As String) As AppliedWrite
            If toolCall Is Nothing OrElse change Is Nothing Then Return Nothing
            If Not change.Success OrElse Not change.Changed OrElse change.RolledBack Then Return Nothing

            Dim entry As New AppliedWrite With {
                .Fingerprint = AppliedWrite.BuildFingerprint(toolCall.ToolId, toolCall.Parameters),
                .ToolId = If(toolCall.ToolId, ""),
                .StepDescription = If(stepDescription, ""),
                .Summary = If(change.Summary, ""),
                .PlanGeneration = PlanGeneration
            }
            _items.Add(entry)
            Return entry
        End Function

        ''' <summary>
        ''' 只匹配更早计划版本里的同款写入。同一版计划内出现两个相同步骤是用户要求的重复动作
        ''' （例如连续插入同样的页），不能被当成重规划造成的重复而跳过。
        ''' </summary>
        Public Function FindAppliedInPriorPlan(toolCall As ToolCall) As AppliedWrite
            If toolCall Is Nothing OrElse _items.Count = 0 Then Return Nothing

            Dim fingerprint = AppliedWrite.BuildFingerprint(toolCall.ToolId, toolCall.Parameters)
            Return _items.FirstOrDefault(
                Function(item) item.PlanGeneration < PlanGeneration AndAlso
                               String.Equals(item.Fingerprint, fingerprint, StringComparison.Ordinal))
        End Function

        ''' <summary>
        ''' 给模型看的已写入清单。台账为空时返回空串，不给提示词添噪声。
        ''' </summary>
        Public Function BuildAppliedWritesText() As String
            If _items.Count = 0 Then Return ""

            Dim sb As New StringBuilder()
            sb.AppendLine()
            sb.AppendLine("【本轮已真实写入，禁止重复执行】APPLIED_WRITES")
            For Each item In _items.Take(PromptListLimit)
                sb.AppendLine($"- 计划版本 {item.PlanGeneration}: {item.ToolId}｜{Compact(item.StepDescription)}｜{Compact(item.Summary)}")
            Next
            If _items.Count > PromptListLimit Then
                sb.AppendLine($"- 另有 {_items.Count - PromptListLimit} 条已写入未列出")
            End If
            sb.AppendLine("重新规划时只安排尚未完成的工作；上述写入再次出现在计划里会被执行器判为重复写入并跳过。")
            Return sb.ToString()
        End Function

        Private Shared Function Compact(value As String) As String
            Dim text = If(value, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            Const maxLength As Integer = 60
            If text.Length <= maxLength Then Return text
            Return text.Substring(0, maxLength) & "…"
        End Function
    End Class

End Namespace
