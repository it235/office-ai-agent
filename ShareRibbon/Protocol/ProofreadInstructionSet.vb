' ShareRibbon\Protocol\ProofreadInstructionSet.vb
' 校对指令集 - 定义所有文档校对相关的DSL指令

Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

''' <summary>
''' 校对指令集 - 包含所有支持的文档校对操作指令
''' </summary>
Public Class ProofreadInstructionSet

    ''' <summary>
    ''' 文字修正指令 - 建议文字修正
    ''' </summary>
    Public Class SuggestCorrectionInstruction
        Inherits DslInstruction

        ''' <summary>目标文本选择器</summary>
        Public Shadows Property Target As TextTarget
        ''' <summary>修正参数</summary>
        Public Shadows Property Params As CorrectionParams

        Public Sub New()
            MyBase.New("suggestCorrection")
        End Sub
    End Class

    ''' <summary>
    ''' 格式修正建议指令
    ''' </summary>
    Public Class SuggestFormatFixInstruction
        Inherits DslInstruction

        ''' <summary>目标区域选择器</summary>
        Public Shadows Property Target As ParagraphTarget
        ''' <summary>格式参数</summary>
        Public Shadows Property Params As FormatFixParams

        Public Sub New()
            MyBase.New("suggestFormatFix")
        End Sub
    End Class

    ''' <summary>
    ''' 样式统一建议指令
    ''' </summary>
    Public Class SuggestStyleUnifyInstruction
        Inherits DslInstruction

        ''' <summary>目标文档</summary>
        Public Shadows Property Target As DocumentTarget
        ''' <summary>统一参数</summary>
        Public Shadows Property Params As StyleUnifyParams

        Public Sub New()
            MyBase.New("suggestStyleUnify")
        End Sub
    End Class

    ''' <summary>
    ''' 标记待审核指令
    ''' </summary>
    Public Class MarkForReviewInstruction
        Inherits DslInstruction

        ''' <summary>目标区域</summary>
        Public Shadows Property Target As TextTarget
        ''' <summary>标记参数</summary>
        Public Shadows Property Params As MarkForReviewParams

        Public Sub New()
            MyBase.New("markForReview")
        End Sub
    End Class

    ''' <summary>
    ''' 文字修正参数
    ''' </summary>
    Public Class CorrectionParams
        ''' <summary>原文</summary>
        Public Property Original As String
        ''' <summary>修正后文字</summary>
        Public Property Suggestion As String
        ''' <summary>问题类型</summary>
        Public Property IssueType As String
        ''' <summary>严重程度</summary>
        Public Property Severity As String
        ''' <summary>解释说明</summary>
        Public Property Explanation As String
    End Class

    ''' <summary>
    ''' 格式修正参数
    ''' </summary>
    Public Class FormatFixParams
        ''' <summary>当前格式</summary>
        Public Property CurrentFormat As String
        ''' <summary>期望格式</summary>
        Public Property ExpectedFormat As String
        ''' <summary>解释说明</summary>
        Public Property Explanation As String
    End Class

    ''' <summary>
    ''' 样式统一参数
    ''' </summary>
    Public Class StyleUnifyParams
        ''' <summary>目标样式</summary>
        Public Property TargetStyle As String
        ''' <summary>期望样式</summary>
        Public Property ExpectedStyle As String
        ''' <summary>不一致的范围</summary>
        Public Property InconsistentRanges As List(Of RangeInfo)
    End Class

    ''' <summary>
    ''' 审核标记参数
    ''' </summary>
    Public Class MarkForReviewParams
        ''' <summary>备注信息</summary>
        Public Property Note As String
        ''' <summary>分类</summary>
        Public Property Category As String
    End Class

    ''' <summary>
    ''' 文本目标选择器
    ''' </summary>
    Public Class TextTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
        ''' <summary>文本匹配模式</summary>
        Public Property Match As String
        ''' <summary>索引位置</summary>
        Public Property Index As Integer?
        ''' <summary>范围信息</summary>
        Public Property Range As RangeInfo
    End Class

    ''' <summary>
    ''' 段落目标选择器
    ''' </summary>
    Public Class ParagraphTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
        ''' <summary>语义选择器</summary>
        Public Property Selector As String
        ''' <summary>索引位置</summary>
        Public Property Index As Integer?
    End Class

    ''' <summary>
    ''' 文档目标选择器
    ''' </summary>
    Public Class DocumentTarget
        ''' <summary>选择器类型</summary>
        Public Property Type As String
    End Class

    ''' <summary>
    ''' 范围信息
    ''' </summary>
    Public Class RangeInfo
        ''' <summary>起始位置</summary>
        Public Property Start As Integer
        ''' <summary>结束位置</summary>
        Public Property [End] As Integer
        ''' <summary>文本内容</summary>
        Public Property Text As String
    End Class

    ''' <summary>
    ''' 从DSL JSON创建校对指令
    ''' </summary>
    Public Shared Function FromJson(json As JObject) As List(Of DslInstruction)
        Dim instructions As New List(Of DslInstruction)()

        If json("instructions") Is Nothing Then
            Return instructions
        End If

        For Each item In json("instructions")
            Dim instruction = CreateSingleInstruction(CType(item, JObject))
            If instruction IsNot Nothing Then
                instructions.Add(instruction)
            End If
        Next

        Return instructions
    End Function

    ''' <summary>
    ''' 创建单个指令
    ''' </summary>
    Private Shared Function CreateSingleInstruction(json As JObject) As DslInstruction
        Dim op = json("op")?.ToString()
        If String.IsNullOrEmpty(op) Then
            Return Nothing
        End If

        Select Case op.ToLower()
            Case "suggestcorrection"
                Return CreateInstruction(Of SuggestCorrectionInstruction)(json)
            Case "suggestformatfix"
                Return CreateInstruction(Of SuggestFormatFixInstruction)(json)
            Case "suggeststyleunify"
                Return CreateInstruction(Of SuggestStyleUnifyInstruction)(json)
            Case "markforreview"
                Return CreateInstruction(Of MarkForReviewInstruction)(json)
            Case Else
                Return New DslInstruction(op)
        End Select
    End Function

    ''' <summary>
    ''' 创建具体指令
    ''' </summary>
    Private Shared Function CreateInstruction(Of T As {DslInstruction, New})(json As JObject) As T
        Dim instruction = New T()

        If json("target") IsNot Nothing Then
            instruction.Target = json("target").ToObject(Of JObject)()
        End If

        If json("params") IsNot Nothing Then
            instruction.Params = json("params").ToObject(Of JObject)()
        End If

        If json("expected") IsNot Nothing Then
            instruction.Expected = json("expected").ToObject(Of JObject)()
        End If

        If json("rollback") IsNot Nothing Then
            instruction.Rollback = json("rollback").ToObject(Of JObject)()
        End If

        Return instruction
    End Function

    ''' <summary>
    ''' 创建修正指令
    ''' </summary>
    Public Shared Function CreateCorrection(
        original As String,
        suggestion As String,
        issueType As String,
        severity As String,
        explanation As String,
        Optional matchText As String = Nothing) As SuggestCorrectionInstruction

        Dim instruction As New SuggestCorrectionInstruction()
        Dim baseTarget = New JObject()
        baseTarget("type") = "textMatch"
        baseTarget("match") = If(matchText, original)
        Dim baseParams = New JObject()
        baseParams("original") = original
        baseParams("suggestion") = suggestion
        baseParams("issueType") = issueType
        baseParams("severity") = severity
        baseParams("explanation") = explanation

        DirectCast(instruction, DslInstruction).Target = baseTarget
        DirectCast(instruction, DslInstruction).Params = baseParams
        DirectCast(instruction, DslInstruction).Expected = New JObject()
        DirectCast(instruction, DslInstruction).Expected("description") = $"建议将'${original}'修正为'${suggestion}'"

        Return instruction
    End Function
End Class
