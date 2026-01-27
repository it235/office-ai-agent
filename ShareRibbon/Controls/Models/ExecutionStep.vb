' ShareRibbon\Controls\Models\ExecutionStep.vb
' 执行步骤模型：用于意图预览和执行计划展示

''' <summary>
''' 执行步骤模型
''' </summary>
Public Class ExecutionStep
    ''' <summary>
    ''' 步骤编号
    ''' </summary>
    Public Property StepNumber As Integer

    ''' <summary>
    ''' 步骤描述（用户友好的描述）
    ''' </summary>
    Public Property Description As String

    ''' <summary>
    ''' 图标类型（formula/chart/data/format/search/clean）
    ''' </summary>
    Public Property Icon As String

    ''' <summary>
    ''' 将修改的范围或内容
    ''' </summary>
    Public Property WillModify As String

    ''' <summary>
    ''' 预计耗时描述
    ''' </summary>
    Public Property EstimatedTime As String

    ''' <summary>
    ''' 需要的数据/依赖
    ''' </summary>
    Public Property RequiresData As List(Of String)

    Public Sub New()
        RequiresData = New List(Of String)()
        Icon = "default"
        EstimatedTime = "1秒"
    End Sub

    Public Sub New(stepNum As Integer, desc As String, Optional iconType As String = "default")
        Me.New()
        StepNumber = stepNum
        Description = desc
        Icon = iconType
    End Sub
End Class

''' <summary>
''' 意图澄清结果
''' </summary>
Public Class IntentClarification
    ''' <summary>
    ''' 用户友好的意图描述
    ''' </summary>
    Public Property Description As String

    ''' <summary>
    ''' 执行计划步骤列表
    ''' </summary>
    Public Property ExecutionPlan As List(Of ExecutionStep)

    ''' <summary>
    ''' 是否需要用户确认
    ''' </summary>
    Public Property RequiresConfirmation As Boolean

    ''' <summary>
    ''' 澄清问题（如果信息不足）
    ''' </summary>
    Public Property ClarifyingQuestions As List(Of String)

    ''' <summary>
    ''' 原始用户输入
    ''' </summary>
    Public Property OriginalInput As String

    Public Sub New()
        ExecutionPlan = New List(Of ExecutionStep)()
        ClarifyingQuestions = New List(Of String)()
        RequiresConfirmation = True
    End Sub
End Class

''' <summary>
''' JSON命令预览结果
''' </summary>
Public Class JsonPreviewResult
    ''' <summary>
    ''' 执行计划步骤
    ''' </summary>
    Public Property ExecutionPlan As List(Of ExecutionStep)

    ''' <summary>
    ''' 单元格变更列表
    ''' </summary>
    Public Property CellChanges As List(Of CellChange)

    ''' <summary>
    ''' 变更摘要
    ''' </summary>
    Public Property Summary As String

    ''' <summary>
    ''' 原始JSON命令
    ''' </summary>
    Public Property OriginalJson As String

    Public Sub New()
        ExecutionPlan = New List(Of ExecutionStep)()
        CellChanges = New List(Of CellChange)()
    End Sub
End Class

''' <summary>
''' 单元格变更
''' </summary>
Public Class CellChange
    ''' <summary>
    ''' 单元格地址
    ''' </summary>
    Public Property Address As String

    ''' <summary>
    ''' 原始值
    ''' </summary>
    Public Property OldValue As Object

    ''' <summary>
    ''' 新值
    ''' </summary>
    Public Property NewValue As Object

    ''' <summary>
    ''' 变更类型（Added/Modified/Deleted）
    ''' </summary>
    Public Property ChangeType As String

    Public Sub New()
        ChangeType = "Modified"
    End Sub

    Public Sub New(addr As String, oldVal As Object, newVal As Object)
        Me.New()
        Address = addr
        OldValue = oldVal
        NewValue = newVal
    End Sub
End Class
