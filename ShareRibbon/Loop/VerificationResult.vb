' ShareRibbon\Loop\VerificationResult.vb
' 执行后验证结果

Imports System.Collections.Generic

''' <summary>
''' 执行后验证结果
''' </summary>
Public Class VerificationResult

    ''' <summary>是否通过验证</summary>
    Public Property IsValid As Boolean = True

    ''' <summary>验证错误列表</summary>
    Public Property Errors As List(Of VerificationError)

    ''' <summary>修改过的Range列表（用于回滚）</summary>
    Public Property ModifiedRanges As List(Of ModifiedRangeInfo)

    ''' <summary>验证通过的指令数</summary>
    Public Property PassedCount As Integer = 0

    ''' <summary>验证失败的指令数</summary>
    Public Property FailedCount As Integer = 0

    Public Sub New()
        Errors = New List(Of VerificationError)()
        ModifiedRanges = New List(Of ModifiedRangeInfo)()
    End Sub

End Class

''' <summary>
''' 验证错误
''' </summary>
Public Class VerificationError

    ''' <summary>关联的指令ID</summary>
    Public Property InstructionId As String

    ''' <summary>错误描述</summary>
    Public Property Message As String

    ''' <summary>期望值</summary>
    Public Property ExpectedValue As String

    ''' <summary>实际值</summary>
    Public Property ActualValue As String

    Public Sub New(instructionId As String, message As String)
        Me.InstructionId = instructionId
        Me.Message = message
        Me.ExpectedValue = String.Empty
        Me.ActualValue = String.Empty
    End Sub

End Class

''' <summary>
''' 修改过的Range信息（用于回滚和验证）</summary>
Public Class ModifiedRangeInfo

    ''' <summary>关联的指令ID</summary>
    Public Property InstructionId As String

    ''' <summary>Range起始位置</summary>
    Public Property Start As Integer

    ''' <summary>Range结束位置</summary>
    Public Property [End] As Integer

    ''' <summary>修改前的文本/样式快照</summary>
    Public Property OriginalSnapshot As String

    ''' <summary>段落索引</summary>
    Public Property ParagraphIndex As Integer = -1

End Class
