' ShareRibbon\Loop\ExecutionResult.vb
' 指令执行结果

Imports System.Collections.Generic

''' <summary>
''' 指令执行结果
''' </summary>
Public Class ExecutionResult

    ''' <summary>是否全部成功</summary>
    Public Property IsSuccess As Boolean = True

    ''' <summary>是否有错误</summary>
    Public Property HasErrors As Boolean = False

    ''' <summary>执行错误列表</summary>
    Public Property Errors As List(Of InstructionError)

    ''' <summary>执行的单个操作结果</summary>
    Public Property Operations As List(Of OperationResult)

    ''' <summary>成功执行的指令数</summary>
    Public Property SuccessCount As Integer = 0

    ''' <summary>失败的指令数</summary>
    Public Property FailureCount As Integer = 0

    ''' <summary>修改过的Range列表（用于回滚）</summary>
    Public Property ModifiedRanges As List(Of ModifiedRangeInfo)

    ''' <summary>执行耗时（毫秒）</summary>
    Public Property ExecutionTimeMs As Long = 0

    Public Sub New()
        Operations = New List(Of OperationResult)()
        ModifiedRanges = New List(Of ModifiedRangeInfo)()
        Errors = New List(Of InstructionError)()
    End Sub

End Class

''' <summary>
''' 单个操作执行结果
''' </summary>
Public Class OperationResult

    ''' <summary>关联的指令</summary>
    Public Property Instruction As Instruction

    ''' <summary>是否成功</summary>
    Public Property IsSuccess As Boolean = True

    ''' <summary>错误消息</summary>
    Public Property ErrorMessage As String = String.Empty

    ''' <summary>目标Range（执行时确定的）</summary>
    Public Property TargetRange As Object = Nothing

    ''' <summary>执行耗时（毫秒）</summary>
    Public Property ExecutionTimeMs As Long = 0

    ''' <summary>受影响项目数</summary>
    Public Property AffectedItems As Integer = 0

    ''' <summary>附加数据</summary>
    Public Property AdditionalData As Object = Nothing

End Class
