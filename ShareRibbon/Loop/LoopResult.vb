' ShareRibbon\Loop\LoopResult.vb
' Loop执行结果模型

Imports System.Collections.Generic

''' <summary>
''' Loop执行结果
''' </summary>
Public Class LoopResult

    ''' <summary>是否成功</summary>
    Public Property IsSuccess As Boolean = False

    ''' <summary>结果类型</summary>
    Public Property ResultType As LoopResultType = LoopResultType.Unknown

    ''' <summary>最终指令列表（校验通过的）</summary>
    Public Property Instructions As List(Of Instruction) = Nothing

    ''' <summary>执行结果</summary>
    Public Property ExecutionResult As ExecutionResult = Nothing

    ''' <summary>错误列表</summary>
    Public Property Errors As List(Of InstructionError)

    ''' <summary>警告列表</summary>
    Public Property Warnings As List(Of String)

    ''' <summary>执行的迭代次数</summary>
    Public Property IterationCount As Integer = 0

    ''' <summary>用户提示消息（用于前端展示）</summary>
    Public Property UserMessage As String = String.Empty

    Public Sub New()
        Errors = New List(Of InstructionError)()
        Warnings = New List(Of String)()
    End Sub

    ''' <summary>创建成功结果</summary>
    Public Shared Function Success(executionResult As ExecutionResult) As LoopResult
        Return New LoopResult With {
            .IsSuccess = True,
            .ResultType = LoopResultType.Success,
            .ExecutionResult = executionResult
        }
    End Function

    ''' <summary>创建发送前校验失败结果</summary>
    Public Shared Function FromPreCheckFailure(checkResult As ContextCheckResult) As LoopResult
        Dim result = New LoopResult With {
            .IsSuccess = False,
            .ResultType = LoopResultType.PreCheckFailed
        }
        If checkResult.Errors IsNot Nothing Then
            For Each e In checkResult.Errors
                result.Errors.Add(New InstructionError(ErrorLevel.Critical, e))
            Next
        End If
        If checkResult.Warnings IsNot Nothing Then
            result.Warnings.AddRange(checkResult.Warnings)
        End If
        result.UserMessage = checkResult.SuggestedClarification
        Return result
    End Function

    ''' <summary>创建规划失败结果</summary>
    Public Shared Function FromPlanningFailure(planResult As PlanningResult) As LoopResult
        Dim result = New LoopResult With {
            .IsSuccess = False,
            .ResultType = LoopResultType.PlanningFailed
        }
        If planResult.ErrorMessage IsNot Nothing Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, planResult.ErrorMessage))
        End If
        Return result
    End Function

    ''' <summary>创建校验失败结果</summary>
    Public Shared Function FromValidationFailure(validationResult As ValidationResult) As LoopResult
        Dim result = New LoopResult With {
            .IsSuccess = False,
            .ResultType = LoopResultType.ValidationFailed
        }
        If validationResult.Errors IsNot Nothing Then
            result.Errors.AddRange(validationResult.Errors)
        End If
        Return result
    End Function

    ''' <summary>创建修正结果</summary>
    Public Shared Function FromCorrection(correctionResult As CorrectionResult) As LoopResult
        Dim result = New LoopResult With {
            .IsSuccess = correctionResult.IsSuccess,
            .ResultType = If(correctionResult.IsSuccess, LoopResultType.Success, LoopResultType.CorrectionFailed),
            .Instructions = correctionResult.Instructions,
            .IterationCount = correctionResult.IterationCount
        }
        If correctionResult.Errors IsNot Nothing Then
            result.Errors.AddRange(correctionResult.Errors)
        End If
        Return result
    End Function

End Class

''' <summary>
''' Loop结果类型
''' </summary>
Public Enum LoopResultType
    Unknown
    Success
    PreCheckFailed
    PlanningFailed
    ValidationFailed
    ExecutionFailed
    CorrectionFailed
    UserCancelled
End Enum

''' <summary>
''' 规划结果
''' </summary>
Public Class PlanningResult
    Public Property IsSuccess As Boolean = False
    Public Property ErrorMessage As String = String.Empty
    Public Property PlannedOperations As List(Of String) = Nothing
    Public Property PlanData As Object = Nothing
End Class
