' ShareRibbon\Loop\CorrectionResult.vb
' 修正Loop结果

Imports System.Collections.Generic

''' <summary>
''' 修正Loop结果
''' </summary>
Public Class CorrectionResult

    ''' <summary>是否成功</summary>
    Public Property IsSuccess As Boolean = False

    ''' <summary>修正后的AI响应内容</summary>
    Public Property CorrectedResponse As String = String.Empty

    ''' <summary>修正后的指令列表</summary>
    Public Property Instructions As List(Of Instruction)

    ''' <summary>迭代次数</summary>
    Public Property IterationCount As Integer = 0

    ''' <summary>最终校验结果（修正后仍失败时使用）</summary>
    Public Property FinalValidation As ValidationResult = Nothing

    ''' <summary>错误列表</summary>
    Public Property Errors As List(Of InstructionError)

    Public Sub New()
        Instructions = New List(Of Instruction)()
        Errors = New List(Of InstructionError)()
    End Sub

    Public Shared Function Success(correctedResponse As String, instructions As List(Of Instruction), iterationCount As Integer) As CorrectionResult
        Return New CorrectionResult With {
            .IsSuccess = True,
            .CorrectedResponse = correctedResponse,
            .Instructions = instructions,
            .IterationCount = iterationCount
        }
    End Function

    Public Shared Function Failure(finalValidation As ValidationResult, iterationCount As Integer) As CorrectionResult
        Return New CorrectionResult With {
            .IsSuccess = False,
            .FinalValidation = finalValidation,
            .IterationCount = iterationCount,
            .Errors = If(finalValidation?.Errors, New List(Of InstructionError)())
        }
    End Function

End Class
