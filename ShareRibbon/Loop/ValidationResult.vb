' ShareRibbon\Loop\ValidationResult.vb
' Flush后校验结果

Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

''' <summary>
''' Flush后校验结果
''' </summary>
Public Class ValidationResult

    ''' <summary>是否通过校验</summary>
    Public Property IsValid As Boolean = True

    ''' <summary>错误列表</summary>
    Public Property Errors As List(Of InstructionError)

    ''' <summary>警告列表</summary>
    Public Property Warnings As List(Of InstructionError)

    ''' <summary>从AI响应中提取的原始内容</summary>
    Public Property ExtractedContent As String = String.Empty

    ''' <summary>解析后的指令列表</summary>
    Public Property ParsedInstructions As List(Of Instruction)

    ''' <summary>是否可自动修正</summary>
    Public Property CanAutoCorrect As Boolean = False

    ''' <summary>AI原始响应用于修正</summary>
    Public Property OriginalResponse As String = String.Empty

    Public Sub New()
        Errors = New List(Of InstructionError)()
        Warnings = New List(Of InstructionError)()
        ParsedInstructions = New List(Of Instruction)()
    End Sub

    ''' <summary>创建失败结果</summary>
    Public Shared Function Failure(errors As List(Of InstructionError)) As ValidationResult
        Return New ValidationResult With {
            .IsValid = False,
            .Errors = errors
        }
    End Function

    ''' <summary>创建失败结果（带提取内容）</summary>
    Public Shared Function Failure(errors As List(Of InstructionError), extractedContent As String) As ValidationResult
        Return New ValidationResult With {
            .IsValid = False,
            .Errors = errors,
            .ExtractedContent = extractedContent
        }
    End Function

End Class

''' <summary>
''' 解析结果
''' </summary>
Public Class ParseResult

    ''' <summary>是否解析成功</summary>
    Public Property IsValid As Boolean = True

    ''' <summary>解析出的指令列表</summary>
    Public Property Instructions As List(Of Instruction)

    ''' <summary>解析错误列表</summary>
    Public Property Errors As List(Of InstructionError)

    Public Sub New()
        Instructions = New List(Of Instruction)()
        Errors = New List(Of InstructionError)()
    End Sub

    Public Shared Function Success(instructions As List(Of Instruction)) As ParseResult
        Return New ParseResult With {
            .IsValid = True,
            .Instructions = instructions
        }
    End Function

    Public Shared Function Failure([error] As InstructionError) As ParseResult
        Dim result = New ParseResult With {.IsValid = False}
        result.Errors.Add([error])
        Return result
    End Function

End Class

''' <summary>
''' 参数校验结果
''' </summary>
Public Class ParamValidationResult

    ''' <summary>是否通过</summary>
    Public Property IsValid As Boolean = True

    ''' <summary>错误消息</summary>
    Public Property ErrorMessage As String = String.Empty

    Public Shared Function Success() As ParamValidationResult
        Return New ParamValidationResult With {.IsValid = True}
    End Function

    Public Shared Function Failure(message As String) As ParamValidationResult
        Return New ParamValidationResult With {
            .IsValid = False,
            .ErrorMessage = message
        }
    End Function

End Class
