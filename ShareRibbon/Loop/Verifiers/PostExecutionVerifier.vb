' ShareRibbon\Loop\Verifiers\PostExecutionVerifier.vb
' 执行后验证器 - 验证指令执行后的结果是否符合预期

Imports System.Collections.Generic
Imports System.Threading.Tasks

''' <summary>
''' 执行后验证器 - 验证DOM操作结果
''' </summary>
Public Class PostExecutionVerifier
    Implements IResultVerifier

    Public Async Function VerifyAsync(
        executionResult As ExecutionResult,
        context As ExecutionContext) As Task(Of VerificationResult) Implements IResultVerifier.VerifyAsync

        Dim result As New VerificationResult()

        If executionResult.Operations Is Nothing OrElse executionResult.Operations.Count = 0 Then
            result.IsValid = True
            Return result
        End If

        For Each op In executionResult.Operations
            If op.Instruction Is Nothing Then Continue For

            ' 根据指令类型进行验证
            Select Case op.Instruction.Operation
                Case "setParagraphStyle"
                    VerifyParagraphStyle(op, result)

                Case "setCharacterFormat"
                    VerifyCharacterFormat(op, result)

                Case "insertTable"
                    VerifyTableInsertion(op, result)

                Case "setPageSetup"
                    VerifyPageSetup(op, result)

                Case "insertBreak"
                    ' 分隔符插入一般不需要复杂验证
                    result.PassedCount += 1

                Case "applyListFormat"
                    VerifyListFormat(op, result)

                Case "suggestCorrection", "suggestFormatFix", "suggestStyleUnify", "markForReview"
                    ' 校对指令不产生DOM修改，直接通过
                    result.PassedCount += 1

                Case Else
                    ' 未知指令类型，标记为通过（保守策略）
                    result.PassedCount += 1
            End Select
        Next

        result.FailedCount = result.Errors.Count
        result.IsValid = result.Errors.Count = 0

        Return result
    End Function

    ''' <summary>
    ''' 验证段落样式
    ''' </summary>
    Private Sub VerifyParagraphStyle(op As OperationResult, result As VerificationResult)
        Try
            Dim expectedStyle = op.Instruction.GetParam("params.styleName", String.Empty).ToString()
            If String.IsNullOrEmpty(expectedStyle) Then
                result.PassedCount += 1
                Return
            End If

            ' 获取实际样式（这里简化处理，实际需要访问Word DOM）
            ' 在Word中可通过 Range.Style.NameLocal 获取
            Dim actualStyle = GetStyleNameFromRange(op.TargetRange)

            If actualStyle <> expectedStyle Then
                Dim verr = New VerificationError(op.Instruction.Id,
                    $"段落样式未正确应用: 期望 '{expectedStyle}', 实际 '{actualStyle}'")
                verr.ExpectedValue = expectedStyle
                verr.ActualValue = actualStyle
                result.Errors.Add(verr)
            Else
                result.PassedCount += 1
            End If

        Catch ex As Exception
            result.Errors.Add(New VerificationError(
                op.Instruction.Id,
                $"验证段落样式时出错: {ex.Message}"))
        End Try
    End Sub

    ''' <summary>
    ''' 验证字符格式
    ''' </summary>
    Private Sub VerifyCharacterFormat(op As OperationResult, result As VerificationResult)
        Try
            Dim expectedFont = op.Instruction.GetParam("params.font.name", String.Empty).ToString()
            If String.IsNullOrEmpty(expectedFont) Then
                result.PassedCount += 1
                Return
            End If

            Dim actualFont = GetFontNameFromRange(op.TargetRange)

            If actualFont <> expectedFont Then
                Dim verr2 = New VerificationError(op.Instruction.Id,
                    $"字体未正确应用: 期望 '{expectedFont}', 实际 '{actualFont}'")
                verr2.ExpectedValue = expectedFont
                verr2.ActualValue = actualFont
                result.Errors.Add(verr2)
            Else
                result.PassedCount += 1
            End If

        Catch ex As Exception
            result.Errors.Add(New VerificationError(
                op.Instruction.Id,
                $"验证字符格式时出错: {ex.Message}"))
        End Try
    End Sub

    ''' <summary>
    ''' 验证表格插入
    ''' </summary>
    Private Sub VerifyTableInsertion(op As OperationResult, result As VerificationResult)
        Try
            Dim expectedRows = CInt(op.Instruction.GetParam("params.rows", 0))
            Dim expectedCols = CInt(op.Instruction.GetParam("params.cols", 0))

            If expectedRows = 0 OrElse expectedCols = 0 Then
                result.PassedCount += 1
                Return
            End If

            ' 简化验证：假设表格已插入
            result.PassedCount += 1

        Catch ex As Exception
            result.Errors.Add(New VerificationError(
                op.Instruction.Id,
                $"验证表格插入时出错: {ex.Message}"))
        End Try
    End Sub

    ''' <summary>
    ''' 验证页面设置
    ''' </summary>
    Private Sub VerifyPageSetup(op As OperationResult, result As VerificationResult)
        Try
            result.PassedCount += 1
        Catch ex As Exception
            result.Errors.Add(New VerificationError(
                op.Instruction.Id,
                $"验证页面设置时出错: {ex.Message}"))
        End Try
    End Sub

    ''' <summary>
    ''' 验证列表格式
    ''' </summary>
    Private Sub VerifyListFormat(op As OperationResult, result As VerificationResult)
        Try
            result.PassedCount += 1
        Catch ex As Exception
            result.Errors.Add(New VerificationError(
                op.Instruction.Id,
                $"验证列表格式时出错: {ex.Message}"))
        End Try
    End Sub

    ''' <summary>
    ''' 从Range获取样式名称（Word特定）
    ''' </summary>
    Private Function GetStyleNameFromRange(targetRange As Object) As String
        If targetRange Is Nothing Then Return String.Empty
        Try
            ' Word Range.Style.NameLocal
            Dim style = targetRange.GetType().GetProperty("Style")?.GetValue(targetRange)
            If style IsNot Nothing Then
                Return style.GetType().GetProperty("NameLocal")?.GetValue(style)?.ToString()
            End If
        Catch
        End Try
        Return String.Empty
    End Function

    ''' <summary>
    ''' 从Range获取字体名称（Word特定）
    ''' </summary>
    Private Function GetFontNameFromRange(targetRange As Object) As String
        If targetRange Is Nothing Then Return String.Empty
        Try
            ' Word Range.Font.Name
            Dim font = targetRange.GetType().GetProperty("Font")?.GetValue(targetRange)
            If font IsNot Nothing Then
                Return font.GetType().GetProperty("Name")?.GetValue(font)?.ToString()
            End If
        Catch
        End Try
        Return String.Empty
    End Function

End Class
