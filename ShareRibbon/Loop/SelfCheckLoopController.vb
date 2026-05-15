' ShareRibbon\Loop\SelfCheckLoopController.vb
' 自检Loop控制器 - 编排整个规划-生成-校验-修正-执行-验证流程

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' 自检Loop控制器 - 编排整个自检流程
''' </summary>
Public Class SelfCheckLoopController

    ' 循环配置
    Public Property MaxPlanningIterations As Integer = 3
    Public Property MaxExecutionIterations As Integer = 3
    Public Property MaxCorrectionIterations As Integer = 2

    ' 各阶段组件
    Private ReadOnly _contextChecker As IContextChecker
    Private ReadOnly _instructionValidator As IInstructionValidator
    Private ReadOnly _instructionExecutor As IInstructionExecutor
    Private ReadOnly _resultVerifier As IResultVerifier

    ' 可选组件
    Public Property InstructionPlanner As IInstructionPlanner = Nothing
    Public Property InstructionGenerator As IInstructionGenerator = Nothing

    ' 状态
    Private _currentPhase As LoopPhase = LoopPhase.Idle
    Private _iterationCount As Integer = 0

    ''' <summary>
    ''' 当前阶段（用于外部监控）
    ''' </summary>
    Public ReadOnly Property CurrentPhase As LoopPhase
        Get
            Return _currentPhase
        End Get
    End Property

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New(
        contextChecker As IContextChecker,
        instructionValidator As IInstructionValidator,
        instructionExecutor As IInstructionExecutor,
        resultVerifier As IResultVerifier)

        _contextChecker = contextChecker
        _instructionValidator = instructionValidator
        _instructionExecutor = instructionExecutor
        _resultVerifier = resultVerifier
    End Sub

    ''' <summary>
    ''' 执行发送前自检
    ''' </summary>
    Public Async Function PreSendCheckAsync(context As ExecutionContext) As Task(Of ContextCheckResult)
        _currentPhase = LoopPhase.PreSendCheck
        Return Await _contextChecker.CheckAsync(context)
    End Function

    ''' <summary>
    ''' 执行完整的自检Loop流程
    ''' </summary>
    Public Async Function ExecuteLoopAsync(context As ExecutionContext) As Task(Of LoopResult)
        Try
            _currentPhase = LoopPhase.Planning
            _iterationCount = 0

            ' 1. 规划阶段（如果提供了规划器）
            Dim plan As PlanningResult = Nothing
            If InstructionPlanner IsNot Nothing Then
                plan = Await RunPlanningLoopAsync(context)
                If Not plan.IsSuccess Then
                    _currentPhase = LoopPhase.Failed
                    Return LoopResult.FromPlanningFailure(plan)
                End If
            End If

            ' 2. 生成+校验阶段
            _currentPhase = LoopPhase.Generating
            Dim validatedInstructions = Await RunGenerationValidationLoopAsync(context, plan)
            If Not validatedInstructions.IsValid Then
                _currentPhase = LoopPhase.Failed
                Return LoopResult.FromValidationFailure(validatedInstructions)
            End If

            ' 3. 执行阶段
            _currentPhase = LoopPhase.Executing
            Dim execution = Await _instructionExecutor.ExecuteAsync(validatedInstructions.ParsedInstructions, context)

            ' 4. 执行后验证
            _currentPhase = LoopPhase.Verifying
            Dim verification = Await _resultVerifier.VerifyAsync(execution, context)

            ' 5. 若验证失败，记录但继续（执行后验证失败不阻塞）
            If Not verification.IsValid Then
                Debug.WriteLine($"[SelfCheckLoopController] 执行后验证未通过")
            End If

            _currentPhase = LoopPhase.Completed
            Return LoopResult.Success(execution)

        Catch ex As Exception
            Debug.WriteLine($"[SelfCheckLoopController] ExecuteLoopAsync 异常: {ex.Message}")
            _currentPhase = LoopPhase.Failed
            Dim errorResult = New LoopResult With {
                .IsSuccess = False,
                .ResultType = LoopResultType.ExecutionFailed,
                .UserMessage = $"执行失败: {ex.Message}"
            }
            errorResult.Errors.Add(New InstructionError(ErrorLevel.Critical, ex.Message))
            Return errorResult
        End Try
    End Function

    ''' <summary>
    ''' 执行Flush后校验（非完整Loop，仅校验AI响应）
    ''' </summary>
    Public Async Function PostFlushValidateAsync(
        aiResponse As String,
        expectedFormat As InstructionFormat,
        context As ExecutionContext) As Task(Of ValidationResult)

        _currentPhase = LoopPhase.Validating
        Return Await _instructionValidator.ValidateAsync(aiResponse, expectedFormat, context)
    End Function

    ''' <summary>
    ''' 运行修正Loop
    ''' </summary>
    Public Async Function RunCorrectionLoopAsync(
        context As ExecutionContext,
        originalResponse As String,
        validationResult As ValidationResult) As Task(Of CorrectionResult)

        If InstructionGenerator Is Nothing Then
            Return CorrectionResult.Failure(validationResult, 0)
        End If

        Dim iteration = 0
        Dim currentResponse = originalResponse
        Dim currentValidation = validationResult

        While iteration < MaxCorrectionIterations
            iteration += 1
            _iterationCount = iteration

            Try
                ' 构建修正请求
                Dim correctionPrompt = BuildCorrectionPrompt(currentResponse, currentValidation.Errors)

                ' 发送修正请求（这里简化处理，实际应调用AI）
                Dim correctedResponse = Await InstructionGenerator.GenerateCorrectionAsync(currentResponse, currentValidation.Errors)

                If String.IsNullOrEmpty(correctedResponse) Then
                    Continue While
                End If

                ' 重新校验
                currentValidation = Await _instructionValidator.ValidateAsync(
                    correctedResponse, context.ExpectedFormat, context)

                If currentValidation.IsValid Then
                    Return CorrectionResult.Success(correctedResponse, currentValidation.ParsedInstructions, iteration)
                End If

                currentResponse = correctedResponse

            Catch ex As Exception
                Debug.WriteLine($"[SelfCheckLoopController] 修正Loop第{iteration}轮异常: {ex.Message}")
            End Try
        End While

        ' 修正Loop耗尽
        Return CorrectionResult.Failure(currentValidation, iteration)
    End Function

    ''' <summary>
    ''' 运行规划Loop
    ''' </summary>
    Private Async Function RunPlanningLoopAsync(context As ExecutionContext) As Task(Of PlanningResult)
        If InstructionPlanner Is Nothing Then
            Return New PlanningResult With {.IsSuccess = True}
        End If

        Dim iteration = 0
        While iteration < MaxPlanningIterations
            iteration += 1
            Dim plan = Await InstructionPlanner.PlanAsync(context)
            If plan.IsSuccess Then
                Return plan
            End If
        End While

        Return New PlanningResult With {
            .IsSuccess = False,
            .ErrorMessage = "规划阶段超过最大迭代次数"
        }
    End Function

    ''' <summary>
    ''' 运行生成+校验Loop
    ''' </summary>
    Private Async Function RunGenerationValidationLoopAsync(
        context As ExecutionContext,
        plan As PlanningResult) As Task(Of ValidationResult)

        If InstructionGenerator Is Nothing Then
            Return New ValidationResult With {.IsValid = True}
        End If

        Dim iteration = 0
        While iteration < MaxExecutionIterations
            iteration += 1

            ' 生成AI请求
            Dim aiResponse = Await InstructionGenerator.GenerateAsync(context, plan)

            If String.IsNullOrEmpty(aiResponse) Then
                Continue While
            End If

            ' 校验响应
            Dim validation = Await _instructionValidator.ValidateAsync(
                aiResponse, context.ExpectedFormat, context)

            If validation.IsValid Then
                Return validation
            End If

            ' 校验失败，尝试修正
            If validation.CanAutoCorrect AndAlso InstructionGenerator IsNot Nothing Then
                Dim correction = Await RunCorrectionLoopAsync(context, aiResponse, validation)
                If correction.IsSuccess Then
                    Return New ValidationResult With {
                        .IsValid = True,
                        .ParsedInstructions = correction.Instructions,
                        .ExtractedContent = correction.CorrectedResponse
                    }
                End If
            End If
        End While

        Return New ValidationResult With {
            .IsValid = False,
            .Errors = New List(Of InstructionError) From {
                New InstructionError(ErrorLevel.Critical, "生成+校验阶段超过最大迭代次数")
            }
        }
    End Function

    ''' <summary>
    ''' 构建修正Prompt
    ''' </summary>
    Private Function BuildCorrectionPrompt(originalResponse As String, errors As List(Of InstructionError)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("你之前返回的操作指令存在以下问题，请修正后重新返回：")
        sb.AppendLine()

        For Each [error] In errors
            sb.AppendLine($"- [{[error].Level}] {[error].Message}")
            If [error].Suggestion IsNot Nothing Then
                sb.AppendLine($"  建议: {[error].Suggestion}")
            End If
        Next

        sb.AppendLine()
        sb.AppendLine("【你之前返回的内容】")
        sb.AppendLine(originalResponse)
        sb.AppendLine()
        sb.AppendLine("【修正要求】")
        sb.AppendLine("1. 仅修正上述错误，不要改变原有操作意图")
        sb.AppendLine("2. 严格按照指令协议格式返回")
        sb.AppendLine("3. 返回纯JSON，不要包含解释文字或代码块标记")

        Return sb.ToString()
    End Function

End Class
