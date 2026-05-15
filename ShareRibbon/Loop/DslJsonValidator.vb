' ShareRibbon\Loop\DslJsonValidator.vb
' DSL JSON验证器 - 验证AI返回的DSL格式指令

Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

''' <summary>
''' DSL JSON验证器
''' </summary>
Public Class DslJsonValidator
    Implements IInstructionValidator

    ''' <summary>
    ''' 验证AI响应是否符合DSL格式
    ''' </summary>
    Public Async Function ValidateAsync(
        aiResponse As String,
        expectedFormat As InstructionFormat,
        context As ExecutionContext) As Task(Of ValidationResult) Implements IInstructionValidator.ValidateAsync

        Dim result As New ValidationResult()
        result.IsValid = False

        Try
            ' 去除可能存在的markdown代码块
            Dim content = CleanMarkdownCodeBlocks(aiResponse)
            If String.IsNullOrWhiteSpace(content) Then
                result.Errors.Add(New InstructionError(ErrorLevel.Critical, "响应内容为空"))
                Return result
            End If

            ' 解析JSON
            Dim json = JObject.Parse(content)

            ' 验证基本结构
            If Not ValidateBasicStructure(json, result) Then
                Return result
            End If

            ' 验证指令列表
            If Not ValidateInstructions(json, result, context) Then
                Return result
            End If

            ' 验证元数据
            ValidateMetadata(json, result)

            result.IsValid = result.Errors.Count = 0
            If result.IsValid Then
                ' 解析成功，提取指令
                result.ParsedInstructions = ExtractInstructions(json)
                result.ExtractedContent = content
            End If

        Catch ex As Exception
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, $"JSON解析失败: {ex.Message}"))
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 去除markdown代码块
    ''' </summary>
    Private Function CleanMarkdownCodeBlocks(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return String.Empty

        ' 去除 ```json 和 ``` 之间的代码块
        Dim cleaned = Regex.Replace(text, "```(?:json)?\s*([\s\S]*?)\s*```", "$1", RegexOptions.IgnoreCase)
        Return cleaned.Trim()
    End Function

    ''' <summary>
    ''' 验证基本JSON结构
    ''' </summary>
    Private Function ValidateBasicStructure(json As JObject, result As ValidationResult) As Boolean
        ' 验证version字段
        If json("version") Is Nothing OrElse Not json("version").Type = JTokenType.String Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, "缺少或无效的version字段"))
            Return False
        End If

        ' 验证protocol字段
        If json("protocol") Is Nothing OrElse Not json("protocol").Type = JTokenType.String Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, "缺少或无效的protocol字段"))
            Return False
        End If

        If Not String.Equals(json("protocol").ToString(), "office-dsl", StringComparison.OrdinalIgnoreCase) Then
            result.Errors.Add(New InstructionError(ErrorLevel.Warning, $"不支持的协议类型: {json("protocol").ToString()}"))
        End If

        ' 验证operation字段
        If json("operation") Is Nothing OrElse Not json("operation").Type = JTokenType.String Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, "缺少或无效的operation字段"))
            Return False
        End If

        ' 验证target字段
        If json("target") IsNot Nothing AndAlso Not json("target").Type = JTokenType.Object Then
            result.Errors.Add(New InstructionError(ErrorLevel.Warning, "target字段必须是对象类型"))
        End If

        Return True
    End Function

    ''' <summary>
    ''' 验证指令列表
    ''' </summary>
    Private Function ValidateInstructions(json As JObject, result As ValidationResult, context As ExecutionContext) As Boolean
        If json("instructions") Is Nothing Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, "缺少instructions字段"))
            Return False
        End If

        If json("instructions").Type <> JTokenType.Array Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, "instructions必须是数组类型"))
            Return False
        End If

        Dim instructionsArray = CType(json("instructions"), JArray)
        If instructionsArray.Count = 0 Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, "指令列表为空"))
            Return True ' 空指令列表可能是有效的（如无需操作）
        End If

        For i = 0 To instructionsArray.Count - 1
            Dim instructionObj = instructionsArray(i)
            If instructionObj.Type <> JTokenType.Object Then
                result.Errors.Add(New InstructionError(ErrorLevel.Warning, $"指令{i}不是对象类型"))
                Continue For
            End If

            ValidateSingleInstruction(instructionObj, i, result, context)
        Next

        Return result.Errors.Count = 0
    End Function

    ''' <summary>
    ''' 验证单条指令
    ''' </summary>
    Private Sub ValidateSingleInstruction(instructionObj As JObject, index As Integer, result As ValidationResult, context As ExecutionContext)
        Dim prefix = $"指令{index}"

        ' 验证id字段
        If instructionObj("id") Is Nothing Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: 缺少id字段"))
        ElseIf Not instructionObj("id").Type = JTokenType.String Then
            result.Errors.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: id必须是字符串类型"))
        End If

        ' 验证op字段
        If instructionObj("op") Is Nothing Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, $"{prefix}: 缺少op字段"))
            Return
        ElseIf Not instructionObj("op").Type = JTokenType.String Then
            result.Errors.Add(New InstructionError(ErrorLevel.Critical, $"{prefix}: op必须是字符串类型"))
            Return
        End If

        Dim operation = instructionObj("op").ToString()
        If Not InstructionRegistry.IsValidOperation(operation) Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: 未知的操作类型: {operation}"))
        End If

        ' 验证target字段
        If instructionObj("target") IsNot Nothing AndAlso Not instructionObj("target").Type = JTokenType.Object Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: target必须是对象类型"))
        End If

        ' 验证params字段
        If instructionObj("params") IsNot Nothing AndAlso Not instructionObj("params").Type = JTokenType.Object Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: params必须是对象类型"))
        End If

        ' 验证expected字段
        If instructionObj("expected") IsNot Nothing AndAlso Not instructionObj("expected").Type = JTokenType.Object Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: expected必须是对象类型"))
        End If

        ' 验证rollback字段
        If instructionObj("rollback") IsNot Nothing AndAlso Not instructionObj("rollback").Type = JTokenType.Object Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: rollback必须是对象类型"))
        End If

        ' 详细验证参数
        If instructionObj("params") IsNot Nothing Then
            Dim paramValidation = InstructionRegistry.ValidateParameters(operation, instructionObj("params"))
            If Not paramValidation.IsValid Then
                result.Errors.Add(New InstructionError(ErrorLevel.Warning, $"{prefix}: 参数验证失败: {paramValidation.ErrorMessage}"))
            End If
        End If
    End Sub

    ''' <summary>
    ''' 验证元数据
    ''' </summary>
    Private Sub ValidateMetadata(json As JObject, result As ValidationResult)
        If json("metadata") Is Nothing Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, "缺少metadata字段"))
            Return
        End If

        If Not json("metadata").Type = JTokenType.Object Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, "metadata必须是对象类型"))
            Return
        End If

        ' 验证estimatedOperations（可选）
        If json("metadata")("estimatedOperations") IsNot Nothing AndAlso Not (json("metadata")("estimatedOperations").Type = JTokenType.Integer OrElse json("metadata")("estimatedOperations").Type = JTokenType.Float) Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, "estimatedOperations必须是数值类型"))
        End If

        ' 验证hasDestructiveOps（可选）
        If json("metadata")("hasDestructiveOps") IsNot Nothing AndAlso Not json("metadata")("hasDestructiveOps").Type = JTokenType.Boolean Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, "hasDestructiveOps必须是布尔类型"))
        End If

        ' 验证requiresConfirmation（可选）
        If json("metadata")("requiresConfirmation") IsNot Nothing AndAlso Not json("metadata")("requiresConfirmation").Type = JTokenType.Boolean Then
            result.Warnings.Add(New InstructionError(ErrorLevel.Warning, "requiresConfirmation必须是布尔类型"))
        End If
    End Sub

    ''' <summary>
    ''' 提取指令列表
    ''' </summary>
    Private Function ExtractInstructions(json As JObject) As List(Of Instruction)
        Dim instructions As New List(Of Instruction)()

        If json("instructions") Is Nothing OrElse json("instructions").Type <> JTokenType.Array Then
            Return instructions
        End If

        Dim instructionsArray = CType(json("instructions"), JArray)
        For Each item In instructionsArray
            If item.Type <> JTokenType.Object Then Continue For

            Dim itemObj = CType(item, JObject)
            If itemObj("op") Is Nothing Then Continue For

            Dim op = itemObj("op").ToString()
            Dim params = If(itemObj("params"), New JObject())
            Dim id = itemObj("id")?.ToString()

            Dim instruction = New Instruction(op, CType(params, JObject), id)
            If itemObj("target") IsNot Nothing Then
                instruction.Target = CType(itemObj("target"), JObject)
            End If
            If itemObj("expected") IsNot Nothing Then
                instruction.Expected = CType(itemObj("expected"), JObject)
            End If
            If itemObj("rollback") IsNot Nothing Then
                instruction.Rollback = CType(itemObj("rollback"), JObject)
            End If
            If itemObj("metadata") IsNot Nothing Then
                instruction.Metadata = CType(itemObj("metadata"), JObject)
            End If

            instructions.Add(instruction)
        Next

        Return instructions
    End Function

End Class
