' ShareRibbon\Loop\Executors\DslInstructionExecutor.vb
' DSL指令执行器 - 将DSL指令转换为Office DOM操作

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' DSL指令执行器 - 通用执行器基类
''' </summary>
Public Class DslInstructionExecutor
    Implements IInstructionExecutor

    ' Office应用特定执行器（由子类或WordAi项目提供实现）
    Protected _wordExecutor As IAppSpecificDslExecutor = Nothing
    Protected _excelExecutor As IAppSpecificDslExecutor = Nothing
    Protected _powerPointExecutor As IAppSpecificDslExecutor = Nothing

    ' 回滚栈
    Protected _undoStack As UndoStack = Nothing

    ''' <summary>
    ''' 回滚栈（外部可访问用于集成）
    ''' </summary>
    Public Property UndoStack As UndoStack
        Get
            If _undoStack Is Nothing Then
                _undoStack = New UndoStack()
            End If
            Return _undoStack
        End Get
        Set(value As UndoStack)
            _undoStack = value
        End Set
    End Property

    Public Async Function ExecuteAsync(
        instructions As List(Of Instruction),
        context As ExecutionContext) As Task(Of ExecutionResult) Implements IInstructionExecutor.ExecuteAsync

        Dim result As New ExecutionResult()
        Dim sw = System.Diagnostics.Stopwatch.StartNew()

        Try
            ' 获取Office应用特定执行器
            Dim appExecutor = GetAppSpecificExecutor(context.OfficeAppType)

            If appExecutor Is Nothing Then
                result.IsSuccess = False
                result.HasErrors = True
                result.Errors.Add(New InstructionError(ErrorLevel.Critical,
                    $"不支持的Office应用类型: {context.OfficeAppType}"))
                Return result
            End If

            ' 逐条执行指令
            For Each instruction In instructions
                Try
                    Dim opResult = Await ExecuteSingleInstructionAsync(instruction, appExecutor, context)
                    result.Operations.Add(opResult)

                    If opResult.IsSuccess Then
                        result.SuccessCount += 1

                        ' 记录回滚信息
                        If instruction.Rollback IsNot Nothing AndAlso instruction.Rollback.Count > 0 Then
                            RecordUndoOperation(instruction, opResult)
                        End If
                    Else
                        result.FailureCount += 1
                        result.HasErrors = True
                    End If

                Catch ex As Exception
                    Debug.WriteLine($"[DslInstructionExecutor] 执行指令 {instruction.Id} 异常: {ex.Message}")
                    result.Operations.Add(New OperationResult With {
                        .Instruction = instruction,
                        .IsSuccess = False,
                        .ErrorMessage = ex.Message
                    })
                    result.FailureCount += 1
                    result.HasErrors = True
                End Try
            Next

            result.IsSuccess = result.FailureCount = 0
            result.ExecutionTimeMs = sw.ElapsedMilliseconds

        Catch ex As Exception
            Debug.WriteLine($"[DslInstructionExecutor] ExecuteAsync 异常: {ex.Message}")
            result.IsSuccess = False
            result.HasErrors = True
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 执行单条指令
    ''' </summary>
    Protected Overridable Async Function ExecuteSingleInstructionAsync(
        instruction As Instruction,
        appExecutor As IAppSpecificDslExecutor,
        context As ExecutionContext) As Task(Of OperationResult)

        Return Await appExecutor.ExecuteInstructionAsync(instruction, context)
    End Function

    ''' <summary>
    ''' 获取Office应用特定执行器
    ''' </summary>
    Protected Function GetAppSpecificExecutor(appType As OfficeAppType) As IAppSpecificDslExecutor
        Select Case appType
            Case OfficeAppType.Word
                Return _wordExecutor
            Case OfficeAppType.Excel
                Return _excelExecutor
            Case OfficeAppType.PowerPoint
                Return _powerPointExecutor
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' 记录回滚操作
    ''' </summary>
    Protected Sub RecordUndoOperation(instruction As Instruction, opResult As OperationResult)
        Try
            Dim undoOp = New DslUndoOperation(instruction, opResult)
            UndoStack.Push(undoOp)
        Catch ex As Exception
            Debug.WriteLine($"[DslInstructionExecutor] 记录回滚操作失败: {ex.Message}")
        End Try
    End Sub

End Class

''' <summary>
''' 应用特定DSL执行器接口
''' </summary>
Public Interface IAppSpecificDslExecutor

    ''' <summary>
    ''' 执行单条指令
    ''' </summary>
    Function ExecuteInstructionAsync(instruction As Instruction, context As ExecutionContext) As Task(Of OperationResult)

End Interface

''' <summary>
''' DSL回滚操作
''' </summary>
Public Class DslUndoOperation
    Implements UndoableOperation

    Private ReadOnly _instruction As Instruction
    Private ReadOnly _originalSnapshot As Object

    Public ReadOnly Property Description As String Implements UndoableOperation.Description
        Get
            Return $"撤销: {_instruction.GetDescription()}"
        End Get
    End Property

    Public ReadOnly Property Timestamp As DateTime Implements UndoableOperation.Timestamp
        Get
            Return DateTime.Now
        End Get
    End Property

    Public ReadOnly Property InstructionId As String Implements UndoableOperation.InstructionId
        Get
            Return _instruction.Id
        End Get
    End Property

    Public Sub New(instruction As Instruction, opResult As OperationResult)
        _instruction = instruction
        _originalSnapshot = CaptureSnapshot(opResult)
    End Sub

    Public Function Undo() As Boolean Implements UndoableOperation.Undo
        Try
            ' 根据回滚信息执行反向操作
            If _instruction.Rollback IsNot Nothing Then
                Dim rollbackOp = _instruction.Rollback("op")?.ToString()
                If Not String.IsNullOrEmpty(rollbackOp) Then
                    ' 执行回滚操作
                    Debug.WriteLine($"[DslUndoOperation] 执行回滚: {rollbackOp}")
                    ' 实际回滚逻辑由子类实现
                    Return True
                End If
            End If
            Return True
        Catch ex As Exception
            Debug.WriteLine($"[DslUndoOperation] Undo失败: {ex.Message}")
            Return False
        End Try
    End Function

    Public Function Redo() As Boolean Implements UndoableOperation.Redo
        ' 重做即重新执行原指令
        Try
            Debug.WriteLine($"[DslUndoOperation] 执行重做: {_instruction.Operation}")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"[DslUndoOperation] Redo失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 捕获快照（用于回滚）
    ''' </summary>
    Private Function CaptureSnapshot(opResult As OperationResult) As Object
        ' 简化实现，实际应捕获Range的完整状态
        Return Nothing
    End Function

End Class
