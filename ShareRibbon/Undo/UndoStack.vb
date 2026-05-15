' ShareRibbon\Undo\UndoStack.vb
' 统一回滚栈 - 支持指令级别的撤销和重做

Imports System.Collections.Generic
Imports System.Diagnostics

''' <summary>
''' 统一回滚栈 - 管理所有可撤销操作
''' </summary>
Public Class UndoStack

    ' 最大回滚步数
    Public Property MaxSize As Integer = 50

    ' 操作栈
    Private ReadOnly _undoStack As New Stack(Of UndoableOperation)()
    Private ReadOnly _redoStack As New Stack(Of UndoableOperation)()

    ''' <summary>当前栈深度</summary>
    Public ReadOnly Property Count As Integer
        Get
            Return _undoStack.Count
        End Get
    End Property

    ''' <summary>是否可以撤销</summary>
    Public ReadOnly Property CanUndo As Boolean
        Get
            Return _undoStack.Count > 0
        End Get
    End Property

    ''' <summary>是否可以重做</summary>
    Public ReadOnly Property CanRedo As Boolean
        Get
            Return _redoStack.Count > 0
        End Get
    End Property

    ''' <summary>
    ''' 压入操作
    ''' </summary>
    Public Sub Push(operation As UndoableOperation)
        If operation Is Nothing Then Return

        _undoStack.Push(operation)
        _redoStack.Clear()

        ' 限制栈大小
        While _undoStack.Count > MaxSize
            ' 移除最旧的操作
            Dim tempStack As New Stack(Of UndoableOperation)()
            Dim skipFirst = True
            For Each op In _undoStack
                If skipFirst Then
                    skipFirst = False
                    Continue For
                End If
                tempStack.Push(op)
            Next
            _undoStack.Clear()
            ' 重新压入（顺序反转）
            Dim reverseList As New List(Of UndoableOperation)(tempStack)
            reverseList.Reverse()
            For Each op In reverseList
                _undoStack.Push(op)
            Next
        End While
    End Sub

    ''' <summary>
    ''' 撤销一步
    ''' </summary>
    Public Function Undo() As Boolean
        If Not CanUndo Then Return False

        Try
            Dim operation = _undoStack.Pop()
            Dim success = operation.Undo()
            If success Then
                _redoStack.Push(operation)
            Else
                ' 撤销失败，放回undo栈
                _undoStack.Push(operation)
                Return False
            End If
            Return True
        Catch ex As Exception
            Debug.WriteLine($"[UndoStack] Undo失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 重做一次
    ''' </summary>
    Public Function Redo() As Boolean
        If Not CanRedo Then Return False

        Try
            Dim operation = _redoStack.Pop()
            Dim success = operation.Redo()
            If success Then
                _undoStack.Push(operation)
            Else
                ' 重做失败，放回redo栈
                _redoStack.Push(operation)
                Return False
            End If
            Return True
        Catch ex As Exception
            Debug.WriteLine($"[UndoStack] Redo失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 清空栈
    ''' </summary>
    Public Sub Clear()
        _undoStack.Clear()
        _redoStack.Clear()
    End Sub

    ''' <summary>
    ''' 获取当前所有操作描述（用于UI展示）
    ''' </summary>
    Public Function GetOperationDescriptions() As List(Of String)
        Dim descriptions As New List(Of String)()
        For Each op In _undoStack
            descriptions.Add(op.Description)
        Next
        descriptions.Reverse()
        Return descriptions
    End Function

End Class

''' <summary>
''' 可撤销操作接口
''' </summary>
Public Interface UndoableOperation

    ''' <summary>操作描述</summary>
    ReadOnly Property Description As String

    ''' <summary>操作时间戳</summary>
    ReadOnly Property Timestamp As DateTime

    ''' <summary>关联的指令ID</summary>
    ReadOnly Property InstructionId As String

    ''' <summary>撤销操作</summary>
    Function Undo() As Boolean

    ''' <summary>重做操作</summary>
    Function Redo() As Boolean

End Interface
