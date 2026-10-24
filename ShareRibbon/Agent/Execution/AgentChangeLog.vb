Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' 一次任务中的单条写操作记录。字段刻意只保留可核对的事实：
    ''' 改了什么、是否真的变了、是否成功、有没有撤销依据。
    ''' </summary>
    Public Class AgentChangeRecord
        Public Property ChangeId As String = Guid.NewGuid().ToString()
        Public Property Seq As Integer
        Public Property SessionId As String = ""
        Public Property AppType As String = ""
        Public Property ToolId As String = ""
        Public Property StepIndex As Integer
        Public Property StepDescription As String = ""
        Public Property Summary As String = ""
        Public Property TargetRefs As New JArray()
        Public Property Changed As Boolean
        Public Property Success As Boolean
        Public Property RolledBack As Boolean
        Public Property PartialApply As Boolean
        Public Property UndoAvailable As Boolean
        Public Property UndoUnits As Integer
        Public Property UndoPointId As String = ""
        Public Property RollbackError As String = ""
        Public Property DocHashBefore As String = ""
        Public Property DocHashAfter As String = ""
        Public Property ObservedAt As DateTime = DateTime.Now

        Public Function ToJson() As JObject
            Return New JObject From {
                {"changeId", ChangeId},
                {"seq", Seq},
                {"toolId", ToolId},
                {"stepIndex", StepIndex},
                {"stepDescription", StepDescription},
                {"summary", Summary},
                {"targetRefs", If(TargetRefs, New JArray())},
                {"changed", Changed},
                {"success", Success},
                {"rolledBack", RolledBack},
                {"partialApply", PartialApply},
                {"undoAvailable", UndoAvailable},
                {"undoUnits", UndoUnits},
                {"undoPointId", UndoPointId},
                {"rollbackError", RollbackError},
                {"docHashBefore", DocHashBefore},
                {"docHashAfter", DocHashAfter},
                {"observedAt", ObservedAt.ToString("yyyy-MM-dd HH:mm:ss")}
            }
        End Function
    End Class

    ''' <summary>
    ''' 单轮变更日志：把每一步写操作的观察聚合成可列出的变更清单。
    ''' 只收录宿主声明 kind=write / code_execution 且带 changed 的结果，
    ''' 不把聊天、读取或没有观察依据的调用混进“已修改”列表。
    ''' </summary>
    Public Class AgentRunChangeLog

        Private ReadOnly _entries As New List(Of AgentChangeRecord)()

        Public ReadOnly Property Entries As List(Of AgentChangeRecord)
            Get
                Return _entries
            End Get
        End Property

        Public ReadOnly Property ChangedCount As Integer
            Get
                Return _entries.Where(Function(item) item.Changed).Count()
            End Get
        End Property

        ''' <summary>
        ''' 真正留在文档里的改动数：回滚成功的不算，避免取消提示把已还原的步骤也报成“已写入”。
        ''' </summary>
        Public ReadOnly Property WrittenCount As Integer
            Get
                Return _entries.Where(Function(item) item.Changed AndAlso Not item.RolledBack).Count()
            End Get
        End Property

        Public ReadOnly Property HasChanges As Boolean
            Get
                Return WrittenCount > 0
            End Get
        End Property

        ''' <summary>
        ''' 宿主声明可撤销的单元总数；0 表示本轮没有可安全撤销的改动。
        ''' </summary>
        Public ReadOnly Property UndoableUnits As Integer
            Get
                Return _entries.
                    Where(Function(item) item.Changed AndAlso Not item.RolledBack AndAlso item.UndoAvailable).
                    Sum(Function(item) item.UndoUnits)
            End Get
        End Property

        ''' <summary>
        ''' 只有所有残留改动都能按宿主单元撤销，且记录到前后文档指纹时，才允许“撤销本轮”。
        ''' </summary>
        Public ReadOnly Property FullyUndoable As Boolean
            Get
                Dim written = _entries.Where(Function(item) item.Changed AndAlso Not item.RolledBack).ToList()
                If written.Count = 0 Then Return False
                If written.Any(Function(item) Not item.UndoAvailable OrElse item.UndoUnits <= 0) Then Return False
                Return FirstBeforeHash.Length > 0 AndAlso LastAfterHash.Length > 0
            End Get
        End Property

        Public ReadOnly Property FirstBeforeHash As String
            Get
                Return _entries.
                    Where(Function(item) item.Changed AndAlso Not item.RolledBack AndAlso item.DocHashBefore <> "").
                    Select(Function(item) item.DocHashBefore).
                    FirstOrDefault()
            End Get
        End Property

        Public ReadOnly Property LastAfterHash As String
            Get
                Return _entries.
                    Where(Function(item) item.Changed AndAlso Not item.RolledBack AndAlso item.DocHashAfter <> "").
                    Select(Function(item) item.DocHashAfter).
                    LastOrDefault()
            End Get
        End Property

        ''' <summary>
        ''' 从一步的工具结果生成变更记录；不是写操作或没有观察对象时返回 Nothing。
        ''' </summary>
        Public Function AddFromStep(toolId As String,
                                    stepIndex As Integer,
                                    stepDescription As String,
                                    toolResult As ToolResult,
                                    undoPointName As String,
                                    canUndo As Boolean) As AgentChangeRecord
            If toolResult Is Nothing Then Return Nothing
            Dim observation = TryReadObservation(toolResult)
            If observation Is Nothing Then Return Nothing

            Dim kind = ReadString(observation, "kind")
            If Not String.Equals(kind, "write", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(kind, "code_execution", StringComparison.OrdinalIgnoreCase) Then
                Return Nothing
            End If

            Dim changedToken = observation("changed")
            If changedToken Is Nothing OrElse changedToken.Type <> JTokenType.Boolean Then Return Nothing
            Dim changed = changedToken.Value(Of Boolean)()

            Dim rollback = TryCast(observation("rollback"), JObject)
            Dim rollbackError = If(rollback IsNot Nothing, ReadString(rollback, "error"), "")
            Dim rolledBack = rollback IsNot Nothing AndAlso ReadBool(rollback, "succeeded")
            Dim partialApply = Not toolResult.Success AndAlso changed AndAlso Not rolledBack
            Dim explicitPartial As JToken = If(rollback IsNot Nothing, rollback("partialApply"), Nothing)
            If explicitPartial IsNot Nothing AndAlso explicitPartial.Type = JTokenType.Boolean Then
                partialApply = explicitPartial.Value(Of Boolean)()
            End If

            Dim beforeHash = ReadDocHash(observation, "before")
            Dim afterHash = ReadDocHash(observation, "after")
            Dim items = TryCast(observation("items"), JArray)
            Dim undoUnits As Integer
            If observation("undoUnits") IsNot Nothing Then
                ' 批次级声明优先：它已经是逐条汇总值，再累加 items 会翻倍撤销
                undoUnits = ReadInt(observation, "undoUnits", 0)
            ElseIf items IsNot Nothing Then
                undoUnits = items.OfType(Of JObject)().
                    Where(Function(item) IsCommittedChange(item)).
                    Sum(Function(item) ReadInt(item, "undoUnits", 0))
            End If
            If items IsNot Nothing Then
                ' 已回滚或未真正改动文档的条目不能参与指纹推导，
                ' 否则“撤销本轮”会拿中途哈希当基线，撤销后必然对不上
                If beforeHash.Length = 0 Then
                    beforeHash = items.OfType(Of JObject)().
                        Where(Function(item) IsCommittedChange(item)).
                        Select(Function(item) ReadDocHash(item, "before")).
                        FirstOrDefault(Function(hash) hash.Length > 0)
                    If beforeHash Is Nothing Then beforeHash = ""
                End If
                If afterHash.Length = 0 Then
                    afterHash = items.OfType(Of JObject)().
                        Where(Function(item) IsCommittedChange(item)).
                        Select(Function(item) ReadDocHash(item, "after")).
                        Where(Function(hash) hash.Length > 0).
                        LastOrDefault()
                    If afterHash Is Nothing Then afterHash = ""
                End If
            End If
            If undoUnits < 0 Then undoUnits = 0

            Dim record As New AgentChangeRecord With {
                .Seq = _entries.Count + 1,
                .ToolId = If(toolId, ""),
                .StepIndex = stepIndex,
                .StepDescription = If(stepDescription, ""),
                .Summary = ReadString(observation, "summary"),
                .TargetRefs = TryCast(observation("targetRefs")?.DeepClone(), JArray),
                .Changed = changed,
                .Success = toolResult.Success,
                .RolledBack = rolledBack,
                .PartialApply = partialApply,
                .UndoAvailable = canUndo AndAlso undoUnits > 0,
                .UndoUnits = undoUnits,
                .UndoPointId = If(undoPointName, ""),
                .RollbackError = rollbackError,
                .DocHashBefore = beforeHash,
                .DocHashAfter = afterHash
            }
            If record.Summary.Length = 0 Then record.Summary = If(toolResult.Message, "")
            If record.TargetRefs Is Nothing Then record.TargetRefs = New JArray()

            _entries.Add(record)
            Return record
        End Function

        Public Function ToJson() As JArray
            Dim result As New JArray()
            For Each entry In _entries
                result.Add(entry.ToJson())
            Next
            Return result
        End Function

        Private Shared Function TryReadObservation(toolResult As ToolResult) As JObject
            If toolResult Is Nothing OrElse toolResult.Observation Is Nothing Then Return Nothing
            Try
                Return TryCast(JToken.FromObject(toolResult.Observation), JObject)
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function ReadString(obj As JObject, name As String) As String
            If obj Is Nothing Then Return ""
            Return If(obj(name)?.ToString(), "")
        End Function

        Private Shared Function ReadBool(obj As JObject, name As String) As Boolean
            If obj Is Nothing Then Return False
            Dim token = obj(name)
            If token Is Nothing OrElse token.Type <> JTokenType.Boolean Then Return False
            Return token.Value(Of Boolean)()
        End Function

        Private Shared Function ReadInt(obj As JObject, name As String, fallback As Integer) As Integer
            If obj Is Nothing Then Return fallback
            Dim token = obj(name)
            If token Is Nothing Then Return fallback
            Try
                Return token.Value(Of Integer)()
            Catch
                Return fallback
            End Try
        End Function

        Private Shared Function IsCommittedChange(item As JObject) As Boolean
            If Not ReadBool(item, "changed") Then Return False
            Dim rollback = TryCast(item("rollback"), JObject)
            Return rollback Is Nothing OrElse Not ReadBool(rollback, "succeeded")
        End Function

        Private Shared Function ReadDocHash(obj As JObject, name As String) As String
            If obj Is Nothing Then Return ""
            Dim inner = TryCast(obj(name), JObject)
            If inner Is Nothing Then Return ""
            Return ReadString(inner, "documentTextHash")
        End Function
    End Class

    ''' <summary>
    ''' 一轮任务的可撤销范围：只有残留改动全部带宿主撤销单元和前后指纹时，才允许自动撤销。
    ''' </summary>
    Public Class AgentRunUndoScope
        Public Property RunId As String = ""
        Public Property WrittenCount As Integer
        Public Property UndoUnits As Integer
        Public Property FullyUndoable As Boolean
        Public Property BaselineHash As String = ""
        Public Property ExpectedHash As String = ""
        Public Property Reason As String = ""
    End Class

    ''' <summary>
    ''' 撤销本轮的结果。Verified 只表示撤销后文档指纹与撤销前一致，不表示用户主观满意。
    ''' </summary>
    Public Class AgentRunUndoResult
        Public Property Success As Boolean
        Public Property Verified As Boolean
        Public Property Message As String = ""

        Public Shared Function Succeeded(message As String, Optional verified As Boolean = False) As AgentRunUndoResult
            Return New AgentRunUndoResult With {.Success = True, .Verified = verified, .Message = If(message, "")}
        End Function

        Public Shared Function Failed(message As String) As AgentRunUndoResult
            Return New AgentRunUndoResult With {.Success = False, .Verified = False, .Message = If(message, "")}
        End Function
    End Class

End Namespace
