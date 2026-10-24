Imports System.Data.SQLite
Imports System.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 单轮变更日志的持久化读写。一次任务结束后按 run_id 可列出本轮所有写操作，
''' 供取消/失败后的“已写入 N 处”说明和后续“撤销本轮”使用。
''' </summary>
Public Class AgentChangeLogRepository

    Public Shared Sub SaveRun(runId As String,
                              sessionId As String,
                              appType As String,
                              changeLog As Agent.AgentRunChangeLog)
        If String.IsNullOrWhiteSpace(runId) OrElse changeLog Is Nothing OrElse changeLog.Entries.Count = 0 Then Return
        OfficeAiDatabase.EnsureInitialized()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using tx = conn.BeginTransaction()
                For Each entry In changeLog.Entries
                    Using cmd As New SQLiteCommand(
                        "INSERT OR REPLACE INTO agent_run_change " &
                        "(change_id, run_id, session_id, app_type, seq, tool_id, step_index, step_description, summary, target_refs_json, " &
                        " changed, success, rolled_back, partial_apply, undo_available, undo_units, undo_point_id, rollback_error, " &
                        " doc_hash_before, doc_hash_after, observed_at) " &
                        "VALUES (@change_id, @run_id, @session_id, @app_type, @seq, @tool_id, @step_index, @step_description, @summary, @target_refs_json, " &
                        " @changed, @success, @rolled_back, @partial_apply, @undo_available, @undo_units, @undo_point_id, @rollback_error, " &
                        " @doc_hash_before, @doc_hash_after, @observed_at)", conn, tx)
                        cmd.Parameters.AddWithValue("@change_id", If(entry.ChangeId, Guid.NewGuid().ToString()))
                        cmd.Parameters.AddWithValue("@run_id", runId)
                        cmd.Parameters.AddWithValue("@session_id", If(sessionId, ""))
                        cmd.Parameters.AddWithValue("@app_type", If(appType, ""))
                        cmd.Parameters.AddWithValue("@seq", entry.Seq)
                        cmd.Parameters.AddWithValue("@tool_id", If(entry.ToolId, ""))
                        cmd.Parameters.AddWithValue("@step_index", entry.StepIndex)
                        cmd.Parameters.AddWithValue("@step_description", If(entry.StepDescription, ""))
                        cmd.Parameters.AddWithValue("@summary", If(entry.Summary, ""))
                        cmd.Parameters.AddWithValue("@target_refs_json", If(entry.TargetRefs, New JArray()).ToString(Formatting.None))
                        cmd.Parameters.AddWithValue("@changed", If(entry.Changed, 1, 0))
                        cmd.Parameters.AddWithValue("@success", If(entry.Success, 1, 0))
                        cmd.Parameters.AddWithValue("@rolled_back", If(entry.RolledBack, 1, 0))
                        cmd.Parameters.AddWithValue("@partial_apply", If(entry.PartialApply, 1, 0))
                        cmd.Parameters.AddWithValue("@undo_available", If(entry.UndoAvailable, 1, 0))
                        cmd.Parameters.AddWithValue("@undo_units", Math.Max(0, entry.UndoUnits))
                        cmd.Parameters.AddWithValue("@undo_point_id", If(entry.UndoPointId, ""))
                        cmd.Parameters.AddWithValue("@rollback_error", If(entry.RollbackError, ""))
                        cmd.Parameters.AddWithValue("@doc_hash_before", If(entry.DocHashBefore, ""))
                        cmd.Parameters.AddWithValue("@doc_hash_after", If(entry.DocHashAfter, ""))
                        cmd.Parameters.AddWithValue("@observed_at", entry.ObservedAt.ToString("yyyy-MM-dd HH:mm:ss"))
                        cmd.ExecuteNonQuery()
                    End Using
                Next
                tx.Commit()
            End Using
        End Using
    End Sub

    Public Shared Function GetByRun(runId As String) As List(Of Agent.AgentChangeRecord)
        Dim result As New List(Of Agent.AgentChangeRecord)()
        If String.IsNullOrWhiteSpace(runId) Then Return result
        OfficeAiDatabase.EnsureInitialized()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT change_id, seq, tool_id, step_index, step_description, summary, target_refs_json, " &
                " changed, success, rolled_back, partial_apply, undo_available, undo_units, undo_point_id, rollback_error, " &
                " doc_hash_before, doc_hash_after, observed_at " &
                "FROM agent_run_change WHERE run_id = @run_id ORDER BY seq", conn)
                cmd.Parameters.AddWithValue("@run_id", runId)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim record As New Agent.AgentChangeRecord With {
                            .ChangeId = ReadText(rdr, 0),
                            .Seq = rdr.GetInt32(1),
                            .ToolId = ReadText(rdr, 2),
                            .StepIndex = If(rdr.IsDBNull(3), 0, rdr.GetInt32(3)),
                            .StepDescription = ReadText(rdr, 4),
                            .Summary = ReadText(rdr, 5),
                            .TargetRefs = ParseArray(ReadText(rdr, 6)),
                            .Changed = ReadBool(rdr, 7),
                            .Success = ReadBool(rdr, 8),
                            .RolledBack = ReadBool(rdr, 9),
                            .PartialApply = ReadBool(rdr, 10),
                            .UndoAvailable = ReadBool(rdr, 11),
                            .UndoUnits = If(rdr.IsDBNull(12), 0, rdr.GetInt32(12)),
                            .UndoPointId = ReadText(rdr, 13),
                            .RollbackError = ReadText(rdr, 14),
                            .DocHashBefore = ReadText(rdr, 15),
                            .DocHashAfter = ReadText(rdr, 16),
                            .ObservedAt = ParseTime(ReadText(rdr, 17))
                        }
                        result.Add(record)
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    ''' <summary>
    ''' 计算某轮的可撤销范围：只统计残留改动，且要求每一步都带宿主撤销单元和文档前后指纹。
    ''' </summary>
    Public Shared Function GetUndoScopeByRun(runId As String) As Agent.AgentRunUndoScope
        Dim scope As New Agent.AgentRunUndoScope With {.RunId = If(runId, "")}
        If String.IsNullOrWhiteSpace(runId) Then
            scope.Reason = "缺少 runId"
            Return scope
        End If

        Dim entries = GetByRun(runId)
        Dim written = entries.Where(Function(item) item.Changed AndAlso Not item.RolledBack).ToList()
        scope.WrittenCount = written.Count
        If written.Count = 0 Then
            scope.Reason = "本轮没有残留改动"
            Return scope
        End If

        scope.UndoUnits = written.
            Where(Function(item) item.UndoAvailable).
            Sum(Function(item) item.UndoUnits)
        scope.BaselineHash = written.
            Where(Function(item) item.DocHashBefore <> "").
            Select(Function(item) item.DocHashBefore).
            FirstOrDefault()
        scope.ExpectedHash = written.
            Where(Function(item) item.DocHashAfter <> "").
            Select(Function(item) item.DocHashAfter).
            LastOrDefault()
        scope.FullyUndoable = written.All(Function(item) item.UndoAvailable AndAlso item.UndoUnits > 0) AndAlso
                              Not String.IsNullOrWhiteSpace(scope.BaselineHash) AndAlso
                              Not String.IsNullOrWhiteSpace(scope.ExpectedHash)
        If Not scope.FullyUndoable Then
            scope.Reason = "本轮包含无法自动撤销的改动（缺少宿主撤销单元或文档指纹）"
        End If
        Return scope
    End Function

    ''' <summary>
    ''' 撤销本轮成功后把残留改动标记为已回滚，后续再次查询不会把同一批改动当成可撤销对象。
    ''' </summary>
    Public Shared Sub MarkRunRolledBack(runId As String)
        If String.IsNullOrWhiteSpace(runId) Then Return
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE agent_run_change SET rolled_back = 1 " &
                "WHERE run_id = @run_id AND changed = 1 AND rolled_back = 0", conn)
                cmd.Parameters.AddWithValue("@run_id", runId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Shared Function ReadText(rdr As SQLiteDataReader, ordinal As Integer) As String
        Return If(rdr.IsDBNull(ordinal), "", rdr.GetString(ordinal))
    End Function

    Private Shared Function ReadBool(rdr As SQLiteDataReader, ordinal As Integer) As Boolean
        Return Not rdr.IsDBNull(ordinal) AndAlso rdr.GetInt32(ordinal) <> 0
    End Function

    Private Shared Function ParseArray(value As String) As JArray
        If String.IsNullOrWhiteSpace(value) Then Return New JArray()
        Try
            Return TryCast(JToken.Parse(value), JArray)
        Catch
            Return New JArray()
        End Try
    End Function

    Private Shared Function ParseTime(value As String) As DateTime
        Dim parsed As DateTime = DateTime.Now
        If DateTime.TryParse(value, parsed) Then Return parsed
        Return DateTime.Now
    End Function

End Class
