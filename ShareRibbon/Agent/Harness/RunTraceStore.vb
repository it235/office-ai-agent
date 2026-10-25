Imports System.Data.SQLite
Imports Newtonsoft.Json

Namespace Agent.Harness

    Public Interface IRunTraceStore
        Sub StartRun(runId As String, turn As UserTurn, startedAt As DateTime)
        Sub AppendStep(runId As String, seq As Integer, toolId As String, status As String, message As String, errorCode As String, observation As Object, startedAt As DateTime, finishedAt As DateTime)
        Sub SetRunStatus(runId As String, status As String, message As String, errorCode As String)
        Sub CompleteRun(runId As String, status As String, finalMessage As String, errorCode As String, finishedAt As DateTime)
    End Interface

    Public Class NoopRunTraceStore
        Implements IRunTraceStore

        Public Sub StartRun(runId As String, turn As UserTurn, startedAt As DateTime) Implements IRunTraceStore.StartRun
        End Sub

        Public Sub AppendStep(runId As String, seq As Integer, toolId As String, status As String, message As String, errorCode As String, observation As Object, startedAt As DateTime, finishedAt As DateTime) Implements IRunTraceStore.AppendStep
        End Sub

        Public Sub SetRunStatus(runId As String, status As String, message As String, errorCode As String) Implements IRunTraceStore.SetRunStatus
        End Sub

        Public Sub CompleteRun(runId As String, status As String, finalMessage As String, errorCode As String, finishedAt As DateTime) Implements IRunTraceStore.CompleteRun
        End Sub

    End Class

    Public Class SqliteRunTraceStore
        Implements IRunTraceStore

        Public Sub StartRun(runId As String, turn As UserTurn, startedAt As DateTime) Implements IRunTraceStore.StartRun
            If String.IsNullOrWhiteSpace(runId) Then Return
            OfficeAiDatabase.EnsureInitialized()

            Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
                conn.Open()
                Using cmd As New SQLiteCommand(
                    "INSERT OR REPLACE INTO agent_run " &
                    "(run_id, turn_id, session_id, app_type, status, user_text, started_at) " &
                    "VALUES (@run_id, @turn_id, @session_id, @app_type, @status, @user_text, @started_at)", conn)
                    cmd.Parameters.AddWithValue("@run_id", runId)
                    cmd.Parameters.AddWithValue("@turn_id", If(turn?.TurnId, ""))
                    cmd.Parameters.AddWithValue("@session_id", If(turn?.SessionId, ""))
                    cmd.Parameters.AddWithValue("@app_type", AppTypeContract.Normalize(turn?.AppType))
                    cmd.Parameters.AddWithValue("@status", "running")
                    cmd.Parameters.AddWithValue("@user_text", DbText(turn?.Text))
                    cmd.Parameters.AddWithValue("@started_at", ToDbTime(startedAt))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub AppendStep(runId As String,
                              seq As Integer,
                              toolId As String,
                              status As String,
                              message As String,
                              errorCode As String,
                              observation As Object,
                              startedAt As DateTime,
                              finishedAt As DateTime) Implements IRunTraceStore.AppendStep
            If String.IsNullOrWhiteSpace(runId) Then Return
            OfficeAiDatabase.EnsureInitialized()

            Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
                conn.Open()
                Using cmd As New SQLiteCommand(
                    "INSERT OR REPLACE INTO agent_run_step " &
                    "(step_id, run_id, seq, tool_id, status, message, error_code, observation_json, started_at, finished_at) " &
                    "VALUES (@step_id, @run_id, @seq, @tool_id, @status, @message, @error_code, @observation_json, @started_at, @finished_at)", conn)
                    cmd.Parameters.AddWithValue("@step_id", $"{runId}:{seq}:{If(toolId, "")}")
                    cmd.Parameters.AddWithValue("@run_id", runId)
                    cmd.Parameters.AddWithValue("@seq", seq)
                    cmd.Parameters.AddWithValue("@tool_id", If(toolId, ""))
                    cmd.Parameters.AddWithValue("@status", If(status, ""))
                    cmd.Parameters.AddWithValue("@message", DbText(message))
                    cmd.Parameters.AddWithValue("@error_code", If(errorCode, ""))
                    cmd.Parameters.AddWithValue("@observation_json", DbText(SerializeObservation(observation)))
                    cmd.Parameters.AddWithValue("@started_at", ToDbTime(startedAt))
                    cmd.Parameters.AddWithValue("@finished_at", ToDbTime(finishedAt))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub CompleteRun(runId As String, status As String, finalMessage As String, errorCode As String, finishedAt As DateTime) Implements IRunTraceStore.CompleteRun
            If String.IsNullOrWhiteSpace(runId) Then Return
            OfficeAiDatabase.EnsureInitialized()

            Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
                conn.Open()
                Using cmd As New SQLiteCommand(
                    "UPDATE agent_run SET status = @status, finished_at = @finished_at, final_message = @final_message, error_code = @error_code WHERE run_id = @run_id", conn)
                    cmd.Parameters.AddWithValue("@run_id", runId)
                    cmd.Parameters.AddWithValue("@status", If(status, ""))
                    cmd.Parameters.AddWithValue("@finished_at", ToDbTime(finishedAt))
                    cmd.Parameters.AddWithValue("@final_message", DbText(finalMessage))
                    cmd.Parameters.AddWithValue("@error_code", If(errorCode, ""))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub SetRunStatus(runId As String, status As String, message As String, errorCode As String) Implements IRunTraceStore.SetRunStatus
            If String.IsNullOrWhiteSpace(runId) Then Return
            OfficeAiDatabase.EnsureInitialized()
            Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
                conn.Open()
                Using cmd As New SQLiteCommand(
                    "UPDATE agent_run SET status = @status, final_message = @message, error_code = @error_code WHERE run_id = @run_id", conn)
                    cmd.Parameters.AddWithValue("@run_id", runId)
                    cmd.Parameters.AddWithValue("@status", If(status, ""))
                    cmd.Parameters.AddWithValue("@message", DbText(message))
                    cmd.Parameters.AddWithValue("@error_code", If(errorCode, ""))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Shared Function DbText(value As String) As Object
            If String.IsNullOrEmpty(value) Then Return DBNull.Value
            Return value
        End Function

        Private Shared Function ToDbTime(value As DateTime) As String
            If value = DateTime.MinValue Then value = DateTime.Now
            Return value.ToString("yyyy-MM-dd HH:mm:ss.fff")
        End Function

        Private Shared Function SerializeObservation(value As Object) As String
            If value Is Nothing Then Return ""
            Try
                Return JsonConvert.SerializeObject(value)
            Catch
                Return value.ToString()
            End Try
        End Function
    End Class

End Namespace
