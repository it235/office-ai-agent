' ShareRibbon\Storage\ConversationRepository.vb
' 会话历史读写：conversation_event 是聊天历史的唯一权威源。

Imports System.Data.SQLite

''' <summary>
''' 单条会话消息 DTO（用于加载历史会话）
''' </summary>
Public Class ConversationMessageDto
    Public Property Role As String
    Public Property Content As String
    Public Property CreateTime As String
End Class

''' <summary>
''' 会话历史读写。历史上聊天消息曾分裂写入 conversation 与 conversation_event 两张表，
''' 现在统一以 conversation_event（event_type=chat_message）为唯一来源，历史加载与记忆
''' 抽取读同一批行，任一侧再也不会与另一侧分叉。
''' </summary>
Public Class ConversationRepository

    ''' <summary>
    ''' 追加一条聊天消息到权威历史表 conversation_event
    ''' </summary>
    Public Shared Sub InsertMessage(sessionId As String, role As String, content As String, Optional isCollected As Boolean = False)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO conversation_event (event_id, session_id, event_type, role, content, is_collected, created_at) " &
                "VALUES (@event_id, @sid, 'chat_message', @role, @content, @collected, datetime('now', 'localtime'))", conn)
                cmd.Parameters.AddWithValue("@event_id", Guid.NewGuid().ToString())
                cmd.Parameters.AddWithValue("@sid", sessionId)
                cmd.Parameters.AddWithValue("@role", role)
                cmd.Parameters.AddWithValue("@content", If(content, ""))
                cmd.Parameters.AddWithValue("@collected", If(isCollected, 1, 0))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 收藏/取消收藏会话内最近一条 assistant 消息
    ''' </summary>
    Public Shared Sub SetLastAssistantCollected(sessionId As String, isCollected As Boolean)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE conversation_event SET is_collected=@c WHERE event_id=" &
                "(SELECT event_id FROM conversation_event WHERE session_id=@sid AND role='assistant' AND event_type='chat_message' " &
                "ORDER BY created_at DESC LIMIT 1)", conn)
                cmd.Parameters.AddWithValue("@c", If(isCollected, 1, 0))
                cmd.Parameters.AddWithValue("@sid", sessionId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 按会话 ID 获取该会话下所有聊天消息（按 created_at 升序），用于加载历史会话到界面
    ''' </summary>
    Public Shared Function GetMessagesBySession(sessionId As String) As List(Of ConversationMessageDto)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of ConversationMessageDto)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT role, content, created_at FROM conversation_event " &
                "WHERE session_id=@sid AND event_type='chat_message' ORDER BY created_at ASC", conn)
                cmd.Parameters.AddWithValue("@sid", sessionId)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        list.Add(New ConversationMessageDto With {
                            .Role = rdr.GetString(0),
                            .Content = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
                            .CreateTime = If(rdr.IsDBNull(2), "", rdr.GetString(2))
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    ''' <summary>
    ''' Authoritative single-transaction cleanup of every persisted row owned by one session:
    ''' conversation events, legacy conversation rows, the session summary, run traces, legacy
    ''' atomic memories, plus the memory items derived from the session's events and their
    ''' embeddings and pending jobs. Deleting a session must not leave orphans behind.
    ''' </summary>
    Public Shared Function DeleteSessionCascade(sessionId As String) As Integer
        If String.IsNullOrWhiteSpace(sessionId) Then Return 0
        OfficeAiDatabase.EnsureInitialized()

        Dim removed As Integer = 0
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using tx = conn.BeginTransaction()
                ' Order: leaf evidence (embeddings) -> jobs -> derived memory -> events -> session-level rows.
                Dim statements As String() = {
                    "DELETE FROM memory_embedding WHERE memory_id IN (" &
                    "  SELECT memory_id FROM memory_item WHERE source_event_id IN (" &
                    "    SELECT event_id FROM conversation_event WHERE session_id = @sid))",
                    "DELETE FROM memory_job WHERE target_id = @sid OR target_id IN (" &
                    "  SELECT event_id FROM conversation_event WHERE session_id = @sid) OR target_id IN (" &
                    "  SELECT memory_id FROM memory_item WHERE source_event_id IN (" &
                    "    SELECT event_id FROM conversation_event WHERE session_id = @sid))",
                    "DELETE FROM memory_item WHERE source_event_id IN (" &
                    "  SELECT event_id FROM conversation_event WHERE session_id = @sid)",
                    "DELETE FROM conversation_event WHERE session_id = @sid",
                    "DELETE FROM conversation WHERE session_id = @sid",
                    "DELETE FROM session_summary WHERE session_id = @sid",
                    "DELETE FROM agent_run WHERE session_id = @sid",
                    "DELETE FROM atomic_memory WHERE session_id = @sid"
                }
                For Each stmt In statements
                    Using cmd As New SQLiteCommand(stmt, conn)
                        cmd.Transaction = tx
                        cmd.Parameters.AddWithValue("@sid", sessionId)
                        removed += cmd.ExecuteNonQuery()
                    End Using
                Next
                tx.Commit()
            End Using
        End Using
        Return removed
    End Function
End Class
