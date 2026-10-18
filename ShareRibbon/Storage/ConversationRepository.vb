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
    ''' 会话列表：直接从权威历史表按会话聚合，标题取首条用户消息、摘要取末条消息。
    ''' 曾由普通聊天路径单独写 session_summary，Agent 主路径的会话因此不在列表里；
    ''' 现在无论走哪条链路，只要有消息就能出现在列表中。
    ''' </summary>
    Public Shared Function GetRecentSessions(limit As Integer) As List(Of SessionSummaryRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of SessionSummaryRecord)()
        If limit <= 0 Then Return list

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT e.session_id, MIN(e.created_at), " &
                "  (SELECT f.content FROM conversation_event f WHERE f.session_id = e.session_id " &
                "   AND f.event_type = 'chat_message' AND f.role = 'user' ORDER BY f.created_at, f.rowid LIMIT 1), " &
                "  (SELECT g.content FROM conversation_event g WHERE g.session_id = e.session_id " &
                "   AND g.event_type = 'chat_message' ORDER BY g.created_at DESC, g.rowid DESC LIMIT 1) " &
                "FROM conversation_event e " &
                "WHERE e.event_type = 'chat_message' AND IFNULL(e.session_id, '') <> '' " &
                "GROUP BY e.session_id ORDER BY MAX(e.created_at) DESC LIMIT @limit", conn)
                cmd.Parameters.AddWithValue("@limit", limit)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        list.Add(New SessionSummaryRecord With {
                            .Id = 0,
                            .SessionId = rdr.GetString(0),
                            .Title = Truncate(ReadText(rdr, 2), 80),
                            .Snippet = Truncate(ReadText(rdr, 3), 200),
                            .CreatedAt = ReadText(rdr, 1)
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    Private Shared Function ReadText(rdr As SQLiteDataReader, ordinal As Integer) As String
        Return If(rdr.IsDBNull(ordinal), "", rdr.GetString(ordinal))
    End Function

    Private Shared Function Truncate(value As String, maxLen As Integer) As String
        Dim text = If(value, "").Replace(vbCr, " "c).Replace(vbLf, " "c)
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "..."
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
                ' Order: leaf evidence (embeddings) -> jobs -> session-scoped memory -> run steps
                ' -> events -> session-level rows.
                ' 只删本会话独占的 session 作用域记忆：共享记忆（user/document/project）是跨会话沉淀，
                ' 删一个会话不得连带删掉别的会话积累的记忆。
                Dim sessionEvents As String =
                    "(SELECT event_id FROM conversation_event WHERE session_id = @sid)"
                Dim sessionMemories As String =
                    $"(SELECT memory_id FROM memory_item WHERE scope = 'session' AND source_event_id IN {sessionEvents})"
                ' 共享记忆只解绑不删除：source_event_id 是指向 conversation_event 的外键，
                ' 置空串会当场触发约束，NULL 才是合法的"无锚点"表示。
                Dim statements As String() = {
                    "DELETE FROM memory_embedding WHERE memory_id IN " & sessionMemories,
                    "DELETE FROM memory_job WHERE target_id = @sid OR target_id IN " & sessionEvents &
                    " OR target_id IN " & sessionMemories,
                    "DELETE FROM memory_item WHERE scope = 'session' AND source_event_id IN " & sessionEvents,
                    "UPDATE memory_item SET source_event_id = NULL WHERE source_event_id IN " & sessionEvents,
                    "DELETE FROM agent_run_step WHERE run_id IN (SELECT run_id FROM agent_run WHERE session_id = @sid)",
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
