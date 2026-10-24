' ShareRibbon\Storage\ConversationRepository.vb
' 会话历史读写：conversation_event 是聊天历史的唯一权威源。

Imports System.Text.RegularExpressions
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
    ''' 追加一条聊天消息到权威历史表 conversation_event，并把该会话登记进 session_summary 档案。
    ''' app_type 决定该会话归属于哪个宿主，document_id 决定它归属于哪一篇文档，
    ''' 历史侧边栏据此隔离 Word/Excel/PowerPoint 并优先显示当前文档的会话。
    ''' </summary>
    Public Shared Sub InsertMessage(sessionId As String, role As String, content As String, Optional isCollected As Boolean = False, Optional appType As String = "", Optional documentId As String = "", Optional documentName As String = "")
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO conversation_event (event_id, session_id, app_type, document_id, event_type, role, content, is_collected, created_at) " &
                "VALUES (@event_id, @sid, @app, @doc, 'chat_message', @role, @content, @collected, datetime('now', 'localtime'))", conn)
                cmd.Parameters.AddWithValue("@event_id", Guid.NewGuid().ToString())
                cmd.Parameters.AddWithValue("@sid", sessionId)
                cmd.Parameters.AddWithValue("@app", If(appType, ""))
                cmd.Parameters.AddWithValue("@doc", If(documentId, ""))
                cmd.Parameters.AddWithValue("@role", role)
                cmd.Parameters.AddWithValue("@content", If(content, ""))
                cmd.Parameters.AddWithValue("@collected", If(isCollected, 1, 0))
                cmd.ExecuteNonQuery()
            End Using
            UpsertSessionProfile(conn, sessionId, If(appType, ""), If(documentId, ""), If(documentName, ""))
        End Using
    End Sub

    ''' <summary>
    ''' 会话档案一行一个会话。文档绑定采用"先写者胜"：会话属于创建它的那篇文档，
    ''' 之后用户切换活动文档不得把已有会话悄悄搬走，否则历史归属变得不可解释。
    ''' </summary>
    Private Shared Sub UpsertSessionProfile(conn As SQLiteConnection, sessionId As String, appType As String, documentId As String, documentName As String)
        If String.IsNullOrWhiteSpace(sessionId) Then Return
        Using cmd As New SQLiteCommand(
            "UPDATE session_summary SET " &
            "  app_type = CASE WHEN IFNULL(app_type, '') = '' THEN @app ELSE app_type END, " &
            "  document_id = CASE WHEN IFNULL(document_id, '') = '' THEN @doc ELSE document_id END, " &
            "  document_name = CASE WHEN IFNULL(document_name, '') = '' THEN @docname ELSE document_name END, " &
            "  updated_at = datetime('now', 'localtime') " &
            "WHERE session_id = @sid", conn)
            cmd.Parameters.AddWithValue("@app", appType)
            cmd.Parameters.AddWithValue("@doc", documentId)
            cmd.Parameters.AddWithValue("@docname", documentName)
            cmd.Parameters.AddWithValue("@sid", sessionId)
            If cmd.ExecuteNonQuery() > 0 Then Return
        End Using

        Using cmd As New SQLiteCommand(
            "INSERT INTO session_summary (session_id, app_type, document_id, document_name, created_at, updated_at) " &
            "VALUES (@sid, @app, @doc, @docname, datetime('now', 'localtime'), datetime('now', 'localtime'))", conn)
            cmd.Parameters.AddWithValue("@app", appType)
            cmd.Parameters.AddWithValue("@doc", documentId)
            cmd.Parameters.AddWithValue("@docname", documentName)
            cmd.Parameters.AddWithValue("@sid", sessionId)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' 用户重命名会话：自定义标题覆盖"首条用户消息"自动标题。
    ''' 传空标题表示放弃自定义、回到自动标题。老会话可能还没有档案行，这里补一行。
    ''' </summary>
    Public Shared Function RenameSession(sessionId As String, title As String) As Boolean
        If String.IsNullOrWhiteSpace(sessionId) Then Return False
        OfficeAiDatabase.EnsureInitialized()

        Dim custom = If(title, "").Trim()
        If custom.Length > MaxTitleLength Then custom = custom.Substring(0, MaxTitleLength)

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE session_summary SET custom_title = @title, updated_at = datetime('now', 'localtime') WHERE session_id = @sid", conn)
                cmd.Parameters.AddWithValue("@title", If(String.IsNullOrEmpty(custom), CType(DBNull.Value, Object), CType(custom, Object)))
                cmd.Parameters.AddWithValue("@sid", sessionId)
                If cmd.ExecuteNonQuery() > 0 Then Return True
            End Using

            Using cmd As New SQLiteCommand(
                "INSERT INTO session_summary (session_id, custom_title, created_at, updated_at) " &
                "SELECT @sid, @title, MIN(e.created_at), datetime('now', 'localtime') FROM conversation_event e " &
                "WHERE e.session_id = @sid AND e.event_type = 'chat_message' " &
                "HAVING COUNT(*) > 0", conn)
                cmd.Parameters.AddWithValue("@title", If(String.IsNullOrEmpty(custom), CType(DBNull.Value, Object), CType(custom, Object)))
                cmd.Parameters.AddWithValue("@sid", sessionId)
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    End Function

    Private Const MaxTitleLength As Integer = 60

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
    ''' 会话列表：直接从权威历史表按会话聚合，标题优先用用户自定义，否则取首条用户消息，摘要取末条消息。
    ''' 曾由普通聊天路径单独写 session_summary，Agent 主路径的会话因此不在列表里；
    ''' 现在无论走哪条链路，只要有消息就能出现在列表中。
    ''' appType 非空时只返回当前宿主创建的会话；升级前的历史行 app_type 为空，仍保留可见。
    ''' scope="document" 只返回当前文档的会话；"all" 返回全部但把当前文档排在最前。
    ''' </summary>
    Public Shared Function GetRecentSessions(limit As Integer,
                                            Optional appType As String = Nothing,
                                            Optional documentId As String = Nothing,
                                            Optional scope As String = SessionScopeAll) As List(Of SessionSummaryRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of SessionSummaryRecord)()
        If limit <= 0 Then Return list
        Dim appFilter = If(appType, "").Trim()
        Dim docFilter = If(documentId, "").Trim()
        Dim docOnly = If(String.Equals(If(scope, SessionScopeAll), SessionScopeDocument, StringComparison.OrdinalIgnoreCase), 1, 0)

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT e.session_id, MIN(e.created_at), " &
                "  (SELECT f.content FROM conversation_event f WHERE f.session_id = e.session_id " &
                "   AND f.event_type = 'chat_message' AND f.role = 'user' ORDER BY f.created_at, f.rowid LIMIT 1), " &
                "  (SELECT g.content FROM conversation_event g WHERE g.session_id = e.session_id " &
                "   AND g.event_type = 'chat_message' ORDER BY g.created_at DESC, g.rowid DESC LIMIT 1), " &
                "  IFNULL(p.custom_title, ''), IFNULL(p.document_id, ''), IFNULL(p.document_name, '') " &
                "FROM conversation_event e " &
                "LEFT JOIN session_summary p ON p.session_id = e.session_id " &
                "WHERE e.event_type = 'chat_message' AND IFNULL(e.session_id, '') <> '' " &
                "  AND (@app = '' OR IFNULL(e.app_type, '') = '' OR e.app_type = @app) " &
                "  AND (@doc_only = 0 OR IFNULL(p.document_id, '') = @doc) " &
                "GROUP BY e.session_id " &
                "ORDER BY (IFNULL(MAX(p.document_id), '') = @doc) DESC, MAX(e.created_at) DESC LIMIT @limit", conn)
                cmd.Parameters.AddWithValue("@app", appFilter)
                cmd.Parameters.AddWithValue("@doc", docFilter)
                cmd.Parameters.AddWithValue("@doc_only", docOnly)
                cmd.Parameters.AddWithValue("@limit", limit)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim customTitle = Truncate(ReadText(rdr, 4), MaxTitleLength)
                        Dim derivedTitle = AutoTitle(ReadText(rdr, 2))
                        list.Add(New SessionSummaryRecord With {
                            .Id = 0,
                            .SessionId = rdr.GetString(0),
                            .Title = If(String.IsNullOrEmpty(customTitle), derivedTitle, customTitle),
                            .Snippet = Truncate(ReadText(rdr, 3), 200),
                            .CreatedAt = ReadText(rdr, 1),
                            .DocumentId = ReadText(rdr, 5),
                            .DocumentName = ReadText(rdr, 6),
                            .IsRenamed = Not String.IsNullOrEmpty(customTitle)
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    Public Const SessionScopeAll As String = "all"
    Public Const SessionScopeDocument As String = "document"

    Private Shared Function ReadText(rdr As SQLiteDataReader, ordinal As Integer) As String
        Return If(rdr.IsDBNull(ordinal), "", rdr.GetString(ordinal))
    End Function

    Private Shared Function Truncate(value As String, maxLen As Integer) As String
        Dim text = If(value, "").Replace(vbCr, " "c).Replace(vbLf, " "c)
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "..."
    End Function

    Private Const AutoTitleMin As Integer = 8
    Private Const AutoTitleMax As Integer = 30
    Private Const AutoTitleSearchWindow As Integer = 40

    ''' <summary>
    ''' 自动标题：折叠空白、优先在句末标点处收尾，其次在词边界截断，让列表里是一条读得懂的任务名
    ''' 而不是硬切开的半句话。不调模型生成标题：为侧边栏多花一次请求不划算，成本闸门（E15）还没落地。
    ''' </summary>
    Private Shared Function AutoTitle(raw As String) As String
        Dim text = Regex.Replace(If(raw, ""), "\s+", " ").Trim()
        If text.Length <= AutoTitleMax Then Return text

        Dim probeEnd = Math.Min(text.Length, AutoTitleSearchWindow)
        For i As Integer = AutoTitleMin To probeEnd - 1
            Dim ch = text(i)
            ' 句中的小数点不该被当成句号："把 3.5 倍预算的表重排" 要留完整语义
            Dim isBreak = (ch = "."c AndAlso (i = 0 OrElse Not Char.IsDigit(text(i - 1)))) OrElse
                          ch = "。"c OrElse ch = "！"c OrElse ch = "？"c OrElse
                          ch = "!"c OrElse ch = "?"c OrElse ch = "；"c OrElse ch = ";"c
            If isBreak Then Return text.Substring(0, i + 1).Trim()
        Next

        Dim lastSpace = text.LastIndexOf(" "c, AutoTitleMax)
        If lastSpace >= AutoTitleMin Then Return text.Substring(0, lastSpace).Trim() & "…"
        Return text.Substring(0, AutoTitleMax).Trim() & "…"
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
                    "DELETE FROM agent_run_change WHERE run_id IN (SELECT run_id FROM agent_run WHERE session_id = @sid)",
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
