' ShareRibbon\Storage\AgentMemoryRepository.vb
' Sidecar memory/event/job repository for the Office AI Agent memory pipeline.

Imports System.Data.SQLite
Imports System.Collections.Generic

Public Class AgentMemoryRepository
    Implements IMemoryEventStore

    ''' <summary>
    ''' 单个任务最多被领取执行几次；用尽后落 dead-letter，不再无限重投。
    ''' </summary>
    Public Const MaxJobAttempts As Integer = 3

    ''' <summary>
    ''' 失败重试退避的封顶分钟数，避免长尾退避把队列拖死。
    ''' </summary>
    Private Const JobBackoffCapMinutes As Integer = 30

    Private Shared Function NewId() As String
        Return Guid.NewGuid().ToString("D")
    End Function

    Private Shared Function DbText(value As String) As Object
        If String.IsNullOrEmpty(value) Then Return DBNull.Value
        Return value
    End Function

    Private Shared Function DbRequiredText(value As String, fallback As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return fallback
        Return value.Trim()
    End Function

    Private Shared Function ReadConversationEvent(rdr As SQLiteDataReader) As ConversationEventRecord
        Return New ConversationEventRecord With {
            .EventId = If(rdr.IsDBNull(0), "", rdr.GetString(0)),
            .SessionId = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
            .AppType = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
            .DocumentId = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
            .EventType = If(rdr.IsDBNull(4), "", rdr.GetString(4)),
            .Role = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
            .Content = If(rdr.IsDBNull(6), "", rdr.GetString(6)),
            .MetadataJson = If(rdr.IsDBNull(7), "", rdr.GetString(7)),
            .CreatedAt = If(rdr.IsDBNull(8), "", rdr.GetString(8)),
            .ProcessedAt = If(rdr.IsDBNull(9), "", rdr.GetString(9))
        }
    End Function

    Private Shared Function ReadMemoryItem(rdr As SQLiteDataReader) As MemoryItemRecord
        Return New MemoryItemRecord With {
            .MemoryId = If(rdr.IsDBNull(0), "", rdr.GetString(0)),
            .SourceEventId = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
            .Scope = If(rdr.IsDBNull(2), "user", rdr.GetString(2)),
            .AppType = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
            .DocumentId = If(rdr.IsDBNull(4), "", rdr.GetString(4)),
            .ProjectId = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
            .MemoryType = If(rdr.IsDBNull(6), "fact", rdr.GetString(6)),
            .Content = If(rdr.IsDBNull(7), "", rdr.GetString(7)),
            .Summary = If(rdr.IsDBNull(8), "", rdr.GetString(8)),
            .Confidence = If(rdr.IsDBNull(9), 0.5R, rdr.GetDouble(9)),
            .Importance = If(rdr.IsDBNull(10), 0.5R, rdr.GetDouble(10)),
            .Status = If(rdr.IsDBNull(11), "active", rdr.GetString(11)),
            .ExpiresAt = If(rdr.IsDBNull(12), "", rdr.GetString(12)),
            .LastVerifiedAt = If(rdr.IsDBNull(13), "", rdr.GetString(13)),
            .CreatedAt = If(rdr.IsDBNull(14), "", rdr.GetString(14)),
            .UpdatedAt = If(rdr.IsDBNull(15), "", rdr.GetString(15))
        }
    End Function

    Private Shared Function ReadMemoryJob(rdr As SQLiteDataReader) As MemoryJobRecord
        Return New MemoryJobRecord With {
            .JobId = If(rdr.IsDBNull(0), "", rdr.GetString(0)),
            .JobType = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
            .TargetId = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
            .PayloadJson = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
            .Status = If(rdr.IsDBNull(4), "pending", rdr.GetString(4)),
            .AttemptCount = If(rdr.IsDBNull(5), 0, rdr.GetInt32(5)),
            .LastError = If(rdr.IsDBNull(6), "", rdr.GetString(6)),
            .CreatedAt = If(rdr.IsDBNull(7), "", rdr.GetString(7)),
            .UpdatedAt = If(rdr.IsDBNull(8), "", rdr.GetString(8)),
            .NextRunAt = If(rdr.IsDBNull(9), "", rdr.GetString(9))
        }
    End Function

    Private Shared Function ReadSkillRegistry(rdr As SQLiteDataReader) As SkillRegistryRecord
        Return New SkillRegistryRecord With {
            .SkillName = If(rdr.IsDBNull(0), "", rdr.GetString(0)),
            .FilePath = If(rdr.IsDBNull(1), "", rdr.GetString(1)),
            .AppScope = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
            .IntentTypes = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
            .TriggerKeywords = If(rdr.IsDBNull(4), "", rdr.GetString(4)),
            .Description = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
            .EmbeddingJson = If(rdr.IsDBNull(6), "", rdr.GetString(6)),
            .UsageCount = If(rdr.IsDBNull(7), 0, rdr.GetInt32(7)),
            .SuccessCount = If(rdr.IsDBNull(8), 0, rdr.GetInt32(8)),
            .LastIndexedAt = If(rdr.IsDBNull(9), "", rdr.GetString(9)),
            .Enabled = Not rdr.IsDBNull(10) AndAlso rdr.GetInt32(10) <> 0
        }
    End Function

    Private Shared Function NormalizeComparableText(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        Dim chars As New List(Of Char)()
        For Each ch In value.Trim().ToLowerInvariant()
            If Not Char.IsWhiteSpace(ch) Then
                chars.Add(ch)
            End If
        Next
        Return New String(chars.ToArray())
    End Function

    Private Shared Function GetSearchTerms(query As String) As List(Of String)
        Dim terms As New List(Of String)()
        If String.IsNullOrWhiteSpace(query) Then Return terms

        Dim full = query.Trim()
        terms.Add(full)

        For Each part In full.Split({" "c, vbTab, ","c, "，"c, ";"c, "；"c, "."c, "。"c, "/"c, "\"c, "|"c, ":"c, "："c}, StringSplitOptions.RemoveEmptyEntries)
            Dim term = part.Trim()
            If term.Length >= 2 AndAlso Not terms.Any(Function(t) String.Equals(t, term, StringComparison.OrdinalIgnoreCase)) Then
                terms.Add(term)
            End If
        Next

        Return terms.Take(8).ToList()
    End Function

    Public Shared Function AppendConversationEvent(record As ConversationEventRecord) As String
        OfficeAiDatabase.EnsureInitialized()
        If record Is Nothing Then Throw New ArgumentNullException(NameOf(record))

        Dim eventId = DbRequiredText(record.EventId, NewId())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO conversation_event " &
                "(event_id, session_id, app_type, document_id, event_type, role, content, metadata_json, processed_at) " &
                "VALUES (@event_id, @session_id, @app_type, @document_id, @event_type, @role, @content, @metadata_json, @processed_at) " &
                "ON CONFLICT(event_id) DO UPDATE SET " &
                "session_id = excluded.session_id, app_type = excluded.app_type, document_id = excluded.document_id, " &
                "event_type = excluded.event_type, role = excluded.role, content = excluded.content, metadata_json = excluded.metadata_json, " &
                "processed_at = excluded.processed_at", conn)
                cmd.Parameters.AddWithValue("@event_id", eventId)
                cmd.Parameters.AddWithValue("@session_id", DbRequiredText(record.SessionId, "default"))
                cmd.Parameters.AddWithValue("@app_type", If(record.AppType, ""))
                cmd.Parameters.AddWithValue("@document_id", If(record.DocumentId, ""))
                cmd.Parameters.AddWithValue("@event_type", DbRequiredText(record.EventType, "message"))
                cmd.Parameters.AddWithValue("@role", If(record.Role, ""))
                cmd.Parameters.AddWithValue("@content", DbText(record.Content))
                cmd.Parameters.AddWithValue("@metadata_json", DbText(record.MetadataJson))
                cmd.Parameters.AddWithValue("@processed_at", DbText(record.ProcessedAt))
                cmd.ExecuteNonQuery()
            End Using
        End Using
        Return eventId
    End Function

    Public Function InsertConversationEvent(record As ConversationEventRecord) As String Implements IMemoryEventStore.InsertConversationEvent
        Return AppendConversationEvent(record)
    End Function

    Public Shared Function GetConversationEventById(eventId As String) As ConversationEventRecord
        If String.IsNullOrWhiteSpace(eventId) Then Return Nothing
        OfficeAiDatabase.EnsureInitialized()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT event_id, session_id, app_type, document_id, event_type, role, content, metadata_json, created_at, processed_at " &
                "FROM conversation_event WHERE event_id = @event_id LIMIT 1", conn)
                cmd.Parameters.AddWithValue("@event_id", eventId)
                Using rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        Return ReadConversationEvent(rdr)
                    End If
                End Using
            End Using
        End Using

        Return Nothing
    End Function

    ''' <summary>
    ''' Returns this session's chat_message events that extraction has not processed yet, oldest
    ''' first. With ChatStateService as the single writer of chat history, extraction is driven by
    ''' this session backlog instead of the event ids the recorder used to stamp into each job.
    ''' </summary>
    Public Shared Function GetUnprocessedConversationEventsBySession(sessionId As String, Optional limit As Integer = 40) As List(Of ConversationEventRecord)
        Dim results As New List(Of ConversationEventRecord)()
        If String.IsNullOrWhiteSpace(sessionId) Then Return results
        OfficeAiDatabase.EnsureInitialized()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT event_id, session_id, app_type, document_id, event_type, role, content, metadata_json, created_at, processed_at " &
                "FROM conversation_event WHERE session_id = @sid AND event_type = 'chat_message' AND processed_at IS NULL " &
                "ORDER BY created_at ASC LIMIT @limit", conn)
                cmd.Parameters.AddWithValue("@sid", sessionId)
                cmd.Parameters.AddWithValue("@limit", Math.Max(1, limit))
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        results.Add(ReadConversationEvent(rdr))
                    End While
                End Using
            End Using
        End Using

        Return results
    End Function

    Public Shared Sub MarkConversationEventProcessed(eventId As String)
        If String.IsNullOrWhiteSpace(eventId) Then Return
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("UPDATE conversation_event SET processed_at = datetime('now', 'localtime') WHERE event_id = @event_id", conn)
                cmd.Parameters.AddWithValue("@event_id", eventId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Shared Function UpsertMemoryItem(record As MemoryItemRecord) As String
        OfficeAiDatabase.EnsureInitialized()
        If record Is Nothing Then Throw New ArgumentNullException(NameOf(record))

        Dim memoryId = DbRequiredText(record.MemoryId, NewId())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO memory_item " &
                "(memory_id, source_event_id, scope, app_type, document_id, project_id, memory_type, content, summary, confidence, importance, status, expires_at, last_verified_at) " &
                "VALUES (@memory_id, @source_event_id, @scope, @app_type, @document_id, @project_id, @memory_type, @content, @summary, @confidence, @importance, @status, @expires_at, @last_verified_at) " &
                "ON CONFLICT(memory_id) DO UPDATE SET " &
                "source_event_id = CASE WHEN IFNULL(memory_item.source_event_id, '') <> '' " &
                "  THEN memory_item.source_event_id ELSE excluded.source_event_id END, " &
                "scope = excluded.scope, app_type = excluded.app_type, document_id = excluded.document_id, " &
                "project_id = excluded.project_id, memory_type = excluded.memory_type, content = excluded.content, summary = excluded.summary, " &
                "confidence = excluded.confidence, importance = excluded.importance, status = excluded.status, expires_at = excluded.expires_at, " &
                "last_verified_at = excluded.last_verified_at, updated_at = datetime('now', 'localtime')", conn)
                cmd.Parameters.AddWithValue("@memory_id", memoryId)
                cmd.Parameters.AddWithValue("@source_event_id", DbText(record.SourceEventId))
                cmd.Parameters.AddWithValue("@scope", DbRequiredText(record.Scope, "user"))
                cmd.Parameters.AddWithValue("@app_type", If(record.AppType, ""))
                cmd.Parameters.AddWithValue("@document_id", If(record.DocumentId, ""))
                cmd.Parameters.AddWithValue("@project_id", If(record.ProjectId, ""))
                cmd.Parameters.AddWithValue("@memory_type", DbRequiredText(record.MemoryType, "fact"))
                cmd.Parameters.AddWithValue("@content", DbRequiredText(record.Content, ""))
                cmd.Parameters.AddWithValue("@summary", DbText(record.Summary))
                cmd.Parameters.AddWithValue("@confidence", record.Confidence)
                cmd.Parameters.AddWithValue("@importance", record.Importance)
                cmd.Parameters.AddWithValue("@status", DbRequiredText(record.Status, "active"))
                cmd.Parameters.AddWithValue("@expires_at", DbText(record.ExpiresAt))
                cmd.Parameters.AddWithValue("@last_verified_at", DbText(record.LastVerifiedAt))
                cmd.ExecuteNonQuery()
            End Using
        End Using
        Return memoryId
    End Function

    Public Shared Function UpsertMemoryItemMerged(record As MemoryItemRecord) As String
        If record Is Nothing Then Throw New ArgumentNullException(NameOf(record))

        Dim duplicate = FindDuplicateMemoryItem(record)
        If duplicate IsNot Nothing Then
            Dim oldContent = If(duplicate.Content, "")
            record.MemoryId = duplicate.MemoryId
            ' 证据锚点遵循先到原则：合并命中已有记忆时沿用旧锚点，避免记忆被后来者改绑、
            ' 从而在原始会话删除时被误判定为本会话产物。
            If Not String.IsNullOrWhiteSpace(duplicate.SourceEventId) Then
                record.SourceEventId = duplicate.SourceEventId
            End If
            record.Content = If(If(record.Content, "").Length >= oldContent.Length, record.Content, duplicate.Content)
            record.Summary = If(String.IsNullOrWhiteSpace(record.Summary), duplicate.Summary, record.Summary)
            record.Confidence = Math.Max(record.Confidence, duplicate.Confidence)
            record.Importance = Math.Max(record.Importance, duplicate.Importance)
            record.Status = If(String.IsNullOrWhiteSpace(record.Status), duplicate.Status, record.Status)
            record.LastVerifiedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            Dim memoryId = UpsertMemoryItem(record)
            If Not String.Equals(oldContent, If(record.Content, ""), StringComparison.Ordinal) Then
                MarkMemoryEmbeddingDirty(memoryId)
            End If
            Return memoryId
        End If

        Return UpsertMemoryItem(record)
    End Function

    Public Shared Function FindDuplicateMemoryItem(record As MemoryItemRecord) As MemoryItemRecord
        If record Is Nothing Then Return Nothing

        Dim contentKey = NormalizeComparableText(record.Content)
        Dim summaryKey = NormalizeComparableText(record.Summary)
        If String.IsNullOrWhiteSpace(contentKey) AndAlso String.IsNullOrWhiteSpace(summaryKey) Then Return Nothing

        OfficeAiDatabase.EnsureInitialized()
        Dim results As New List(Of MemoryItemRecord)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT memory_id, source_event_id, scope, app_type, document_id, project_id, memory_type, content, summary, confidence, importance, status, expires_at, last_verified_at, created_at, updated_at " &
                "FROM memory_item WHERE status = 'active' " &
                "AND scope = @scope AND memory_type = @memory_type " &
                "AND (app_type = @app_type OR app_type IS NULL OR app_type = '') " &
                "AND (document_id = @document_id OR document_id IS NULL OR document_id = '') " &
                "ORDER BY updated_at DESC LIMIT 50", conn)
                cmd.Parameters.AddWithValue("@scope", DbRequiredText(record.Scope, "user"))
                cmd.Parameters.AddWithValue("@memory_type", DbRequiredText(record.MemoryType, "fact"))
                cmd.Parameters.AddWithValue("@app_type", If(record.AppType, ""))
                cmd.Parameters.AddWithValue("@document_id", If(record.DocumentId, ""))
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        results.Add(ReadMemoryItem(rdr))
                    End While
                End Using
            End Using
        End Using

        For Each candidate In results
            Dim candidateContentKey = NormalizeComparableText(candidate.Content)
            Dim candidateSummaryKey = NormalizeComparableText(candidate.Summary)
            If Not String.IsNullOrWhiteSpace(contentKey) AndAlso contentKey = candidateContentKey Then Return candidate
            If Not String.IsNullOrWhiteSpace(summaryKey) AndAlso summaryKey = candidateSummaryKey Then Return candidate
        Next

        Return Nothing
    End Function

    Public Shared Sub MarkMemoryEmbeddingDirty(memoryId As String)
        If String.IsNullOrWhiteSpace(memoryId) Then Return
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE memory_embedding SET vector_status = 'dirty', updated_at = datetime('now', 'localtime') WHERE memory_id = @memory_id", conn)
                cmd.Parameters.AddWithValue("@memory_id", memoryId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Shared Function UpsertMemoryEmbedding(record As MemoryEmbeddingRecord) As String
        OfficeAiDatabase.EnsureInitialized()
        If record Is Nothing Then Throw New ArgumentNullException(NameOf(record))

        Dim embeddingId = DbRequiredText(record.EmbeddingId, NewId())
        Dim memoryId = DbRequiredText(record.MemoryId, "")
        Dim embeddingModel = DbRequiredText(record.EmbeddingModel, EmbeddingService.GetConfiguredEmbeddingModelName())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO memory_embedding " &
                "(embedding_id, memory_id, embedding_model, embedding_dim, embedding_json, vector_status, last_error) " &
                "VALUES (@embedding_id, @memory_id, @embedding_model, @embedding_dim, @embedding_json, @vector_status, @last_error) " &
                "ON CONFLICT(memory_id, embedding_model) DO UPDATE SET " &
                "embedding_dim = excluded.embedding_dim, embedding_json = excluded.embedding_json, vector_status = excluded.vector_status, " &
                "last_error = excluded.last_error, updated_at = datetime('now', 'localtime')", conn)
                cmd.Parameters.AddWithValue("@embedding_id", embeddingId)
                cmd.Parameters.AddWithValue("@memory_id", memoryId)
                cmd.Parameters.AddWithValue("@embedding_model", embeddingModel)
                cmd.Parameters.AddWithValue("@embedding_dim", record.EmbeddingDim)
                cmd.Parameters.AddWithValue("@embedding_json", DbText(record.EmbeddingJson))
                cmd.Parameters.AddWithValue("@vector_status", DbRequiredText(record.VectorStatus, "pending"))
                cmd.Parameters.AddWithValue("@last_error", DbText(record.LastError))
                cmd.ExecuteNonQuery()
            End Using

            Using lookup As New SQLiteCommand("SELECT embedding_id FROM memory_embedding WHERE memory_id = @memory_id AND embedding_model = @embedding_model LIMIT 1", conn)
                lookup.Parameters.AddWithValue("@memory_id", memoryId)
                lookup.Parameters.AddWithValue("@embedding_model", embeddingModel)
                Dim obj = lookup.ExecuteScalar()
                If obj IsNot Nothing AndAlso obj IsNot DBNull.Value Then
                    Return obj.ToString()
                End If
            End Using
        End Using
        Return embeddingId
    End Function

    Public Shared Function RetrieveMemoryItems(query As String, appType As String, documentId As String, projectId As String, Optional topN As Integer = 8) As List(Of MemoryItemRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim results As New List(Of MemoryItemRecord)()
        Dim whereParts As New List(Of String) From {
            "status = 'active'",
            "(expires_at IS NULL OR expires_at = '' OR expires_at > datetime('now', 'localtime'))"
        }
        Dim searchTerms = GetSearchTerms(query)

        If Not String.IsNullOrWhiteSpace(appType) Then whereParts.Add("(app_type = @app_type OR app_type IS NULL OR app_type = '')")
        If Not String.IsNullOrWhiteSpace(documentId) Then whereParts.Add("(document_id = @document_id OR document_id IS NULL OR document_id = '')")
        If Not String.IsNullOrWhiteSpace(projectId) Then whereParts.Add("(project_id = @project_id OR project_id IS NULL OR project_id = '')")
        If searchTerms.Count > 0 Then
            Dim termParts As New List(Of String)()
            For i = 0 To searchTerms.Count - 1
                termParts.Add($"content LIKE @term{i} OR summary LIKE @term{i}")
            Next
            whereParts.Add("(" & String.Join(" OR ", termParts) & ")")
        End If

        Dim sql = "SELECT memory_id, source_event_id, scope, app_type, document_id, project_id, memory_type, content, summary, confidence, importance, status, expires_at, last_verified_at, created_at, updated_at " &
                  "FROM memory_item WHERE " & String.Join(" AND ", whereParts) &
                  " ORDER BY importance DESC, confidence DESC, updated_at DESC LIMIT @limit"

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(sql, conn)
                If Not String.IsNullOrWhiteSpace(appType) Then cmd.Parameters.AddWithValue("@app_type", appType.Trim())
                If Not String.IsNullOrWhiteSpace(documentId) Then cmd.Parameters.AddWithValue("@document_id", documentId.Trim())
                If Not String.IsNullOrWhiteSpace(projectId) Then cmd.Parameters.AddWithValue("@project_id", projectId.Trim())
                For i = 0 To searchTerms.Count - 1
                    cmd.Parameters.AddWithValue($"@term{i}", "%" & searchTerms(i) & "%")
                Next
                cmd.Parameters.AddWithValue("@limit", Math.Max(1, topN))
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        results.Add(ReadMemoryItem(rdr))
                    End While
                End Using
            End Using
        End Using
        Return results
    End Function

    Public Shared Function HasReadyMemoryEmbeddings(Optional appType As String = Nothing) As Boolean
        Try
            OfficeAiDatabase.EnsureInitialized()
            Dim app = If(appType, "").Trim()
            Dim sql = "SELECT COUNT(1) FROM memory_item m INNER JOIN memory_embedding e ON e.memory_id = m.memory_id " &
                      "WHERE m.status = 'active' AND e.vector_status = 'ready' AND e.embedding_json IS NOT NULL AND e.embedding_json != ''"
            If Not String.IsNullOrWhiteSpace(app) Then
                sql &= " AND (m.app_type = @app_type OR m.app_type IS NULL OR m.app_type = '')"
            End If

            Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
                conn.Open()
                Using cmd As New SQLiteCommand(sql, conn)
                    If Not String.IsNullOrWhiteSpace(app) Then cmd.Parameters.AddWithValue("@app_type", app)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[AgentMemoryRepository] HasReadyMemoryEmbeddings 失败: {ex.Message}")
            Return False
        End Try
    End Function

    Public Shared Function RetrieveMemoryItemsByVector(queryEmbedding As Single(), query As String, appType As String, documentId As String, projectId As String, Optional topN As Integer = 8) As List(Of MemoryItemRecord)
        Dim results As New List(Of MemoryItemRecord)()
        If queryEmbedding Is Nothing OrElse queryEmbedding.Length = 0 Then Return results

        OfficeAiDatabase.EnsureInitialized()
        MemoryRepository.EnsureVectorFunctionsRegistered()

        Dim queryEmbeddingJson = EmbeddingService.SerializeVector(queryEmbedding)
        If String.IsNullOrWhiteSpace(queryEmbeddingJson) Then Return results

        Dim whereParts As New List(Of String) From {
            "m.status = 'active'",
            "(m.expires_at IS NULL OR m.expires_at = '' OR m.expires_at > datetime('now', 'localtime'))",
            "e.vector_status = 'ready'",
            "e.embedding_json IS NOT NULL",
            "e.embedding_json != ''"
        }

        If Not String.IsNullOrWhiteSpace(appType) Then whereParts.Add("(m.app_type = @app_type OR m.app_type IS NULL OR m.app_type = '')")
        If Not String.IsNullOrWhiteSpace(documentId) Then whereParts.Add("(m.document_id = @document_id OR m.document_id IS NULL OR m.document_id = '')")
        If Not String.IsNullOrWhiteSpace(projectId) Then whereParts.Add("(m.project_id = @project_id OR m.project_id IS NULL OR m.project_id = '')")
        If Not String.IsNullOrWhiteSpace(query) Then whereParts.Add("(m.content LIKE @query OR m.summary LIKE @query OR cosine_similarity_json(e.embedding_json, @query_embedding) >= @threshold)")

        Dim fields = "m.memory_id, m.source_event_id, m.scope, m.app_type, m.document_id, m.project_id, m.memory_type, m.content, m.summary, m.confidence, m.importance, m.status, m.expires_at, m.last_verified_at, m.created_at, m.updated_at"
        Dim scoreExpr = "(cosine_similarity_json(e.embedding_json, @query_embedding) * (0.7 + COALESCE(m.importance, 0.5) * 0.3) * (0.7 + COALESCE(m.confidence, 0.5) * 0.3))"
        Dim sql = $"SELECT {fields}, {scoreExpr} AS final_score " &
                  "FROM memory_item m INNER JOIN memory_embedding e ON e.memory_id = m.memory_id " &
                  $"WHERE {String.Join(" AND ", whereParts)} " &
                  "ORDER BY final_score DESC, m.updated_at DESC LIMIT @limit"

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@query_embedding", queryEmbeddingJson)
                cmd.Parameters.AddWithValue("@threshold", MemoryConfig.RagSimilarityThreshold)
                cmd.Parameters.AddWithValue("@limit", Math.Max(1, topN))
                If Not String.IsNullOrWhiteSpace(appType) Then cmd.Parameters.AddWithValue("@app_type", appType.Trim())
                If Not String.IsNullOrWhiteSpace(documentId) Then cmd.Parameters.AddWithValue("@document_id", documentId.Trim())
                If Not String.IsNullOrWhiteSpace(projectId) Then cmd.Parameters.AddWithValue("@project_id", projectId.Trim())
                If Not String.IsNullOrWhiteSpace(query) Then cmd.Parameters.AddWithValue("@query", "%" & query.Trim() & "%")

                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim item = ReadMemoryItem(rdr)
                        item.Score = If(rdr.IsDBNull(16), 0.0R, rdr.GetDouble(16))
                        results.Add(item)
                    End While
                End Using
            End Using
        End Using

        Return results
    End Function

    Public Shared Function GetMemoryItemById(memoryId As String) As MemoryItemRecord
        If String.IsNullOrWhiteSpace(memoryId) Then Return Nothing
        OfficeAiDatabase.EnsureInitialized()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT memory_id, source_event_id, scope, app_type, document_id, project_id, memory_type, content, summary, confidence, importance, status, expires_at, last_verified_at, created_at, updated_at " &
                "FROM memory_item WHERE memory_id = @memory_id LIMIT 1", conn)
                cmd.Parameters.AddWithValue("@memory_id", memoryId)
                Using rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        Return ReadMemoryItem(rdr)
                    End If
                End Using
            End Using
        End Using

        Return Nothing
    End Function

    Public Shared Function EnqueueJob(record As MemoryJobRecord) As String
        OfficeAiDatabase.EnsureInitialized()
        If record Is Nothing Then Throw New ArgumentNullException(NameOf(record))

        Dim jobType = DbRequiredText(record.JobType, "extract_memory")
        Dim status = DbRequiredText(record.Status, "pending")
        Dim jobId = DbRequiredText(record.JobId, NewId())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            ' 幂等入队：(job_type, target_id) 上存在未终结任务时由部分唯一索引挡住重复投递，
            ' 这里回读那条在队任务并复用它的 id，同一批事件不会被抽取两次。
            Using cmd As New SQLiteCommand(
                "INSERT OR IGNORE INTO memory_job (job_id, job_type, target_id, payload_json, status, attempt_count, last_error, next_run_at) " &
                "VALUES (@job_id, @job_type, @target_id, @payload_json, @status, @attempt_count, @last_error, @next_run_at)", conn)
                cmd.Parameters.AddWithValue("@job_id", jobId)
                cmd.Parameters.AddWithValue("@job_type", jobType)
                cmd.Parameters.AddWithValue("@target_id", DbText(record.TargetId))
                cmd.Parameters.AddWithValue("@payload_json", DbText(record.PayloadJson))
                cmd.Parameters.AddWithValue("@status", status)
                cmd.Parameters.AddWithValue("@attempt_count", record.AttemptCount)
                cmd.Parameters.AddWithValue("@last_error", DbText(record.LastError))
                cmd.Parameters.AddWithValue("@next_run_at", DbText(record.NextRunAt))
                If cmd.ExecuteNonQuery() > 0 Then Return jobId
            End Using

            Using sel As New SQLiteCommand(
                "SELECT job_id FROM memory_job WHERE job_type = @job_type AND target_id = @target_id " &
                "AND status IN ('pending', 'processing') ORDER BY created_at ASC LIMIT 1", conn)
                sel.Parameters.AddWithValue("@job_type", jobType)
                sel.Parameters.AddWithValue("@target_id", DbText(record.TargetId))
                Dim existing = sel.ExecuteScalar()
                If existing IsNot Nothing Then Return CStr(existing)
            End Using
        End Using
        Return jobId
    End Function

    Public Function EnqueueMemoryJob(record As MemoryJobRecord) As String Implements IMemoryEventStore.EnqueueMemoryJob
        Return EnqueueJob(record)
    End Function

    Public Shared Function GetPendingMemoryJobs(Optional limit As Integer = 20) As List(Of MemoryJobRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim results As New List(Of MemoryJobRecord)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT job_id, job_type, target_id, payload_json, status, attempt_count, last_error, created_at, updated_at, next_run_at " &
                "FROM memory_job WHERE status = 'pending' AND (next_run_at IS NULL OR next_run_at = '' OR next_run_at <= datetime('now', 'localtime')) " &
                "ORDER BY created_at ASC LIMIT @limit", conn)
                cmd.Parameters.AddWithValue("@limit", Math.Max(1, limit))
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        results.Add(ReadMemoryJob(rdr))
                    End While
                End Using
            End Using
        End Using
        Return results
    End Function

    Public Shared Function EnqueuePendingEmbeddingJobs(Optional limit As Integer = 20) As Integer
        OfficeAiDatabase.EnsureInitialized()

        Dim model = EmbeddingService.GetConfiguredEmbeddingModelName()
        If String.IsNullOrWhiteSpace(model) Then Return 0

        Dim memoryIds As New List(Of String)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT m.memory_id " &
                "FROM memory_item m " &
                "LEFT JOIN memory_embedding e ON e.memory_id = m.memory_id AND e.embedding_model = @model " &
                "WHERE m.status = 'active' " &
                "AND (e.embedding_id IS NULL OR e.vector_status IN ('pending', 'failed', 'dirty')) " &
                "AND NOT EXISTS (" &
                "  SELECT 1 FROM memory_job j " &
                "  WHERE j.target_id = m.memory_id " &
                "  AND j.job_type IN ('embed_memory', 'rebuild_embedding') " &
                "  AND j.status IN ('pending', 'processing')" &
                ") " &
                "ORDER BY m.updated_at DESC LIMIT @limit", conn)
                cmd.Parameters.AddWithValue("@model", model)
                cmd.Parameters.AddWithValue("@limit", Math.Max(1, limit))
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        If Not rdr.IsDBNull(0) Then memoryIds.Add(rdr.GetString(0))
                    End While
                End Using
            End Using
        End Using

        Dim count As Integer = 0
        For Each memoryId In memoryIds
            Dim payload As New Newtonsoft.Json.Linq.JObject()
            payload("memory_id") = memoryId
            EnqueueJob(New MemoryJobRecord With {
                .JobType = "embed_memory",
                .TargetId = memoryId,
                .PayloadJson = payload.ToString(Newtonsoft.Json.Formatting.None),
                .Status = "pending"
            })
            count += 1
        Next

        Return count
    End Function

    Public Shared Sub MarkJobProcessing(jobId As String)
        UpdateJobStatus(jobId, "processing", Nothing)
    End Sub

    Public Shared Sub MarkJobCompleted(jobId As String)
        UpdateJobStatus(jobId, "completed", Nothing)
    End Sub

    ''' <summary>
    ''' 失败不再一次定终态：还有重试额度时带退避回到 pending，用尽才落 failed（dead-letter）。
    ''' attempt_count 在 ClaimPendingJobs 领取时已自增，这里只判断剩余额度。
    ''' 返回任务的新状态（"pending" 表示还会重试，"failed" 表示已放弃）。
    ''' </summary>
    Public Shared Function MarkJobFailed(jobId As String,
                                        errorMessage As String,
                                        Optional maxAttempts As Integer = MaxJobAttempts) As String
        If String.IsNullOrWhiteSpace(jobId) Then Return Nothing

        OfficeAiDatabase.EnsureInitialized()
        Dim cap = Math.Max(1, maxAttempts)
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using tx = conn.BeginTransaction()
                Dim attempts As Integer = 0
                Using sel As New SQLiteCommand("SELECT IFNULL(attempt_count, 0) FROM memory_job WHERE job_id = @job_id", conn)
                    sel.Transaction = tx
                    sel.Parameters.AddWithValue("@job_id", jobId)
                    Dim value = sel.ExecuteScalar()
                    If value IsNot Nothing Then attempts = Convert.ToInt32(value)
                End Using

                Dim exhausted = attempts >= cap
                Dim backoffMinutes = Math.Min(JobBackoffCapMinutes, 2 ^ Math.Max(0, attempts - 1))
                Using upd As New SQLiteCommand(
                    "UPDATE memory_job SET status = @status, last_error = @last_error, " &
                    "next_run_at = datetime('now', 'localtime', @window), " &
                    "updated_at = datetime('now', 'localtime') WHERE job_id = @job_id", conn)
                    upd.Transaction = tx
                    upd.Parameters.AddWithValue("@status", If(exhausted, "failed", "pending"))
                    upd.Parameters.AddWithValue("@last_error", DbText(errorMessage))
                    upd.Parameters.AddWithValue("@window", $"+{If(exhausted, 0, backoffMinutes)} minutes")
                    upd.Parameters.AddWithValue("@job_id", jobId)
                    upd.ExecuteNonQuery()
                End Using
                tx.Commit()
                Return If(exhausted, "failed", "pending")
            End Using
        End Using
    End Function

    ''' <summary>
    ''' 重试次数用尽、或压根跑不完的任务写回 failed，避免回收时反复重投。
    ''' </summary>
    Public Shared Function DeadLetterJob(jobId As String, errorMessage As String) As Boolean
        If String.IsNullOrWhiteSpace(jobId) Then Return False
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE memory_job SET status = 'failed', last_error = @last_error, " &
                "updated_at = datetime('now', 'localtime') WHERE job_id = @job_id", conn)
                cmd.Parameters.AddWithValue("@last_error", DbText(errorMessage))
                cmd.Parameters.AddWithValue("@job_id", jobId)
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    End Function

    Private Shared Sub UpdateJobStatus(jobId As String, status As String, errorMessage As String)
        If String.IsNullOrWhiteSpace(jobId) Then Return
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE memory_job SET status = @status, last_error = @last_error, updated_at = datetime('now', 'localtime') WHERE job_id = @job_id", conn)
                cmd.Parameters.AddWithValue("@status", status)
                cmd.Parameters.AddWithValue("@last_error", DbText(errorMessage))
                cmd.Parameters.AddWithValue("@job_id", jobId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Atomically claims up to <paramref name="limit"/> pending jobs in a single transaction:
    ''' selects candidate ids and flips them to 'processing' together, so a concurrent worker
    ''' (another host sharing the same SQLite file, or a re-entrant kick) cannot claim the same row.
    ''' Replaces the old SELECT-then-MarkJobProcessing split that left a claim window open.
    ''' </summary>
    Public Shared Function ClaimPendingJobs(Optional limit As Integer = 5) As List(Of MemoryJobRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim results As New List(Of MemoryJobRecord)()
        Dim max = Math.Max(1, limit)

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using tx = conn.BeginTransaction()
                Dim ids As New List(Of String)()
                Using cmd As New SQLiteCommand(
                    "SELECT job_id FROM memory_job WHERE status = 'pending' " &
                    "AND (next_run_at IS NULL OR next_run_at = '' OR next_run_at <= datetime('now', 'localtime')) " &
                    "ORDER BY created_at ASC LIMIT @limit", conn)
                    cmd.Transaction = tx
                    cmd.Parameters.AddWithValue("@limit", max)
                    Using rdr = cmd.ExecuteReader()
                        While rdr.Read()
                            If Not rdr.IsDBNull(0) Then ids.Add(rdr.GetString(0))
                        End While
                    End Using
                End Using

                If ids.Count = 0 Then
                    tx.Commit()
                    Return results
                End If

                Dim placeholders As New List(Of String)()
                For i = 0 To ids.Count - 1
                    placeholders.Add("@id" & i)
                Next
                Dim idList = String.Join(",", placeholders)

                Using upd As New SQLiteCommand(
                    "UPDATE memory_job SET status = 'processing', attempt_count = attempt_count + 1, " &
                    "updated_at = datetime('now', 'localtime') " &
                    "WHERE status = 'pending' AND job_id IN (" & idList & ")", conn)
                    upd.Transaction = tx
                    For i = 0 To ids.Count - 1
                        upd.Parameters.AddWithValue("@id" & i, ids(i))
                    Next
                    upd.ExecuteNonQuery()
                End Using

                Using sel As New SQLiteCommand(
                    "SELECT job_id, job_type, target_id, payload_json, status, attempt_count, last_error, created_at, updated_at, next_run_at " &
                    "FROM memory_job WHERE job_id IN (" & idList & ")", conn)
                    sel.Transaction = tx
                    For i = 0 To ids.Count - 1
                        sel.Parameters.AddWithValue("@id" & i, ids(i))
                    Next
                    Using rdr = sel.ExecuteReader()
                        While rdr.Read()
                            results.Add(ReadMemoryJob(rdr))
                        End While
                    End Using
                End Using

                tx.Commit()
            End Using
        End Using

        Return results
    End Function

    ''' <summary>
    ''' Returns jobs left stuck in 'processing' (crash, host closed mid-run) back to 'pending'
    ''' once they are older than <paramref name="staleMinutes"/>, so the queue self-heals.
    ''' 重试次数已用尽的卡死任务不再回到 pending，而是直接落 failed（dead-letter），
    ''' 否则一条永远跑不完的任务会被反复回收，把队列容量吃光。
    ''' </summary>
    Public Shared Function ReclaimStaleProcessingJobs(Optional staleMinutes As Integer = 30,
                                                      Optional maxAttempts As Integer = MaxJobAttempts) As Integer
        OfficeAiDatabase.EnsureInitialized()
        Dim cap = Math.Max(1, maxAttempts)
        Dim window = "-" & Math.Max(1, staleMinutes) & " minutes"
        Dim recovered As Integer = 0

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using tx = conn.BeginTransaction()
                Using dead As New SQLiteCommand(
                    "UPDATE memory_job SET status = 'failed', " &
                    "last_error = IFNULL(last_error, 'stale: attempt budget exhausted'), " &
                    "updated_at = datetime('now', 'localtime') " &
                    "WHERE status = 'processing' AND updated_at < datetime('now', 'localtime', @window) " &
                    "AND IFNULL(attempt_count, 0) >= @cap", conn)
                    dead.Transaction = tx
                    dead.Parameters.AddWithValue("@window", window)
                    dead.Parameters.AddWithValue("@cap", cap)
                    recovered += dead.ExecuteNonQuery()
                End Using

                Using requeue As New SQLiteCommand(
                    "UPDATE memory_job SET status = 'pending', updated_at = datetime('now', 'localtime') " &
                    "WHERE status = 'processing' AND updated_at < datetime('now', 'localtime', @window)", conn)
                    requeue.Transaction = tx
                    requeue.Parameters.AddWithValue("@window", window)
                    recovered += requeue.ExecuteNonQuery()
                End Using

                tx.Commit()
            End Using
        End Using
        Return recovered
    End Function

    Public Shared Sub UpsertSkillRegistry(record As SkillRegistryRecord)
        OfficeAiDatabase.EnsureInitialized()
        If record Is Nothing Then Throw New ArgumentNullException(NameOf(record))

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO skills_registry " &
                "(skill_name, file_path, app_scope, intent_types, trigger_keywords, description, embedding_json, usage_count, success_count, last_indexed_at, enabled) " &
                "VALUES (@skill_name, @file_path, @app_scope, @intent_types, @trigger_keywords, @description, @embedding_json, @usage_count, @success_count, datetime('now', 'localtime'), @enabled) " &
                "ON CONFLICT(skill_name) DO UPDATE SET " &
                "file_path = excluded.file_path, app_scope = excluded.app_scope, intent_types = excluded.intent_types, trigger_keywords = excluded.trigger_keywords, " &
                "description = excluded.description, embedding_json = excluded.embedding_json, last_indexed_at = excluded.last_indexed_at, enabled = excluded.enabled", conn)
                cmd.Parameters.AddWithValue("@skill_name", DbRequiredText(record.SkillName, "unnamed"))
                cmd.Parameters.AddWithValue("@file_path", DbRequiredText(record.FilePath, ""))
                cmd.Parameters.AddWithValue("@app_scope", If(record.AppScope, ""))
                cmd.Parameters.AddWithValue("@intent_types", If(record.IntentTypes, ""))
                cmd.Parameters.AddWithValue("@trigger_keywords", If(record.TriggerKeywords, ""))
                cmd.Parameters.AddWithValue("@description", DbText(record.Description))
                cmd.Parameters.AddWithValue("@embedding_json", DbText(record.EmbeddingJson))
                cmd.Parameters.AddWithValue("@usage_count", record.UsageCount)
                cmd.Parameters.AddWithValue("@success_count", record.SuccessCount)
                cmd.Parameters.AddWithValue("@enabled", If(record.Enabled, 1, 0))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Shared Function RetrieveSkillRegistry(query As String, intentType As String, appType As String, Optional topN As Integer = 5) As List(Of SkillRegistryRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim results As New List(Of SkillRegistryRecord)()
        Dim q = If(query, "").Trim()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT skill_name, file_path, app_scope, intent_types, trigger_keywords, description, embedding_json, usage_count, success_count, last_indexed_at, enabled, " &
                "(CASE WHEN skill_name LIKE @q THEN 40 ELSE 0 END + " &
                " CASE WHEN trigger_keywords LIKE @q THEN 30 ELSE 0 END + " &
                " CASE WHEN description LIKE @q THEN 15 ELSE 0 END + " &
                " CASE WHEN @has_intent = 1 AND intent_types LIKE @intent THEN 10 ELSE 0 END + " &
                " usage_count * 0.2 + success_count * 0.3) AS score " &
                "FROM skills_registry WHERE enabled = 1 " &
                "AND (@app_type = '' OR app_scope = '' OR app_scope LIKE @app_like) " &
                "AND (@q = '%%' OR skill_name LIKE @q OR trigger_keywords LIKE @q OR description LIKE @q OR (@has_intent = 1 AND intent_types LIKE @intent)) " &
                "ORDER BY score DESC, last_indexed_at DESC LIMIT @limit", conn)
                cmd.Parameters.AddWithValue("@q", "%" & q & "%")
                Dim intent = If(intentType, "").Trim()
                cmd.Parameters.AddWithValue("@intent", "%" & intent & "%")
                cmd.Parameters.AddWithValue("@has_intent", If(String.IsNullOrWhiteSpace(intent), 0, 1))
                cmd.Parameters.AddWithValue("@app_type", If(appType, "").Trim())
                cmd.Parameters.AddWithValue("@app_like", "%" & If(appType, "").Trim() & "%")
                cmd.Parameters.AddWithValue("@limit", Math.Max(1, topN))
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim skill = ReadSkillRegistry(rdr)
                        skill.Score = If(rdr.IsDBNull(11), 0.0R, Convert.ToDouble(rdr.GetValue(11)))
                        results.Add(skill)
                    End While
                End Using
            End Using
        End Using

        Return results
    End Function

    Public Shared Function GetSkillRegistryByName(skillName As String) As SkillRegistryRecord
        If String.IsNullOrWhiteSpace(skillName) Then Return Nothing
        OfficeAiDatabase.EnsureInitialized()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT skill_name, file_path, app_scope, intent_types, trigger_keywords, description, embedding_json, usage_count, success_count, last_indexed_at, enabled " &
                "FROM skills_registry WHERE skill_name = @skill_name LIMIT 1", conn)
                cmd.Parameters.AddWithValue("@skill_name", skillName.Trim())
                Using rdr = cmd.ExecuteReader()
                    If rdr.Read() Then Return ReadSkillRegistry(rdr)
                End Using
            End Using
        End Using

        Return Nothing
    End Function

    Public Shared Sub DisableSkillsNotIn(activeSkillNames As IEnumerable(Of String))
        OfficeAiDatabase.EnsureInitialized()
        Dim active As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        If activeSkillNames IsNot Nothing Then
            For Each name In activeSkillNames
                If Not String.IsNullOrWhiteSpace(name) Then active.Add(name.Trim())
            Next
        End If

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT skill_name, file_path, app_scope, intent_types, trigger_keywords, description, embedding_json, usage_count, success_count, last_indexed_at, enabled FROM skills_registry WHERE enabled = 1", conn)
                Dim toDisable As New List(Of String)()
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim name = If(rdr.IsDBNull(0), "", rdr.GetString(0))
                        If Not active.Contains(name) Then toDisable.Add(name)
                    End While
                End Using

                For Each name In toDisable
                    Using updateCmd As New SQLiteCommand("UPDATE skills_registry SET enabled = 0, last_indexed_at = datetime('now', 'localtime') WHERE skill_name = @skill_name", conn)
                        updateCmd.Parameters.AddWithValue("@skill_name", name)
                        updateCmd.ExecuteNonQuery()
                    End Using
                Next
            End Using
        End Using
    End Sub

    Public Shared Sub RecordSkillRegistryUsage(skillName As String, Optional success As Boolean = True)
        If String.IsNullOrWhiteSpace(skillName) Then Return
        OfficeAiDatabase.EnsureInitialized()

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE skills_registry SET usage_count = usage_count + 1, success_count = success_count + @success_delta, last_indexed_at = datetime('now', 'localtime') WHERE skill_name = @skill_name", conn)
                cmd.Parameters.AddWithValue("@success_delta", If(success, 1, 0))
                cmd.Parameters.AddWithValue("@skill_name", skillName.Trim())
                Dim affected = cmd.ExecuteNonQuery()
                If affected = 0 Then
                    Using insertCmd As New SQLiteCommand(
                        "INSERT INTO skills_registry " &
                        "(skill_name, file_path, app_scope, intent_types, trigger_keywords, description, embedding_json, usage_count, success_count, last_indexed_at, enabled) " &
                        "VALUES (@skill_name, '', '', '', '', NULL, NULL, 1, @success_delta, datetime('now', 'localtime'), 1)", conn)
                        insertCmd.Parameters.AddWithValue("@skill_name", skillName.Trim())
                        insertCmd.Parameters.AddWithValue("@success_delta", If(success, 1, 0))
                        insertCmd.ExecuteNonQuery()
                    End Using
                End If
            End Using
        End Using
    End Sub
End Class
