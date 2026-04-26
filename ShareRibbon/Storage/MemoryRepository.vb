' ShareRibbon\Storage\MemoryRepository.vb
' 记忆相关表的 CRUD 访问

Imports System.Data.SQLite

''' <summary>
''' 原子记忆实体
''' </summary>
Public Class AtomicMemoryRecord
    Public Property Id As Long
    Public Property Timestamp As Long
    Public Property Content As String
    Public Property Tags As String
    Public Property SessionId As String
    Public Property CreateTime As String
    Public Property Embedding As String
    Public Property MemoryType As String
    Public Property Importance As Double
    Public Property AccessCount As Integer
    Public Property LastAccess As String
    Public Property SourceType As String
    Public Property LinkedMemories As String
    Public Property SimilarityScore As Single
End Class

''' <summary>
''' 记忆关联实体
''' </summary>
Public Class MemoryGraphRecord
    Public Property Id As Long
    Public Property SourceId As Long
    Public Property TargetId As Long
    Public Property RelationType As String
    Public Property Weight As Double
    Public Property CreatedAt As String
End Class

''' <summary>
''' 用户画像项实体
''' </summary>
Public Class UserProfileItem
    Public Property Id As Long
    Public Property Key As String
    Public Property Value As String
    Public Property Category As String
    Public Property Confidence As Double
    Public Property LastUpdated As String
    Public Property ObservationCount As Integer
End Class

''' <summary>
''' 对话分支实体
''' </summary>
Public Class ConversationBranchRecord
    Public Property Id As Long
    Public Property ConversationId As Long
    Public Property ParentMessageId As Long?
    Public Property BranchName As String
    Public Property IsActive As Boolean
    Public Property CreatedAt As String
End Class

''' <summary>
''' 技能使用统计实体
''' </summary>
Public Class SkillUsageRecord
    Public Property Id As Long
    Public Property SkillName As String
    Public Property UsageCount As Integer
    Public Property SuccessCount As Integer
    Public Property TotalTokens As Long
    Public Property LastUsedAt As String
    Public Property CreatedAt As String
    Public Property UpdatedAt As String
End Class

''' <summary>
''' 带评分的记忆结果
''' </summary>
Public Class MemoryWithScore
    Public Property Memory As AtomicMemoryRecord
    Public Property Score As Double
    Public Property Components As Dictionary(Of String, Double)
End Class

''' <summary>
''' 会话摘要实体
''' </summary>
Public Class SessionSummaryRecord
    Public Property Id As Long
    Public Property SessionId As String
    Public Property Title As String
    Public Property Snippet As String
    Public Property CreatedAt As String
End Class

''' <summary>
''' 记忆表 CRUD 访问
''' </summary>
Public Class MemoryRepository

    ''' <summary>
    ''' 插入原子记忆。appType 为当前宿主（Excel/Word/PowerPoint），用于按应用筛选。
    ''' </summary>
    Public Shared Sub InsertAtomicMemory(content As String, Optional tags As String = Nothing, Optional sessionId As String = Nothing, Optional appType As String = Nothing, Optional embedding As String = Nothing, Optional memoryType As String = "short_term")
        OfficeAiDatabase.EnsureInitialized()
        Dim ts = CType((DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long)
        Dim app = If(String.IsNullOrEmpty(appType), "", appType.Trim())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO atomic_memory (timestamp, content, tags, session_id, app_type, embedding, memory_type) VALUES (@ts, @content, @tags, @sid, @app, @emb, @mtype)", conn)
                cmd.Parameters.AddWithValue("@ts", ts)
                cmd.Parameters.AddWithValue("@content", If(content, ""))
                cmd.Parameters.AddWithValue("@tags", If(tags, ""))
                cmd.Parameters.AddWithValue("@sid", If(sessionId, ""))
                cmd.Parameters.AddWithValue("@app", app)
                cmd.Parameters.AddWithValue("@emb", If(embedding, DBNull.Value))
                cmd.Parameters.AddWithValue("@mtype", If(String.IsNullOrEmpty(memoryType), "short_term", memoryType))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 列出原子记忆（分页，供管理界面用）。appType 为空时不过滤，否则只返回该宿主下的记录。
    ''' </summary>
    Public Shared Function ListAtomicMemories(Optional limit As Integer = 100, Optional offset As Integer = 0, Optional appType As String = Nothing) As List(Of AtomicMemoryRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of AtomicMemoryRecord)()
        Dim app = If(String.IsNullOrEmpty(appType), "", appType.Trim())
        Dim hasApp = Not String.IsNullOrEmpty(app)
        Dim sql = "SELECT id, timestamp, content, tags, session_id, create_time, embedding, memory_type, importance, access_count, last_access, source_type, linked_memories FROM atomic_memory WHERE 1=1"
        ' 按应用过滤：仅显示当前宿主或历史无 app_type 的记录
        If hasApp Then sql &= " AND (app_type = @app OR app_type IS NULL OR app_type = '')"
        sql &= " ORDER BY timestamp DESC LIMIT @limit OFFSET @offset"
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(sql, conn)
                If hasApp Then cmd.Parameters.AddWithValue("@app", app)
                cmd.Parameters.AddWithValue("@limit", limit)
                cmd.Parameters.AddWithValue("@offset", offset)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        list.Add(New AtomicMemoryRecord With {
                            .Id = rdr.GetInt64(0),
                            .Timestamp = rdr.GetInt64(1),
                            .Content = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
                            .Tags = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
                            .SessionId = If(rdr.IsDBNull(4), "", rdr.GetString(4)),
                            .CreateTime = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
                            .Embedding = If(rdr.IsDBNull(6), Nothing, rdr.GetString(6)),
                            .MemoryType = If(rdr.IsDBNull(7), "short_term", rdr.GetString(7)),
                            .Importance = If(rdr.IsDBNull(8), 0.5, rdr.GetDouble(8)),
                            .AccessCount = If(rdr.IsDBNull(9), 0, rdr.GetInt32(9)),
                            .LastAccess = If(rdr.IsDBNull(10), "", rdr.GetString(10)),
                            .SourceType = If(rdr.IsDBNull(11), "general", rdr.GetString(11)),
                            .LinkedMemories = If(rdr.IsDBNull(12), "", rdr.GetString(12))
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    ''' <summary>
    ''' 快速检查数据库中是否存在带 embedding 的长期记忆（避免无谓的向量 API 调用）
    ''' </summary>
    Public Shared Function HasMemoriesWithEmbedding(Optional appType As String = Nothing) As Boolean
        Try
            OfficeAiDatabase.EnsureInitialized()
            Dim app = If(String.IsNullOrEmpty(appType), "", appType.Trim())
            Dim hasApp = Not String.IsNullOrEmpty(app)
            Dim sql = "SELECT COUNT(1) FROM atomic_memory WHERE memory_type = 'long_term' AND embedding IS NOT NULL AND embedding != ''"
            If hasApp Then sql &= " AND (app_type = @app OR app_type IS NULL OR app_type = '')"
            Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
                conn.Open()
                Using cmd As New SQLiteCommand(sql, conn)
                    If hasApp Then cmd.Parameters.AddWithValue("@app", app)
                    Dim count = CInt(cmd.ExecuteScalar())
                    Return count > 0
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[MemoryRepository] HasMemoriesWithEmbedding 失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 删除原子记忆
    ''' </summary>
    Public Shared Sub DeleteAtomicMemory(id As Long)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("DELETE FROM atomic_memory WHERE id=@id", conn)
                cmd.Parameters.AddWithValue("@id", id)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 按向量相似度检索原子记忆（RAG）。支持 appType 过滤、相似度阈值、时间衰减。
    ''' </summary>
    Public Shared Function GetRelevantMemories(query As String, topN As Integer, Optional queryEmbedding As Single() = Nothing, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing, Optional appType As String = Nothing) As List(Of AtomicMemoryRecord)
        OfficeAiDatabase.EnsureInitialized()

        Dim app = If(String.IsNullOrEmpty(appType), "", appType.Trim())
        Dim hasApp = Not String.IsNullOrEmpty(app)
        Dim nowUnix = CType((DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long)

        Dim allMemories As New List(Of AtomicMemoryRecord)()
        Dim sql = "SELECT id, timestamp, content, tags, session_id, create_time, embedding, memory_type, importance, access_count, last_access, source_type, linked_memories FROM atomic_memory WHERE memory_type = 'long_term'"

        If hasApp Then sql &= " AND (app_type = @app OR app_type IS NULL OR app_type = '')"
        If startTime.HasValue Then sql &= " AND timestamp >= @st"
        If endTime.HasValue Then sql &= " AND timestamp <= @et"
        sql &= " ORDER BY timestamp DESC LIMIT 500"

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(sql, conn)
                If hasApp Then cmd.Parameters.AddWithValue("@app", app)
                If startTime.HasValue Then
                    cmd.Parameters.AddWithValue("@st", CType((startTime.Value.ToUniversalTime() - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long))
                End If
                If endTime.HasValue Then
                    cmd.Parameters.AddWithValue("@et", CType((endTime.Value.ToUniversalTime() - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long))
                End If

                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        allMemories.Add(New AtomicMemoryRecord With {
                            .Id = rdr.GetInt64(0),
                            .Timestamp = rdr.GetInt64(1),
                            .Content = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
                            .Tags = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
                            .SessionId = If(rdr.IsDBNull(4), "", rdr.GetString(4)),
                            .CreateTime = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
                            .Embedding = If(rdr.IsDBNull(6), Nothing, rdr.GetString(6)),
                            .MemoryType = If(rdr.IsDBNull(7), "short_term", rdr.GetString(7)),
                            .Importance = If(rdr.IsDBNull(8), 0.5, rdr.GetDouble(8)),
                            .AccessCount = If(rdr.IsDBNull(9), 0, rdr.GetInt32(9)),
                            .LastAccess = If(rdr.IsDBNull(10), "", rdr.GetString(10)),
                            .SourceType = If(rdr.IsDBNull(11), "general", rdr.GetString(11)),
                            .LinkedMemories = If(rdr.IsDBNull(12), "", rdr.GetString(12))
                        })
                    End While
                End Using
            End Using
        End Using

        If queryEmbedding IsNot Nothing AndAlso queryEmbedding.Length > 0 Then
            Dim memoriesWithEmbedding = allMemories.Where(Function(m) Not String.IsNullOrWhiteSpace(m.Embedding)).ToList()

            If memoriesWithEmbedding.Count > 0 Then
                Debug.WriteLine($"[MemoryRepository] 使用向量检索，共有 {memoriesWithEmbedding.Count} 条带 embedding 的记忆")

                Dim threshold = MemoryConfig.RagSimilarityThreshold
                Dim decayRate = MemoryConfig.RagTimeDecayRate
                Dim scoredMemories As New List(Of Tuple(Of AtomicMemoryRecord, Single))()

                For Each mem In memoriesWithEmbedding
                    Dim memEmbedding = EmbeddingService.DeserializeVector(mem.Embedding)
                    If memEmbedding IsNot Nothing Then
                        Dim similarity = EmbeddingService.CosineSimilarity(queryEmbedding, memEmbedding)
                        mem.SimilarityScore = similarity ' 保存相似度供后续使用
                        Dim daysSinceCreation = CSng(Math.Max(0, nowUnix - mem.Timestamp)) / 86400.0F
                        Dim timeDecay = 1.0F / (1.0F + daysSinceCreation * decayRate)
                        Dim finalScore = similarity * timeDecay

                        If finalScore >= threshold Then
                            scoredMemories.Add(Tuple.Create(mem, finalScore))
                        End If
                    End If
                Next

                Dim sorted = scoredMemories.OrderByDescending(Function(t) t.Item2).Take(topN).ToList()

                Debug.WriteLine($"[MemoryRepository] 向量检索完成，阈值={threshold:F2}，返回 {sorted.Count} 条")
                For i = 0 To Math.Min(5, sorted.Count) - 1
                    Debug.WriteLine($"[MemoryRepository]   {i + 1}. 分数: {sorted(i).Item2:F4}, 内容: {sorted(i).Item1.Content.Substring(0, Math.Min(50, sorted(i).Item1.Content.Length))}...")
                Next

                If sorted.Count > 0 Then
                    Return sorted.Select(Function(t) t.Item1).ToList()
                End If
            End If
        End If

        Debug.WriteLine($"[MemoryRepository] 退回到 LIKE 查询，query: {If(query?.Length > 50, query.Substring(0, 50) & "...", query)}")

        Dim fallbackList As New List(Of AtomicMemoryRecord)()
        Dim fallbackSql = "SELECT id, timestamp, content, tags, session_id, create_time, embedding, memory_type, importance, access_count, last_access, source_type, linked_memories FROM atomic_memory WHERE memory_type = 'long_term'"

        If Not String.IsNullOrWhiteSpace(query) Then
            fallbackSql &= " AND (content LIKE @q OR tags LIKE @q)"
        End If
        If hasApp Then fallbackSql &= " AND (app_type = @app OR app_type IS NULL OR app_type = '')"
        If startTime.HasValue Then fallbackSql &= " AND timestamp >= @st"
        If endTime.HasValue Then fallbackSql &= " AND timestamp <= @et"
        fallbackSql &= " ORDER BY timestamp DESC LIMIT @limit"

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(fallbackSql, conn)
                If Not String.IsNullOrWhiteSpace(query) Then
                    cmd.Parameters.AddWithValue("@q", "%" & query & "%")
                End If
                If hasApp Then cmd.Parameters.AddWithValue("@app", app)
                If startTime.HasValue Then
                    cmd.Parameters.AddWithValue("@st", CType((startTime.Value.ToUniversalTime() - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long))
                End If
                If endTime.HasValue Then
                    cmd.Parameters.AddWithValue("@et", CType((endTime.Value.ToUniversalTime() - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long))
                End If
                cmd.Parameters.AddWithValue("@limit", topN)

                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        fallbackList.Add(New AtomicMemoryRecord With {
                            .Id = rdr.GetInt64(0),
                            .Timestamp = rdr.GetInt64(1),
                            .Content = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
                            .Tags = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
                            .SessionId = If(rdr.IsDBNull(4), "", rdr.GetString(4)),
                            .CreateTime = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
                            .Embedding = If(rdr.IsDBNull(6), Nothing, rdr.GetString(6)),
                            .MemoryType = If(rdr.IsDBNull(7), "short_term", rdr.GetString(7)),
                            .Importance = If(rdr.IsDBNull(8), 0.5, rdr.GetDouble(8)),
                            .AccessCount = If(rdr.IsDBNull(9), 0, rdr.GetInt32(9)),
                            .LastAccess = If(rdr.IsDBNull(10), "", rdr.GetString(10)),
                            .SourceType = If(rdr.IsDBNull(11), "general", rdr.GetString(11)),
                            .LinkedMemories = If(rdr.IsDBNull(12), "", rdr.GetString(12))
                        })
                    End While
                End Using
            End Using
        End Using

        Return fallbackList
    End Function

    ''' <summary>
    ''' 获取用户画像
    ''' </summary>
    Public Shared Function GetUserProfile() As String
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT content FROM user_profile ORDER BY id DESC LIMIT 1", conn)
                Dim obj = cmd.ExecuteScalar()
                Return If(obj Is Nothing OrElse obj Is DBNull.Value, "", obj.ToString())
            End Using
        End Using
    End Function

    ''' <summary>
    ''' 更新用户画像
    ''' </summary>
    Public Shared Sub UpdateUserProfile(content As String)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            ' 若存在则更新，否则插入
            Using check As New SQLiteCommand("SELECT COUNT(*) FROM user_profile", conn)
                Dim cnt = Convert.ToInt32(check.ExecuteScalar())
                If cnt > 0 Then
                    Using cmd As New SQLiteCommand("UPDATE user_profile SET content=@c, updated_at=datetime('now','localtime')", conn)
                        cmd.Parameters.AddWithValue("@c", If(content, ""))
                        cmd.ExecuteNonQuery()
                    End Using
                Else
                    Using cmd As New SQLiteCommand("INSERT INTO user_profile (content) VALUES (@c)", conn)
                        cmd.Parameters.AddWithValue("@c", If(content, ""))
                        cmd.ExecuteNonQuery()
                    End Using
                End If
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 获取近期会话摘要
    ''' </summary>
    Public Shared Function GetRecentSessionSummaries(limit As Integer) As List(Of SessionSummaryRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of SessionSummaryRecord)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "SELECT id, session_id, title, snippet, created_at FROM session_summary ORDER BY created_at DESC LIMIT @limit", conn)
                cmd.Parameters.AddWithValue("@limit", limit)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        list.Add(New SessionSummaryRecord With {
                            .Id = rdr.GetInt64(0),
                            .SessionId = rdr.GetString(1),
                            .Title = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
                            .Snippet = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
                            .CreatedAt = rdr.GetString(4)
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    ''' <summary>
    ''' 插入会话摘要
    ''' </summary>
    Public Shared Sub InsertSessionSummary(sessionId As String, title As String, snippet As String)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO session_summary (session_id, title, snippet) VALUES (@sid, @title, @snippet)", conn)
                cmd.Parameters.AddWithValue("@sid", sessionId)
                cmd.Parameters.AddWithValue("@title", If(title, ""))
                cmd.Parameters.AddWithValue("@snippet", If(snippet, ""))
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 插入原子记忆（增强版，支持新字段）
    ''' </summary>
    Public Shared Function InsertMemory(content As String, Optional embedding As String = Nothing, Optional sessionId As String = Nothing, Optional appType As String = Nothing, Optional memoryType As String = "long_term", Optional importance As Double = 0.5, Optional sourceType As String = "general") As Long
        OfficeAiDatabase.EnsureInitialized()
        Dim ts = CType((DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long)
        Dim app = If(String.IsNullOrEmpty(appType), "", appType.Trim())
        Dim newId As Long = 0
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO atomic_memory (timestamp, content, tags, session_id, app_type, embedding, memory_type, importance, source_type) VALUES (@ts, @content, @tags, @sid, @app, @emb, @mtype, @imp, @stype); SELECT last_insert_rowid();", conn)
                cmd.Parameters.AddWithValue("@ts", ts)
                cmd.Parameters.AddWithValue("@content", If(content, ""))
                cmd.Parameters.AddWithValue("@tags", "")
                cmd.Parameters.AddWithValue("@sid", If(sessionId, ""))
                cmd.Parameters.AddWithValue("@app", app)
                cmd.Parameters.AddWithValue("@emb", If(embedding, DBNull.Value))
                cmd.Parameters.AddWithValue("@mtype", If(String.IsNullOrEmpty(memoryType), "long_term", memoryType))
                cmd.Parameters.AddWithValue("@imp", importance)
                cmd.Parameters.AddWithValue("@stype", If(String.IsNullOrEmpty(sourceType), "general", sourceType))
                Dim obj = cmd.ExecuteScalar()
                If obj IsNot Nothing AndAlso Not IsDBNull(obj) Then
                    newId = Convert.ToInt64(obj)
                End If
            End Using
        End Using
        Return newId
    End Function

    ''' <summary>
    ''' 更新记忆访问次数和时间
    ''' </summary>
    Public Shared Sub UpdateMemoryAccess(id As Long)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE atomic_memory SET access_count = access_count + 1, last_access = datetime('now', 'localtime') WHERE id = @id", conn)
                cmd.Parameters.AddWithValue("@id", id)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 获取相似记忆（用于建立关联）
    ''' </summary>
    Public Shared Function GetSimilarMemories(memoryId As Long, Optional topK As Integer = 5) As List(Of AtomicMemoryRecord)
        OfficeAiDatabase.EnsureInitialized()
        Dim list As New List(Of AtomicMemoryRecord)()
        Dim memory As AtomicMemoryRecord = Nothing
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT id, embedding FROM atomic_memory WHERE id = @id", conn)
                cmd.Parameters.AddWithValue("@id", memoryId)
                Using rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        memory = New AtomicMemoryRecord With {
                            .Id = rdr.GetInt64(0),
                            .Embedding = If(rdr.IsDBNull(1), Nothing, rdr.GetString(1))
                        }
                    End If
                End Using
            End Using
            If memory IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(memory.Embedding) Then
                Dim memEmbedding = EmbeddingService.DeserializeVector(memory.Embedding)
                If memEmbedding IsNot Nothing Then
                    Using cmd As New SQLiteCommand("SELECT id, timestamp, content, tags, session_id, create_time, embedding, memory_type, importance, access_count, last_access, source_type, linked_memories FROM atomic_memory WHERE id != @id AND memory_type = 'long_term' AND embedding IS NOT NULL AND embedding != '' LIMIT 100", conn)
                        cmd.Parameters.AddWithValue("@id", memoryId)
                        Using rdr = cmd.ExecuteReader()
                            While rdr.Read()
                                Dim mem = New AtomicMemoryRecord With {
                                    .Id = rdr.GetInt64(0),
                                    .Timestamp = rdr.GetInt64(1),
                                    .Content = If(rdr.IsDBNull(2), "", rdr.GetString(2)),
                                    .Tags = If(rdr.IsDBNull(3), "", rdr.GetString(3)),
                                    .SessionId = If(rdr.IsDBNull(4), "", rdr.GetString(4)),
                                    .CreateTime = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
                                    .Embedding = If(rdr.IsDBNull(6), Nothing, rdr.GetString(6)),
                                    .MemoryType = If(rdr.IsDBNull(7), "short_term", rdr.GetString(7)),
                                    .Importance = If(rdr.IsDBNull(8), 0.5, rdr.GetDouble(8)),
                                    .AccessCount = If(rdr.IsDBNull(9), 0, rdr.GetInt32(9)),
                                    .LastAccess = If(rdr.IsDBNull(10), "", rdr.GetString(10)),
                                    .SourceType = If(rdr.IsDBNull(11), "general", rdr.GetString(11)),
                                    .LinkedMemories = If(rdr.IsDBNull(12), "", rdr.GetString(12))
                                }
                                Dim otherEmbedding = EmbeddingService.DeserializeVector(mem.Embedding)
                                If otherEmbedding IsNot Nothing Then
                                    mem.SimilarityScore = EmbeddingService.CosineSimilarity(memEmbedding, otherEmbedding)
                                    list.Add(mem)
                                End If
                            End While
                        End Using
                    End Using
                    list = list.OrderByDescending(Function(m) m.SimilarityScore).Take(topK).ToList()
                End If
            End If
        End Using
        Return list
    End Function

    ''' <summary>
    ''' 添加记忆关联
    ''' </summary>
    Public Shared Sub AddMemoryRelation(sourceId As Long, targetId As Long, relationType As String, weight As Double)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO memory_graph (source_id, target_id, relation_type, weight) VALUES (@sid, @tid, @rtype, @w)", conn)
                cmd.Parameters.AddWithValue("@sid", sourceId)
                cmd.Parameters.AddWithValue("@tid", targetId)
                cmd.Parameters.AddWithValue("@rtype", If(String.IsNullOrEmpty(relationType), "similar", relationType))
                cmd.Parameters.AddWithValue("@w", weight)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 过期低重要性的记忆
    ''' </summary>
    Public Shared Sub ExpireLowImportanceMemories(sessionId As String, threshold As Double)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "UPDATE atomic_memory SET memory_type = 'expired' WHERE session_id = @sid AND importance < @t AND memory_type = 'short_term'", conn)
                cmd.Parameters.AddWithValue("@sid", If(sessionId, ""))
                cmd.Parameters.AddWithValue("@t", threshold)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 获取所有用户画像项
    ''' </summary>
    Public Shared Function GetAllUserProfile() As Dictionary(Of String, UserProfileItem)
        OfficeAiDatabase.EnsureInitialized()
        Dim dict As New Dictionary(Of String, UserProfileItem)()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT id, key, value, category, confidence, last_updated, observation_count FROM user_profile", conn)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim item = New UserProfileItem With {
                            .Id = rdr.GetInt64(0),
                            .Key = rdr.GetString(1),
                            .Value = rdr.GetString(2),
                            .Category = If(rdr.IsDBNull(3), "preference", rdr.GetString(3)),
                            .Confidence = If(rdr.IsDBNull(4), 0.5, rdr.GetDouble(4)),
                            .LastUpdated = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
                            .ObservationCount = If(rdr.IsDBNull(6), 1, rdr.GetInt32(6))
                        }
                        dict(item.Key) = item
                    End While
                End Using
            End Using
        End Using
        Return dict
    End Function

    ''' <summary>
    ''' 更新或插入用户画像项
    ''' </summary>
    Public Shared Sub UpsertUserProfile(key As String, value As String, category As String, Optional confidence As Double = 0.5)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using checkCmd As New SQLiteCommand("SELECT id, observation_count FROM user_profile WHERE key = @k", conn)
                checkCmd.Parameters.AddWithValue("@k", key)
                Dim existingId As Long? = Nothing
                Dim existingCount As Integer = 1
                Using rdr = checkCmd.ExecuteReader()
                    If rdr.Read() Then
                        existingId = rdr.GetInt64(0)
                        existingCount = If(rdr.IsDBNull(1), 1, rdr.GetInt32(1)) + 1
                    End If
                End Using
                If existingId.HasValue Then
                    Using updateCmd As New SQLiteCommand(
                        "UPDATE user_profile SET value = @v, category = @c, confidence = @conf, observation_count = @cnt, last_updated = datetime('now', 'localtime') WHERE id = @id", conn)
                        updateCmd.Parameters.AddWithValue("@id", existingId.Value)
                        updateCmd.Parameters.AddWithValue("@v", value)
                        updateCmd.Parameters.AddWithValue("@c", If(String.IsNullOrEmpty(category), "preference", category))
                        updateCmd.Parameters.AddWithValue("@conf", Math.Min(1.0, confidence))
                        updateCmd.Parameters.AddWithValue("@cnt", existingCount)
                        updateCmd.ExecuteNonQuery()
                    End Using
                Else
                    Using insertCmd As New SQLiteCommand(
                        "INSERT INTO user_profile (key, value, category, confidence) VALUES (@k, @v, @c, @conf)", conn)
                        insertCmd.Parameters.AddWithValue("@k", key)
                        insertCmd.Parameters.AddWithValue("@v", value)
                        insertCmd.Parameters.AddWithValue("@c", If(String.IsNullOrEmpty(category), "preference", category))
                        insertCmd.Parameters.AddWithValue("@conf", Math.Min(1.0, confidence))
                        insertCmd.ExecuteNonQuery()
                    End Using
                End If
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 获取技能使用统计
    ''' </summary>
    Public Shared Function GetSkillUsage(skillName As String) As SkillUsageRecord
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand("SELECT id, skill_name, usage_count, success_count, total_tokens, last_used_at, created_at, updated_at FROM skills_usage WHERE skill_name = @name", conn)
                cmd.Parameters.AddWithValue("@name", skillName)
                Using rdr = cmd.ExecuteReader()
                    If rdr.Read() Then
                        Return New SkillUsageRecord With {
                            .Id = rdr.GetInt64(0),
                            .SkillName = rdr.GetString(1),
                            .UsageCount = If(rdr.IsDBNull(2), 0, rdr.GetInt32(2)),
                            .SuccessCount = If(rdr.IsDBNull(3), 0, rdr.GetInt32(3)),
                            .TotalTokens = If(rdr.IsDBNull(4), 0, rdr.GetInt64(4)),
                            .LastUsedAt = If(rdr.IsDBNull(5), "", rdr.GetString(5)),
                            .CreatedAt = If(rdr.IsDBNull(6), "", rdr.GetString(6)),
                            .UpdatedAt = If(rdr.IsDBNull(7), "", rdr.GetString(7))
                        }
                    End If
                End Using
            End Using
        End Using
        Return Nothing
    End Function

    ''' <summary>
    ''' 记录技能使用
    ''' </summary>
    Public Shared Sub RecordSkillUsage(skillName As String, success As Boolean, Optional tokensUsed As Long = 0)
        OfficeAiDatabase.EnsureInitialized()
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using checkCmd As New SQLiteCommand("SELECT id FROM skills_usage WHERE skill_name = @name", conn)
                checkCmd.Parameters.AddWithValue("@name", skillName)
                Dim existingId As Object = checkCmd.ExecuteScalar()
                If existingId IsNot Nothing AndAlso Not IsDBNull(existingId) Then
                    Using updateCmd As New SQLiteCommand(
                        "UPDATE skills_usage SET usage_count = usage_count + 1, success_count = success_count + @s, total_tokens = total_tokens + @t, last_used_at = datetime('now', 'localtime'), updated_at = datetime('now', 'localtime') WHERE id = @id", conn)
                        updateCmd.Parameters.AddWithValue("@id", Convert.ToInt64(existingId))
                        updateCmd.Parameters.AddWithValue("@s", If(success, 1, 0))
                        updateCmd.Parameters.AddWithValue("@t", tokensUsed)
                        updateCmd.ExecuteNonQuery()
                    End Using
                Else
                    Using insertCmd As New SQLiteCommand(
                        "INSERT INTO skills_usage (skill_name, usage_count, success_count, total_tokens, last_used_at) VALUES (@name, 1, @s, @t, datetime('now', 'localtime'))", conn)
                        insertCmd.Parameters.AddWithValue("@name", skillName)
                        insertCmd.Parameters.AddWithValue("@s", If(success, 1, 0))
                        insertCmd.Parameters.AddWithValue("@t", tokensUsed)
                        insertCmd.ExecuteNonQuery()
                    End Using
                End If
            End Using
        End Using
    End Sub
End Class
