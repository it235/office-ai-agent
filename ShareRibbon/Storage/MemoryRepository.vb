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
    Public Shared Sub InsertAtomicMemory(content As String, Optional tags As String = Nothing, Optional sessionId As String = Nothing, Optional appType As String = Nothing, Optional embedding As String = Nothing)
        OfficeAiDatabase.EnsureInitialized()
        Dim ts = CType((DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalSeconds, Long)
        Dim app = If(String.IsNullOrEmpty(appType), "", appType.Trim())
        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(
                "INSERT INTO atomic_memory (timestamp, content, tags, session_id, app_type, embedding) VALUES (@ts, @content, @tags, @sid, @app, @emb)", conn)
                cmd.Parameters.AddWithValue("@ts", ts)
                cmd.Parameters.AddWithValue("@content", If(content, ""))
                cmd.Parameters.AddWithValue("@tags", If(tags, ""))
                cmd.Parameters.AddWithValue("@sid", If(sessionId, ""))
                cmd.Parameters.AddWithValue("@app", app)
                cmd.Parameters.AddWithValue("@emb", If(embedding, DBNull.Value))
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
        Dim sql = "SELECT id, timestamp, content, tags, session_id, create_time, embedding FROM atomic_memory WHERE 1=1"
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
                            .Embedding = If(rdr.IsDBNull(6), Nothing, rdr.GetString(6))
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
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
    ''' 按向量相似度检索原子记忆（RAG）。如果 query 有向量，则用余弦相似度排序；否则退回到 LIKE 查询
    ''' </summary>
    Public Shared Function GetRelevantMemories(query As String, topN As Integer, Optional queryEmbedding As Single() = Nothing, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing) As List(Of AtomicMemoryRecord)
        OfficeAiDatabase.EnsureInitialized()

        ' 首先获取所有符合时间范围的记忆（带 embedding）
        Dim allMemories As New List(Of AtomicMemoryRecord)()
        Dim sql = "SELECT id, timestamp, content, tags, session_id, create_time, embedding FROM atomic_memory WHERE 1=1"

        If startTime.HasValue Then
            sql &= " AND timestamp >= @st"
        End If
        If endTime.HasValue Then
            sql &= " AND timestamp <= @et"
        End If
        sql &= " ORDER BY timestamp DESC"

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(sql, conn)
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
                            .Embedding = If(rdr.IsDBNull(6), Nothing, rdr.GetString(6))
                        })
                    End While
                End Using
            End Using
        End Using

        ' 如果有查询向量且有带 embedding 的记忆，使用余弦相似度排序
        If queryEmbedding IsNot Nothing AndAlso queryEmbedding.Length > 0 Then
            Dim memoriesWithEmbedding = allMemories.Where(Function(m) Not String.IsNullOrWhiteSpace(m.Embedding)).ToList()

            If memoriesWithEmbedding.Count > 0 Then
                Debug.WriteLine($"[MemoryRepository] 使用向量检索，共有 {memoriesWithEmbedding.Count} 条带 embedding 的记忆")

                ' 计算相似度并排序
                Dim scoredMemories As New List(Of Tuple(Of AtomicMemoryRecord, Single))()

                For Each mem In memoriesWithEmbedding
                    Dim memEmbedding = EmbeddingService.DeserializeVector(mem.Embedding)
                    If memEmbedding IsNot Nothing Then
                        Dim similarity = EmbeddingService.CosineSimilarity(queryEmbedding, memEmbedding)
                        scoredMemories.Add(Tuple.Create(mem, similarity))
                    End If
                Next

                ' 按相似度降序排序，取 topN
                Dim sorted = scoredMemories.OrderByDescending(Function(t) t.Item2).Take(topN).ToList()

                Debug.WriteLine($"[MemoryRepository] 向量检索完成，返回 {sorted.Count} 条最相关记忆")
                For i = 0 To Math.Min(5, sorted.Count) - 1
                    Debug.WriteLine($"[MemoryRepository]   {i + 1}. 相似度: {sorted(i).Item2:F4}, 内容: {sorted(i).Item1.Content.Substring(0, Math.Min(50, sorted(i).Item1.Content.Length))}...")
                Next

                Return sorted.Select(Function(t) t.Item1).ToList()
            End If
        End If

        ' 退回到 LIKE 查询（无向量或向量不可用时）
        Debug.WriteLine($"[MemoryRepository] 退回到 LIKE 查询，query: {If(query?.Length > 50, query.Substring(0, 50) & "...", query)}")

        Dim fallbackList As New List(Of AtomicMemoryRecord)()
        Dim fallbackSql = "SELECT id, timestamp, content, tags, session_id, create_time, embedding FROM atomic_memory WHERE 1=1"

        If Not String.IsNullOrWhiteSpace(query) Then
            fallbackSql &= " AND (content LIKE @q OR tags LIKE @q)"
        End If
        If startTime.HasValue Then
            fallbackSql &= " AND timestamp >= @st"
        End If
        If endTime.HasValue Then
            fallbackSql &= " AND timestamp <= @et"
        End If
        fallbackSql &= " ORDER BY timestamp DESC LIMIT @limit"

        Using conn As New SQLiteConnection(OfficeAiDatabase.GetConnectionString())
            conn.Open()
            Using cmd As New SQLiteCommand(fallbackSql, conn)
                If Not String.IsNullOrWhiteSpace(query) Then
                    cmd.Parameters.AddWithValue("@q", "%" & query & "%")
                End If
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
                            .Embedding = If(rdr.IsDBNull(6), Nothing, rdr.GetString(6))
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
End Class
