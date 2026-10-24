' ShareRibbon\Controls\Services\MemoryService.vb
' 记忆服务：封装 RAG、用户画像、会话摘要、异步写入

Imports System.Threading.Tasks

''' <summary>
''' 记忆服务：被动 RAG、用户画像、近期会话摘要、异步原子记忆写入
''' </summary>
Public Class MemoryService

    ''' <summary>
    ''' Sync RAG entry for legacy callers. Prefer GetRelevantMemoriesAsync on chat/agent paths.
    ''' Embedding is resolved via SyncOverAsync (thread-pool, timeout) — never UI-context .Result.
    ''' </summary>
    Public Shared Function GetRelevantMemories(query As String, Optional topN As Integer? = Nothing, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing, Optional appType As String = Nothing) As List(Of AtomicMemoryRecord)
        Dim result = SyncOverAsync.Run(Function() GetRelevantMemoriesAsync(query, topN, startTime, endTime, appType), 5000)
        Return If(result, New List(Of AtomicMemoryRecord)())
    End Function

    ''' <summary>
    ''' 被动 RAG 异步版本：避免 UI 线程阻塞
    ''' </summary>
    Public Shared Async Function GetRelevantMemoriesAsync(query As String, Optional topN As Integer? = Nothing, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing, Optional appType As String = Nothing) As Task(Of List(Of AtomicMemoryRecord))
        Dim n = If(topN.HasValue, topN.Value, MemoryConfig.RagTopN)

        Dim queryEmbedding As Single() = Nothing
        Try
            If Not String.IsNullOrWhiteSpace(query) AndAlso EmbeddingService.IsEmbeddingAvailable() AndAlso
               MemoryRepository.HasMemoriesWithEmbedding(appType) Then
                Debug.WriteLine("[MemoryService] generating query embedding (async)...")
                queryEmbedding = Await EmbeddingService.GetEmbeddingAsync(query).ConfigureAwait(False)
                If queryEmbedding IsNot Nothing Then
                    Debug.WriteLine($"[MemoryService] query embedding dim={queryEmbedding.Length}")
                End If
            ElseIf Not EmbeddingService.IsEmbeddingAvailable() Then
                Debug.WriteLine("[MemoryService] Embedding unavailable; keyword search only")
            Else
                Debug.WriteLine("[MemoryService] No embedded memories; skip vector API")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] async embedding failed: {ex.Message}")
        End Try

        Return Await Task.Run(Function() MemoryRepository.GetRelevantMemories(query, n, queryEmbedding, startTime, endTime, appType)).ConfigureAwait(False)
    End Function

    Public Shared Function GetRelevantStructuredMemories(query As String, Optional topN As Integer? = Nothing, Optional appType As String = Nothing, Optional documentId As String = Nothing, Optional projectId As String = Nothing) As List(Of MemoryItemRecord)
        Dim result = SyncOverAsync.Run(Function() GetRelevantStructuredMemoriesAsync(query, topN, appType, documentId, projectId), 5000)
        Return If(result, New List(Of MemoryItemRecord)())
    End Function

    Public Shared Async Function GetRelevantStructuredMemoriesAsync(query As String, Optional topN As Integer? = Nothing, Optional appType As String = Nothing, Optional documentId As String = Nothing, Optional projectId As String = Nothing) As Task(Of List(Of MemoryItemRecord))
        Dim n = If(topN.HasValue, topN.Value, MemoryConfig.RagTopN)
        Dim queryEmbedding As Single() = Nothing

        Try
            If Not String.IsNullOrWhiteSpace(query) AndAlso EmbeddingService.IsEmbeddingAvailable() AndAlso
               AgentMemoryRepository.HasReadyMemoryEmbeddings(appType) Then
                queryEmbedding = Await EmbeddingService.GetEmbeddingAsync(query).ConfigureAwait(False)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] structured embedding failed: {ex.Message}")
        End Try

        If queryEmbedding IsNot Nothing AndAlso queryEmbedding.Length > 0 Then
            Dim vectorResults = Await Task.Run(Function() AgentMemoryRepository.RetrieveMemoryItemsByVector(queryEmbedding, query, appType, documentId, projectId, n)).ConfigureAwait(False)
            If vectorResults IsNot Nothing AndAlso vectorResults.Count > 0 Then
                Return vectorResults
            End If
        End If

        Return Await Task.Run(Function() AgentMemoryRepository.RetrieveMemoryItems(query, appType, documentId, projectId, n)).ConfigureAwait(False)
    End Function

    ''' <summary>
    ''' 获取用户画像
    ''' </summary>
    Public Shared Function GetUserProfile() As String
        If Not MemoryConfig.EnableUserProfile Then Return ""
        Return MemoryRepository.GetUserProfile()
    End Function

    ''' <summary>
    ''' 获取近期会话摘要：与历史侧边栏同源，都从 conversation_event 聚合，不再依赖单独写入的 session_summary
    ''' </summary>
    Public Shared Function GetRecentSessionSummaries(Optional limit As Integer? = Nothing, Optional appType As String = Nothing) As List(Of SessionSummaryRecord)
        Dim n = If(limit.HasValue, limit.Value, MemoryConfig.SessionSummaryLimit)
        Return ConversationRepository.GetRecentSessions(n, appType)
    End Function

    ''' <summary>
    ''' 保存文件解析内容到记忆（用于在收到AI回复前保存引用的文件内容）- 同步保存确保立即可用
    ''' </summary>
    Public Shared Sub SaveFileContentToMemory(userPrompt As String, fileContent As String, sessionId As String, Optional appType As String = Nothing)
        Try
            If String.IsNullOrWhiteSpace(userPrompt) AndAlso String.IsNullOrWhiteSpace(fileContent) Then Return

            ' 取用户问题和文件内容的摘要保存
            Dim maxLen = MemoryConfig.AtomicContentMaxLength
            Dim u = (If(userPrompt, "").Trim())
            Dim f = (If(fileContent, "").Trim())
            Dim uPart = If(u.Length > maxLen \ 2, u.Substring(0, maxLen \ 2), u)
            Dim fPart = If(f.Length > maxLen \ 2, f.Substring(0, maxLen \ 2), f)
            Dim candidate = uPart & " [文件内容] " & fPart
            If String.IsNullOrWhiteSpace(candidate) OrElse candidate.Length < 10 Then Return

            Dim prefix = candidate.Substring(0, Math.Min(50, candidate.Length))
            If MemoryRepository.ExistsMemoryWithPrefix(sessionId, prefix, "short_term") Then Return

            Dim importance = UnifiedMemoryService.CalculateImportance(candidate, "knowledge", Nothing)
            Dim memoryId = MemoryRepository.InsertMemory(candidate, Nothing, sessionId, appType, "short_term", importance, "file_content")
            Debug.WriteLine($"[MemoryService] 已同步保存文件内容到记忆(ID={memoryId})，长度: {candidate.Length}")
        Catch ex As Exception
            Debug.WriteLine($"SaveFileContentToMemory 失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 主动 RAG 工具：按 keyword 和可选时间范围检索
    ''' </summary>
    Public Shared Function SearchMemories(keyword As String, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing, Optional appType As String = Nothing) As List(Of AtomicMemoryRecord)
        Dim result = SyncOverAsync.Run(Function() SearchMemoriesAsync(keyword, startTime, endTime, appType), 5000)
        Return If(result, New List(Of AtomicMemoryRecord)())
    End Function

    Public Shared Async Function SearchMemoriesAsync(keyword As String, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing, Optional appType As String = Nothing) As Task(Of List(Of AtomicMemoryRecord))
        Dim queryEmbedding As Single() = Nothing
        Try
            If Not String.IsNullOrWhiteSpace(keyword) AndAlso EmbeddingService.IsEmbeddingAvailable() AndAlso
               MemoryRepository.HasMemoriesWithEmbedding(appType) Then
                queryEmbedding = Await EmbeddingService.GetEmbeddingAsync(keyword).ConfigureAwait(False)
            Else
                Debug.WriteLine("[MemoryService] SearchMemoriesAsync: skip vector API")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] SearchMemoriesAsync embedding failed: {ex.Message}")
        End Try

        Return Await Task.Run(Function() MemoryRepository.GetRelevantMemories(keyword, MemoryConfig.RagTopN, queryEmbedding, startTime, endTime, appType)).ConfigureAwait(False)
    End Function

    Public Shared Function PromoteMemoryToLongTerm(memoryId As Long) As Boolean
        Return MemoryRepository.PromoteMemoryToLongTerm(memoryId)
    End Function

    Public Shared Function PromoteImportantShortTermMemories(sessionId As String, Optional threshold As Double = 0.65, Optional limit As Integer = 20) As Integer
        Return MemoryRepository.PromoteImportantShortTermMemories(sessionId, threshold, limit)
    End Function

    Public Shared Sub ConsolidateSessionMemoriesAsync(sessionId As String)
        If String.IsNullOrWhiteSpace(sessionId) Then Return

        Task.Run(Sub()
                     Try
                         Dim promotedByImportance = MemoryRepository.PromoteImportantShortTermMemories(sessionId, 0.65, 20)
                         Dim promotedByAccess = MemoryRepository.PromoteAccessedShortTermMemories(sessionId, 2, 20)
                         Dim conflicts = MemoryRepository.RecordPotentialConflictsForSession(sessionId)
                         MemoryRepository.ExpireLowImportanceMemories(sessionId, 0.15)
                         Debug.WriteLine($"[MemoryService] 会话记忆整理完成: importance={promotedByImportance}, access={promotedByAccess}, conflicts={conflicts}, session={sessionId}")
                     Catch ex As Exception
                         Debug.WriteLine($"[MemoryService] 会话记忆整理失败: {ex.Message}")
                     End Try
                 End Sub)
    End Sub

    ''' <summary>
    ''' 保存一轮对话（user + assistant）为两条独立的原子记忆，含 embedding 和改进的去重。
    ''' </summary>
    Public Shared Sub SaveConversationTurnAsync(userContent As String, assistantContent As String, sessionId As String, Optional appType As String = Nothing)
        Task.Run(Async Function()
                     Try
                         Dim maxLen = MemoryConfig.AtomicContentMaxLength

                         ' 保存 user 消息
                         If Not String.IsNullOrWhiteSpace(userContent) Then
                             Dim uTrimmed = userContent.Trim()
                             Dim uStore = If(uTrimmed.Length > maxLen, uTrimmed.Substring(0, maxLen), uTrimmed)
                             If uStore.Length >= 10 AndAlso Not IsDuplicate(sessionId, uStore) Then
                                 Dim embJson = Await GenerateEmbeddingJson(uStore)
                                 Dim uImportance = UnifiedMemoryService.CalculateImportance(uStore, "conversation", Nothing)
                                MemoryRepository.InsertMemory(uStore, embJson, sessionId, appType, "short_term", uImportance, "user_query")
                                 Debug.WriteLine($"[MemoryService] 保存 user 记忆，长度: {uStore.Length}, 重要性: {uImportance:F2}")
                             End If
                         End If

                         ' 保存 assistant 回复
                         If Not String.IsNullOrWhiteSpace(assistantContent) Then
                             Dim aTrimmed = assistantContent.Trim()
                             Dim aStore = If(aTrimmed.Length > maxLen, aTrimmed.Substring(0, maxLen), aTrimmed)
                             If aStore.Length >= 10 AndAlso Not IsDuplicate(sessionId, aStore) Then
                                 Dim embJson = Await GenerateEmbeddingJson(aStore)
                                 Dim aImportance = UnifiedMemoryService.CalculateImportance(aStore, "assistant_solution", Nothing)
                                MemoryRepository.InsertMemory(aStore, embJson, sessionId, appType, "short_term", aImportance, "assistant_reply")
                                 Debug.WriteLine($"[MemoryService] 保存 assistant 记忆，长度: {aStore.Length}, 重要性: {aImportance:F2}")
                             End If
                         End If

                         ConsolidateSessionMemoriesAsync(sessionId)
                     Catch ex As Exception
                         Debug.WriteLine($"SaveConversationTurnAsync 失败: {ex.Message}")
                     End Try
                 End Function)
    End Sub

    Private Shared Function IsDuplicate(sessionId As String, content As String) As Boolean
        Try
            Dim prefix = If(content.Length > 50, content.Substring(0, 50), content)
            Return MemoryRepository.ExistsMemoryWithPrefix(sessionId, prefix, "short_term")
        Catch
        End Try
        Return False
    End Function

    Private Shared Async Function GenerateEmbeddingJson(text As String) As Task(Of String)
        Try
            If Not EmbeddingService.IsEmbeddingAvailable() Then Return Nothing
            Dim embedding = Await EmbeddingService.GetEmbeddingAsync(text)
            If embedding IsNot Nothing Then
                Return EmbeddingService.SerializeVector(embedding)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] 生成记忆向量失败: {ex.Message}")
        End Try
        Return Nothing
    End Function
End Class
