' ShareRibbon\Controls\Services\MemoryService.vb
' 记忆服务：封装 RAG、用户画像、会话摘要、异步写入

Imports System.Threading.Tasks

''' <summary>
''' 记忆服务：被动 RAG、用户画像、近期会话摘要、异步原子记忆写入
''' </summary>
Public Class MemoryService

    ''' <summary>
    ''' 被动 RAG：按 query 检索 top-N 条相关原子记忆（使用向量相似度）
    ''' </summary>
    Public Shared Function GetRelevantMemories(query As String, Optional topN As Integer? = Nothing, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing, Optional appType As String = Nothing) As List(Of AtomicMemoryRecord)
        Dim n = If(topN.HasValue, topN.Value, MemoryConfig.RagTopN)

        Dim queryEmbedding As Single() = Nothing
        Try
            ' 优化：仅在数据库中确实存在带 embedding 的记忆时才调用向量 API，避免浪费
            If Not String.IsNullOrWhiteSpace(query) AndAlso EmbeddingService.IsEmbeddingAvailable() AndAlso
               MemoryRepository.HasMemoriesWithEmbedding(appType) Then
                Debug.WriteLine($"[MemoryService] 检测到带 embedding 的记忆，生成查询向量...")
                ' 用 Task.Run 将 async 任务派发到线程池线程执行：
                ' 直接在同步方法里对 async Task 调用 .Wait() 会捕获当前 SynchronizationContext，
                ' 若调用方在 UI 线程（WinForms/VSTO），极易造成死锁。
                ' Task.Run 的包装让 async 代码运行在无 SynchronizationContext 的线程池线程，
                ' .Wait(timeout) 只是普通的线程阻塞，不会死锁。
                Dim embTask = Task.Run(Async Function()
                    Return Await EmbeddingService.GetEmbeddingAsync(query)
                End Function)
                If embTask.Wait(3000) Then
                    queryEmbedding = embTask.Result
                    If queryEmbedding IsNot Nothing Then
                        Debug.WriteLine($"[MemoryService] 查询向量生成成功，维度: {queryEmbedding.Length}")
                    End If
                Else
                    Debug.WriteLine($"[MemoryService] 查询向量生成超时或失败")
                End If
            ElseIf Not EmbeddingService.IsEmbeddingAvailable() Then
                Debug.WriteLine($"[MemoryService] Embedding 不可用，直接使用关键词检索")
            Else
                Debug.WriteLine($"[MemoryService] 数据库中无带 embedding 的记忆，跳过向量 API 调用")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] 生成查询向量失败: {ex.Message}")
        End Try

        Return MemoryRepository.GetRelevantMemories(query, n, queryEmbedding, startTime, endTime, appType)
    End Function

    ''' <summary>
    ''' 获取用户画像
    ''' </summary>
    Public Shared Function GetUserProfile() As String
        If Not MemoryConfig.EnableUserProfile Then Return ""
        Return MemoryRepository.GetUserProfile()
    End Function

    ''' <summary>
    ''' 获取近期会话摘要
    ''' </summary>
    Public Shared Function GetRecentSessionSummaries(Optional limit As Integer? = Nothing) As List(Of SessionSummaryRecord)
        Dim n = If(limit.HasValue, limit.Value, MemoryConfig.SessionSummaryLimit)
        Return MemoryRepository.GetRecentSessionSummaries(n)
    End Function

    ''' <summary>
    ''' 异步写入原子记忆（fire-and-forget），含简单去重。appType 为当前宿主（Excel/Word/PowerPoint），用于按应用筛选展示。
    ''' </summary>
    Public Shared Sub SaveAtomicMemoryAsync(userPrompt As String, assistantReply As String, sessionId As String, Optional appType As String = Nothing)
        Task.Run(Async Function()
                     Try
                         If String.IsNullOrWhiteSpace(userPrompt) AndAlso String.IsNullOrWhiteSpace(assistantReply) Then Return

                         ' 简化：取 user 前 N 字 + assistant 前 M 字 作为候选 content
                         Dim maxLen = MemoryConfig.AtomicContentMaxLength
                         Dim u = (If(userPrompt, "").Trim())
                         Dim a = (If(assistantReply, "").Trim())
                         Dim uPart = If(u.Length > maxLen \ 2, u.Substring(0, maxLen \ 2), u)
                         Dim aPart = If(a.Length > maxLen \ 2, a.Substring(0, maxLen \ 2), a)
                         Dim candidate = uPart & " | " & aPart
                         If String.IsNullOrWhiteSpace(candidate) OrElse candidate.Length < 10 Then Return

                         ' 简单去重：若已有相似 content（LIKE）则跳过
                         Dim existing = MemoryRepository.GetRelevantMemories(candidate.Substring(0, Math.Min(20, candidate.Length)), 3)
                         For Each ex In existing
                             Dim exC = If(ex.Content, "")
                             Dim subC = If(candidate.Length > 30, candidate.Substring(0, 30), candidate)
                             Dim subEx = If(exC.Length > 30, exC.Substring(0, 30), exC)
                             If (exC.Length > 0 AndAlso exC.Contains(subC)) OrElse (subEx.Length > 0 AndAlso candidate.Contains(subEx)) Then
                                 Return
                             End If
                         Next

                         ' 异步生成向量嵌入
                         Dim embeddingJson As String = Nothing
                         Try
                             Debug.WriteLine($"[MemoryService] 正在生成记忆向量...")
                             Dim embedding = Await EmbeddingService.GetEmbeddingAsync(candidate)
                             If embedding IsNot Nothing Then
                                 embeddingJson = EmbeddingService.SerializeVector(embedding)
                                 Debug.WriteLine($"[MemoryService] 记忆向量生成成功，维度: {embedding.Length}")
                             End If
                         Catch vecEx As Exception
                             Debug.WriteLine($"[MemoryService] 生成记忆向量失败: {vecEx.Message}")
                         End Try

                         MemoryRepository.InsertAtomicMemory(candidate, Nothing, sessionId, appType, embeddingJson)
                         Debug.WriteLine($"[MemoryService] 原子记忆已保存，长度: {candidate.Length}, 有向量: {Not String.IsNullOrWhiteSpace(embeddingJson)}")
                     Catch ex As Exception
                         Debug.WriteLine($"SaveAtomicMemoryAsync 失败: {ex.Message}")
                     End Try
                 End Function)
    End Sub

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

            ' 简单去重
            Dim existing = MemoryRepository.GetRelevantMemories(candidate.Substring(0, Math.Min(20, candidate.Length)), 3)
            For Each ex In existing
                Dim exC = If(ex.Content, "")
                Dim subC = If(candidate.Length > 30, candidate.Substring(0, 30), candidate)
                Dim subEx = If(exC.Length > 30, exC.Substring(0, 30), exC)
                If (exC.Length > 0 AndAlso exC.Contains(subC)) OrElse (subEx.Length > 0 AndAlso candidate.Contains(subEx)) Then
                    Return
                End If
            Next

            MemoryRepository.InsertAtomicMemory(candidate, Nothing, sessionId, appType)
            Debug.WriteLine($"[MemoryService] 已同步保存文件内容到记忆，长度: {candidate.Length}")
        Catch ex As Exception
            Debug.WriteLine($"SaveFileContentToMemory 失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 主动 RAG 工具：按 keyword 和可选时间范围检索
    ''' </summary>
    Public Shared Function SearchMemories(keyword As String, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing, Optional appType As String = Nothing) As List(Of AtomicMemoryRecord)
        Dim queryEmbedding As Single() = Nothing
        Try
            ' 优化：仅在数据库中存在带 embedding 的记忆时才调用向量 API
            If Not String.IsNullOrWhiteSpace(keyword) AndAlso EmbeddingService.IsEmbeddingAvailable() AndAlso
               MemoryRepository.HasMemoriesWithEmbedding(appType) Then
                ' 同 GetRelevantMemories：用 Task.Run 包装避免在 UI 线程上死锁
                Dim embTask = Task.Run(Async Function()
                    Return Await EmbeddingService.GetEmbeddingAsync(keyword)
                End Function)
                If embTask.Wait(3000) Then
                    queryEmbedding = embTask.Result
                End If
            Else
                Debug.WriteLine($"[MemoryService] SearchMemories: 无带 embedding 记忆或 embedding 不可用，跳过向量 API")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] SearchMemories 生成向量失败: {ex.Message}")
        End Try

        Return MemoryRepository.GetRelevantMemories(keyword, MemoryConfig.RagTopN, queryEmbedding, startTime, endTime, appType)
    End Function

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
                                 MemoryRepository.InsertAtomicMemory(uStore, Nothing, sessionId, appType, embJson, "short_term")
                                 Debug.WriteLine($"[MemoryService] 保存 user 记忆，长度: {uStore.Length}, 有向量: {Not String.IsNullOrWhiteSpace(embJson)}")
                             End If
                         End If

                         ' 保存 assistant 回复
                         If Not String.IsNullOrWhiteSpace(assistantContent) Then
                             Dim aTrimmed = assistantContent.Trim()
                             Dim aStore = If(aTrimmed.Length > maxLen, aTrimmed.Substring(0, maxLen), aTrimmed)
                             If aStore.Length >= 10 AndAlso Not IsDuplicate(sessionId, aStore) Then
                                 Dim embJson = Await GenerateEmbeddingJson(aStore)
                                 MemoryRepository.InsertAtomicMemory(aStore, Nothing, sessionId, appType, embJson, "short_term")
                                 Debug.WriteLine($"[MemoryService] 保存 assistant 记忆，长度: {aStore.Length}, 有向量: {Not String.IsNullOrWhiteSpace(embJson)}")
                             End If
                         End If
                     Catch ex As Exception
                         Debug.WriteLine($"SaveConversationTurnAsync 失败: {ex.Message}")
                     End Try
                 End Function)
    End Sub

    Private Shared Function IsDuplicate(sessionId As String, content As String) As Boolean
        Try
            Dim prefix = If(content.Length > 50, content.Substring(0, 50), content)
            Dim existing = MemoryRepository.GetRelevantMemories(prefix, 3)
            For Each ex In existing
                If ex.SessionId = sessionId AndAlso Not String.IsNullOrWhiteSpace(ex.Content) Then
                    Dim exPrefix = If(ex.Content.Length > 50, ex.Content.Substring(0, 50), ex.Content)
                    If exPrefix = prefix Then Return True
                End If
            Next
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

    ''' <summary>
    ''' 插入会话摘要
    ''' </summary>
    Public Shared Sub SaveSessionSummary(sessionId As String, title As String, snippet As String)
        MemoryRepository.InsertSessionSummary(sessionId, title, snippet)
    End Sub
End Class
