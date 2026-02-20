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
    Public Shared Function GetRelevantMemories(query As String, Optional topN As Integer? = Nothing, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing) As List(Of AtomicMemoryRecord)
        Dim n = If(topN.HasValue, topN.Value, MemoryConfig.RagTopN)
        
        ' 首先尝试生成查询向量（同步等待，为了兼容现有调用）
        Dim queryEmbedding As Single() = Nothing
        Try
            If Not String.IsNullOrWhiteSpace(query) Then
                Debug.WriteLine($"[MemoryService] 正在生成查询向量...")
                ' 使用同步等待的方式调用异步方法（这不是最佳实践，但为了兼容性）
                Dim task = EmbeddingService.GetEmbeddingAsync(query)
                task.Wait(TimeSpan.FromSeconds(10))
                If task.IsCompleted Then
                    queryEmbedding = task.Result
                    If queryEmbedding IsNot Nothing Then
                        Debug.WriteLine($"[MemoryService] 查询向量生成成功，维度: {queryEmbedding.Length}")
                    End If
                Else
                    Debug.WriteLine($"[MemoryService] 查询向量生成超时或失败")
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] 生成查询向量失败: {ex.Message}")
        End Try
        
        ' 调用 Repository 进行检索（会自动判断是否有向量）
        Return MemoryRepository.GetRelevantMemories(query, n, queryEmbedding, startTime, endTime)
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
    Public Shared Function SearchMemories(keyword As String, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing) As List(Of AtomicMemoryRecord)
        Dim queryEmbedding As Single() = Nothing
        Try
            If Not String.IsNullOrWhiteSpace(keyword) Then
                Dim taskx = EmbeddingService.GetEmbeddingAsync(keyword)
                taskx.Wait(TimeSpan.FromSeconds(10))
                If taskx.IsCompleted Then
                    queryEmbedding = taskx.Result
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[MemoryService] SearchMemories 生成向量失败: {ex.Message}")
        End Try
        
        Return MemoryRepository.GetRelevantMemories(keyword, MemoryConfig.RagTopN, queryEmbedding, startTime, endTime)
    End Function

    ''' <summary>
    ''' 插入会话摘要
    ''' </summary>
    Public Shared Sub SaveSessionSummary(sessionId As String, title As String, snippet As String)
        MemoryRepository.InsertSessionSummary(sessionId, title, snippet)
    End Sub
End Class
