' ShareRibbon\Services\UnifiedMemoryService.vb
' 统一记忆管理服务 - 整合原子记忆与Ralph Loop记忆

Imports System.Threading.Tasks

''' <summary>
''' 统一记忆管理服务
''' </summary>
Public Class UnifiedMemoryService

    ''' <summary>
    ''' 保存记忆并自动计算重要性
    ''' </summary>
    Public Shared Async Function SaveMemoryAsync(
        content As String,
        memoryType As String,
        sessionId As String,
        Optional appType As String = Nothing,
        Optional metadata As Dictionary(Of String, Object) = Nothing) As Task(Of Long)

        ' 1. 计算重要性（同步计算，无IO操作）
        Dim importance = CalculateImportance(content, memoryType, metadata)

        ' 2. 生成向量嵌入
        Dim embeddingJson As String = Nothing
        Try
            If Not String.IsNullOrWhiteSpace(content) AndAlso EmbeddingService.IsEmbeddingAvailable() Then
                Debug.WriteLine($"[UnifiedMemoryService] 正在生成记忆向量...")
                Dim embedding = Await EmbeddingService.GetEmbeddingAsync(content)
                If embedding IsNot Nothing Then
                    embeddingJson = EmbeddingService.SerializeVector(embedding)
                    Debug.WriteLine($"[UnifiedMemoryService] 记忆向量生成成功，维度: {embedding.Length}")
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[UnifiedMemoryService] 生成向量失败: {ex.Message}")
        End Try

        ' 3. 保存到数据库
        Dim memoryId = MemoryRepository.InsertMemory(content, embeddingJson, sessionId, appType, memoryType, importance)

        ' 4. 异步建立知识关联（在后台线程执行）
        If memoryId > 0 Then
            Task.Run(Sub() BuildMemoryAssociations(memoryId, content))
        End If

        Return memoryId
    End Function

    ''' <summary>
    ''' 计算记忆重要性（基于内容、用户反馈、访问频率）
    ''' </summary>
    Private Shared Function CalculateImportance(
        content As String,
        memoryType As String,
        metadata As Dictionary(Of String, Object)) As Double

        Dim baseScore As Double = 0.5

        ' 1. 根据类型调整基础分
        Select Case memoryType?.ToLowerInvariant()
            Case "user_explicit_intent"
                baseScore = 0.9
            Case "assistant_solution", "task_result"
                baseScore = 0.8
            Case "user_feedback"
                baseScore = 0.85
            Case "knowledge"
                baseScore = 0.75
            Case "skill_feedback"
                baseScore = 0.7
        End Select

        ' 2. 内容特征分析
        If Not String.IsNullOrWhiteSpace(content) Then
            If content.Length > 100 Then baseScore += 0.1
            If content.Contains("?") OrElse content.Contains("如何") OrElse content.Contains("怎么") Then baseScore += 0.05
            If content.Contains("!") OrElse content.Contains("重要") OrElse content.Contains("关键") Then baseScore += 0.08
        End If

        ' 3. 元数据调整
        If metadata IsNot Nothing Then
            If metadata.ContainsKey("user_explicit_save") AndAlso CBool(metadata("user_explicit_save")) Then
                baseScore += 0.15
            End If
            If metadata.ContainsKey("user_rating") Then
                Dim rating = Convert.ToInt32(metadata("user_rating"))
                baseScore += (rating - 3) * 0.1  ' 5星+0.2，3星0，1星-0.2
            End If
        End If

        Return Math.Min(1.0, Math.Max(0.1, baseScore))
    End Function

    ''' <summary>
    ''' 建立记忆关联（后台线程执行）
    ''' </summary>
    Private Shared Sub BuildMemoryAssociations(memoryId As Long, content As String)
        Try
            ' 1. 查找相似记忆
            Dim similarMemories = MemoryRepository.GetSimilarMemories(memoryId, topK:=5)

            ' 2. 创建关联边
            For Each similar In similarMemories
                If similar.Id <> memoryId AndAlso similar.SimilarityScore > 0.7 Then
                    MemoryRepository.AddMemoryRelation(
                        memoryId,
                        similar.Id,
                        "similar",
                        similar.SimilarityScore
                    )
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"[UnifiedMemoryService] 建立记忆关联失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 智能检索记忆（混合关键词+向量+重要性+时间衰减）
    ''' </summary>
    Public Shared Function RetrieveMemories(
        query As String,
        Optional topN As Integer = 10,
        Optional sessionId As String = Nothing,
        Optional startTime As DateTime? = Nothing,
        Optional endTime As DateTime? = Nothing,
        Optional appType As String = Nothing) As List(Of MemoryWithScore)

        ' 1. 获取查询向量
        Dim queryEmbedding As Single() = Nothing
        Try
            If Not String.IsNullOrWhiteSpace(query) AndAlso EmbeddingService.IsEmbeddingAvailable() AndAlso
               MemoryRepository.HasMemoriesWithEmbedding(appType) Then
                Debug.WriteLine($"[UnifiedMemoryService] 生成查询向量...")
                Dim embTask = Task.Run(Function() EmbeddingService.GetEmbeddingAsync(query))
                If embTask.Wait(3000) Then
                    queryEmbedding = embTask.Result
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"[UnifiedMemoryService] 生成查询向量失败: {ex.Message}")
        End Try

        ' 2. 获取候选记忆
        Dim candidates = MemoryRepository.GetRelevantMemories(query, topN * 3, queryEmbedding, startTime, endTime, appType)

        ' 3. 综合评分
        Dim results As New List(Of MemoryWithScore)()
        Dim now = DateTime.UtcNow
        Dim accessedIds As New List(Of Long)()

        For Each candidate In candidates
            ' 收集需要更新访问计数的ID（批量更新）
            accessedIds.Add(candidate.Id)

            ' a. 向量相似度（如果有）
            Dim similarity = candidate.SimilarityScore
            If similarity = 0 Then similarity = 0.3  ' 默认基础分

            ' b. 时间衰减（半衰期30天）
            Dim createDateTime = DateTime.UtcNow
            Try
                Dim ts = New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(candidate.Timestamp)
                createDateTime = ts
            Catch
            End Try
            Dim hoursSinceCreation = Math.Max(0, (now - createDateTime).TotalHours)
            Dim timeDecay = Math.Exp(-hoursSinceCreation / 720.0)  ' 30天 = 720小时

            ' c. 重要性分数
            Dim importance = candidate.Importance
            If importance = 0 Then importance = 0.5

            ' d. 访问频率（对数衰减）
            Dim accessBoost = Math.Log(1 + candidate.AccessCount) * 0.05

            ' e. 会话相关性
            Dim sessionBoost = If(candidate.SessionId = sessionId, 0.2, 0)

            ' 综合得分
            Dim totalScore = (similarity * 0.4) +
                           (importance * 0.3) +
                           (timeDecay * 0.15) +
                           (accessBoost * 0.1) +
                           (sessionBoost * 0.05)

            results.Add(New MemoryWithScore With {
                .Memory = candidate,
                .Score = totalScore,
                .Components = New Dictionary(Of String, Double) From {
                    {"similarity", similarity},
                    {"importance", importance},
                    {"timeDecay", timeDecay},
                    {"accessBoost", accessBoost},
                    {"sessionBoost", sessionBoost}
                }
            })
        Next

        ' 4. 批量更新访问计数（避免N+1问题）
        If accessedIds.Count > 0 Then
            Task.Run(Sub()
                         Try
                             For Each id In accessedIds
                                 MemoryRepository.UpdateMemoryAccess(id)
                             Next
                         Catch ex As Exception
                             Debug.WriteLine($"[UnifiedMemoryService] 批量更新访问计数失败: {ex.Message}")
                         End Try
                     End Sub)
        End If

        ' 5. 排序返回
        Return results.OrderByDescending(Function(m) m.Score).Take(topN).ToList()
    End Function

    ''' <summary>
    ''' 短期记忆迁移为长期记忆（会话结束时调用）
    ''' </summary>
    Public Shared Sub MigrateShortTermToLongTerm(sessionId As String)
        Try
            MemoryRepository.ExpireLowImportanceMemories(sessionId, threshold:=0.3)
            Debug.WriteLine($"[UnifiedMemoryService] 已清理会话 {sessionId} 的低重要性记忆")
        Catch ex As Exception
            Debug.WriteLine($"[UnifiedMemoryService] 迁移记忆失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 更新用户画像（基于对话历史学习）
    ''' </summary>
    Public Shared Sub UpdateUserProfile(
        recentConversations As List(Of ConversationHistoryItem))

        Try
            ' 分析用户行为模式
            Dim observations = AnalyzeUserBehavior(recentConversations)

            For Each obs In observations
                ' 更新或添加画像项
                MemoryRepository.UpsertUserProfile(
                    obs.Key,
                    obs.Value,
                    obs.Category,
                    obs.Confidence
                )
            Next

            Debug.WriteLine($"[UnifiedMemoryService] 已更新用户画像，共 {observations.Count} 项")
        Catch ex As Exception
            Debug.WriteLine($"[UnifiedMemoryService] 更新用户画像失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 分析用户行为模式
    ''' </summary>
    Private Shared Function AnalyzeUserBehavior(
        recentConversations As List(Of ConversationHistoryItem)) As List(Of UserProfileObservation)

        Dim observations As New List(Of UserProfileObservation)()

        If recentConversations Is Nothing OrElse recentConversations.Count = 0 Then
            Return observations
        End If

        ' 简单启发式分析
        Dim userMessages = recentConversations.Where(Function(c) c.Role = "user").ToList()
        Dim assistantMessages = recentConversations.Where(Function(c) c.Role = "assistant").ToList()

        ' 统计应用使用频率
        Dim appCounts As New Dictionary(Of String, Integer)()
        For Each msg In recentConversations
            If Not String.IsNullOrWhiteSpace(msg.AppType) Then
                If Not appCounts.ContainsKey(msg.AppType) Then appCounts(msg.AppType) = 0
                appCounts(msg.AppType) += 1
            End If
        Next

        If appCounts.Count > 0 Then
            Dim topApp = appCounts.OrderByDescending(Function(k) k.Value).First()
            observations.Add(New UserProfileObservation With {
                .Key = "preferred_application",
                .Value = topApp.Key,
                .Category = "preference",
                .Confidence = Math.Min(0.95, topApp.Value * 0.1)
            })
        End If

        ' 分析消息长度偏好
        If userMessages.Count > 0 Then
            Dim avgLength = userMessages.Average(Function(m) m.Content.Length)
            Dim complexity = If(avgLength > 200, "detailed", If(avgLength > 50, "moderate", "concise"))
            observations.Add(New UserProfileObservation With {
                .Key = "message_preference",
                .Value = complexity,
                .Category = "preference",
                .Confidence = 0.7
            })
        End If

        ' 领域偏好（关键词检测）
        Dim domainKeywords As New Dictionary(Of String, List(Of String)) From {
            {"excel", New List(Of String) From {"excel", "表格", "vba", "公式", "数据"}},
            {"word", New List(Of String) From {"word", "文档", "排版", "格式"}},
            {"powerpoint", New List(Of String) From {"ppt", "powerpoint", "演示", "幻灯片"}},
            {"programming", New List(Of String) From {"代码", "编程", "程序", "python", "javascript"}},
            {"data_analysis", New List(Of String) From {"分析", "统计", "数据", "图表"}}
        }

        For Each domain In domainKeywords
            Dim matchCount = Enumerable.Count(userMessages, Function(m)
                Dim lower = m.Content.ToLowerInvariant()
                Return domain.Value.Any(Function(k) lower.Contains(k))
            End Function)
            If matchCount > 0 Then
                observations.Add(New UserProfileObservation With {
                    .Key = "domains",
                    .Value = domain.Key,
                    .Category = "interest",
                    .Confidence = Math.Min(0.8, matchCount * 0.15)
                })
            End If
        Next

        Return observations
    End Function

End Class

''' <summary>
''' 对话历史项
''' </summary>
Public Class ConversationHistoryItem
    Public Property Role As String
    Public Property Content As String
    Public Property AppType As String
    Public Property Timestamp As Long
End Class

''' <summary>
''' 用户画像观察结果
''' </summary>
Public Class UserProfileObservation
    Public Property Key As String
    Public Property Value As String
    Public Property Category As String
    Public Property Confidence As Double
End Class
