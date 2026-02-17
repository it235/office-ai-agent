' ShareRibbon\Controls\Services\MemoryService.vb
' 记忆服务：封装 RAG、用户画像、会话摘要、异步写入

Imports System.Threading.Tasks

''' <summary>
''' 记忆服务：被动 RAG、用户画像、近期会话摘要、异步原子记忆写入
''' </summary>
Public Class MemoryService

    ''' <summary>
    ''' 被动 RAG：按 query 检索 top-N 条相关原子记忆
    ''' </summary>
    Public Shared Function GetRelevantMemories(query As String, Optional topN As Integer? = Nothing, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing) As List(Of AtomicMemoryRecord)
        Dim n = If(topN.HasValue, topN.Value, MemoryConfig.RagTopN)
        Return MemoryRepository.GetRelevantMemories(query, n, startTime, endTime)
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
        Task.Run(Sub()
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

                         MemoryRepository.InsertAtomicMemory(candidate, Nothing, sessionId, appType)
                     Catch ex As Exception
                         Debug.WriteLine($"SaveAtomicMemoryAsync 失败: {ex.Message}")
                     End Try
                 End Sub)
    End Sub

    ''' <summary>
    ''' 主动 RAG 工具：按 keyword 和可选时间范围检索
    ''' </summary>
    Public Shared Function SearchMemories(keyword As String, Optional startTime As DateTime? = Nothing, Optional endTime As DateTime? = Nothing) As List(Of AtomicMemoryRecord)
        Return MemoryRepository.GetRelevantMemories(keyword, MemoryConfig.RagTopN, startTime, endTime)
    End Function

    ''' <summary>
    ''' 插入会话摘要
    ''' </summary>
    Public Shared Sub SaveSessionSummary(sessionId As String, title As String, snippet As String)
        MemoryRepository.InsertSessionSummary(sessionId, title, snippet)
    End Sub
End Class
