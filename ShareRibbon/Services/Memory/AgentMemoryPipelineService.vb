' ShareRibbon\Services\Memory\AgentMemoryPipelineService.vb
' Incremental processor for sidecar memory jobs.

Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Public Class AgentMemoryPipelineService
    Private Shared ReadOnly _processLock As New Object()
    Private Shared _isProcessing As Boolean = False
    Private Shared _extractor As IMemoryExtractor = New LlmMemoryExtractor()

    Public Shared Property Extractor As IMemoryExtractor
        Get
            Return _extractor
        End Get
        Set(value As IMemoryExtractor)
            _extractor = If(value, New LlmMemoryExtractor())
        End Set
    End Property

    Public Shared Sub KickoffPendingJobs(Optional limit As Integer = 5)
        SyncLock _processLock
            If _isProcessing Then Return
            _isProcessing = True
        End SyncLock

        Task.Run(Async Function()
                     Try
                         Dim reclaimed = AgentMemoryRepository.ReclaimStaleProcessingJobs()
                         If reclaimed > 0 Then
                             Debug.WriteLine($"[AgentMemoryPipeline] 回收 {reclaimed} 个卡住的 processing job")
                         End If
                         If EmbeddingService.IsEmbeddingAvailable() Then
                             Dim queued = AgentMemoryRepository.EnqueuePendingEmbeddingJobs(limit)
                             If queued > 0 Then
                                 Debug.WriteLine($"[AgentMemoryPipeline] 已补充 {queued} 个 embedding job")
                             End If
                         End If
                         Await ProcessPendingJobsAsync(limit)
                     Catch ex As Exception
                         Debug.WriteLine($"[AgentMemoryPipeline] 后台处理失败: {ex.Message}")
                     Finally
                         SyncLock _processLock
                             _isProcessing = False
                         End SyncLock
                     End Try
                 End Function)
    End Sub

    Public Shared Async Function ProcessPendingJobsAsync(Optional limit As Integer = 10) As Task
        Dim processed As Integer = 0
        Dim maxItems = Math.Max(1, limit)

        While processed < maxItems
            Dim batchSize = Math.Min(5, maxItems - processed)
            Dim jobs = AgentMemoryRepository.ClaimPendingJobs(batchSize)
            If jobs Is Nothing OrElse jobs.Count = 0 Then Exit While

            For Each job In jobs
                Try
                    Select Case If(job.JobType, "").Trim().ToLowerInvariant()
                        Case "extract_memory"
                            Await ProcessExtractMemoryJobAsync(job)
                        Case "embed_memory", "rebuild_embedding"
                            Await ProcessEmbeddingJobAsync(job)
                        Case Else
                            Debug.WriteLine($"[AgentMemoryPipeline] 未支持的任务类型: {job.JobType}")
                    End Select

                    AgentMemoryRepository.MarkJobCompleted(job.JobId)
                Catch ex As Exception
                    AgentMemoryRepository.MarkJobFailed(job.JobId, ex.Message)
                    Debug.WriteLine($"[AgentMemoryPipeline] 任务失败 {job.JobId}: {ex.Message}")
                End Try

                processed += 1
                If processed >= maxItems Then Exit For
            Next
        End While
    End Function

    Private Shared Async Function ProcessExtractMemoryJobAsync(job As MemoryJobRecord) As Task
        Dim payload = ParsePayload(job.PayloadJson)
        Dim sessionId = If(Not String.IsNullOrWhiteSpace(job.TargetId), job.TargetId, GetPayloadText(payload, "session_id"))

        Dim events = AgentMemoryRepository.GetUnprocessedConversationEventsBySession(sessionId)
        If events Is Nothing OrElse events.Count = 0 Then Return

        ' Anchor derived memories to the newest assistant event so a session delete can trace and
        ' clean them through source_event_id even when the extractor did not populate it.
        Dim anchorEventId = events.
            LastOrDefault(Function(e) String.Equals(e.Role, "assistant", StringComparison.OrdinalIgnoreCase))?.EventId
        If String.IsNullOrWhiteSpace(anchorEventId) Then anchorEventId = events(events.Count - 1).EventId

        Dim memories = Extractor.ExtractMemories(events, job.PayloadJson)
        If memories IsNot Nothing Then
            For Each memory In memories
                If String.IsNullOrWhiteSpace(memory.SourceEventId) Then memory.SourceEventId = anchorEventId
                Dim memoryId = AgentMemoryRepository.UpsertMemoryItemMerged(memory)
                Dim embedPayload As New JObject()
                embedPayload("memory_id") = memoryId
                AgentMemoryRepository.EnqueueJob(New MemoryJobRecord With {
                    .JobType = "embed_memory",
                    .TargetId = memoryId,
                    .PayloadJson = embedPayload.ToString(Newtonsoft.Json.Formatting.None),
                    .Status = "pending"
                })
            Next
        End If

        For Each evt In events
            AgentMemoryRepository.MarkConversationEventProcessed(evt.EventId)
        Next

        Await Task.CompletedTask
    End Function

    Private Shared Async Function ProcessEmbeddingJobAsync(job As MemoryJobRecord) As Task
        Dim payload = ParsePayload(job.PayloadJson)
        Dim memoryId = If(Not String.IsNullOrWhiteSpace(job.TargetId), job.TargetId, GetPayloadText(payload, "memory_id"))
        If String.IsNullOrWhiteSpace(memoryId) Then
            Throw New InvalidOperationException("embed_memory 缺少 memory_id")
        End If

        Dim memory = AgentMemoryRepository.GetMemoryItemById(memoryId)
        If memory Is Nothing OrElse String.IsNullOrWhiteSpace(memory.Content) Then
            Throw New InvalidOperationException("未找到可向量化的 memory_item")
        End If

        If Not EmbeddingService.IsEmbeddingAvailable() Then
            AgentMemoryRepository.UpsertMemoryEmbedding(New MemoryEmbeddingRecord With {
                .MemoryId = memoryId,
                .EmbeddingModel = EmbeddingService.GetConfiguredEmbeddingModelName(),
                .EmbeddingDim = 0,
                .EmbeddingJson = Nothing,
                .VectorStatus = "pending",
                .LastError = "Embedding 配置不可用"
            })
            Return
        End If

        Dim embedding = Await EmbeddingService.GetEmbeddingAsync(memory.Content)
        If embedding Is Nothing OrElse embedding.Length = 0 Then
            AgentMemoryRepository.UpsertMemoryEmbedding(New MemoryEmbeddingRecord With {
                .MemoryId = memoryId,
                .EmbeddingModel = EmbeddingService.GetConfiguredEmbeddingModelName(),
                .EmbeddingDim = 0,
                .EmbeddingJson = Nothing,
                .VectorStatus = "failed",
                .LastError = "Embedding 生成失败"
            })
            Return
        End If

        AgentMemoryRepository.UpsertMemoryEmbedding(New MemoryEmbeddingRecord With {
            .MemoryId = memoryId,
            .EmbeddingModel = EmbeddingService.GetConfiguredEmbeddingModelName(),
            .EmbeddingDim = embedding.Length,
            .EmbeddingJson = EmbeddingService.SerializeVector(embedding),
            .VectorStatus = "ready"
        })
    End Function

    Private Shared Function ParsePayload(payloadJson As String) As JObject
        If String.IsNullOrWhiteSpace(payloadJson) Then Return New JObject()
        Try
            Return JObject.Parse(payloadJson)
        Catch
            Return New JObject()
        End Try
    End Function

    Private Shared Function GetPayloadText(payload As JObject, key As String) As String
        If payload Is Nothing OrElse payload(key) Is Nothing Then Return ""
        Return payload(key).ToString()
    End Function
End Class
