Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class MemoryTurnRecorder
    Public Sub RecordConversationTurn(
        userQuestion As String,
        assistantAnswer As String,
        sessionId As String,
        responseMode As String,
        addHistory As Boolean,
        appType As String)

        If String.IsNullOrWhiteSpace(userQuestion) AndAlso String.IsNullOrWhiteSpace(assistantAnswer) Then Return

        Dim capturedUserQuestion = If(userQuestion, "")
        Dim capturedAssistantAnswer = If(assistantAnswer, "")
        Dim capturedSessionId = If(String.IsNullOrWhiteSpace(sessionId), Guid.NewGuid().ToString(), sessionId)
        Dim capturedResponseMode = If(responseMode, "")
        Dim capturedAppType = If(appType, "")

        ' ChatStateService.AddMessage 已把每条 user/assistant 消息写入 conversation_event
        ' （event_type=chat_message），这里不再重复 append，只负责触发一次记忆抽取；
        ' 抽取按 session 拾取尚未处理的 chat_message 事件，天然幂等。
        Task.Run(Sub()
                     Try
                         Dim payload As New JObject()
                         payload("session_id") = capturedSessionId
                         payload("app_type") = capturedAppType
                         payload("response_mode") = capturedResponseMode
                         payload("has_user_content") = Not String.IsNullOrWhiteSpace(capturedUserQuestion)
                         payload("has_assistant_content") = Not String.IsNullOrWhiteSpace(capturedAssistantAnswer)

                         AgentMemoryRepository.EnqueueJob(New MemoryJobRecord With {
                             .JobType = "extract_memory",
                             .TargetId = capturedSessionId,
                             .PayloadJson = payload.ToString(Formatting.None),
                             .Status = "pending"
                         })
                         AgentMemoryPipelineService.KickoffPendingJobs()
                     Catch ex As Exception
                         Debug.WriteLine($"MemoryTurnRecorder failed: {ex.Message}")
                     End Try
                 End Sub)
    End Sub
End Class
