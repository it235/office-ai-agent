' ShareRibbon\Controls\Services\ChatRoutingOrchestrator.vb
' Extracts smart-mode routing (analyze → precheck → agent/chat) from BaseChatControl.

Imports System.Diagnostics
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' Host callbacks required by smart-mode chat routing.
''' Implemented by BaseChatControl so the orchestrator stays free of UI inheritance.
''' </summary>
Public Interface IChatRoutingHost
    Function GetHistoryMessages() As List(Of HistoryMessage)
    Function IsFollowUpQuestionAsync(question As String, history As List(Of HistoryMessage)) As Task(Of Boolean)
    Function GetContextSnapshot() As JObject
    Sub EnrichContextForIntent(snapshot As JObject,
                               originalQuestion As String,
                               filePaths As List(Of String),
                               selectedContents As List(Of SendMessageReferenceContentItem))
    Function GetApplicationType() As String
    Function CaptureOfficeContext(appType As String) As Agent.Context.OfficeContext
    Function AnalyzeAiNativeAsync(request As Agent.AiNativeRequest) As Task(Of Agent.AiNativeRuntimeResult)
    Function BuildExecutionContextAsync(originalQuestion As String,
                                        filePaths As List(Of String),
                                        selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ExecutionContext)
    Function PreSendCheckAsync(execContext As ExecutionContext) As Task(Of ContextCheckResult)
    Sub ShowContextHints(ragCount As Integer, intentDescription As String, contextTrace As ChatContextTrace)
    Sub ShowWarning(message As String)
    Sub SendChatMessage(message As String)
    Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
    Sub StartAgentPlanningFlow(message As String, intent As IntentResult, analysis As Agent.AiNativeRuntimeResult)
    Sub SetCurrentIntentResult(intent As IntentResult)
End Interface

''' <summary>
''' Decision outcome of smart-mode routing (for tests / diagnostics).
''' </summary>
Public Enum ChatRouteDecision
    FollowUpChat
    PlainChat
    BlockedByPreCheck
    AgentKernel
    ''' <summary>Only used when agent start fails and policy allows chat fallback, or analyze throws.</summary>
    FallbackChat
End Enum

''' <summary>
''' Coordinates AI Native analysis and the smart-mode route to AgentKernel.
''' Product primary path (P0-2): Analyze → AgentKernel/Loop → Tool/Capability Executor.
''' </summary>
Public Class ChatRoutingOrchestrator
    Private ReadOnly _host As IChatRoutingHost

    Public Sub New(host As IChatRoutingHost)
        If host Is Nothing Then Throw New ArgumentNullException(NameOf(host))
        _host = host
    End Sub

    ''' <summary>
    ''' Smart mode: follow-up detection → AiNative analyze → pre-check → AgentKernel.
    ''' </summary>
    Public Async Function RouteSmartModeAsync(finalMessageToLLM As String,
                                              originalQuestion As String,
                                              filePaths As List(Of String),
                                              selectedContents As List(Of SendMessageReferenceContentItem)) As Task(Of ChatRouteDecision)

        Try
            Dim hasReferences As Boolean =
                (filePaths IsNot Nothing AndAlso filePaths.Count > 0) OrElse
                (selectedContents IsNot Nothing AndAlso selectedContents.Count > 0)

            Dim history = _host.GetHistoryMessages()
            If history Is Nothing Then history = New List(Of HistoryMessage)()

            Dim isFollowUp As Boolean = False
            If history.Count >= 2 AndAlso Not String.IsNullOrWhiteSpace(originalQuestion) Then
                isFollowUp = Await _host.IsFollowUpQuestionAsync(originalQuestion, history)
                Debug.WriteLine($"[ChatRoutingOrchestrator] follow-up check: isFollowUp={isFollowUp}")
            End If

            ' Short follow-ups stay on plain chat to avoid re-planning every clarification turn.
            If isFollowUp Then
                Debug.WriteLine("[ChatRoutingOrchestrator] follow-up → plain chat (not a parallel product path)")
                _host.SendChatMessage(finalMessageToLLM)
                Return ChatRouteDecision.FollowUpChat
            End If

            Dim contextSnapshot = _host.GetContextSnapshot()
            If contextSnapshot Is Nothing Then contextSnapshot = New JObject()
            _host.EnrichContextForIntent(contextSnapshot, originalQuestion, filePaths, selectedContents)

            Dim recentHistory = history.
                Where(Function(m) m.role <> "system" AndAlso Not String.IsNullOrEmpty(m.content)).
                ToList()
            If recentHistory.Count > 0 Then
                Dim historyArray As New JArray()
                Dim takeCount = Math.Min(6, recentHistory.Count)
                For i = recentHistory.Count - takeCount To recentHistory.Count - 1
                    Dim hMsg = recentHistory(i)
                    Dim content = hMsg.content
                    If content IsNot Nothing AndAlso content.Length > 300 Then
                        content = content.Substring(0, 300) & "..."
                    End If
                    historyArray.Add(New JObject From {
                        {"role", hMsg.role},
                        {"content", content}
                    })
                Next
                contextSnapshot("conversationHistory") = historyArray
            End If

            Dim appType = _host.GetApplicationType()
            Dim aiNativeRequest As New Agent.AiNativeRequest With {
                .UserInput = originalQuestion,
                .AppType = appType,
                .SystemPrompt = "",
                .RequestUuid = Guid.NewGuid().ToString(),
                .OfficeContext = _host.CaptureOfficeContext(appType),
                .ContextSnapshot = contextSnapshot,
                .HistoryMessages = recentHistory,
                .EnableMemory = MemoryConfig.EnableUserProfile OrElse MemoryConfig.RagTopN > 0,
                .UseContextBuilder = MemoryConfig.UseContextBuilder
            }

            Dim aiNativeResult = Await _host.AnalyzeAiNativeAsync(aiNativeRequest)
            Dim intent = If(aiNativeResult?.Intent, New IntentResult())
            intent.OriginalInput = originalQuestion
            _host.SetCurrentIntentResult(intent)

            If aiNativeResult IsNot Nothing AndAlso aiNativeResult.ContextTrace IsNot Nothing Then
                _host.ShowContextHints(
                    aiNativeResult.RagCount,
                    If(intent.UserFriendlyDescription, ""),
                    aiNativeResult.ContextTrace)
            End If

            Dim isFormattingOrProofreading As Boolean =
                intent.OfficeIntent = OfficeIntentType.FORMAT_STYLE OrElse
                intent.OfficeIntent = OfficeIntentType.TEXT_FORMAT

            If isFormattingOrProofreading Then
                Try
                    Dim execContext = Await _host.BuildExecutionContextAsync(originalQuestion, filePaths, selectedContents)
                    If execContext IsNot Nothing Then
                        execContext.IntentResult = intent
                        Dim preCheck = Await _host.PreSendCheckAsync(execContext)
                        If preCheck IsNot Nothing AndAlso Not preCheck.IsValid Then
                            Dim errors = String.Join(";", preCheck.Errors)
                            Debug.WriteLine($"[ChatRoutingOrchestrator] PreSendCheck blocked: {errors}")
                            _host.ShowWarning($"请求未通过预检: {errors}")
                            Return ChatRouteDecision.BlockedByPreCheck
                        End If
                        If preCheck IsNot Nothing AndAlso preCheck.Warnings IsNot Nothing AndAlso preCheck.Warnings.Count > 0 Then
                            Debug.WriteLine($"[ChatRoutingOrchestrator] PreSendCheck warnings: {String.Join(";", preCheck.Warnings)}")
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"[ChatRoutingOrchestrator] PreSendCheck exception: {ex.Message}")
                End Try
            End If

            If hasReferences AndAlso String.IsNullOrWhiteSpace(originalQuestion) Then
                intent.UserFriendlyDescription = "已根据引用内容自动识别处理意图"
                _host.SetCurrentIntentResult(intent)
            End If

            Dim interactionMode = If(intent.ResponseMode, "").Trim().ToLowerInvariant()
            Dim taskSpecRequiresExecution = aiNativeResult?.TaskSpec IsNot Nothing AndAlso
                (aiNativeResult.TaskSpec.ExpectedSlideCount > 0 OrElse
                 (aiNativeResult.TaskSpec.ExpectedOutputs IsNot Nothing AndAlso
                  aiNativeResult.TaskSpec.ExpectedOutputs.Count > 0))
            Dim shouldUsePlainChat = interactionMode = "answer" OrElse interactionMode = "clarify" OrElse
                (String.IsNullOrWhiteSpace(interactionMode) AndAlso
                 intent.OfficeIntent = OfficeIntentType.GENERAL_QUERY AndAlso
                 Not taskSpecRequiresExecution)
            If shouldUsePlainChat Then
                Debug.WriteLine($"[ChatRoutingOrchestrator] interactionMode={If(interactionMode, "compat-general")} → plain chat")
                _host.SendChatMessageWithIntent(finalMessageToLLM, intent)
                Return ChatRouteDecision.PlainChat
            End If

            ' Primary product path always uses AgentKernel.
            Debug.WriteLine($"[ChatRoutingOrchestrator] primary path AgentKernel intent={intent.OfficeIntent}, confidence={intent.Confidence:F2}")
            _host.StartAgentPlanningFlow(finalMessageToLLM, intent, aiNativeResult)
            Return ChatRouteDecision.AgentKernel

        Catch ex As Exception
            Debug.WriteLine($"[ChatRoutingOrchestrator] analyze/route failed, fallback chat: {ex.Message}")
            If Agent.ExecutionPathPolicy.AllowChatFallbackOnAgentFailure Then
                _host.SendChatMessage(finalMessageToLLM)
            End If
            Return ChatRouteDecision.FallbackChat
        End Try
    End Function
End Class
