Imports System.Threading
Imports System.Threading.Tasks

Namespace Agent.Harness

    Public Class OfficeHarness
        Implements IOfficeHarness

        Private ReadOnly _kernel As AgentKernel
        Private ReadOnly _runTraceStore As IRunTraceStore
        Private _currentRunId As String = ""
        Private ReadOnly _runLock As New Object()
        Private ReadOnly _pendingRuns As New Dictionary(Of String, PendingHarnessRun)(StringComparer.OrdinalIgnoreCase)

        Public Event PhaseChanged As EventHandler(Of HarnessPhaseChangedEventArgs) Implements IOfficeHarness.PhaseChanged
        Public Event StepChanged As EventHandler(Of HarnessStepChangedEventArgs) Implements IOfficeHarness.StepChanged
        Public Event ContextReady As EventHandler(Of HarnessContextEventArgs) Implements IOfficeHarness.ContextReady

        Public Sub New(kernel As AgentKernel, Optional runTraceStore As IRunTraceStore = Nothing)
            If kernel Is Nothing Then Throw New ArgumentNullException(NameOf(kernel))
            _kernel = kernel
            _runTraceStore = If(runTraceStore, New NoopRunTraceStore())
            AddHandler _kernel.OnStatusChanged, AddressOf HandleKernelStatusChanged
            AddHandler _kernel.OnStepCompleted, AddressOf HandleKernelStepCompleted
            AddHandler _kernel.OnPlanGenerated, AddressOf HandleKernelPlanGenerated
            AddHandler _kernel.OnExecutionExplained, AddressOf HandleKernelExecutionExplained
            AddHandler _kernel.OnRequestApproval, AddressOf HandleKernelRequestApproval
        End Sub

        Public Async Function RunAsync(turn As UserTurn,
                                       cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.RunAsync
            Dim runId = Guid.NewGuid().ToString()
            Dim startedAt = DateTime.Now
            Dim keepPending As Boolean = False

            SyncLock _runLock
                If _pendingRuns.Values.Any(Function(item) Not item.IsCompleted) Then
                    Return New HarnessRunResult With {
                        .RunId = runId,
                        .Status = HarnessRunStatus.Failed,
                        .UserMessage = "当前会话已有正在执行或等待审批的任务",
                        .DebugMessage = "RUN_ALREADY_ACTIVE",
                        .StartedAt = startedAt,
                        .FinishedAt = DateTime.Now
                    }
                End If
            End SyncLock

            If turn Is Nothing Then
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = HarnessRunStatus.Failed,
                    .UserMessage = "用户请求为空",
                    .DebugMessage = "UserTurn is null",
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            End If
            _currentRunId = runId

            Try
                cancellationToken.ThrowIfCancellationRequested()
                SafeStartRun(runId, turn, startedAt)
                RaisePhase(runId, "starting", "Harness run started")

                Dim officeContext = turn.OfficeContext
                If officeContext Is Nothing Then
                    officeContext = New Agent.Context.OfficeContext With {.AppType = turn.AppType}
                End If
                Dim contextPack = If(turn.ContextPack,
                                     Agent.Context.ContextPack.FromOfficeContext(officeContext, turn.HostContextText))
                turn.ContextPack = contextPack

                RaiseEvent ContextReady(Me, New HarnessContextEventArgs With {
                    .RunId = runId,
                    .AppType = turn.AppType,
                    .ContextText = contextPack.ToPromptText(),
                    .ContextPackJson = contextPack.ToJson()
                })

                RaisePhase(runId, "executing", "AgentKernel executing")
                Dim pending As New PendingHarnessRun With {
                    .RunId = runId,
                    .Turn = turn,
                    .StartedAt = startedAt,
                    .ApprovalSignal = New TaskCompletionSource(Of Boolean)(TaskCreationOptions.RunContinuationsAsynchronously)
                }
                SyncLock _runLock
                    _pendingRuns(runId) = pending
                End SyncLock

                pending.AgentTask = _kernel.ExecuteAsync(If(turn.Text, ""),
                                                         If(turn.AppType, ""),
                                                         If(turn.HostContextText, ""),
                                                         officeContext,
                                                         contextPack,
                                                         turn.TaskSpec,
                                                         turn.SelectedSkills,
                                                         cancellationToken)
                Dim first = Await Task.WhenAny(pending.AgentTask, pending.ApprovalSignal.Task)
                If first Is pending.ApprovalSignal.Task AndAlso Not pending.AgentTask.IsCompleted Then
                    keepPending = True
                    SafeSetRunStatus(runId, "awaiting_approval", pending.ApprovalMessage, ExceptionClassifier.CodeSafetyNeedsApproval)
                    RaisePhase(runId, "awaiting_approval", pending.ApprovalMessage)
                    Return New HarnessRunResult With {
                        .RunId = runId,
                        .Status = HarnessRunStatus.AwaitingApproval,
                        .UserMessage = pending.ApprovalMessage,
                        .StartedAt = startedAt
                    }
                End If

                Dim agentResult = Await pending.AgentTask
                pending.IsCompleted = True

                If agentResult IsNot Nothing AndAlso agentResult.IsCancelled Then
                    RaisePhase(runId, "cancelled", "已取消")
                    SafeCompleteRun(runId, "cancelled", "已取消", ExceptionClassifier.CodeCancelled, DateTime.Now)
                    Return New HarnessRunResult With {
                        .RunId = runId,
                        .Status = HarnessRunStatus.Cancelled,
                        .UserMessage = If(agentResult.Message, "已取消"),
                        .DebugMessage = "CANCELLED_AT_SAFE_POINT",
                        .AgentSessionId = If(agentResult.SessionId, ""),
                        .StartedAt = startedAt,
                        .FinishedAt = DateTime.Now
                    }
                End If

                cancellationToken.ThrowIfCancellationRequested()

                Dim succeeded = agentResult IsNot Nothing AndAlso agentResult.Success
                Dim status = If(succeeded, HarnessRunStatus.Succeeded, HarnessRunStatus.Failed)
                Dim message = If(agentResult?.Message, If(succeeded, "执行完成", "执行失败"))

                RaisePhase(runId, If(succeeded, "completed", "failed"), message)
                SafeCompleteRun(runId, If(succeeded, "succeeded", "failed"), message, If(succeeded, "", ExceptionClassifier.CodeUnknown), DateTime.Now)
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = status,
                    .UserMessage = message,
                    .DebugMessage = message,
                    .AgentSessionId = If(agentResult?.SessionId, ""),
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            Catch ex As OperationCanceledException
                RaisePhase(runId, "cancelled", "Harness run cancelled")
                SafeCompleteRun(runId, "cancelled", "已取消", ExceptionClassifier.CodeCancelled, DateTime.Now)
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = HarnessRunStatus.Cancelled,
                    .UserMessage = "已取消",
                    .DebugMessage = ex.Message,
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            Catch ex As Exception
                AppLogger.Error("OfficeHarness", "RunAsync exception", ex)
                Dim classified = ExceptionClassifier.Classify(ex)
                RaisePhase(runId, "failed", classified.UserMessage)
                SafeCompleteRun(runId, "failed", classified.UserMessage, classified.ErrorCode, DateTime.Now)
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = HarnessRunStatus.Failed,
                    .UserMessage = classified.UserMessage,
                    .DebugMessage = classified.DebugDetail,
                    .StartedAt = startedAt,
                    .FinishedAt = DateTime.Now
                }
            Finally
                If Not keepPending Then
                    RemovePendingRun(runId)
                    _currentRunId = ""
                End If
            End Try
        End Function

        Public Async Function ApproveAsync(runId As String,
                                           approved As Boolean,
                                           cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.ApproveAsync
            Dim pending = GetPendingRun(runId)
            If pending Is Nothing OrElse pending.AgentTask Is Nothing Then
                Return MissingRunResult(runId)
            End If

            cancellationToken.ThrowIfCancellationRequested()
            Dim callback = pending.ApprovalCallback
            pending.ApprovalCallback = Nothing
            If callback Is Nothing Then Return MissingRunResult(runId)
            SafeAppendApprovalStep(runId,
                                   -2,
                                   "approval.decision",
                                   If(approved, "approved", "rejected"),
                                   If(approved, "用户已批准高风险操作", "用户已拒绝高风险操作"),
                                   If(approved, "", ExceptionClassifier.CodeSafetyBlocked),
                                   New With {.approved = approved},
                                   DateTime.Now)
            callback(approved)
            SafeSetRunStatus(runId, "running", If(approved, "审批通过，继续执行", "审批拒绝，正在收敛执行结果"), "")

            Dim agentResult = Await pending.AgentTask
            pending.IsCompleted = True

            If agentResult IsNot Nothing AndAlso agentResult.IsCancelled Then
                SafeCompleteRun(runId, "cancelled", "已取消", ExceptionClassifier.CodeCancelled, DateTime.Now)
                RaisePhase(runId, "cancelled", "已取消")
                RemovePendingRun(runId)
                _currentRunId = ""
                Return New HarnessRunResult With {
                    .RunId = runId,
                    .Status = HarnessRunStatus.Cancelled,
                    .UserMessage = If(agentResult.Message, "已取消"),
                    .DebugMessage = "CANCELLED_AT_SAFE_POINT",
                    .AgentSessionId = If(agentResult.SessionId, ""),
                    .StartedAt = pending.StartedAt,
                    .FinishedAt = DateTime.Now
                }
            End If

            Dim succeeded = agentResult IsNot Nothing AndAlso agentResult.Success
            Dim status = If(succeeded, HarnessRunStatus.Succeeded, HarnessRunStatus.Failed)
            Dim message = If(agentResult?.Message, If(succeeded, "执行完成", "执行失败"))
            SafeCompleteRun(runId, If(succeeded, "succeeded", "failed"), message, If(succeeded, "", ExceptionClassifier.CodeUnknown), DateTime.Now)
            RaisePhase(runId, If(succeeded, "completed", "failed"), message)
            RemovePendingRun(runId)
            _currentRunId = ""
            Return New HarnessRunResult With {
                .RunId = runId,
                .Status = status,
                .UserMessage = message,
                .DebugMessage = message,
                .AgentSessionId = If(agentResult?.SessionId, ""),
                .StartedAt = pending.StartedAt,
                .FinishedAt = DateTime.Now
            }
        End Function

        Public Async Function CancelAsync(runId As String,
                                          cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.CancelAsync
            Dim pending = GetPendingRun(runId)
            If pending Is Nothing Then Return MissingRunResult(runId)
            If pending.ApprovalCallback IsNot Nothing Then
                Return Await ApproveAsync(runId, False, cancellationToken)
            End If

            Return New HarnessRunResult With {
                .RunId = runId,
                .Status = HarnessRunStatus.Failed,
                .UserMessage = "当前步骤正在执行，尚不能安全中断宿主 COM 操作",
                .DebugMessage = "CANCEL_NOT_AT_SAFE_POINT",
                .StartedAt = pending.StartedAt,
                .FinishedAt = DateTime.Now
            }
        End Function

        Public Function ResumeAsync(runId As String,
                                    cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.ResumeAsync
            Return ApproveAsync(runId, True, cancellationToken)
        End Function

        Private Sub HandleKernelStatusChanged(status As String)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            RaisePhase(_currentRunId, "status", If(status, ""))
        End Sub

        Private Sub HandleKernelStepCompleted(stepIndex As Integer, success As Boolean, message As String)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            RaiseEvent StepChanged(Me, New HarnessStepChangedEventArgs With {
                .RunId = _currentRunId,
                .StepIndex = stepIndex,
                .Status = If(success, "completed", "failed"),
                .Message = If(message, "")
            })
        End Sub

        Private Sub HandleKernelPlanGenerated(plan As ExecutionPlan)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            Dim count = If(plan?.Steps?.Count, 0)
            RaisePhase(_currentRunId, "planned", $"生成执行计划：{count} 步")
        End Sub

        Private Sub HandleKernelExecutionExplained(explanation As ExecutionExplanation)
            If String.IsNullOrWhiteSpace(_currentRunId) OrElse explanation Is Nothing Then Return
            Try
                _runTraceStore.AppendStep(_currentRunId,
                                          explanation.StepIndex,
                                          explanation.ToolId,
                                          If(explanation.Success, "succeeded", "failed"),
                                          If(explanation.Message, explanation.ExplanationText),
                                          If(explanation.FailureReason, ""),
                                          explanation,
                                          explanation.StartedAt,
                                          explanation.FinishedAt)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Append run trace step failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub HandleKernelRequestApproval(message As String, callback As Action(Of Boolean))
            Dim pending = GetPendingRun(_currentRunId)
            If pending Is Nothing Then
                callback(False)
                Return
            End If
            pending.ApprovalMessage = If(message, "该操作需要用户确认")
            pending.ApprovalCallback = callback
            SafeAppendApprovalStep(pending.RunId,
                                   -1,
                                   "approval.request",
                                   "awaiting_approval",
                                   pending.ApprovalMessage,
                                   ExceptionClassifier.CodeSafetyNeedsApproval,
                                   New With {.message = pending.ApprovalMessage},
                                   DateTime.Now)
            pending.ApprovalSignal.TrySetResult(True)
        End Sub

        Private Sub SafeAppendApprovalStep(runId As String,
                                           seq As Integer,
                                           toolId As String,
                                           status As String,
                                           message As String,
                                           errorCode As String,
                                           observation As Object,
                                           occurredAt As DateTime)
            Try
                _runTraceStore.AppendStep(runId,
                                          seq,
                                          toolId,
                                          status,
                                          message,
                                          errorCode,
                                          observation,
                                          occurredAt,
                                          occurredAt)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Append approval trace failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Function GetPendingRun(runId As String) As PendingHarnessRun
            If String.IsNullOrWhiteSpace(runId) Then Return Nothing
            SyncLock _runLock
                Dim pending As PendingHarnessRun = Nothing
                If _pendingRuns.TryGetValue(runId, pending) Then Return pending
            End SyncLock
            Return Nothing
        End Function

        Private Sub RemovePendingRun(runId As String)
            SyncLock _runLock
                _pendingRuns.Remove(runId)
            End SyncLock
        End Sub

        Private Shared Function MissingRunResult(runId As String) As HarnessRunResult
            Return New HarnessRunResult With {
                .RunId = If(runId, ""),
                .Status = HarnessRunStatus.Failed,
                .UserMessage = "未找到等待审批的任务",
                .DebugMessage = "RUN_NOT_AWAITING_APPROVAL",
                .FinishedAt = DateTime.Now
            }
        End Function

        Private Class PendingHarnessRun
            Public Property RunId As String
            Public Property Turn As UserTurn
            Public Property StartedAt As DateTime
            Public Property AgentTask As Task(Of AgentResult)
            Public Property ApprovalSignal As TaskCompletionSource(Of Boolean)
            Public Property ApprovalCallback As Action(Of Boolean)
            Public Property ApprovalMessage As String = ""
            Public Property IsCompleted As Boolean
        End Class

        Private Sub SafeStartRun(runId As String, turn As UserTurn, startedAt As DateTime)
            Try
                _runTraceStore.StartRun(runId, turn, startedAt)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Start run trace failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub SafeCompleteRun(runId As String, status As String, finalMessage As String, errorCode As String, finishedAt As DateTime)
            Try
                _runTraceStore.CompleteRun(runId, status, finalMessage, errorCode, finishedAt)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Complete run trace failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub SafeSetRunStatus(runId As String, status As String, message As String, errorCode As String)
            Try
                _runTraceStore.SetRunStatus(runId, status, message, errorCode)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Set run trace status failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub RaisePhase(runId As String, phase As String, message As String)
            RaiseEvent PhaseChanged(Me, New HarnessPhaseChangedEventArgs With {
                .RunId = If(runId, ""),
                .Phase = If(phase, ""),
                .Message = If(message, "")
            })
        End Sub
    End Class

End Namespace
