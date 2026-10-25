Imports System.Threading
Imports System.Threading.Tasks

Namespace Agent.Harness

    Public Partial Class OfficeHarness
        Implements IOfficeHarness

        Private ReadOnly _kernel As AgentKernel
        Private ReadOnly _runTraceStore As IRunTraceStore
        Private _currentRunId As String = ""
        Private ReadOnly _runLock As New Object()
        Private ReadOnly _pendingRuns As New Dictionary(Of String, PendingHarnessRun)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' 审批等待的最长时长。超时后按"拒绝"收敛，绝不因为没人点按钮就把会话永久锁在
        ''' 一个未决的高风险操作上；默认值宽松，留出人工确认的时间。
        ''' </summary>
        Public Property ApprovalTimeout As TimeSpan = TimeSpan.FromMinutes(10)

        ''' <summary>
        ''' 撤销待决审批后，等待 Loop 在安全点收敛的时间上限。
        ''' </summary>
        Private Shared ReadOnly ApprovalSettleGrace As TimeSpan = TimeSpan.FromSeconds(5)

        Public Event PhaseChanged As EventHandler(Of HarnessPhaseChangedEventArgs) Implements IOfficeHarness.PhaseChanged
        Public Event StepChanged As EventHandler(Of HarnessStepChangedEventArgs) Implements IOfficeHarness.StepChanged
        Public Event ContextReady As EventHandler(Of HarnessContextEventArgs) Implements IOfficeHarness.ContextReady

        Public Sub New(kernel As AgentKernel, Optional runTraceStore As IRunTraceStore = Nothing)
            If kernel Is Nothing Then Throw New ArgumentNullException(NameOf(kernel))
            _kernel = kernel
            _runTraceStore = If(runTraceStore, New NoopRunTraceStore())
            AddHandler _kernel.OnStatusChanged, AddressOf HandleKernelStatusChanged
            AddHandler _kernel.OnStepStatusChanged, AddressOf HandleKernelStepStatusChanged
            AddHandler _kernel.OnPlanGenerated, AddressOf HandleKernelPlanGenerated
            AddHandler _kernel.OnExecutionExplained, AddressOf HandleKernelExecutionExplained
            AddHandler _kernel.OnRequestApproval, AddressOf HandleKernelRequestApproval
        End Sub

        Public Async Function RunAsync(turn As UserTurn,
                                       cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.RunAsync
            Dim runId = Guid.NewGuid().ToString()
            Dim startedAt = DateTime.Now
            Dim keepPending As Boolean = False

            Dim superseded As PendingHarnessRun = Nothing
            SyncLock _runLock
                Dim active = _pendingRuns.Values.FirstOrDefault(Function(item) Not item.IsCompleted)
                If active IsNot Nothing Then
                    ' 只有"等待用户裁决"的 run 能被新一轮对话放弃；正在操作宿主的必须拒绝，
                    ' 否则两个任务会并发抢占同一个单例 AgentKernel。
                    If Not active.AwaitingApproval Then
                        Return BlockedResult(runId, startedAt, "当前会话已有正在执行的任务，请先停止或等待它结束")
                    End If
                    superseded = active
                End If
            End SyncLock

            If superseded IsNot Nothing Then
                ' 用户没有点批准/拒绝而是直接说了新需求：视为放弃上一轮未确认的操作。
                ' 未确认的写操作绝不擅自执行，按拒绝收敛后才能腾出 kernel 跑新一轮。
                Dim settled = Await SettleApprovalAsync(superseded,
                                                        ApprovalDecision.Abort,
                                                        "superseded",
                                                        "用户发起了新一轮对话，上一轮未确认的操作已放弃",
                                                        ExceptionClassifier.CodeCancelled)
                If settled.HasValue Then
                    ReleaseAbandonedRun(superseded, "用户发起了新一轮对话，上一轮未确认的操作已放弃",
                                        ExceptionClassifier.CodeCancelled)
                End If
                If HasActiveRun() Then
                    Return BlockedResult(runId, startedAt, "上一轮任务正在收尾，请稍后再试")
                End If
            End If

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
                                                         turn.EntryActionId,
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
                SafePersistChangeLog(runId, turn, agentResult)

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

        Public Function ApproveAsync(runId As String,
                                     approved As Boolean,
                                     cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.ApproveAsync
            Return DecideAsync(runId, If(approved, ApprovalDecision.Once, ApprovalDecision.Deny), cancellationToken)
        End Function

        Public Async Function DecideAsync(runId As String,
                                          decision As ApprovalDecision,
                                          cancellationToken As CancellationToken) As Task(Of HarnessRunResult) Implements IOfficeHarness.DecideAsync
            Dim pending = GetPendingRun(runId)
            If pending Is Nothing OrElse pending.AgentTask Is Nothing Then
                Return MissingRunResult(runId)
            End If

            cancellationToken.ThrowIfCancellationRequested()
            Dim allowed = decision = ApprovalDecision.Once OrElse decision = ApprovalDecision.Session
            Dim decisionName As String
            If allowed Then
                decisionName = If(decision = ApprovalDecision.Session, "approved_for_session", "approved_once")
            Else
                decisionName = If(decision = ApprovalDecision.Abort, "aborted", "rejected")
            End If
            Dim decisionMessage As String
            If decision = ApprovalDecision.Session Then
                decisionMessage = "用户已允许本会话使用该工具"
            ElseIf decision = ApprovalDecision.Once Then
                decisionMessage = "用户已批准本次高风险操作"
            ElseIf decision = ApprovalDecision.Abort Then
                decisionMessage = "审批已放弃，正在停止本轮执行"
            Else
                decisionMessage = "用户已拒绝当前高风险操作"
            End If
            Dim settled = Await SettleApprovalAsync(pending,
                                                    decision,
                                                    decisionName,
                                                    decisionMessage,
                                                    If(allowed, "", ExceptionClassifier.CodeSafetyBlocked))
            If Not settled.HasValue Then Return MissingRunResult(runId)
            SafeSetRunStatus(runId, If(decision = ApprovalDecision.Abort, "cancelling", "running"),
                             If(allowed,
                                "审批通过，继续执行",
                                If(decision = ApprovalDecision.Abort, "审批已放弃，正在停止本轮执行", "审批拒绝，正在跳过当前步骤")),
                             "")

            Dim agentResult = Await pending.AgentTask
            pending.IsCompleted = True
            SafePersistChangeLog(runId, pending.Turn, agentResult)

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
                Return Await DecideAsync(runId, ApprovalDecision.Abort, cancellationToken)
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
            Return DecideAsync(runId, ApprovalDecision.Once, cancellationToken)
        End Function

        Private Sub HandleKernelStatusChanged(status As String)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            RaisePhase(_currentRunId, "status", If(status, ""))
        End Sub

        Private Sub HandleKernelStepStatusChanged(stepIndex As Integer, status As String, message As String)
            If String.IsNullOrWhiteSpace(_currentRunId) Then Return
            RaiseEvent StepChanged(Me, New HarnessStepChangedEventArgs With {
                .RunId = _currentRunId,
                .StepIndex = stepIndex,
                .Status = If(String.IsNullOrWhiteSpace(status), "failed", status),
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

        Private Shared Function BlockedResult(runId As String, startedAt As DateTime, message As String) As HarnessRunResult
            Return New HarnessRunResult With {
                .RunId = runId,
                .Status = HarnessRunStatus.Failed,
                .UserMessage = message,
                .DebugMessage = "RUN_ALREADY_ACTIVE",
                .StartedAt = startedAt,
                .FinishedAt = DateTime.Now
            }
        End Function

        Private Function HasActiveRun() As Boolean
            SyncLock _runLock
                Return _pendingRuns.Values.Any(Function(item) Not item.IsCompleted)
            End SyncLock
        End Function

        Private Class PendingHarnessRun
            Public Property RunId As String
            Public Property Turn As UserTurn
            Public Property StartedAt As DateTime
            Public Property AgentTask As Task(Of AgentResult)
            Public Property ApprovalSignal As TaskCompletionSource(Of Boolean)
            Public Property ApprovalCallback As Action(Of ApprovalDecision)
            Public Property ApprovalMessage As String = ""
            Public Property ApprovalToolId As String = ""
            Public Property ApprovalRiskLevel As String = ""
            Public Property IsCompleted As Boolean
            ''' <summary>
            ''' 仅表示"当前正卡在等用户裁决"，是新一轮对话能否放弃它的选择性依据。
            ''' </summary>
            Public Property AwaitingApproval As Boolean
        End Class

        ''' <summary>
        ''' 一轮结束后持久化写操作变更日志；失败或取消也要保存已落地的部分。
        ''' </summary>
        Private Sub SafePersistChangeLog(runId As String, turn As UserTurn, agentResult As AgentResult)
            Try
                If agentResult Is Nothing OrElse agentResult.ChangeLog Is Nothing Then Return
                Dim sessionId = If(String.IsNullOrWhiteSpace(agentResult.SessionId), If(turn?.SessionId, ""), agentResult.SessionId)
                AgentChangeLogRepository.SaveRun(runId, sessionId, If(turn?.AppType, ""), agentResult.ChangeLog)
            Catch ex As Exception
                AppLogger.Warn("OfficeHarness", $"Persist change log failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

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
