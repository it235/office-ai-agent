' ShareRibbon\Agent\Harness\OfficeHarness.Approvals.vb
' Approval handshake for a harness run: request, decision, timeout exit and abandonment.

Imports System.Threading.Tasks

Namespace Agent.Harness

    Public Partial Class OfficeHarness

        #Region "Approval Handshake"

        Private Sub HandleKernelRequestApproval(message As String, callback As Action(Of Boolean))
            Dim pending = GetPendingRun(_currentRunId)
            If pending Is Nothing Then
                callback(False)
                Return
            End If
            SyncLock _runLock
                pending.ApprovalMessage = If(message, "该操作需要用户确认")
                pending.ApprovalCallback = callback
                pending.AwaitingApproval = True
            End SyncLock
            SafeAppendApprovalStep(pending.RunId,
                                   -1,
                                   "approval.request",
                                   "awaiting_approval",
                                   pending.ApprovalMessage,
                                   ExceptionClassifier.CodeSafetyNeedsApproval,
                                   New With {.message = pending.ApprovalMessage},
                                   DateTime.Now)
            ' 每个待决审批自带超时哨兵：第二次、第三次请求审批同样不会把会话无限挂住。
            ArmApprovalTimeout(pending)
            pending.ApprovalSignal.TrySetResult(True)
        End Sub

        ''' <summary>
        ''' 到点仍无人裁决时按拒绝收敛。拒绝是安全默认：未经确认的写操作绝不因为等待太久而擅自执行。
        ''' </summary>
        Private Sub ArmApprovalTimeout(pending As PendingHarnessRun)
            Dim runId = pending.RunId
            Task.Run(Async Function()
                         Try
                             Await Task.Delay(ApprovalTimeout)
                             If Not pending.AwaitingApproval OrElse pending.IsCompleted Then Return
                             Dim settled = Await SettleApprovalAsync(pending,
                                                                     False,
                                                                     "timeout",
                                                                     "审批等待超时，该操作已放弃",
                                                                     ExceptionClassifier.CodeApprovalTimeout)
                             If Not settled.HasValue Then Return
                             ReleaseAbandonedRun(pending, "审批等待超时，该操作已放弃", ExceptionClassifier.CodeApprovalTimeout)
                         Catch ex As Exception
                             AppLogger.Warn("OfficeHarness", $"Approval timeout watcher failed: {AppLogger.Redact(ex.Message)}")
                         End Try
                     End Function)
        End Sub

        ''' <summary>
        ''' 取走当前待决审批并让 Loop 在安全点收敛。
        ''' Nothing 表示审批已被别的路径接管；True/False 表示本次接管后 Loop 是否在宽限期内收敛。
        ''' </summary>
        Private Async Function SettleApprovalAsync(pending As PendingHarnessRun,
                                                   approved As Boolean,
                                                   decision As String,
                                                   message As String,
                                                   errorCode As String) As Task(Of Boolean?)
            Dim callback As Action(Of Boolean) = Nothing
            SyncLock _runLock
                If pending.ApprovalCallback Is Nothing Then Return Nothing
                callback = pending.ApprovalCallback
                pending.ApprovalCallback = Nothing
                pending.AwaitingApproval = False
            End SyncLock

            SafeAppendApprovalStep(pending.RunId,
                                   -2,
                                   "approval.decision",
                                   decision,
                                   message,
                                   errorCode,
                                   New With {.approved = approved, .decision = decision},
                                   DateTime.Now)
            callback(approved)

            Dim agentTask = pending.AgentTask
            If agentTask Is Nothing Then Return True
            Dim finished = Await Task.WhenAny(agentTask, Task.Delay(ApprovalSettleGrace))
            If finished Is agentTask Then Return True

            ' 拒绝后 Loop 直接走不可恢复失败出口，正常都在毫秒级收敛；没收敛时由续接负责释放。
            AppLogger.Warn("OfficeHarness",
                           $"Run {pending.RunId} did not converge within grace after approval {decision}.")
            Return False
        End Function

        ''' <summary>
        ''' 审批被放弃（超时，或被新一轮对话取代）后的收尾。Loop 已结束就立刻落终态摘除；
        ''' 还没结束则续接摘除，期间宁可拒绝新任务，也不让两个执行体同时抢占同一个 AgentKernel 写文档。
        ''' </summary>
        Private Sub ReleaseAbandonedRun(pending As PendingHarnessRun, message As String, errorCode As String)
            If pending.AgentTask Is Nothing OrElse pending.AgentTask.IsCompleted Then
                FinishAbandonedRun(pending, message, errorCode)
                Return
            End If
            pending.AgentTask.ContinueWith(Sub(t) FinishAbandonedRun(pending, message, errorCode))
        End Sub

        Private Sub FinishAbandonedRun(pending As PendingHarnessRun, message As String, errorCode As String)
            SafeCompleteRun(pending.RunId, "cancelled", message, errorCode, DateTime.Now)
            pending.IsCompleted = True
            RemovePendingRun(pending.RunId)
            AppLogger.Warn("OfficeHarness", $"Abandoned awaiting run {pending.RunId}: {AppLogger.Redact(message)}")
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

        Private Shared Function MissingRunResult(runId As String) As HarnessRunResult
            Return New HarnessRunResult With {
                .RunId = If(runId, ""),
                .Status = HarnessRunStatus.Failed,
                .UserMessage = "未找到等待审批的任务",
                .DebugMessage = "RUN_NOT_AWAITING_APPROVAL",
                .FinishedAt = DateTime.Now
            }
        End Function


        #End Region

    End Class

End Namespace
