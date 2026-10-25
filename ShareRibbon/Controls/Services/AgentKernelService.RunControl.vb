' ShareRibbon\Controls\Services\AgentKernelService.RunControl.vb
' Run control for the agent service: stop, approval routing, and refusing a turn that never started.

Imports System.Threading.Tasks

Public Partial Class AgentKernelService

    ''' <summary>
    ''' 终止当前 Agent 任务：触发协作式取消，等待循环在安全点收敛，而不是伪造"已终止"。
    ''' </summary>
    Public Sub AbortAgent()
        Try
            Dim signaledRunning = False
            If _runCts IsNot Nothing AndAlso Not _runCts.IsCancellationRequested Then
                _runCts.Cancel()
                signaledRunning = True
            End If

            ' 若正卡在等待审批，取消令牌无法立刻唤醒，需主动放弃审批以释放挂起回调。
            Dim hadPendingApproval = False
            If Not String.IsNullOrWhiteSpace(CurrentHarnessRunId) Then
                hadPendingApproval = True
                CancelHarnessRunAsync(CurrentHarnessRunId)
            End If

            If signaledRunning OrElse hadPendingApproval Then
                OnKernelStatusChanged("正在停止，等待当前步骤安全结束…")
                Return
            End If

            ' 没有正在运行的任务：直接复位 UI。
            Dim sessionId = CurrentAgentSessionId
            AgentThinkingUuid = Nothing
            AgentOriginalUserRequest = Nothing
            AgentFullUserMessage = Nothing
            AgentEntryActionId = Nothing
            CurrentAgentSessionId = Nothing
            CurrentHarnessRunId = Nothing

            _executeScript($"completeAgent({UtilsService.ToJsStringLiteral(sessionId)}, false, {UtilsService.ToJsStringLiteral("已终止")})")
            GlobalStatusStrip.ShowInfo("已终止Agent")
        Catch ex As Exception
            AppLogger.Error("AgentKernelService", "AbortAgent exception", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 面板对单个步骤的纠偏（E10-01）。这一层只回答"收下了没有"，不预测何时生效：
    ''' 落地由循环在步骤边界做，宿主改步骤状态等于第二套执行真相。返回空串表示已排队。
    ''' </summary>
    Public Function RequestStepAction(sessionId As String, stepIndex As Integer, action As String) As String
        If _agentKernel Is Nothing Then Return "Agent 还没有初始化，请稍后再试"
        If stepIndex < 0 Then Return "缺少步骤序号，无法定位这一步"
        If _agentKernel.QueueStepAction(sessionId, stepIndex, action) Then
            AppLogger.Info("AgentKernelService", $"Step action queued step={stepIndex + 1} action={action}")
            Return ""
        End If
        ' 按了没反应比没有这个按钮更糟：必须把"为什么这次不算"送回界面。
        AppLogger.Info("AgentKernelService", $"Step action refused step={stepIndex + 1} action={action}")
        Return "这一步已经不在当前这一轮里了（本轮已结束，或计划已被换成新版本）"
    End Function

    ''' <summary>
    ''' 用户允许一次当前高风险操作
    ''' </summary>
    Public Sub Approve()
        ResolveHarnessApprovalAsync(Agent.ApprovalDecision.Once)
    End Sub

    ''' <summary>
    ''' 用户允许本轮会话内该工具的所有调用
    ''' </summary>
    Public Sub ApproveForSession()
        ResolveHarnessApprovalAsync(Agent.ApprovalDecision.Session)
    End Sub

    ''' <summary>
    ''' 用户拒绝当前操作（只跳过当前步骤）
    ''' </summary>
    Public Sub Reject()
        ResolveHarnessApprovalAsync(Agent.ApprovalDecision.Deny)
    End Sub

    Private Async Sub ResolveHarnessApprovalAsync(decision As Agent.ApprovalDecision)
        If _officeHarness Is Nothing OrElse String.IsNullOrWhiteSpace(CurrentHarnessRunId) Then Return
        Try
            Dim result = Await _officeHarness.DecideAsync(CurrentHarnessRunId, decision, Threading.CancellationToken.None)
            If result IsNot Nothing AndAlso result.Status <> Agent.Harness.HarnessRunStatus.AwaitingApproval Then
                CurrentHarnessRunId = Nothing
            End If
        Catch ex As Exception
            AppLogger.Error("AgentKernelService", "Resolve harness approval failed", ex)
        End Try
    End Sub

    Private Async Sub CancelHarnessRunAsync(runId As String)
        Try
            Dim result = Await _officeHarness.CancelAsync(runId, Threading.CancellationToken.None)
            ' 运行中无法安全抢占宿主 COM 属预期：真正停止由 _runCts 在步骤边界收敛；此处仅如实记录。
            If result Is Nothing Then
                AppLogger.Warn("AgentKernelService", "Cancel harness run returned null")
            ElseIf result.Status = Agent.Harness.HarnessRunStatus.Cancelled OrElse
                   result.Status = Agent.Harness.HarnessRunStatus.Succeeded OrElse
                   result.Status = Agent.Harness.HarnessRunStatus.Failed Then
                AppLogger.Info("AgentKernelService", $"Cancel harness run: status={result.Status} debug={result.DebugMessage}")
            End If
        Catch ex As Exception
            AppLogger.Warn("AgentKernelService", $"Cancel harness run threw: {AppLogger.Redact(ex.Message)}")
        End Try
    End Sub

    ''' <summary>
    ''' 本轮从未启动（harness 无返回，或上一轮仍在途而被拒绝）：把在途 run 的会话与取消令牌原样交还，
    ''' 只提示不接管；没有在途 run 时才按普通失败收敛界面。
    ''' </summary>
    Private Function RefuseUnstartedRun(sessionId As String,
                                        thinkingUuid As String,
                                        originalRequest As String,
                                        fullMessage As String,
                                        entryActionId As String,
                                        carriedCts As Threading.CancellationTokenSource,
                                        rejectedCts As Threading.CancellationTokenSource,
                                        message As String) As Boolean
        CurrentAgentSessionId = sessionId
        AgentThinkingUuid = thinkingUuid
        AgentOriginalUserRequest = originalRequest
        AgentFullUserMessage = fullMessage
        AgentEntryActionId = entryActionId
        _runCts = carriedCts
        rejectedCts.Dispose()
        If String.IsNullOrWhiteSpace(sessionId) Then
            FinalizeAgentUi(False, message)
        Else
            GlobalStatusStrip.ShowWarning(message)
        End If
        Return False
    End Function

End Class
