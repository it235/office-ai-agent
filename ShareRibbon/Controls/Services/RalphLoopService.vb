Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' Ralph Loop 服务：封装 Loop 规划、执行、取消及完成回调逻辑
''' </summary>
Public Class RalphLoopService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _sendRequest As Func(Of String, String, Boolean, String, Task)
    Private ReadOnly _escapeJs As Func(Of String, String)
    Private ReadOnly _getApplicationType As Func(Of String)
    Private ReadOnly _ralphLoopController As New RalphLoopController()

    ' 状态字段
    Private _isRalphLoopPlanning As Boolean = False
    Private _currentRalphLoopStep As RalphLoopStep = Nothing

    ''' <summary>暴露底层控制器供外部直接调用（如 HandleConfirmIntent）</summary>
    Public ReadOnly Property Controller As RalphLoopController
        Get
            Return _ralphLoopController
        End Get
    End Property

    ''' <summary>规划模式标记，供外部（如 HandleConfirmIntent）设置</summary>
    Public Property IsPlanning As Boolean
        Get
            Return _isRalphLoopPlanning
        End Get
        Set(value As Boolean)
            _isRalphLoopPlanning = value
        End Set
    End Property

    Public Sub New(
        executeScript As Func(Of String, Task),
        sendRequest As Func(Of String, String, Boolean, String, Task),
        escapeJs As Func(Of String, String),
        getApplicationType As Func(Of String))

        _executeScript = executeScript
        _sendRequest = sendRequest
        _escapeJs = escapeJs
        _getApplicationType = getApplicationType
    End Sub

    ''' <summary>
    ''' 启动Ralph Loop - 用户输入目标后调用
    ''' </summary>
    Public Async Function StartRalphLoop(userGoal As String) As Task
        Try
            Debug.WriteLine($"[RalphLoop] 启动循环，目标: {userGoal}")

            Dim appType = _getApplicationType()
            Dim loopSession = Await _ralphLoopController.StartNewLoop(userGoal, appType)

            Dim loopDataJson = $"{{""goal"":""{_escapeJs(userGoal)}"",""steps"":[],""status"":""planning""}}"
            Await _executeScript($"showLoopPlanCard({loopDataJson})")

            GlobalStatusStrip.ShowInfo("正在规划任务...")

            Dim planningPrompt = _ralphLoopController.GetPlanningPrompt(userGoal)
            _isRalphLoopPlanning = True
            Await _sendRequest(planningPrompt, "", False, "")

        Catch ex As Exception
            Debug.WriteLine($"[RalphLoop] 启动失败: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"启动循环失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 处理前端 startLoop 消息
    ''' </summary>
    Public Sub HandleStartLoop(jsonDoc As JObject)
        Try
            Dim goal = jsonDoc("goal")?.ToString()
            If Not String.IsNullOrEmpty(goal) Then
                StartRalphLoop(goal)
            End If
        Catch ex As Exception
            Debug.WriteLine($"HandleStartLoop 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理继续执行循环
    ''' </summary>
    Public Async Function HandleContinueLoop() As Task
        Try
            Debug.WriteLine("[RalphLoop] 用户点击继续执行")

            Dim nextStep = _ralphLoopController.ExecuteNextStep()
            If nextStep Is Nothing Then
                Debug.WriteLine("[RalphLoop] 没有更多步骤")
                Await _executeScript("updateLoopStatus('completed')")
                GlobalStatusStrip.ShowInfo("所有步骤已完成")
                Return
            End If

            Await _executeScript($"updateLoopStep({nextStep.StepNumber - 1}, 'running')")
            Await _executeScript("updateLoopStatus('running')")
            GlobalStatusStrip.ShowInfo($"正在执行步骤 {nextStep.StepNumber}: {nextStep.Description}")

            _currentRalphLoopStep = nextStep
            Await _sendRequest(nextStep.Description, "", True, "")

        Catch ex As Exception
            Debug.WriteLine($"HandleContinueLoop 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"执行步骤失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 处理取消循环
    ''' </summary>
    Public Sub HandleCancelLoop()
        Try
            Debug.WriteLine("[RalphLoop] 用户取消循环")

            _ralphLoopController.ClearAndEndLoop()
            _isRalphLoopPlanning = False
            _currentRalphLoopStep = Nothing

            _executeScript("hideLoopPlanCard()")
            GlobalStatusStrip.ShowInfo("已取消循环任务")

        Catch ex As Exception
            Debug.WriteLine($"HandleCancelLoop 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 在流完成后检查是否需要处理 Ralph Loop，由 BaseChatControl 流完成时调用
    ''' </summary>
    Public Sub CheckRalphLoopCompletion(responseContent As String)
        Try
            If _isRalphLoopPlanning Then
                _isRalphLoopPlanning = False

                If _ralphLoopController.ParsePlanningResult(responseContent) Then
                    Dim loopSession = _ralphLoopController.GetActiveLoop()
                    If loopSession IsNot Nothing Then
                        Dim stepsJson = BuildStepsJson(loopSession.Steps)
                        Dim loopDataJson = $"{{""goal"":""{_escapeJs(loopSession.OriginalGoal)}"",""steps"":{stepsJson},""status"":""ready""}}"
                        _executeScript($"showLoopPlanCard({loopDataJson})")
                        GlobalStatusStrip.ShowInfo("规划完成，点击[继续执行]开始")
                    End If
                Else
                    GlobalStatusStrip.ShowWarning("规划结果解析失败")
                    _executeScript("hideLoopPlanCard()")
                End If
                Return
            End If

            If _currentRalphLoopStep IsNot Nothing Then
                Dim stepNum = _currentRalphLoopStep.StepNumber
                _ralphLoopController.CompleteCurrentStep(responseContent, True)
                _currentRalphLoopStep = Nothing

                _executeScript($"updateLoopStep({stepNum - 1}, 'completed')")

                Dim loopSession = _ralphLoopController.GetActiveLoop()
                If loopSession IsNot Nothing Then
                    If loopSession.Status = RalphLoopStatus.Paused Then
                        _executeScript("updateLoopStatus('paused')")
                        GlobalStatusStrip.ShowInfo($"步骤 {stepNum} 完成，点击继续执行下一步")
                    ElseIf loopSession.Status = RalphLoopStatus.Completed Then
                        _executeScript("updateLoopStatus('completed')")
                        GlobalStatusStrip.ShowInfo("所有步骤已完成！")
                    End If
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"CheckRalphLoopCompletion 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理重新规划循环
    ''' </summary>
    Public Sub HandleReplanLoop(jsonDoc As JObject)
        Try
            Dim feedback = jsonDoc("feedback")?.ToString()
            If String.IsNullOrEmpty(feedback) Then Return

            _executeScript("addThinkingMessage('正在根据您的意见重新规划...')")

            Task.Run(Async Function()
                Await ReplanAsync(feedback)
            End Function)
        Catch ex As Exception
            Debug.WriteLine($"HandleReplanLoop 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 根据用户反馈重新规划
    ''' </summary>
    Private Async Function ReplanAsync(feedback As String) As Task
        Try
            Dim loopSession = _ralphLoopController.GetActiveLoop()
            If loopSession Is Nothing Then Return

            Dim originalGoal = loopSession.OriginalGoal
            Dim appType = loopSession.ApplicationType

            ' 重置当前会话步骤
            loopSession.Steps.Clear()
            loopSession.CurrentStep = 0
            loopSession.Status = RalphLoopStatus.Planning

            Dim loopDataJson = $"{{""goal"":""{_escapeJs(originalGoal)}"",""steps"":[],""status"":""planning""}}"
            Await _executeScript($"showLoopPlanCard({loopDataJson})")
            GlobalStatusStrip.ShowInfo("正在重新规划任务...")

            Dim refinedGoal = originalGoal & vbCrLf & "[用户反馈] " & feedback
            Dim planningPrompt = _ralphLoopController.GetPlanningPrompt(refinedGoal)
            _isRalphLoopPlanning = True
            Await _sendRequest(planningPrompt, "", False, "")
        Catch ex As Exception
            Debug.WriteLine($"ReplanAsync 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"重新规划失败: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 构建步骤 JSON 字符串
    ''' </summary>
    Public Function BuildStepsJson(steps As List(Of RalphLoopStep)) As String
        Dim sb As New StringBuilder()
        sb.Append("[")
        For i = 0 To steps.Count - 1
            If i > 0 Then sb.Append(",")
            Dim s = steps(i)
            Dim statusStr = s.Status.ToString().ToLower()
            sb.Append($"{{""description"":""{_escapeJs(s.Description)}"",""status"":""{statusStr}""}}")
        Next
        sb.Append("]")
        Return sb.ToString()
    End Function

End Class
