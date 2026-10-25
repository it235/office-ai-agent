Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' Agent 会话 - 统一的数据模型
    ''' </summary>
    Public Class AgentSession
        Public Property Id As String = Guid.NewGuid().ToString()
        Public Property UserRequest As String
        Public Property AppType As String
        Public Property CurrentContent As String
        Public Property Skill As AgentSkill
        Public Property SelectedSkill As SkillFileDefinition
        Public Property Spec As AgentTaskSpec

        ''' <summary>
        ''' 本轮来自哪个入口动作（QuickActionCatalog 里的 id）；自由文本输入为空串。
        ''' 这是任务包断言唯一的绑定依据：没有入口身份就不核对，绝不靠关键词猜一个包。
        ''' </summary>
        Public Property EntryActionId As String = ""

        ' ReAct 循环记录
        Public Property Iterations As New List(Of ReActIteration)
        Public Property Status As AgentStatus = AgentStatus.Idle
        Public Property StartTime As DateTime = DateTime.Now
        Public Property CurrentIteration As Integer = 0
        Public Property MaxIterations As Integer = 15

        ' 执行计划
        Public Property Plan As ExecutionPlan

        ' 本轮写操作变更日志（目标、前后摘要、是否成功）
        Public Property ChangeLog As New AgentRunChangeLog()

        ' 本轮真实落盘写入的台账，供 replan 判定哪些步骤已经做完（E1-06）
        Public Property Writes As New AgentWriteLedger()

        ' 本轮用户明确选择“本会话允许”的工具；只在这一轮 Agent 会话内生效
        Public Property SessionApprovedTools As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ' 最终结果
        Public Property Result As AgentResult

        Public Sub New(userRequest As String, appType As String, currentContent As String)
            Me.UserRequest = userRequest
            Me.AppType = appType
            Me.CurrentContent = currentContent
        End Sub
    End Class

    ''' <summary>
    ''' ReAct 迭代记录
    ''' </summary>
    Public Class ReActIteration
        Public Property Index As Integer
        Public Property Thought As String
        Public Property Action As ToolCall
        Public Property Observation As String
        Public Property Explanation As ExecutionExplanation
        Public Property Timestamp As DateTime = DateTime.Now
    End Class

    ''' <summary>
    ''' 执行计划
    ''' </summary>
    Public Class ExecutionPlan
        Public Property Understanding As String
        Public Property Steps As New List(Of PlanStep)
        Public Property Summary As String
        Public Property Complexity As String = "medium"
        Public Property CapabilityGap As String = ""
        ''' <summary>发给 UI 前由 LoopEngine 补齐的计划事实；模型不参与生成，见 E1-08。</summary>
        Public Property Trust As PlanTrust
    End Class

    ''' <summary>
    ''' 计划卡的信任信息：这一版计划要动多大范围、几步几个工具、其中哪些需要确认。
    ''' 重规划会产生新版本，卡片按版本号原位重绘，时间线才不会停留在旧计划上。
    ''' </summary>
    Public Class PlanTrust
        Public Property PlanVersion As Integer = 1
        Public Property ScopeText As String = ""
        Public Property StepCount As Integer
        Public Property ToolCount As Integer
        Public Property RiskLevel As String = "safe"
        Public Property RiskyTools As New List(Of String)()
    End Class

    ''' <summary>
    ''' 计划步骤
    ''' </summary>
    Public Class PlanStep
        Public Property StepNumber As Integer
        Public Property Description As String
        Public Property Code As String
        Public Property Language As String = "json"
        Public Property Status As StepStatus = StepStatus.Pending
        Public Property ErrorMessage As String
        Public Property LastExplanation As ExecutionExplanation
    End Class

    Public Class ExecutionExplanation
        Public Property StepIndex As Integer
        Public Property StepDescription As String
        Public Property ToolId As String
        Public Property ToolName As String
        Public Property ToolCategory As String
        Public Property RiskLevel As String
        Public Property ParametersJson As String
        Public Property StartedAt As DateTime
        Public Property FinishedAt As DateTime
        Public Property ElapsedMs As Long
        Public Property BeforeSummary As String
        Public Property AfterSummary As String
        Public Property ObservationJson As String
        Public Property DataSummaryJson As String
        Public Property Success As Boolean
        Public Property Message As String
        Public Property SkillName As String
        Public Property ScriptFileName As String
        Public Property McpToolName As String
        Public Property McpStatus As String
        Public Property FailureReason As String
        Public Property UndoPointName As String
        Public Property UndoHint As String
        Public Property CanUndo As Boolean
        Public Property AutoRepairSummary As String
        Public Property FixAttempts As Integer
        Public Property ExplanationText As String
    End Class

    ''' <summary>
    ''' Agent 执行结果
    ''' </summary>
    Public Class AgentResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property SessionId As String
        Public Property IterationsCompleted As Integer
        ''' <summary>用户协作式停止产生的终态，区别于普通失败。</summary>
        Public Property IsCancelled As Boolean
        ''' <summary>本轮所有写操作的变更日志；失败或取消时同样保留已落地的记录。</summary>
        Public Property ChangeLog As New AgentRunChangeLog()

        Public Shared Function SuccessResult(sessionId As String,
                                             Optional message As String = "") As AgentResult
            Return New AgentResult With {
                .Success = True,
                .SessionId = sessionId,
                .Message = message
            }
        End Function

        Public Shared Function Failed(sessionId As String, message As String) As AgentResult
            Return New AgentResult With {
                .Success = False,
                .SessionId = sessionId,
                .Message = message
            }
        End Function

        Public Shared Function CancelledResult(sessionId As String, message As String) As AgentResult
            Return New AgentResult With {
                .Success = False,
                .IsCancelled = True,
                .SessionId = sessionId,
                .Message = message
            }
        End Function
    End Class

    ''' <summary>
    ''' Agent 状态枚举
    ''' </summary>
    Public Enum AgentStatus
        Idle
        Thinking
        Planning
        WaitingApproval
        Executing
        Observing
        Reflecting
        Completed
        Failed
        Aborted
    End Enum

    ''' <summary>
    ''' 步骤状态枚举
    ''' </summary>
    Public Enum StepStatus
        Pending
        Running
        Completed
        Failed
        Skipped
    End Enum

    ''' <summary>
    ''' 任务规格（Spec驱动）
    ''' </summary>
    Public Class AgentTaskSpec
        Public Property Goal As String = ""
        Public Property TargetObject As String = ""
        Public Property Constraints As New List(Of String)()
        Public Property SuccessCriteria As New List(Of String)()
        Public Property RequiredTools As New List(Of String)()
        Public Property RiskLevel As String = "safe"
        Public Property Complexity As String = "medium"
        Public Property ExpectedOutputs As New List(Of String)()
        Public Property ExpectedSlideCount As Integer = 0

        Public ReadOnly Property IsSimple As Boolean
            Get
                Return String.Equals(Complexity, "simple", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property
    End Class

End Namespace
