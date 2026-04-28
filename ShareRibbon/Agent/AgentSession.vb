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
        Public Property Spec As AgentTaskSpec

        ' ReAct 循环记录
        Public Property Iterations As New List(Of ReActIteration)
        Public Property Status As AgentStatus = AgentStatus.Idle
        Public Property StartTime As DateTime = DateTime.Now
        Public Property CurrentIteration As Integer = 0
        Public Property MaxIterations As Integer = 15

        ' 执行计划
        Public Property Plan As ExecutionPlan

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
    End Class

    ''' <summary>
    ''' Agent 执行结果
    ''' </summary>
    Public Class AgentResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property SessionId As String
        Public Property IterationsCompleted As Integer
        Public Property FinalOutput As String

        Public Shared Function SuccessResult(sessionId As String,
                                              Optional message As String = "",
                                              Optional finalOutput As String = "") As AgentResult
            Return New AgentResult With {
                .Success = True,
                .SessionId = sessionId,
                .Message = message,
                .FinalOutput = finalOutput
            }
        End Function

        Public Shared Function Failed(sessionId As String, message As String) As AgentResult
            Return New AgentResult With {
                .Success = False,
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
        Public Property Constraints As New List(Of String)()
        Public Property SuccessCriteria As New List(Of String)()
        Public Property Complexity As String = "medium"

        Public ReadOnly Property IsSimple As Boolean
            Get
                Return String.Equals(Complexity, "simple", StringComparison.OrdinalIgnoreCase)
            End Get
        End Property
    End Class

End Namespace
