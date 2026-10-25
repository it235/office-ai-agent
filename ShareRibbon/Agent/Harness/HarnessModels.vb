Imports System.Collections.Generic

Namespace Agent.Harness

    Public Enum HarnessRunStatus
        Running
        AwaitingApproval
        Succeeded
        Failed
        Cancelled
    End Enum

    Public Class UserTurn
        Public Property TurnId As String = Guid.NewGuid().ToString()
        Public Property SessionId As String = ""
        Public Property AppType As String = ""
        Public Property Text As String = ""
        Public Property Mode As String = "agent"
        Public Property References As New List(Of String)()
        Public Property HostContextText As String = ""
        Public Property OfficeContext As Agent.Context.OfficeContext
        Public Property ContextPack As Agent.Context.ContextPack
        Public Property TaskSpec As Agent.AgentTaskSpec
        Public Property SelectedSkills As New List(Of SkillFileDefinition)()
        ''' <summary>入口动作 id；空表示自由文本，任务包断言据此决定要不要核对（E6-05）。</summary>
        Public Property EntryActionId As String = ""
    End Class

    Public Class HarnessRunResult
        Public Property RunId As String = Guid.NewGuid().ToString()
        Public Property Status As HarnessRunStatus = HarnessRunStatus.Running
        Public Property UserMessage As String = ""
        Public Property DebugMessage As String = ""
        Public Property AgentSessionId As String = ""
        Public Property StartedAt As DateTime
        Public Property FinishedAt As DateTime
    End Class

    Public Class HarnessPhaseChangedEventArgs
        Inherits EventArgs

        Public Property RunId As String = ""
        Public Property Phase As String = ""
        Public Property Message As String = ""
    End Class

    Public Class HarnessStepChangedEventArgs
        Inherits EventArgs

        Public Property RunId As String = ""
        Public Property StepIndex As Integer
        Public Property ToolId As String = ""
        Public Property Description As String = ""
        Public Property Status As String = ""
        Public Property Message As String = ""
        Public Property ErrorCode As String = ""
    End Class

    Public Class HarnessContextEventArgs
        Inherits EventArgs

        Public Property RunId As String = ""
        Public Property AppType As String = ""
        Public Property ContextText As String = ""
        Public Property ContextPackJson As String = ""
    End Class

End Namespace
