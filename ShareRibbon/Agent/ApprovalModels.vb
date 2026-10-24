Namespace Agent

    ''' <summary>
    ''' 用户对高风险工具调用的裁决粒度。
    ''' Once 只放行当前参数组合；Session 在本轮 Agent 会话内放行该工具的全部调用；
    ''' Deny 跳过当前步骤；Abort 用于超时、取消或对话被替代，直接停止本轮。
    ''' </summary>
    Public Enum ApprovalDecision
        Deny = 0
        Once = 1
        Session = 2
        Abort = 3
    End Enum

    ''' <summary>
    ''' 审批请求：工具标识与风险信息随提示一起送到 UI，避免用户只看到一句泛泛的“高风险操作”。
    ''' </summary>
    Public Class ApprovalRequest
        Public Property ToolId As String = ""
        Public Property Message As String = ""
        Public Property RiskLevel As String = "risky"
    End Class

End Namespace
