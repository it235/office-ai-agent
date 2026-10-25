Imports System.Collections.Generic

Namespace Agent

    ''' <summary>
    ''' 用户对某一步的决定。retry / skip 的词汇刻意与反思分支一致，只是把决定权从模型
    ''' 手里换回用户——多一套说法就多一份需要人工同步的语义，也会让"失败步骤能做什么"出现两种合同。
    ''' edit 是第三条：不推进计划版本，只把这一步的说明与参数换成用户的版本（E10-02）。
    ''' </summary>
    Public Enum StepIntervention
        None = 0
        Retry = 1
        Skip = 2
        Edit = 3
    End Enum

    ''' <summary>
    ''' 一次单步编辑能不能排队。机器原因用枚举，面向用户的那句话留给 UI 层——
    ''' 把句子写在这里，没有面板的调用方（自动化、真机脚本）就只能拿一句中文去判断分支。
    ''' </summary>
    Public Enum StepEditResult
        Queued = 0
        NoLiveRun = 1
        StepOutOfRange = 2
        StepNotEditable = 3
        EmptyEdit = 4
        CodeInvalid = 5
        ToolChanged = 6
    End Enum

    ''' <summary>
    ''' 一条决定连同它的载荷。edit 要带上改写后的说明与参数文本，retry/skip 两个字段留空。
    ''' 载荷必须跟着决定一起被取走：分两次取会出现"决定生效了、参数没跟上"的中间态。
    ''' </summary>
    Public Class StepInterventionRequest
        Public Property Action As StepIntervention = StepIntervention.None
        Public Property Description As String = ""
        Public Property Code As String = ""
    End Class

    ''' <summary>
    ''' 本轮用户纠偏收件箱。面板只允许往一个步号上放决定，真正取用发生在 LoopEngine 的步骤边界：
    ''' 循环自己拉、UI 不往执行中段推，所以不需要第二条裁决通道，也不必给审批握手加第二种用途。
    ''' 按步号寻址、同号覆盖（重复点击等于同一个决定），计划被整体替换时必须清空——新计划里的同号不是同号。
    ''' </summary>
    Public NotInheritable Class AgentStepInterventions

        Private ReadOnly _lock As New Object()
        Private ReadOnly _pending As New Dictionary(Of Integer, StepInterventionRequest)

        ''' <summary>
        ''' 收下一个不需要载荷的决定。edit 不从这里进：没有说明与参数文本的编辑等于让用户按了没反应，
        ''' 所以它只有 <see cref="QueueEdit"/> 一个入口。
        ''' </summary>
        Public Sub Queue(stepIndex As Integer, action As StepIntervention)
            If stepIndex < 0 OrElse action = StepIntervention.None OrElse action = StepIntervention.Edit Then Return
            SyncLock _lock
                _pending(stepIndex) = New StepInterventionRequest With {.Action = action}
            End SyncLock
        End Sub

        ''' <summary>改这一步的说明与参数。同号后写的覆盖先写的——用户在生效前改两次，以最后一次为准。</summary>
        Public Sub QueueEdit(stepIndex As Integer, description As String, code As String)
            If stepIndex < 0 OrElse String.IsNullOrWhiteSpace(code) Then Return
            SyncLock _lock
                _pending(stepIndex) = New StepInterventionRequest With {
                    .Action = StepIntervention.Edit,
                    .Description = If(description, ""),
                    .Code = code
                }
            End SyncLock
        End Sub

        ''' <summary>取走这一步的用户决定并在同一瞬间失效：一条决定只许作用于一次失败或一次重跑。</summary>
        Public Function Take(stepIndex As Integer) As StepInterventionRequest
            SyncLock _lock
                Dim request As StepInterventionRequest = Nothing
                If Not _pending.TryGetValue(stepIndex, request) Then Return Nothing
                _pending.Remove(stepIndex)
                Return request
            End SyncLock
        End Function

        Public Function PendingCount() As Integer
            SyncLock _lock
                Return _pending.Count
            End SyncLock
        End Function

        ''' <summary>待办步号的快照（升序）。循环按这个顺序落地，回退才总是退到最靠前的那一步。</summary>
        Public Function PendingIndices() As List(Of Integer)
            SyncLock _lock
                Dim keys = New List(Of Integer)(_pending.Keys)
                keys.Sort()
                Return keys
            End SyncLock
        End Function

        Public Sub Clear()
            SyncLock _lock
                _pending.Clear()
            End SyncLock
        End Sub

        ''' <summary>把面板传来的字符串收成枚举；认不出来就是 None，不能猜成"用户想重试"。</summary>
        Public Shared Function Parse(value As String) As StepIntervention
            If String.Equals(value, "retry", StringComparison.OrdinalIgnoreCase) Then Return StepIntervention.Retry
            If String.Equals(value, "skip", StringComparison.OrdinalIgnoreCase) Then Return StepIntervention.Skip
            If String.Equals(value, "edit", StringComparison.OrdinalIgnoreCase) Then Return StepIntervention.Edit
            Return StepIntervention.None
        End Function
    End Class

End Namespace
