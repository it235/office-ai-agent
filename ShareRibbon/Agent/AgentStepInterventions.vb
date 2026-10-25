Imports System.Collections.Generic

Namespace Agent

    ''' <summary>
    ''' 用户对某一步的纠偏决定。词汇刻意与反思分支一致（retry / skip），只是把决定权从模型
    ''' 手里换回用户——多一套说法就多一份需要人工同步的语义，也会让"失败步骤能做什么"出现两种合同。
    ''' </summary>
    Public Enum StepIntervention
        None = 0
        Retry = 1
        Skip = 2
    End Enum

    ''' <summary>
    ''' 本轮用户纠偏收件箱。面板只允许往一个步号上放决定，真正取用发生在 LoopEngine 的失败分支里：
    ''' 循环自己拉、UI 不往执行中段推，所以不需要第二条裁决通道，也不必给审批握手加第二种用途。
    ''' 按步号寻址、同号覆盖（重复点击等于同一个决定），计划被整体替换时必须清空——新计划里的同号不是同号。
    ''' </summary>
    Public NotInheritable Class AgentStepInterventions

        Private ReadOnly _lock As New Object()
        Private ReadOnly _pending As New Dictionary(Of Integer, StepIntervention)

        Public Sub Queue(stepIndex As Integer, action As StepIntervention)
            If stepIndex < 0 OrElse action = StepIntervention.None Then Return
            SyncLock _lock
                _pending(stepIndex) = action
            End SyncLock
        End Sub

        ''' <summary>取走这一步的用户决定并在同一瞬间失效：一条决定只许作用于一次失败。</summary>
        Public Function Take(stepIndex As Integer) As StepIntervention
            SyncLock _lock
                Dim action As StepIntervention
                If Not _pending.TryGetValue(stepIndex, action) Then Return StepIntervention.None
                _pending.Remove(stepIndex)
                Return action
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
            Return StepIntervention.None
        End Function
    End Class

End Namespace
