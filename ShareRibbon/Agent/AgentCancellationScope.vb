' ShareRibbon\Agent\AgentCancellationScope.vb
' Ambient cancellation scope for one user turn, so a deep LLM call can observe "停止"
' without threading a CancellationToken through every AI-request delegate in the chain.

Imports System.Threading

Namespace Agent

    ''' <summary>
    ''' 当前用户回合的取消作用域。只有真正属于这次回合的网络请求才读取它；
    ''' 记忆沉淀、embedding、意图识别等后台链路不读，因此不会被用户点"停止"连带掐断。
    ''' </summary>
    Public Class AgentCancellationScope

        Private Shared ReadOnly _current As New AsyncLocal(Of CancellationToken)

        ''' <summary>
        ''' 当前流程可见的取消令牌；不在任何回合作用域内时为 CancellationToken.None。
        ''' </summary>
        Public Shared ReadOnly Property Current As CancellationToken
            Get
                Return _current.Value
            End Get
        End Property

        Public Shared ReadOnly Property IsActive As Boolean
            Get
                Return _current.Value.CanBeCanceled
            End Get
        End Property

        ''' <summary>
        ''' 进入回合作用域。Using 结束时自动复位，避免同一线程的后续请求继承上一次回合的令牌。
        ''' </summary>
        Public Shared Function Enter(token As CancellationToken) As IDisposable
            Dim previous = _current.Value
            _current.Value = token
            Return New Scope(previous)
        End Function

        Private NotInheritable Class Scope
            Implements IDisposable

            Private ReadOnly _previous As CancellationToken
            Private _disposed As Boolean

            Public Sub New(previous As CancellationToken)
                _previous = previous
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If _disposed Then Return
                _disposed = True
                _current.Value = _previous
            End Sub
        End Class
    End Class

End Namespace
