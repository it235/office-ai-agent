' ShareRibbon\Loop\ContextCheckResult.vb
' 发送前上下文校验结果

Imports System.Collections.Generic

''' <summary>
''' 发送前上下文校验结果
''' </summary>
Public Class ContextCheckResult

    ''' <summary>是否通过校验</summary>
    Public Property IsValid As Boolean = True

    ''' <summary>错误信息列表（阻止继续）</summary>
    Public Property Errors As List(Of String)

    ''' <summary>警告信息列表（可继续但需记录）</summary>
    Public Property Warnings As List(Of String)

    ''' <summary>建议的用户澄清提示</summary>
    Public Property SuggestedClarification As String = String.Empty

    ''' <summary>是否需要用户确认后继续</summary>
    Public Property NeedsConfirmation As Boolean = False

    Public Sub New()
        Errors = New List(Of String)()
        Warnings = New List(Of String)()
    End Sub

End Class
