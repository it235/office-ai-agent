' ShareRibbon\Storage\AppTypeContract.vb
' app_type 的合法值口径。六张表（conversation_event / session_summary / agent_run /
' agent_run_change / atomic_memory / memory_item）都用这一列做宿主过滤，读侧一律是
' 「等于当前宿主 或 空串」，所以第三种值既不等于任何宿主也拿不到空串豁免。

''' <summary>
''' 会话/记忆归属的唯一口径：三个规范宿主名，或空串表示通用数据（三端都可见）。
''' 任何写这六列的地方都必须先过 Normalize，禁止把调用方传来的原始字符串直接落库。
''' </summary>
Public Class AppTypeContract

    Public Const Word As String = "Word"
    Public Const Excel As String = "Excel"
    Public Const PowerPoint As String = "PowerPoint"
    Public Const Generic As String = ""

    ''' <summary>
    ''' 压成规范名；白名单之外（"Unknown"、大小写与首尾空白变体、空）一律返回空串。
    ''' 空串不是"归属未知所以谁都看不见"，而是"通用数据，三端都看得见"：
    ''' 老数据宁可多端可见，也不能在升级后凭空消失。迁移 019 用同一条规则收口历史行。
    ''' </summary>
    Public Shared Function Normalize(appType As String) As String
        Select Case If(appType, "").Trim().ToLowerInvariant()
            Case "word"
                Return Word
            Case "excel"
                Return Excel
            Case "powerpoint"
                Return PowerPoint
            Case Else
                Return Generic
        End Select
    End Function

End Class
