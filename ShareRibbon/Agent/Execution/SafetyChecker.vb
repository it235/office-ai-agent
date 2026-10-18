Namespace Agent.Execution

    ''' <summary>
    ''' 代码安全检查器 - 拦截危险操作
    ''' </summary>
    Public Class SafetyChecker

        ''' <summary>
        ''' 危险操作黑名单（绝对禁止执行）。全部使用去空白后的小写片段，
        ''' 避免 "Kill (""a.txt"")"、"Shell (" 这类加空格或改大小写的绕过。
        ''' </summary>
        Private Shared ReadOnly DANGEROUS() As String = {
            "kill(",
            "kill""",
            "shell(",
            "shell""",
            "shell.application",
            "wscript.shell",
            "createobject(""wscript",
            "createobject(""scripting.filesystemobject",
            "scriptcontrol",
            ".run(",
            ".deletefile",
            ".deletefolder",
            ".movehere",
            ".quit(",
            "environ(",
            "regwrite",
            "regdelete",
            "application.quit",
            "activeworkbook.closefalse",
            "documents.close",
            "presentation.close"
        }

        ''' <summary>
        ''' 需要用户确认的操作（允许但需确认）
        ''' </summary>
        Private Shared ReadOnly NEEDS_CONFIRM() As String = {
            ".delete(",
            ".delete",
            ".clear",
            ".clearcontents",
            "workbooks.close",
            ".saveas",
            ".savecopyas",
            "activedocument.close",
            "selection.delete",
            "range.delete",
            "worksheet.delete",
            "slide.delete",
            ".export(",
            ".printout(",
            ".sendmail(",
            ".protect"
        }

        ''' <summary>
        ''' 检查代码安全性
        ''' </summary>
        ''' <param name="code">待检查的代码</param>
        ''' <returns>安全检查结果</returns>
        Public Shared Function Check(code As String) As SafetyResult
            If String.IsNullOrWhiteSpace(code) Then
                Return SafetyResult.Safe()
            End If

            Dim normalized = Normalize(code)
            Dim matchedNeedConfirm As String = Nothing

            ' 1. 检测危险操作（绝对禁止）
            For Each danger In DANGEROUS
                If normalized.Contains(danger) Then
                    Return SafetyResult.Blocked($"检测到危险操作: {danger}")
                End If
            Next

            ' 2. 检测需要确认的操作
            For Each item In NEEDS_CONFIRM
                If normalized.Contains(item) Then
                    If matchedNeedConfirm Is Nothing Then matchedNeedConfirm = item
                End If
            Next

            If matchedNeedConfirm IsNot Nothing Then
                Return SafetyResult.NeedConfirm($"将执行: {matchedNeedConfirm}")
            End If

            ' 3. 通过安全检查
            Return SafetyResult.Safe()
        End Function

        ''' <summary>
        ''' 去掉空白并转小写，使黑名单无法用空格或大小写绕过
        ''' </summary>
        Private Shared Function Normalize(code As String) As String
            Dim builder As New Text.StringBuilder(code.Length)
            For Each ch In code
                If Not Char.IsWhiteSpace(ch) Then builder.Append(Char.ToLowerInvariant(ch))
            Next
            Return builder.ToString()
        End Function

    End Class

    ''' <summary>
    ''' 安全检查结果
    ''' </summary>
    Public Class SafetyResult

        ''' <summary>
        ''' 是否安全（可以执行）
        ''' </summary>
        Public Property IsSafe As Boolean

        ''' <summary>
        ''' 是否需要用户确认
        ''' </summary>
        Public Property NeedsConfirm As Boolean

        ''' <summary>
        ''' 原因描述
        ''' </summary>
        Public Property Reason As String

        ''' <summary>
        ''' 检测到的危险操作列表
        ''' </summary>
        Public Property DetectedIssues As List(Of String)

        Public Sub New()
            DetectedIssues = New List(Of String)()
        End Sub

        ''' <summary>
        ''' 创建"安全通过"结果
        ''' </summary>
        Public Shared Function Safe() As SafetyResult
            Return New SafetyResult With {
                .IsSafe = True,
                .NeedsConfirm = False,
                .Reason = "安全检查通过"
            }
        End Function

        ''' <summary>
        ''' 创建"被拦截"结果
        ''' </summary>
        Public Shared Function Blocked(reason As String) As SafetyResult
            Return New SafetyResult With {
                .IsSafe = False,
                .NeedsConfirm = False,
                .Reason = reason
            }
        End Function

        ''' <summary>
        ''' 创建"需要确认"结果
        ''' </summary>
        Public Shared Function NeedConfirm(reason As String) As SafetyResult
            Return New SafetyResult With {
                .IsSafe = True,
                .NeedsConfirm = True,
                .Reason = reason
            }
        End Function

        ''' <summary>
        ''' 转为用户友好的提示文本
        ''' </summary>
        Public Function ToUserMessage() As String
            If Not IsSafe Then
                Return $"⛔ 安全拦截: {Reason}"
            ElseIf NeedsConfirm Then
                Return $"⚠️ 需要确认: {Reason}"
            Else
                Return $"✅ {Reason}"
            End If
        End Function

    End Class

End Namespace
