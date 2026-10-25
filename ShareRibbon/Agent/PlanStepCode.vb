Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agent

    ''' <summary>
    ''' 计划步骤 code 的唯一解析者。规划提示词规定 code 是
    ''' {"command":"工具ID","params":{...}} 的 JSON 文本，面板改参数时宿主必须用同一份判定——
    ''' 两份解析的后果不是报错，而是"卡片说能保存、执行时却按另一套读法"。
    ''' 只认这一个形状：模型偶尔会写别的键名，那种步骤由 Think 路径执行，不走用户改写的直达路径。
    ''' </summary>
    Public NotInheritable Class PlanStepCode

        Private Sub New()
        End Sub

        Public Shared Function TryParse(code As String, ByRef toolCall As ToolCall) As Boolean
            toolCall = Nothing
            If String.IsNullOrWhiteSpace(code) Then Return False

            Try
                Dim obj = JObject.Parse(code)
                If obj Is Nothing Then Return False

                Dim toolId = obj("command")?.ToString()
                If String.IsNullOrWhiteSpace(toolId) Then Return False

                Dim paramsToken = obj("params")
                Dim parameters As JObject = Nothing
                If paramsToken Is Nothing OrElse paramsToken.Type = JTokenType.Null Then
                    parameters = New JObject()
                Else
                    parameters = TryCast(paramsToken, JObject)
                    ' 参数不是对象时宁可判为不可用：把写错的参数袋悄悄换成空袋，等于执行了用户没写的操作。
                    If parameters Is Nothing Then Return False
                End If

                toolCall = New ToolCall With {
                    .ToolId = toolId.Trim(),
                    .Parameters = parameters
                }
                Return True
            Catch ex As JsonException
                Return False
            End Try
        End Function

        ''' <summary>只关心工具身份时用这个（面板改写必须保持工具不变）。</summary>
        Public Shared Function TryGetToolId(code As String, ByRef toolId As String) As Boolean
            toolId = ""
            Dim toolCall As ToolCall = Nothing
            If Not TryParse(code, toolCall) Then Return False
            toolId = toolCall.ToolId
            Return True
        End Function
    End Class

End Namespace
