Imports Newtonsoft.Json.Linq

''' <summary>
''' Provider 回包里的 token 用量。全仓只有这一处负责“从回包里读出用量”：OpenAI 兼容用
''' prompt_tokens/completion_tokens/total_tokens，Anthropic 用 input_tokens/output_tokens，
''' 两个形状在这里收成一份合同，调用方不再各自 SelectToken。
''' 读不到 usage 时返回 Nothing——用量缺失是可以说清楚的，编一个数字不行。
''' </summary>
Public Class TokenUsage
    Public Property PromptTokens As Integer
    Public Property CompletionTokens As Integer
    Public Property TotalTokens As Integer

    Public Function Copy() As TokenUsage
        Return New TokenUsage With {
            .PromptTokens = PromptTokens,
            .CompletionTokens = CompletionTokens,
            .TotalTokens = TotalTokens
        }
    End Function

    ''' <summary>把一次回包的用量累加进当前账本（账本自身可为空，返回累加后的同一个实例）。</summary>
    Public Function Add(usage As TokenUsage) As TokenUsage
        If usage Is Nothing Then Return Me
        PromptTokens += usage.PromptTokens
        CompletionTokens += usage.CompletionTokens
        TotalTokens += usage.TotalTokens
        Return Me
    End Function

    Public Shared Function FromResponse(responseText As String) As TokenUsage
        If String.IsNullOrWhiteSpace(responseText) Then Return Nothing
        Try
            Return FromJson(TryCast(JObject.Parse(responseText), JObject))
        Catch ex As Exception
            Debug.WriteLine($"[TokenUsage] usage 解析失败: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Public Shared Function FromJson(root As JObject) As TokenUsage
        If root Is Nothing Then Return Nothing

        Dim usage = TryCast(root("usage"), JObject)
        If usage Is Nothing Then Return Nothing

        Dim prompt = ReadInt(usage, "prompt_tokens", "input_tokens")
        Dim completion = ReadInt(usage, "completion_tokens", "output_tokens")
        Dim total = ReadInt(usage, "total_tokens")
        If prompt < 0 AndAlso completion < 0 AndAlso total < 0 Then Return Nothing

        Dim result As New TokenUsage With {
            .PromptTokens = Math.Max(0, prompt),
            .CompletionTokens = Math.Max(0, completion)
        }
        ' 有 total 就用 provider 的：部分接口把推理 token 单独计入 total，自己相加会低估。
        If total > 0 Then
            result.TotalTokens = total
        Else
            result.TotalTokens = result.PromptTokens + result.CompletionTokens
        End If
        Return result
    End Function

    ''' <summary>按别名顺序取第一个存在的非负整数；都没有时返回 -1，表示“这个形状里没有”。</summary>
    Private Shared Function ReadInt(usage As JObject, ParamArray names() As String) As Integer
        For Each name In names
            Dim token = usage(name)
            If token IsNot Nothing AndAlso token.Type <> JTokenType.Null Then
                Dim parsed As Integer
                If Integer.TryParse(token.ToString(), parsed) AndAlso parsed >= 0 Then Return parsed
            End If
        Next
        Return -1
    End Function
End Class
