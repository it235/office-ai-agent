' ShareRibbon\Services\Reformat\SemanticPromptBuilder.vb
' 统一构建语义标注提示词

Imports System.Text

''' <summary>
''' 语义提示词构建器 - 为AI构建语义标注的系统提示词
''' 模板排版和规范排版共用同一构建逻辑
''' </summary>
Public Class SemanticPromptBuilder

    ''' <summary>
    ''' 构建语义标注提示词
    ''' </summary>
    ''' <param name="mapping">语义样式映射（包含可用标签）</param>
    ''' <param name="paragraphs">采样段落列表</param>
    Public Shared Function BuildSemanticTaggingPrompt(
        mapping As SemanticStyleMapping,
        paragraphs As List(Of String)) As String

        Dim sb As New StringBuilder()

        ' 系统角色
        sb.AppendLine("你是文档语义分层专家。请为每个段落标注语义标签。")
        sb.AppendLine()

        ' 可用标签列表
        sb.AppendLine("【可用标签】")
        For Each tag In mapping.SemanticTags
            sb.Append($"- {tag.TagId}: {tag.DisplayName}")
            If Not String.IsNullOrEmpty(tag.MatchHint) Then
                sb.Append($"（提示：{tag.MatchHint}）")
            End If
            sb.AppendLine()
        Next
        sb.AppendLine()

        ' 严格要求
        sb.AppendLine("【严格要求】")
        sb.AppendLine("1. 仅使用上述标签，禁止自创标签")
        sb.AppendLine("2. 不要输出任何格式参数（字体、字号、颜色、缩进等）")
        sb.AppendLine("3. 返回纯JSON数组，不要包含markdown代码块标记")
        sb.AppendLine("4. 格式: [{""paraIndex"":0, ""tag"":""title.1""}, ...]")
        sb.AppendLine("5. 层级合理：title.1 后可接 title.2 或 body，不能直接接 title.3")
        sb.AppendLine("6. 每个段落必须且只能有一个标签")
        sb.AppendLine()

        ' 文档段落
        sb.AppendLine("【文档段落】")
        For i = 0 To paragraphs.Count - 1
            ' 截取段落前100字符，避免token浪费
            Dim text = paragraphs(i)
            If text.Length > 100 Then text = text.Substring(0, 100) & "..."
            sb.AppendLine($"[{i}] {text}")
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建带重试提示的标注提示词（当校验失败时使用）
    ''' </summary>
    Public Shared Function BuildRetryPrompt(
        mapping As SemanticStyleMapping,
        paragraphs As List(Of String),
        errors As List(Of String)) As String

        Dim sb As New StringBuilder()

        ' 原始提示词
        sb.Append(BuildSemanticTaggingPrompt(mapping, paragraphs))
        sb.AppendLine()

        ' 错误反馈
        sb.AppendLine("【上次输出存在以下错误，请修正】")
        For Each errMsg In errors
            sb.AppendLine($"- {errMsg}")
        Next
        sb.AppendLine()
        sb.AppendLine("请重新输出正确的JSON数组。")

        Return sb.ToString()
    End Function
End Class
