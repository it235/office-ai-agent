' ShareRibbon\Services\Reformat\SemanticPromptBuilder.vb
' 统一构建语义标注提示词

Imports System.Text

''' <summary>
''' 语义提示词构建器 - 为AI构建语义标注的系统提示词
''' 模板排版和规范排版共用同一构建逻辑
''' </summary>
Public Class SemanticPromptBuilder

    ''' <summary>
    ''' 构建语义标注提示词（带样式上下文）
    ''' </summary>
    ''' <param name="mapping">语义样式映射（包含可用标签）</param>
    ''' <param name="paragraphs">段落文本列表（仅文本段落，非文本已过滤）</param>
    ''' <param name="paragraphStyles">段落样式名称（仅文本段落，与paragraphs一一对应）</param>
    ''' <param name="originalParaIndices">原文档中的段落索引（仅文本段落，用于映射回正确位置）</param>
    ''' <param name="detectedHeadings">DocumentAnalyzer检测到的标题信息</param>
    Public Shared Function BuildSemanticTaggingPrompt(
        mapping As SemanticStyleMapping,
        paragraphs As List(Of String),
        Optional paragraphStyles As List(Of String) = Nothing,
        Optional originalParaIndices As List(Of Integer) = Nothing,
        Optional detectedHeadings As String = Nothing,
        Optional documentTypeContext As String = Nothing,
        Optional paragraphFontSizes As List(Of Single) = Nothing,
        Optional paragraphIsBold As List(Of Boolean) = Nothing) As String

        Dim sb As New StringBuilder()

        ' 系统角色 — 极其重要，必须前置
        sb.AppendLine("你是一个严格的JSON输出器。你必须只输出一个JSON数组，不要输出任何其他内容。")
        sb.AppendLine()
        sb.AppendLine("【绝对禁止】")
        sb.AppendLine("- 禁止输出VBA代码、Sub/End Sub、宏代码")
        sb.AppendLine("- 禁止输出Python代码、JavaScript代码")
        sb.AppendLine("- 禁止输出markdown代码块（```json 或 ```）")
        sb.AppendLine("- 禁止输出任何解释、说明、注释")
        sb.AppendLine("- 禁止输出格式参数（字体名、字号数值、颜色代码等），只输出语义标签")
        sb.AppendLine("- 如果文档类型不明，仍然要标注，使用最通用的body.normal标签")
        sb.AppendLine()

        ' 文档类型上下文 — 帮助AI理解应该用什么视角标注
        If Not String.IsNullOrEmpty(documentTypeContext) Then
            sb.AppendLine("【文档类型与排版标准】")
            sb.AppendLine(documentTypeContext)
            sb.AppendLine("请基于上述文档类型和排版标准进行语义标注。例如公文中的「发文字号」段应标注为header.refno，而非简单的heading。")
            sb.AppendLine()
        End If

        sb.AppendLine("【你的唯一输出格式】")
        sb.AppendLine("[{""paraIndex"":0, ""tag"":""body.normal""}, {""paraIndex"":1, ""tag"":""heading.1""}]")
        sb.AppendLine()
        sb.AppendLine("【重要提示】")
        sb.AppendLine("1. 每个段落给出了「原文样式」和「原文格式线索」，请结合文本内容综合判断该段落的语义角色。")
        sb.AppendLine("2. 原文样式名为「标题 1」「Heading 1」「标题 2」等的段落，大概率就是对应级别的标题。")
        sb.AppendLine("3. 原文样式名为「正文」「Normal」且字数多(>100字)的段落，通常是正文。")
        sb.AppendLine("4. 加粗且字号偏大的短段落，大概率是标题。")
        sb.AppendLine("5. 字数很少(<30字)、居中的段落，通常是标题。")
        sb.AppendLine("6. 包含发文号格式(如「XX发〔20XX〕X号」)的段落，是公文发文字号。")
        sb.AppendLine("7. 包含日期且位于文末的短段落，通常是落款/署名。")
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

        ' 自动检测到的标题结构（来自DocumentAnalyzer）
        If Not String.IsNullOrEmpty(detectedHeadings) Then
            sb.AppendLine("【AI自动检测到的标题结构（仅供参考）】")
            sb.AppendLine(detectedHeadings)
            sb.AppendLine()
        End If

        ' 严格要求
        sb.AppendLine("【严格要求】")
        sb.AppendLine("1. 仅使用上述标签，禁止自创标签")
        sb.AppendLine("2. 不要输出任何格式参数（字体、字号、颜色、缩进等）")
        sb.AppendLine("3. 返回纯JSON数组，不要包含markdown代码块标记")
        sb.AppendLine("4. 格式: [{""paraIndex"":0, ""tag"":""title.1""}, ...]")
        sb.AppendLine("5. 层级合理：title.1 后可接 title.2 或 body，不能直接接 title.3")
        sb.AppendLine("6. 每个段落必须且只能有一个标签")
        sb.AppendLine("7. paraIndex 使用上面给出的段落索引号（第一个数字）")
        sb.AppendLine()

        ' 文档段落（仅文本段落）
        sb.AppendLine("【文档段落】")
        Dim hasStyles = paragraphStyles IsNot Nothing AndAlso paragraphStyles.Count = paragraphs.Count
        Dim hasOrigIdx = originalParaIndices IsNot Nothing AndAlso originalParaIndices.Count = paragraphs.Count
        Dim hasFontSizes = paragraphFontSizes IsNot Nothing AndAlso paragraphFontSizes.Count = paragraphs.Count
        Dim hasBold = paragraphIsBold IsNot Nothing AndAlso paragraphIsBold.Count = paragraphs.Count

        For i = 0 To paragraphs.Count - 1
            Dim origIdx = If(hasOrigIdx, originalParaIndices(i), i)
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            ' 样式提示
            Dim styleHint As String = ""
            If hasStyles AndAlso Not String.IsNullOrEmpty(paragraphStyles(i)) Then
                styleHint = $" [样式:{paragraphStyles(i)}]"
            End If

            ' 字号+加粗提示
            Dim formatHint As String = ""
            If hasFontSizes Then
                formatHint = $" {paragraphFontSizes(i):F0}pt"
            End If
            If hasBold AndAlso paragraphIsBold(i) Then
                formatHint &= " 加粗"
            End If
            If formatHint <> "" Then
                formatHint = $" [格式:{formatHint.Trim()}]"
            End If

            ' 截取段落前120字符
            If text.Length > 120 Then text = text.Substring(0, 120) & "..."

            sb.Append($"[{origIdx}]{styleHint}{formatHint} {text}")
            sb.AppendLine()
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
