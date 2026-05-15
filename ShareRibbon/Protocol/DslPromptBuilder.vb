' ShareRibbon\Protocol\DslPromptBuilder.vb
' DSL Prompt构建器 - 生成让AI产生DSL指令的Prompt

Imports System.Collections.Generic
Imports System.Text

''' <summary>
''' DSL Prompt构建器
''' </summary>
Public Class DslPromptBuilder

    ''' <summary>
    ''' 构建排版DSL Prompt
    ''' </summary>
    Public Shared Function BuildReformatPrompt(
        userRequest As String,
        documentType As String,
        scope As String,
        paragraphCount As Integer,
        selectedParagraphs As List(Of String)) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("你是Office文档智能排版助手。用户要求你对文档进行排版操作。")
        sb.AppendLine()
        sb.AppendLine("【当前文档信息】")
        sb.AppendLine($"- 文档类型: {documentType}")
        sb.AppendLine($"- 处理范围: {scope}")
        sb.AppendLine($"- 段落数: {paragraphCount}")

        If selectedParagraphs IsNot Nothing AndAlso selectedParagraphs.Count > 0 Then
            sb.AppendLine("- 选中段落预览:")
            For i = 0 To Math.Min(selectedParagraphs.Count - 1, 5)
                Dim preview = If(selectedParagraphs(i).Length > 50,
                    selectedParagraphs(i).Substring(0, 50) & "...",
                    selectedParagraphs(i))
                sb.AppendLine($"  [{i}] {preview}")
            Next
            If selectedParagraphs.Count > 5 Then
                sb.AppendLine($"  ... 还有 {selectedParagraphs.Count - 5} 个段落")
            End If
        End If

        sb.AppendLine()
        sb.AppendLine("【你的任务】")
        sb.AppendLine("请分析用户需求，生成一组结构化的操作指令（DSL）。")
        sb.AppendLine()
        sb.AppendLine("【重要规则】")
        sb.AppendLine("1. 你只负责描述'期望的格式状态'，具体的Word DOM操作由系统代码执行")
        sb.AppendLine("2. 不要生成OpenXML，不要生成VBA代码")
        sb.AppendLine("3. 使用声明式指令，描述'应该是什么样子'而非'怎么操作'")
        sb.AppendLine("4. 每个指令必须包含id、op、target、params、expected字段")
        sb.AppendLine("5. target使用语义选择器（如'paragraph[role='heading1']'），不要使用绝对位置")
        sb.AppendLine("6. 如果用户需求不明确，返回空instructions数组并在metadata中说明")
        sb.AppendLine("7. 所有中文字符串值使用正确的引号包裹")
        sb.AppendLine("8. 不要在JSON末尾添加多余的逗号")
        sb.AppendLine()
        sb.AppendLine("【输出格式】")
        sb.AppendLine("必须且只能返回以下JSON格式（不要markdown代码块，不要解释文字）：")
        sb.AppendLine()
        sb.AppendLine("{")
        sb.AppendLine("  ""version"": ""2.0"",")
        sb.AppendLine("  ""protocol"": ""office-dsl"",")
        sb.AppendLine("  ""operation"": ""reformat"",")
        sb.AppendLine("  ""target"": { ""scope"": ""..."", ""documentType"": ""..."" },")
        sb.AppendLine("  ""instructions"": [")
        sb.AppendLine("    {")
        sb.AppendLine("      ""id"": ""inst-001"",")
        sb.AppendLine("      ""op"": ""setParagraphStyle"",")
        sb.AppendLine("      ""target"": { ""type"": ""semantic"", ""selector"": ""paragraph[role='heading1']"", ""index"": 0 },")
        sb.AppendLine("      ""params"": { ""styleName"": ""标题 1"", ""font"": { ""name"": ""仿宋"", ""size"": 16 }, ""alignment"": ""center"" },")
        sb.AppendLine("      ""expected"": { ""description"": ""标题居中仿宋16磅"", ""verifyBy"": ""styleNameAndFont"" },")
        sb.AppendLine("      ""rollback"": { ""op"": ""setParagraphStyle"", ""params"": { ""styleName"": ""__original__"" } }")
        sb.AppendLine("    }")
        sb.AppendLine("  ],")
        sb.AppendLine("  ""metadata"": { ""estimatedOperations"": 1, ""hasDestructiveOps"": false, ""requiresConfirmation"": false }")
        sb.AppendLine("}")
        sb.AppendLine()
        sb.AppendLine(InstructionRegistry.BuildPromptDocumentation("reformat"))
        sb.AppendLine()
        sb.AppendLine("【语义选择器语法】")
        sb.AppendLine("- paragraph[role='heading1'] - 第1级标题段落")
        sb.AppendLine("- paragraph[role='heading2'] - 第2级标题段落")
        sb.AppendLine("- paragraph[role='body'] - 正文段落")
        sb.AppendLine("- paragraph[role='title'] - 文档标题")
        sb.AppendLine("- paragraph:first - 第一个段落")
        sb.AppendLine("- paragraph:last - 最后一个段落")
        sb.AppendLine("- selection - 当前选区")
        sb.AppendLine()
        sb.AppendLine("【支持的排版指令】")
        sb.AppendLine("- setParagraphStyle - 设置段落样式")
        sb.AppendLine("- setCharacterFormat - 设置字符格式")
        sb.AppendLine("- insertTable - 插入表格")
        sb.AppendLine("- formatTable - 格式化表格")
        sb.AppendLine("- setPageSetup - 页面设置")
        sb.AppendLine("- insertBreak - 插入分隔符")
        sb.AppendLine("- applyListFormat - 应用列表格式")
        sb.AppendLine("- setColumnFormat - 分栏设置")
        sb.AppendLine("- insertHeaderFooter - 插入页眉页脚")
        sb.AppendLine("- generateToc - 生成目录")
        sb.AppendLine()
        sb.AppendLine("【用户需求】")
        sb.AppendLine(userRequest)

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建校对DSL Prompt
    ''' </summary>
    Public Shared Function BuildProofreadPrompt(
        userRequest As String,
        paragraphCount As Integer,
        paragraphs As List(Of String),
        checkTypes As List(Of String)) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("你是Word文档智能校对助手。请仔细检查以下文档内容，识别所有需要修正的问题。")
        sb.AppendLine()
        sb.AppendLine("【校对范围】")

        If checkTypes IsNot Nothing AndAlso checkTypes.Count > 0 Then
            For Each ct In checkTypes
                Select Case ct.ToLower()
                    Case "spelling"
                        sb.AppendLine("- 错别字和拼写错误")
                    Case "grammar"
                        sb.AppendLine("- 语法和语病问题")
                    Case "punctuation"
                        sb.AppendLine("- 标点符号错误（中英文标点混用、缺失或多余）")
                    Case "wordusage"
                        sb.AppendLine("- 词语使用错误（的地得混用、他/她/它混用等）")
                    Case "format"
                        sb.AppendLine("- 格式一致性问题")
                End Select
            Next
        Else
            sb.AppendLine("- 错别字和拼写错误")
            sb.AppendLine("- 词语使用错误（的地得混用等）")
            sb.AppendLine("- 标点符号错误")
            sb.AppendLine("- 语法和语病问题")
            sb.AppendLine("- 表达不通顺或容易引起歧义的地方")
        End If

        sb.AppendLine()
        sb.AppendLine("【文档内容】")
        If paragraphs IsNot Nothing Then
            For i = 0 To paragraphs.Count - 1
                Dim para = paragraphs(i)
                If Not String.IsNullOrWhiteSpace(para) Then
                    Dim preview = If(para.Length > 200, para.Substring(0, 200) & "...", para)
                    sb.AppendLine($"[段落{i}] {preview}")
                End If
            Next
        End If

        sb.AppendLine()
        sb.AppendLine("【输出格式】")
        sb.AppendLine("请生成DSL指令，每个建议修正为一个suggestCorrection指令：")
        sb.AppendLine()
        sb.AppendLine("{")
        sb.AppendLine("  ""version"": ""2.0"",")
        sb.AppendLine("  ""protocol"": ""office-dsl"",")
        sb.AppendLine("  ""operation"": ""proofread"",")
        sb.AppendLine("  ""target"": { ""scope"": ""selection"", ""checkTypes"": [""spelling"", ""grammar"", ""punctuation"", ""wordUsage""] },")
        sb.AppendLine("  ""instructions"": [")
        sb.AppendLine("    {")
        sb.AppendLine("      ""id"": ""proof-001"",")
        sb.AppendLine("      ""op"": ""suggestCorrection"",")
        sb.AppendLine("      ""target"": { ""type"": ""textMatch"", ""match"": ""的地得用法错误的地"" },")
        sb.AppendLine("      ""params"": {")
        sb.AppendLine("        ""original"": ""的地得用法错误的地"",")
        sb.AppendLine("        ""suggestion"": ""的地得用法错误的地"",")
        sb.AppendLine("        ""issueType"": ""wordUsageError"",")
        sb.AppendLine("        ""severity"": ""high"",")
        sb.AppendLine("        ""explanation"": ""此处应为'的'，修饰名词""")
        sb.AppendLine("      },")
        sb.AppendLine("      ""expected"": { ""description"": ""建议将'的地得用法错误的地'修正为'的地得用法错误的地'"" }")
        sb.AppendLine("    }")
        sb.AppendLine("  ],")
        sb.AppendLine("  ""metadata"": { ""totalIssues"": 12, ""highSeverity"": 3, ""mediumSeverity"": 5, ""lowSeverity"": 4 }")
        sb.AppendLine("}")
        sb.AppendLine()
        sb.AppendLine("【重要规则】")
        sb.AppendLine("1. original字段必须精确匹配文档原文（包括标点和空格）")
        sb.AppendLine("2. suggestion只写修正后的文本，不要加说明")
        sb.AppendLine("3. 如果文档没有问题，返回空instructions数组")
        sb.AppendLine("4. 不要生成markdown代码块，直接返回JSON")
        sb.AppendLine()
        sb.AppendLine("【支持的校对指令】")
        sb.AppendLine("- suggestCorrection - 文字修正建议")
        sb.AppendLine("- suggestFormatFix - 格式修正建议")
        sb.AppendLine("- suggestStyleUnify - 样式统一建议")
        sb.AppendLine("- markForReview - 标记待审核")
        sb.AppendLine()
        sb.AppendLine("【问题类型】")
        sb.AppendLine("- spellingError - 拼写错误")
        sb.AppendLine("- wordUsageError - 词语使用错误")
        sb.AppendLine("- punctuationError - 标点符号错误")
        sb.AppendLine("- grammaticalError - 语法错误")
        sb.AppendLine("- expressionError - 表达错误")
        sb.AppendLine("- formatError - 格式错误")
        sb.AppendLine()
        sb.AppendLine("【严重程度】")
        sb.AppendLine("- high - 高（严重影响理解）")
        sb.AppendLine("- medium - 中（需要修正但不影响理解）")
        sb.AppendLine("- low - 低（建议优化）")
        sb.AppendLine()
        sb.AppendLine(InstructionRegistry.BuildPromptDocumentation("proofread"))
        sb.AppendLine()
        sb.AppendLine("【用户需求】")
        sb.AppendLine(userRequest)

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建修正Prompt（校验失败时使用）
    ''' </summary>
    Public Shared Function BuildCorrectionPrompt(
        originalResponse As String,
        errors As List(Of InstructionError),
        operation As String) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("你之前返回的操作指令存在以下问题，请修正后重新返回：")
        sb.AppendLine()

        For Each [error] In errors
            sb.AppendLine($"- [{[error].Level}] {[error].Message}")
            If [error].Suggestion IsNot Nothing Then
                sb.AppendLine($"  建议: {[error].Suggestion}")
            End If
        Next

        sb.AppendLine()
        sb.AppendLine("【你之前返回的内容】")
        sb.AppendLine(originalResponse)
        sb.AppendLine()
        sb.AppendLine("【修正要求】")
        sb.AppendLine("1. 仅修正上述错误，不要改变原有操作意图")
        sb.AppendLine("2. 严格按照指令协议格式返回")
        sb.AppendLine("3. 返回纯JSON，不要包含解释文字或代码块标记")
        sb.AppendLine("4. 不要在JSON末尾添加多余的逗号")
        sb.AppendLine("5. 所有字符串使用英文双引号包裹")

        If operation = "reformat" Then
            sb.AppendLine()
            sb.AppendLine(InstructionRegistry.BuildPromptDocumentation("reformat"))
        ElseIf operation = "proofread" Then
            sb.AppendLine()
            sb.AppendLine(InstructionRegistry.BuildPromptDocumentation("proofread"))
        End If

        Return sb.ToString()
    End Function

End Class

''' <summary>
''' DSL协议检测器
''' </summary>
Public Class DslProtocolDetector

    ''' <summary>
    ''' 检测是否为DSL格式
    ''' </summary>
    Public Shared Function IsDslFormat(jsonText As String) As Boolean
        Try
            Dim json = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
            ' DSL格式特征：有version、protocol、instructions字段
            If json("version") IsNot Nothing AndAlso
               json("protocol") IsNot Nothing AndAlso
               json("instructions") IsNot Nothing Then
                Return True
            End If
            ' 或者operation字段为reformat/proofread
            If json("operation") IsNot Nothing Then
                Dim op = json("operation").ToString().ToLower()
                If op = "reformat" OrElse op = "proofread" Then
                    Return True
                End If
            End If
        Catch
        End Try
        Return False
    End Function

    ''' <summary>
    ''' 检测是否为旧版JSON命令格式
    ''' </summary>
    Public Shared Function IsLegacyJsonCommandFormat(jsonText As String) As Boolean
        Try
            Dim json = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
            ' 旧版格式特征：有command字段或commands数组
            If json("command") IsNot Nothing OrElse json("commands") IsNot Nothing Then
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    ''' <summary>
    ''' 检测格式类型
    ''' </summary>
    Public Shared Function DetectFormat(jsonText As String) As InstructionFormat
        If String.IsNullOrWhiteSpace(jsonText) Then Return InstructionFormat.None

        If IsDslFormat(jsonText) Then Return InstructionFormat.DslJson
        If IsLegacyJsonCommandFormat(jsonText) Then Return InstructionFormat.LegacyJsonCommand

        ' 尝试检测是否为校对JSON数组
        Try
            If jsonText.Trim().StartsWith("[") Then
                Return InstructionFormat.ProofreadJson
            End If
        Catch
        End Try

        Return InstructionFormat.None
    End Function

End Class
