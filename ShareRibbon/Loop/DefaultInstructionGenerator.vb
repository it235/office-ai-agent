' ShareRibbon\Loop\DefaultInstructionGenerator.vb
' 默认指令生成器 - 生成AI提示词和修正请求

Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

''' <summary>
''' 默认指令生成器
''' </summary>
Public Class DefaultInstructionGenerator
    Implements IInstructionGenerator

    ''' <summary>
    ''' 根据执行上下文生成AI提示词
    ''' </summary>
    Public Async Function GenerateAsync(context As ExecutionContext, plan As PlanningResult) As Task(Of String) Implements IInstructionGenerator.GenerateAsync
        Try
            ' 构建DSL指令生成提示词
            Dim prompt = BuildDslPrompt(context, plan)

            ' 模拟AI调用（实际项目中应调用真实的AI接口）
            Dim aiResponse = Await CallAIGenerateAsync(prompt, context)

            Return aiResponse

        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' 根据错误信息生成修正请求
    ''' </summary>
    Public Async Function GenerateCorrectionAsync(
        originalResponse As String,
        errors As List(Of InstructionError)) As Task(Of String) Implements IInstructionGenerator.GenerateCorrectionAsync

        Try
            Dim prompt = BuildCorrectionPrompt(originalResponse, errors)
            Dim aiResponse = Await CallAIGenerateAsync(prompt, Nothing)
            Return aiResponse

        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' 构建DSL指令生成提示词
    ''' </summary>
    Private Function BuildDslPrompt(context As ExecutionContext, plan As PlanningResult) As String
        Dim sb As New System.Text.StringBuilder()

        ' 基础提示词
        sb.AppendLine("你是Office文档智能排版助手。用户要求你对文档进行排版操作。")
        sb.AppendLine()

        ' 文档信息
        sb.AppendLine("【当前文档信息】")
        sb.AppendLine($"- 文档类型: {If(context.OfficeContent IsNot Nothing, context.OfficeContent.DocumentType, "未知")}")
        sb.AppendLine($"- 处理范围: {If(context.RequiresSelection AndAlso context.SelectionInfo IsNot Nothing, "选中区域", "整个文档")}")
        sb.AppendLine($"- 段落数: {context.Paragraphs.Count}")

        If context.Paragraphs IsNot Nothing AndAlso context.Paragraphs.Count > 0 Then
            sb.AppendLine("- 选中段落预览:")
            For i = 0 To Math.Min(context.Paragraphs.Count - 1, 5)
                Dim preview = If(context.Paragraphs(i).Length > 50,
                    context.Paragraphs(i).Substring(0, 50) & "...",
                    context.Paragraphs(i))
                sb.AppendLine($"  [{i}] {preview}")
            Next
            If context.Paragraphs.Count > 5 Then
                sb.AppendLine($"  ... 还有 {context.Paragraphs.Count - 5} 个段落")
            End If
        End If

        sb.AppendLine()
        sb.AppendLine("【你的任务】")
        sb.AppendLine("请分析用户需求，生成一组结构化的操作指令（DSL）。")
        sb.AppendLine()

        ' 重要规则
        sb.AppendLine("【重要规则】")
        sb.AppendLine("1. 你只负责描述'期望的格式状态'，具体的Word DOM操作由系统代码执行")
        sb.AppendLine("2. 不要生成OpenXML，不要生成VBA代码")
        sb.AppendLine("3. 使用声明式指令，描述'应该是什么样子'而非'怎么操作'")
        sb.AppendLine("4. 每个指令必须包含id、op、target、params、expected字段")
        sb.AppendLine("5. target使用语义选择器（如'paragraph[role='heading1']'），不要使用绝对位置")
        sb.AppendLine("6. 如果用户需求不明确，返回空instructions数组并在metadata中说明")
        sb.AppendLine("7. 所有中文字符串值使用正确的引号包裹")
        sb.AppendLine("8. 不要在JSON末尾添加多余的逗号")

        ' 输出格式
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

        ' 语义选择器语法
        sb.AppendLine("【语义选择器语法】")
        sb.AppendLine("- paragraph[role='heading1'] - 第1级标题段落")
        sb.AppendLine("- paragraph[role='heading2'] - 第2级标题段落")
        sb.AppendLine("- paragraph[role='body'] - 正文段落")
        sb.AppendLine("- paragraph[role='title'] - 文档标题")
        sb.AppendLine("- paragraph:first - 第一个段落")
        sb.AppendLine("- paragraph:last - 最后一个段落")
        sb.AppendLine("- selection - 当前选区")

        ' 用户需求
        sb.AppendLine()
        sb.AppendLine("【用户需求】")
        sb.AppendLine(context.UserMessage)

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 构建修正提示词
    ''' </summary>
    Private Function BuildCorrectionPrompt(originalResponse As String, errors As List(Of InstructionError)) As String
        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("你之前返回的操作指令存在以下问题，请修正后重新返回：")
        sb.AppendLine()

        For Each [error] In errors
            sb.AppendLine($"- [{[error].Level}] {[error].Message}")
            If Not String.IsNullOrEmpty([error].Suggestion) Then
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

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 模拟AI调用（实际项目中应实现真实的AI接口）
    ''' </summary>
    Private Async Function CallAIGenerateAsync(prompt As String, context As ExecutionContext) As Task(Of String)
        ' 简化实现：返回示例DSL指令
        Await Task.Delay(1000)

        Dim exampleDsl = "{
  ""version"": ""2.0"",
  ""protocol"": ""office-dsl"",
  ""operation"": ""reformat"",
  ""target"": { ""scope"": ""document"", ""documentType"": ""official"" },
  ""instructions"": [
    {
      ""id"": ""inst-001"",
      ""op"": ""setParagraphStyle"",
      ""target"": { ""type"": ""semantic"", ""selector"": ""paragraph[role='title']"", ""index"": 0 },
      ""params"": { ""styleName"": ""标题 1"", ""font"": { ""name"": ""仿宋"", ""size"": 16 }, ""alignment"": ""center"" },
      ""expected"": { ""description"": ""标题居中仿宋16磅"", ""verifyBy"": ""styleNameAndFont"" },
      ""rollback"": { ""op"": ""setParagraphStyle"", ""params"": { ""styleName"": ""__original__"" } }
    },
    {
      ""id"": ""inst-002"",
      ""op"": ""setParagraphStyle"",
      ""target"": { ""type"": ""semantic"", ""selector"": ""paragraph[role='heading1']"", ""index"": 1 },
      ""params"": { ""styleName"": ""标题 1"", ""font"": { ""name"": ""黑体"", ""size"": 14 }, ""alignment"": ""left"" },
      ""expected"": { ""description"": ""一级标题左对齐黑体14磅"", ""verifyBy"": ""styleNameAndFont"" },
      ""rollback"": { ""op"": ""setParagraphStyle"", ""params"": { ""styleName"": ""__original__"" } }
    },
    {
      ""id"": ""inst-003"",
      ""op"": ""setParagraphStyle"",
      ""target"": { ""type"": ""semantic"", ""selector"": ""paragraph[role='body']"" },
      ""params"": { ""styleName"": ""正文"", ""font"": { ""name"": ""仿宋_GB2312"", ""size"": 16 }, ""alignment"": ""justify"", ""indent"": { ""firstLine"": 2 } },
      ""expected"": { ""description"": ""正文两端对齐仿宋16磅首行缩进2字符"", ""verifyBy"": ""styleNameAndFont"" },
      ""rollback"": { ""op"": ""setParagraphStyle"", ""params"": { ""styleName"": ""__original__"" } }
    }
  ],
  ""metadata"": { ""estimatedOperations"": 3, ""hasDestructiveOps"": false, ""requiresConfirmation"": false }
}"

        Return exampleDsl
    End Function
End Class
