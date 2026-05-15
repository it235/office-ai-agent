' ShareRibbon\Services\Proofread\ProofreadPromptBuilder.vb
' 校对Prompt构建器 - 构建AI校对Prompt，解析校对结果

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 校对Prompt构建器
''' </summary>
Public Class ProofreadPromptBuilder

    ''' <summary>
    ''' 构建全文校对Prompt
    ''' </summary>
    Public Shared Function BuildFullDocumentPrompt(paragraphs As List(Of String)) As String
        Dim sb As New StringBuilder()
        
        sb.AppendLine("你是Word文档智能校对助手。请仔细检查以下文档内容，识别所有需要修正的问题。")
        sb.AppendLine()
        sb.AppendLine("【校对范围】")
        sb.AppendLine("1. 错别字和拼写错误")
        sb.AppendLine("2. 词语使用错误（包括但不限于）：")
        sb.AppendLine("   - 的地得混用（的地得是最常见的词语错误）")
        sb.AppendLine("   - 他/她/它在表示指代时的混用")
        sb.AppendLine("   - 的在/得/地混用")
        sb.AppendLine("   - 其他常见用词错误")
        sb.AppendLine("3. 标点符号错误：")
        sb.AppendLine("   - 中英文标点混用（如中文句子里用了英文逗号）")
        sb.AppendLine("   - 标点缺失或多余")
        sb.AppendLine("   - 引号、括号不匹配")
        sb.AppendLine("4. 语法和语病问题")
        sb.AppendLine("5. 表达不通顺或容易引起歧义的地方")
        sb.AppendLine()
        
        sb.AppendLine("【文档内容】")
        For i = 0 To paragraphs.Count - 1
            Dim para = paragraphs(i)
            If Not String.IsNullOrWhiteSpace(para) Then
                sb.AppendLine($"[段落{i}] {para}")
            End If
        Next
        sb.AppendLine()
        
        sb.AppendLine("【输出要求】")
        sb.AppendLine("请以JSON数组格式返回校对结果：")
        sb.AppendLine("[")
        sb.AppendLine("  {")
        sb.AppendLine("    ""paragraphIndex"": 0,")
        sb.AppendLine("    ""original"": ""需要修正的原文片段（必须精确匹配）"",")
        sb.AppendLine("    ""suggestion"": ""修正后的文本（只写修正内容，不要加说明）"",")
        sb.AppendLine("    ""issueType"": ""WordUsageError"",")
        sb.AppendLine("    ""severity"": ""High"",")
        sb.AppendLine("    ""explanation"": ""简要说明修改原因""")
        sb.AppendLine("  }")
        sb.AppendLine("]")
        sb.AppendLine()
        
        sb.AppendLine("【issueType可选值】")
        sb.AppendLine("- spellingError: 拼写错误")
        sb.AppendLine("- wordUsageError: 用词错误")
        sb.AppendLine("- punctuationError: 标点错误")
        sb.AppendLine("- grammaticalError: 语法错误")
        sb.AppendLine("- expressionError: 表达问题")
        sb.AppendLine()
        
        sb.AppendLine("【severity可选值】")
        sb.AppendLine("- High: 必须修改（如错别字、严重语法错误）")
        sb.AppendLine("- Medium: 建议修改（如用词不当、轻微语病）")
        sb.AppendLine("- Low: 可选优化（如表达可以更精炼）")
        sb.AppendLine()
        
        sb.AppendLine("【注意事项】")
        sb.AppendLine("1. original必须精确匹配文档原文，包括标点和空格")
        sb.AppendLine("2. 同一段落有多处问题时，需要返回多个条目")
        sb.AppendLine("3. 只返回需要修改的内容，没问题的段落不要包含在结果中")
        sb.AppendLine("4. 如果文档没有需要修改的内容，请返回空数组：[]")
        sb.AppendLine("5. 请尽量全面地检查，不要遗漏明显的问题")
        
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 解析AI返回的校对结果
    ''' </summary>
    Public Shared Function ParseProofreadResponse(
        aiResponse As String,
        Optional paragraphs As List(Of String) = Nothing) As List(Of ProofreadIssue)
        
        Dim issues As New List(Of ProofreadIssue)()
        
        Try
            ' 清理响应，提取JSON
            Dim jsonContent = ExtractJson(aiResponse)
            If String.IsNullOrEmpty(jsonContent) Then Return issues
            
            ' 解析JSON数组
            Dim jsonArray = JArray.Parse(jsonContent)
            
            For Each item In jsonArray
                Dim issue As New ProofreadIssue()
                
                ' 解析基本字段
                If item("paragraphIndex") IsNot Nothing Then
                    issue.ParagraphIndex = CInt(item("paragraphIndex"))
                End If
                
                If item("original") IsNot Nothing Then
                    issue.Original = item("original").ToString()
                End If
                
                If item("suggestion") IsNot Nothing Then
                    issue.Suggestion = item("suggestion").ToString()
                End If
                
                If item("issueType") IsNot Nothing Then
                    issue.IssueType = ParseIssueType(item("issueType").ToString())
                End If
                
                If item("severity") IsNot Nothing Then
                    issue.Severity = ParseSeverity(item("severity").ToString())
                End If
                
                If item("explanation") IsNot Nothing Then
                    issue.Explanation = item("explanation").ToString()
                End If
                
                ' 生成唯一ID
                issue.Id = Guid.NewGuid().ToString()
                
                ' 验证数据有效性
                If Not String.IsNullOrEmpty(issue.Original) AndAlso 
                   Not String.IsNullOrEmpty(issue.Suggestion) AndAlso
                   issue.ParagraphIndex >= 0 Then
                    issues.Add(issue)
                End If
            Next
            
        Catch ex As Exception
            Debug.WriteLine($"[ProofreadPromptBuilder] 解析校对结果失败: {ex.Message}")
        End Try
        
        Return issues
    End Function

    ''' <summary>
    ''' 从AI响应中提取JSON内容
    ''' </summary>
    Private Shared Function ExtractJson(content As String) As String
        If String.IsNullOrEmpty(content) Then Return ""
        
        ' 尝试提取代码块中的JSON
        Dim jsonMatch = Regex.Match(content, "```(?:json)?\s*([\s\S]*?)\s*```", RegexOptions.IgnoreCase)
        If jsonMatch.Success Then
            Return jsonMatch.Groups(1).Value.Trim()
        End If
        
        ' 尝试直接解析
        If content.Trim().StartsWith("[") Then
            Return content.Trim()
        End If
        
        ' 尝试找到JSON数组的开始和结束
        Dim startIdx = content.IndexOf("[")
        Dim endIdx = content.LastIndexOf("]")
        If startIdx >= 0 AndAlso endIdx > startIdx Then
            Return content.Substring(startIdx, endIdx - startIdx + 1)
        End If
        
        Return ""
    End Function

    ''' <summary>
    ''' 解析问题类型
    ''' </summary>
    Private Shared Function ParseIssueType(typeStr As String) As IssueType
        Select Case typeStr.ToLower()
            Case "spellingerror", "spelling", "spell"
                Return IssueType.SpellingError
            Case "wordusageerror", "wordusage", "word"
                Return IssueType.WordUsageError
            Case "punctuationerror", "punctuation", "punct"
                Return IssueType.PunctuationError
            Case "grammaticalerror", "grammar", "grammatical"
                Return IssueType.GrammaticalError
            Case "expressionerror", "expression", "express"
                Return IssueType.ExpressionError
            Case "formaterror", "format"
                Return IssueType.FormatError
            Case Else
                Return IssueType.ExpressionError
        End Select
    End Function

    ''' <summary>
    ''' 解析严重程度
    ''' </summary>
    Private Shared Function ParseSeverity(severityStr As String) As IssueSeverity
        Select Case severityStr.ToLower()
            Case "high", "必须", "must", "error"
                Return IssueSeverity.High
            Case "medium", "建议", "should", "warning"
                Return IssueSeverity.Medium
            Case "low", "可选", "could", "info", "suggestion"
                Return IssueSeverity.Low
            Case Else
                Return IssueSeverity.Medium
        End Select
    End Function

End Class
