' ShareRibbon\Controls\Services\ChatFormatterAgent.vb
' Chat格式化代理 - 处理Chat中的排版对话消息、生成排版卡片HTML、解析自然语言排版指令

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 微调指令 - 从用户自然语言中解析出的排版微调操作
''' </summary>
Public Class RefinementCommand
    ''' <summary>目标区域: "title" / "body" / "page" / "heading" / "all"</summary>
    Public Property Target As String = ""
    ''' <summary>操作: "fontSize" / "alignment" / "color" / "spacing" / "fontFamily" / "indent"</summary>
    Public Property Action As String = ""
    ''' <summary>操作值: "+2pt" / "center" / "#FF0000" / "1.5"</summary>
    Public Property Value As String = ""
    ''' <summary>用户原始消息</summary>
    Public Property OriginalText As String = ""
End Class

''' <summary>
''' Chat格式化代理 - 在Chat对话中处理排版相关的用户交互
''' </summary>
Public Class ChatFormatterAgent

    Private ReadOnly _orchestrator As SmartFormattingOrchestrator
    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _escapeJs As Func(Of String, String)
    Private ReadOnly _textAnalyzer As Func(Of String, String, Task(Of String))

    ''' <summary>
    ''' 存储最后AI标注的段落结果（用于应用排版时）
    ''' </summary>
    Private _lastTaggedParagraphs As List(Of TaggedParagraph) = Nothing

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New(
        executeScript As Func(Of String, Task),
        escapeJs As Func(Of String, String),
        Optional textAnalyzer As Func(Of String, String, Task(Of String)) = Nothing,
        Optional orchestrator As SmartFormattingOrchestrator = Nothing)

        _executeScript = executeScript
        _escapeJs = escapeJs
        _textAnalyzer = textAnalyzer
        _orchestrator = If(orchestrator, New SmartFormattingOrchestrator())
    End Sub

    ''' <summary>访问编排器实例</summary>
    Public ReadOnly Property Orchestrator As SmartFormattingOrchestrator
        Get
            Return _orchestrator
        End Get
    End Property

    ''' <summary>
    ''' 处理Chat中的排版消息。
    ''' 根据消息内容自动判断是"首次排版请求"还是"微调指令"。
    ''' </summary>
    ''' <param name="userMessage">用户消息文本</param>
    ''' <param name="paragraphs">文档段落文本列表</param>
    ''' <param name="wordParagraphs">Word段落对象列表</param>
    ''' <param name="responseUuid">响应的UUID（用于推送HTML到前端）</param>
    ''' <returns>是否已处理（False表示非排版消息，应由其他处理器处理）</returns>
    Public Async Function HandleFormattingMessage(
        userMessage As String,
        paragraphs As List(Of String),
        wordParagraphs As List(Of Object),
        responseUuid As String) As Task(Of Boolean)

        ' 判断是否为排版相关消息
        If Not IsFormattingRelated(userMessage) Then
            Return False
        End If

        Try
            ' 始终通过ChatReformatAsync解析用户意图，避免忽略用户指令
            Dim lastPlan = _orchestrator.RefinementContext.LastPlan
            Dim plan = Await _orchestrator.ChatReformatAsync(userMessage, paragraphs, wordParagraphs)

            ' 有活动上下文且非全新请求时显示微调对比卡片，否则显示完整预览
            Dim isRefinement = _orchestrator.HasActiveContext() AndAlso
                               Not IsNewFormattingRequest(userMessage) AndAlso
                               lastPlan IsNot Nothing

            If isRefinement Then
                Dim html = GenerateRefinementCardHtml(lastPlan, plan)
                If String.IsNullOrEmpty(responseUuid) Then
                    responseUuid = Guid.NewGuid().ToString()
                End If
                Dim jsonPayload As New JObject()
                jsonPayload("uuid") = responseUuid
                jsonPayload("html") = html
                Await _executeScript($"appendFormattingCard({jsonPayload.ToString(Newtonsoft.Json.Formatting.None)});")
            Else
                Dim html = GenerateFormattingCardHtml(plan)
                If String.IsNullOrEmpty(responseUuid) Then
                    responseUuid = Guid.NewGuid().ToString()
                End If
                Dim jsonPayload As New JObject()
                jsonPayload("uuid") = responseUuid
                jsonPayload("html") = html
                Await _executeScript($"appendFormattingCard({jsonPayload.ToString(Newtonsoft.Json.Formatting.None)});")
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] 处理排版消息失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 生成排版建议卡片HTML
    ''' </summary>
    Public Function GenerateFormattingCardHtml(plan As ReformatPreviewPlan) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F4CB;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版建议</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        ' 文档类型与标准
        If plan.DetectedType <> DocumentType.Unknown Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文档类型: <strong>{plan.DocumentTypeName}</strong> (置信度{Math.Round(plan.TypeConfidence * 100)}%)</div>")
        End If
        If Not String.IsNullOrEmpty(plan.StandardName) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">推荐标准: <strong>{plan.StandardName}</strong></div>")
        End If

        ' 变更列表 — 按NewTag分组显示
        sb.AppendLine("    <div class=""formatting-changes"">")
        sb.AppendLine("      <div class=""formatting-changes-title"">即将修改:</div>")

        ' 按NewTag分组
        Dim grouped = plan.Changes.GroupBy(Function(c) If(String.IsNullOrEmpty(c.NewTag), "__pending__", c.NewTag)).ToList()
        For Each group In grouped
            Dim tagName = If(group.Key = "__pending__", "AI待标注", group.Key)
            Dim count = group.Count()
            Dim sampleDesc = group.FirstOrDefault()?.ChangeDescription
            sb.AppendLine($"      <div class=""formatting-change-item"">")
            sb.AppendLine($"        <span class=""formatting-change-section"">{System.Web.HttpUtility.HtmlEncode(tagName)}</span>")
            sb.AppendLine($"        <span class=""formatting-change-count"">({count}处)</span>")
            If Not String.IsNullOrEmpty(sampleDesc) Then
                sb.AppendLine($"        <span class=""formatting-change-desc"">: {System.Web.HttpUtility.HtmlEncode(sampleDesc)}</span>")
            End If
            sb.AppendLine($"      </div>")
        Next

        sb.AppendLine($"      <div class=""formatting-change-summary"">合计: {plan.TotalChanges}处段落, {grouped.Count}个样式区</div>")
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">应用排版</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-secondary"" onclick=""previewReformat();"">预览对比</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-outline"" onclick=""alternateReformat();"">换一种</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-ghost"" onclick=""startRefinement();"">微调</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成微调对比卡片HTML（显示变更前后对比）
    ''' </summary>
    Public Function GenerateRefinementCardHtml(
        before As ReformatPreviewPlan,
        after As ReformatPreviewPlan) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card formatting-card-refinement"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F504;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版已微调</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        ' 变更对比 — 按ParagraphIndex匹配
        sb.AppendLine("    <div class=""formatting-diff"">")
        For Each c In after.Changes
            Dim beforeChange = before.Changes.FirstOrDefault(Function(b) b.ParagraphIndex = c.ParagraphIndex)
            If beforeChange IsNot Nothing Then
                Dim oldDesc = If(String.IsNullOrEmpty(beforeChange.ChangeDescription), beforeChange.NewTag, beforeChange.ChangeDescription)
                Dim newDesc = If(String.IsNullOrEmpty(c.ChangeDescription), c.NewTag, c.ChangeDescription)
                If oldDesc <> newDesc Then
                    sb.AppendLine("      <div class=""formatting-diff-item"">")
                    sb.AppendLine($"        <span class=""formatting-diff-section"">{System.Web.HttpUtility.HtmlEncode(c.ParagraphPreview)}:</span>")
                    sb.AppendLine($"        <span class=""formatting-diff-old"">{System.Web.HttpUtility.HtmlEncode(oldDesc)}</span>")
                    sb.AppendLine($"        <span class=""formatting-diff-arrow"">&rarr;</span>")
                    sb.AppendLine($"        <span class=""formatting-diff-new"">{System.Web.HttpUtility.HtmlEncode(newDesc)}</span>")
                    sb.AppendLine("      </div>")
                End If
            End If
        Next
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">应用排版</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-ghost"" onclick=""startRefinement();"">继续微调</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 解析用户消息中的微调指令
    ''' </summary>
    Public Shared Function ParseRefinementCommand(userMessage As String) As RefinementCommand
        Dim cmd As New RefinementCommand()
        cmd.OriginalText = userMessage

        If String.IsNullOrWhiteSpace(userMessage) Then Return cmd

        ' 简单文本解析微调指令
        Dim msg = userMessage.ToLower().Trim()
        If msg.Contains("大") OrElse msg.Contains("小") Then
            cmd.Action = "fontSize"
            cmd.Value = If(msg.Contains("大"), "+1pt", "-1pt")
        ElseIf msg.Contains("行距") Then
            cmd.Action = "spacing"
            cmd.Value = "1.5"
        ElseIf msg.Contains("红") OrElse msg.Contains("蓝") OrElse msg.Contains("颜色") Then
            cmd.Action = "color"
        ElseIf msg.Contains("居中") OrElse msg.Contains("对齐") Then
            cmd.Action = "alignment"
            cmd.Value = "center"
        End If

        cmd.Target = "all"
        Return cmd
    End Function

    ''' <summary>
    ''' 判断消息是否与排版相关
    ''' </summary>
    Public Shared Function IsFormattingRelated(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False

        Dim keywords As String() = {
            "排版", "格式", "样式", "字体", "字号", "行距", "对齐",
            "缩进", "页边距", "居中", "加粗", "红色", "标题", "正文",
            "仿宋", "宋体", "黑体", "楷体", "微软雅黑", "小标宋",
            "公文", "国标", "标准", "模板", "美化", "整理", "规范",
            "GB/T", "gbt", "克隆", "照这个", "参照", "按照"
        }

        Return keywords.Any(Function(k) message.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
    End Function

    ''' <summary>
    ''' 判断是否为全新的排版请求（而非微调）
    ''' 如果是全新的格式化请求，会重新做完整分析
    ''' </summary>
    Private Shared Function IsNewFormattingRequest(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False

        Dim newRequestKeywords As String() = {
            "重新", "再来", "换一种", "重新排", "重新排版",
            "换一个", "用另一个", "换成", "改用", "不要这个"
        }

        Return newRequestKeywords.Any(Function(k) message.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
    End Function

    ''' <summary>
    ''' 获取最后AI标注的段落结果（用于应用排版时）
    ''' </summary>
    Public Function GetLastTaggedParagraphs() As List(Of TaggedParagraph)
        Return _lastTaggedParagraphs
    End Function

    ''' <summary>
    ''' 使用AI进行语义标注
    ''' 调用AI分析文档内容，返回每个段落对应的语义标签
    ''' </summary>
    ''' <param name="paragraphs">文档段落文本列表</param>
    ''' <param name="mapping">当前排版标准的语义样式映射</param>
    ''' <param name="paragraphStyles">段落样式名称列表（可选，用于增强AI判断）</param>
    ''' <param name="documentTypeContext">文档类型上下文描述（可选）</param>
    ''' <param name="detectedHeadings">已检测到的标题结构（可选）</param>
    Public Async Function PerformAISemanticTaggingAsync(
        paragraphs As List(Of String),
        mapping As SemanticStyleMapping,
        Optional paragraphStyles As List(Of String) = Nothing,
        Optional documentTypeContext As String = Nothing,
        Optional detectedHeadings As String = Nothing) As Task(Of List(Of TaggedParagraph))

        _lastTaggedParagraphs = Nothing

        ' 如果没有AI分析器，回退到基于规则的简单标注（全部body.normal）
        If _textAnalyzer Is Nothing Then
            Debug.WriteLine("[ChatFormatterAgent] 没有配置AI分析器，使用默认标注")
            Dim fallback As New List(Of TaggedParagraph)()
            For i = 0 To paragraphs.Count - 1
                fallback.Add(New TaggedParagraph(i, "body.normal"))
            Next
            _lastTaggedParagraphs = fallback
            Return fallback
        End If

        Try
            ' 构建AI提示词
            Dim originalIndices As New List(Of Integer)()
            For i = 0 To paragraphs.Count - 1
                originalIndices.Add(i)
            Next

            Dim prompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(
                mapping,
                paragraphs,
                paragraphStyles,
                originalIndices,
                detectedHeadings,
                documentTypeContext)

            ' 调用AI获取标注结果
            Debug.WriteLine("[ChatFormatterAgent] 正在调用AI进行语义标注...")
            Dim aiResponse = Await _textAnalyzer("semantic_tagging", prompt)

            If String.IsNullOrWhiteSpace(aiResponse) Then
                Debug.WriteLine("[ChatFormatterAgent] AI返回为空，使用默认标注")
                Return Await GetDefaultTaggingAsync(paragraphs)
            End If

            ' 解析AI响应
            Dim taggedParagraphs = ParseAITagResponse(aiResponse, paragraphs.Count)
            _lastTaggedParagraphs = taggedParagraphs

            Debug.WriteLine($"[ChatFormatterAgent] AI标注完成: {taggedParagraphs.Count}个段落")
            Return taggedParagraphs

        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] AI标注失败: {ex.Message}")
        End Try

        ' 如果解析失败或结果为空，返回默认标注
        Return Await GetDefaultTaggingAsync(paragraphs)
    End Function

    ''' <summary>
    ''' 获取默认标注（全部标记为body.normal）
    ''' </summary>
    Private Async Function GetDefaultTaggingAsync(paragraphs As List(Of String)) As Task(Of List(Of TaggedParagraph))
        Dim result As New List(Of TaggedParagraph)()
        For i = 0 To paragraphs.Count - 1
            result.Add(New TaggedParagraph(i, "body.normal"))
        Next
        _lastTaggedParagraphs = result
        Return result
    End Function

    ''' <summary>
    ''' 解析AI标注响应
    ''' </summary>
    Private Function ParseAITagResponse(response As String, paragraphCount As Integer) As List(Of TaggedParagraph)
        Dim result As New List(Of TaggedParagraph)()

        Try
            ' 清理响应文本，移除可能的markdown代码块标记
            Dim cleanResponse = response.Trim()
            If cleanResponse.StartsWith("```json") Then
                cleanResponse = cleanResponse.Substring(7)
            ElseIf cleanResponse.StartsWith("```") Then
                cleanResponse = cleanResponse.Substring(3)
            End If
            If cleanResponse.EndsWith("```") Then
                cleanResponse = cleanResponse.Substring(0, cleanResponse.Length - 3)
            End If
            cleanResponse = cleanResponse.Trim()

            ' 尝试解析JSON数组
            Dim taggerd As List(Of TaggedParagraph) = Nothing
            Try
                taggerd = JsonConvert.DeserializeObject(Of List(Of TaggedParagraph))(cleanResponse)
            Catch ex As Exception
                ' JSON解析失败，尝试正则提取
                Debug.WriteLine($"[ChatFormatterAgent] JSON解析失败: {ex.Message}")
                taggerd = ParseTagResponseWithRegex(cleanResponse, paragraphCount)
            End Try

            If taggerd IsNot Nothing AndAlso taggerd.Count > 0 Then
                result.AddRange(taggerd)
            End If

        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] 解析标注响应失败: {ex.Message}")
        End Try

        ' 如果解析失败或结果为空，返回默认标注
        If result.Count = 0 Then
            For i = 0 To paragraphCount - 1
                result.Add(New TaggedParagraph(i, "body.normal"))
            Next
        End If

        Return result
    End Function

    ''' <summary>
    ''' 使用正则表达式解析标注响应（当JSON解析失败时）
    ''' </summary>
    Private Function ParseTagResponseWithRegex(response As String, paragraphCount As Integer) As List(Of TaggedParagraph)
        Dim result As New List(Of TaggedParagraph)()

        ' 匹配 {paraIndex:数字, tag:"标签"} 模式
        Dim pattern = """paraIndex""\s*:\s*(\d+)\s*,\s*""tag""\s*:\s*""([^""]+)"""
        Dim matches = System.Text.RegularExpressions.Regex.Matches(response, pattern)

        For Each match In matches
            Dim paraIndex = Integer.Parse(match.Groups(1).Value)
            Dim tag = match.Groups(2).Value
            If paraIndex >= 0 AndAlso paraIndex < paragraphCount Then
                result.Add(New TaggedParagraph(paraIndex, tag))
            End If
        Next

        ' 如果正则也没有匹配到，返回空列表（将使用默认标注）
        Return result
    End Function

End Class
