' ShareRibbon\Controls\Services\ChatFormatterAgent.vb
' Chat格式化代理 - 处理Chat中的排版对话消息并生成排版卡片HTML。

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Web
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Chat格式化代理 - 在Chat对话中处理排版相关的用户交互
''' Phase 2改进：
''' 1. AI语义标注为必经步骤（无AI时明确告知用户）
''' 2. 先分析后确认：先展示AI分析结果，用户确认后才应用
''' </summary>
Public Class ChatFormatterAgent

    Private ReadOnly _orchestrator As SmartFormattingOrchestrator
    Private ReadOnly _executeScript As Func(Of String, Task)
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
        Optional textAnalyzer As Func(Of String, String, Task(Of String)) = Nothing,
        Optional orchestrator As SmartFormattingOrchestrator = Nothing)

        _executeScript = executeScript
        _textAnalyzer = textAnalyzer
        _orchestrator = If(orchestrator, New SmartFormattingOrchestrator())
    End Sub

    ''' <summary>访问编排器实例</summary>
    Public ReadOnly Property Orchestrator As SmartFormattingOrchestrator
        Get
            Return _orchestrator
        End Get
    End Property

    Public Async Function RecognizeReformatIntentAsync(userMessage As String,
                                                       paragraphs As List(Of String),
                                                       Optional paragraphStyles As List(Of String) = Nothing,
                                                       Optional paragraphFontSizes As List(Of Single) = Nothing,
                                                       Optional paragraphIsBold As List(Of Boolean) = Nothing) As Task(Of ReformatIntentRecognitionResult)
        Dim analysis As DocumentAnalysisResult = Nothing
        Try
            Dim analyzer As New DocumentAnalyzer()
            If paragraphStyles IsNot Nothing AndAlso paragraphFontSizes IsNot Nothing AndAlso paragraphIsBold IsNot Nothing Then
                analysis = analyzer.Analyze(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
            Else
                analysis = analyzer.Analyze(paragraphs)
            End If
        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] 文档上下文分析失败，使用空上下文识别意图: {ex.Message}")
            analysis = New DocumentAnalysisResult()
        End Try

        Dim recognizer As New ReformatIntentRecognizer(
            _textAnalyzer,
            Function(message, context) _orchestrator.ParseUserIntent(message, context))

        Return Await recognizer.RecognizeAsync(userMessage, analysis, paragraphs)
    End Function

    ''' <summary>
    ''' 处理Chat中的排版消息。
    ''' 根据消息内容自动判断是"首次排版请求"还是"微调指令"。
    ''' Phase 2改进：始终走AI标注流程，先展示分析结果再确认。
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
            ' 通过ChatReformatAsync解析用户意图，获取排版方案
            Dim lastPlan = _orchestrator.RefinementContext.LastPlan
            Dim plan = Await _orchestrator.ChatReformatAsync(userMessage, paragraphs, wordParagraphs)

            If plan Is Nothing OrElse plan.Changes.Count = 0 Then
                ' 没有生成方案，展示提示
                Dim html = GenerateNoPlanCardHtml(userMessage)
                Await PushCardHtml(html, responseUuid)
                Return True
            End If

            ' 判断是否为微调请求
            Dim isRefinement = _orchestrator.HasActiveContext() AndAlso
                               Not IsNewFormattingRequest(userMessage) AndAlso
                               lastPlan IsNot Nothing

            If isRefinement Then
                ' 微调模式：AI标注结果已存在，展示微调对比卡片
                Dim html = GenerateRefinementCardHtml(lastPlan, plan)
                Await PushCardHtml(html, responseUuid)
            Else
                ' 首次排版：执行AI语义标注（必经步骤）
                If _textAnalyzer IsNot Nothing Then
                    ' AI可用：执行语义标注，展示带分析结果的卡片
                    Dim taggedParagraphs = Await PerformAISemanticTaggingAsync(
                        paragraphs,
                        plan.SemanticMapping,
                        documentTypeContext:=plan.StandardName)

                    ' 将AI标注结果合并到plan中
                    UpdatePlanWithAITags(plan, taggedParagraphs)

                    Dim html = GenerateAnalysisCardHtml(plan, taggedParagraphs)
                    Await PushCardHtml(html, responseUuid)
                Else
                    ' AI不可用：展示警告卡片，说明效果有限
                    Dim html = GenerateNoAICardHtml(plan)
                    Await PushCardHtml(html, responseUuid)
                End If
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] 处理排版消息失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 将AI标注结果更新到预览方案中
    ''' </summary>
    Private Sub UpdatePlanWithAITags(plan As ReformatPreviewPlan, taggedParagraphs As List(Of TaggedParagraph))
        If plan Is Nothing OrElse taggedParagraphs Is Nothing Then Return

        ' 更新plan.Changes中的NewTag为AI标注结果
        For Each change In plan.Changes
            Dim tagged = taggedParagraphs.FirstOrDefault(Function(t) t.ParaIndex = change.ParagraphIndex)
            If tagged IsNot Nothing AndAlso Not String.IsNullOrEmpty(tagged.TagId) Then
                change.NewTag = tagged.TagId
            End If
        Next
    End Sub

    ''' <summary>
    ''' 推送卡片HTML到前端
    ''' </summary>
    Private Async Function PushCardHtml(html As String, responseUuid As String) As Task
        Dim uuid = If(responseUuid, Guid.NewGuid().ToString())
        Dim jsonPayload As New JObject()
        jsonPayload("uuid") = uuid
        jsonPayload("html") = html
        Await _executeScript($"appendFormattingCard({UtilsService.SerializeForScript(jsonPayload)});")
    End Function

    ''' <summary>
    ''' 卡片HTML里的文本一律转义：模型输出的名称/类型不能当成标记写进 innerHTML
    ''' </summary>
    Private Shared Function Esc(value As String) As String
        Return System.Web.HttpUtility.HtmlEncode(If(value, String.Empty))
    End Function

    Private Shared Function ShouldShowDocumentType(plan As ReformatPreviewPlan) As Boolean
        Return plan IsNot Nothing AndAlso
               plan.DetectedType <> DocumentType.Unknown AndAlso
               plan.TypeConfidence >= 0.35
    End Function

    Private Shared Function GetDocumentTypeLabel(plan As ReformatPreviewPlan) As String
        If plan Is Nothing Then Return ""
        If Not String.IsNullOrWhiteSpace(plan.DocumentTypeName) Then Return plan.DocumentTypeName
        Return plan.DetectedType.ToString()
    End Function

    ''' <summary>
    ''' 生成AI分析结果卡片HTML（Phase 2新增：先分析后确认）
    ''' </summary>
    Public Function GenerateAnalysisCardHtml(plan As ReformatPreviewPlan, taggedParagraphs As List(Of TaggedParagraph)) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card formatting-card-analysis"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F50D;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">AI分析完成</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        ' 文档类型与标准
        If Not String.IsNullOrWhiteSpace(plan.ScopeSummary) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">作用范围: <strong>{System.Web.HttpUtility.HtmlEncode(plan.ScopeSummary)}</strong></div>")
        End If
        If plan.TextParagraphCount > 0 AndAlso plan.TextParagraphCount <> plan.TotalParagraphs Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文本段落: <strong>{plan.TextParagraphCount}</strong> / 总段落 {plan.TotalParagraphs}</div>")
        End If
        If ShouldShowDocumentType(plan) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文档类型: <strong>{Esc(GetDocumentTypeLabel(plan))}</strong> (置信度{Math.Round(plan.TypeConfidence * 100)}%)</div>")
        End If
        If Not String.IsNullOrEmpty(plan.StandardName) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">排版标准: <strong>{Esc(plan.StandardName)}</strong></div>")
        End If

        ' AI标注摘要
        sb.AppendLine("    <div class=""formatting-changes"">")
        sb.AppendLine("      <div class=""formatting-changes-title"">AI识别的文档结构:</div>")

        ' 按tagId分组
        Dim grouped = taggedParagraphs.GroupBy(Function(t) t.TagId).OrderByDescending(Function(g) g.Count()).ToList()
        For Each group In grouped
            Dim tagName = group.Key
            Dim count = group.Count()
            ' 取一条reason作为示例
            Dim sampleReason = group.FirstOrDefault()?.Reason
            Dim displayName = GetTagDisplayName(tagName, plan.SemanticMapping)

            sb.AppendLine($"      <div class=""formatting-change-item"">")
            sb.AppendLine($"        <span class=""formatting-change-section"">{System.Web.HttpUtility.HtmlEncode(displayName)}</span>")
            sb.AppendLine($"        <span class=""formatting-change-count"">({count}处)</span>")
            If Not String.IsNullOrEmpty(sampleReason) Then
                sb.AppendLine($"        <span class=""formatting-change-reason"">- {System.Web.HttpUtility.HtmlEncode(sampleReason)}</span>")
            End If
            sb.AppendLine($"      </div>")
        Next

        sb.AppendLine($"      <div class=""formatting-change-summary"">合计: {taggedParagraphs.Count}个段落, {grouped.Count}种样式</div>")
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">确认应用</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-secondary"" onclick=""previewReformat();"">预览对比</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-outline"" onclick=""alternateReformat();"">换一种</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-ghost"" onclick=""startRefinement();"">微调</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成无AI时的卡片HTML（告知用户AI不可用，效果有限）
    ''' </summary>
    Public Function GenerateNoAICardHtml(plan As ReformatPreviewPlan) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card formatting-card-warning"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x26A0;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版建议（基础模式）</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")

        sb.AppendLine("    <div class=""formatting-info-row formatting-info-warning"">AI分析器未启用，将使用基于规则的分析，精度有限。建议配置AI模型以获得更精准的排版效果。</div>")

        ' 文档类型与标准
        If ShouldShowDocumentType(plan) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文档类型: <strong>{Esc(GetDocumentTypeLabel(plan))}</strong></div>")
        End If
        If Not String.IsNullOrEmpty(plan.StandardName) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">排版标准: <strong>{Esc(plan.StandardName)}</strong></div>")
        End If

        ' 变更列表
        sb.AppendLine("    <div class=""formatting-changes"">")
        sb.AppendLine("      <div class=""formatting-changes-title"">即将修改:</div>")
        Dim grouped = plan.Changes.GroupBy(Function(c) If(String.IsNullOrEmpty(c.NewTag), "__pending__", c.NewTag)).ToList()
        For Each group In grouped
            Dim tagName = If(group.Key = "__pending__", "待标注", group.Key)
            Dim count = group.Count()
            sb.AppendLine($"      <div class=""formatting-change-item"">")
            sb.AppendLine($"        <span class=""formatting-change-section"">{System.Web.HttpUtility.HtmlEncode(tagName)}</span>")
            sb.AppendLine($"        <span class=""formatting-change-count"">({count}处)</span>")
            sb.AppendLine($"      </div>")
        Next
        sb.AppendLine($"      <div class=""formatting-change-summary"">合计: {plan.TotalChanges}处段落</div>")
        sb.AppendLine("    </div>")

        ' 操作按钮
        sb.AppendLine("    <div class=""formatting-card-actions"">")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-primary"" onclick=""applyReformat();"">应用排版</button>")
        sb.AppendLine("      <button class=""formatting-btn formatting-btn-outline"" onclick=""alternateReformat();"">换一种</button>")
        sb.AppendLine("    </div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成无方案时的卡片HTML
    ''' </summary>
    Private Function GenerateNoPlanCardHtml(userMessage As String) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<div class=""formatting-card"">")
        sb.AppendLine("  <div class=""formatting-card-header"">")
        sb.AppendLine("    <span class=""formatting-card-icon"">&#x1F4CB;</span>")
        sb.AppendLine("    <span class=""formatting-card-title"">排版建议</span>")
        sb.AppendLine("  </div>")
        sb.AppendLine("  <div class=""formatting-card-body"">")
        sb.AppendLine("    <div class=""formatting-info-row"">未能生成排版方案。请尝试更明确的指令，如""按公文排版""或""整理格式""。</div>")
        sb.AppendLine("  </div>")
        sb.AppendLine("</div>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成排版建议卡片HTML（兼容旧版调用）
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
        If ShouldShowDocumentType(plan) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">文档类型: <strong>{Esc(GetDocumentTypeLabel(plan))}</strong> (置信度{Math.Round(plan.TypeConfidence * 100)}%)</div>")
        End If
        If Not String.IsNullOrEmpty(plan.StandardName) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">推荐标准: <strong>{Esc(plan.StandardName)}</strong></div>")
        End If
        If Not String.IsNullOrWhiteSpace(plan.StandardDescription) Then
            sb.AppendLine($"    <div class=""formatting-info-row"">方案说明: <strong>{System.Web.HttpUtility.HtmlEncode(plan.StandardDescription)}</strong></div>")
        End If

        ' 变更列表 — 按NewTag分组显示，面向用户展示语义名称而不是技术标签。
        sb.AppendLine("    <div class=""formatting-changes"">")
        sb.AppendLine("      <div class=""formatting-changes-title"">即将修改:</div>")

        ' 按NewTag分组
        Dim grouped = plan.Changes.GroupBy(Function(c) If(String.IsNullOrEmpty(c.NewTag), "__pending__", c.NewTag)).ToList()
        If grouped.Count = 0 Then
            sb.AppendLine("      <div class=""formatting-change-item"">")
            sb.AppendLine("        <span class=""formatting-change-section"">已完成文档结构分析</span>")
            sb.AppendLine("        <span class=""formatting-change-desc"">: 暂未发现需要立即调整的样式区，可继续微调或换一种方案。</span>")
            sb.AppendLine("      </div>")
        Else
            For Each group In grouped
                Dim tagName = If(group.Key = "__pending__", "AI待标注", GetTagDisplayName(group.Key, plan.SemanticMapping))
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
        End If

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
    ''' 获取标签的显示名称（优先从mapping中获取，否则返回tagId）
    ''' </summary>
    Private Function GetTagDisplayName(tagId As String, mapping As SemanticStyleMapping) As String
        If mapping IsNot Nothing Then
            Dim tag = mapping.FindTag(tagId)
            If tag IsNot Nothing AndAlso Not String.IsNullOrEmpty(tag.DisplayName) Then
                Return tag.DisplayName
            End If
        End If

        ' 从注册表获取默认名称
        Dim layer2Tags = SemanticTagRegistry.GetCommonLayer2Tags()
        If layer2Tags.ContainsKey(tagId) Then Return layer2Tags(tagId)

        Dim layer1Tags = SemanticTagRegistry.GetLayer1TagDescriptions()
        Dim parentTag = SemanticTagRegistry.GetParentTag(tagId)
        If layer1Tags.ContainsKey(parentTag) Then Return $"{layer1Tags(parentTag)}.{tagId}"

        Return tagId
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
    ''' 使用AI进行语义标注（必经步骤）
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

        ' AI不可用时返回规则推断结果（不再静默回退为body.normal）
        If _textAnalyzer Is Nothing Then
            Debug.WriteLine("[ChatFormatterAgent] 没有配置AI分析器，使用规则推断标注")
            Dim fallback As New List(Of TaggedParagraph)()
            For i = 0 To paragraphs.Count - 1
                ' 使用InferDefaultTag进行基于规则的推断（而非全部body.normal）
                Dim inferredTag = SmartFormattingOrchestrator.InferDefaultTagPublic(i, paragraphs.Count, paragraphs(i), Nothing)
                fallback.Add(New TaggedParagraph(i, inferredTag, "规则推断（无AI）"))
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
                Debug.WriteLine("[ChatFormatterAgent] AI返回为空，使用规则推断标注")
                Return Await GetRuleBasedTaggingAsync(paragraphs, mapping)
            End If

            ' 解析AI响应
            Dim taggedParagraphs = ParseAITagResponse(aiResponse, paragraphs.Count)
            _lastTaggedParagraphs = taggedParagraphs

            Debug.WriteLine($"[ChatFormatterAgent] AI标注完成: {taggedParagraphs.Count}个段落")
            Return taggedParagraphs

        Catch ex As Exception
            Debug.WriteLine($"[ChatFormatterAgent] AI标注失败: {ex.Message}")
        End Try

        ' 如果解析失败或结果为空，返回规则推断标注
        Return Await GetRuleBasedTaggingAsync(paragraphs, mapping)
    End Function

    ''' <summary>
    ''' 基于规则的标注（AI不可用时的替代方案，使用InferDefaultTag而非全部body.normal）
    ''' </summary>
    Private Function GetRuleBasedTaggingAsync(paragraphs As List(Of String), mapping As SemanticStyleMapping) As Task(Of List(Of TaggedParagraph))
        Dim result As New List(Of TaggedParagraph)()
        For i = 0 To paragraphs.Count - 1
            Dim inferredTag = SmartFormattingOrchestrator.InferDefaultTagPublic(i, paragraphs.Count, paragraphs(i), Nothing)
            result.Add(New TaggedParagraph(i, inferredTag, "规则推断"))
        Next
        _lastTaggedParagraphs = result
        Return Task.FromResult(result)
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

        ' 如果解析失败或结果为空，返回空列表（调用方会降级到规则推断）
        If result.Count = 0 Then
            For i = 0 To paragraphCount - 1
                result.Add(New TaggedParagraph(i, "body.normal", "解析失败降级"))
            Next
        End If

        Return result
    End Function

    ''' <summary>
    ''' 使用正则表达式解析标注响应（当JSON解析失败时）
    ''' </summary>
    Private Function ParseTagResponseWithRegex(response As String, paragraphCount As Integer) As List(Of TaggedParagraph)
        Dim result As New List(Of TaggedParagraph)()

        ' 匹配 {paraIndex:数字, tag:"标签", reason:"理由"(可选)} 模式
        Dim pattern = """paraIndex""\s*:\s*(\d+)\s*,\s*""tag""\s*:\s*""([^""]+)""(?:\s*,\s*""reason""\s*:\s*""([^""]*)"")?"
        Dim matches = System.Text.RegularExpressions.Regex.Matches(response, pattern)

        For Each match In matches
            Dim paraIndex = Integer.Parse(match.Groups(1).Value)
            Dim tag = match.Groups(2).Value
            Dim reason = If(match.Groups(3).Success, match.Groups(3).Value, "")
            If paraIndex >= 0 AndAlso paraIndex < paragraphCount Then
                result.Add(New TaggedParagraph(paraIndex, tag, reason))
            End If
        Next

        ' 如果正则也没有匹配到，返回空列表（将使用规则推断）
        Return result
    End Function

End Class
