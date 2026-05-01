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

End Class
