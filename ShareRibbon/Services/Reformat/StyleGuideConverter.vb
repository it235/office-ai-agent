' ShareRibbon\Services\Reformat\StyleGuideConverter.vb
' 将文本规范通过AI转换为SemanticStyleMapping

Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 规范转换器 - 将文本格式规范（markdown/txt）转换为SemanticStyleMapping
''' 需要AI参与解析规范文本中的格式要求
''' </summary>
Public Class StyleGuideConverter

    ''' <summary>
    ''' 构建用于AI转换规范的系统提示词
    ''' </summary>
    ''' <param name="guideContent">规范原文内容</param>
    Public Shared Function BuildConversionPrompt(guideContent As String) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("你是文档格式规范解析专家。请从以下格式规范文本中提取格式参数，转换为结构化JSON。")
        sb.AppendLine()

        ' 标签体系说明
        sb.AppendLine("【可用语义标签ID】")
        sb.AppendLine("- title.1: 一级标题")
        sb.AppendLine("- title.2: 二级标题")
        sb.AppendLine("- title.3: 三级标题")
        sb.AppendLine("- body.normal: 正文")
        sb.AppendLine("- body.emphasis: 强调段落")
        sb.AppendLine("- list.ordered: 有序列表")
        sb.AppendLine("- list.unordered: 无序列表")
        sb.AppendLine("- quote: 引用")
        sb.AppendLine("- caption: 图表题注")
        sb.AppendLine()

        ' 输出格式要求
        sb.AppendLine("【输出格式要求】")
        sb.AppendLine("返回纯JSON对象，不要包含markdown代码块标记。格式如下：")
        sb.AppendLine("{")
        sb.AppendLine("  ""semanticTags"": [")
        sb.AppendLine("    {")
        sb.AppendLine("      ""tagId"": ""title.1"",")
        sb.AppendLine("      ""displayName"": ""一级标题"",")
        sb.AppendLine("      ""matchHint"": ""包含'第X章'开头"",")
        sb.AppendLine("      ""font"": {""fontNameCN"": ""黑体"", ""fontNameEN"": ""Arial"", ""fontSize"": 22, ""bold"": true},")
        sb.AppendLine("      ""paragraph"": {""alignment"": ""center"", ""firstLineIndent"": 0, ""lineSpacing"": 1.5, ""spaceBefore"": 1, ""spaceAfter"": 0.5},")
        sb.AppendLine("      ""color"": {""fontColor"": ""#000000""}")
        sb.AppendLine("    }")
        sb.AppendLine("  ],")
        sb.AppendLine("  ""pageConfig"": {")
        sb.AppendLine("    ""margins"": {""top"": 2.54, ""bottom"": 2.54, ""left"": 3.18, ""right"": 3.18}")
        sb.AppendLine("  }")
        sb.AppendLine("}")
        sb.AppendLine()

        ' 规范原文
        sb.AppendLine("【格式规范原文】")
        sb.AppendLine("---BEGIN STYLE GUIDE---")
        sb.AppendLine(guideContent)
        sb.AppendLine("---END STYLE GUIDE---")
        sb.AppendLine()
        sb.AppendLine("请从以上规范中提取所有格式要求，对应到可用标签。如果规范中未明确某项参数，使用合理的默认值。")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 将AI返回的JSON解析为SemanticStyleMapping
    ''' </summary>
    ''' <param name="aiResponseJson">AI返回的JSON字符串</param>
    ''' <param name="guideName">规范名称</param>
    ''' <param name="guideId">规范来源ID</param>
    Public Shared Function ParseAiResponse(
        aiResponseJson As String,
        guideName As String,
        guideId As String) As SemanticStyleMapping

        Dim mapping As New SemanticStyleMapping()
        mapping.Name = guideName
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide
        mapping.SourceId = guideId

        Try
            Dim json = CleanJsonResponse(aiResponseJson)
            Dim obj = JObject.Parse(json)

            ' 解析语义标签
            Dim tagsArray = obj("semanticTags")
            If TypeOf tagsArray Is JArray Then
                For Each tagObj In CType(tagsArray, JArray)
                    Dim tag As New SemanticTag()
                    tag.TagId = If(tagObj("tagId")?.ToString(), "")
                    tag.DisplayName = If(tagObj("displayName")?.ToString(), "")
                    tag.MatchHint = If(tagObj("matchHint")?.ToString(), "")
                    tag.ParentTagId = SemanticTagRegistry.GetParentTag(tag.TagId)
                    tag.Level = SemanticTagRegistry.GetTagLevel(tag.TagId)

                    ' 解析字体
                    Dim fontObj = tagObj("font")
                    If fontObj IsNot Nothing Then
                        tag.Font.FontNameCN = If(fontObj("fontNameCN")?.ToString(), tag.Font.FontNameCN)
                        tag.Font.FontNameEN = If(fontObj("fontNameEN")?.ToString(), tag.Font.FontNameEN)
                        If fontObj("fontSize") IsNot Nothing Then tag.Font.FontSize = CDbl(fontObj("fontSize"))
                        If fontObj("bold") IsNot Nothing Then tag.Font.Bold = CBool(fontObj("bold"))
                        If fontObj("italic") IsNot Nothing Then tag.Font.Italic = CBool(fontObj("italic"))
                        If fontObj("underline") IsNot Nothing Then tag.Font.Underline = CBool(fontObj("underline"))
                    End If

                    ' 解析段落
                    Dim paraObj = tagObj("paragraph")
                    If paraObj IsNot Nothing Then
                        tag.Paragraph.Alignment = If(paraObj("alignment")?.ToString(), tag.Paragraph.Alignment)
                        If paraObj("firstLineIndent") IsNot Nothing Then tag.Paragraph.FirstLineIndent = CDbl(paraObj("firstLineIndent"))
                        If paraObj("lineSpacing") IsNot Nothing Then tag.Paragraph.LineSpacing = CDbl(paraObj("lineSpacing"))
                        If paraObj("spaceBefore") IsNot Nothing Then tag.Paragraph.SpaceBefore = CDbl(paraObj("spaceBefore"))
                        If paraObj("spaceAfter") IsNot Nothing Then tag.Paragraph.SpaceAfter = CDbl(paraObj("spaceAfter"))
                        If paraObj("leftIndent") IsNot Nothing Then tag.Paragraph.LeftIndent = CDbl(paraObj("leftIndent"))
                    End If

                    ' 解析颜色
                    Dim colorObj = tagObj("color")
                    If colorObj IsNot Nothing Then
                        tag.Color.FontColor = If(colorObj("fontColor")?.ToString(), tag.Color.FontColor)
                        tag.Color.BackgroundColor = If(colorObj("backgroundColor")?.ToString(), tag.Color.BackgroundColor)
                    End If

                    If Not String.IsNullOrEmpty(tag.TagId) Then
                        mapping.SemanticTags.Add(tag)
                    End If
                Next
            End If

            ' 解析页面设置
            Dim pageObj = obj("pageConfig")
            If pageObj IsNot Nothing Then
                Dim marginsObj = pageObj("margins")
                If marginsObj IsNot Nothing Then
                    If marginsObj("top") IsNot Nothing Then mapping.PageConfig.Margins.Top = CDbl(marginsObj("top"))
                    If marginsObj("bottom") IsNot Nothing Then mapping.PageConfig.Margins.Bottom = CDbl(marginsObj("bottom"))
                    If marginsObj("left") IsNot Nothing Then mapping.PageConfig.Margins.Left = CDbl(marginsObj("left"))
                    If marginsObj("right") IsNot Nothing Then mapping.PageConfig.Margins.Right = CDbl(marginsObj("right"))
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"解析AI规范转换结果失败: {ex.Message}")
        End Try

        ' 确保基础标签
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_BODY_NORMAL) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_BODY_NORMAL, "正文",
                SemanticTagRegistry.TAG_BODY, 2, "普通正文段落"))
        End If

        Return mapping
    End Function

    ''' <summary>清理AI返回的JSON（去除markdown标记等）</summary>
    Private Shared Function CleanJsonResponse(response As String) As String
        If String.IsNullOrWhiteSpace(response) Then Return "{}"

        Dim json = response.Trim()

        ' 移除markdown代码块
        If json.StartsWith("```") Then
            Dim firstNewline = json.IndexOf(vbLf)
            If firstNewline > 0 Then
                json = json.Substring(firstNewline + 1)
            End If
            If json.EndsWith("```") Then
                json = json.Substring(0, json.LastIndexOf("```"))
            End If
            json = json.Trim()
        End If

        Return json
    End Function
End Class
