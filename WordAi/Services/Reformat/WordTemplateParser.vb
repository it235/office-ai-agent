' WordAi\Services\Reformat\WordTemplateParser.vb
' 从.docx文件提取样式，生成SemanticStyleMapping

Imports System.Diagnostics
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports ShareRibbon

''' <summary>
''' Word模板解析器 - 从.docx文件提取样式规则
''' 使用OpenXML SDK解析，不需要打开Word实例
''' 支持样式继承链解析，完整提取所有标题/正文/列表等样式
''' </summary>
Public Class WordTemplateParser

    ''' <summary>样式继承链缓存</summary>
    Private Shared _styleMap As Dictionary(Of String, Style) = Nothing

    ''' <summary>
    ''' 从.docx文件提取语义样式映射
    ''' </summary>
    Public Shared Function ExtractFromDocx(filePath As String) As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Name = System.IO.Path.GetFileNameWithoutExtension(filePath)
        mapping.SourceType = SemanticMappingSourceType.FromDocxTemplate

        Try
            Using doc = WordprocessingDocument.Open(filePath, False)
                ' 构建样式继承缓存
                BuildStyleMap(doc)

                ' 提取样式定义中的标签
                ExtractStyleDefinitions(doc, mapping)

                ' 扫描文档正文，补充和修正样式信息
                EnrichFromDocumentBody(doc, mapping)

                ' 提取页面设置
                ExtractPageSetup(doc, mapping)

                ' 清理缓存
                _styleMap = Nothing
            End Using
        Catch ex As Exception
            Debug.WriteLine($"解析.docx模板失败: {ex.Message}")
        End Try

        ' 确保至少有基础标签
        EnsureBasicTags(mapping)

        Return mapping
    End Function

#Region "样式继承链"

    ''' <summary>构建样式ID → Style对象的字典</summary>
    Private Shared Sub BuildStyleMap(doc As WordprocessingDocument)
        _styleMap = New Dictionary(Of String, Style)(StringComparer.OrdinalIgnoreCase)

        Dim stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart
        If stylesPart?.Styles Is Nothing Then Return

        For Each style In stylesPart.Styles.Elements(Of Style)()
            Dim sid = If(style.StyleId?.Value, "")
            If Not String.IsNullOrEmpty(sid) AndAlso Not _styleMap.ContainsKey(sid) Then
                _styleMap(sid) = style
            End If
        Next
    End Sub

    ''' <summary>
    ''' 沿继承链解析完整的字体配置
    ''' OpenXML样式可以BasedOn另一个样式，属性逐级继承
    ''' </summary>
    Private Shared Function ResolveFont(styleId As String) As FontConfig
        Dim font As New FontConfig()
        Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim chain As New List(Of String)()

        ' 收集继承链（从子到父）
        Dim currentId = styleId
        While Not String.IsNullOrEmpty(currentId) AndAlso Not visited.Contains(currentId)
            visited.Add(currentId)
            chain.Add(currentId)
            If _styleMap IsNot Nothing AndAlso _styleMap.ContainsKey(currentId) Then
                Dim s = _styleMap(currentId)
                currentId = If(s.BasedOn?.Val?.Value, "")
            Else
                Exit While
            End If
        End While

        ' 从父到子逐级覆盖（父先，子后覆盖）
        chain.Reverse()
        For Each sid In chain
            If _styleMap IsNot Nothing AndAlso _styleMap.ContainsKey(sid) Then
                Dim s = _styleMap(sid)
                ApplyRunPropertiesToFont(s.StyleRunProperties, font)
            End If
        Next

        Return font
    End Function

    ''' <summary>沿继承链解析完整的段落配置</summary>
    Private Shared Function ResolveParagraph(styleId As String) As ParagraphConfig
        Dim para As New ParagraphConfig()
        Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim chain As New List(Of String)()

        Dim currentId = styleId
        While Not String.IsNullOrEmpty(currentId) AndAlso Not visited.Contains(currentId)
            visited.Add(currentId)
            chain.Add(currentId)
            If _styleMap IsNot Nothing AndAlso _styleMap.ContainsKey(currentId) Then
                Dim s = _styleMap(currentId)
                currentId = If(s.BasedOn?.Val?.Value, "")
            Else
                Exit While
            End If
        End While

        chain.Reverse()
        For Each sid In chain
            If _styleMap IsNot Nothing AndAlso _styleMap.ContainsKey(sid) Then
                Dim s = _styleMap(sid)
                ApplyParagraphPropertiesToConfig(s.StyleParagraphProperties, para)
            End If
        Next

        Return para
    End Function

#End Region

#Region "样式定义提取"

    ''' <summary>从StylesPart提取所有样式定义</summary>
    Private Shared Sub ExtractStyleDefinitions(doc As WordprocessingDocument, mapping As SemanticStyleMapping)
        If _styleMap Is Nothing OrElse _styleMap.Count = 0 Then Return

        For Each kvp In _styleMap
            Dim style = kvp.Value
            If style.Type Is Nothing OrElse style.Type.Value <> StyleValues.Paragraph Then Continue For

            Dim styleId = If(style.StyleId?.Value, "")
            Dim styleName = If(style.StyleName?.Val?.Value, styleId)

            ' 匹配样式到语义标签
            Dim tagInfo = MatchStyleToTag(styleId, styleName)
            If tagInfo Is Nothing Then Continue For

            ' 沿继承链解析完整属性
            Dim font = ResolveFont(styleId)
            Dim para = ResolveParagraph(styleId)

            Dim tag As New SemanticTag(tagInfo.TagId, tagInfo.DisplayName, tagInfo.ParentTagId,
                                       SemanticTagRegistry.GetTagLevel(tagInfo.TagId), tagInfo.MatchHint)
            tag.Font = font
            tag.Paragraph = para

            ' 避免重复添加（同一个tagId只保留第一个）
            If Not mapping.SemanticTags.Any(Function(t) t.TagId = tag.TagId) Then
                mapping.SemanticTags.Add(tag)
            End If
        Next
    End Sub

    ''' <summary>匹配样式到语义标签信息</summary>
    Private Shared Function MatchStyleToTag(styleId As String, styleName As String) As TagMatchInfo
        Dim id = If(styleId, "").ToLower()
        Dim name = If(styleName, "")

        ' === 标题类 ===
        ' Heading 1-9 / 标题 1-9
        For level As Integer = 1 To 9
            If id = "heading" & level.ToString() OrElse
               id = level.ToString() OrElse
               name = "Heading " & level.ToString() OrElse
               name = "heading " & level.ToString() OrElse
               name.Contains("标题 " & level.ToString()) OrElse
               name.Contains("标题" & level.ToString()) Then

                Dim tagId As String
                Dim displayName As String
                Dim matchHint As String

                Select Case level
                    Case 1
                        tagId = SemanticTagRegistry.TAG_TITLE_1
                        displayName = "一级标题"
                        matchHint = "包含'第X章'、'一、'或文档主要章节标题"
                    Case 2
                        tagId = SemanticTagRegistry.TAG_TITLE_2
                        displayName = "二级标题"
                        matchHint = "包含'1.1'、'（一）'或章节子标题"
                    Case 3
                        tagId = SemanticTagRegistry.TAG_TITLE_3
                        displayName = "三级标题"
                        matchHint = "包含'1.1.1'、'1.'或小节标题"
                    Case Else
                        tagId = $"title.{level}"
                        displayName = $"{level}级标题"
                        matchHint = $"{level}级子标题"
                End Select

                Return New TagMatchInfo(tagId, displayName, SemanticTagRegistry.TAG_TITLE, matchHint)
            End If
        Next

        ' === 正文类 ===
        If id = "normal" OrElse id = "bodytext" OrElse id = "body" OrElse
           name = "Normal" OrElse name = "正文" OrElse name = "Body Text" OrElse
           name.Contains("正文") Then
            Return New TagMatchInfo(SemanticTagRegistry.TAG_BODY_NORMAL, "正文", SemanticTagRegistry.TAG_BODY, "普通正文段落")
        End If

        ' === 列表类 ===
        If id = "listparagraph" OrElse id.Contains("list") OrElse
           name.Contains("列表") OrElse name = "List Paragraph" Then
            Return New TagMatchInfo(SemanticTagRegistry.TAG_LIST_ORDERED, "列表", SemanticTagRegistry.TAG_LIST, "列表项")
        End If

        ' === 引用类 ===
        If id.Contains("quote") OrElse id.Contains("blockquote") OrElse
           name.Contains("引用") OrElse name.Contains("Quote") Then
            Return New TagMatchInfo(SemanticTagRegistry.TAG_QUOTE, "引用", "", "引用段落")
        End If

        ' === 题注类 ===
        If id.Contains("caption") OrElse name.Contains("题注") OrElse name.Contains("Caption") Then
            Return New TagMatchInfo(SemanticTagRegistry.TAG_CAPTION, "题注", "", "图表题注")
        End If

        ' === 目录类（跳过） ===
        If id.StartsWith("toc") OrElse name.Contains("目录") Then
            Return Nothing
        End If

        ' === 页眉页脚（跳过）===
        If id = "header" OrElse id = "footer" OrElse name.Contains("页眉") OrElse name.Contains("页脚") Then
            Return Nothing
        End If

        Return Nothing
    End Function

#End Region

#Region "文档正文扫描补充"

    ''' <summary>
    ''' 扫描文档正文中实际使用的段落，补充样式定义中可能缺失的格式信息
    ''' 当样式定义的字体/字号为空时，用实际段落的直接格式填充
    ''' </summary>
    Private Shared Sub EnrichFromDocumentBody(doc As WordprocessingDocument, mapping As SemanticStyleMapping)
        Dim body = doc.MainDocumentPart?.Document?.Body
        If body Is Nothing Then Return

        ' 按样式ID分组收集实际段落的格式信息
        Dim samplesByStyle As New Dictionary(Of String, Paragraph)(StringComparer.OrdinalIgnoreCase)

        For Each p In body.Elements(Of Paragraph)()
            Dim pStyleId = If(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value, "Normal")
            ' 只保留每种样式的第一个样本
            If Not samplesByStyle.ContainsKey(pStyleId) Then
                samplesByStyle(pStyleId) = p
            End If
        Next

        ' 用实际段落的格式补充映射中空缺的属性
        For Each tag In mapping.SemanticTags
            ' 查找匹配此标签的样式ID
            Dim matchedStyleId As String = Nothing
            For Each kvp In samplesByStyle
                Dim tagInfo = MatchStyleToTag(kvp.Key, GetStyleName(kvp.Key))
                If tagInfo IsNot Nothing AndAlso tagInfo.TagId = tag.TagId Then
                    matchedStyleId = kvp.Key
                    Exit For
                End If
            Next

            If matchedStyleId Is Nothing OrElse Not samplesByStyle.ContainsKey(matchedStyleId) Then Continue For

            Dim samplePara = samplesByStyle(matchedStyleId)

            ' 从段落的首个Run提取直接格式（补充样式定义中缺失的部分）
            Dim firstRun = samplePara.Elements(Of Run)().FirstOrDefault()
            If firstRun IsNot Nothing AndAlso firstRun.RunProperties IsNot Nothing Then
                Dim rPr = firstRun.RunProperties

                ' 只在样式定义没有值时才用直接格式补充
                If String.IsNullOrEmpty(tag.Font.FontNameCN) Then
                    Dim rf = rPr.GetFirstChild(Of RunFonts)()
                    If rf?.EastAsia IsNot Nothing Then tag.Font.FontNameCN = rf.EastAsia.Value
                End If
                If String.IsNullOrEmpty(tag.Font.FontNameEN) Then
                    Dim rf = rPr.GetFirstChild(Of RunFonts)()
                    If rf?.Ascii IsNot Nothing Then tag.Font.FontNameEN = rf.Ascii.Value
                End If
                If tag.Font.FontSize <= 0 Then
                    Dim fs = rPr.GetFirstChild(Of FontSize)()
                    If fs?.Val IsNot Nothing Then
                        Dim halfPt As Double
                        If Double.TryParse(fs.Val.Value, halfPt) Then
                            tag.Font.FontSize = halfPt / 2.0
                        End If
                    End If
                End If
            End If

            ' 从段落属性补充段落格式
            If samplePara.ParagraphProperties IsNot Nothing Then
                Dim pPr = samplePara.ParagraphProperties

                ' 对齐
                If String.IsNullOrEmpty(tag.Paragraph.Alignment) Then
                    Dim jc = pPr.GetFirstChild(Of Justification)()
                    If jc?.Val IsNot Nothing Then
                        Select Case jc.Val.Value
                            Case JustificationValues.Center : tag.Paragraph.Alignment = "center"
                            Case JustificationValues.Right : tag.Paragraph.Alignment = "right"
                            Case JustificationValues.Both, JustificationValues.Distribute : tag.Paragraph.Alignment = "justify"
                            Case Else : tag.Paragraph.Alignment = "left"
                        End Select
                    End If
                End If

                ' 缩进
                If tag.Paragraph.FirstLineIndent <= 0 Then
                    Dim indent = pPr.GetFirstChild(Of Indentation)()
                    If indent?.FirstLineChars IsNot Nothing Then
                        Dim chars As Integer
                        If Integer.TryParse(indent.FirstLineChars.Value, chars) Then
                            tag.Paragraph.FirstLineIndent = chars / 100.0
                        End If
                    End If
                End If

                ' 行距
                If tag.Paragraph.LineSpacing <= 0 Then
                    Dim spacing = pPr.GetFirstChild(Of SpacingBetweenLines)()
                    If spacing?.Line IsNot Nothing Then
                        Dim lineVal As Integer
                        If Integer.TryParse(spacing.Line.Value, lineVal) Then
                            tag.Paragraph.LineSpacing = lineVal / 240.0
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>获取样式名称</summary>
    Private Shared Function GetStyleName(styleId As String) As String
        If _styleMap IsNot Nothing AndAlso _styleMap.ContainsKey(styleId) Then
            Return If(_styleMap(styleId).StyleName?.Val?.Value, styleId)
        End If
        Return styleId
    End Function

#End Region

#Region "属性提取辅助方法"

    ''' <summary>将RunProperties中的属性应用到FontConfig（增量覆盖）</summary>
    Private Shared Sub ApplyRunPropertiesToFont(rPr As StyleRunProperties, font As FontConfig)
        If rPr Is Nothing Then Return

        Dim runFonts = rPr.GetFirstChild(Of RunFonts)()
        If runFonts IsNot Nothing Then
            If runFonts.EastAsia IsNot Nothing Then font.FontNameCN = runFonts.EastAsia.Value
            If runFonts.Ascii IsNot Nothing Then font.FontNameEN = runFonts.Ascii.Value
        End If

        Dim fontSize = rPr.GetFirstChild(Of FontSize)()
        If fontSize?.Val IsNot Nothing Then
            Dim halfPt As Double
            If Double.TryParse(fontSize.Val.Value, halfPt) Then
                font.FontSize = halfPt / 2.0
            End If
        End If

        Dim bold = rPr.GetFirstChild(Of Bold)()
        If bold IsNot Nothing Then
            font.Bold = (bold.Val Is Nothing OrElse bold.Val.Value)
        End If

        Dim italic = rPr.GetFirstChild(Of Italic)()
        If italic IsNot Nothing Then
            font.Italic = (italic.Val Is Nothing OrElse italic.Val.Value)
        End If

        Dim underline = rPr.GetFirstChild(Of Underline)()
        If underline?.Val IsNot Nothing Then
            font.Underline = (underline.Val.Value <> UnderlineValues.None)
        End If
    End Sub

    ''' <summary>将ParagraphProperties中的属性应用到ParagraphConfig（增量覆盖）</summary>
    Private Shared Sub ApplyParagraphPropertiesToConfig(pPr As StyleParagraphProperties, para As ParagraphConfig)
        If pPr Is Nothing Then Return

        Dim jc = pPr.GetFirstChild(Of Justification)()
        If jc?.Val IsNot Nothing Then
            Select Case jc.Val.Value
                Case JustificationValues.Center : para.Alignment = "center"
                Case JustificationValues.Right : para.Alignment = "right"
                Case JustificationValues.Both, JustificationValues.Distribute : para.Alignment = "justify"
                Case Else : para.Alignment = "left"
            End Select
        End If

        Dim indent = pPr.GetFirstChild(Of Indentation)()
        If indent IsNot Nothing Then
            If indent.FirstLineChars IsNot Nothing Then
                Dim chars As Integer
                If Integer.TryParse(indent.FirstLineChars.Value, chars) Then
                    para.FirstLineIndent = chars / 100.0
                End If
            End If
            If indent.Left IsNot Nothing Then
                Dim twips As Integer
                If Integer.TryParse(indent.Left.Value, twips) Then
                    para.LeftIndent = twips / 567.0
                End If
            End If
        End If

        Dim spacing = pPr.GetFirstChild(Of SpacingBetweenLines)()
        If spacing IsNot Nothing Then
            If spacing.Line IsNot Nothing Then
                Dim lineVal As Integer
                If Integer.TryParse(spacing.Line.Value, lineVal) Then
                    para.LineSpacing = lineVal / 240.0
                End If
            End If
            If spacing.Before IsNot Nothing Then
                Dim beforeVal As Integer
                If Integer.TryParse(spacing.Before.Value, beforeVal) Then
                    para.SpaceBefore = beforeVal / 240.0
                End If
            End If
            If spacing.After IsNot Nothing Then
                Dim afterVal As Integer
                If Integer.TryParse(spacing.After.Value, afterVal) Then
                    para.SpaceAfter = afterVal / 240.0
                End If
            End If
        End If
    End Sub

#End Region

#Region "页面设置提取"

    ''' <summary>提取页面设置</summary>
    Private Shared Sub ExtractPageSetup(doc As WordprocessingDocument, mapping As SemanticStyleMapping)
        Dim body = doc.MainDocumentPart?.Document?.Body
        If body Is Nothing Then Return

        ' 查找最后一个SectionProperties（文档默认节）
        Dim sectPr = body.GetFirstChild(Of SectionProperties)()
        If sectPr Is Nothing Then
            ' 也可能在最后一个段落的ParagraphProperties中
            Dim lastPara = body.Elements(Of Paragraph)().LastOrDefault()
            sectPr = lastPara?.ParagraphProperties?.GetFirstChild(Of SectionProperties)()
        End If
        If sectPr Is Nothing Then Return

        ' 页边距（OpenXML单位为twips，1cm = 567 twips）
        Dim pgMar = sectPr.GetFirstChild(Of PageMargin)()
        If pgMar IsNot Nothing Then
            If pgMar.Top IsNot Nothing Then mapping.PageConfig.Margins.Top = CInt(pgMar.Top.Value) / 567.0
            If pgMar.Bottom IsNot Nothing Then mapping.PageConfig.Margins.Bottom = CInt(pgMar.Bottom.Value) / 567.0
            If pgMar.Left IsNot Nothing Then mapping.PageConfig.Margins.Left = CUInt(pgMar.Left.Value) / 567.0
            If pgMar.Right IsNot Nothing Then mapping.PageConfig.Margins.Right = CUInt(pgMar.Right.Value) / 567.0
        End If

        ' 页眉页脚
        For Each headerRef In sectPr.Elements(Of HeaderReference)()
            Try
                Dim headerPart = doc.MainDocumentPart.GetPartById(headerRef.Id.Value)
                If TypeOf headerPart Is HeaderPart Then
                    mapping.PageConfig.Header.Enabled = True
                    Dim headerBody = DirectCast(headerPart, HeaderPart).Header
                    If headerBody IsNot Nothing Then
                        mapping.PageConfig.Header.Content = headerBody.InnerText
                    End If
                End If
            Catch
            End Try
        Next

        For Each footerRef In sectPr.Elements(Of FooterReference)()
            Try
                Dim footerPart = doc.MainDocumentPart.GetPartById(footerRef.Id.Value)
                If TypeOf footerPart Is FooterPart Then
                    mapping.PageConfig.Footer.Enabled = True
                    Dim footerBody = DirectCast(footerPart, FooterPart).Footer
                    If footerBody IsNot Nothing Then
                        mapping.PageConfig.Footer.Content = footerBody.InnerText
                    End If
                End If
            Catch
            End Try
        Next
    End Sub

#End Region

#Region "兜底"

    ''' <summary>确保映射至少包含基础标签</summary>
    Private Shared Sub EnsureBasicTags(mapping As SemanticStyleMapping)
        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_BODY_NORMAL) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_BODY_NORMAL, "正文",
                SemanticTagRegistry.TAG_BODY, 2, "普通正文段落"))
        End If

        If Not mapping.SemanticTags.Any(Function(t) t.TagId = SemanticTagRegistry.TAG_TITLE_1) Then
            mapping.SemanticTags.Add(New SemanticTag(
                SemanticTagRegistry.TAG_TITLE_1, "一级标题",
                SemanticTagRegistry.TAG_TITLE, 2, "主要章节标题"))
        End If
    End Sub

#End Region

    ''' <summary>标签匹配信息（内部辅助类）</summary>
    Private Class TagMatchInfo
        Public Property TagId As String
        Public Property DisplayName As String
        Public Property ParentTagId As String
        Public Property MatchHint As String

        Public Sub New(tagId As String, displayName As String, parentTagId As String, matchHint As String)
            Me.TagId = tagId
            Me.DisplayName = displayName
            Me.ParentTagId = parentTagId
            Me.MatchHint = matchHint
        End Sub
    End Class
End Class
