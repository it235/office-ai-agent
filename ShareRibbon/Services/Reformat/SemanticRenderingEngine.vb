' ShareRibbon\Services\Reformat\SemanticRenderingEngine.vb
' 确定性渲染引擎 - 根据语义标签应用Word格式

Imports Newtonsoft.Json.Linq

''' <summary>
''' 语义渲染引擎 - 根据语义标签确定性地应用Word格式
''' 核心类：接收AI标注结果 + SemanticStyleMapping，渲染到Word段落
''' </summary>
Public Class SemanticRenderingEngine

    ''' <summary>渲染结果统计</summary>
    Public Class RenderResult
        Public Property AppliedCount As Integer = 0
        Public Property SkippedCount As Integer = 0
        Public Property TagUsage As New Dictionary(Of String, Integer)()
        Public Property Errors As New List(Of String)()

        ''' <summary>转换为JSON（用于推送前端）</summary>
        Public Function ToJson() As JObject
            Dim result As New JObject()
            result("appliedCount") = AppliedCount
            result("skippedCount") = SkippedCount
            Dim tagsObj As New JObject()
            For Each kvp In TagUsage
                tagsObj(kvp.Key) = kvp.Value
            Next
            result("tags") = tagsObj
            Return result
        End Function
    End Class

    ''' <summary>
    ''' 应用语义排版到Word段落
    ''' </summary>
    ''' <param name="taggedParagraphs">AI标注结果: List of (paraIndex, tagId)</param>
    ''' <param name="mapping">语义样式映射</param>
    ''' <param name="wordParagraphs">Word段落对象列表</param>
    ''' <param name="paragraphTypes">段落类型列表（text/image/table/formula）</param>
    ''' <param name="wordApp">Word Application对象（用于页面设置）</param>
    Public Shared Function ApplySemanticFormatting(
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping,
        wordParagraphs As List(Of Object),
        paragraphTypes As List(Of String),
        Optional wordApp As Object = Nothing) As RenderResult

        Dim result As New RenderResult()

        ' 构建 tagId → SemanticTag 查找字典
        Dim tagDict As New Dictionary(Of String, SemanticTag)()
        For Each tag In mapping.SemanticTags
            If Not tagDict.ContainsKey(tag.TagId) Then
                tagDict(tag.TagId) = tag
            End If
        Next

        ' 遍历标注结果，逐段落应用格式
        For Each tagged In taggedParagraphs
            If tagged.ParaIndex < 0 OrElse tagged.ParaIndex >= wordParagraphs.Count Then
                result.Errors.Add($"段落索引越界: {tagged.ParaIndex}")
                Continue For
            End If

            ' 跳过非文本段落
            If paragraphTypes IsNot Nothing AndAlso tagged.ParaIndex < paragraphTypes.Count Then
                Dim pType = paragraphTypes(tagged.ParaIndex)
                If pType <> "text" Then
                    result.SkippedCount += 1
                    Continue For
                End If
            End If

            ' 查找语义标签（精确匹配 → 父级回退）
            Dim semanticTag = FindTagWithFallback(tagged.TagId, tagDict, mapping)
            If semanticTag Is Nothing Then
                result.Errors.Add($"未找到标签: {tagged.TagId}")
                result.SkippedCount += 1
                Continue For
            End If

            ' 应用格式到段落
            Try
                Dim para = wordParagraphs(tagged.ParaIndex)
                Dim range = para.Range
                ApplyFormatToRange(range, semanticTag)
                result.AppliedCount += 1

                ' 记录标签使用次数
                If result.TagUsage.ContainsKey(tagged.TagId) Then
                    result.TagUsage(tagged.TagId) += 1
                Else
                    result.TagUsage(tagged.TagId) = 1
                End If
            Catch ex As Exception
                result.Errors.Add($"段落{tagged.ParaIndex}格式应用失败: {ex.Message}")
            End Try
        Next

        ' 应用页面设置（如果提供了Word应用对象）
        If wordApp IsNot Nothing AndAlso mapping.PageConfig IsNot Nothing Then
            Try
                ApplyPageConfig(wordApp, mapping.PageConfig)
            Catch ex As Exception
                result.Errors.Add($"页面设置应用失败: {ex.Message}")
            End Try
        End If

        Return result
    End Function

    ''' <summary>查找标签（精确匹配 → 父级回退）</summary>
    Private Shared Function FindTagWithFallback(
        tagId As String,
        tagDict As Dictionary(Of String, SemanticTag),
        mapping As SemanticStyleMapping) As SemanticTag

        ' 精确匹配
        If tagDict.ContainsKey(tagId) Then Return tagDict(tagId)

        ' 父级回退
        Dim parentId = SemanticTagRegistry.GetParentTag(tagId)
        If Not String.IsNullOrEmpty(parentId) AndAlso tagDict.ContainsKey(parentId) Then
            Return tagDict(parentId)
        End If

        ' 通过mapping的FindTag方法
        Return mapping.FindTag(tagId)
    End Function

    ''' <summary>
    ''' 将语义标签的格式应用到Word Range
    ''' </summary>
    Public Shared Sub ApplyFormatToRange(targetRange As Object, tag As SemanticTag)
        If targetRange Is Nothing OrElse tag Is Nothing Then Return

        ' 应用字体
        ApplyFontConfig(targetRange, tag.Font)

        ' 应用段落格式
        ApplyParagraphConfig(targetRange, tag.Paragraph)

        ' 应用颜色
        ApplyColorConfig(targetRange, tag.Color)
    End Sub

    ''' <summary>应用字体配置</summary>
    Private Shared Sub ApplyFontConfig(targetRange As Object, font As FontConfig)
        If font Is Nothing Then Return

        Try
            ' 中文字体
            If Not String.IsNullOrEmpty(font.FontNameCN) Then
                targetRange.Font.NameFarEast = font.FontNameCN
            End If

            ' 英文字体
            If Not String.IsNullOrEmpty(font.FontNameEN) Then
                targetRange.Font.Name = font.FontNameEN
            End If

            ' 字号
            If font.FontSize > 0 Then
                targetRange.Font.Size = CSng(font.FontSize)
            End If

            ' 加粗
            targetRange.Font.Bold = If(font.Bold, -1, 0)

            ' 斜体
            targetRange.Font.Italic = If(font.Italic, -1, 0)

            ' 下划线
            If font.Underline Then
                targetRange.Font.Underline = 1 ' wdUnderlineSingle
            Else
                targetRange.Font.Underline = 0 ' wdUnderlineNone
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用字体配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>应用段落配置</summary>
    Private Shared Sub ApplyParagraphConfig(targetRange As Object, para As ParagraphConfig)
        If para Is Nothing Then Return

        Try
            ' 对齐方式
            If Not String.IsNullOrEmpty(para.Alignment) Then
                Select Case para.Alignment.ToLower()
                    Case "center"
                        targetRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                    Case "right"
                        targetRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
                    Case "justify"
                        targetRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
                    Case Else
                        targetRange.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                End Select
            End If

            ' 首行缩进
            If para.FirstLineIndent > 0 Then
                Try
                    targetRange.ParagraphFormat.CharacterUnitFirstLineIndent = CSng(para.FirstLineIndent)
                Catch
                    ' 回退: 使用磅值（1字符约10.5磅）
                    targetRange.ParagraphFormat.FirstLineIndent = CSng(para.FirstLineIndent * 10.5)
                End Try
            End If

            ' 行距
            If para.LineSpacing > 0 Then
                If para.LineSpacing = 1.0 Then
                    targetRange.ParagraphFormat.LineSpacingRule = 0 ' wdLineSpaceSingle
                ElseIf para.LineSpacing = 1.5 Then
                    targetRange.ParagraphFormat.LineSpacingRule = 1 ' wdLineSpace1pt5
                ElseIf para.LineSpacing = 2.0 Then
                    targetRange.ParagraphFormat.LineSpacingRule = 2 ' wdLineSpaceDouble
                Else
                    targetRange.ParagraphFormat.LineSpacingRule = 5 ' wdLineSpaceMultiple
                    targetRange.ParagraphFormat.LineSpacing = CSng(12 * para.LineSpacing)
                End If
            End If

            ' 段前间距
            If para.SpaceBefore > 0 Then
                targetRange.ParagraphFormat.SpaceBefore = CSng(para.SpaceBefore * 12) ' 行→磅
            End If

            ' 段后间距
            If para.SpaceAfter > 0 Then
                targetRange.ParagraphFormat.SpaceAfter = CSng(para.SpaceAfter * 12) ' 行→磅
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用段落配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>应用颜色配置</summary>
    Private Shared Sub ApplyColorConfig(targetRange As Object, color As ColorConfig)
        If color Is Nothing Then Return

        Try
            If Not String.IsNullOrEmpty(color.FontColor) AndAlso color.FontColor <> "#000000" Then
                Dim clr = System.Drawing.ColorTranslator.FromHtml(color.FontColor)
                targetRange.Font.Color = System.Drawing.ColorTranslator.ToOle(clr)
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用颜色配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>应用页面设置</summary>
    Private Shared Sub ApplyPageConfig(wordApp As Object, config As PageConfig)
        If config Is Nothing Then Return

        Try
            Dim doc = wordApp.ActiveDocument
            Dim pageSetup = doc.PageSetup

            ' 页边距（cm → 磅，1cm = 28.35磅）
            If config.Margins IsNot Nothing Then
                Dim cmToPt As Double = 28.35
                pageSetup.TopMargin = config.Margins.Top * cmToPt
                pageSetup.BottomMargin = config.Margins.Bottom * cmToPt
                pageSetup.LeftMargin = config.Margins.Left * cmToPt
                pageSetup.RightMargin = config.Margins.Right * cmToPt
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用页面设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 排版后校验 - 检查渲染结果是否匹配预期
    ''' </summary>
    Public Shared Function ValidateRenderedDocument(
        wordParagraphs As List(Of Object),
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping) As List(Of String)

        Dim deviations As New List(Of String)()

        For Each tagged In taggedParagraphs
            If tagged.ParaIndex < 0 OrElse tagged.ParaIndex >= wordParagraphs.Count Then Continue For

            Dim expectedTag = mapping.FindTag(tagged.TagId)
            If expectedTag Is Nothing Then Continue For

            Try
                Dim para = wordParagraphs(tagged.ParaIndex)
                Dim range = para.Range

                ' 检查字号
                If expectedTag.Font.FontSize > 0 Then
                    Dim actualSize As Double = CDbl(range.Font.Size)
                    If Math.Abs(actualSize - expectedTag.Font.FontSize) > 0.5 Then
                        deviations.Add($"段落{tagged.ParaIndex}: 字号偏差 期望{expectedTag.Font.FontSize}pt 实际{actualSize}pt")
                        ' 自动修正
                        range.Font.Size = CSng(expectedTag.Font.FontSize)
                    End If
                End If
            Catch ex As Exception
                ' 校验失败不影响主流程
                Debug.WriteLine($"校验段落{tagged.ParaIndex}失败: {ex.Message}")
            End Try
        Next

        Return deviations
    End Function
End Class

''' <summary>
''' AI标注的段落结构
''' </summary>
Public Class TaggedParagraph
    ''' <summary>段落索引</summary>
    Public Property ParaIndex As Integer

    ''' <summary>语义标签ID</summary>
    Public Property TagId As String = ""

    Public Sub New()
    End Sub

    Public Sub New(paraIndex As Integer, tagId As String)
        Me.ParaIndex = paraIndex
        Me.TagId = tagId
    End Sub
End Class
