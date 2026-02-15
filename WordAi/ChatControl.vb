Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Math
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Imports DocumentFormat.OpenXml.Packaging
Imports HtmlToOpenXml

Public Class ChatControl
    Inherits BaseChatControl


    Private sheetContentItems As New Dictionary(Of String, Tuple(Of System.Windows.Forms.Label, System.Windows.Forms.Button))

    ' 排版上下文：存储待格式化的段落和样式信息
    Private _reformatParagraphs As List(Of Object) = Nothing
    Private _reformatStyles As List(Of String) = Nothing
    Private _reformatTypes As List(Of String) = Nothing ' text/image/table/formula
    Private _reformatMapping As SemanticStyleMapping = Nothing ' 语义映射上下文
    Private _reformatRetryCount As Integer = 0 ' 重试计数器，防止死循环
    Private Const MAX_REFORMAT_RETRIES As Integer = 2 ' 最大重试次数

    ''' <summary>
    ''' 设置排版上下文，用于规则匹配后应用格式
    ''' </summary>
    Public Sub SetReformatContext(paragraphs As List(Of Object), styles As List(Of String), Optional types As List(Of String) = Nothing, Optional mapping As SemanticStyleMapping = Nothing)
        _reformatParagraphs = paragraphs
        _reformatStyles = styles
        _reformatTypes = types
        _reformatMapping = mapping
        _reformatRetryCount = 0 ' 重置重试计数器
    End Sub

    ''' <summary>
    ''' 获取当前 Office 应用程序名称
    ''' </summary>
    Protected Overrides Function GetOfficeApplicationName() As String
        Return "Word"
    End Function

    ''' <summary>
    ''' 处理模板预览请求（Word 特定实现）
    ''' </summary>
    Protected Overrides Sub HandlePreviewTemplateInWord(jsonDoc As Newtonsoft.Json.Linq.JObject)
        Dim templateId As String = jsonDoc("templateId")?.ToString()
        If String.IsNullOrEmpty(templateId) Then
                GlobalStatusStrip.ShowWarning("模板ID不能为空")
                Return
            End If

            ' docx映射卡片：直接打开关联的.docx文件预览
            If templateId.StartsWith("docx_") Then
                Dim mappingId = templateId.Substring(5)
                Dim mapping = SemanticMappingManager.Instance.GetMappingById(mappingId)
                If mapping Is Nothing Then
                    GlobalStatusStrip.ShowWarning("语义映射不存在")
                    Return
                End If
                If String.IsNullOrEmpty(mapping.SourceFilePath) OrElse Not IO.File.Exists(mapping.SourceFilePath) Then
                    GlobalStatusStrip.ShowWarning("原始模板文件已丢失，请重新上传")
                    Return
                End If

                ' 直接用系统默认方式打开文档（新Word实例）
                Try
                    Process.Start(mapping.SourceFilePath)
                    GlobalStatusStrip.ShowInfo($"已打开模板文档预览: {mapping.Name}")
                Catch ex As Exception
                    GlobalStatusStrip.ShowWarning($"打开文档失败: {ex.Message}")
                End Try
                Return
            End If

            ' 常规模板预览：先保存为临时文件，然后用系统默认方式打开
            Dim template As ReformatTemplate = ReformatTemplateManager.Instance.GetTemplateById(templateId)
            If template Is Nothing Then
                GlobalStatusStrip.ShowWarning($"找不到ID为 {templateId} 的模板")
                Return
            End If

            Try
                ' 创建临时文件
                Dim tempPath = IO.Path.Combine(IO.Path.GetTempPath(), $"模板预览_{template.Name}_{DateTime.Now:yyyyMMddHHmmss}.docx")

                ' 使用当前Word实例创建临时文档
                Dim currentApp = Globals.ThisAddIn.Application
                Dim tempDoc = currentApp.Documents.Add()

                ApplyTemplateToDocument(tempDoc, template)

                ' 保存并关闭临时文档
                tempDoc.SaveAs2(tempPath)
                tempDoc.Close(SaveChanges:=False)

                ' 用系统默认方式打开（新Word实例）
                Process.Start(tempPath)

                GlobalStatusStrip.ShowInfo($"已打开模板预览: {template.Name}")
            Catch ex As Exception
                GlobalStatusStrip.ShowWarning($"预览模板失败: {ex.Message}")
                Debug.WriteLine($"HandlePreviewTemplateInWord 错误: {ex.Message}")
            End Try
    End Sub

    ''' <summary>
    ''' 将模板应用到Word文档（预览用）
    ''' </summary>
    Private Sub ApplyTemplateToDocument(doc As Microsoft.Office.Interop.Word.Document, template As ReformatTemplate)
        Try
            ' 添加预览标记标题
            Dim para = doc.Paragraphs.Add()
            para.Range.Text = $"【模板预览】{template.Name}"
            para.Range.Font.Size = 16
            para.Range.Font.Bold = 1
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            para.Range.InsertParagraphAfter()

            ' 添加分隔线
            para = doc.Paragraphs.Add()
            para.Range.Text = "─────────────────────────────"
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            para.Range.InsertParagraphAfter()

            ' 如果有布局配置，添加相应的示例内容
            If template.Layout IsNot Nothing AndAlso template.Layout.Elements IsNot Nothing Then
                For i As Integer = 0 To Math.Min(template.Layout.Elements.Count - 1, 5)
                    Dim element = template.Layout.Elements(i)

                    ' 添加新段落
                    para = doc.Paragraphs.Add()
                    para.Range.Text = If(String.IsNullOrEmpty(element.DefaultValue), $"【{element.Name}】示例内容", element.DefaultValue)

                    ' 安全地应用字体设置
                    If element.Font IsNot Nothing Then
                        If Not String.IsNullOrEmpty(element.Font.FontNameCN) Then
                            para.Range.Font.Name = element.Font.FontNameCN
                            para.Range.Font.NameFarEast = element.Font.FontNameCN
                        End If
                        If element.Font.FontSize > 0 Then
                            para.Range.Font.Size = CSng(element.Font.FontSize)
                        End If
                        para.Range.Font.Bold = If(element.Font.Bold, 1, 0)
                    End If

                    ' 安全地应用段落设置
                    If element.Paragraph IsNot Nothing Then
                        Select Case element.Paragraph.Alignment?.ToLower()
                            Case "center"
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            Case "right"
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                            Case "justify"
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                            Case Else
                                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                        End Select
                    End If

                    ' 应用颜色设置
                    If element.Color IsNot Nothing AndAlso Not String.IsNullOrEmpty(element.Color.FontColor) Then
                        Try
                            Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(element.Color.FontColor)
                            para.Range.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                        Catch ex As Exception
                            Debug.WriteLine($"应用版式元素颜色失败: {ex.Message}")
                        End Try
                    End If

                    para.Range.InsertParagraphAfter()
                Next
            End If

            ' 添加正文样式预览
            If template.BodyStyles IsNot Nothing AndAlso template.BodyStyles.Count > 0 Then
                para = doc.Paragraphs.Add()
                para.Range.Text = "─────正文样式预览─────"
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                para.Range.InsertParagraphAfter()

                For Each style In template.BodyStyles
                    para = doc.Paragraphs.Add()
                    para.Range.Text = $"【{style.RuleName}】这是正文样式的示例文本，用于展示排版效果。"

                    If style.Font IsNot Nothing Then
                        If Not String.IsNullOrEmpty(style.Font.FontNameCN) Then
                            para.Range.Font.Name = style.Font.FontNameCN
                            para.Range.Font.NameFarEast = style.Font.FontNameCN
                        End If
                        If style.Font.FontSize > 0 Then
                            para.Range.Font.Size = CSng(style.Font.FontSize)
                        End If
                        para.Range.Font.Bold = If(style.Font.Bold, 1, 0)
                    End If

                    If style.Paragraph IsNot Nothing Then
                        If style.Paragraph.FirstLineIndent > 0 AndAlso para.Range.Font.Size > 0 Then
                            para.Range.ParagraphFormat.FirstLineIndent = CSng(style.Paragraph.FirstLineIndent * para.Range.Font.Size)
                        End If
                    End If

                    ' 应用颜色设置
                    If style.Color IsNot Nothing AndAlso Not String.IsNullOrEmpty(style.Color.FontColor) Then
                        Try
                            Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(style.Color.FontColor)
                            para.Range.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                        Catch ex As Exception
                            Debug.WriteLine($"应用正文样式颜色失败: {ex.Message}")
                        End Try
                    End If

                    para.Range.InsertParagraphAfter()
                Next
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyTemplateToDocument 错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用模板进行排版（覆盖基类方法）- 语义标注流水线
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithTemplate(template As ReformatTemplate)
        Try
            ' 转换旧模板为SemanticStyleMapping
            Dim mapping = LegacyTemplateConverter.Convert(template)
            If mapping Is Nothing Then
                GlobalStatusStrip.ShowWarning("模板转换失败")
                Return
            End If

            Await StartSemanticReformatPipeline(mapping, template.Name)
        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithTemplate 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"排版失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 使用docx解析的SemanticStyleMapping直接排版（覆盖基类方法）
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithMapping(mapping As SemanticStyleMapping)
        Try
            Await StartSemanticReformatPipeline(mapping, mapping.Name)
        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithMapping 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"排版失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 共用语义排版流水线：收集段落 → 构建提示词 → 发送AI标注请求
    ''' </summary>
    Private Async Function StartSemanticReformatPipeline(mapping As SemanticStyleMapping, displayName As String) As Task
        Dim wordApp = Globals.ThisAddIn.Application
        Dim selText As String = String.Empty

        Try
            If wordApp IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                selText = If(wordApp.Selection.Range IsNot Nothing, wordApp.Selection.Range.Text, String.Empty)
            End If
        Catch
            selText = String.Empty
        End Try

        If String.IsNullOrWhiteSpace(selText) Then
            GlobalStatusStrip.ShowWarning("请先选中需要排版的文本内容。")
            Return
        End If

        Dim selRange = wordApp.Selection.Range

        ' 收集所有段落信息
        Dim allParagraphs As New List(Of Microsoft.Office.Interop.Word.Paragraph)()
        Dim paragraphStyles As New List(Of String)()
        Dim paragraphTypes As New List(Of String)()
        Dim paragraphTexts As New List(Of String)()

        For Each p As Microsoft.Office.Interop.Word.Paragraph In selRange.Paragraphs
            Dim paraText As String = If(p.Range.Text IsNot Nothing, p.Range.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)

            Dim paraType As String = "text"
            Try
                If p.Range.InlineShapes.Count > 0 Then
                    paraType = "image"
                ElseIf p.Range.Tables.Count > 0 Then
                    paraType = "table"
                ElseIf p.Range.OMaths.Count > 0 Then
                    paraType = "formula"
                End If
            Catch
            End Try

            If Not String.IsNullOrWhiteSpace(paraText) OrElse paraType <> "text" Then
                allParagraphs.Add(p)

                Dim styleName As String = ""
                Try
                    styleName = p.Style.NameLocal
                Catch
                    styleName = "正文"
                End Try
                paragraphStyles.Add(styleName)
                paragraphTypes.Add(paraType)
                paragraphTexts.Add(paraText)
            End If
        Next

        If allParagraphs.Count = 0 Then
            GlobalStatusStrip.ShowWarning("选中的内容没有有效段落。")
            Return
        End If

        ' 显示排版模式吸顶提示
        Await ExecuteJavaScriptAsyncJS("showReformatModeIndicator();")

        ' 退出模板选择模式（显示聊天区域供后续交互），但保留mapping上下文
        ExitReformatTemplateMode()

        ' 构建语义标注提示词
        Dim systemPrompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(mapping, paragraphTexts)

        ' 保存段落上下文
        SetReformatContext(allParagraphs.Cast(Of Object).ToList(), paragraphStyles, paragraphTypes, mapping)

        ' 发送语义标注请求
        Await Send("请使用「" & displayName & "」对选中内容进行语义标注。", systemPrompt, False, "semantic_reformat")

        GlobalStatusStrip.ShowInfo("正在使用「" & displayName & "」排版...")
    End Function

    ''' <summary>
    ''' 使用规范进行排版（覆盖基类方法）- 语义标注流水线
    ''' </summary>
    Protected Overrides Async Sub ApplyReformatWithStyleGuide(guide As StyleGuideResource)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            Dim selText As String = String.Empty

            Try
                If wordApp IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                    selText = If(wordApp.Selection.Range IsNot Nothing, wordApp.Selection.Range.Text, String.Empty)
                End If
            Catch
                selText = String.Empty
            End Try

            ' 必须先选中内容
            If String.IsNullOrWhiteSpace(selText) Then
                GlobalStatusStrip.ShowWarning("请先选中需要排版的文本内容。")
                Return
            End If

            Dim selRange = wordApp.Selection.Range

            ' 收集所有段落信息
            Dim allParagraphs As New List(Of Microsoft.Office.Interop.Word.Paragraph)()
            Dim paragraphStyles As New List(Of String)()
            Dim paragraphTypes As New List(Of String)()
            Dim paragraphTexts As New List(Of String)()

            For Each p As Microsoft.Office.Interop.Word.Paragraph In selRange.Paragraphs
                Dim paraText As String = If(p.Range.Text IsNot Nothing, p.Range.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)

                Dim paraType As String = "text"
                Try
                    If p.Range.InlineShapes.Count > 0 Then
                        paraType = "image"
                    ElseIf p.Range.Tables.Count > 0 Then
                        paraType = "table"
                    ElseIf p.Range.OMaths.Count > 0 Then
                        paraType = "formula"
                    End If
                Catch
                End Try

                If Not String.IsNullOrWhiteSpace(paraText) OrElse paraType <> "text" Then
                    allParagraphs.Add(p)

                    Dim styleName As String = ""
                    Try
                        styleName = p.Style.NameLocal
                    Catch
                        styleName = "正文"
                    End Try
                    paragraphStyles.Add(styleName)
                    paragraphTypes.Add(paraType)
                    paragraphTexts.Add(paraText)
                End If
            Next

            If allParagraphs.Count = 0 Then
                GlobalStatusStrip.ShowWarning("选中的内容没有有效段落。")
                Return
            End If

            ' 显示排版模式吸顶提示
            Await ExecuteJavaScriptAsyncJS("showReformatModeIndicator();")

            ' 成功开始处理，退出模板选择模式
            ExitReformatTemplateMode()

            ' 检查是否有缓存的SemanticStyleMapping
            Dim mapping = SemanticMappingManager.Instance.GetMappingBySourceId(guide.Id)

            If mapping IsNot Nothing Then
                ' 使用缓存的映射，直接进行语义标注
                Dim systemPrompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(mapping, paragraphTexts)
                SetReformatContext(allParagraphs.Cast(Of Object).ToList(), paragraphStyles, paragraphTypes, mapping)

                ' 开始Undo快照
                Try
                    wordApp.UndoRecord.StartCustomRecord("AI排版")
                Catch ex As Exception
                    Debug.WriteLine($"StartCustomRecord 失败: {ex.Message}")
                End Try

                Await Send("请使用「" & guide.Name & "」排版规范对选中内容进行语义标注。", systemPrompt, False, "semantic_reformat")
            Else
                ' 没有缓存，需要先让AI将规范转为映射，再进行语义标注
                ' 使用两步模式：先转换规范，再标注
                Dim conversionPrompt = StyleGuideConverter.BuildConversionPrompt(guide.GuideContent)

                ' 保存段落信息（mapping暂时为Nothing，在收到转换结果后设置）
                SetReformatContext(allParagraphs.Cast(Of Object).ToList(), paragraphStyles, paragraphTypes, Nothing)

                ' 发送规范转换请求（mode为styleguide_convert）
                Await Send("请解析「" & guide.Name & "」排版规范并提取格式参数。", conversionPrompt, False, "styleguide_convert")
            End If

            GlobalStatusStrip.ShowInfo("正在使用「" & guide.Name & "」规范排版...")

        Catch ex As Exception
            Debug.WriteLine($"ApplyReformatWithStyleGuide 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"排版失败: {ex.Message}")
        End Try
    End Sub


    Public Sub New()
        ' 此调用是设计师所必需的。
        InitializeComponent()

        ' 确保WebView2控件可以正常交互
        ChatBrowser.BringToFront()

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)

        ' 订阅Word的SelectionChange 事件
        ' 帮我补全word选择的内容事件
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf GetSelectionContent
    End Sub

    '获取选中的内容
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' 转换为 Word.Selection 对象
            Dim selection = TryCast(Globals.ThisAddIn.Application.Selection, Microsoft.Office.Interop.Word.Selection)
            If selection Is Nothing Then
                Return
            End If

            ' 检查是否有实际选中内容（通过比较Start和End位置）
            If selection.Start = selection.End Then
                ' 光标在单一位置，没有选中内容，清除之前的选中显示
                ClearSelectedContentBySheetName("Word文档")
                Return
            End If

            ' 获取选中内容的详细信息
            Dim content As String = String.Empty

            ' 检查是否选中了表格
            If selection.Tables.Count > 0 Then
                ' 如果选中的是表格
                Dim table = selection.Tables(1)
                Dim sb As New StringBuilder()

                ' 遍历表格内容
                For row As Integer = 1 To table.Rows.Count
                    For col As Integer = 1 To table.Columns.Count
                        sb.Append(table.Cell(row, col).Range.Text.TrimEnd(ChrW(13), ChrW(7)))
                        If col < table.Columns.Count Then sb.Append(vbTab)
                    Next
                    sb.AppendLine()
                Next
                content = sb.ToString()

            ElseIf selection.InlineShapes.Count > 0 OrElse selection.ShapeRange.Count > 0 Then
                ' 如果选中的是图片或形状
                content = "[图片或形状]"
            Else
                ' 普通文本选择
                content = selection.Text
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' 添加到选中内容列表
                AddSelectedContentItem(
                "Word文档",  ' 使用文档名称作为标识
                If(selection.Tables.Count > 0,
                   "[表格内容]",
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
            )
            Else
                ClearSelectedContentBySheetName("Word文档")
            End If

        Catch ex As Exception
            Debug.WriteLine($"获取Word选中内容时出错: {ex.Message}")
        End Try
    End Sub


    ' 获取选中内容的详细信息
    Private Function GetSelectionDetails(selection As Microsoft.Office.Interop.Word.Selection) As String
        Dim details As New StringBuilder()

        ' 添加基本信息
        details.AppendLine($"开始位置: {selection.Start}")
        details.AppendLine($"结束位置: {selection.End}")
        details.AppendLine($"字符数: {selection.Characters.Count}")

        ' 如果是表格，添加表格信息
        If selection.Tables.Count > 0 Then
            Dim table = selection.Tables(1)
            details.AppendLine($"表格大小: {table.Rows.Count}行 x {table.Columns.Count}列")
        End If

        Return details.ToString()
    End Function

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


    ' 返回应用信息
    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Word", OfficeApplicationType.Word)
    End Function

    ' 返回Office应用类型
    Protected Overrides Function GetOfficeAppType() As String
        Return "Word"
    End Function

    ''' <summary>
    ''' 获取样式预览回调（Word实现：将样式应用到选中的文本）
    ''' </summary>
    Protected Overrides Function GetStylePreviewCallback() As PreviewStyleCallback
        Return AddressOf ApplyStylePreviewToSelection
    End Function

    ''' <summary>
    ''' 显示模板编辑器面板（使用 CustomTaskPane）
    ''' </summary>
    Protected Overrides Function ShowTemplateEditorPane(template As ReformatTemplate) As Boolean
        Try
            Globals.ThisAddIn.ShowTemplateEditorTaskPane(template)
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ShowTemplateEditorPane 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 将样式预览应用到Word中选中的文本（Public供外部调用）
    ''' </summary>
    Public Sub ApplyStylePreviewToSelection(fontName As String, fontSize As Double, bold As Boolean, alignment As String, firstLineIndent As Double, lineSpacing As Double)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            Dim doc = wordApp.ActiveDocument
            If doc Is Nothing Then Return

            Dim selection = wordApp.Selection
            Dim previewRange As Word.Range = Nothing

            ' 检查是否有有效的选中内容
            Dim hasValidSelection = selection IsNot Nothing AndAlso
                                   selection.Type <> Word.WdSelectionType.wdNoSelection AndAlso
                                   Not String.IsNullOrWhiteSpace(selection.Text?.Replace(vbCr, "").Replace(vbLf, ""))

            If hasValidSelection Then
                ' 使用选中的文本
                previewRange = selection.Range
            Else
                ' 没有选中文本，查找或创建预览段落
                Dim previewMarker = "【样式预览】"
                Dim found = False

                ' 查找已有的预览段落
                For Each para As Word.Paragraph In doc.Paragraphs
                    If para.Range.Text.StartsWith(previewMarker) Then
                        previewRange = para.Range
                        found = True
                        Exit For
                    End If
                Next

                ' 如果没有找到，在文档末尾创建预览段落
                If Not found Then
                    Dim endRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                    endRange.InsertParagraphAfter()
                    endRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                    endRange.Text = previewMarker & "这是用于预览样式效果的示例文本，您可以在此查看字体、字号、对齐等效果。"
                    previewRange = doc.Paragraphs.Last.Range
                End If
            End If

            If previewRange Is Nothing Then Return

            ' 应用字体
            If Not String.IsNullOrEmpty(fontName) Then
                previewRange.Font.Name = fontName
                previewRange.Font.NameFarEast = fontName
            End If

            ' 应用字号（磅值）
            If fontSize > 0 Then
                previewRange.Font.Size = CSng(fontSize)
            End If

            ' 应用加粗
            previewRange.Font.Bold = If(bold, 1, 0)

            ' 应用对齐方式
            Select Case alignment?.ToLower()
                Case "left"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                Case "center"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                Case "right"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                Case "justify"
                    previewRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
            End Select

            ' 应用首行缩进（字符数转换为磅值）
            If firstLineIndent > 0 AndAlso fontSize > 0 Then
                previewRange.ParagraphFormat.FirstLineIndent = CSng(firstLineIndent * fontSize)
            Else
                previewRange.ParagraphFormat.FirstLineIndent = 0
            End If

            ' 应用行距
            If lineSpacing > 0 Then
                previewRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple
                previewRange.ParagraphFormat.LineSpacing = CSng(lineSpacing * 12)
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyStylePreviewToSelection 出错: {ex.Message}")
        End Try
    End Sub

    ' 返回 Word Application 对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function

    ' 返回当前文档的 VBProject（可能为 Nothing）
    Protected Overrides Function GetVBProject() As VBProject
        Try
            Return Globals.ThisAddIn.Application.ActiveDocument.VBProject
        Catch
            Return Nothing
        End Try
    End Function

    ' 预览运行：展示代码并询问是否继续（返回 True 执行）
    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
        If Not preview Then Return True
        Dim prompt As String = "预览将要执行的 VBA 代码，是否继续？" & vbCrLf & "----" & vbCrLf & vbaCode
        Return (MessageBox.Show(prompt, "VBA 预览", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)
    End Function

    ' 真正执行宏（通过 Application.Run 调用模块.过程）
    Protected Overrides Function RunCode(vbaCode As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(vbaCode)
        Catch ex As Exception
            MessageBox.Show("执行宏失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

    ' 将要发送到 LLM 的消息委托到底层 Send 方法（异步）
    Protected Overrides Sub SendChatMessage(message As String)
        Task.Run(Async Function()
                     Await Send(message, "", True, "")
                 End Function)
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息（重写基类方法）
    ''' </summary>
    Protected Overrides Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"Word使用意图优化提示词: {intent.IntentType}, 置信度: {intent.Confidence:F2}")

            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "")
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

    ' 解析 Word 文件为文本（用于 file 引用）
    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' 创建一个新的Word应用程序实例（避免影响当前文档）
            Dim wordApp As New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
            Try
                doc = wordApp.Documents.Open(FileName:=filePath, ReadOnly:=True, Visible:=False)
                Dim contentBuilder As New StringBuilder()

                contentBuilder.AppendLine($"文件: {Path.GetFileName(filePath)}")
                contentBuilder.AppendLine($"共 {doc.Paragraphs.Count} 个段落")
                contentBuilder.AppendLine()

                ' 限制处理的段落数量
                Dim maxParagraphs As Integer = Math.Min(doc.Paragraphs.Count, 50)
                Dim paraIndex As Integer = 0

                For Each para As Microsoft.Office.Interop.Word.Paragraph In doc.Paragraphs
                    paraIndex += 1
                    If paraIndex > maxParagraphs Then Exit For

                    Dim text As String = para.Range.Text.Trim()
                    If Not String.IsNullOrEmpty(text) AndAlso text <> vbCr Then
                        ' 获取段落样式
                        Dim styleName As String = ""
                        Try
                            styleName = para.Style.NameLocal
                        Catch
                        End Try

                        ' 判断是否是标题
                        Dim prefix As String = $"段落{paraIndex}"
                        If styleName.Contains("标题") OrElse styleName.ToLower().Contains("heading") Then
                            prefix = $"[{styleName}]"
                        End If

                        contentBuilder.AppendLine($"{prefix}: {text}")
                    End If
                Next

                ' 处理表格
                If doc.Tables.Count > 0 Then
                    contentBuilder.AppendLine()
                    contentBuilder.AppendLine($"=== 文档包含 {doc.Tables.Count} 个表格 ===")

                    Dim tableIndex As Integer = 0
                    For Each tbl As Microsoft.Office.Interop.Word.Table In doc.Tables
                        tableIndex += 1
                        If tableIndex > 5 Then Exit For ' 限制表格数量

                        contentBuilder.AppendLine($"表格 {tableIndex}: {tbl.Rows.Count}行×{tbl.Columns.Count}列")

                        ' 读取表格前几行
                        Dim maxRows = Math.Min(tbl.Rows.Count, 5)
                        For rowIdx = 1 To maxRows
                            Dim rowContent As New StringBuilder("  ")
                            For colIdx = 1 To tbl.Columns.Count
                                Try
                                    Dim cellText = tbl.Cell(rowIdx, colIdx).Range.Text.Trim()
                                    cellText = cellText.Replace(vbCr, "").Replace(Chr(7), "")
                                    If cellText.Length > 20 Then cellText = cellText.Substring(0, 17) & "..."
                                    rowContent.Append(cellText & " | ")
                                Catch
                                End Try
                            Next
                            contentBuilder.AppendLine(rowContent.ToString().TrimEnd(" |".ToCharArray()))
                        Next
                        contentBuilder.AppendLine()
                    Next
                End If

                If doc.Paragraphs.Count > maxParagraphs Then
                    contentBuilder.AppendLine()
                    contentBuilder.AppendLine($"... 共 {doc.Paragraphs.Count} 个段落，仅显示前 {maxParagraphs} 个")
                End If

                Return New FileContentResult With {
                    .FileName = Path.GetFileName(filePath),
                    .FileType = "Word",
                    .ParsedContent = contentBuilder.ToString(),
                    .RawData = Nothing
                }

            Finally
                If doc IsNot Nothing Then
                    doc.Close(False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc)
                End If
                wordApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            Debug.WriteLine($"解析Word文件时出错: {ex.Message}")
            Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Word",
                .ParsedContent = $"[解析Word文件时出错: {ex.Message}]"
            }
        End Try
    End Function

    ' 返回当前文档所在目录（未保存返回空字符串）
    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            Dim p = Globals.ThisAddIn.Application.ActiveDocument.Path
            If String.IsNullOrEmpty(p) Then Return String.Empty
            Return p
        Catch
            Return String.Empty
        End Try
    End Function

    ' 将当前选区内容附加到提示，并记录 PendingSelectionInfo 供写回使用
    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            Dim txt As String = If(sel IsNot Nothing AndAlso sel.Range IsNot Nothing, sel.Range.Text, String.Empty)

            Dim info As New SelectionInfo()
            info.DocumentPath = If(Globals.ThisAddIn.Application.ActiveDocument.Path, "")
            info.SelectedText = txt
            Try
                info.StartPos = sel.Range.Start
                info.EndPos = sel.Range.End
            Catch
                info.StartPos = 0
                info.EndPos = 0
            End Try

            PendingSelectionInfo = info

            If String.IsNullOrWhiteSpace(txt) Then
                Return message
            Else
                Return message & vbCrLf & vbCrLf & txt
            End If
        Catch
            Return message
        End Try
    End Function


    ' 修订、审阅功能（简化版：使用段落索引定位）
    Protected Overrides Sub HandleApplyRevisionSegment(jsonDoc As JObject)
        Try
            ' 期望收到字段： uuid, paraIndex, original, corrected
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim paraIndex As Integer = If(jsonDoc("paraIndex") IsNot Nothing, CInt(jsonDoc("paraIndex")), -1)
            Dim original As String = If(jsonDoc("original") IsNot Nothing, jsonDoc("original").ToString(), String.Empty)
            Dim corrected As String = If(jsonDoc("corrected") IsNot Nothing, jsonDoc("corrected").ToString(), String.Empty)

            If paraIndex < 0 Then
                GlobalStatusStrip.ShowWarning("缺少 paraIndex 参数")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("校对功能仅在 Word 环境下支持")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try
            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象")
                Return
            End If

            Dim doc = officeApp.ActiveDocument
            Dim selRange = officeApp.Selection.Range

            ' 使用选中范围内的段落索引定位
            If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(selRange.Text) Then
                GlobalStatusStrip.ShowWarning("请先选中需要校对的内容")
                Return
            End If

            ' 获取选中范围内的段落
            Dim paragraphs = selRange.Paragraphs
            If paraIndex >= paragraphs.Count Then
                GlobalStatusStrip.ShowWarning($"段落索引 {paraIndex} 超出范围")
                Return
            End If

            ' 定位目标段落（段落索引从1开始）
            Dim targetPara = paragraphs(paraIndex + 1)
            Dim targetRange = targetPara.Range

            ' 在目标段落中查找并替换原文
            If Not String.IsNullOrEmpty(original) Then
                Dim paraText As String = targetRange.Text
                Dim startPos As Integer = paraText.IndexOf(original, StringComparison.Ordinal)
                If startPos >= 0 Then
                    ' 创建精确的替换范围
                    Dim replaceRange = doc.Range(targetRange.Start + startPos, targetRange.Start + startPos + original.Length)

                    ' 开启审阅模式
                    Try
                        doc.TrackRevisions = True
                    Catch
                    End Try

                    ' 执行替换
                    replaceRange.Text = corrected
                    GlobalStatusStrip.ShowInfo($"已替换段落 {paraIndex} 中的内容（审阅模式）")
                Else
                    GlobalStatusStrip.ShowWarning($"在段落 {paraIndex} 中未找到原文：{original}")
                End If
            Else
                GlobalStatusStrip.ShowWarning("缺少原文内容")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionSegment 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("校对写回异常: " & ex.Message)
        End Try
    End Sub

    ' 新增：在 Range 插入 WordProcessingML（OpenXML）片段
    Private Function InsertOpenXmlIntoRange(openXml As String, targetRange As Object) As Boolean
        Try
            If String.IsNullOrEmpty(openXml) OrElse targetRange Is Nothing Then Return False

            ' Word Range.InsertXML 需要完整的 WordProcessingML 文档结构
            ' 如果传入的只是片段（如 <w:p>），需要包装成完整结构
            Dim wrappedXml As String = WrapXmlFragment(openXml)

            Try
                Debug.Print("InsertOpenXmlIntoRange: " & wrappedXml.Substring(0, Math.Min(500, wrappedXml.Length)))
                targetRange.InsertXML(wrappedXml)
                Return True
            Catch ex As Exception
                Debug.WriteLine("InsertOpenXmlIntoRange: InsertXML 失败: " & ex.Message)
                ' 回退：尝试直接设置文本
                Try
                    Dim plainText As String = ExtractTextFromXml(openXml)
                    If Not String.IsNullOrEmpty(plainText) Then
                        targetRange.Text = plainText
                        Return True
                    End If
                Catch
                End Try
                Return False
            End Try
        Catch ex As Exception
            Debug.WriteLine("InsertOpenXmlIntoRange 出错: " & ex.Message)
            Return False
        End Try
    End Function

    ' 将 OpenXML 片段包装成完整的 WordProcessingML 文档
    Private Function WrapXmlFragment(fragment As String) As String
        If String.IsNullOrEmpty(fragment) Then Return String.Empty

        ' 检查是否已经是完整的文档结构
        If fragment.Contains("<w:document") OrElse fragment.Contains("<pkg:package") Then
            Return fragment
        End If

        ' 定义命名空间
        Const wNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        Const rNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        ' 包装成完整的 WordProcessingML 文档
        Dim sb As New StringBuilder()
        sb.Append("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
        sb.Append($"<w:document xmlns:w=""{wNs}"" xmlns:r=""{rNs}"">")
        sb.Append("<w:body>")
        sb.Append(fragment)
        sb.Append("</w:body>")
        sb.Append("</w:document>")

        Return sb.ToString()
    End Function

    ' 从 OpenXML 片段中提取纯文本（作为回退方案）
    Private Function ExtractTextFromXml(xml As String) As String
        Try
            If String.IsNullOrEmpty(xml) Then Return String.Empty
            ' 简单的正则提取 <w:t> 标签内容
            Dim matches = System.Text.RegularExpressions.Regex.Matches(xml, "<w:t[^>]*>([^<]*)</w:t>")
            Dim result As New StringBuilder()
            For Each m As System.Text.RegularExpressions.Match In matches
                If m.Groups.Count > 1 Then
                    result.Append(m.Groups(1).Value)
                End If
            Next
            Return result.ToString()
        Catch
            Return String.Empty
        End Try
    End Function

    ' applyRevision
    Protected Overrides Sub HandleApplyRevisionAll(jsonDoc As JObject)
        Try
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim newContent As String = If(jsonDoc("newContent") IsNot Nothing, jsonDoc("newContent").ToString(), String.Empty)

            If String.IsNullOrWhiteSpace(newContent) Then
                GlobalStatusStrip.ShowWarning("没有接收到写回的新内容")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("写回操作仅在 Word 环境下支持（默认实现）")
                Return
            End If

            ' 使用 GetOfficeApplicationObject 获取宿主 Word Application 对象（子类需实现）
            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try

            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象，写回失败")
                Return
            End If

            Try
                ' 在审阅模式下写回：先开启 TrackRevisions，再执行删除/插入以产生审阅记录
                Dim doc = officeApp.ActiveDocument
                Dim selRange = officeApp.Selection.Range
                Dim useRange = Nothing

                If selRange IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(selRange.Text) Then
                    useRange = selRange
                Else
                    useRange = doc.Content
                End If

                ' 开启审阅模式
                Try
                    doc.TrackRevisions = True
                Catch
                    ' 忽略，如果宿主不支持
                End Try

                ' 删除原文本（此操作会被记录为删除），然后插入新文本（被记录为插入）
                useRange.Delete()
                useRange.InsertAfter(newContent)

                GlobalStatusStrip.ShowInfo("写回已完成（审阅模式）。请在审阅中查看修改。")
            Catch ex As Exception
                Debug.WriteLine("写回失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("写回失败: " & ex.Message)
            End Try

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionAll 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("写回操作异常")
        End Try
    End Sub

    Protected Overrides Sub HandleApplyRevisionAccept(jsonDoc As JObject)
        Try
            ' 期望 { type:'applyRevisionAccept', responseUuid:..., globalIndex: n }
            Dim responseUuid As String = If(jsonDoc("responseUuid") IsNot Nothing, jsonDoc("responseUuid").ToString(), If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty))
            Dim globalIndex As Integer = If(jsonDoc("globalIndex") IsNot Nothing, CInt(jsonDoc("globalIndex")), -1)

            If globalIndex < 0 Then
                GlobalStatusStrip.ShowWarning("applyRevisionAccept: 缺少 globalIndex")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("接受单个修订仅在 Word 环境下支持（默认实现）")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try

            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象，接受修订失败")
                Return
            End If

            Try
                Dim doc = officeApp.ActiveDocument
                ' Word Revisions 集合是 1 基的；尝试保护性调用
                If globalIndex >= 1 And globalIndex <= doc.Revisions.Count Then
                    doc.Revisions(globalIndex).Accept()
                    GlobalStatusStrip.ShowInfo($"已接受修订 #{globalIndex}")
                Else
                    GlobalStatusStrip.ShowWarning("指定的修订索引超出范围或不存在")
                End If
            Catch ex As Exception
                Debug.WriteLine("接受修订失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("接受修订失败: " & ex.Message)
            End Try
        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionAccept 出错: {ex.Message}")
        End Try
    End Sub

    Protected Overrides Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)

        ' 如果此次会话绑定了选区信息，则发送对比预览（原文 vs AI 输出）
        Try
            ' 使用 response->request 的映射查找对应的选区信息（修正原有逻辑中使用 _finalUuid 直接查找的错误）
            Dim requestId As String = Nothing
            If _responseToRequestMap.ContainsKey(_finalUuid) Then
                requestId = _responseToRequestMap(_finalUuid)
            End If

            Dim mode As String = ""
            If _responseModeMap.ContainsKey(_finalUuid) Then
                mode = _responseModeMap(_finalUuid)
            End If

            ' 语义标注排版模式：AI返回标注JSON，由渲染引擎确定性应用格式
            If String.Equals(mode, "semantic_reformat", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    ApplySemanticTaggingResult(aiText)
                Catch ex As Exception
                    Debug.WriteLine("语义标注排版处理失败: " & ex.Message)
                End Try
                MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
                Return
            End If

            ' 规范转换模式：AI返回规范的结构化JSON，解析后进入语义标注阶段
            If String.Equals(mode, "styleguide_convert", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    HandleStyleGuideConversionResult(aiText)
                Catch ex As Exception
                    Debug.WriteLine("规范转换处理失败: " & ex.Message)
                End Try
                MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
                Return
            End If

            ' 如果是排版重构动作，则触发 showComparison
            If _responseSelectionMap.ContainsKey(_finalUuid) AndAlso String.Equals(mode, "reformat", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim selInfo = _responseSelectionMap(_finalUuid)
                    Dim originalText As String = If(selInfo?.SelectedText, "")
                    Dim aiFinal As String = allPlainMarkdownBuffer.ToString()

                    Dim js As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(aiFinal)});"
                    ExecuteJavaScriptAsyncJS(js)
                Catch ex As Exception
                    Debug.WriteLine("尝试解析并发送 comparison 时出错: " & ex.Message)
                End Try
            End If

            ' 如果是审阅修订动作，解析并展示 revisions（前端会处理）
            If String.Equals(mode, "proofread", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    Dim revisions As JArray = TryExtractJsonArrayFromText(aiText)
                    If revisions IsNot Nothing AndAlso revisions.Count > 0 Then
                        _revisionsMap(_finalUuid) = revisions
                        Dim jsRev As String = $"showRevisions('{_finalUuid}', {revisions.ToString(Formatting.None)});"
                        ExecuteJavaScriptAsyncJS(jsRev)
                    End If
                Catch ex As Exception
                    Debug.WriteLine("尝试解析并发送 revisions 时出错: " & ex.Message)
                End Try
            End If

            ' 解析并发送文档计划或预览 HTML 给前端，作为唯一内容
            If String.Equals(mode, "documentPlan", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(mode, "previewHtml", StringComparison.OrdinalIgnoreCase) Then
                Try
                    ' 尝试直接解析 JSON 对象（可能是 documentPlan 数组 / previewHtml / previewHtmlMap / 单个 planItem）
                    Dim rawText As String = allPlainMarkdownBuffer.ToString()
                    Dim jsonObj As JObject = TryExtractJsonObjectFromText(rawText)

                    If jsonObj IsNot Nothing Then
                        ' 如果后端/模型仅返回单个 planItem（键为 planItem），将其包装为 documentPlan 数组以便前端统一处理
                        Dim sendObj As JObject = Nothing
                        If jsonObj("planItem") IsNot Nothing Then
                            Dim arr As New JArray()
                            arr.Add(jsonObj("planItem"))
                            sendObj = New JObject()
                            sendObj("documentPlan") = arr
                            ' 若同时包含 previewHtmlMap，保留之
                            If jsonObj("previewHtmlMap") IsNot Nothing Then
                                sendObj("previewHtmlMap") = jsonObj("previewHtmlMap")
                            End If
                            ' 若 planItem 自身已包含 previewHtmlMap（极少见），合并也可按需处理
                        Else
                            ' 直接使用解析到的对象：可能含 documentPlan、previewHtml、previewHtmlMap 等
                            sendObj = jsonObj
                        End If

                        ' 获取原始选区文本（若存在）
                        Dim originalText As String = ""
                        If _responseSelectionMap.ContainsKey(_finalUuid) Then
                            Dim selInfo = _responseSelectionMap(_finalUuid)
                            originalText = If(selInfo?.SelectedText, "")
                        End If

                        ' 将整个对象序列化为字符串后传给前端的 showComparison，前端会解析 previewHtmlMap 或 documentPlan
                        Dim payload As String = sendObj.ToString(Formatting.None)
                        Dim jsPlan As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(payload)});"
                        ExecuteJavaScriptAsyncJS(jsPlan)
                    Else
                        ' 退回尝试解析为 JSON 数组（旧版可能只返回数组）
                        Dim planArr As JArray = TryExtractJsonArrayFromText(rawText)
                        If planArr IsNot Nothing AndAlso planArr.Count > 0 Then
                            Dim wrapper As New JObject()
                            wrapper("documentPlan") = planArr

                            Dim originalText As String = ""
                            If _responseSelectionMap.ContainsKey(_finalUuid) Then
                                Dim selInfo = _responseSelectionMap(_finalUuid)
                                originalText = If(selInfo?.SelectedText, "")
                            End If

                            Dim payload As String = wrapper.ToString(Formatting.None)
                            Dim jsPlan As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(payload)});"
                            ExecuteJavaScriptAsyncJS(jsPlan)
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine("处理 documentPlan/previewHtml 失败: " & ex.Message)
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"发送对比预览失败: {ex.Message}")
        End Try

        ' 调用基类处理续写模式
        MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
    End Sub

    ' 提取文本中第一个 JSON 数组（严格数组格式），返回 JArray 或 Nothing
    Private Function TryExtractJsonArrayFromText(text As String) As JArray
        Try
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            ' 尝试用正则抽取第一个 [...] 数组块（Singleline 允许跨行）
            Dim m As Match = Regex.Match(text, "\[.*\]", RegexOptions.Singleline)
            If m.Success Then
                Dim jsonCandidate As String = m.Value.Trim()
                ' 验证并解析为 JArray
                Try
                    Dim arr As JArray = JArray.Parse(jsonCandidate)
                    Return arr
                Catch ex As Exception
                    Debug.WriteLine("解析 JSON 数组失败: " & ex.Message)
                    Return Nothing
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("TryExtractJsonArrayFromText 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' 提取文本中第一个 JSON 对象（如 {"documentPlan":..., "previewHtml":...}），返回 JObject 或 Nothing
    Private Function TryExtractJsonObjectFromText(text As String) As JObject
        Try
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            ' 尝试用正则抽取第一个 { ... } 对象块（Singleline 允许跨行）
            Dim m As Match = Regex.Match(text, "\{[\s\S]*\}", RegexOptions.Singleline)
            If m.Success Then
                Dim jsonCandidate As String = m.Value.Trim()
                ' 验证并解析为 JObject
                Try
                    Dim obj As JObject = JObject.Parse(jsonCandidate)
                    Return obj
                Catch ex As Exception
                    Debug.WriteLine("解析 JSON 对象失败: " & ex.Message)
                    Return Nothing
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("TryExtractJsonObjectFromText 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' 排版功能（支持语义标注模式、规则模式和旧的逐段落模式）
    Protected Overrides Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
        Try
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)

            ' 语义标注模式（新流水线）：有tags字段
            If jsonDoc("tags") IsNot Nothing AndAlso jsonDoc("tags").Type = JTokenType.Array Then
                ApplySemanticTaggingResult(jsonDoc("tags").ToString(Newtonsoft.Json.Formatting.None))
                Return
            End If

            ' 检测是否为新的规则模式（有rules字段）
            If jsonDoc("rules") IsNot Nothing AndAlso jsonDoc("rules").Type = JTokenType.Array Then
                ApplyReformatRules(jsonDoc)
                Return
            End If

            ' 旧模式：逐段落格式化（保留向后兼容）
            Dim paraIndex As Integer = If(jsonDoc("paraIndex") IsNot Nothing, CInt(jsonDoc("paraIndex")), -1)
            Dim formatting As JObject = Nothing
            If jsonDoc("formatting") IsNot Nothing Then
                formatting = DirectCast(jsonDoc("formatting"), JObject)
            End If

            If paraIndex < 0 Then
                GlobalStatusStrip.ShowWarning("缺少 paraIndex 参数")
                Return
            End If

            If formatting Is Nothing Then
                GlobalStatusStrip.ShowWarning("缺少 formatting 参数")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("排版功能仅在 Word 环境下支持")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try
            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象")
                Return
            End If

            Dim doc = officeApp.ActiveDocument
            Dim selRange = officeApp.Selection.Range

            If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(selRange.Text) Then
                GlobalStatusStrip.ShowWarning("请先选中需要排版的内容")
                Return
            End If

            ' 获取选中范围内的段落
            Dim paragraphs = selRange.Paragraphs
            If paraIndex >= paragraphs.Count Then
                GlobalStatusStrip.ShowWarning($"段落索引 {paraIndex} 超出范围")
                Return
            End If

            ' 定位目标段落
            Dim targetPara = paragraphs(paraIndex + 1)
            Dim targetRange = targetPara.Range

            ' 使用Word对象模型应用格式化
            Try
                ApplyFormattingToRange(targetRange, formatting)
                GlobalStatusStrip.ShowInfo($"已应用段落 {paraIndex} 的排版")
            Catch ex As Exception
                Debug.WriteLine("排版写回失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("排版写回失败: " & ex.Message)
            End Try

        Catch ex As Exception
            Debug.WriteLine("HandleApplyDocumentPlanItem 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("排版应用出错: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 应用规则模式的排版（优化版：减少token消耗）
    ''' </summary>
    Private Sub ApplyReformatRules(jsonDoc As JObject)
        Try
            Dim rules = jsonDoc("rules").ToObject(Of List(Of JObject))()
            Dim sampleClassification = jsonDoc("sampleClassification")?.ToObject(Of List(Of JObject))()

            If rules Is Nothing OrElse rules.Count = 0 Then
                GlobalStatusStrip.ShowWarning("没有收到有效的排版规则")
                Return
            End If

            ' 构建规则字典
            Dim ruleDict As New Dictionary(Of String, JObject)()
            For Each rule In rules
                Dim ruleType = rule("type")?.ToString()
                If Not String.IsNullOrEmpty(ruleType) AndAlso rule("formatting") IsNot Nothing Then
                    ruleDict(ruleType) = DirectCast(rule("formatting"), JObject)
                End If
            Next

            ' 如果没有保存的段落上下文，使用当前选中内容
            If _reformatParagraphs Is Nothing OrElse _reformatParagraphs.Count = 0 Then
                GlobalStatusStrip.ShowWarning("排版上下文丢失，请重新选择内容并排版")
                Return
            End If

            ' 基于样本分类推断所有段落的规则
            Dim sampleRuleMap As New Dictionary(Of Integer, String)()
            If sampleClassification IsNot Nothing Then
                For Each sc In sampleClassification
                    Dim idx = sc("sampleIndex")?.ToObject(Of Integer)()
                    Dim appliedRule = sc("appliedRule")?.ToString()
                    If idx IsNot Nothing AndAlso Not String.IsNullOrEmpty(appliedRule) Then
                        sampleRuleMap(idx) = appliedRule
                    End If
                Next
            End If

            ' 应用格式到所有段落
            Dim appliedCount As Integer = 0
            Dim skippedCount As Integer = 0
            Dim defaultRule As String = If(ruleDict.ContainsKey("body"), "body", ruleDict.Keys.FirstOrDefault())

            For i As Integer = 0 To _reformatParagraphs.Count - 1
                Try
                    ' 检查段落类型，跳过非文本元素
                    Dim paraType As String = "text"
                    If _reformatTypes IsNot Nothing AndAlso i < _reformatTypes.Count Then
                        paraType = _reformatTypes(i)
                    End If

                    If paraType <> "text" Then
                        ' 跳过图片、表格、公式等非文本元素
                        skippedCount += 1
                        Continue For
                    End If

                    Dim para = _reformatParagraphs(i)
                    Dim styleName = If(i < _reformatStyles.Count, _reformatStyles(i), "")

                    ' 确定使用哪个规则
                    Dim ruleToApply As String = defaultRule

                    ' 先检查是否有样本分类指定
                    If sampleRuleMap.ContainsKey(i) Then
                        ruleToApply = sampleRuleMap(i)
                    Else
                        ' 基于样式名推断规则
                        If styleName.Contains("标题") OrElse styleName.ToLower().Contains("heading") Then
                            ' 找到第一个标题类规则
                            For Each key In ruleDict.Keys
                                If key.ToLower().Contains("title") OrElse key.ToLower().Contains("heading") Then
                                    ruleToApply = key
                                    Exit For
                                End If
                            Next
                        End If
                    End If

                    ' 应用规则
                    If ruleDict.ContainsKey(ruleToApply) Then
                        Dim formatting = ruleDict(ruleToApply)
                        ApplyFormattingToRange(para.Range, formatting)
                        appliedCount += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"应用段落 {i} 格式失败: " & ex.Message)
                End Try
            Next

            ' 清理上下文
            _reformatParagraphs = Nothing
            _reformatStyles = Nothing
            _reformatTypes = Nothing

            Dim resultMsg = $"排版完成，共处理 {appliedCount} 个文本段落"
            If skippedCount > 0 Then
                resultMsg &= $"，跳过 {skippedCount} 个特殊元素"
            End If
            GlobalStatusStrip.ShowInfo(resultMsg)

        Catch ex As Exception
            Debug.WriteLine("ApplyReformatRules 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("应用排版规则出错: " & ex.Message)
        End Try
    End Sub

    ' 使用Word对象模型应用格式化属性
    Private Sub ApplyFormattingToRange(targetRange As Object, formatting As JObject)
        Try
            ' 字体名称（中文）
            If formatting("fontNameCN") IsNot Nothing Then
                Dim fontNameCN As String = formatting("fontNameCN").ToString()
                If Not String.IsNullOrEmpty(fontNameCN) Then
                    Try
                        targetRange.Font.NameFarEast = fontNameCN
                    Catch
                        ' 某些 Word 版本可能不支持 NameFarEast
                    End Try
                End If
            End If

            ' 字体名称（英文/西文）
            If formatting("fontNameEN") IsNot Nothing Then
                Dim fontNameEN As String = formatting("fontNameEN").ToString()
                If Not String.IsNullOrEmpty(fontNameEN) Then
                    Try
                        targetRange.Font.Name = fontNameEN
                    Catch
                    End Try
                End If
            End If

            ' 字号
            If formatting("fontSize") IsNot Nothing Then
                Dim fontSize As Single = 0
                Single.TryParse(formatting("fontSize").ToString(), fontSize)
                If fontSize > 0 Then
                    Try
                        targetRange.Font.Size = fontSize
                    Catch
                    End Try
                End If
            End If

            ' 加粗
            If formatting("bold") IsNot Nothing Then
                Try
                    Dim bold As Boolean = formatting("bold").ToObject(Of Boolean)()
                    targetRange.Font.Bold = If(bold, -1, 0) ' Word: -1 = True, 0 = False
                Catch
                End Try
            End If

            ' 对齐方式
            If formatting("alignment") IsNot Nothing Then
                Dim alignment As String = formatting("alignment").ToString().ToLower()
                Try
                    Select Case alignment
                        Case "left"
                            targetRange.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                        Case "center"
                            targetRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                        Case "right"
                            targetRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
                        Case "justify", "justified"
                            targetRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
                    End Select
                Catch
                End Try
            End If

            ' 首行缩进（字符数）
            If formatting("firstLineIndent") IsNot Nothing Then
                Dim indent As Single = 0
                Single.TryParse(formatting("firstLineIndent").ToString(), indent)
                If indent > 0 Then
                    Try
                        ' CharacterUnitFirstLineIndent 以字符为单位
                        targetRange.ParagraphFormat.CharacterUnitFirstLineIndent = indent
                    Catch
                        ' 回退：使用磅值（1字符约=10.5磅 for 五号字）
                        Try
                            targetRange.ParagraphFormat.FirstLineIndent = indent * 10.5
                        Catch
                        End Try
                    End Try
                End If
            End If

            ' 行距
            If formatting("lineSpacing") IsNot Nothing Then
                Dim lineSpacing As Single = 0
                Single.TryParse(formatting("lineSpacing").ToString(), lineSpacing)
                If lineSpacing > 0 Then
                    Try
                        ' LineSpacingRule: 0=wdLineSpaceSingle, 1=wdLineSpace1pt5, 2=wdLineSpaceDouble, 5=wdLineSpaceMultiple
                        If lineSpacing = 1.0 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 0 ' wdLineSpaceSingle
                        ElseIf lineSpacing = 1.5 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 1 ' wdLineSpace1pt5
                        ElseIf lineSpacing = 2.0 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 2 ' wdLineSpaceDouble
                        Else
                            ' 使用多倍行距
                            targetRange.ParagraphFormat.LineSpacingRule = 5 ' wdLineSpaceMultiple
                            targetRange.ParagraphFormat.LineSpacing = 12 * lineSpacing ' 12磅 * 倍数
                        End If
                    Catch
                    End Try
                End If
            End If

            ' 字体颜色
            If formatting("color") IsNot Nothing Then
                Dim colorStr As String = formatting("color").ToString()
                If Not String.IsNullOrEmpty(colorStr) AndAlso colorStr <> "auto" Then
                    Try
                        Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                        targetRange.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                    Catch ex As Exception
                        Debug.WriteLine("设置字体颜色失败: " & ex.Message)
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine("ApplyFormattingToRange 出错: " & ex.Message)
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 处理AI语义标注结果 - 渲染引擎确定性应用格式
    ''' </summary>
    Private Async Sub ApplySemanticTaggingResult(taggingJson As String)
        Try
            If _reformatMapping Is Nothing Then
                GlobalStatusStrip.ShowWarning("排版映射上下文丢失")
                Return
            End If

            If _reformatParagraphs Is Nothing OrElse _reformatParagraphs.Count = 0 Then
                GlobalStatusStrip.ShowWarning("排版段落上下文丢失")
                Return
            End If

            ' 校验AI标注结果
            Dim validation = TaggingValidator.Validate(taggingJson, _reformatMapping, _reformatParagraphs.Count)

            If Not validation.IsValid Then
                ' 严重错误：尝试重试
                Debug.WriteLine("语义标注校验失败，错误数: " & validation.Errors.Count)

                ' 检查重试次数限制
                _reformatRetryCount += 1
                If _reformatRetryCount > MAX_REFORMAT_RETRIES Then
                    ' 超过重试限制，显示错误给用户
                    Debug.WriteLine($"重试次数超过限制({MAX_REFORMAT_RETRIES})，停止重试")
                    GlobalStatusStrip.ShowWarning($"AI标注解析失败，已重试{MAX_REFORMAT_RETRIES}次")
                    _reformatRetryCount = 0 ' 重置计数器

                    ' 显示错误详情
                    Dim errorMsg = String.Join(vbCrLf, validation.Errors.Take(5))
                    If validation.Errors.Count > 5 Then
                        errorMsg &= vbCrLf & $"...还有{validation.Errors.Count - 5}个错误"
                    End If
                    Await ShowReformatError($"AI标注解析失败:{vbCrLf}{errorMsg}")
                    Return
                End If

                ' 收集段落文本用于重试
                Dim paragraphTexts As New List(Of String)()
                For Each para In _reformatParagraphs
                    Try
                        paragraphTexts.Add(If(para.Range.Text?.ToString().TrimEnd(vbCr, vbLf), ""))
                    Catch
                        paragraphTexts.Add("")
                    End Try
                Next

                Debug.WriteLine($"第{_reformatRetryCount}次重试...")
                Dim retryPrompt = SemanticPromptBuilder.BuildRetryPrompt(_reformatMapping, paragraphTexts, validation.Errors)
                Await Send("标注结果存在错误，请修正。", retryPrompt, False, "semantic_reformat")
                Return
            End If

            ' 校验通过，重置重试计数器
            _reformatRetryCount = 0

            ' 使用渲染引擎应用格式
            Dim wordApp = Globals.ThisAddIn.Application

            ' 开始Undo快照（必须紧贴格式操作，跨async边界会失效）
            Try
                wordApp.UndoRecord.StartCustomRecord("AI排版")
            Catch ex As Exception
                Debug.WriteLine($"StartCustomRecord 失败: {ex.Message}")
            End Try

            Dim renderResult = SemanticRenderingEngine.ApplySemanticFormatting(
                validation.ValidatedTags, _reformatMapping, _reformatParagraphs, _reformatTypes, wordApp)

            ' 结束Undo快照
            Try
                wordApp.UndoRecord.EndCustomRecord()
            Catch ex As Exception
                Debug.WriteLine($"EndCustomRecord 失败: {ex.Message}")
            End Try

            ' 推送排版结果到前端
            Dim resultJson = renderResult.ToJson()
            Await ExecuteJavaScriptAsyncJS($"showReformatResult({resultJson.ToString(Formatting.None)});")

            ' 显示状态
            Dim resultMsg = $"排版完成，共处理 {renderResult.AppliedCount} 个段落"
            If renderResult.SkippedCount > 0 Then
                resultMsg &= $"，跳过 {renderResult.SkippedCount} 个特殊元素"
            End If
            If validation.AutoFixedCount > 0 Then
                resultMsg &= $"，自动修正 {validation.AutoFixedCount} 个标签"
            End If
            GlobalStatusStrip.ShowInfo(resultMsg)

            ' 清理段落引用（段落对象会失效），但保留映射上下文以支持后续排版
            _reformatParagraphs = Nothing
            _reformatStyles = Nothing
            _reformatTypes = Nothing
            ' _reformatMapping 保留，用户可继续使用同一映射排版其他内容

        Catch ex As Exception
            Debug.WriteLine($"ApplySemanticTaggingResult 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"语义排版应用失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 显示排版错误到前端
    ''' </summary>
    Private Async Function ShowReformatError(errorMsg As String) As Task
        Try
            Dim errorJson = New JObject From {
                {"success", False},
                {"error", errorMsg}
            }
            Await ExecuteJavaScriptAsyncJS($"showReformatResult({errorJson.ToString(Formatting.None)});")
        Catch ex As Exception
            Debug.WriteLine($"ShowReformatError 出错: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 处理规范转换结果 - AI将规范文本转为结构化映射后，继续发起语义标注
    ''' </summary>
    Private Async Sub HandleStyleGuideConversionResult(aiResponseJson As String)
        Try
            ' 解析AI返回的规范映射
            Dim guideName As String = "排版规范"
            Dim guideId As String = ""

            ' 尝试从响应映射获取规范信息
            Dim mapping = StyleGuideConverter.ParseAiResponse(aiResponseJson, guideName, guideId)

            If mapping Is Nothing OrElse mapping.SemanticTags.Count = 0 Then
                GlobalStatusStrip.ShowWarning("规范转换结果解析失败")
                Return
            End If

            ' 缓存映射
            SemanticMappingManager.Instance.AddMapping(mapping)

            ' 更新上下文中的mapping
            _reformatMapping = mapping

            ' 收集段落文本
            Dim paragraphTexts As New List(Of String)()
            If _reformatParagraphs IsNot Nothing Then
                For Each para In _reformatParagraphs
                    Try
                        paragraphTexts.Add(If(para.Range.Text?.ToString().TrimEnd(vbCr, vbLf), ""))
                    Catch
                        paragraphTexts.Add("")
                    End Try
                Next
            End If

            ' 构建语义标注提示词
            Dim systemPrompt = SemanticPromptBuilder.BuildSemanticTaggingPrompt(mapping, paragraphTexts)

            ' 开始Undo快照
            Dim wordApp = Globals.ThisAddIn.Application
            Try
                wordApp.UndoRecord.StartCustomRecord("AI排版")
            Catch ex As Exception
                Debug.WriteLine($"StartCustomRecord 失败: {ex.Message}")
            End Try

            ' 发送语义标注请求
            Await Send("规范已解析，现在进行语义标注。", systemPrompt, False, "semantic_reformat")

        Catch ex As Exception
            Debug.WriteLine($"HandleStyleGuideConversionResult 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"规范转换后标注失败: {ex.Message}")
        End Try
    End Sub

    ' 辅助：由纯文本生成最简单的 WordProcessingML OpenXML 片段（每个换行生成一个段落）
    Private Function BuildOpenXmlFromText(text As String) As String
        Try
            If String.IsNullOrEmpty(text) Then Return String.Empty
            Dim ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            Dim sb As New StringBuilder()
            sb.Append($"<w:document xmlns:w=""{ns}""><w:body>")
            Dim lines = text.Replace(vbCrLf, vbLf).Split(New Char() {vbLf})
            For Each line In lines
                Dim escaped = line
                escaped = escaped.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                ' 保留前后空格
                sb.Append($"<w:p><w:r><w:t xml:space=""preserve"">{escaped}</w:t></w:r></w:p>")
            Next
            sb.Append("</w:body></w:document>")
            Return sb.ToString()
        Catch ex As Exception
            Debug.WriteLine("BuildOpenXmlFromText 出错: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    Protected Overrides Function CaptureCurrentSelectionInfo(mode As String) As SelectionInfo
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            Dim txt As String = If(sel IsNot Nothing AndAlso sel.Range IsNot Nothing, sel.Range.Text, String.Empty)
            If String.IsNullOrEmpty(txt) Then
                Return Nothing
            End If

            Dim info As New SelectionInfo()
            info.SelectedText = txt
            info.DocumentPath = Globals.ThisAddIn.Application.ActiveDocument.FullName

            Try
                info.StartPos = sel.Range.Start
                info.EndPos = sel.Range.End
            Catch
                info.StartPos = 0
                info.EndPos = 0
            End Try

            Return info
        Catch
            Return Nothing
        End Try
    End Function

    ' ========== 续写功能 ==========

    Private _continuationService As WordContinuationService
    Private _cachedContinuationContext As ContinuationContext ' 缓存续写上下文，用于多轮续写

    ''' <summary>
    ''' 触发续写 - 获取光标上下文并发送AI请求
    ''' </summary>
    Protected Overrides Sub HandleTriggerContinuation(jsonDoc As JObject)
        Try
            ' 提取参数
            Dim style As String = ""
            Dim isContinuationMode As Boolean = False

            If jsonDoc IsNot Nothing Then
                If jsonDoc("style") IsNot Nothing Then
                    style = jsonDoc("style").ToString()
                End If
                If jsonDoc("isContinuationMode") IsNot Nothing Then
                    isContinuationMode = jsonDoc("isContinuationMode").ToObject(Of Boolean)()
                End If
            End If

            ' 初始化续写服务
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 检查是否可以续写
            If Not _continuationService.CanContinue() Then
                GlobalStatusStrip.ShowWarning("无法获取文档信息，请确保文档已打开")
                Return
            End If

            Dim context As ContinuationContext

            ' 如果是续写模式的后续请求，并且有缓存的上下文，则复用
            If isContinuationMode AndAlso _cachedContinuationContext IsNot Nothing Then
                ' 多轮续写：使用缓存的上下文，但style作为新的调整要求
                context = _cachedContinuationContext
                GlobalStatusStrip.ShowInfo("继续续写...")
            Else
                ' 首次续写或非续写模式：重新获取上下文
                context = _continuationService.GetCursorContext(3, 3)
                If context Is Nothing Then
                    GlobalStatusStrip.ShowWarning("无法获取文档上下文")
                    Return
                End If
                ' 缓存上下文
                _cachedContinuationContext = context
                GlobalStatusStrip.ShowInfo("正在分析上下文并生成续写内容...")
            End If

            ' 发送续写请求（带上风格参数）
            SendContinuationRequest(context, style)

        Catch ex As Exception
            Debug.WriteLine($"HandleTriggerContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"触发续写时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用续写结果到Word文档
    ''' </summary>
    Protected Overrides Sub HandleApplyContinuation(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("续写内容为空")
                Return
            End If

            ' 确保续写服务已初始化
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor
            End Select

            ' 插入续写内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("续写内容已插入文档")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeContinuationActions('{uuid}');")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"插入续写内容时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用模板渲染结果到Word文档
    ''' </summary>
    Protected Overrides Sub HandleApplyTemplateContent(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("模板内容为空")
                Return
            End If

            ' 确保续写服务已初始化（复用其插入逻辑）
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor
            End Select

            ' 插入模板内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("模板内容已插入文档")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeTemplateActions('{uuid}');")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyTemplateContent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"插入模板内容时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取当前Word上下文快照（用于自动补全）
    ''' </summary>
    Protected Overrides Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = "Word"

        Try
            Dim selection = Globals.ThisAddIn.Application.Selection
            If selection IsNot Nothing AndAlso selection.Start <> selection.End Then
                ' 有选中内容
                Dim selText = selection.Text
                If Not String.IsNullOrEmpty(selText) AndAlso selText.Length > 500 Then
                    selText = selText.Substring(0, 500) & "..."
                End If
                snapshot("selection") = If(selText, "")
            Else
                snapshot("selection") = ""
            End If

            ' 获取文档标题
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc IsNot Nothing Then
                snapshot("documentName") = If(doc.Name, "")
            End If

        Catch ex As Exception
            Debug.WriteLine($"GetContextSnapshot 出错: {ex.Message}")
        End Try

        Return snapshot
    End Function

    ''' <summary>
    ''' 重写保存设置方法，同步更新Word补全管理器状态
    ''' </summary>
    Protected Overrides Sub HandleSaveSettings(jsonDoc As JObject)
        MyBase.HandleSaveSettings(jsonDoc)

        ' 同步更新Word补全管理器的启用状态
        Try
            Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
            WordCompletionManager.Instance.Enabled = enableAutocomplete
            Debug.WriteLine($"[WordChatControl] 补全设置已同步: Enabled={enableAutocomplete}")
        Catch ex As Exception
            Debug.WriteLine($"[WordChatControl] 同步补全设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 执行JSON命令（重写基类方法）- 带严格验证
    ''' </summary>
    Protected Overrides Function ExecuteJsonCommand(jsonCode As String, preview As Boolean) As Boolean
        Try
            ' 预览模式下跳过自动执行（排版/校对模式的JSON用于预览，由用户手动点击应用）
            If IsInPreviewMode() Then
                Debug.WriteLine($"[WordChatControl] 预览模式({GetCurrentResponseMode()})下跳过JSON命令自动执行")
                Return True ' 返回True表示"成功处理"，避免显示错误
            End If

            ' 使用严格的结构验证
            Dim errorMessage As String = ""
            Dim normalizedJson As JToken = Nothing

            If Not WordJsonCommandSchema.ValidateJsonStructure(jsonCode, errorMessage, normalizedJson) Then
                ' 格式验证失败
                Debug.WriteLine($"Word JSON格式验证失败: {errorMessage}")
                Debug.WriteLine($"原始JSON: {jsonCode.Substring(0, Math.Min(200, jsonCode.Length))}...")

                ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式不符合规范: {errorMessage}")
                Return False
            End If

            ' 验证通过，根据类型执行
            If normalizedJson.Type = JTokenType.Object Then
                Dim jsonObj = CType(normalizedJson, JObject)

                ' 命令数组格式
                If jsonObj("commands") IsNot Nothing Then
                    Return ExecuteWordCommandsArray(jsonObj("commands"), jsonCode, preview)
                End If

                ' 单命令格式
                Return ExecuteWordSingleCommand(jsonObj, jsonCode, preview)
            End If

            ShareRibbon.GlobalStatusStrip.ShowWarning("无效的JSON格式")
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式无效: {ex.Message}")
            Return False
        Catch ex As Exception
            ShareRibbon.GlobalStatusStrip.ShowWarning($"执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行Word命令数组
    ''' </summary>
    Private Function ExecuteWordCommandsArray(commandsArray As JToken, originalJson As String, preview As Boolean) As Boolean
        Try
            Dim commands = CType(commandsArray, JArray)
            If commands.Count = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("命令数组为空")
                Return False
            End If

            ' 预览所有命令
            If preview Then
                ' 使用增强的预览表单
                If Not ShareRibbon.CommandPreviewForm.ShowPreview($"Word命令预览 - 共 {commands.Count} 个命令", commandsArray) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 执行所有命令
            Dim successCount = 0
            Dim failCount = 0

            For Each cmd In commands
                If cmd.Type = JTokenType.Object Then
                    Dim cmdObj = CType(cmd, JObject)
                    If ExecuteWordCommand(cmdObj) Then
                        successCount += 1
                    Else
                        failCount += 1
                    End If
                End If
            Next

            If failCount = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"所有 {successCount} 个命令执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"执行完成: {successCount} 成功, {failCount} 失败")
            End If

            Return failCount = 0

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommandsArray 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"批量执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行单个Word命令
    ''' </summary>
    Private Function ExecuteWordSingleCommand(commandJson As JObject, processedJson As String, preview As Boolean) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()

            ' 预览 - 使用增强的预览表单
            If preview Then
                If Not ShareRibbon.CommandPreviewForm.ShowPreview("Word命令预览", commandJson) Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 执行命令
            Dim success = ExecuteWordCommand(commandJson)

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"命令 '{command}' 执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"命令 '{command}' 执行失败")
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordSingleCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行具体的Word命令
    ''' </summary>
    Private Function ExecuteWordCommand(commandJson As JObject) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            Dim params = commandJson("params")

            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selection = Globals.ThisAddIn.Application.Selection

            Select Case command.ToLower()
                Case "inserttext"
                    Return ExecuteInsertText(params, selection)
                Case "formattext"
                    Return ExecuteFormatText(params, selection)
                Case "replacetext"
                    Return ExecuteReplaceText(params, doc)
                Case "inserttable"
                    Return ExecuteInsertTable(params, selection)
                Case "applystyle"
                    Return ExecuteApplyStyle(params, selection)
                Case "generatetoc"
                    Return ExecuteGenerateTOC(params, doc)
                Case "beautifydocument"
                    Return ExecuteBeautifyDocument(params, doc)
                Case Else
                    Debug.WriteLine($"不支持的Word命令: {command}")
                    Return False
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertText(params As JToken, selection As Object) As Boolean
        Try
            Dim content = params("content")?.ToString()
            Dim position = If(params("position")?.ToString(), "cursor")

            Select Case position.ToLower()
                Case "start"
                    Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0).InsertBefore(content)
                Case "end"
                    Globals.ThisAddIn.Application.ActiveDocument.Content.InsertAfter(content)
                Case Else ' cursor
                    selection.TypeText(content)
            End Select

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteFormatText(params As JToken, selection As Object) As Boolean
        Try
            Dim range = selection.Range

            If params("bold") IsNot Nothing Then
                range.Font.Bold = If(params("bold").Value(Of Boolean)(), -1, 0)
            End If

            If params("italic") IsNot Nothing Then
                range.Font.Italic = If(params("italic").Value(Of Boolean)(), -1, 0)
            End If

            If params("underline") IsNot Nothing Then
                range.Font.Underline = If(params("underline").Value(Of Boolean)(), 1, 0)
            End If

            If params("fontSize") IsNot Nothing Then
                range.Font.Size = params("fontSize").Value(Of Single)()
            End If

            If params("fontName") IsNot Nothing Then
                range.Font.Name = params("fontName").ToString()
            End If

            If params("color") IsNot Nothing Then
                Dim colorStr As String = params("color").ToString()
                If Not String.IsNullOrEmpty(colorStr) AndAlso colorStr <> "auto" Then
                    Try
                        Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                        range.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                    Catch ex As Exception
                        Debug.WriteLine("设置字体颜色失败: " & ex.Message)
                    End Try
                End If
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteReplaceText(params As JToken, doc As Object) As Boolean
        Try
            Dim find = params("find")?.ToString()
            Dim replace = If(params("replace")?.ToString(), "")
            Dim matchCase = If(params("matchCase")?.Value(Of Boolean)(), False)

            Dim findObj = doc.Content.Find
            findObj.ClearFormatting()
            findObj.Replacement.ClearFormatting()
            findObj.Text = find
            findObj.Replacement.Text = replace
            findObj.Forward = True
            findObj.Wrap = 1 ' wdFindContinue
            findObj.MatchCase = matchCase
            findObj.Execute(Replace:=2) ' wdReplaceAll

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteReplaceText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertTable(params As JToken, selection As Object) As Boolean
        Try
            Dim rows = params("rows")?.Value(Of Integer)()
            Dim cols = params("cols")?.Value(Of Integer)()

            If rows <= 0 OrElse cols <= 0 Then Return False

            Dim table = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
                selection.Range, rows, cols)

            ' 如果有data，填充表格
            Dim data = params("data")
            If data IsNot Nothing AndAlso data.Type = JTokenType.Array Then
                Dim dataArr = CType(data, JArray)
                Dim x As Integer = dataArr.Count - 1
                Dim x2 As Integer = rows - 1
                For rowIdx = 0 To Math.Min(x, x2)
                    Dim rowData = dataArr(rowIdx)
                    If rowData.Type = JTokenType.Array Then
                        Dim rowArr = CType(rowData, JArray)
                        Dim y As Integer = rowArr.Count - 1
                        Dim y1 As Integer = cols - 1
                        For colIdx = 0 To Math.Min(y, y1)
                            table.Cell(rowIdx + 1, colIdx + 1).Range.Text = rowArr(colIdx).ToString()
                        Next
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertTable 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteApplyStyle(params As JToken, selection As Object) As Boolean
        Try
            Dim styleName = params("styleName")?.ToString()
            If String.IsNullOrEmpty(styleName) Then Return False

            ' 检查样式是否存在
            Dim doc As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
            Dim styleExists As Boolean = False
            Try
                Dim testStyle = doc.Styles(styleName)
                styleExists = True
            Catch
                styleExists = False
            End Try

            If Not styleExists Then
                Debug.WriteLine($"ExecuteApplyStyle: 样式 '{styleName}' 不存在，跳过应用")
                ' 尝试使用内置样式名称映射
                Dim builtinStyleName = MapToBuiltinStyle(styleName)
                If Not String.IsNullOrEmpty(builtinStyleName) Then
                    Try
                        selection.Style = builtinStyleName
                        Return True
                    Catch
                        Debug.WriteLine($"ExecuteApplyStyle: 内置样式 '{builtinStyleName}' 也无法应用")
                    End Try
                End If
                Return True ' 返回True避免中断后续命令执行
            End If

            selection.Style = styleName
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyStyle 出错: {ex.Message}")
            Return True ' 返回True避免因样式问题中断整个流程
        End Try
    End Function

    ''' <summary>
    ''' 将常见样式名称映射到Word内置样式
    ''' </summary>
    Private Function MapToBuiltinStyle(styleName As String) As String
        Dim styleMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"标题", "标题 1"},
            {"Title", "Title"},
            {"标题1", "标题 1"},
            {"标题2", "标题 2"},
            {"标题3", "标题 3"},
            {"Heading1", "Heading 1"},
            {"Heading2", "Heading 2"},
            {"Heading3", "Heading 3"},
            {"正文", "正文"},
            {"Normal", "Normal"},
            {"副标题", "副标题"},
            {"Subtitle", "Subtitle"}
        }

        If styleMap.ContainsKey(styleName) Then
            Return styleMap(styleName)
        End If
        Return Nothing
    End Function

#Region "高级Word命令实现"

    ''' <summary>
    ''' 生成目录
    ''' </summary>
    Private Function ExecuteGenerateTOC(params As JToken, doc As Object) As Boolean
        Try
            Dim position = If(params("position")?.ToString(), "start")
            Dim levels = If(params("levels")?.Value(Of Integer)(), 3)
            Dim includePageNumbers = If(params("includePageNumbers")?.Value(Of Boolean)(), True)

            ' 确定插入位置
            Dim range As Object
            If position.ToLower() = "start" Then
                range = doc.Range(0, 0)
            Else
                range = Globals.ThisAddIn.Application.Selection.Range
            End If

            ' 删除已有目录
            For Each toc In doc.TablesOfContents
                toc.Delete()
            Next

            ' 插入新目录
            Dim newToc = doc.TablesOfContents.Add(
                Range:=range,
                UseHeadingStyles:=True,
                UpperHeadingLevel:=1,
                LowerHeadingLevel:=levels,
                IncludePageNumbers:=includePageNumbers
            )

            ' 更新目录
            newToc.Update()

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已生成{levels}级目录")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteGenerateTOC 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 文档美化
    ''' </summary>
    Private Function ExecuteBeautifyDocument(params As JToken, doc As Object) As Boolean
        Try
            Dim theme = params("theme")
            Dim margins = params("margins")

            ' 应用页边距
            If margins IsNot Nothing Then
                ApplyMargins(doc, margins)
            End If

            ' 应用主题样式
            If theme IsNot Nothing Then
                ApplyThemeStyles(doc, theme)
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo("文档美化完成")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteBeautifyDocument 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 应用页边距
    ''' </summary>
    Private Sub ApplyMargins(doc As Object, margins As JToken)
        Try
            Dim pageSetup = doc.PageSetup

            ' 单位转换: 厘米 -> 磅 (1cm = 28.35磅)
            Const cmToPoints As Single = 28.35F

            If margins("top") IsNot Nothing Then
                pageSetup.TopMargin = margins("top").Value(Of Single)() * cmToPoints
            End If
            If margins("bottom") IsNot Nothing Then
                pageSetup.BottomMargin = margins("bottom").Value(Of Single)() * cmToPoints
            End If
            If margins("left") IsNot Nothing Then
                pageSetup.LeftMargin = margins("left").Value(Of Single)() * cmToPoints
            End If
            If margins("right") IsNot Nothing Then
                pageSetup.RightMargin = margins("right").Value(Of Single)() * cmToPoints
            End If
        Catch ex As Exception
            Debug.WriteLine($"ApplyMargins 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用主题样式到文档
    ''' </summary>
    Private Sub ApplyThemeStyles(doc As Object, theme As JToken)
        Try
            ' 应用标题1样式
            Dim h1Theme = theme("h1")
            If h1Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 1", h1Theme)
            End If

            ' 应用标题2样式
            Dim h2Theme = theme("h2")
            If h2Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 2", h2Theme)
            End If

            ' 应用标题3样式
            Dim h3Theme = theme("h3")
            If h3Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 3", h3Theme)
            End If

            ' 应用正文样式
            Dim bodyTheme = theme("body")
            If bodyTheme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "正文", bodyTheme)

                ' 应用行间距到所有段落
                If bodyTheme("lineSpacing") IsNot Nothing Then
                    Dim lineSpacing = bodyTheme("lineSpacing").Value(Of Single)()
                    For Each para In doc.Paragraphs
                        Try
                            para.LineSpacingRule = 5 ' wdLineSpaceMultiple
                            para.LineSpacing = 12 * lineSpacing ' 12磅 * 倍数
                        Catch
                        End Try
                    Next
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyThemeStyles 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用样式设置
    ''' </summary>
    Private Sub ApplyStyleFromTheme(doc As Object, styleName As String, themeSettings As JToken)
        Try
            Dim style = doc.Styles(styleName)

            If themeSettings("font") IsNot Nothing Then
                style.Font.Name = themeSettings("font").ToString()
            End If
            If themeSettings("size") IsNot Nothing Then
                style.Font.Size = themeSettings("size").Value(Of Single)()
            End If
            If themeSettings("color") IsNot Nothing Then
                Dim colorStr = themeSettings("color").ToString()
                Dim color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                style.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
            End If
            If themeSettings("bold") IsNot Nothing Then
                style.Font.Bold = If(themeSettings("bold").Value(Of Boolean)(), -1, 0)
            End If
            If themeSettings("italic") IsNot Nothing Then
                style.Font.Italic = If(themeSettings("italic").Value(Of Boolean)(), -1, 0)
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyStyleFromTheme ({styleName}) 出错: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "语义排版 - .docx模板解析"

    ''' <summary>
    ''' 覆盖基类方法：从.docx文件解析SemanticStyleMapping并推送前端预览
    ''' </summary>
    Protected Overrides Sub HandleUploadDocxTemplateFromPath(filePath As String)
        Try
            Dim mapping = WordTemplateParser.ExtractFromDocx(filePath)
            If mapping Is Nothing OrElse mapping.SemanticTags.Count = 0 Then
                GlobalStatusStrip.ShowWarning("模板解析失败，未能提取到有效的样式信息")
                Return
            End If

            ' 将原始docx拷贝到数据目录，关联到映射
            Try
                Dim templateDir = IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    ConfigSettings.OfficeAiAppDataFolder, "docx_templates")
                If Not IO.Directory.Exists(templateDir) Then
                    IO.Directory.CreateDirectory(templateDir)
                End If
                Dim destPath = IO.Path.Combine(templateDir, mapping.Id & IO.Path.GetExtension(filePath))
                IO.File.Copy(filePath, destPath, True)
                mapping.SourceFilePath = destPath
            Catch ex As Exception
                Debug.WriteLine($"拷贝模板文件到数据目录失败: {ex.Message}")
                ' 非致命错误，继续保存映射
            End Try

            ' 保存映射
            SemanticMappingManager.Instance.AddMapping(mapping)

            ' 序列化为JSON并推送前端预览
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(mapping, Newtonsoft.Json.Formatting.None)
            ExecuteJavaScriptAsyncJS($"showMappingPreview({json});")

            GlobalStatusStrip.ShowInfo($"模板「{mapping.Name}」解析完成，提取到 {mapping.SemanticTags.Count} 个语义标签")
        Catch ex As Exception
            Debug.WriteLine($"HandleUploadDocxTemplateFromPath 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"模板解析失败: {ex.Message}")
        End Try
    End Sub

#End Region

End Class
