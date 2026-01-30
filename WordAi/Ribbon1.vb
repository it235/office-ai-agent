' WordAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Xml
Imports AngleSharp
Imports Microsoft.Office.Tools.Ribbon
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon  ' 添加此引用

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Protected Overrides Async Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Async Sub WebResearchButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowDataCaptureTaskPane()
    End Sub
    Protected Overrides Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs)
        'Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' Word 特定的数据分析逻辑
        MessageBox.Show("Word数据分析功能正在开发中...")
    End Sub

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Word", OfficeApplicationType.Word)
    End Function

    Protected Overrides Sub DeepseekButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowDeepseekTaskPane()
    End Sub

    Protected Overrides Sub DoubaoButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowDoubaoTaskPane()
    End Sub
    Protected Overrides Sub BatchDataGenButton_Click(sender As Object, e As RibbonControlEventArgs)
    End Sub

    Protected Overrides Sub MCPButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' 创建并显示MCP配置表单
        Dim mcpConfigForm As New MCPConfigForm()
        If mcpConfigForm.ShowDialog() = DialogResult.OK Then
            ' 在需要时可以集成到ChatControl调用MCP服务
        End If
    End Sub

    ' Proofread 按钮 — 校对功能（简化版：仅校对选中内容，使用段落索引定位）
    Protected Overrides Async Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            Dim selText As String = String.Empty

            Try
                If wordApp IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                    selText = If(wordApp.Selection.Range IsNot Nothing, wordApp.Selection.Range.Text, String.Empty)
                End If
            Catch ex As Exception
                selText = String.Empty
            End Try

            ' 必须先选中内容
            If String.IsNullOrWhiteSpace(selText) Then
                MessageBox.Show("请先选中需要校对的文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selRange = wordApp.Selection.Range

            ' 按段落分割选中内容，构建带索引的文本
            Dim paragraphs As New List(Of String)()
            Dim paraIndex As Integer = 0
            Dim sb As New StringBuilder()
            sb.AppendLine("以下是需要校对的内容（按段落编号）：")

            For Each p In selRange.Paragraphs
                Dim paraText As String = If(p.Range.Text IsNot Nothing, p.Range.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)
                If Not String.IsNullOrWhiteSpace(paraText) Then
                    paragraphs.Add(paraText)
                    sb.AppendLine($"[段落{paraIndex}] {paraText}")
                    paraIndex += 1
                End If
            Next

            If paragraphs.Count = 0 Then
                MessageBox.Show("选中的内容没有有效段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 打开侧栏
            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(250)

            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                MessageBox.Show("无法获取聊天控件实例，请确认 Chat 面板已打开。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 显示校对模式吸顶提示
            Await chatCtrl.ExecuteJavaScriptAsyncJS("showProofreadModeIndicator();")

            ' 构建前端提示
            buildHtmlHint(chatCtrl, "正在向模型发起校对请求，请耐心等待")

            ' 构建简化的提示词
            Dim systemPrompt As New StringBuilder()
            systemPrompt.AppendLine("你是Word内容校对助手。请检查以下内容中的错字、错标点或不当换行，并给出修正建议。")
            systemPrompt.AppendLine("必须且仅返回一个严格的JSON数组，格式如下：")
            systemPrompt.AppendLine("[{")
            systemPrompt.AppendLine("  ""paraIndex"": 0,")
            systemPrompt.AppendLine("  ""original"": ""原文片段"",")
            systemPrompt.AppendLine("  ""corrected"": ""修正后的文字"",")
            systemPrompt.AppendLine("  ""reason"": ""简短说明修正原因""")
            systemPrompt.AppendLine("}]")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("注意：")
            systemPrompt.AppendLine("- paraIndex是段落编号，从0开始")
            systemPrompt.AppendLine("- 如果没有需要修正的内容，返回空数组[]")
            systemPrompt.AppendLine("- 不要输出任何非JSON内容")

            Await chatCtrl.Send(sb.ToString(), systemPrompt.ToString(), False, "proofread")
        Catch ex As Exception
            MessageBox.Show("执行校对过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub buildHtmlHint(chatCtrl As ChatControl, displayContent As String)

        Try
            Dim responseUuid As String = Guid.NewGuid().ToString()
            Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName
            Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCreate)
            Dim js = $"appendRenderer('{responseUuid}','{displayContent}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine("ExecuteJavaScriptAsyncJS 调用失败: " & ex.Message)
        End Try
    End Sub

    ' 排版功能（简化版：按段落获取WordOpenXML并替换）
    Protected Overrides Async Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
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
                MessageBox.Show("请先选中需要排版的文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selRange = wordApp.Selection.Range

            ' 按段落分割，构建简化的blocks数组
            Dim blocks As New Newtonsoft.Json.Linq.JArray()
            Dim paraIndex As Integer = 0

            For Each p In selRange.Paragraphs
                Dim r = p.Range
                Dim paraText As String = If(r.Text IsNot Nothing, r.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)

                If Not String.IsNullOrWhiteSpace(paraText) Then
                    Dim paraObj As New Newtonsoft.Json.Linq.JObject()
                    paraObj("paraIndex") = paraIndex
                    paraObj("text") = paraText
                    ' 不再传递 WordOpenXML，只传递文本

                    blocks.Add(paraObj)
                    paraIndex += 1
                End If
            Next

            If blocks.Count = 0 Then
                MessageBox.Show("选中的内容没有有效段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(250)

            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                MessageBox.Show("无法获取聊天控件实例，请确认 Chat 面板已打开。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 显示排版模式吸顶提示
            Await chatCtrl.ExecuteJavaScriptAsyncJS("showReformatModeIndicator();")

            ' 构建前端提示
            buildHtmlHint(chatCtrl, "正在向模型发起排版请求，请耐心等待")

            ' 构建简化的系统提示 - 返回格式化属性而非XML
            Dim systemPrompt As New System.Text.StringBuilder()
            systemPrompt.AppendLine("你是Word排版助手。我提供文档段落，请帮我优化排版。")
            systemPrompt.AppendLine("排版规则：")
            systemPrompt.AppendLine("1. 中文字体使用宋体，英文使用Times New Roman")
            systemPrompt.AppendLine("2. 正文字号12pt（小四），标题根据级别设置（如16pt/14pt）")
            systemPrompt.AppendLine("3. 段落首行缩进2字符")
            systemPrompt.AppendLine("4. 行距1.5倍")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("必须且仅返回一个严格的JSON数组，格式如下：")
            systemPrompt.AppendLine("[{")
            systemPrompt.AppendLine("  ""paraIndex"": 0,")
            systemPrompt.AppendLine("  ""formatting"": {")
            systemPrompt.AppendLine("    ""fontNameCN"": ""宋体"",")
            systemPrompt.AppendLine("    ""fontNameEN"": ""Times New Roman"",")
            systemPrompt.AppendLine("    ""fontSize"": 12,")
            systemPrompt.AppendLine("    ""bold"": false,")
            systemPrompt.AppendLine("    ""alignment"": ""left"",")
            systemPrompt.AppendLine("    ""firstLineIndent"": 2,")
            systemPrompt.AppendLine("    ""lineSpacing"": 1.5")
            systemPrompt.AppendLine("  },")
            systemPrompt.AppendLine("  ""previewText"": ""格式化后的纯文本预览"",")
            systemPrompt.AppendLine("  ""changes"": ""简述做了哪些格式修改""")
            systemPrompt.AppendLine("}]")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("formatting字段说明：")
            systemPrompt.AppendLine("- fontNameCN: 中文字体名称")
            systemPrompt.AppendLine("- fontNameEN: 英文字体名称")
            systemPrompt.AppendLine("- fontSize: 字号(pt)")
            systemPrompt.AppendLine("- bold: 是否加粗(true/false)")
            systemPrompt.AppendLine("- alignment: 对齐方式(left/center/right/justify)")
            systemPrompt.AppendLine("- firstLineIndent: 首行缩进字符数(0表示无缩进)")
            systemPrompt.AppendLine("- lineSpacing: 行距倍数(1/1.5/2等)")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("注意：")
            systemPrompt.AppendLine("- paraIndex是段落编号，与输入对应")
            systemPrompt.AppendLine("- 如果段落无需修改，可以不包含该段落")
            systemPrompt.AppendLine("- 不要输出任何非JSON内容")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("以下是需要排版的段落：")
            systemPrompt.AppendLine(blocks.ToString(Newtonsoft.Json.Formatting.Indented))

            Await chatCtrl.Send("请基于提供的段落进行排版优化。", systemPrompt.ToString(), False, "reformat")

        Catch ex As Exception
            MessageBox.Show("执行排版过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 一键翻译功能
    Protected Overrides Async Sub TranslateButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim wordApp = Globals.ThisAddIn.Application

            ' 检查是否有选中内容
            Dim hasSelection As Boolean = False
            Try
                If wordApp?.Selection?.Range IsNot Nothing Then
                    Dim selText = wordApp.Selection.Range.Text
                    hasSelection = Not String.IsNullOrWhiteSpace(selText)
                End If
            Catch
                hasSelection = False
            End Try

            ' 显示翻译操作对话框
            Dim actionForm As New ShareRibbon.TranslateActionForm(hasSelection, "Word")
            If actionForm.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            ' 创建翻译服务
            Dim translateService As New WordDocumentTranslateService(wordApp)

            ' 更新设置
            Dim settings = ShareRibbon.TranslateSettings.Load()
            settings.SourceLanguage = actionForm.SourceLanguage
            settings.TargetLanguage = actionForm.TargetLanguage
            settings.CurrentDomain = actionForm.SelectedDomain
            settings.OutputMode = actionForm.OutputMode
            settings.Save()

            ' 显示进度
            ShareRibbon.GlobalStatusStripAll.ShowWarning("正在准备翻译...")

            ' 绑定进度事件
            AddHandler translateService.ProgressChanged, Sub(s, args)
                                                             ShareRibbon.GlobalStatusStripAll.ShowWarning(args.Message)
                                                         End Sub

            ' 执行翻译
            Dim results As List(Of ShareRibbon.TranslateParagraphResult)
            If actionForm.TranslateAll Then
                results = Await translateService.TranslateAllAsync()
            Else
                results = Await translateService.TranslateSelectionAsync()
            End If

            ' 应用翻译结果
            If actionForm.OutputMode = ShareRibbon.TranslateOutputMode.SidePanel Then
                ' 在侧栏显示
                Globals.ThisAddIn.ShowChatTaskPane()
                Await Task.Delay(250)

                Dim chatCtrl = Globals.ThisAddIn.chatControl
                If chatCtrl IsNot Nothing Then
                    Dim displayText = translateService.FormatResultsForDisplay(results, True)
                    Dim responseUuid As String = Guid.NewGuid().ToString()
                    Dim aiName As String = "AI翻译助手"
                    Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
                    Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCreate)

                    ' 转义特殊字符
                    Dim escapedText = displayText.Replace("\", "\\").Replace("'", "\'").Replace(vbCr, "\n").Replace(vbLf, "")
                    Dim js = $"appendRenderer('{responseUuid}','{escapedText}');"
                    Await chatCtrl.ExecuteJavaScriptAsyncJS(js)
                End If
            Else
                ' 应用到文档
                If actionForm.TranslateAll Then
                    translateService.ApplyTranslation(results, actionForm.OutputMode)
                Else
                    translateService.ApplyTranslationToSelection(results, actionForm.OutputMode)
                End If
            End If

            ShareRibbon.GlobalStatusStripAll.ShowWarning($"翻译完成，共处理 {results.Count} 个段落")

        Catch ex As Exception
            MessageBox.Show("翻译过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' AI续写功能
    Protected Overrides Sub ContinuationButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ' 确保侧栏已打开
            Globals.ThisAddIn.ShowChatTaskPane()

            ' 获取ChatControl并触发续写（自动模式，显示对话框）
            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl IsNot Nothing Then
                ' 稍等一下让WebView2加载完成，然后显示续写按钮并触发续写对话框
                Task.Run(Async Function()
                             Await Task.Delay(300)
                             ' 先显示续写按钮
                             Await chatCtrl.ExecuteJavaScriptAsyncJS("setContinuationButtonVisible(true);")
                             ' 再触发续写对话框
                             Await chatCtrl.ExecuteJavaScriptAsyncJS("triggerContinuation(true);")
                         End Function)
            Else
                MessageBox.Show("请先打开AI助手面板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("触发AI续写时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 接受补全功能
    Protected Overrides Sub AcceptCompletionButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim completionManager = WordCompletionManager.Instance
            If completionManager IsNot Nothing AndAlso completionManager.HasGhostText Then
                completionManager.AcceptCurrentCompletion()
            Else
                ' 没有可接受的补全时，显示提示
                MessageBox.Show("当前没有可接受的补全建议。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("接受补全时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 模板排版功能 - Word实现（使用JSON格式完整提取模板结构）
    Protected Overrides Sub TemplateFormatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ' 1. 打开文件对话框选择模板文件
            Using openDialog As New OpenFileDialog()
                openDialog.Title = "选择Word模板文件"
                openDialog.Filter = "Word文档|*.docx;*.doc|所有文件|*.*"
                openDialog.FilterIndex = 1

                If openDialog.ShowDialog() <> DialogResult.OK Then Return

                Dim templatePath = openDialog.FileName
                Dim templateName = System.IO.Path.GetFileName(templatePath)

                ' 2. 读取模板文件内容 - 使用JSON格式完整提取
                Dim wordApp = Globals.ThisAddIn.Application
                Dim templateJson As JObject = Nothing

                ' 打开模板文档（只读）
                Dim templateDoc As Microsoft.Office.Interop.Word.Document = Nothing
                Try
                    templateDoc = wordApp.Documents.Open(templatePath, ReadOnly:=True, Visible:=False)

                    ' 构建JSON结构
                    templateJson = ExtractTemplateStructure(templateDoc, templateName)
                Finally
                    If templateDoc IsNot Nothing Then
                        templateDoc.Close(SaveChanges:=False)
                    End If
                End Try

                If templateJson Is Nothing Then
                    MessageBox.Show("无法解析模板文件内容。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                ' 3. 打开Chat面板并进入模板渲染模式
                Globals.ThisAddIn.ShowChatTaskPane()
                Dim chatCtrl = Globals.ThisAddIn.chatControl
                If chatCtrl IsNot Nothing Then
                    ' 将JSON转为字符串传递给JS
                    Dim templateContent = templateJson.ToString(Newtonsoft.Json.Formatting.Indented)

                    ' 调用JS进入模板渲染模式
                    Task.Run(Async Function()
                                 Await Task.Delay(500) ' 等待WebView加载
                                 Dim jsCall = $"enterTemplateMode(`{EscapeForJs(templateContent)}`, `{EscapeForJs(templateName)}`);"
                                 Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCall)
                             End Function)

                    MessageBox.Show("已进入模板渲染模式！" & vbCrLf & vbCrLf &
                                    "模板结构已解析完成（包含段落、样式、字体、图片等信息）。" & vbCrLf &
                                    "现在您可以在Chat中输入内容需求，AI将按照模板格式生成内容。" & vbCrLf &
                                    "生成完成后可选择插入位置将内容插入到文档中。",
                                    "模板模式已激活", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("加载模板时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 提取Word文档的完整结构为JSON格式
    ''' </summary>
    Private Function ExtractTemplateStructure(doc As Microsoft.Office.Interop.Word.Document, templateName As String) As JObject
        Dim result As New JObject()
        result("templateName") = templateName
        result("totalParagraphs") = doc.Paragraphs.Count

        ' 元素数组：包含段落、图片、表格等
        Dim elements As New JArray()
        Dim elementIndex As Integer = 0

        ' 收集样式信息
        Dim stylesDict As New Dictionary(Of String, JObject)()

        ' 遍历段落（最多200段）
        For i = 1 To Math.Min(doc.Paragraphs.Count, 200)
            Dim para = doc.Paragraphs(i)
            Dim r = para.Range
            Dim text As String = If(r.Text IsNot Nothing, r.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)

            ' 获取段落样式
            Dim style = TryCast(para.Style, Microsoft.Office.Interop.Word.Style)
            Dim styleName As String = If(style?.NameLocal, "Normal")

            ' 收集样式详情
            If style IsNot Nothing AndAlso Not stylesDict.ContainsKey(styleName) Then
                Dim styleObj As New JObject()
                styleObj("fontName") = If(style.Font.Name, "")
                styleObj("fontSize") = If(style.Font.Size > 0, CDec(style.Font.Size), 12)
                styleObj("bold") = (style.Font.Bold = -1)
                styleObj("italic") = (style.Font.Italic = -1)
                stylesDict(styleName) = styleObj
            End If

            ' 创建段落元素
            Dim paraObj As New JObject()
            paraObj("type") = "paragraph"
            paraObj("index") = elementIndex
            paraObj("text") = text
            paraObj("styleName") = styleName

            ' 提取段落的详细格式信息
            Dim formatting As New JObject()
            Try
                ' 字体信息
                formatting("fontName") = If(r.Font.Name, "")
                formatting("fontSize") = If(r.Font.Size > 0, CDec(r.Font.Size), 12)
                formatting("bold") = (r.Font.Bold = -1)
                formatting("italic") = (r.Font.Italic = -1)
                formatting("underline") = (r.Font.Underline <> Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone)
                formatting("color") = If(r.Font.Color <> Microsoft.Office.Interop.Word.WdColor.wdColorAutomatic,
                                        ColorToHex(CInt(r.Font.Color)), "auto")

                ' 段落格式
                formatting("alignment") = GetAlignmentString(para.Alignment)
                formatting("firstLineIndent") = Math.Round(CDec(para.FirstLineIndent) / 28.35, 2) ' 转换为字符
                formatting("leftIndent") = Math.Round(CDec(para.LeftIndent) / 28.35, 2)
                formatting("lineSpacing") = GetLineSpacingValue(para)
                formatting("spaceBefore") = Math.Round(CDec(para.SpaceBefore), 1)
                formatting("spaceAfter") = Math.Round(CDec(para.SpaceAfter), 1)
            Catch ex As Exception
                Debug.WriteLine($"提取段落 {i} 格式时出错: {ex.Message}")
            End Try

            paraObj("formatting") = formatting

            ' 检查是否包含图片
            If r.InlineShapes.Count > 0 Then
                paraObj("hasImages") = True
                paraObj("imageCount") = r.InlineShapes.Count
            End If

            ' 检查是否包含公式
            If r.OMaths.Count > 0 Then
                paraObj("hasFormulas") = True
                paraObj("formulaCount") = r.OMaths.Count
            End If

            elements.Add(paraObj)
            elementIndex += 1
        Next

        ' 检查文档中的表格
        If doc.Tables.Count > 0 Then
            For t = 1 To Math.Min(doc.Tables.Count, 20)
                Dim table = doc.Tables(t)
                Dim tableObj As New JObject()
                tableObj("type") = "table"
                tableObj("index") = elementIndex
                tableObj("rows") = table.Rows.Count
                tableObj("columns") = table.Columns.Count

                ' 提取表格首行内容作为表头示例
                Dim headerCells As New JArray()
                Try
                    For c = 1 To table.Columns.Count
                        Dim cellText = table.Cell(1, c).Range.Text
                        cellText = cellText.TrimEnd(vbCr, vbLf, ChrW(7))
                        headerCells.Add(cellText)
                    Next
                    tableObj("headerCells") = headerCells
                Catch
                    ' 忽略合并单元格等情况
                End Try

                elements.Add(tableObj)
                elementIndex += 1
            Next
        End If

        ' 检查文档中的图片（非内嵌）
        If doc.Shapes.Count > 0 Then
            For s = 1 To Math.Min(doc.Shapes.Count, 20)
                Dim shape = doc.Shapes(s)
                If shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture OrElse
                   shape.Type = Microsoft.Office.Core.MsoShapeType.msoLinkedPicture Then
                    Dim imgObj As New JObject()
                    imgObj("type") = "image"
                    imgObj("index") = elementIndex
                    imgObj("width") = Math.Round(CDec(shape.Width), 1)
                    imgObj("height") = Math.Round(CDec(shape.Height), 1)
                    imgObj("description") = "浮动图片"
                    elements.Add(imgObj)
                    elementIndex += 1
                End If
            Next
        End If

        result("elements") = elements

        ' 添加样式集合
        Dim stylesObj As New JObject()
        For Each kvp In stylesDict
            stylesObj(kvp.Key) = kvp.Value
        Next
        result("styles") = stylesObj

        Return result
    End Function

    ''' <summary>
    ''' 将对齐方式转换为字符串
    ''' </summary>
    Private Function GetAlignmentString(alignment As Microsoft.Office.Interop.Word.WdParagraphAlignment) As String
        Select Case alignment
            Case Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                Return "left"
            Case Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                Return "center"
            Case Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                Return "right"
            Case Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
                Return "justify"
            Case Else
                Return "left"
        End Select
    End Function

    ''' <summary>
    ''' 获取行距值（返回倍数）
    ''' </summary>
    Private Function GetLineSpacingValue(para As Microsoft.Office.Interop.Word.Paragraph) As Decimal
        Try
            Select Case para.LineSpacingRule
                Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle
                    Return 1.0D
                Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpace1pt5
                    Return 1.5D
                Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceDouble
                    Return 2.0D
                Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceMultiple
                    Return Math.Round(CDec(para.LineSpacing) / 12, 2)
                Case Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly,
                     Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceAtLeast
                    Return Math.Round(CDec(para.LineSpacing) / 12, 2)
                Case Else
                    Return 1.0D
            End Select
        Catch
            Return 1.0D
        End Try
    End Function

    ''' <summary>
    ''' 将Word颜色值转换为十六进制字符串
    ''' </summary>
    Private Function ColorToHex(colorValue As Integer) As String
        Try
            Dim r = colorValue And &HFF
            Dim g = (colorValue >> 8) And &HFF
            Dim b = (colorValue >> 16) And &HFF
            Return $"#{r:X2}{g:X2}{b:X2}"
        Catch
            Return "auto"
        End Try
    End Function

    Private Function EscapeForJs(text As String) As String
        Return text.Replace("\", "\\").Replace("`", "\`").Replace("$", "\$").Replace(vbCr, "").Replace(vbLf, "\n")
    End Function

End Class
