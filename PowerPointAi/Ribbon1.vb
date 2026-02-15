' PowerPointAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms
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
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub

    Protected Overrides Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs)
        'Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' Word 特定的数据分析逻辑
        MessageBox.Show("Word数据分析功能正在开发中...")
    End Sub

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("PowerPoint", OfficeApplicationType.PowerPoint)
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

    Protected Overrides Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show("PowerPoint校对功能正在开发中...", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ''' <summary>
    ''' 在聊天面板显示提示信息
    ''' </summary>
    Private Async Sub buildHtmlHint(chatCtrl As ChatControl, displayContent As String)
        Try
            Dim responseUuid As String = Guid.NewGuid().ToString()
            Dim aiName As String = ShareRibbon.ConfigSettings.platform & " " & ShareRibbon.ConfigSettings.ModelName
            Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCreate)
            Dim js = $"appendRenderer('{responseUuid}','{displayContent}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine("ExecuteJavaScriptAsyncJS 调用失败: " & ex.Message)
        End Try
    End Sub

    ' 排版功能 - 进入模板选择模式
    Protected Overrides Async Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ' 打开Chat面板并进入模板选择模式（不再预先检查选中内容，改为选择模板后再检查）
            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(250)

            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                MessageBox.Show("无法获取聊天控件实例，请确认 Chat 面板已打开。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 进入模板选择模式
            chatCtrl.EnterReformatTemplateMode()

        Catch ex As Exception
            MessageBox.Show("进入排版模式出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 判断形状类型（标题/副标题/正文）
    ''' </summary>
    Private Function GetShapeType(shp As Microsoft.Office.Interop.PowerPoint.Shape) As String
        Try
            If shp.PlaceholderFormat IsNot Nothing Then
                Select Case shp.PlaceholderFormat.Type
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                         Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle
                        Return "标题"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle
                        Return "副标题"
                    Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody
                        Return "正文"
                End Select
            End If
        Catch
        End Try
        Return "文本框"
    End Function

    ' 一键翻译功能 - PowerPoint实现
    Protected Overrides Async Sub TranslateButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim pptApp = Globals.ThisAddIn.Application

            ' 检查是否有选中内容
            Dim hasSelection As Boolean = False
            Try
                Dim sel = pptApp.ActiveWindow.Selection
                hasSelection = (sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText OrElse
                               sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes OrElse
                               sel.Type = Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides)
            Catch
                hasSelection = False
            End Try

            ' 显示翻译操作对话框
            Dim actionForm As New ShareRibbon.TranslateActionForm(hasSelection, "PowerPoint")
            If actionForm.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            ' 创建翻译服务
            Dim translateService As New PowerPointDocumentTranslateService(pptApp)

            ' 更新设置
            Dim settings = ShareRibbon.TranslateSettings.Load()
            settings.SourceLanguage = actionForm.SourceLanguage
            settings.TargetLanguage = actionForm.TargetLanguage
            settings.CurrentDomain = actionForm.SelectedDomain
            settings.OutputMode = actionForm.OutputMode
            settings.Save()

            ' 显示进度
            ShareRibbon.GlobalStatusStripAll.ShowWarning("正在准备翻译... " & translateService.GetStatistics())

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
                ' 应用到演示文稿
                If actionForm.TranslateAll Then
                    translateService.ApplyTranslation(results, actionForm.OutputMode)
                Else
                    translateService.ApplyTranslationToSelection(results, actionForm.OutputMode)
                End If
            End If

            ShareRibbon.GlobalStatusStripAll.ShowWarning($"翻译完成，共处理 {results.Count} 个文本块")

        Catch ex As Exception
            MessageBox.Show("翻译过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' AI续写功能 - PowerPoint实现
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

    ' 接受补全功能 - PowerPoint实现
    Protected Sub AcceptCompletionButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim completionManager = PowerPointCompletionManager.Instance
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

    ' 模板排版功能 - PowerPoint实现（使用JSON格式完整提取模板结构）
    Protected Overrides Sub TemplateFormatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ' 1. 打开文件对话框选择模板文件
            Using openDialog As New OpenFileDialog()
                openDialog.Title = "选择PowerPoint模板文件"
                openDialog.Filter = "PowerPoint文件|*.pptx;*.ppt|所有文件|*.*"
                openDialog.FilterIndex = 1

                If openDialog.ShowDialog() <> DialogResult.OK Then Return

                Dim templatePath = openDialog.FileName
                Dim templateName = System.IO.Path.GetFileName(templatePath)

                ' 2. 读取模板文件内容 - 使用JSON格式完整提取
                Dim pptApp = Globals.ThisAddIn.Application
                Dim templateJson As JObject = Nothing

                ' 打开模板演示文稿（只读）
                Dim templatePres As Microsoft.Office.Interop.PowerPoint.Presentation = Nothing
                Try
                    templatePres = pptApp.Presentations.Open(templatePath, ReadOnly:=Microsoft.Office.Core.MsoTriState.msoTrue, WithWindow:=Microsoft.Office.Core.MsoTriState.msoFalse)

                    ' 构建JSON结构
                    templateJson = ExtractPresentationStructure(templatePres, templateName)
                Finally
                    If templatePres IsNot Nothing Then
                        templatePres.Close()
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
                    Dim templateContent = templateJson.ToString(Formatting.Indented)

                    ' 调用JS进入模板渲染模式
                    Task.Run(Async Function()
                                 Await Task.Delay(500)
                                 Dim jsCall = $"enterTemplateMode(`{EscapeForJs(templateContent)}`, `{EscapeForJs(templateName)}`);"
                                 Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCall)
                             End Function)

                    MessageBox.Show("已进入模板渲染模式！" & vbCrLf & vbCrLf &
                                    "模板结构已解析完成（包含幻灯片、文本、样式、图片等信息）。" & vbCrLf &
                                    "现在您可以在Chat中输入内容需求，AI将按照模板格式生成内容。" & vbCrLf &
                                    "生成完成后可选择插入位置将内容插入到演示文稿中。",
                                    "模板模式已激活", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("加载模板时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 提取PPT演示文稿的完整结构为JSON格式
    ''' </summary>
    Private Function ExtractPresentationStructure(pres As Microsoft.Office.Interop.PowerPoint.Presentation, templateName As String) As JObject
        Dim result As New JObject()
        result("templateName") = templateName
        result("totalSlides") = pres.Slides.Count
        result("slideWidth") = pres.PageSetup.SlideWidth
        result("slideHeight") = pres.PageSetup.SlideHeight

        ' 幻灯片数组
        Dim slides As New JArray()

        ' 遍历幻灯片（最多30张）
        For i = 1 To Math.Min(pres.Slides.Count, 30)
            Dim slide = pres.Slides(i)
            Dim slideObj As New JObject()
            slideObj("slideIndex") = i
            slideObj("slideLayout") = GetLayoutName(slide.Layout)

            ' 元素数组：包含文本框、图片、表格等
            Dim elements As New JArray()
            Dim elementIndex As Integer = 0

            ' 遍历幻灯片中的形状
            For Each shape As Microsoft.Office.Interop.PowerPoint.Shape In slide.Shapes
                Dim elemObj As New JObject()
                elemObj("index") = elementIndex

                ' 判断形状类型
                If shape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' 文本框/占位符
                    Dim text = shape.TextFrame.TextRange.Text.Trim()
                    elemObj("type") = "textbox"
                    elemObj("text") = text
                    elemObj("placeholderType") = GetPlaceholderTypeName(shape)

                    ' 提取文本格式
                    Dim formatting As New JObject()
                    Try
                        Dim textRange = shape.TextFrame.TextRange
                        formatting("fontName") = If(textRange.Font.Name, "")
                        formatting("fontSize") = If(textRange.Font.Size > 0, CDec(textRange.Font.Size), 18)
                        formatting("bold") = (textRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue)
                        formatting("italic") = (textRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue)
                        formatting("underline") = (textRange.Font.Underline = Microsoft.Office.Core.MsoTriState.msoTrue)

                        ' 颜色
                        Try
                            Dim rgb = textRange.Font.Color.RGB
                            formatting("color") = $"#{rgb And &HFF:X2}{(rgb >> 8) And &HFF:X2}{(rgb >> 16) And &HFF:X2}"
                        Catch
                            formatting("color") = "auto"
                        End Try

                        ' 对齐方式
                        formatting("alignment") = GetPPTAlignmentString(textRange.ParagraphFormat.Alignment)
                    Catch ex As Exception
                        Debug.WriteLine($"提取PPT文本格式时出错: {ex.Message}")
                    End Try
                    elemObj("formatting") = formatting

                    ' 位置和大小
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                ElseIf shape.Type = Microsoft.Office.Core.MsoShapeType.msoPicture OrElse
                       shape.Type = Microsoft.Office.Core.MsoShapeType.msoLinkedPicture Then
                    ' 图片
                    elemObj("type") = "image"
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                ElseIf shape.HasTable = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' 表格
                    elemObj("type") = "table"
                    elemObj("rows") = shape.Table.Rows.Count
                    elemObj("columns") = shape.Table.Columns.Count
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                    ' 提取表格首行作为示例
                    Dim headerCells As New JArray()
                    Try
                        For c = 1 To shape.Table.Columns.Count
                            Dim cellText = shape.Table.Cell(1, c).Shape.TextFrame.TextRange.Text.Trim()
                            headerCells.Add(cellText)
                        Next
                        elemObj("headerCells") = headerCells
                    Catch
                        ' 忽略合并单元格等情况
                    End Try

                ElseIf shape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
                    ' 图表
                    elemObj("type") = "chart"
                    elemObj("chartType") = shape.Chart.ChartType.ToString()
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)

                Else
                    ' 其他形状
                    elemObj("type") = "shape"
                    elemObj("shapeType") = shape.Type.ToString()
                    elemObj("left") = Math.Round(CDec(shape.Left), 1)
                    elemObj("top") = Math.Round(CDec(shape.Top), 1)
                    elemObj("width") = Math.Round(CDec(shape.Width), 1)
                    elemObj("height") = Math.Round(CDec(shape.Height), 1)
                End If

                elements.Add(elemObj)
                elementIndex += 1
            Next

            slideObj("elements") = elements
            slides.Add(slideObj)
        Next

        result("slides") = slides
        Return result
    End Function

    Private Function GetLayoutName(layout As Microsoft.Office.Interop.PowerPoint.PpSlideLayout) As String
        Select Case layout
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitle : Return "标题幻灯片"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly : Return "仅标题"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText : Return "标题和内容"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTwoColumnText : Return "两栏内容"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank : Return "空白"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutContentWithCaption : Return "内容与标题"
            Case Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutPictureWithCaption : Return "图片与标题"
            Case Else : Return "自定义"
        End Select
    End Function

    Private Function GetPPTAlignmentString(alignment As Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment) As String
        Select Case alignment
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft : Return "left"
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter : Return "center"
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignRight : Return "right"
            Case Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignJustify : Return "justify"
            Case Else : Return "left"
        End Select
    End Function

    Private Function GetPlaceholderTypeName(shape As Microsoft.Office.Interop.PowerPoint.Shape) As String
        Try
            If shape.PlaceholderFormat Is Nothing Then Return "文本框"
            Select Case shape.PlaceholderFormat.Type
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderTitle : Return "标题"
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle : Return "居中标题"
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle : Return "副标题"
                Case Microsoft.Office.Interop.PowerPoint.PpPlaceholderType.ppPlaceholderBody : Return "正文"
                Case Else : Return "内容"
            End Select
        Catch
            Return "文本"
        End Try
    End Function

    Private Function EscapeForJs(text As String) As String
        Return text.Replace("\", "\\").Replace("`", "\`").Replace("$", "\$").Replace(vbCr, "").Replace(vbLf, "\n")
    End Function
End Class