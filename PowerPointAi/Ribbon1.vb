' WordAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
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

    Protected Overrides Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show("PowerPoint排版功能正在开发中...", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

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
                ' 稍等一下让WebView2加载完成，然后触发续写对话框
                Task.Run(Async Function()
                             Await Task.Delay(300)
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
    Protected Overrides Sub AcceptCompletionButton_Click(sender As Object, e As RibbonControlEventArgs)
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
End Class