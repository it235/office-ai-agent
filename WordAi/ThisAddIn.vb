Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports ShareRibbon
Public Class ThisAddIn

    Public Shared chatTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public Shared chatControl As ChatControl
    Private translateService As WordTranslateService

    Private captureTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public Shared dataCapturePane As WebDataCapturePane

    ' 在类中添加以下变量
    Private _deepseekControl As DeepseekControl
    Private _deepseekTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private _doubaoControl As DoubaoChat
    Private _doubaoTaskPane As Microsoft.Office.Tools.CustomTaskPane
    
    ' 模板编辑器
    Private _templateEditorControl As ReformatTemplateEditorControl
    Private _templateEditorTaskPane As Microsoft.Office.Tools.CustomTaskPane
    
    ' Word补全管理器
    Private _completionManager As WordCompletionManager

    Private Sub WordAi_Startup() Handles Me.Startup

        Try
            WebView2Loader.EnsureWebView2Loader()
        Catch ex As Exception
            MessageBox.Show($"WebView2 初始化失败: {ex.Message}")
        End Try

        ' 处理工作表和工作簿切换事件
        Application_WorkbookActivate()
        ' 初始化 Timer，用于WPS中扩大聊天区域的宽度
        widthTimer = New Timer()
        AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
        widthTimer.Interval = 100 ' 设置延迟时间，单位为毫秒
        ' 初始化 Timer，用于WPS中扩大聊天区域的宽度
        widthTimer1 = New Timer()
        AddHandler widthTimer1.Tick, AddressOf WidthTimer1_Tick
        widthTimer1.Interval = 100 ' 设置延迟时间，单位为毫秒

        translateService = New WordTranslateService()
        
        ' 预加载聊天设置（确保补全配置在CompletionManager初始化前已加载）
        Dim chatSettings As New ChatSettings(New ApplicationInfo("Word", OfficeApplicationType.Word))
        
        ' 初始化Word补全管理器（已禁用 - 观察期）
        ' InitializeCompletionManager()

    End Sub
    
    ''' <summary>
    ''' 初始化Word补全管理器（已禁用 - 观察期）
    ''' </summary>
    Private Sub InitializeCompletionManager()
        ' 补全功能已禁用，跳过初始化
        Debug.WriteLine("[Word] 补全管理器已跳过初始化（观察期）")
    End Sub
    
    ''' <summary>
    ''' 启用/禁用Word补全功能
    ''' </summary>
    Public Sub SetCompletionEnabled(enabled As Boolean)
        If _completionManager IsNot Nothing Then
            _completionManager.Enabled = enabled
        End If
    End Sub

    ''' <summary>
    ''' 自动补全设置保存事件处理
    ''' </summary>
    Private Sub OnAutocompleteSettingsSaved(sender As Object, e As AutocompleteSettingsSavedEventArgs)
        Try
            If _completionManager IsNot Nothing Then
                _completionManager.Enabled = e.EnableAutocomplete
                Debug.WriteLine($"[Word] 补全设置已同步: Enabled={e.EnableAutocomplete}")
            End If
        Catch ex As Exception
            Debug.WriteLine($"[Word] 同步补全设置失败: {ex.Message}")
        End Try
    End Sub


    Private Function IsWpsActive() As Boolean
        Try
            Return Process.GetProcessesByName("WPS").Length > 0
        Catch
            Return False
        End Try
    End Function


    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 补全功能已禁用，无需取消订阅
    End Sub


    ' 为新工作簿创建任务窗格
    Private Sub CreateChatTaskPane()
        Try
            chatControl = New ChatControl()
            chatTaskPane = Me.CustomTaskPanes.Add(chatControl, "Word AI智能助手")
                chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                chatTaskPane.Width = 420

        Catch ex As Exception
            MessageBox.Show($"初始化新建工作簿任务窗格失败: {ex.Message}")
        End Try
    End Sub

    '    ' 切换工作表时重新

    Private Sub Application_WorkbookActivate()
        Try
            ' 为新工作簿创建任务窗格
            dataCapturePane = New WebDataCapturePane()
            captureTaskPane = Me.CustomTaskPanes.Add(dataCapturePane, "Word爬虫")
            captureTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            captureTaskPane.Width = 420
            'AddHandler captureTaskPane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged
            captureTaskPane.Visible = False


        Catch ex As Exception
            MessageBox.Show($"初始化新建工作簿任务窗格失败: {ex.Message}")
        End Try
    End Sub

    Private widthTimer As Timer
    Private widthTimer1 As Timer
    ' 解决WPS中无法显示正常宽度的问题
    Private Sub ChatTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If IsWpsActive() Then
                widthTimer.Start()
            End If
        End If
    End Sub

    Private Sub DeepseekTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If IsWpsActive() Then
                widthTimer1.Start()
            End If
        End If
    End Sub
    Private Sub CreateDeepseekTaskPane()
        Try
            If _deepseekControl Is Nothing Then
                ' 为新工作簿创建任务窗格
                _deepseekControl = New DeepseekControl()
                _deepseekTaskPane = Me.CustomTaskPanes.Add(_deepseekControl, "Deepseek AI智能助手")
                _deepseekTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _deepseekTaskPane.Width = 420
                AddHandler _deepseekTaskPane.VisibleChanged, AddressOf DeepseekTaskPane_VisibleChanged
                _deepseekTaskPane.Visible = False
            End If
        Catch ex As Exception
            MessageBox.Show($"初始化任务窗格失败: {ex.Message}")
        End Try
    End Sub

    Private Async Function CreateDoubaoTaskPane() As Task
        Try
            If _doubaoControl Is Nothing Then
                ' 为新工作簿创建任务窗格
                _doubaoControl = New DoubaoChat()
                Await _doubaoControl.InitializeAsync()
                _doubaoTaskPane = Me.CustomTaskPanes.Add(_doubaoControl, "Doubao AI智能助手")
                _doubaoTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _doubaoTaskPane.Width = 420
            End If
        Catch ex As Exception
            MessageBox.Show($"初始化Doubao任务窗格失败: {ex.Message}")
        End Try
    End Function

    Private Sub WidthTimer_Tick(sender As Object, e As EventArgs)
        widthTimer.Stop()
        If IsWpsActive() AndAlso chatTaskPane IsNot Nothing Then
            chatTaskPane.Width = 420
        End If
    End Sub
    Private Sub WidthTimer1_Tick(sender As Object, e As EventArgs)
        widthTimer1.Stop()
        If IsWpsActive() AndAlso _deepseekTaskPane IsNot Nothing Then
            _deepseekTaskPane.Width = 420
        End If
    End Sub
    Private Sub AiHelper_Shutdown() Handles Me.Shutdown
        ' 清理资源
        'RemoveHandler Globals.ThisAddIn.Application.WorkbookActivate, AddressOf Me.Application_WorkbookActivate
    End Sub

    Dim loadChatHtml As Boolean = True
    Dim loadDataCaptureHtml As Boolean = True

    Public Async Sub ShowChatTaskPane()
        CreateChatTaskPane()
        If chatTaskPane Is Nothing Then Return
        chatTaskPane.Visible = True
        If loadChatHtml Then
            loadChatHtml = False
            Await chatControl.LoadLocalHtmlFile()
        End If
    End Sub

    Public Async Sub ShowDataCaptureTaskPane()
        If captureTaskPane Is Nothing Then Return
        captureTaskPane.Visible = True
    End Sub

    Public Async Sub ShowDeepseekTaskPane()
        CreateDeepseekTaskPane()
        If _deepseekTaskPane Is Nothing Then Return
        _deepseekTaskPane.Visible = True
    End Sub

    Public Async Sub ShowDoubaoTaskPane()
        Await CreateDoubaoTaskPane()
        If _doubaoTaskPane Is Nothing Then Return
        _doubaoTaskPane.Visible = True
    End Sub
    
    ''' <summary>
    ''' 显示模板编辑器任务窗格
    ''' </summary>
    Public Sub ShowTemplateEditorTaskPane(Optional template As ReformatTemplate = Nothing)
        Try
            ' 如果已存在，先关闭
            If _templateEditorTaskPane IsNot Nothing Then
                Try
                    Me.CustomTaskPanes.Remove(_templateEditorTaskPane)
                Catch
                End Try
                _templateEditorTaskPane = Nothing
                _templateEditorControl = Nothing
            End If
            
            ' 创建预览回调（安全处理 chatControl 可能为 Nothing 的情况）
            Dim previewCallback As PreviewStyleCallback = Nothing
            If chatControl IsNot Nothing Then
                previewCallback = AddressOf chatControl.ApplyStylePreviewToSelection
            End If
            
            ' 创建占位符预览回调
            Dim placeholderPreviewCallback As TemplatePlaceholderPreviewCallback = AddressOf ApplyPlaceholderPreviewToDocument
            
            ' 创建新的编辑器控件
            _templateEditorControl = New ReformatTemplateEditorControl(template, previewCallback, placeholderPreviewCallback)
            
            ' 绑定事件
            AddHandler _templateEditorControl.TemplateSaved, Sub(s, t)
                GlobalStatusStrip.ShowInfo($"模板「{t.Name}」已保存")
                HideTemplateEditorTaskPane()
                ' 刷新前端模板列表
                If chatControl IsNot Nothing Then
                    chatControl.RefreshReformatTemplates()
                End If
            End Sub
            
            AddHandler _templateEditorControl.EditorClosed, Sub(s, e)
                HideTemplateEditorTaskPane()
            End Sub
            
            ' 创建TaskPane
            _templateEditorTaskPane = Me.CustomTaskPanes.Add(_templateEditorControl, "排版模板编辑器")
            _templateEditorTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            _templateEditorTaskPane.Width = 380
            _templateEditorTaskPane.Visible = True
            
        Catch ex As Exception
            MessageBox.Show($"打开模板编辑器失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"ShowTemplateEditorTaskPane 错误: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 隐藏模板编辑器任务窗格
    ''' </summary>
    Public Sub HideTemplateEditorTaskPane()
        Try
            If _templateEditorTaskPane IsNot Nothing Then
                _templateEditorTaskPane.Visible = False
                Me.CustomTaskPanes.Remove(_templateEditorTaskPane)
                _templateEditorTaskPane = Nothing
                _templateEditorControl = Nothing
            End If
        Catch ex As Exception
            Debug.WriteLine($"HideTemplateEditorTaskPane 错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用占位符预览到文档
    ''' </summary>
    Private Sub ApplyPlaceholderPreviewToDocument(placeholderId As String, content As String, fontConfig As FontConfig, paragraphConfig As ParagraphConfig, colorConfig As ColorConfig)
        Try
            Dim wordApp = Me.Application
            If wordApp Is Nothing OrElse wordApp.Selection Is Nothing Then Return

            Dim selRange = wordApp.Selection.Range

            ' 清除当前选择内容
            selRange.Text = ""

            ' 应用占位符内容
            selRange.Text = content

            ' 应用字体设置
            If fontConfig IsNot Nothing Then
                If Not String.IsNullOrEmpty(fontConfig.FontNameCN) Then
                    selRange.Font.NameFarEast = fontConfig.FontNameCN
                End If
                If fontConfig.FontSize > 0 Then
                    selRange.Font.Size = CSng(fontConfig.FontSize)
                End If
                selRange.Font.Bold = If(fontConfig.Bold, 0, 0)
                selRange.Font.Italic = If(fontConfig.Italic, 0, 0)
                selRange.Font.Underline = If(fontConfig.Underline, 0, 0)
            End If

            ' 应用颜色设置
            If colorConfig IsNot Nothing AndAlso Not String.IsNullOrEmpty(colorConfig.FontColor) Then
                Try
                    Dim color As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml(colorConfig.FontColor)
                    selRange.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
                Catch ex As Exception
                    Debug.WriteLine($"应用颜色失败: {ex.Message}")
                End Try
            End If

            ' 应用段落设置
            If paragraphConfig IsNot Nothing Then
                Select Case paragraphConfig.Alignment?.ToLower()
                    Case "center"
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                    Case "right"
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                    Case "justify"
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                    Case Else
                        selRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                End Select

                If paragraphConfig.FirstLineIndent > 0 AndAlso selRange.Font.Size > 0 Then
                    selRange.ParagraphFormat.FirstLineIndent = CSng(paragraphConfig.FirstLineIndent * selRange.Font.Size)
                End If

                If paragraphConfig.LineSpacing > 0 Then
                    selRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple
                    selRange.ParagraphFormat.LineSpacing = CSng(paragraphConfig.LineSpacing * 12)
                End If
            End If

            Debug.WriteLine($"应用占位符预览: {placeholderId} -> {content}")

        Catch ex As Exception
            Debug.WriteLine($"ApplyPlaceholderPreviewToDocument 错误: {ex.Message}")
        End Try
    End Sub

End Class
