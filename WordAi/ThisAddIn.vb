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
    ' 翻译服务：延迟初始化，首次使用时创建

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

    ' 延迟初始化：WebView2 和 SQLite 仅在首次使用时加载
    Private _lazyWebView2 As New Lazy(Of Boolean)(Function()
        WebView2Loader.EnsureWebView2Loader()
        Return True
    End Function)

    Private _lazySqlite As New Lazy(Of Boolean)(Function()
        SqliteNativeLoader.EnsureLoaded()
        Return True
    End Function)

    ' WPS 宽度修复定时器
    Private widthTimer As Timer
    Private widthTimer1 As Timer

    ' E4-01 选区工具条控制器（随面板首次打开才创建，不进 Startup 关键路径）
    Private selectionToolbarController As WordAi.Selection.WordSelectionToolbarController

    ' E4-02 右键菜单入口（同上：面板打开过才挂，关闭加载项时必须按 Tag 摘干净）
    Private contextMenu As WordAi.HostEntry.WordContextMenu

    Private Sub WordAi_Startup() Handles Me.Startup
        ' Phase 0: 仅注册事件处理器 + 状态栏初始化（微秒级，不阻塞启动）
        PhaseStartupManager.Instance.RunCriticalPhase(Me.Application)
    End Sub

    ''' <summary>
    ''' 选区工具条只在用户已经打开 AI 面板后才存在：面板没打开时点了动作没有承接方，
    ''' 同时保证没用过该功能的用户不受浮层影响、启动耗时也不变。
    ''' </summary>
    Private Sub EnsureSelectionToolbar()
        Try
            If selectionToolbarController Is Nothing Then
                selectionToolbarController = New WordAi.Selection.WordSelectionToolbarController(Me.Application)
            End If
            ' 控制器内部有幂等守卫，重复打开面板不会重复绑定选区事件
            selectionToolbarController.Initialize()
        Catch ex As Exception
            Debug.WriteLine($"Word 选区工具条初始化失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' E4-02：把 Word 的 AI 动作挂到文本右键菜单上。同样只在面板打开过之后挂：
    ''' 菜单项点下去要有承接方，而且改宿主菜单属于锦上添花，不能让没用过 AI 的用户被改界面。
    ''' </summary>
    Private Sub EnsureContextMenu()
        If Not ChatSettings.EnableContextMenu Then Return

        Try
            If contextMenu Is Nothing Then
                contextMenu = New WordAi.HostEntry.WordContextMenu(Me.Application)
            End If
            contextMenu.Attach(AddressOf OnContextMenuAction)
        Catch ex As Exception
            Debug.WriteLine($"Word 右键菜单入口初始化失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 菜单项只把入口动作交回面板的统一入口：闸门、上下文采集、预览与风险确认、变更摘要、撤销都在主链上。
    ''' </summary>
    Private Sub OnContextMenuAction(action As ShareRibbon.Agent.QuickAction)
        Try
            If chatControl Is Nothing Then Return
            If chatTaskPane IsNot Nothing Then chatTaskPane.Visible = True
            chatControl.StartAgentFromHostEntry(action)
        Catch ex As Exception
            Debug.WriteLine($"Word 右键菜单动作执行失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 确保核心服务已加载（WebView2 + SQLite），首次调用时初始化
    ''' 如果 Phase 2 后台预加载已完成，则直接跳过
    ''' </summary>
    Private Sub EnsureCoreServicesLoaded()
        Try
            Dim webView2Init = _lazyWebView2.Value
        Catch ex As Exception
            MessageBox.Show($"WebView2 初始化失败: {ex.Message}")
        End Try
        Try
            Dim sqliteInit = _lazySqlite.Value
        Catch ex As Exception
            MessageBox.Show($"SQLite 原生库加载失败，Skills/记忆功能可能不可用: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 确保 WPS 宽度修复定时器已初始化（仅在需要时创建）
    ''' </summary>
    Private Sub EnsureWidthTimers()
        If widthTimer Is Nothing Then
            widthTimer = New Timer()
            AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
            widthTimer.Interval = 100
        End If
        If widthTimer1 Is Nothing Then
            widthTimer1 = New Timer()
            AddHandler widthTimer1.Tick, AddressOf WidthTimer1_Tick
            widthTimer1.Interval = 100
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 清理翻译服务资源（右键菜单按钮）
        ' 清理定时器资源
        If widthTimer IsNot Nothing Then
            widthTimer.Stop()
            widthTimer.Dispose()
            widthTimer = Nothing
        End If
        If widthTimer1 IsNot Nothing Then
            widthTimer1.Stop()
            widthTimer1.Dispose()
            widthTimer1 = Nothing
        End If

        ' 选区工具条：解除 WindowSelectionChange 订阅并销毁浮层，否则关闭加载项后事件仍挂在 Word 上
        If selectionToolbarController IsNot Nothing Then
            Try
                selectionToolbarController.Dispose()
            Catch ex As Exception
                Debug.WriteLine($"Word 选区工具条释放失败: {ex.Message}")
            End Try
            selectionToolbarController = Nothing
        End If

        ' 右键菜单：按 Tag 摘掉本加载项挂上去的菜单项并解绑事件，否则宿主菜单里会留下死项
        If contextMenu IsNot Nothing Then
            Try
                contextMenu.Detach()
            Catch ex As Exception
                Debug.WriteLine($"Word 右键菜单入口释放失败: {ex.Message}")
            End Try
            contextMenu = Nothing
        End If
    End Sub


    ' 为新工作簿创建任务窗格
    Private Sub CreateChatTaskPane()
        Try
            If chatControl IsNot Nothing AndAlso chatTaskPane IsNot Nothing Then
                Return
            End If

            chatControl = New ChatControl()
            chatTaskPane = Me.CustomTaskPanes.Add(chatControl, "Word AI智能助手")
            chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            chatTaskPane.Width = 420
            AddHandler chatTaskPane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged

        Catch ex As Exception
            MessageBox.Show($"初始化 Word AI 任务窗格失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 创建网页爬虫任务窗格（延迟初始化，仅在用户点击爬虫按钮时创建）
    ''' </summary>
    Private Sub EnsureDataCapturePaneCreated()
        If dataCapturePane IsNot Nothing Then Return
        Try
            dataCapturePane = New WebDataCapturePane()
            captureTaskPane = Me.CustomTaskPanes.Add(dataCapturePane, "Word爬虫")
            captureTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            captureTaskPane.Width = 420
            captureTaskPane.Visible = False
        Catch ex As Exception
            MessageBox.Show($"初始化 Word 爬虫任务窗格失败: {ex.Message}")
        End Try
    End Sub

    ' 解决WPS中无法显示正常宽度的问题
    Private Sub ChatTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If LLMUtil.IsWpsActive() Then
                EnsureWidthTimers()
                widthTimer.Start()
            End If
            ' 挂在可见性事件上而不是 ShowChatTaskPane 里：用户也能用任务窗格自己的折叠按钮打开面板，
            ' 那条路不经过我们的方法；启动阶段面板不可见，所以不会把浮层创建拖进 Startup 关键路径。
            ' E4-05：这两个入口都读配置文件里的开关，而此刻面板可能还没装载过任何 ChatSettings，
            ' 读到的就是代码里的默认值——用户开过的开关在进程内第一次不生效。先确保加载过一次。
            ChatSettings.EnsureLoaded(New ApplicationInfo("Word", OfficeApplicationType.Word))
            EnsureSelectionToolbar()
            EnsureContextMenu()
        ElseIf selectionToolbarController IsNot Nothing Then
            ' 面板收起后动作条点下去没有承接方，必须一起收掉
            selectionToolbarController.HideStrip()
        End If
    End Sub

    Private Sub DeepseekTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If LLMUtil.IsWpsActive() Then
                EnsureWidthTimers()
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
        If LLMUtil.IsWpsActive() AndAlso chatTaskPane IsNot Nothing Then
            chatTaskPane.Width = 420
        End If
    End Sub
    Private Sub WidthTimer1_Tick(sender As Object, e As EventArgs)
        widthTimer1.Stop()
        If LLMUtil.IsWpsActive() AndAlso _deepseekTaskPane IsNot Nothing Then
            _deepseekTaskPane.Width = 420
        End If
    End Sub

    Public Sub ShowChatTaskPane()
        EnsureCoreServicesLoaded()
        CreateChatTaskPane()
        If chatTaskPane Is Nothing Then Return
        chatTaskPane.Visible = True
    End Sub

    Public Sub ShowDataCaptureTaskPane()
        EnsureDataCapturePaneCreated()
        If captureTaskPane Is Nothing Then Return
        captureTaskPane.Visible = True
    End Sub

    Public Sub ShowDeepseekTaskPane()
        EnsureCoreServicesLoaded()
        CreateDeepseekTaskPane()
        If _deepseekTaskPane Is Nothing Then Return
        _deepseekTaskPane.Visible = True
    End Sub

    Public Async Sub ShowDoubaoTaskPane()
        EnsureCoreServicesLoaded()
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
