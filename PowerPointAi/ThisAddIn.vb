Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports ShareRibbon
Public Class ThisAddIn

    Private chatTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Public Shared chatControl As ChatControl
    ' 翻译服务：延迟初始化，首次使用时创建
    Private _translateService As PowerPointTranslateService

    ' 在类中添加以下变量
    Private _deepseekControl As DeepseekControl
    Private _deepseekTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private _doubaoControl As DoubaoChat
    Private _doubaoTaskPane As Microsoft.Office.Tools.CustomTaskPane

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

    Private Sub PowerPointAi_Startup() Handles Me.Startup
        ' Phase 0: 仅注册事件处理器（微秒级，不阻塞启动）
        PhaseStartupManager.Instance.RunCriticalPhase(Me.Application)
    End Sub

    ''' <summary>
    ''' 确保核心服务已加载（WebView2 + SQLite），首次调用时初始化
    ''' </summary>
    Private Sub EnsureCoreServicesLoaded()
        If PhaseStartupManager.Instance.IsBackgroundReady Then Return
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
                'Await _doubaoControl.InitializeAsync()
                _doubaoTaskPane = Me.CustomTaskPanes.Add(_doubaoControl, "Doubao AI智能助手")
                _doubaoTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
                _doubaoTaskPane.Width = 420
            End If
        Catch ex As Exception
            MessageBox.Show($"初始化Doubao任务窗格失败: {ex.Message}")
        End Try
    End Function

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
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
    End Sub

    ' 创建聊天任务窗格
    Private Sub CreateChatTaskPane()
        Try
            ' 为新工作簿创建任务窗格
            chatControl = New ChatControl()
            chatTaskPane = Me.CustomTaskPanes.Add(chatControl, "PPT AI智能助手")
            chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            chatTaskPane.Width = 420
        Catch ex As Exception
            MessageBox.Show($"初始化 PPT AI 任务窗格失败: {ex.Message}")
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

    Dim loadChatHtml As Boolean = True

    Public Async Sub ShowChatTaskPane()
        EnsureCoreServicesLoaded()
        CreateChatTaskPane()
        If chatTaskPane Is Nothing Then Return
        chatTaskPane.Visible = True
        If loadChatHtml Then
            loadChatHtml = False
            Await chatControl.LoadLocalHtmlFile()
        End If
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
End Class
