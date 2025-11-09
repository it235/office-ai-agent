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

    End Sub


    Private Function IsWpsActive() As Boolean
        Try
            Return Process.GetProcessesByName("WPS").Length > 0
        Catch
            Return False
        End Try
    End Function


    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
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
        chatTaskPane.Visible = True
        If loadChatHtml Then
            loadChatHtml = False
            Await chatControl.LoadLocalHtmlFile()
        End If
    End Sub

    Public Async Sub ShowDataCaptureTaskPane()
        captureTaskPane.Visible = True
    End Sub

    Public Async Sub ShowDeepseekTaskPane()
        CreateDeepseekTaskPane()
        _deepseekTaskPane.Visible = True
    End Sub
End Class
