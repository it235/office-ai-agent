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

    Private Sub WordAi_Startup() Handles Me.Startup
        ' 初始化 GlobalStatusStrip
        Try
            Debug.WriteLine("正在初始化GlobalStatusStrip...")
            GlobalStatusStripAll.InitializeApplication(Me.Application)
            Debug.WriteLine("GlobalStatusStrip初始化完成")

            ' 测试状态栏是否正常工作
            'GlobalStatusStripAll.ShowWarning("Excel加载项已启动")
        Catch ex As Exception
            Debug.WriteLine("初始化GlobalStatusStrip时出错: " & ex.Message)
            MessageBox.Show("初始化状态栏时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            WebView2Loader.EnsureWebView2Loader()
        Catch ex As Exception
            MessageBox.Show($"WebView2 初始化失败: {ex.Message}")
        End Try

        ' 处理工作表和工作簿切换事件
        'AddHandler Globals.ThisAddIn.Application.ActiveDocument, AddressOf Me.Application_WorkbookActivate
        Application_WorkbookActivate()
        ' 初始化 Timer，用于WPS中扩大聊天区域的宽度
        widthTimer = New Timer()
        AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
        widthTimer.Interval = 100 ' 设置延迟时间，单位为毫秒

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


    '    ' 切换工作表时重新

    Private Sub Application_WorkbookActivate()
        Try
            ' 为新工作簿创建任务窗格
            chatControl = New ChatControl()
            chatTaskPane = Me.CustomTaskPanes.Add(chatControl, "Excel AI智能助手")
            chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            chatTaskPane.Width = 420
            AddHandler chatTaskPane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged
            chatTaskPane.Visible = False

        Catch ex As Exception
            MessageBox.Show($"初始化新建工作簿任务窗格失败: {ex.Message}")
        End Try
    End Sub

    Private widthTimer As Timer
    ' 解决WPS中无法显示正常宽度的问题
    Private Sub ChatTaskPane_VisibleChanged(sender As Object, e As EventArgs)
        Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = CType(sender, Microsoft.Office.Tools.CustomTaskPane)
        If taskPane.Visible Then
            If IsWpsActive() Then
                widthTimer.Start()
            End If
        End If
    End Sub

    Private Sub WidthTimer_Tick(sender As Object, e As EventArgs)
        widthTimer.Stop()
        If IsWpsActive() AndAlso chatTaskPane IsNot Nothing Then
            chatTaskPane.Width = 420
        End If
    End Sub
    Private Sub AiHelper_Shutdown() Handles Me.Shutdown
        ' 清理资源
        'RemoveHandler Globals.ThisAddIn.Application.WorkbookActivate, AddressOf Me.Application_WorkbookActivate
    End Sub

    Dim loadChatHtml As Boolean = True

    Public Async Sub ShowChatTaskPane()
        chatTaskPane.Visible = True
        If loadChatHtml Then
            loadChatHtml = False
            Await chatControl.LoadLocalHtmlFile()
        End If
    End Sub
End Class
