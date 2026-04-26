Imports System.Windows.Forms
Imports System.Drawing
Imports System.Timers

Public Module GlobalStatusStripAll
    Private WithEvents StatusTimer As New System.Timers.Timer()
    Private _application As Object = Nothing
    Private _messagePending As Boolean = False
    Private notificationForm As NotificationForm = Nothing

    Sub New()
        StatusTimer.Interval = 5000 ' 设置显示提示时间为5秒
        AddHandler StatusTimer.Elapsed, AddressOf StatusTimer_Elapsed
    End Sub

    Public Sub InitializeApplication(application As Object)
        _application = application
        Debug.WriteLine("GlobalStatusStrip已初始化，应用程序类型: " & application.GetType().FullName)
    End Sub

    Public Sub ShowWarningStatus(message As String)
        ShowWarning(message) ' 简化代码，使用同一个方法
    End Sub

    ''' <summary>
    ''' 显示进度信息（只更新状态栏，不弹窗）
    ''' </summary>
    Public Sub ShowProgress(message As String)
        Try
            Debug.WriteLine("[Progress] " & message)

            If _application IsNot Nothing Then
                Try
                    Dim app = TryCast(_application, Object)
                    app.StatusBar = "AI: " & message

                    ' 进度状态10秒后清除（比警告长，避免翻译过程中过早消失）
                    _messagePending = True
                    StatusTimer.Stop()
                    StatusTimer.Interval = 10000
                    StatusTimer.Start()
                Catch ex As Exception
                    Debug.WriteLine($"设置进度状态栏失败: {ex.Message}")
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("显示进度失败: " & ex.Message)
        End Try
    End Sub

    Public Sub ShowWarning(message As String)
        Try
            Debug.WriteLine("显示警告通知: " & message)

            ' 方法1: 直接使用应用程序的状态栏
            If _application IsNot Nothing Then
                Try
                    ' 使用动态类型，避免反射
                    Dim app = TryCast(_application, Object)
                    app.StatusBar = "AI: " & message

                    ' 启动定时器，5秒后清除状态栏
                    _messagePending = True
                    StatusTimer.Stop() ' 确保先停止之前的计时器
                    StatusTimer.Interval = 5000
                    StatusTimer.Start()

                    Debug.WriteLine("状态栏消息设置成功")
                Catch ex As Exception
                    Debug.WriteLine($"设置状态栏失败: {ex.Message}")
                End Try
            End If

            ' 方法2: 总是显示一个漂亮的通知窗口
            ShowNotificationForm(message)

            ' 备用方案：输出到调试
            Debug.WriteLine("AI提示: " & message)

        Catch ex As Exception
            Debug.WriteLine("显示通知失败: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 显示实时进度Toast（单实例，可更新内容）
    ''' </summary>
    Public Sub ShowToast(message As String)
        Try
            Debug.WriteLine("[Toast] " & message)

            ' 更新 Office 状态栏
            If _application IsNot Nothing Then
                Try
                    Dim app = TryCast(_application, Object)
                    app.StatusBar = "AI: " & message
                    _messagePending = True
                    StatusTimer.Stop()
                    StatusTimer.Interval = 10000
                    StatusTimer.Start()
                Catch ex As Exception
                End Try
            End If

            ' 更新或创建 Toast 弹框
            If notificationForm IsNot Nothing AndAlso Not notificationForm.IsDisposed Then
                Try
                    If notificationForm.InvokeRequired Then
                        notificationForm.Invoke(Sub()
                                                    notificationForm.UpdateMessage(message)
                                                    notificationForm.ResetTimer()
                                                End Sub)
                    Else
                        notificationForm.UpdateMessage(message)
                        notificationForm.ResetTimer()
                    End If
                Catch ex As Exception
                    notificationForm = Nothing
                    ShowNotificationForm(message)
                End Try
            Else
                ShowNotificationForm(message)
            End If
        Catch ex As Exception
            Debug.WriteLine("ShowToast failed: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 关闭当前 Toast 弹框
    ''' </summary>
    Public Sub CloseToast()
        Try
            If notificationForm IsNot Nothing AndAlso Not notificationForm.IsDisposed Then
                Try
                    If notificationForm.InvokeRequired Then
                        notificationForm.Invoke(Sub() notificationForm.Close())
                    Else
                        notificationForm.Close()
                    End If
                Catch ex As Exception
                End Try
                notificationForm = Nothing
            End If
        Catch ex As Exception
            Debug.WriteLine("CloseToast failed: " & ex.Message)
        End Try
    End Sub

    ' 显示自定义通知窗口
    Private Sub ShowNotificationForm(message As String)
        Try
            ' 如果已有通知窗体，先关闭
            If notificationForm IsNot Nothing AndAlso Not notificationForm.IsDisposed Then
                Try
                    notificationForm.Invoke(Sub()
                                                notificationForm.Close()
                                                notificationForm = Nothing
                                            End Sub)
                Catch
                    notificationForm = Nothing
                End Try
            End If

            ' 创建并显示通知窗口
            Dim thread As New System.Threading.Thread(
                Sub()
                    Try
                        notificationForm = New NotificationForm(message)
                        AddHandler notificationForm.FormClosed, Sub() notificationForm = Nothing
                        notificationForm.ShowDialog() ' 使用ShowDialog以保持线程运行直到窗体关闭
                    Catch ex As Exception
                        Debug.WriteLine($"显示通知窗口失败: {ex.Message}")
                    End Try
                End Sub)

            thread.SetApartmentState(System.Threading.ApartmentState.STA)
            thread.IsBackground = True ' 设置为后台线程，这样不会阻止应用程序退出
            thread.Start()
        Catch ex As Exception
            Debug.WriteLine($"创建通知线程失败: {ex.Message}")
        End Try
    End Sub

    Private Sub StatusTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Try
            StatusTimer.Stop()

            If Not _messagePending OrElse _application Is Nothing Then
                _messagePending = False
                Return
            End If

            Debug.WriteLine("计时器触发，准备清除状态栏")

            Try
                ' 清除状态栏的简单方法
                Dim app = TryCast(_application, Object)
                app.StatusBar = False
                Debug.WriteLine("状态栏已清除")
            Catch ex As Exception
                Debug.WriteLine($"清除状态栏失败: {ex.Message}")
            End Try

            _messagePending = False
        Catch ex As Exception
            Debug.WriteLine("定时器事件处理失败: " & ex.Message)
            _messagePending = False
        End Try
    End Sub
End Module

' 自定义通知窗体
Public Class NotificationForm : Inherits Form
    Private WithEvents closeTimer As New System.Windows.Forms.Timer()
    Private fadeTimer As New System.Windows.Forms.Timer()
    Private _opacity As Double = 1.0
    Private _lblMessage As Label

    Public Sub New(message As String)
        ' 设置窗体属性
        Me.FormBorderStyle = FormBorderStyle.None
        Me.StartPosition = FormStartPosition.Manual
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.Size = New Size(300, 80)
        Me.BackColor = Color.FromArgb(50, 50, 50)
        Me.Opacity = 0.9

        ' 圆角效果
        Dim path As New Drawing2D.GraphicsPath()
        path.AddArc(0, 0, 20, 20, 180, 90)
        path.AddArc(Me.Width - 20, 0, 20, 20, 270, 90)
        path.AddArc(Me.Width - 20, Me.Height - 20, 20, 20, 0, 90)
        path.AddArc(0, Me.Height - 20, 20, 20, 90, 90)
        Me.Region = New Region(path)

        ' 添加消息标签
        _lblMessage = New Label()
        _lblMessage.Text = message
        _lblMessage.ForeColor = Color.White
        _lblMessage.Font = New Font("Microsoft YaHei UI", 9.0F, FontStyle.Regular)
        _lblMessage.AutoSize = False
        _lblMessage.Size = New Size(280, 60)
        _lblMessage.Location = New Point(10, 10)
        _lblMessage.TextAlign = ContentAlignment.MiddleCenter
        Me.Controls.Add(_lblMessage)

        ' 设置窗体位置 - 右下角
        Dim screenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
        Me.Location = New Point(screenWidth - Me.Width - 20, screenHeight - Me.Height - 20)

        ' 设置自动关闭定时器（延长到8秒，给进度更新留出时间）
        closeTimer.Interval = 8000
        closeTimer.Start()

        ' 设置渐隐效果定时器
        fadeTimer.Interval = 50 ' 每50毫秒更新一次透明度
        AddHandler fadeTimer.Tick, AddressOf FadeTimer_Tick
    End Sub

    ''' <summary>
    ''' 更新弹框消息文本
    ''' </summary>
    Public Sub UpdateMessage(newMessage As String)
        If _lblMessage IsNot Nothing AndAlso Not _lblMessage.IsDisposed Then
            _lblMessage.Text = newMessage
        End If
    End Sub

    ''' <summary>
    ''' 重置关闭计时器（用于进度更新时延长显示）
    ''' </summary>
    Public Sub ResetTimer()
        closeTimer.Stop()
        fadeTimer.Stop()
        _opacity = 1.0
        Me.Opacity = 0.9
        closeTimer.Start()
    End Sub

    Private Sub CloseTimer_Tick(sender As Object, e As EventArgs) Handles closeTimer.Tick
        closeTimer.Stop()
        fadeTimer.Start() ' 开始渐隐效果
    End Sub

    Private Sub FadeTimer_Tick(sender As Object, e As EventArgs)
        _opacity -= 0.05
        If _opacity <= 0 Then
            fadeTimer.Stop()
            Me.Close()
        Else
            Me.Opacity = _opacity
        End If
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        ' 绘制边框
        Using g As Graphics = e.Graphics
            g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            Using pen As New Pen(Color.FromArgb(100, 100, 100), 1)
                g.DrawPath(pen, GetRoundedRectPath(Me.ClientRectangle, 10))
            End Using
        End Using
    End Sub

    Private Function GetRoundedRectPath(rect As Rectangle, radius As Integer) As Drawing2D.GraphicsPath
        Dim path As New Drawing2D.GraphicsPath()
        path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90)
        path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90)
        path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90)
        path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90)
        path.CloseFigure()
        Return path
    End Function
End Class