Imports System.Windows.Forms
Imports System.Drawing
Imports System.Timers

Public Module GlobalTips
    Private WithEvents StatusTimer As New System.Timers.Timer()
    Private _application As Object = Nothing
    Private _messagePending As Boolean = False
    Private notificationForm As NotificationForm2 = Nothing

    Sub New()
        StatusTimer.Interval = 5000 ' 默认状态栏显示时间为5秒
        AddHandler StatusTimer.Elapsed, AddressOf StatusTimer_Elapsed
    End Sub

    Public Sub InitializeApplication(application As Object)
        _application = application
        Debug.WriteLine("GlobalStatusStrip已初始化应用对象: " & application.GetType().FullName)
    End Sub

    Public Sub ShowWarningStatus(message As String)
        ShowWarning(message) ' 兼容旧接口，使用默认行为
    End Sub

    ' 新的 ShowWarning：允许指定弹窗停留时间（毫秒）和是否需要手动关闭
    Public Sub ShowWarning(message As String, Optional durationMs As Integer = 3000, Optional requireManualClose As Boolean = False)
        Try
            Debug.WriteLine("显示通知: " & message)

            ' 方案1: 使用宿主应用状态栏显示
            If _application IsNot Nothing Then
                Try
                    Dim app = TryCast(_application, Object)
                    app.StatusBar = "AI: " & message

                    ' 设置计时器控制状态栏恢复
                    _messagePending = True
                    StatusTimer.Stop()
                    StatusTimer.Start()

                    Debug.WriteLine("状态栏信息设置成功")
                Catch ex As Exception
                    Debug.WriteLine($"设置状态栏失败: {ex.Message}")
                End Try
            End If

            ' 方案2: 弹出自定义通知窗体（异步、不会阻塞主线程）
            ShowNotificationForm(message, durationMs, requireManualClose)

            Debug.WriteLine("AI显示: " & message)
        Catch ex As Exception
            Debug.WriteLine("显示通知失败: " & ex.Message)
        End Try
    End Sub

    ' 兼容旧方法（不带参数）
    Public Sub ShowWarning(message As String)
        ShowWarning(message, 3000, False)
    End Sub

    ' 弹出自定义通知窗体（在单独的 STA 线程中 ShowDialog，不阻塞主线程）
    Private Sub ShowNotificationForm(message As String, durationMs As Integer, requireManualClose As Boolean)
        Try
            Dim thread As New System.Threading.Thread(
                Sub()
                    Try
                        Dim form As New NotificationForm2(message, durationMs, requireManualClose)
                        form.ShowDialog()
                    Catch ex As Exception
                        Debug.WriteLine($"显示通知窗体失败: {ex.Message}")
                    End Try
                End Sub)

            thread.SetApartmentState(System.Threading.ApartmentState.STA)
            thread.IsBackground = True
            thread.Start()
        Catch ex As Exception
            Debug.WriteLine($"启动通知线程失败: {ex.Message}")
        End Try
    End Sub

    Private Sub StatusTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Try
            StatusTimer.Stop()

            If Not _messagePending OrElse _application Is Nothing Then
                _messagePending = False
                Return
            End If

            Debug.WriteLine("定时器触发，准备清除状态栏")

            Try
                Dim app = TryCast(_application, Object)
                app.StatusBar = False
                Debug.WriteLine("状态栏已清除")
            Catch ex As Exception
                Debug.WriteLine($"清除状态栏失败: {ex.Message}")
            End Try

            _messagePending = False
        Catch ex As Exception
            Debug.WriteLine("定时器处理失败: " & ex.Message)
            _messagePending = False
        End Try
    End Sub
End Module

' 自定义通知窗体：支持可复制文本、长内容滚动、可选自动消失或手动关闭、淡出动画
Public Class NotificationForm2 : Inherits Form
    Private WithEvents closeTimer As New System.Windows.Forms.Timer()
    Private fadeTimer As New System.Windows.Forms.Timer()
    Private _opacity As Double = 1.0

    Public Sub New(message As String)
        Me.New(message, 3000, False)
    End Sub

    Public Sub New(message As String, durationMs As Integer, requireManualClose As Boolean)
        ' 无边框、托盘外显示、置顶
        Me.FormBorderStyle = FormBorderStyle.None
        Me.StartPosition = FormStartPosition.Manual
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.BackColor = Color.FromArgb(50, 50, 50)
        Me.Opacity = 0.95

        ' 初始尺寸（可根据内容动态调整高度，保留最大高度以便滚动）
        Dim width = 360
        Dim maxHeight = 200
        Dim preferredHeight = 120
        Me.Size = New Size(width, Math.Min(preferredHeight, maxHeight))

        ' 圆角
        Dim path As New Drawing2D.GraphicsPath()
        path.AddArc(0, 0, 20, 20, 180, 90)
        path.AddArc(Me.Width - 20, 0, 20, 20, 270, 90)
        path.AddArc(Me.Width - 20, Me.Height - 20, 20, 20, 0, 90)
        path.AddArc(0, Me.Height - 20, 20, 20, 90, 90)
        Me.Region = New Region(path)

        ' 可复制、多行、带垂直滚动条的文本控件（RichTextBox）
        Dim rtb As New RichTextBox()
        rtb.ReadOnly = True
        rtb.BorderStyle = BorderStyle.None
        rtb.BackColor = Me.BackColor
        rtb.ForeColor = Color.White
        rtb.Font = New Font("Microsoft YaHei UI", 9.0F, FontStyle.Regular)
        rtb.Location = New Point(10, 10)
        rtb.Size = New Size(width - 20, Me.Height - 20)
        rtb.ScrollBars = RichTextBoxScrollBars.Vertical
        rtb.Text = message
        rtb.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        rtb.DetectUrls = False

        ' 根据内容计算需要的高度（若超过最大高度则保持滚动）
        Using g As Graphics = rtb.CreateGraphics()
            Dim textSize = g.MeasureString(message, rtb.Font, rtb.Width)
            Dim neededHeight = CInt(Math.Ceiling(textSize.Height)) + 20
            Dim finalHeight = Math.Min(Math.Max(80, neededHeight), maxHeight)
            Me.Size = New Size(width, finalHeight + 20)
            ' 重新设置圆角区域和 richtextbox 尺寸
            Dim p As New Drawing2D.GraphicsPath()
            p.AddArc(0, 0, 20, 20, 180, 90)
            p.AddArc(Me.Width - 20, 0, 20, 20, 270, 90)
            p.AddArc(Me.Width - 20, Me.Height - 20, 20, 20, 0, 90)
            p.AddArc(0, Me.Height - 20, 20, 20, 90, 90)
            Me.Region = New Region(p)

            rtb.Size = New Size(width - 20, Me.Height - 40)
        End Using

        Me.Controls.Add(rtb)

        ' 上下文菜单：复制、全选
        Dim ctx As New ContextMenuStrip()
        ctx.Items.Add("复制", Nothing, Sub(s, e) If rtb.SelectedText.Length > 0 Then rtb.Copy() Else Clipboard.SetText(rtb.Text))
        ctx.Items.Add("全选", Nothing, Sub(s, e) rtb.SelectAll())
        rtb.ContextMenuStrip = ctx

        ' 如果允许手动关闭，显示一个关闭按钮；否则使用定时器自动关闭
        Dim closeButton As New Button()
        closeButton.Text = "关闭"
        closeButton.Font = New Font("Microsoft YaHei UI", 8.0F)
        closeButton.FlatStyle = FlatStyle.Flat
        closeButton.FlatAppearance.BorderSize = 0
        closeButton.BackColor = Color.FromArgb(80, 80, 80)
        closeButton.ForeColor = Color.White
        closeButton.Size = New Size(56, 22)
        closeButton.Location = New Point(Me.Width - closeButton.Width - 12, Me.Height - closeButton.Height - 8)
        closeButton.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        closeButton.Visible = requireManualClose
        AddHandler closeButton.Click, Sub(s, e)
                                          closeTimer.Stop()
                                          fadeTimer.Start()
                                      End Sub
        Me.Controls.Add(closeButton)

        ' 将文本区高度再调整，以避免遮挡关闭按钮
        If requireManualClose Then
            rtb.Size = New Size(Me.Width - 20, Me.Height - 50)
        End If

        ' 位置：右下角
        Dim screenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
        Me.Location = New Point(screenWidth - Me.Width - 20, screenHeight - Me.Height - 20)

        ' 关闭定时器配置
        If requireManualClose Then
            closeTimer.Stop()
        Else
            closeTimer.Interval = Math.Max(500, durationMs) ' 最小500ms
            closeTimer.Start()
        End If

        ' 淡出动画定时器
        fadeTimer.Interval = 50
        AddHandler fadeTimer.Tick, AddressOf FadeTimer_Tick

        ' 鼠标进入时阻止自动关闭（可选：增强可读性）
        AddHandler Me.MouseEnter, Sub() If Not requireManualClose Then closeTimer.Stop()
        AddHandler Me.MouseLeave, Sub() If Not requireManualClose Then closeTimer.Start()
        AddHandler rtb.MouseEnter, Sub() If Not requireManualClose Then closeTimer.Stop()
        AddHandler rtb.MouseLeave, Sub() If Not requireManualClose Then closeTimer.Start()
    End Sub

    Private Sub CloseTimer_Tick(sender As Object, e As EventArgs) Handles closeTimer.Tick
        closeTimer.Stop()
        fadeTimer.Start()
    End Sub

    Private Sub FadeTimer_Tick(sender As Object, e As EventArgs)
        _opacity -= 0.06
        If _opacity <= 0 Then
            fadeTimer.Stop()
            Try
                Me.Close()
            Catch ex As Exception
                Debug.WriteLine("关闭通知窗体时出错: " & ex.Message)
            End Try
        Else
            Me.Opacity = _opacity
        End If
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        ' 边框
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