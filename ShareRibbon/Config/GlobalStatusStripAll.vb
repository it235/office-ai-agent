Imports System.Windows.Forms
Imports System.Drawing
Imports System.Timers
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports Timer = System.Windows.Forms.Timer

Public Module GlobalStatusStripAll
    Private WithEvents StatusTimer As New System.Timers.Timer()
    Private _application As Object = Nothing
    Private _messagePending As Boolean = False

    ' 用于检查我们是否在UI线程上的变量
    Private _mainThreadId As Integer = Threading.Thread.CurrentThread.ManagedThreadId

    Sub New()
        StatusTimer.Interval = 5000 ' 设置显示提示时间为5秒
        AddHandler StatusTimer.Elapsed, AddressOf StatusTimer_Elapsed
    End Sub

    Public Sub InitializeApplication(application As Object)
        _application = application
        Debug.WriteLine("GlobalStatusStrip已初始化，应用程序类型: " & application.GetType().FullName)
    End Sub
    Public Sub ShowWarning(message As String)
        Try
            Debug.WriteLine("显示警告通知: " & Message)

            ' 创建并显示简单的非模态通知
            Dim showAction As Action = Sub()
                                           ' 创建通知窗口 - 调整大小更小一些
                                           Dim notification As New Form() With {
                .Text = "提示信息 (3)",  ' 初始显示3秒倒计时
                .FormBorderStyle = FormBorderStyle.FixedToolWindow,
                .StartPosition = FormStartPosition.Manual,
                .ShowInTaskbar = False,
                .TopMost = True,
                .Size = New Size(280, 80),  ' 调整更小的尺寸
                .BackColor = Color.FromArgb(255, 240, 240),
                .Opacity = 0.85  ' 增加透明度
            }

                                           ' 创建消息标签
                                           Dim lblMessage As New Label() With {
                .Text = Message,
                .ForeColor = Color.DarkRed,
                .Font = New Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular),  ' 略微调小字体
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter
            }

                                           notification.Controls.Add(lblMessage)

                                           ' 获取Excel窗口位置
                                           Try
                                               Dim excelRect As New RECT()
                                               GetWindowRect(Process.GetCurrentProcess().MainWindowHandle, excelRect)

                                               ' 定位在Excel窗口右下角
                                               notification.Left = excelRect.right - notification.Width - 20
                                               notification.Top = excelRect.bottom - notification.Height - 40  ' 改为底部位置
                                           Catch ex As Exception
                                               ' 如果获取Excel窗口位置失败，则居中显示
                                               notification.StartPosition = FormStartPosition.CenterScreen
                                           End Try

                                           ' 显示通知并设置自动关闭
                                           notification.Show()

                                           ' 倒计时变量
                                           Dim remainingSeconds As Integer = 3

                                           ' 设置倒计时计时器
                                           Dim countdownTimer As New Timer() With {
                .Interval = 1000  ' 每秒更新一次
            }

                                           ' 设置自动关闭计时器
                                           Dim closeTimer As New Timer() With {
                .Interval = 3000  ' 3秒后关闭
            }

                                           ' 倒计时处理
                                           AddHandler countdownTimer.Tick, Sub(s, e)
                                                                               remainingSeconds -= 1
                                                                               If remainingSeconds > 0 Then
                                                                                   notification.Text = $"提示信息 ({remainingSeconds})"
                                                                               Else
                                                                                   countdownTimer.Stop()
                                                                               End If
                                                                           End Sub

                                           ' 关闭处理
                                           AddHandler closeTimer.Tick, Sub(s, e)
                                                                           closeTimer.Stop()
                                                                           countdownTimer.Stop()
                                                                           notification.Close()
                                                                           notification.Dispose()
                                                                       End Sub

                                           ' 启动计时器
                                           countdownTimer.Start()
                                           closeTimer.Start()
                                       End Sub

            ' 在UI线程上执行
            If Threading.Thread.CurrentThread.ManagedThreadId = _mainThreadId Then
                showAction.Invoke()
            Else
                Dim form As New Form()
                form.Invoke(showAction)
            End If

        Catch ex As Exception
            Debug.WriteLine("显示通知失败: " & ex.Message)
        End Try
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    <DllImport("user32.dll")>
    Private Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    Private Sub UpdateStatusBarDirectly(message As String)
        Try
            ' 使用动态类型处理不同的Office应用程序
            Debug.WriteLine("正在设置状态栏文本: " & message)

            ' 直接访问Excel/Word/PowerPoint的StatusBar属性
            Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
            If propertyInfo IsNot Nothing Then
                propertyInfo.SetValue(_application, "提示：" & message, Nothing)
                Debug.WriteLine("状态栏消息设置成功")
            Else
                Debug.WriteLine("无法找到StatusBar属性")
                ' 尝试另一种方法
                Try
                    ' 某些Office版本可能使用不同的方法
                    _application.GetType().InvokeMember("StatusBar",
                        Reflection.BindingFlags.SetProperty,
                        Nothing,
                        _application,
                        New Object() {"提示：" & message})
                    Debug.WriteLine("通过反射设置状态栏消息成功")
                Catch ex As Exception
                    Debug.WriteLine("通过反射设置状态栏失败: " & ex.Message)
                    Throw ' 重新抛出异常以便上层处理
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("更新状态栏失败: " & ex.Message)
            Throw ' 重新抛出异常以便上层处理
        End Try
    End Sub

    Private Sub StatusTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Try
            If Not _messagePending Then Return

            Debug.WriteLine("计时器触发，准备清除状态栏")

            If _application Is Nothing Then
                StatusTimer.Stop()
                _messagePending = False
                Return
            End If

            ' 尝试在UI线程上执行
            Dim form As New Form()
            form.Invoke(New Action(Sub()
                                       Try
                                           ' 尝试清除状态栏
                                           Debug.WriteLine("正在清除状态栏")

                                           ' 使用反射设置StatusBar属性
                                           Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
                                           If propertyInfo IsNot Nothing Then
                                               propertyInfo.SetValue(_application, False, Nothing)
                                               Debug.WriteLine("状态栏已清除")
                                           Else
                                               ' 尝试另一种方法
                                               Try
                                                   _application.GetType().InvokeMember("StatusBar",
                                                       Reflection.BindingFlags.SetProperty,
                                                       Nothing,
                                                       _application,
                                                       New Object() {False})
                                                   Debug.WriteLine("通过反射清除状态栏成功")
                                               Catch ex As Exception
                                                   Debug.WriteLine("通过反射清除状态栏失败: " & ex.Message)
                                               End Try
                                           End If
                                       Catch ex As Exception
                                           Debug.WriteLine("清除状态栏失败: " & ex.Message)
                                       End Try
                                   End Sub))

            StatusTimer.Stop()
            _messagePending = False
        Catch ex As Exception
            Debug.WriteLine("定时器事件处理失败: " & ex.Message)
            StatusTimer.Stop()
            _messagePending = False
        End Try
    End Sub
End Module