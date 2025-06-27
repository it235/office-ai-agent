Imports System.Windows.Forms
Imports System.Drawing
Imports System.Timers
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Imports System.Diagnostics
Imports Timer = System.Windows.Forms.Timer

Public Module GlobalStatusStripAll
    Private WithEvents StatusTimer As New System.Timers.Timer()
    Private _application As Object = Nothing
    Private _messagePending As Boolean = False

    ' 用于检查我们是否在UI线程上的变量
    Private _mainThreadId As Integer = Threading.Thread.CurrentThread.ManagedThreadId

    Sub New()
        StatusTimer.Interval = 8000 ' 设置显示提示时间为8秒
        AddHandler StatusTimer.Elapsed, AddressOf StatusTimer_Elapsed
    End Sub

    Public Sub InitializeApplication(application As Object)
        _application = application
        Debug.WriteLine("GlobalStatusStrip已初始化，应用程序类型: " & application.GetType().FullName)
    End Sub

    Public Sub ShowWarningStatus(message As String)
        Try
            Debug.WriteLine("显示警告通知: " & message)
            Debug.WriteLine($"_application状态: {If(_application Is Nothing, "Nothing", "已初始化")}")

            ' 尝试更新状态栏 - 不使用Excel-DNA特有的方法
            If _application IsNot Nothing Then
                Try
                    ' 直接尝试设置，在主线程上可能成功
                    _application.StatusBar = "AI: " & message
                    Debug.WriteLine("直接设置状态栏成功")
                    Return
                Catch directEx As Exception
                    Debug.WriteLine($"直接设置失败: {directEx.Message}")

                    ' 备用：使用反射方式
                    Try
                        Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
                        If propertyInfo IsNot Nothing Then
                            propertyInfo.SetValue(_application, "AI: " & message, Nothing)
                            Debug.WriteLine("反射设置状态栏成功")
                            Return
                        End If
                    Catch reflectionEx As Exception
                        Debug.WriteLine($"反射设置失败: {reflectionEx.Message}")
                    End Try
                End Try
            Else
                Debug.WriteLine("_application为Nothing，无法设置状态栏")
            End If

            ' 备用方案：输出到调试
            Console.WriteLine("AI提示: " & message)
            Debug.WriteLine("AI提示: " & message)

        Catch ex As Exception
            Debug.WriteLine("显示通知失败: " & ex.Message)
        End Try
    End Sub

    Public Sub ShowWarning(message As String)
        Try
            Debug.WriteLine("显示警告通知: " & message)

            ' 首先尝试更新Excel状态栏（非阻塞方式）
            Try
                UpdateStatusBarDirectly(message)
                Return ' 如果状态栏更新成功，就不显示弹窗了
            Catch ex As Exception
                Debug.WriteLine("状态栏更新失败，改用非阻塞通知: " & ex.Message)
            End Try

            ' 如果状态栏更新失败，使用异步非阻塞通知
            ShowAsyncNotification(message)

        Catch ex As Exception
            Debug.WriteLine("显示通知失败: " & ex.Message)
        End Try
    End Sub

    Private Sub ShowAsyncNotification(message As String)
        ' 使用Task异步显示通知，避免阻塞
        Task.Run(Sub()
                     Try
                         ' 在后台线程中创建通知
                         Dim notification As Form = Nothing

                         ' 必须在UI线程上创建Form
                         Dim createAction As Action = Sub()
                                                          notification = New Form() With {
                                                              .Text = "AI提示",
                                                              .FormBorderStyle = FormBorderStyle.FixedToolWindow,
                                                              .StartPosition = FormStartPosition.Manual,
                                                              .ShowInTaskbar = False,
                                                              .TopMost = True,
                                                              .Size = New Size(350, 100),
                                                              .BackColor = Color.FromArgb(245, 245, 245),
                                                              .Opacity = 0.9
                                                          }

                                                          ' 创建消息标签 - 修复文本显示问题
                                                          Dim lblMessage As New Label() With {
                                                              .Text = message,
                                                              .ForeColor = Color.FromArgb(51, 51, 51),
                                                              .Font = New Font("Microsoft YaHei UI", 9.0F, FontStyle.Regular),
                                                              .AutoSize = False,
                                                              .Size = New Size(330, 60),
                                                              .Location = New Point(10, 20),
                                                              .TextAlign = ContentAlignment.MiddleLeft
                                                          }

                                                          notification.Controls.Add(lblMessage)

                                                          ' 获取Excel窗口位置（非阻塞方式）
                                                          Try
                                                              Dim excelRect As New RECT()
                                                              If GetWindowRect(Process.GetCurrentProcess().MainWindowHandle, excelRect) Then
                                                                  ' 定位在Excel窗口右下角
                                                                  notification.Left = excelRect.right - notification.Width - 20
                                                                  notification.Top = excelRect.bottom - notification.Height - 60
                                                              Else
                                                                  ' 如果获取失败，显示在屏幕右下角
                                                                  notification.Left = Screen.PrimaryScreen.WorkingArea.Width - notification.Width - 20
                                                                  notification.Top = Screen.PrimaryScreen.WorkingArea.Height - notification.Height - 60
                                                              End If
                                                          Catch
                                                              ' 异常时使用默认位置
                                                              notification.Left = Screen.PrimaryScreen.WorkingArea.Width - notification.Width - 20
                                                              notification.Top = Screen.PrimaryScreen.WorkingArea.Height - notification.Height - 60
                                                          End Try
                                                      End Sub

                         ' 在UI线程上创建窗口
                         If Application.OpenForms.Count > 0 Then
                             Application.OpenForms(0).Invoke(createAction)
                         Else
                             ' 如果没有可用的Form，直接在当前线程创建
                             createAction()
                         End If

                         ' 如果通知创建成功，显示并设置自动关闭
                         If notification IsNot Nothing Then
                             ' 异步显示通知
                             Task.Run(Sub()
                                          Try
                                              ' 非阻塞显示
                                              If Application.OpenForms.Count > 0 Then
                                                  Application.OpenForms(0).BeginInvoke(New Action(Sub()
                                                                                                      notification.Show()
                                                                                                  End Sub))
                                              Else
                                                  notification.Show()
                                              End If

                                              ' 3秒后自动关闭
                                              Threading.Thread.Sleep(3000)

                                              ' 异步关闭
                                              If Application.OpenForms.Count > 0 Then
                                                  Application.OpenForms(0).BeginInvoke(New Action(Sub()
                                                                                                      Try
                                                                                                          If notification IsNot Nothing AndAlso Not notification.IsDisposed Then
                                                                                                              notification.Close()
                                                                                                              notification.Dispose()
                                                                                                          End If
                                                                                                      Catch
                                                                                                          ' 忽略关闭时的异常
                                                                                                      End Try
                                                                                                  End Sub))
                                              End If
                                          Catch ex As Exception
                                              Debug.WriteLine("异步通知显示失败: " & ex.Message)
                                          End Try
                                      End Sub)
                         End If

                     Catch ex As Exception
                         Debug.WriteLine("创建异步通知失败: " & ex.Message)
                     End Try
                 End Sub)
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
            If _application Is Nothing Then
                Throw New InvalidOperationException("应用程序未初始化")
            End If

            Debug.WriteLine("正在设置状态栏文本: " & message)

            ' 直接访问Excel的StatusBar属性
            Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
            If propertyInfo IsNot Nothing Then
                propertyInfo.SetValue(_application, "AI提示: " & message, Nothing)
                Debug.WriteLine("状态栏消息设置成功")
                _messagePending = True
                StatusTimer.Start() ' 启动定时器，稍后清除状态栏
            Else
                Throw New NotSupportedException("无法找到StatusBar属性")
            End If
        Catch ex As Exception
            Debug.WriteLine("更新状态栏失败: " & ex.Message)
            Throw ' 重新抛出异常以便上层处理
        End Try
    End Sub

    Private Sub StatusTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Try
            If Not _messagePending OrElse _application Is Nothing Then
                StatusTimer.Stop()
                _messagePending = False
                Return
            End If

            Debug.WriteLine("计时器触发，准备清除状态栏")

            ' 异步清除状态栏，避免阻塞
            Task.Run(Sub()
                         Try
                             ' 使用反射清除StatusBar
                             Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
                             If propertyInfo IsNot Nothing Then
                                 propertyInfo.SetValue(_application, False, Nothing)
                                 Debug.WriteLine("状态栏已清除")
                             End If
                         Catch ex As Exception
                             Debug.WriteLine("清除状态栏失败: " & ex.Message)
                         End Try
                     End Sub)

            StatusTimer.Stop()
            _messagePending = False
        Catch ex As Exception
            Debug.WriteLine("定时器事件处理失败: " & ex.Message)
            StatusTimer.Stop()
            _messagePending = False
        End Try
    End Sub
End Module