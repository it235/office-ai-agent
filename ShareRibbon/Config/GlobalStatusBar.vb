Imports System.Windows.Forms
Imports System.Drawing
Imports System.Timers
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Imports System.Diagnostics
Imports Timer = System.Windows.Forms.Timer

Public Module GlobalStatusBar
    Private WithEvents StatusTimer As New System.Timers.Timer()
    Private _application As Object = Nothing
    Private _messagePending As Boolean = False

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
                    StatusTimer.Start()

                    Debug.WriteLine("状态栏消息设置成功")
                    Return
                Catch ex As Exception
                    Debug.WriteLine($"设置状态栏失败: {ex.Message}")
                End Try
            End If

            ' 方法2: 如果状态栏设置失败，使用一个简单的MessageBox
            Dim thread As New System.Threading.Thread(
                Sub()
                    Try
                        MessageBox.Show(message, "AI提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        Debug.WriteLine($"显示消息框失败: {ex.Message}")
                    End Try
                End Sub)

            thread.SetApartmentState(System.Threading.ApartmentState.STA)
            thread.Start()

            ' 备用方案：输出到调试
            Debug.WriteLine("AI提示: " & message)

        Catch ex As Exception
            Debug.WriteLine("显示通知失败: " & ex.Message)
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