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

    ' ���ڼ�������Ƿ���UI�߳��ϵı���
    Private _mainThreadId As Integer = Threading.Thread.CurrentThread.ManagedThreadId

    Sub New()
        StatusTimer.Interval = 5000 ' ������ʾ��ʾʱ��Ϊ5��
        AddHandler StatusTimer.Elapsed, AddressOf StatusTimer_Elapsed
    End Sub

    Public Sub InitializeApplication(application As Object)
        _application = application
        Debug.WriteLine("GlobalStatusStrip�ѳ�ʼ����Ӧ�ó�������: " & application.GetType().FullName)
    End Sub
    Public Sub ShowWarning(message As String)
        Try
            Debug.WriteLine("��ʾ����֪ͨ: " & Message)

            ' ��������ʾ�򵥵ķ�ģ̬֪ͨ
            Dim showAction As Action = Sub()
                                           ' ����֪ͨ���� - ������С��СһЩ
                                           Dim notification As New Form() With {
                .Text = "��ʾ��Ϣ (3)",  ' ��ʼ��ʾ3�뵹��ʱ
                .FormBorderStyle = FormBorderStyle.FixedToolWindow,
                .StartPosition = FormStartPosition.Manual,
                .ShowInTaskbar = False,
                .TopMost = True,
                .Size = New Size(280, 80),  ' ������С�ĳߴ�
                .BackColor = Color.FromArgb(255, 240, 240),
                .Opacity = 0.85  ' ����͸����
            }

                                           ' ������Ϣ��ǩ
                                           Dim lblMessage As New Label() With {
                .Text = Message,
                .ForeColor = Color.DarkRed,
                .Font = New Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular),  ' ��΢��С����
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter
            }

                                           notification.Controls.Add(lblMessage)

                                           ' ��ȡExcel����λ��
                                           Try
                                               Dim excelRect As New RECT()
                                               GetWindowRect(Process.GetCurrentProcess().MainWindowHandle, excelRect)

                                               ' ��λ��Excel�������½�
                                               notification.Left = excelRect.right - notification.Width - 20
                                               notification.Top = excelRect.bottom - notification.Height - 40  ' ��Ϊ�ײ�λ��
                                           Catch ex As Exception
                                               ' �����ȡExcel����λ��ʧ�ܣ��������ʾ
                                               notification.StartPosition = FormStartPosition.CenterScreen
                                           End Try

                                           ' ��ʾ֪ͨ�������Զ��ر�
                                           notification.Show()

                                           ' ����ʱ����
                                           Dim remainingSeconds As Integer = 3

                                           ' ���õ���ʱ��ʱ��
                                           Dim countdownTimer As New Timer() With {
                .Interval = 1000  ' ÿ�����һ��
            }

                                           ' �����Զ��رռ�ʱ��
                                           Dim closeTimer As New Timer() With {
                .Interval = 3000  ' 3���ر�
            }

                                           ' ����ʱ����
                                           AddHandler countdownTimer.Tick, Sub(s, e)
                                                                               remainingSeconds -= 1
                                                                               If remainingSeconds > 0 Then
                                                                                   notification.Text = $"��ʾ��Ϣ ({remainingSeconds})"
                                                                               Else
                                                                                   countdownTimer.Stop()
                                                                               End If
                                                                           End Sub

                                           ' �رմ���
                                           AddHandler closeTimer.Tick, Sub(s, e)
                                                                           closeTimer.Stop()
                                                                           countdownTimer.Stop()
                                                                           notification.Close()
                                                                           notification.Dispose()
                                                                       End Sub

                                           ' ������ʱ��
                                           countdownTimer.Start()
                                           closeTimer.Start()
                                       End Sub

            ' ��UI�߳���ִ��
            If Threading.Thread.CurrentThread.ManagedThreadId = _mainThreadId Then
                showAction.Invoke()
            Else
                Dim form As New Form()
                form.Invoke(showAction)
            End If

        Catch ex As Exception
            Debug.WriteLine("��ʾ֪ͨʧ��: " & ex.Message)
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
            ' ʹ�ö�̬���ʹ���ͬ��OfficeӦ�ó���
            Debug.WriteLine("��������״̬���ı�: " & message)

            ' ֱ�ӷ���Excel/Word/PowerPoint��StatusBar����
            Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
            If propertyInfo IsNot Nothing Then
                propertyInfo.SetValue(_application, "��ʾ��" & message, Nothing)
                Debug.WriteLine("״̬����Ϣ���óɹ�")
            Else
                Debug.WriteLine("�޷��ҵ�StatusBar����")
                ' ������һ�ַ���
                Try
                    ' ĳЩOffice�汾����ʹ�ò�ͬ�ķ���
                    _application.GetType().InvokeMember("StatusBar",
                        Reflection.BindingFlags.SetProperty,
                        Nothing,
                        _application,
                        New Object() {"��ʾ��" & message})
                    Debug.WriteLine("ͨ����������״̬����Ϣ�ɹ�")
                Catch ex As Exception
                    Debug.WriteLine("ͨ����������״̬��ʧ��: " & ex.Message)
                    Throw ' �����׳��쳣�Ա��ϲ㴦��
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("����״̬��ʧ��: " & ex.Message)
            Throw ' �����׳��쳣�Ա��ϲ㴦��
        End Try
    End Sub

    Private Sub StatusTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Try
            If Not _messagePending Then Return

            Debug.WriteLine("��ʱ��������׼�����״̬��")

            If _application Is Nothing Then
                StatusTimer.Stop()
                _messagePending = False
                Return
            End If

            ' ������UI�߳���ִ��
            Dim form As New Form()
            form.Invoke(New Action(Sub()
                                       Try
                                           ' �������״̬��
                                           Debug.WriteLine("�������״̬��")

                                           ' ʹ�÷�������StatusBar����
                                           Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
                                           If propertyInfo IsNot Nothing Then
                                               propertyInfo.SetValue(_application, False, Nothing)
                                               Debug.WriteLine("״̬�������")
                                           Else
                                               ' ������һ�ַ���
                                               Try
                                                   _application.GetType().InvokeMember("StatusBar",
                                                       Reflection.BindingFlags.SetProperty,
                                                       Nothing,
                                                       _application,
                                                       New Object() {False})
                                                   Debug.WriteLine("ͨ���������״̬���ɹ�")
                                               Catch ex As Exception
                                                   Debug.WriteLine("ͨ���������״̬��ʧ��: " & ex.Message)
                                               End Try
                                           End If
                                       Catch ex As Exception
                                           Debug.WriteLine("���״̬��ʧ��: " & ex.Message)
                                       End Try
                                   End Sub))

            StatusTimer.Stop()
            _messagePending = False
        Catch ex As Exception
            Debug.WriteLine("��ʱ���¼�����ʧ��: " & ex.Message)
            StatusTimer.Stop()
            _messagePending = False
        End Try
    End Sub
End Module