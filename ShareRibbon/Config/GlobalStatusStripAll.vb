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

    ' ���ڼ�������Ƿ���UI�߳��ϵı���
    Private _mainThreadId As Integer = Threading.Thread.CurrentThread.ManagedThreadId

    Sub New()
        StatusTimer.Interval = 8000 ' ������ʾ��ʾʱ��Ϊ8��
        AddHandler StatusTimer.Elapsed, AddressOf StatusTimer_Elapsed
    End Sub

    Public Sub InitializeApplication(application As Object)
        _application = application
        Debug.WriteLine("GlobalStatusStrip�ѳ�ʼ����Ӧ�ó�������: " & application.GetType().FullName)
    End Sub

    Public Sub ShowWarningStatus(message As String)
        Try
            Debug.WriteLine("��ʾ����֪ͨ: " & message)
            Debug.WriteLine($"_application״̬: {If(_application Is Nothing, "Nothing", "�ѳ�ʼ��")}")

            ' ���Ը���״̬�� - ��ʹ��Excel-DNA���еķ���
            If _application IsNot Nothing Then
                Try
                    ' ֱ�ӳ������ã������߳��Ͽ��ܳɹ�
                    _application.StatusBar = "AI: " & message
                    Debug.WriteLine("ֱ������״̬���ɹ�")
                    Return
                Catch directEx As Exception
                    Debug.WriteLine($"ֱ������ʧ��: {directEx.Message}")

                    ' ���ã�ʹ�÷��䷽ʽ
                    Try
                        Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
                        If propertyInfo IsNot Nothing Then
                            propertyInfo.SetValue(_application, "AI: " & message, Nothing)
                            Debug.WriteLine("��������״̬���ɹ�")
                            Return
                        End If
                    Catch reflectionEx As Exception
                        Debug.WriteLine($"��������ʧ��: {reflectionEx.Message}")
                    End Try
                End Try
            Else
                Debug.WriteLine("_applicationΪNothing���޷�����״̬��")
            End If

            ' ���÷��������������
            Console.WriteLine("AI��ʾ: " & message)
            Debug.WriteLine("AI��ʾ: " & message)

        Catch ex As Exception
            Debug.WriteLine("��ʾ֪ͨʧ��: " & ex.Message)
        End Try
    End Sub

    Public Sub ShowWarning(message As String)
        Try
            Debug.WriteLine("��ʾ����֪ͨ: " & message)

            ' ���ȳ��Ը���Excel״̬������������ʽ��
            Try
                UpdateStatusBarDirectly(message)
                Return ' ���״̬�����³ɹ����Ͳ���ʾ������
            Catch ex As Exception
                Debug.WriteLine("״̬������ʧ�ܣ����÷�����֪ͨ: " & ex.Message)
            End Try

            ' ���״̬������ʧ�ܣ�ʹ���첽������֪ͨ
            ShowAsyncNotification(message)

        Catch ex As Exception
            Debug.WriteLine("��ʾ֪ͨʧ��: " & ex.Message)
        End Try
    End Sub

    Private Sub ShowAsyncNotification(message As String)
        ' ʹ��Task�첽��ʾ֪ͨ����������
        Task.Run(Sub()
                     Try
                         ' �ں�̨�߳��д���֪ͨ
                         Dim notification As Form = Nothing

                         ' ������UI�߳��ϴ���Form
                         Dim createAction As Action = Sub()
                                                          notification = New Form() With {
                                                              .Text = "AI��ʾ",
                                                              .FormBorderStyle = FormBorderStyle.FixedToolWindow,
                                                              .StartPosition = FormStartPosition.Manual,
                                                              .ShowInTaskbar = False,
                                                              .TopMost = True,
                                                              .Size = New Size(350, 100),
                                                              .BackColor = Color.FromArgb(245, 245, 245),
                                                              .Opacity = 0.9
                                                          }

                                                          ' ������Ϣ��ǩ - �޸��ı���ʾ����
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

                                                          ' ��ȡExcel����λ�ã���������ʽ��
                                                          Try
                                                              Dim excelRect As New RECT()
                                                              If GetWindowRect(Process.GetCurrentProcess().MainWindowHandle, excelRect) Then
                                                                  ' ��λ��Excel�������½�
                                                                  notification.Left = excelRect.right - notification.Width - 20
                                                                  notification.Top = excelRect.bottom - notification.Height - 60
                                                              Else
                                                                  ' �����ȡʧ�ܣ���ʾ����Ļ���½�
                                                                  notification.Left = Screen.PrimaryScreen.WorkingArea.Width - notification.Width - 20
                                                                  notification.Top = Screen.PrimaryScreen.WorkingArea.Height - notification.Height - 60
                                                              End If
                                                          Catch
                                                              ' �쳣ʱʹ��Ĭ��λ��
                                                              notification.Left = Screen.PrimaryScreen.WorkingArea.Width - notification.Width - 20
                                                              notification.Top = Screen.PrimaryScreen.WorkingArea.Height - notification.Height - 60
                                                          End Try
                                                      End Sub

                         ' ��UI�߳��ϴ�������
                         If Application.OpenForms.Count > 0 Then
                             Application.OpenForms(0).Invoke(createAction)
                         Else
                             ' ���û�п��õ�Form��ֱ���ڵ�ǰ�̴߳���
                             createAction()
                         End If

                         ' ���֪ͨ�����ɹ�����ʾ�������Զ��ر�
                         If notification IsNot Nothing Then
                             ' �첽��ʾ֪ͨ
                             Task.Run(Sub()
                                          Try
                                              ' ��������ʾ
                                              If Application.OpenForms.Count > 0 Then
                                                  Application.OpenForms(0).BeginInvoke(New Action(Sub()
                                                                                                      notification.Show()
                                                                                                  End Sub))
                                              Else
                                                  notification.Show()
                                              End If

                                              ' 3����Զ��ر�
                                              Threading.Thread.Sleep(3000)

                                              ' �첽�ر�
                                              If Application.OpenForms.Count > 0 Then
                                                  Application.OpenForms(0).BeginInvoke(New Action(Sub()
                                                                                                      Try
                                                                                                          If notification IsNot Nothing AndAlso Not notification.IsDisposed Then
                                                                                                              notification.Close()
                                                                                                              notification.Dispose()
                                                                                                          End If
                                                                                                      Catch
                                                                                                          ' ���Թر�ʱ���쳣
                                                                                                      End Try
                                                                                                  End Sub))
                                              End If
                                          Catch ex As Exception
                                              Debug.WriteLine("�첽֪ͨ��ʾʧ��: " & ex.Message)
                                          End Try
                                      End Sub)
                         End If

                     Catch ex As Exception
                         Debug.WriteLine("�����첽֪ͨʧ��: " & ex.Message)
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
                Throw New InvalidOperationException("Ӧ�ó���δ��ʼ��")
            End If

            Debug.WriteLine("��������״̬���ı�: " & message)

            ' ֱ�ӷ���Excel��StatusBar����
            Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
            If propertyInfo IsNot Nothing Then
                propertyInfo.SetValue(_application, "AI��ʾ: " & message, Nothing)
                Debug.WriteLine("״̬����Ϣ���óɹ�")
                _messagePending = True
                StatusTimer.Start() ' ������ʱ�����Ժ����״̬��
            Else
                Throw New NotSupportedException("�޷��ҵ�StatusBar����")
            End If
        Catch ex As Exception
            Debug.WriteLine("����״̬��ʧ��: " & ex.Message)
            Throw ' �����׳��쳣�Ա��ϲ㴦��
        End Try
    End Sub

    Private Sub StatusTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Try
            If Not _messagePending OrElse _application Is Nothing Then
                StatusTimer.Stop()
                _messagePending = False
                Return
            End If

            Debug.WriteLine("��ʱ��������׼�����״̬��")

            ' �첽���״̬������������
            Task.Run(Sub()
                         Try
                             ' ʹ�÷������StatusBar
                             Dim propertyInfo = _application.GetType().GetProperty("StatusBar")
                             If propertyInfo IsNot Nothing Then
                                 propertyInfo.SetValue(_application, False, Nothing)
                                 Debug.WriteLine("״̬�������")
                             End If
                         Catch ex As Exception
                             Debug.WriteLine("���״̬��ʧ��: " & ex.Message)
                         End Try
                     End Sub)

            StatusTimer.Stop()
            _messagePending = False
        Catch ex As Exception
            Debug.WriteLine("��ʱ���¼�����ʧ��: " & ex.Message)
            StatusTimer.Stop()
            _messagePending = False
        End Try
    End Sub
End Module