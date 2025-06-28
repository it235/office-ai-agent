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
        StatusTimer.Interval = 5000 ' ������ʾ��ʾʱ��Ϊ5��
        AddHandler StatusTimer.Elapsed, AddressOf StatusTimer_Elapsed
    End Sub

    Public Sub InitializeApplication(application As Object)
        _application = application
        Debug.WriteLine("GlobalStatusStrip�ѳ�ʼ����Ӧ�ó�������: " & application.GetType().FullName)
    End Sub

    Public Sub ShowWarningStatus(message As String)
        ShowWarning(message) ' �򻯴��룬ʹ��ͬһ������
    End Sub

    Public Sub ShowWarning(message As String)
        Try
            Debug.WriteLine("��ʾ����֪ͨ: " & message)

            ' ����1: ֱ��ʹ��Ӧ�ó����״̬��
            If _application IsNot Nothing Then
                Try
                    ' ʹ�ö�̬���ͣ����ⷴ��
                    Dim app = TryCast(_application, Object)
                    app.StatusBar = "AI: " & message

                    ' ������ʱ����5������״̬��
                    _messagePending = True
                    StatusTimer.Stop() ' ȷ����ֹ֮ͣǰ�ļ�ʱ��
                    StatusTimer.Start()

                    Debug.WriteLine("״̬����Ϣ���óɹ�")
                    Return
                Catch ex As Exception
                    Debug.WriteLine($"����״̬��ʧ��: {ex.Message}")
                End Try
            End If

            ' ����2: ���״̬������ʧ�ܣ�ʹ��һ���򵥵�MessageBox
            Dim thread As New System.Threading.Thread(
                Sub()
                    Try
                        MessageBox.Show(message, "AI��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        Debug.WriteLine($"��ʾ��Ϣ��ʧ��: {ex.Message}")
                    End Try
                End Sub)

            thread.SetApartmentState(System.Threading.ApartmentState.STA)
            thread.Start()

            ' ���÷��������������
            Debug.WriteLine("AI��ʾ: " & message)

        Catch ex As Exception
            Debug.WriteLine("��ʾ֪ͨʧ��: " & ex.Message)
        End Try
    End Sub

    Private Sub StatusTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Try
            StatusTimer.Stop()

            If Not _messagePending OrElse _application Is Nothing Then
                _messagePending = False
                Return
            End If

            Debug.WriteLine("��ʱ��������׼�����״̬��")

            Try
                ' ���״̬���ļ򵥷���
                Dim app = TryCast(_application, Object)
                app.StatusBar = False
                Debug.WriteLine("״̬�������")
            Catch ex As Exception
                Debug.WriteLine($"���״̬��ʧ��: {ex.Message}")
            End Try

            _messagePending = False
        Catch ex As Exception
            Debug.WriteLine("��ʱ���¼�����ʧ��: " & ex.Message)
            _messagePending = False
        End Try
    End Sub
End Module