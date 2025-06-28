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
    Private notificationForm As NotificationForm = Nothing

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
                Catch ex As Exception
                    Debug.WriteLine($"����״̬��ʧ��: {ex.Message}")
                End Try
            End If

            ' ����2: ������ʾһ��Ư����֪ͨ����
            ShowNotificationForm(message)

            ' ���÷��������������
            Debug.WriteLine("AI��ʾ: " & message)

        Catch ex As Exception
            Debug.WriteLine("��ʾ֪ͨʧ��: " & ex.Message)
        End Try
    End Sub

    ' ��ʾ�Զ���֪ͨ����
    Private Sub ShowNotificationForm(message As String)
        Try
            ' ��������ʾ֪ͨ����
            Dim thread As New System.Threading.Thread(
                Sub()
                    Try
                        ' ����֪ͨ����
                        Dim form As New NotificationForm(message)
                        form.ShowDialog() ' ʹ��ShowDialog�Ա����߳�����ֱ������ر�
                    Catch ex As Exception
                        Debug.WriteLine($"��ʾ֪ͨ����ʧ��: {ex.Message}")
                    End Try
                End Sub)

            thread.SetApartmentState(System.Threading.ApartmentState.STA)
            thread.IsBackground = True ' ����Ϊ��̨�̣߳�����������ֹӦ�ó����˳�
            thread.Start()
        Catch ex As Exception
            Debug.WriteLine($"����֪ͨ�߳�ʧ��: {ex.Message}")
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

' �Զ���֪ͨ����
Public Class NotificationForm : Inherits Form
    Private WithEvents closeTimer As New System.Windows.Forms.Timer()
    Private fadeTimer As New System.Windows.Forms.Timer()
    Private _opacity As Double = 1.0

    Public Sub New(message As String)
        ' ���ô�������
        Me.FormBorderStyle = FormBorderStyle.None
        Me.StartPosition = FormStartPosition.Manual
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.Size = New Size(300, 80)
        Me.BackColor = Color.FromArgb(50, 50, 50)
        Me.Opacity = 0.9

        ' Բ��Ч��
        Dim path As New Drawing2D.GraphicsPath()
        path.AddArc(0, 0, 20, 20, 180, 90)
        path.AddArc(Me.Width - 20, 0, 20, 20, 270, 90)
        path.AddArc(Me.Width - 20, Me.Height - 20, 20, 20, 0, 90)
        path.AddArc(0, Me.Height - 20, 20, 20, 90, 90)
        Me.Region = New Region(path)

        ' �����Ϣ��ǩ
        Dim lblMessage As New Label()
        lblMessage.Text = message
        lblMessage.ForeColor = Color.White
        lblMessage.Font = New Font("Microsoft YaHei UI", 9.0F, FontStyle.Regular)
        lblMessage.AutoSize = False
        lblMessage.Size = New Size(280, 60)
        lblMessage.Location = New Point(10, 10)
        lblMessage.TextAlign = ContentAlignment.MiddleCenter
        Me.Controls.Add(lblMessage)

        ' ���ô���λ�� - ���½�
        Dim screenWidth As Integer = Screen.PrimaryScreen.WorkingArea.Width
        Dim screenHeight As Integer = Screen.PrimaryScreen.WorkingArea.Height
        Me.Location = New Point(screenWidth - Me.Width - 20, screenHeight - Me.Height - 20)

        ' �����Զ��رն�ʱ��
        closeTimer.Interval = 3000 ' 3���ʼ����
        closeTimer.Start()

        ' ���ý���Ч����ʱ��
        fadeTimer.Interval = 50 ' ÿ50�������һ��͸����
        AddHandler fadeTimer.Tick, AddressOf FadeTimer_Tick
    End Sub

    Private Sub CloseTimer_Tick(sender As Object, e As EventArgs) Handles closeTimer.Tick
        closeTimer.Stop()
        fadeTimer.Start() ' ��ʼ����Ч��
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

        ' ���Ʊ߿�
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