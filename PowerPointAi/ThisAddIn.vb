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

        Try
            WebView2Loader.EnsureWebView2Loader()
        Catch ex As Exception
            MessageBox.Show($"WebView2 ��ʼ��ʧ��: {ex.Message}")
        End Try

        ' ��������͹������л��¼�
        'AddHandler Globals.ThisAddIn.Application.ActiveDocument, AddressOf Me.Application_WorkbookActivate
        Application_WorkbookActivate()
        ' ��ʼ�� Timer������WPS��������������Ŀ��
        widthTimer = New Timer()
        AddHandler widthTimer.Tick, AddressOf WidthTimer_Tick
        widthTimer.Interval = 100 ' �����ӳ�ʱ�䣬��λΪ����

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


    '    ' �л�������ʱ����

    Private Sub Application_WorkbookActivate()
        Try
            ' Ϊ�¹������������񴰸�
            chatControl = New ChatControl()
            chatTaskPane = Me.CustomTaskPanes.Add(chatControl, "PPT AI��������")
            chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            chatTaskPane.Width = 420
            AddHandler chatTaskPane.VisibleChanged, AddressOf ChatTaskPane_VisibleChanged
            chatTaskPane.Visible = False

        Catch ex As Exception
            MessageBox.Show($"��ʼ���½����������񴰸�ʧ��: {ex.Message}")
        End Try
    End Sub

    Private widthTimer As Timer
    ' ���WPS���޷���ʾ������ȵ�����
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
        ' ������Դ
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
