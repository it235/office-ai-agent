Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Vbe.Interop
Imports ShareRibbon
Public Class WebDataCapturePane
    Inherits BaseDataCapturePane

    Private isViewInitialized As Boolean = False
    Public Sub New()
        MyBase.New()
        ' ���� ChatControl ʵ��
        ' ����AI���������¼�
        AddHandler AiChatRequested, AddressOf HandleAiChatRequest
        ' ֱ�ӵ����첽��ʼ������
        InitializeWebViewAsync()
    End Sub

    ' �������첽��ʼ������
    ' �첽��ʼ������
    Private Async Sub InitializeWebViewAsync()
        Try
            Debug.WriteLine("Starting WebView initialization from WebDataCapturePane")
            ' ���û���ĳ�ʼ������
            Await InitializeWebView2()
        Catch ex As Exception
            MessageBox.Show($"��ʼ����ҳ��ͼʧ��: {ex.Message}", "����",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub HandleAiChatRequest(sender As Object, content As String)
        ' ��ʾ���촰��
        Globals.ThisAddIn.ShowChatTaskPane()
        ' ���ѡ�е����ݵ�������
        Globals.ThisAddIn.chatControl.AddSelectedContentItem(
                "������ҳ",  ' ʹ���ĵ�������Ϊ��ʶ
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
    End Sub

    ' �����񴴽�
    Protected Overrides Function CreateTable(tableData As TableData) As String
        Try
            ' ��ȡ��ǰ�ĵ���ѡ����Χ
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selection = doc.Application.Selection

            ' �������
            Dim table = doc.Tables.Add(
                Range:=selection.Range,
                NumRows:=tableData.Rows,
                NumColumns:=tableData.Columns)

            ' �������
            For i = 0 To tableData.Data.Count - 1
                For j = 0 To tableData.Data(i).Count - 1
                    table.Cell(i + 1, j + 1).Range.Text = tableData.Data(i)(j)
                Next
            Next

            ' ����б�ͷ�����ñ�ͷ��ʽ
            If tableData.Headers.Count > 0 Then
                table.Rows(1).HeadingFormat = True
                table.Rows(1).Range.Bold = True
            End If

            ' ���ñ����ʽ
            table.Style = "������"
            table.AllowAutoFit = True
            table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

            Return "[����Ѳ���]" & vbCrLf
        Catch ex As Exception
            MessageBox.Show($"�������ʱ����: {ex.Message}", "����")
            Return String.Empty
        End Try
    End Function

    Protected Overrides Sub HandleExtractedContent(content As String)
        Try
            ' ��ȡ��ĵ�
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc IsNot Nothing Then
                ' �ڵ�ǰ���λ�ò�������
                Dim selection = doc.Application.Selection
                If selection IsNot Nothing Then
                    ' ��������
                    selection.TypeText(content)
                    'selection.TypeText(vbCrLf & vbCrLf)

                    ' ����ָ���
                    'selection.TypeText(vbCrLf & "----------------------------------------" & vbCrLf)
                    'selection.TypeText("��Դ: " & ChatBrowser.CoreWebView2.DocumentTitle & vbCrLf)
                    'selection.TypeText("URL: " & ChatBrowser.CoreWebView2.Source & " " & "ʱ��: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & vbCrLf)
                    'selection.TypeText("----------------------------------------" & vbCrLf & vbCrLf)

                    'MessageBox.Show("�����ѳɹ���ȡ�����뵽�ĵ���", "�ɹ�")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"������ȡ����ʱ����: {ex.Message}", "����",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' �����ͼ���ٴ���
    Protected Overrides Sub OnHandleDestroyed(e As EventArgs)
        isViewInitialized = False
        MyBase.OnHandleDestroyed(e)
    End Sub

End Class