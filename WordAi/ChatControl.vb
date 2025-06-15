Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Public Class ChatControl
    Inherits BaseChatControl


    Private sheetContentItems As New Dictionary(Of String, Tuple(Of System.Windows.Forms.Label, System.Windows.Forms.Button))


    Public Sub New()
        ' �˵��������ʦ������ġ�
        InitializeComponent()

        ' ȷ��WebView2�ؼ�������������
        ChatBrowser.BringToFront()

        '����ײ��澯��
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)

        ' ����Word��SelectionChange �¼�
        ' ���Ҳ�ȫwordѡ��������¼�
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf GetSelectionContent
    End Sub

    '��ȡѡ�е�����
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' ת��Ϊ Word.Selection ����
            Dim selection = TryCast(Globals.ThisAddIn.Application.Selection, Microsoft.Office.Interop.Word.Selection)
            If selection Is Nothing Then
                Return
            End If

            ' ��ȡѡ�����ݵ���ϸ��Ϣ
            Dim content As String = String.Empty

            ' ����Ƿ�ѡ���˱��
            If selection.Tables.Count > 0 Then
                ' ���ѡ�е��Ǳ��
                Dim table = selection.Tables(1)
                Dim sb As New StringBuilder()

                ' �����������
                For row As Integer = 1 To table.Rows.Count
                    For col As Integer = 1 To table.Columns.Count
                        sb.Append(table.Cell(row, col).Range.Text.TrimEnd(ChrW(13), ChrW(7)))
                        If col < table.Columns.Count Then sb.Append(vbTab)
                    Next
                    sb.AppendLine()
                Next
                content = sb.ToString()

            ElseIf selection.InlineShapes.Count > 0 OrElse selection.ShapeRange.Count > 0 Then
                ' ���ѡ�е���ͼƬ����״
                content = "[ͼƬ����״]"
            Else
                ' ��ͨ�ı�ѡ��
                content = selection.Text
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' ��ӵ�ѡ�������б�
                AddSelectedContentItem(
                "Word�ĵ�",  ' ʹ���ĵ�������Ϊ��ʶ
                If(selection.Tables.Count > 0,
                   "[�������]",
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
            )
            End If

        Catch ex As Exception
            Debug.WriteLine($"��ȡWordѡ������ʱ����: {ex.Message}")
        End Try
    End Sub


    ' ��ȡѡ�����ݵ���ϸ��Ϣ
    Private Function GetSelectionDetails(selection As Microsoft.Office.Interop.Word.Selection) As String
        Dim details As New StringBuilder()

        ' ��ӻ�����Ϣ
        details.AppendLine($"��ʼλ��: {selection.Start}")
        details.AppendLine($"����λ��: {selection.End}")
        details.AppendLine($"�ַ���: {selection.Characters.Count}")

        ' ����Ǳ����ӱ����Ϣ
        If selection.Tables.Count > 0 Then
            Dim table = selection.Tables(1)
            details.AppendLine($"����С: {table.Rows.Count}�� x {table.Columns.Count}��")
        End If

        Return details.ToString()
    End Function

    ' ��ʼ��ʱע����� HTML �ṹ
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ��ʼ�� WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


    Protected Overrides Function GetVBProject() As VBProject
        Try
            Dim project = Globals.ThisAddIn.Application.VBE.ActiveVBProject
            Return project
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return Nothing
        End Try
    End Function

    Protected Overrides Function RunCode(code As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(code)
            Return True
        Catch ex As Runtime.InteropServices.COMException
            VBAxceptionHandle(ex)
            Return False
        Catch ex As Exception
            MessageBox.Show("ִ�д���ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Word", OfficeApplicationType.Word)
    End Function

    Protected Overrides Sub SendChatMessage(message As String)
        ' �������ʵ��word�������߼�
        Send(message)
    End Sub


End Class

