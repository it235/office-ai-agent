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


    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' ����һ�� Word Ӧ�ó���ʵ��
            Dim wordApp As New Microsoft.Office.Interop.Word.Application
            wordApp.Visible = False

            Dim document As Microsoft.Office.Interop.Word.Document = Nothing
            Try
                document = wordApp.Documents.Open(filePath, ReadOnly:=True)
                Dim contentBuilder As New StringBuilder()

                contentBuilder.AppendLine($"�ļ�: {Path.GetFileName(filePath)} ������������:")

                ' ��ȡ�ĵ��ı�
                Dim text As String = document.Content.Text

                ' �����ı�����
                Dim maxTextLength As Integer = 2000
                If text.Length > maxTextLength Then
                    contentBuilder.AppendLine(text.Substring(0, maxTextLength) & "...")
                    contentBuilder.AppendLine($"[�ĵ�̫����ֻ��ʾǰ {maxTextLength} ���ַ����ܳ���: {text.Length} ���ַ�]")
                Else
                    contentBuilder.AppendLine(text)
                End If

                Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Word",
                .ParsedContent = contentBuilder.ToString(),
                .RawData = Nothing
            }

            Finally
                ' ������Դ
                If document IsNot Nothing Then
                    document.Close(SaveChanges:=False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(document)
                End If

                wordApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            Debug.WriteLine($"���� Word �ļ�ʱ����: {ex.Message}")
            Return New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "Word",
            .ParsedContent = $"[���� Word �ļ�ʱ����: {ex.Message}]"
        }
        End Try
    End Function
    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            ' ��ȡ��ǰ���������·��
            If Globals.ThisAddIn.Application.ActiveWorkbook IsNot Nothing Then
                Return Globals.ThisAddIn.Application.ActiveWorkbook.Path
            End If
        Catch ex As Exception
            Debug.WriteLine($"��ȡ��ǰ����Ŀ¼ʱ����: {ex.Message}")
        End Try

        ' ����޷���ȡ������·�����򷵻�Ӧ�ó���Ŀ¼
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' ����Ƿ�������ѡ����
            If Not selectedCellChecked Then
                Return message
            End If

            ' ��ȡ��ǰ Word �ĵ��е�ѡ��
            Dim selection = Globals.ThisAddIn.Application.Selection
            If selection Is Nothing Then
                Return message
            End If

            ' �������ݹ���������ʽ��ѡ������
            Dim contentBuilder As New StringBuilder()
            contentBuilder.AppendLine(vbCrLf & "--- �û�ѡ�е� Word ���� ---")

            ' ����ĵ���Ϣ
            Dim activeDocument = Globals.ThisAddIn.Application.ActiveDocument
            If activeDocument IsNot Nothing Then
                contentBuilder.AppendLine($"�ĵ�: {Path.GetFileName(activeDocument.FullName)}")
            End If

            ' ѡ��Χ��Ϣ
            contentBuilder.AppendLine($"ѡ��Χ: �� {selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdFirstCharacterLineNumber)} ����")
            contentBuilder.AppendLine($"ѡ���ַ���: {selection.Characters.Count}")

            ' ����ѡ������
            If selection.Tables.Count > 0 Then
                ' ������
                contentBuilder.AppendLine("ѡ����������: ���")
                AppendTableContent(contentBuilder, selection)
            ElseIf selection.InlineShapes.Count > 0 OrElse selection.ShapeRange.Count > 0 Then
                ' ����ͼƬ����״
                contentBuilder.AppendLine("ѡ����������: ͼƬ����״")
                contentBuilder.AppendLine("[ͼƬ����״�����޷�ֱ��ת��Ϊ�ı�]")
            Else
                ' ������ͨ�ı�
                contentBuilder.AppendLine("ѡ����������: �ı�")
                Dim text As String = selection.Text.Trim()

                ' �����ı�����
                Dim maxLength As Integer = 2000
                If text.Length > maxLength Then
                    contentBuilder.AppendLine(text.Substring(0, maxLength) & "...")
                    contentBuilder.AppendLine($"[ѡ���ı�̫����ֻ��ʾǰ {maxLength} ���ַ����ܳ���: {text.Length} ���ַ�]")
                Else
                    contentBuilder.AppendLine(text)
                End If
            End If

            contentBuilder.AppendLine("--- ѡ�����ݽ��� ---" & vbCrLf)

            ' ����ԭʼ��Ϣ����ѡ������
            Return message & contentBuilder.ToString()

        Catch ex As Exception
            Debug.WriteLine($"����Wordѡ������ʱ����: {ex.Message}")
            Return message ' ����ʱ����ԭʼ��Ϣ
        End Try
    End Function

    ' ��������������������
    Private Sub AppendTableContent(builder As StringBuilder, selection As Microsoft.Office.Interop.Word.Selection)
        Try
            ' ��ȡѡ�еı��
            Dim table As Microsoft.Office.Interop.Word.Table = Nothing

            ' �����������������1. ѡ����������� 2. ѡ���˱���еĵ�Ԫ��
            If selection.Tables.Count > 0 Then
                table = selection.Tables(1)
            ElseIf selection.Cells.Count > 0 Then
                ' ���ֻѡ���˵�Ԫ�񣬻�ȡ������Щ��Ԫ��ı��
                table = selection.Cells(1).Range.Tables(1)
            End If

            If table Is Nothing Then
                builder.AppendLine("[�޷���ȡ�������]")
                Return
            End If

            ' ��ӱ����Ϣ
            builder.AppendLine($"����С: {table.Rows.Count} �� �� {table.Columns.Count} ��")
            builder.AppendLine()

            ' ������ʾ��������
            Dim maxRows As Integer = Math.Min(table.Rows.Count, 20)
            Dim maxCols As Integer = Math.Min(table.Columns.Count, 10)

            ' ������ͷ��������һ�У�
            If table.Rows.Count > 0 Then
                ' ������ͷ�ָ���
                Dim headerBuilder As New StringBuilder()
                Dim separatorBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(1, col).Range.Text
                        ' �Ƴ������ַ�
                        cellText = cellText.TrimEnd(ChrW(13), ChrW(7), ChrW(9), ChrW(10), ChrW(32))

                        ' ���Ƶ�Ԫ���ı�����
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' ����ͷ
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append(cellText)
                        separatorBuilder.Append(New String("-"c, Math.Max(cellText.Length, 3)))
                    Catch ex As Exception
                        ' ���Ե�Ԫ�������
                        If col > 1 Then
                            headerBuilder.Append(" | ")
                            separatorBuilder.Append("-+-")
                        End If
                        headerBuilder.Append("N/A")
                        separatorBuilder.Append("---")
                    End Try
                Next

                ' ��ӱ�ͷ�ͷָ���
                builder.AppendLine(headerBuilder.ToString())
                builder.AppendLine(separatorBuilder.ToString())
            End If

            ' ������������
            For row As Integer = 2 To maxRows ' �ӵ�2�п�ʼ��������ͷ��
                Dim rowBuilder As New StringBuilder()

                For col As Integer = 1 To maxCols
                    Try
                        Dim cellText As String = table.Cell(row, col).Range.Text
                        ' �Ƴ������ַ�
                        cellText = cellText.TrimEnd(ChrW(13), ChrW(7), ChrW(9), ChrW(10), ChrW(32))

                        ' ���Ƶ�Ԫ���ı�����
                        If cellText.Length > 20 Then
                            cellText = cellText.Substring(0, 17) & "..."
                        End If

                        ' ���������
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append(cellText)
                    Catch ex As Exception
                        ' ���Ե�Ԫ�������
                        If col > 1 Then
                            rowBuilder.Append(" | ")
                        End If
                        rowBuilder.Append("N/A")
                    End Try
                Next

                ' ���������
                builder.AppendLine(rowBuilder.ToString())
            Next

            ' ����и�����δ��ʾ�������ʾ
            If table.Rows.Count > maxRows Then
                builder.AppendLine($"... [����� {table.Rows.Count} �У�����ʾǰ {maxRows} ��]")
            End If

            ' ����и�����δ��ʾ�������ʾ
            If table.Columns.Count > maxCols Then
                builder.AppendLine($"... [����� {table.Columns.Count} �У�����ʾǰ {maxCols} ��]")
            End If

        Catch ex As Exception
            builder.AppendLine($"[����������ʱ����: {ex.Message}]")
        End Try
    End Sub
End Class

