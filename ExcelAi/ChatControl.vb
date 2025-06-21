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

        ' ���� SelectionChange �¼� - ʹ���µ����ط���
        AddHandler Globals.ThisAddIn.Application.SheetSelectionChange, AddressOf GetSelectionContentExcel

    End Sub

    ' ����ԭ�е�Override�����Լ��ݻ���
    Protected Overrides Sub GetSelectionContent(target As Object)
        ' ����Ǵ�Excel��SheetSelectionChange�¼����ã�targetӦ����Worksheet
        If TypeOf target Is Microsoft.Office.Interop.Excel.Worksheet Then
            ' ��ȡ��ǰѡ�еķ�Χ
            Dim selection = Globals.ThisAddIn.Application.Selection
            If TypeOf selection Is Microsoft.Office.Interop.Excel.Range Then
                GetSelectionContentExcel(target, DirectCast(selection, Microsoft.Office.Interop.Excel.Range))
            End If
        End If
    End Sub

    ' ���һ���µ����ط���������Excel���¼�
    Private Sub GetSelectionContentExcel(Sh As Microsoft.Office.Interop.Excel.Worksheet, Target As Microsoft.Office.Interop.Excel.Range)
        If Me.Visible AndAlso selectedCellChecked Then
            Dim sheetName As String = Sh.Name
            Dim address As String = Target.Address(False, False)
            Dim key As String = $"{sheetName}"

            ' ���ѡ�з�Χ�ĵ�Ԫ������
            Dim cellCount As Integer = Target.Cells.Count

            ' ���ѡ���˶����Ԫ���������Ϊ���ã������Ƿ�������
            If cellCount > 1 Then
                AddSelectedContentItem(key, address)
            Else
                ' ֻ�е�����Ԫ��ʱ���ż���Ƿ�������
                Dim hasContent As Boolean = False
                For Each cell As Microsoft.Office.Interop.Excel.Range In Target
                    If cell.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(cell.Value.ToString()) Then
                        hasContent = True
                        Exit For
                    End If
                Next

                If hasContent Then
                    ' ѡ�е�Ԫ�������ݣ�����µ���
                    AddSelectedContentItem(key, address)
                Else
                    ' ѡ�е�Ԫ��û�����ݣ������ͬ sheetName ������
                    ClearSelectedContentBySheetName(key)
                End If
            End If
        End If
    End Sub

    Private Async Sub AddSelectedContentItem(sheetName As String, address As String)
        'Dim ctrlKey As Boolean = False
        Dim ctrlKey As Boolean = (Control.ModifierKeys And Keys.Control) = Keys.Control

        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
    $"addSelectedContentItem({JsonConvert.SerializeObject(sheetName)}, {JsonConvert.SerializeObject(address)}, {ctrlKey.ToString().ToLower()})"
)
    End Sub

    ' �������ض� sheetName �ķ���
    Private Async Sub ClearSelectedContentBySheetName(sheetName As String)
        Await ChatBrowser.CoreWebView2.ExecuteScriptAsync(
        $"clearSelectedContentBySheetName({JsonConvert.SerializeObject(sheetName)})"
    )
    End Sub
    ' ��ʼ��ʱע����� HTML �ṹ
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ��ʼ�� WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
        InitializeSettings()
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

    'Protected Overrides Function RunCode(code As String) As Object
    '    Try
    '        Globals.ThisAddIn.Application.Run(code)
    '        Return True
    '    Catch ex As Runtime.InteropServices.COMException
    '        VBAxceptionHandle(ex)
    '        Return False
    '    Catch ex As Exception
    '        MessageBox.Show("ִ�д���ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Return False
    '    End Try
    'End Function
    'Protected Overrides Function RunCode(vbaTepModel As String, vbaCode As String) As Object
    '    Try
    '        ' ����ԭʼ���������ã����������Ϊ�ر�/�л�������NullReferenceException
    '        Dim originalWorkbook As Microsoft.Office.Interop.Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook

    '        ' �쳣��������һ��ʼ��û�д򿪵Ĺ�����
    '        If originalWorkbook Is Nothing Then
    '            MessageBox.Show("��ǰ�޻���������޷�ִ�С�", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            Return False
    '        End If

    '        Dim previewTool As New EnhancedPreviewAndConfirm()

    '        ' �����û�Ԥ��������
    '        If previewTool.PreviewAndConfirmVbaExecution(vbaTepModel, vbaCode) Then
    '            Debug.Print("ִ�д���: " & vbaCode)
    '            Globals.ThisAddIn.Application.Run(vbaTepModel)
    '            Return True
    '        Else
    '            ' �û�ȡ����ܾ�
    '            Return False
    '        End If

    '    Catch ex As Runtime.InteropServices.COMException
    '        VBAxceptionHandle(ex)
    '        Return False
    '    Catch ex As Exception
    '        MessageBox.Show("ִ�д���ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Return False
    '    End Try
    'End Function


    ' ִ��ǰ�˴����� VBA ����Ƭ��
    Protected Overrides Function RunCode(vbaCode As String)

        Dim previewTool As New EnhancedPreviewAndConfirm()

        ' �����û�Ԥ��������
        If previewTool.PreviewAndConfirmVbaExecution(vbaCode) Then
            Debug.Print("ִ�д���: " & vbaCode)
        Else
            ' �û�ȡ����ܾ�
            Return False
        End If

        ' ��ȡ VBA ��Ŀ
        Dim vbProj As VBProject = GetVBProject()

        ' ��ӿ�ֵ���
        If vbProj Is Nothing Then
            Return False
        End If

        Dim vbComp As VBComponent = Nothing
        Dim tempModuleName As String = "TempMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

        Try
            ' ������ʱģ��
            vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            vbComp.Name = tempModuleName

            ' �������Ƿ��Ѱ��� Sub/Function ����
            If ContainsProcedureDeclaration(vbaCode) Then
                ' �����Ѱ�������������ֱ�����
                vbComp.CodeModule.AddFromString(vbaCode)

                ' ���ҵ�һ����������ִ��
                Dim procName As String = FindFirstProcedureName(vbComp)
                If Not String.IsNullOrEmpty(procName) Then
                    Globals.ThisAddIn.Application.Run(tempModuleName & "." & procName)
                Else
                    'MessageBox.Show("�޷��ڴ������ҵ���ִ�еĹ���")
                    GlobalStatusStrip.ShowWarning("�޷��ڴ������ҵ���ִ�еĹ���")
                End If
            Else
                ' ���벻�������������������װ�� Auto_Run ������
                Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                           vbaCode & vbNewLine &
                                           "End Sub"
                vbComp.CodeModule.AddFromString(wrappedCode)

                ' ִ�� Auto_Run ����
                Globals.ThisAddIn.Application.Run(tempModuleName & ".Auto_Run")

            End If

        Catch ex As Exception
            MessageBox.Show("ִ�� VBA ����ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' ���۳ɹ�����ʧ�ܣ���ɾ����ʱģ��
            Try
                If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                    vbProj.VBComponents.Remove(vbComp)
                End If
            Catch
                ' �����������
            End Try
        End Try
    End Function

    Private Function GetSelectedRangeContent() As String
        Try
            ' ��ȡ sheetContentItems ������
            Dim selectedContents As String = String.Join("|", sheetContentItems.Values.Select(Function(item) item.Item1.Text))

            ' ���� selectedContents ����ȡÿ����������ѡ���ĵ�Ԫ������
            Dim parsedContents As New StringBuilder()
            If Not String.IsNullOrEmpty(selectedContents) Then
                Dim sheetSelections = selectedContents.Split("|"c)
                For Each sheetSelection In sheetSelections
                    Dim parts = sheetSelection.Split("["c)
                    If parts.Length = 2 Then
                        Dim sheetName = parts(0)
                        Dim ranges = parts(1).TrimEnd("]"c).Split(","c)
                        For Each range In ranges
                            Dim content = GetRangeContent(sheetName, range)
                            If Not String.IsNullOrEmpty(content) Then
                                parsedContents.AppendLine($"{sheetName}��{range}:{content}")
                            End If
                        Next
                    End If
                Next
            End If

            ' �� parsedContents ���뵽 question ��
            If parsedContents.Length > 0 Then
                Return "�����ṩ��ѡ�е�������Ϊ�ο���{" & parsedContents.ToString() & "}"
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Private Function GetRangeContent(sheetName As String, rangeAddress As String) As String
        Try
            Dim sheet = Globals.ThisAddIn.Application.Sheets(sheetName)
            Dim range = sheet.Range(rangeAddress)
            Dim value = range.Value2

            If value Is Nothing Then
                Return String.Empty
            End If

            If TypeOf value Is System.Object(,) Then
                Dim array = DirectCast(value, System.Object(,))
                Dim rows = array.GetLength(0)
                Dim cols = array.GetLength(1)
                Dim result As New StringBuilder()

                For i = 1 To rows
                    For j = 1 To cols
                        If array(i, j) IsNot Nothing Then
                            result.Append(array(i, j).ToString() & vbTab)
                        End If
                    Next
                    result.AppendLine()
                Next

                Return result.ToString().TrimEnd()
            Else
                Return value.ToString()
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Excel", OfficeApplicationType.Excel)
    End Function
    Protected Overrides Sub SendChatMessage(message As String)
        ' �������ʵ��word�������߼�
        Debug.Print(message)
        Send(message)
    End Sub

    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            ' ��ȡ��ǰ��������ѡ������
            Dim activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
            Dim selection = Globals.ThisAddIn.Application.Selection

            ' �����ѡ��������Ϊ Range ����
            If selection IsNot Nothing AndAlso TypeOf selection Is Microsoft.Office.Interop.Excel.Range Then
                Dim selectedRange As Microsoft.Office.Interop.Excel.Range = DirectCast(selection, Microsoft.Office.Interop.Excel.Range)

                ' �������ݹ����������� ParseFile �Ľṹ
                Dim contentBuilder As New StringBuilder()
                contentBuilder.AppendLine(vbCrLf & "--- �û�ѡ�е�WorkbookSheet�ο��������� ---")

                ' ��ӻ��������Ϣ
                contentBuilder.AppendLine($"������: {Path.GetFileName(activeWorkbook.FullName)}")

                ' ��ȡѡ��Ĺ�������Ϣ
                Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = selectedRange.Worksheet
                Dim sheetName As String = worksheet.Name

                ' ��ӹ�������Ϣ
                contentBuilder.AppendLine($"������: {sheetName}")

                ' ��ȡѡ������ķ�Χ��ַ
                Dim address As String = selectedRange.Address(False, False)
                contentBuilder.AppendLine($"  ʹ�÷�Χ: {address}")

                ' ��ȡѡ�������еĵ�Ԫ������
                Dim usedRange As Microsoft.Office.Interop.Excel.Range = selectedRange

                ' ��ȡ�����������Ϣ
                Dim firstRow As Integer = usedRange.Row
                Dim firstCol As Integer = usedRange.Column
                Dim lastRow As Integer = firstRow + usedRange.Rows.Count - 1
                Dim lastCol As Integer = firstCol + usedRange.Columns.Count - 1

                ' ���ƶ�ȡ�ĵ�Ԫ����������ֹ���ݹ���
                Dim maxRows As Integer = Math.Min(lastRow, firstRow + 30)
                Dim maxCols As Integer = Math.Min(lastCol, firstCol + 10)

                ' �����Ԫ���ȡ����
                For rowIndex As Integer = firstRow To maxRows
                    For colIndex As Integer = firstCol To maxCols
                        Try
                            Dim cell As Microsoft.Office.Interop.Excel.Range = worksheet.Cells(rowIndex, colIndex)
                            Dim cellValue As Object = cell.Value

                            If cellValue IsNot Nothing Then
                                Dim cellAddress As String = $"{GetExcelColumnName(colIndex)}{rowIndex}"
                                contentBuilder.AppendLine($"  {cellAddress}: {cellValue}")
                            End If
                        Catch cellEx As Exception
                            Debug.WriteLine($"��ȡ��Ԫ��ʱ����: {cellEx.Message}")
                            ' ����������һ����Ԫ��
                        End Try
                    Next
                Next

                ' ����и����л���δ��ʾ�������ʾ
                If lastRow > maxRows Then
                    contentBuilder.AppendLine($"  ... ���� {lastRow - firstRow + 1} �У�����ʾǰ {maxRows - firstRow + 1} ��")
                End If
                If lastCol > maxCols Then
                    contentBuilder.AppendLine($"  ... ���� {lastCol - firstCol + 1} �У�����ʾǰ {maxCols - firstCol + 1} ��")
                End If

                contentBuilder.AppendLine("--- WorkbookSheet�ο����ݵ������ ---" & vbCrLf)

                ' ��ѡ��������ӵ���Ϣ��
                message &= contentBuilder.ToString()
            End If
        Catch ex As Exception
            Debug.WriteLine($"��ȡѡ�е�Ԫ������ʱ����: {ex.Message}")
            ' ����ʱ�����ѡ�����ݣ���������ԭʼ��Ϣ
        End Try
        Return message
    End Function

    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            ' ����һ���µ� Excel Ӧ�ó���ʵ����Ϊ����Ӱ�쵱ǰ��������
            Dim excelApp As New Microsoft.Office.Interop.Excel.Application
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Try
                workbook = excelApp.Workbooks.Open(filePath, ReadOnly:=True)
                Dim contentBuilder As New StringBuilder()

                contentBuilder.AppendLine($"�ļ�: {Path.GetFileName(filePath)} ������������:")

                ' ����ÿ��������
                For Each worksheet As Microsoft.Office.Interop.Excel.Worksheet In workbook.Worksheets
                    Dim sheetName As String = worksheet.Name
                    contentBuilder.AppendLine($"������: {sheetName}")

                    ' ��ȡʹ�÷�Χ
                    Dim usedRange As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
                    If usedRange IsNot Nothing Then
                        Dim lastRow As Integer = usedRange.Row + usedRange.Rows.Count - 1
                        Dim lastCol As Integer = usedRange.Column + usedRange.Columns.Count - 1

                        ' ���ƶ�ȡ�ĵ�Ԫ����������ֹ�ļ�����
                        Dim maxRows As Integer = Math.Min(lastRow, 30)
                        Dim maxCols As Integer = Math.Min(lastCol, 10)

                        contentBuilder.AppendLine($"  ʹ�÷�Χ: {GetExcelColumnName(usedRange.Column)}{usedRange.Row}:{GetExcelColumnName(lastCol)}{lastRow}")

                        ' ��ȡ��Ԫ������
                        For rowIndex As Integer = usedRange.Row To maxRows
                            For colIndex As Integer = usedRange.Column To maxCols
                                Try
                                    Dim cell As Microsoft.Office.Interop.Excel.Range = worksheet.Cells(rowIndex, colIndex)
                                    Dim cellValue As Object = cell.Value

                                    If cellValue IsNot Nothing Then
                                        Dim cellAddress As String = $"{GetExcelColumnName(colIndex)}{rowIndex}"
                                        contentBuilder.AppendLine($"  {cellAddress}: {cellValue}")
                                    End If
                                Catch cellEx As Exception
                                    Debug.WriteLine($"��ȡ��Ԫ��ʱ����: {cellEx.Message}")
                                    ' ����������һ����Ԫ��
                                End Try
                            Next
                        Next

                        ' ����и����л���δ��ʾ�������ʾ
                        If lastRow > maxRows Then
                            contentBuilder.AppendLine($"  ... ���� {lastRow - usedRange.Row + 1} �У�����ʾǰ {maxRows - usedRange.Row + 1} ��")
                        End If
                        If lastCol > maxCols Then
                            contentBuilder.AppendLine($"  ... ���� {lastCol - usedRange.Column + 1} �У�����ʾǰ {maxCols - usedRange.Column + 1} ��")
                        End If
                    End If

                    contentBuilder.AppendLine()
                Next

                Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Excel",
                .ParsedContent = contentBuilder.ToString(),
                .RawData = Nothing ' ����ѡ��洢�������ݹ���������
            }

            Finally
                ' ������Դ
                If workbook IsNot Nothing Then
                    workbook.Close(SaveChanges:=False)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
                End If

                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        Catch ex As Exception
            Debug.WriteLine($"���� Excel �ļ�ʱ����: {ex.Message}")
            Return New FileContentResult With {
            .FileName = Path.GetFileName(filePath),
            .FileType = "Excel",
            .ParsedContent = $"[���� Excel �ļ�ʱ����: {ex.Message}]"
        }
        End Try
    End Function

    ' ������������������ת��Ϊ Excel �������� 1->A, 27->AA��
    Private Function GetExcelColumnName(columnIndex As Integer) As String
        Dim dividend As Integer = columnIndex
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Chr(65 + modulo) & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

    ' ʵ�ֻ�ȡ��ǰ Excel ����Ŀ¼�ķ���
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
End Class

