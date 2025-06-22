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
                    ' ѡ��û�����ݣ������ͬ sheetName ������
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


    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean)
        ' �����ҪԤ��
        Dim previewTool As New EnhancedPreviewAndConfirm()
        ' �����û�Ԥ��������
        If previewTool.PreviewAndConfirmVbaExecution(vbaCode) Then
            Debug.Print("Ԥ���������û�ͬ��ִ�д���: " & vbaCode)
            Return True
        Else
            ' �û�ȡ����ܾ�
            Return False
        End If
    End Function


    ' ִ�� JavaScript ���룬֧�ֲ���Excel����
    Protected Function ExecuteJavaScript(jsCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(jsCode, preview) Then
                    Return False
                End If
            End If

            ' ���������� - ��ͨJS����Excel����JS
            Dim isExcelJS As Boolean = jsCode.Contains("Excel.") OrElse
                                  jsCode.Contains("ActiveXObject") OrElse
                                  jsCode.Contains("Application") OrElse
                                  jsCode.Contains("Workbook")

            If isExcelJS Then
                ' �����ű�����������ִ�в���Excel��JavaScript
                Dim scriptEngine As Object = CreateObject("MSScriptControl.ScriptControl")
                scriptEngine.Language = "JScript"

                ' ���ö�ExcelӦ�ó��������
                scriptEngine.AddObject("excelApp", Globals.ThisAddIn.Application, True)

                ' ����ִ�д���
                Dim scriptCode As String =
                "function executeExcelJS() {" & vbCrLf &
                "  try {" & vbCrLf &
                "    // Excel����ΪexcelApp�����ṩ" & vbCrLf &
                "    " & jsCode & vbCrLf &
                "    return 'JS����ִ�гɹ�';" & vbCrLf &
                "  } catch(e) {" & vbCrLf &
                "    return 'JSִ�д���: ' + e.message;" & vbCrLf &
                "  }" & vbCrLf &
                "}" & vbCrLf &
                "executeExcelJS();"

                ' ִ��JavaScript����
                Dim result As String = scriptEngine.Eval(scriptCode)
                GlobalStatusStrip.ShowInfo(result)
                Return True
            Else
                ' ������ͨJavaScript��ʹ��WebView2ִ��
                Dim scriptResult As Task(Of String) = ChatBrowser.ExecuteScriptAsync(jsCode)
                scriptResult.Wait() ' �ȴ�ִ�����

                ' ��ʾ���
                If Not String.IsNullOrEmpty(scriptResult.Result) Then
                    Dim resultStr As String = scriptResult.Result.Trim(""""c) ' �Ƴ�JSON�ַ�������
                    GlobalStatusStrip.ShowInfo("JSִ�н��: " + resultStr)
                End If

                Return True
            End If
        Catch ex As Exception
            MessageBox.Show("ִ��JavaScript����ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' �ṩExcelӦ�ó������
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function

    ' ʵ��Excel��ʽ����' ִ��Excel��ʽ���� - ��ǿ��֧�ָ�ֵ��Ԥ��

    Protected Overrides Function EvaluateFormula(formulaCode As String, preview As Boolean) As Boolean
        Try
            ' ����Ƿ��Ǹ�ֵ��� (���� C1=A1+B1)
            Dim isAssignment As Boolean = Regex.IsMatch(formulaCode, "^[A-Za-z]+[0-9]+\s*=")

            If isAssignment Then
                ' ������ֵ���
                Dim parts As String() = formulaCode.Split(New Char() {"="c}, 2)
                Dim targetCell As String = parts(0).Trim()
                Dim formula As String = parts(1).Trim()

                ' �����ʽ��=��ͷ�����Ƴ�
                If formula.StartsWith("=") Then
                    formula = formula.Substring(1)
                End If

                ' �����ҪԤ������ʾԤ���Ի���
                If preview Then
                    Dim excel As Object = Globals.ThisAddIn.Application
                    Dim currentValue As Object = Nothing
                    Try
                        currentValue = excel.Range(targetCell).Value
                    Catch ex As Exception
                        ' ��Ԫ����ܲ�����ֵ
                    End Try

                    ' ������ֵ
                    Dim newValue As Object = excel.Evaluate(formula)

                    ' ����Ԥ���Ի���
                    Dim previewMsg As String = $"��Ҫ�ڵ�Ԫ�� {targetCell} ��Ӧ�ù�ʽ:" & vbCrLf & vbCrLf &
                                          $"={formula}" & vbCrLf & vbCrLf &
                                          $"��ǰֵ: {If(currentValue Is Nothing, "(��)", currentValue)}" & vbCrLf &
                                          $"��ֵ: {If(newValue Is Nothing, "(��)", newValue)}"

                    Dim result As DialogResult = MessageBox.Show(previewMsg, "Excel��ʽԤ��",
                                                          MessageBoxButtons.OKCancel,
                                                          MessageBoxIcon.Information)

                    If result <> DialogResult.OK Then
                        Return False
                    End If
                End If

                ' ִ�и�ֵ
                Dim range As Object = Globals.ThisAddIn.Application.Range(targetCell)
                range.Formula = "=" & formula

                GlobalStatusStrip.ShowInfo($"��ʽ '={formula}' ��Ӧ�õ���Ԫ�� {targetCell}")
                Return True
            Else
                ' ��ͨ��ʽ���� (��������ֵ)
                ' ȥ�����ܵĵȺ�ǰ׺
                If formulaCode.StartsWith("=") Then
                    formulaCode = formulaCode.Substring(1)
                End If

                ' ���㹫ʽ���
                Dim result As Object = Globals.ThisAddIn.Application.Evaluate(formulaCode)

                ' �����ҪԤ������ʾ������
                If preview Then
                    Dim previewMsg As String = $"��ʽ������:" & vbCrLf & vbCrLf &
                                         $"={formulaCode}" & vbCrLf & vbCrLf &
                                         $"���: {If(result Is Nothing, "(��)", result)}"

                    MessageBox.Show(previewMsg, "Excel��ʽ���", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    ' ��ʾ���
                    GlobalStatusStrip.ShowInfo($"��ʽ '={formulaCode}' �ļ�����: {result}")
                End If

                Return True
            End If
        Catch ex As Exception
            MessageBox.Show("ִ��Excel��ʽʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function


    ' ִ��SQL��ѯ
    Protected Function ExecuteSqlQuery(sqlCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(sqlCode, preview) Then
                    Return False
                End If
            End If

            ' ��ȡӦ�ó�����Ϣ
            Dim appInfo As ApplicationInfo = GetApplication()

            Dim activeWorkbook As Object = Globals.ThisAddIn.Application.ActiveWorkbook

            ' ������ѯ��
            Dim activeSheet As Object = Globals.ThisAddIn.Application.ActiveSheet
            Dim queryTable As Object = Nothing

                ' ��ȡ���õĵ�Ԫ������
                Dim targetCell As Object = activeSheet.Range("A1")

                ' ����SQL�����ַ��� (ʾ��ʹ�õ�ǰ��������Ϊ����Դ)
                Dim connString As String = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" &
                                      activeWorkbook.FullName & ";Extended Properties='Excel 12.0 Xml;HDR=YES';"

                ' ������ѯ����
                queryTable = activeSheet.QueryTables.Add(connString, targetCell, sqlCode)

                ' ���ò�ѯ����
                queryTable.RefreshStyle = 1 ' xlOverwriteCells
                queryTable.BackgroundQuery = False

                ' ִ�в�ѯ
                queryTable.Refresh(False)

                GlobalStatusStrip.ShowWarning("SQL��ѯ��ִ��")
            Return True
        Catch ex As Exception
            MessageBox.Show("ִ��SQL��ѯʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' ִ��PowerQuery/M����
    Protected Function ExecutePowerQuery(mCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(mCode, preview) Then
                    Return False
                End If
            End If

            ' ��ȡӦ�ó�����Ϣ
            Dim appInfo As ApplicationInfo = GetApplication()

            ' PowerQueryִ����Ҫ�ϸ��ӵ�ʵ�֣�������ṩ�������
            Dim excelApp = Globals.ThisAddIn.Application
                Dim wb As Object = excelApp.ActiveWorkbook

                ' ���Excel�汾�Ƿ�֧��PowerQuery
                Dim versionSupported As Boolean = excelApp.Version >= 15 ' Excel 2013�����ϰ汾

                If Not versionSupported Then
                    GlobalStatusStrip.ShowWarning("PowerQuery��ҪExcel 2013����߰汾")
                    Return False
                End If

                ' PowerQueryִ���߼���Ҫ���ݾ�������ʵ��
                GlobalStatusStrip.ShowWarning("PowerQuery����ִ�й������ڿ�����")
            Return True
        Catch ex As Exception
            MessageBox.Show("ִ��PowerQuery����ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' ִ��Python����
    Protected Function ExecutePython(pythonCode As String, preview As Boolean) As Boolean
        Try
            If preview Then
                If Not RunCodePreview(pythonCode, preview) Then
                    Return False
                End If
            End If

            ' ��ȡӦ�ó�����Ϣ
            Dim appInfo As ApplicationInfo = GetApplication()

            Dim excelApp = Globals.ThisAddIn.Application

                ' ���Excel�汾�Ƿ�֧��Python (Excel 365)
                Dim versionSupported As Boolean = False

                Try
                    ' ���Է���Python���������֧�ֻ��׳��쳣
                    Dim pythonObj As Object = excelApp.PythonExecute("print('test')")
                    versionSupported = True
                Catch
                    versionSupported = False
                End Try

                If Not versionSupported Then
                    ' �������Python�����ã����Գ���ͨ���ⲿPython������ִ��
                    GlobalStatusStrip.ShowWarning("��Excel�汾��֧������Python������ʹ���ⲿPython...")

                    ' ������ʱPython�ļ�
                    Dim tempFile As String = Path.Combine(Path.GetTempPath(), "excel_python_" & Guid.NewGuid().ToString() & ".py")
                    File.WriteAllText(tempFile, pythonCode)

                    ' ʹ��Process��ִ��Python�ű�
                    Dim startInfo As New ProcessStartInfo With {
                    .FileName = "python", ' ����Python�Ѱ�װ����PATH��
                    .Arguments = tempFile,
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True
                }

                    Using process As Process = Process.Start(startInfo)
                        Dim output As String = process.StandardOutput.ReadToEnd()
                        Dim error1 As String = process.StandardError.ReadToEnd()
                        process.WaitForExit()

                        If Not String.IsNullOrEmpty(error1) Then
                        GlobalStatusStrip.ShowWarning("Pythonִ�д���: " & error1)
                    Else
                            GlobalStatusStrip.ShowWarning("Pythonִ�н��: " & output)
                        End If
                    End Using

                    ' ɾ����ʱ�ļ�
                    Try
                        File.Delete(tempFile)
                    Catch
                        ' �����������
                    End Try
                Else
                    ' ʹ��Excel����Pythonִ�д���
                    Dim result As Object = excelApp.PythonExecute(pythonCode)
                    GlobalStatusStrip.ShowWarning("Python������ִ��")
                End If

            Return True
        Catch ex As Exception
            MessageBox.Show("ִ��Python����ʱ����: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
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

