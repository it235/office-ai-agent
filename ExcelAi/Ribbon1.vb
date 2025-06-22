' WordAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' ��Ӵ�����
Imports Newtonsoft.Json.Linq
Imports Microsoft.Office.Interop.Excel

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Protected Overrides Async Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Async Sub WebResearchButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub ' �޸� SpotlightButton_Click ������������˫��
    Protected Overrides Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ' ��ȡ�۹��ʵ��
            Dim spotlight As Spotlight = Spotlight.GetInstance()

            ' �ж��Ƿ���˫��
            Dim button As RibbonButton = TryCast(sender, RibbonButton)

            ' ����Ƿ�˫�� (��ʱ�����ж�˫��)
            If IsDoubleClick() Then
                ' ˫�� - ��ʾ��ɫѡ��Ի���
                spotlight.ShowColorDialog()
            Else
                ' ���� - �л��۹��״̬
                spotlight.Toggle()
            End If
        Catch ex As Exception
            MsgBox("����۹�ƹ���ʱ����" & ex.Message, vbCritical)
        End Try
    End Sub

    ' ���ڼ��˫���ı���
    Private _lastClickTime As DateTime = DateTime.MinValue

    ' ����Ƿ�Ϊ˫����������ε�����С��300���룬����Ϊ˫����
    Private Function IsDoubleClick() As Boolean
        Dim currentTime As DateTime = DateTime.Now
        Dim isDouble As Boolean = (currentTime - _lastClickTime).TotalMilliseconds < 300

        ' �������˫��������������ʱ��
        If Not isDouble Then
            _lastClickTime = currentTime
        Else
            ' �����˫����������ʱ�䣬����������ε��������Ϊ���˫��
            _lastClickTime = DateTime.MinValue
        End If

        Return isDouble
    End Function

    Protected Overrides Async Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        If String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) Then
            MsgBox("������ApiKey��")
            Return
        End If

        If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) Then
            MsgBox("��ѡ���ģ�ͣ�")
            Return
        End If

        ' ��ȡѡ�еĵ�Ԫ������
        Dim selection As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selection IsNot Nothing Then
            Dim cellValues As New StringBuilder()

            Dim cellIndices As New StringBuilder()
            Dim cellList As New List(Of String)

            ' ���б�����ÿ���þֲ�������¼����������
            For col As Integer = selection.Column To selection.Column + selection.Columns.Count
                Dim emptyCount As Integer = 0
                For row As Integer = selection.Row To selection.Row + selection.Rows.Count - 1
                    Dim cell As Excel.Range = selection.Worksheet.Cells(row, col)
                    ' ������ڷǿ����ݣ����������ÿռ���
                    If cell.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cell.Value.ToString()) Then
                        cellValues.AppendLine(cell.Value.ToString())
                        cellList.Add(cell.Address(False, False))
                        emptyCount = 0
                    Else
                        emptyCount += 1
                        If emptyCount >= 50 Then
                            Exit For  ' ��������50��Ϊ�գ��˳���ǰ��ѭ��
                        End If
                    End If
                Next
            Next


            ' ���վ���չ����ʽ��ʾ��Ԫ������
            Dim groupedCells = cellList.GroupBy(Function(c) Regex.Replace(c, "\d", ""))
            For Each group In groupedCells
                cellIndices.AppendLine(String.Join(",", group))
            Next

            ' ��ʾ���е�Ԫ���ֵ
            If cellValues.Length > 0 Then
                Dim previewForm As New TextPreviewForm(cellIndices.ToString())
                previewForm.ShowDialog()

                If previewForm.IsConfirmed Then
                    ' ��ȡ��ѯ���ݺ�����
                    Dim question As String = cellValues.ToString
                    question = previewForm.InputText & ������ֻ��Ҫ����markdown��ʽ�ı�񼴿ɣ����ʲô����Ҫ˵����Ҫ�κ�������������֡�ԭʼ�������£��� & question

                    Dim requestBody As String = CreateRequestBody(question)

                    ' ���� HTTP ���󲢻�ȡ��Ӧ
                    Dim response As String = Await SendHttpRequest(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)

                    ' �����ӦΪ�գ�����ִֹ��
                    If String.IsNullOrEmpty(response) Then
                        Return
                    End If

                    ' ������д����Ӧ����
                    WriteResponseToSheet(response)
                End If
            Else
                MsgBox("ѡ�еĵ�Ԫ�����ı����ݣ�")
            End If
        Else
            MsgBox("��ѡ��һ����Ԫ������")

        End If

    End Sub

    Private Sub WriteResponseToSheet(response As String)
        Try
            Dim parsedResponse As JObject = JObject.Parse(response)
            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()

            Dim lines() As String = Split(cellValue, vbLf)

            Dim wsOutput As Worksheet = GetOrCreateSheet("AI���")
            ' �������
            wsOutput.Activate()
            ' ��������
            wsOutput.Cells.Clear()

            'wsOutput.Range("F8").Value = cellValue

            ' ��ͣ��Ļ���ºͼ���
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.Calculation = XlCalculation.xlCalculationManual

            ' д���ͷ
            Dim columns() As String = Split(Trim(lines(0)), "|")
            For i As Integer = 1 To UBound(columns)
                wsOutput.Cells(1, i).Value = Trim(columns(i))
            Next i


            ' д�������ݣ������ָ��ߺͱ�ͷ��
            For i As Integer = 2 To UBound(lines)
                If Trim(lines(i)) <> "" And Not Left(Trim(lines(i)), 1) = "-" Then ' �������кͷָ���
                    columns = Split(Trim(lines(i)), "|")
                    For j As Integer = 1 To UBound(columns) - 1
                        wsOutput.Cells(i, j).Value = Trim(columns(j))
                    Next j
                End If
            Next i

            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Calculation = XlCalculation.xlCalculationAutomatic

            ' ��ʾ���
            GlobalStatusStripAll.ShowWarning("�����ѳɹ�д�� AI�����")
        Catch ex As Exception
            MsgBox("������Ӧʱ����" & ex.Message, vbCritical)
        End Try
    End Sub

    Private Function GetOrCreateSheet(sheetName As String) As Worksheet
        Dim ws As Worksheet = Nothing
        Try
            ws = Globals.ThisAddIn.Application.Sheets(sheetName)
        Catch ex As Exception
            ' ������������ڣ��򴴽�һ���µĹ�����
            ws = Globals.ThisAddIn.Application.Sheets.Add()
            ws.Name = sheetName
        End Try
        Return ws
    End Function

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Excel", OfficeApplicationType.Excel)
    End Function
End Class