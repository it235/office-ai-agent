' WordAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' 添加此引用
Imports Newtonsoft.Json.Linq
Imports Microsoft.Office.Interop.Excel

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Protected Overrides Async Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub

    Protected Overrides Async Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        If String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) Then
            MsgBox("请输入ApiKey！")
            Return
        End If

        If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) Then
            MsgBox("请选择大模型！")
            Return
        End If

        ' 获取选中的单元格区域
        Dim selection As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selection IsNot Nothing Then
            Dim cellValues As New StringBuilder()

            Dim cellIndices As New StringBuilder()
            Dim cellList As New List(Of String)

            ' 按列遍历，每列用局部变量记录连续空行数
            For col As Integer = selection.Column To selection.Column + selection.Columns.Count
                Dim emptyCount As Integer = 0
                For row As Integer = selection.Row To selection.Row + selection.Rows.Count - 1
                    Dim cell As Excel.Range = selection.Worksheet.Cells(row, col)
                    ' 如果存在非空内容，则处理，并重置空计数
                    If cell.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cell.Value.ToString()) Then
                        cellValues.AppendLine(cell.Value.ToString())
                        cellList.Add(cell.Address(False, False))
                        emptyCount = 0
                    Else
                        emptyCount += 1
                        If emptyCount >= 50 Then
                            Exit For  ' 本列连续50行为空，退出当前列循环
                        End If
                    End If
                Next
            Next


            ' 按照矩阵展开方式显示单元格索引
            Dim groupedCells = cellList.GroupBy(Function(c) Regex.Replace(c, "\d", ""))
            For Each group In groupedCells
                cellIndices.AppendLine(String.Join(",", group))
            Next

            ' 显示所有单元格的值
            If cellValues.Length > 0 Then
                Dim previewForm As New TextPreviewForm(cellIndices.ToString())
                previewForm.ShowDialog()

                If previewForm.IsConfirmed Then
                    ' 获取查询内容和数据
                    Dim question As String = cellValues.ToString
                    question = previewForm.InputText & “。你只需要返回markdown格式的表格即可，别的什么都不要说，不要任何其他多余的文字。原始数据如下：“ & question

                    Dim requestBody As String = CreateRequestBody(question)

                    ' 发送 HTTP 请求并获取响应
                    Dim response As String = Await SendHttpRequest(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)

                    ' 如果响应为空，则终止执行
                    If String.IsNullOrEmpty(response) Then
                        Return
                    End If

                    ' 解析并写入响应数据
                    WriteResponseToSheet(response)
                End If
            Else
                MsgBox("选中的单元格无文本内容！")
            End If
        Else
            MsgBox("请选择一个单元格区域！")

        End If

    End Sub

    Private Sub WriteResponseToSheet(response As String)
        Try
            Dim parsedResponse As JObject = JObject.Parse(response)
            Dim cellValue As String = parsedResponse("choices")(0)("message")("content").ToString()

            Dim lines() As String = Split(cellValue, vbLf)

            Dim wsOutput As Worksheet = GetOrCreateSheet("AI结果")
            ' 激活工作表
            wsOutput.Activate()
            ' 清空输出表
            wsOutput.Cells.Clear()

            'wsOutput.Range("F8").Value = cellValue

            ' 暂停屏幕更新和计算
            Globals.ThisAddIn.Application.ScreenUpdating = False
            Globals.ThisAddIn.Application.Calculation = XlCalculation.xlCalculationManual

            ' 写入表头
            Dim columns() As String = Split(Trim(lines(0)), "|")
            For i As Integer = 1 To UBound(columns)
                wsOutput.Cells(1, i).Value = Trim(columns(i))
            Next i


            ' 写入表格数据（跳过分隔线和表头）
            For i As Integer = 2 To UBound(lines)
                If Trim(lines(i)) <> "" And Not Left(Trim(lines(i)), 1) = "-" Then ' 跳过空行和分隔线
                    columns = Split(Trim(lines(i)), "|")
                    For j As Integer = 1 To UBound(columns) - 1
                        wsOutput.Cells(i, j).Value = Trim(columns(j))
                    Next j
                End If
            Next i

            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Calculation = XlCalculation.xlCalculationAutomatic

            ' 提示完成
            MsgBox("数据已成功写入 AI结果！", vbInformation)
        Catch ex As Exception
            MsgBox("解析响应时出错：" & ex.Message, vbCritical)
        End Try
    End Sub

    Private Function GetOrCreateSheet(sheetName As String) As Worksheet
        Dim ws As Worksheet = Nothing
        Try
            ws = Globals.ThisAddIn.Application.Sheets(sheetName)
        Catch ex As Exception
            ' 如果工作表不存在，则创建一个新的工作表
            ws = Globals.ThisAddIn.Application.Sheets.Add()
            ws.Name = sheetName
        End Try
        Return ws
    End Function

    Protected Overrides Function GetApplication() As Object
        Return Globals.ThisAddIn.Application
    End Function
End Class