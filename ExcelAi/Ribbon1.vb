' ExcelAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' 添加此引用
Imports Newtonsoft.Json.Linq
Imports Microsoft.Office.Interop.Excel

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Protected Overrides Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Sub WebResearchButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub ' 修改 SpotlightButton_Click 方法处理单击和双击
    Protected Overrides Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            ' 获取聚光灯实例
            Dim spotlight As Spotlight = Spotlight.GetInstance()

            ' 判断是否是双击
            Dim button As RibbonButton = TryCast(sender, RibbonButton)

            ' 检查是否双击 (用时间间隔判断双击)
            If IsDoubleClick() Then
                ' 双击 - 显示颜色选择对话框
                spotlight.ShowColorDialog()
            Else
                ' 单击 - 切换聚光灯状态
                spotlight.Toggle()
            End If
        Catch ex As Exception
            MsgBox("激活聚光灯功能时出错：" & ex.Message, vbCritical)
        End Try
    End Sub

    ' 用于检测双击的变量
    Private _lastClickTime As DateTime = DateTime.MinValue

    ' 检查是否为双击（如果两次点击间隔小于300毫秒，则视为双击）
    Private Function IsDoubleClick() As Boolean
        Dim currentTime As DateTime = DateTime.Now
        Dim isDouble As Boolean = (currentTime - _lastClickTime).TotalMilliseconds < 300

        ' 如果不是双击，则更新最后点击时间
        If Not isDouble Then
            _lastClickTime = currentTime
        Else
            ' 如果是双击，则重置时间，以免连续多次点击被误判为多次双击
            _lastClickTime = DateTime.MinValue
        End If

        Return isDouble
    End Function

    Protected Overrides Async Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        If String.IsNullOrWhiteSpace(ConfigSettings.ApiKey) Then
            GlobalStatusStripAll.ShowWarning("请输入ApiKey！")
            Return
        End If

        If String.IsNullOrWhiteSpace(ConfigSettings.ApiUrl) Then
            GlobalStatusStripAll.ShowWarning("请选择大模型！")
            Return
        End If

        ' 获取选中的单元格区域
        Dim selection As Excel.Range = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
        If selection IsNot Nothing Then
            Dim cellValues As New StringBuilder()

            Dim cellIndices As New StringBuilder()
            Dim cellList As New List(Of String)

            ' 按列遍历，每列用局部变量记录连续空行数
            For col As Integer = selection.Column To selection.Column + selection.Columns.Count - 1
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

                    Dim requestBody As String = LLMUtil.CreateRequestBody(question)

                    ' 发送 HTTP 请求并获取响应
                    Dim response As String = Await LLMUtil.SendHttpRequest(ConfigSettings.ApiUrl, ConfigSettings.ApiKey, requestBody)

                    ' 如果响应为空，则终止执行
                    If String.IsNullOrEmpty(response) Then
                        Return
                    End If

                    ' 解析并写入响应数据
                    WriteResponseToSheet(response)
                End If
            Else
                GlobalStatusStripAll.ShowWarning("选中的单元格无文本内容！")
            End If
        Else
            GlobalStatusStripAll.ShowWarning("请选择一个单元格区域！")

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
            GlobalStatusStripAll.ShowWarning("数据已成功写入 AI结果！")
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

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Excel", OfficeApplicationType.Excel)
    End Function

    ' Deepseek按钮点击事件实现
    Protected Overrides Sub DeepseekButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowDeepseekTaskPane()
    End Sub

    ' Doubao按钮点击事件实现
    Protected Overrides Sub DoubaoButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowDoubaoTaskPane()
    End Sub

    ' 批量数据生成按钮点击事件实现
    Protected Overrides Async Sub BatchDataGenButton_Click(sender As Object, e As RibbonControlEventArgs)
        Using batchDataForm As New BatchDataGenerationForm()
            If batchDataForm.ShowDialog() <> DialogResult.OK Then Return

            Dim fields = batchDataForm.Fields
            Dim rowCount = batchDataForm.RowCount

            Dim excelApp As Excel.Application = Globals.ThisAddIn.Application
            Dim activeSheet As Excel.Worksheet = TryCast(excelApp.ActiveSheet, Excel.Worksheet)
            If activeSheet Is Nothing Then
                GlobalStatusStripAll.ShowWarning("无法获取当前工作表")
                Return
            End If

            Try
                GlobalStatusStripAll.ShowWarning($"正在生成 {rowCount} 条数据，请稍候...")
                Dim svc As New BatchDataService()
                Dim jsonText = Await svc.GenerateBatchDataAsync(fields, rowCount)

                If String.IsNullOrEmpty(jsonText) Then
                    GlobalStatusStripAll.ShowWarning("数据生成失败，请检查 AI 配置")
                    Return
                End If

                ' 提取 JSON 数组（去掉可能的 Markdown 代码块）
                Dim cleanJson = jsonText.Trim()
                Dim startIdx = cleanJson.IndexOf("[")
                Dim endIdx = cleanJson.LastIndexOf("]")
                If startIdx < 0 OrElse endIdx <= startIdx Then
                    GlobalStatusStripAll.ShowWarning("AI 返回格式异常，未能解析 JSON 数组")
                    Return
                End If
                cleanJson = cleanJson.Substring(startIdx, endIdx - startIdx + 1)

                Dim rows = Newtonsoft.Json.Linq.JArray.Parse(cleanJson)

                ' 写入表头（第1行）
                Dim headerRow = 1
                For Each field In fields
                    Dim col = ColumnLetterToIndex(field.CellColumn)
                    If col > 0 Then activeSheet.Cells(headerRow, col).Value = field.FieldName
                Next

                ' 写入数据（从第2行开始）
                For i = 0 To rows.Count - 1
                    Dim rowObj = TryCast(rows(i), Newtonsoft.Json.Linq.JObject)
                    If rowObj Is Nothing Then Continue For
                    For Each field In fields
                        Dim col = ColumnLetterToIndex(field.CellColumn)
                        If col > 0 Then
                            activeSheet.Cells(headerRow + 1 + i, col).Value = rowObj(field.FieldName)?.ToString()
                        End If
                    Next
                Next

                GlobalStatusStripAll.ShowWarning($"成功生成 {rows.Count} 条数据")
            Catch ex As Exception
                GlobalStatusStripAll.ShowWarning($"数据生成失败: {ex.Message}")
                Debug.WriteLine($"[BatchDataGen] 错误: {ex}")
            End Try
        End Using
    End Sub

    ''' <summary>将列字母转换为列索引（A→1，B→2，AA→27）</summary>
    Private Function ColumnLetterToIndex(col As String) As Integer
        If String.IsNullOrWhiteSpace(col) Then Return 0
        col = col.Trim().ToUpper()
        Dim result As Integer = 0
        For Each ch As Char In col
            If ch < "A"c OrElse ch > "Z"c Then Return 0
            result = result * 26 + (AscW(ch) - AscW("A"c) + 1)
        Next
        Return result
    End Function

    ' MCP按钮点击事件实现
    Protected Overrides Sub MCPButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' 创建并显示MCP配置表单
        Dim mcpConfigForm As New MCPConfigForm()
        If mcpConfigForm.ShowDialog() = DialogResult.OK Then
            ' 在需要时可以集成到ChatControl调用MCP服务
        End If
    End Sub

    Protected Overrides Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show("Excel校对功能正在开发中...", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Protected Overrides Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show("Excel排版功能正在开发中...", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' 一键翻译功能 - Excel实现（翻译选中单元格内容）
    Protected Overrides Async Sub TranslateButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim excelApp = Globals.ThisAddIn.Application
            Dim selection As Excel.Range = TryCast(excelApp.Selection, Excel.Range)

            If selection Is Nothing OrElse selection.Cells.Count = 0 Then
                MessageBox.Show("请先选择要翻译的单元格区域。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 显示翻译操作对话框
            Dim actionForm As New ShareRibbon.TranslateActionForm(True, "Excel")
            If actionForm.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            ' 收集单元格内容
            Dim cellTexts As New List(Of String)()
            Dim cellRanges As New List(Of Excel.Range)()

            For Each cell As Excel.Range In selection.Cells
                If cell.Value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cell.Value.ToString()) Then
                    cellTexts.Add(cell.Value.ToString())
                    cellRanges.Add(cell)
                End If
            Next

            If cellTexts.Count = 0 Then
                MessageBox.Show("选中的单元格没有文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 更新设置
            Dim settings = ShareRibbon.TranslateSettings.Load()
            settings.SourceLanguage = actionForm.SourceLanguage
            settings.TargetLanguage = actionForm.TargetLanguage
            settings.CurrentDomain = actionForm.SelectedDomain
            settings.OutputMode = actionForm.OutputMode
            settings.Save()

            ShareRibbon.GlobalStatusStripAll.ShowWarning($"正在翻译 {cellTexts.Count} 个单元格...")

            ' 使用Excel文档翻译服务翻译
            Dim translateService As New ExcelDocumentTranslateService()
            Dim results = Await translateService.TranslateCellsAsync(cellTexts, cellRanges, settings)

            ShareRibbon.GlobalStatusStripAll.ShowWarning($"翻译完成，共处理 {cellTexts.Count} 个单元格")

        Catch ex As Exception
            MessageBox.Show("翻译过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' AI续写功能 - Excel暂不支持（续写主要用于文档类型）
    Protected Overrides Sub ContinuationButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show("AI续写功能主要用于Word和PowerPoint文档，Excel暂不支持此功能。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub



    ' 模板排版功能 - Excel暂不支持
    Protected Overrides Sub TemplateFormatButton_Click(sender As Object, e As RibbonControlEventArgs)
        MessageBox.Show("模板排版功能主要用于Word和PowerPoint文档，Excel暂不支持此功能。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class