' ExcelAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Tools.Ribbon
Imports ShareRibbon  ' 添加此引用
Imports Newtonsoft.Json
Imports Microsoft.Office.Interop.Excel

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Protected Overrides Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Async Sub WebCaptureButton_Click(sender As Object, e As RibbonControlEventArgs)
        Await StartAgentFromRibbonAsync(Agent.QuickActionCatalog.ExcelRibbonWebCapture)
    End Sub
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
        Await StartAgentFromRibbonAsync(Agent.QuickActionCatalog.ExcelRibbonAnalyze)
    End Sub

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

    ' MCPButton_Click 已在 BaseOfficeRibbon 中提供共用实现，Excel 不需要差异化逻辑，故不再重写。

    Protected Overrides Async Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs)
        Await StartAgentFromRibbonAsync(Agent.QuickActionCatalog.ExcelRibbonProofread)
    End Sub

    Protected Overrides Async Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Await StartAgentFromRibbonAsync(Agent.QuickActionCatalog.ExcelRibbonReformat)
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

    ' AI续写功能 - Excel 交给 Agent 根据表格上下文自动补全
    Protected Overrides Async Sub ContinuationButton_Click(sender As Object, e As RibbonControlEventArgs)
        Await StartAgentFromRibbonAsync(Agent.QuickActionCatalog.ExcelRibbonContinuation)
    End Sub



    ' 模板排版功能 - 交给 Agent 自动识别当前表格结构和可用模板/样式
    Protected Overrides Async Sub TemplateFormatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Await StartAgentFromRibbonAsync(Agent.QuickActionCatalog.ExcelRibbonTemplateFormat)
    End Sub

    Private Async Function StartAgentFromRibbonAsync(action As Agent.QuickAction) As Task
        Try
            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(350)

            Dim chatCtrl = ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                GlobalStatusStripAll.ShowWarning("无法获取 AI 助手面板")
                Return
            End If

            chatCtrl.StartAgentFromHostEntry(action)
        Catch ex As Exception
            GlobalStatusStripAll.ShowWarning($"启动 AI Agent 失败: {ex.Message}")
        End Try
    End Function
End Class
