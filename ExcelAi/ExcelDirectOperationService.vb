' ExcelAi\ExcelDirectOperationService.vb
' Excel直接操作服务：执行JSON命令进行Excel操作，不依赖VBA

Imports System.Diagnostics
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq

''' <summary>
''' Excel直接操作服务
''' 支持通过JSON命令直接操作Excel，无需生成VBA代码
''' </summary>
Public Class ExcelDirectOperationService

    Private ReadOnly _excelApp As Application

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <param name="excelApp">Excel应用程序实例</param>
    Public Sub New(excelApp As Application)
        _excelApp = excelApp
    End Sub

#Region "公共方法"

    ''' <summary>
    ''' 执行JSON命令
    ''' </summary>
    ''' <param name="commandJson">命令JSON对象</param>
    ''' <returns>执行是否成功</returns>
    Public Function ExecuteCommand(commandJson As JObject) As Boolean
        Try
            If commandJson Is Nothing Then
                Return False
            End If

            Dim command = commandJson("command")?.ToString()
            Dim params = commandJson("params")

            If String.IsNullOrEmpty(command) Then
                Debug.WriteLine("ExecuteCommand: 命令为空")
                Return False
            End If

            Debug.WriteLine($"执行命令: {command}")

            Select Case command.ToLower()
                Case "writedata", "write", "setvalue", "setvalues"
                    Return ExecuteWriteData(params)
                Case "applyformula", "formula", "calculatesum", "calculate", "range_operations"
                    Return ExecuteApplyFormulaFlexible(commandJson)
                Case "formatrange", "format", "style"
                    Return ExecuteFormatRange(params)
                Case "createchart", "chart"
                    Return ExecuteCreateChart(params)
                Case "cleandata", "clean"
                    Return ExecuteCleanData(params)
                Case "dataanalysis", "analyze"
                    Return ExecuteDataAnalysis(params)
                Case "transformdata", "transform"
                    Return ExecuteTransformData(params)
                Case "generatereport", "report"
                    Return ExecuteGenerateReport(params)
                Case Else
                    Debug.WriteLine($"未知命令: {command}")
                    Return False
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 尝试从AI响应中提取JSON命令
    ''' </summary>
    ''' <param name="aiResponse">AI响应文本</param>
    ''' <returns>提取的命令列表</returns>
    Public Shared Function ExtractCommandsFromResponse(aiResponse As String) As List(Of JObject)
        Dim commands As New List(Of JObject)()

        Try
            ' 查找JSON代码块
            Dim pattern = "```json\s*([\s\S]*?)```"
            Dim matches = System.Text.RegularExpressions.Regex.Matches(aiResponse, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            For Each match As System.Text.RegularExpressions.Match In matches
                Dim jsonStr = match.Groups(1).Value.Trim()
                Try
                    Dim json = JObject.Parse(jsonStr)
                    If json("command") IsNot Nothing Then
                        commands.Add(json)
                    End If
                Catch
                    ' 忽略无效JSON
                End Try
            Next

            ' 如果没有找到代码块，尝试直接解析
            If commands.Count = 0 Then
                Try
                    Dim json = JObject.Parse(aiResponse.Trim())
                    If json("command") IsNot Nothing Then
                        commands.Add(json)
                    End If
                Catch
                    ' 忽略
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"ExtractCommandsFromResponse 出错: {ex.Message}")
        End Try

        Return commands
    End Function

#End Region

#Region "命令执行方法"

    ''' <summary>
    ''' 执行写入数据命令
    ''' </summary>
    Private Function ExecuteWriteData(params As JToken) As Boolean
        Try
            ' 支持多种参数名：targetRange, startCell, range
            Dim targetRange = params("targetRange")?.ToString()
            If String.IsNullOrEmpty(targetRange) Then
                targetRange = params("startCell")?.ToString()
            End If
            If String.IsNullOrEmpty(targetRange) Then
                targetRange = params("range")?.ToString()
            End If
            
            ' 如果有targetSheet，组合成完整地址
            Dim targetSheet = params("targetSheet")?.ToString()
            If Not String.IsNullOrEmpty(targetSheet) AndAlso Not String.IsNullOrEmpty(targetRange) Then
                ' 如果targetRange不包含!，则添加工作表名
                If Not targetRange.Contains("!") Then
                    targetRange = $"{targetSheet}!{targetRange}"
                End If
            End If
            
            ' 支持data或targetData
            Dim data = params("data")
            If data Is Nothing Then
                data = params("targetData")
            End If

            If String.IsNullOrEmpty(targetRange) OrElse data Is Nothing Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("WriteData缺少必要参数：targetRange/startCell 和 data")
                Return False
            End If

            ' 解析目标范围（可能包含工作表名）
            Dim ws As Worksheet
            Dim cellAddress As String = targetRange
            
            If targetRange.Contains("!") Then
                ' 格式: "SheetName!A1" 或 "'Sheet Name'!A1"
                Dim parts = targetRange.Split("!"c)
                Dim sheetName = parts(0).Trim("'"c)
                cellAddress = parts(1)
                
                ' 检查工作表是否存在，不存在则创建
                Try
                    ws = _excelApp.Worksheets(sheetName)
                Catch
                    ' 工作表不存在，创建新的
                    ws = _excelApp.Worksheets.Add()
                    ws.Name = sheetName
                    ShareRibbon.GlobalStatusStrip.ShowInfo($"已创建新工作表: {sheetName}")
                End Try
            Else
                ws = _excelApp.ActiveSheet
            End If
            
            Dim range As Range = ws.Range(cellAddress)

            ' 支持单值或数组
            If data.Type = JTokenType.Array Then
                Dim dataArray = data.ToObject(Of Object()())()
                If dataArray IsNot Nothing AndAlso dataArray.Length > 0 Then
                    Dim rows = dataArray.Length
                    Dim cols = dataArray(0).Length
                    Dim values(rows - 1, cols - 1) As Object

                    For i = 0 To rows - 1
                        For j = 0 To cols - 1
                            values(i, j) = dataArray(i)(j)
                        Next
                    Next

                    range.Resize(rows, cols).Value2 = values
                End If
            Else
                range.Value2 = data.ToString()
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"数据已写入 {targetRange}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWriteData 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"写入数据失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行应用公式命令
    ''' </summary>
    Private Function ExecuteApplyFormula(params As JToken) As Boolean
        Try
            Dim targetRange = params("targetRange")?.ToString()
            Dim formula = params("formula")?.ToString()
            Dim fillDown = If(params("fillDown")?.Value(Of Boolean)(), False)

            If String.IsNullOrEmpty(targetRange) OrElse String.IsNullOrEmpty(formula) Then
                Return False
            End If

            ' 解析目标范围（可能包含工作表名）
            Dim ws As Worksheet
            Dim cellAddress As String = targetRange
            
            If targetRange.Contains("!") Then
                ' 格式: "SheetName!A1:B10" 或 "'Sheet Name'!A1:B10"
                Dim parts = targetRange.Split("!"c)
                Dim sheetName = parts(0).Trim("'"c)
                cellAddress = parts(1)
                
                ' 检查工作表是否存在，不存在则创建
                Try
                    ws = _excelApp.Worksheets(sheetName)
                Catch
                    ws = _excelApp.Worksheets.Add()
                    ws.Name = sheetName
                    ShareRibbon.GlobalStatusStrip.ShowInfo($"已创建新工作表: {sheetName}")
                End Try
            Else
                ws = _excelApp.ActiveSheet
            End If
            
            Dim range As Range = ws.Range(cellAddress)

            ' 确保公式以=开头
            If Not formula.StartsWith("=") Then
                formula = "=" & formula
            End If

            If fillDown AndAlso range.Rows.Count > 1 Then
                ' 只设置第一个单元格，然后向下填充
                range.Cells(1, 1).Formula = formula
                range.Cells(1, 1).AutoFill(range, XlAutoFillType.xlFillDefault)
            Else
                range.Formula = formula
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"公式已应用到 {targetRange}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyFormula 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"应用公式失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 灵活的公式应用方法 - 支持多种JSON格式
    ''' </summary>
    Private Function ExecuteApplyFormulaFlexible(commandJson As JObject) As Boolean
        Try
            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim params = commandJson("params")

            ' 尝试多种参数格式（支持 targetRange, range, target 等别名）
            Dim targetRange = If(params?("targetRange")?.ToString(),
                             If(params?("range")?.ToString(),
                             If(params?("target")?.ToString(),
                             If(commandJson("target")?.ToString(),
                             If(commandJson("targetRange")?.ToString(),
                             If(commandJson("range")?.ToString(), ""))))))

            Dim formula = If(params?("formula")?.ToString(),
                          If(commandJson("formula")?.ToString(), ""))

            Dim fillDown = If(params?("fillDown")?.Value(Of Boolean)(),
                           If(params?("autoFill")?.Value(Of Boolean)(),
                           If(commandJson("autoFill")?.Value(Of Boolean)(), False)))

            ' 处理 operations 数组格式
            Dim operations = If(commandJson("operations"), params?("operations"))
            If operations IsNot Nothing AndAlso operations.Type = JTokenType.Array Then
                For Each op In operations
                    Dim opRange = op("range")?.ToString()
                    Dim opFormula = op("formula")?.ToString()
                    If Not String.IsNullOrEmpty(opRange) AndAlso Not String.IsNullOrEmpty(opFormula) Then
                        ApplySingleFormula(ws, opRange, opFormula, fillDown)
                    End If
                Next
                Return True
            End If

            ' 单一公式格式
            If Not String.IsNullOrEmpty(formula) Then
                ' 如果没有目标范围，尝试从公式中推断（如 C1=A1+B1）
                If String.IsNullOrEmpty(targetRange) AndAlso formula.Contains("=") Then
                    Dim parts = formula.Split("="c)
                    If parts.Length >= 2 AndAlso System.Text.RegularExpressions.Regex.IsMatch(parts(0).Trim(), "^[A-Za-z]+\d+") Then
                        targetRange = parts(0).Trim()
                        formula = "=" & String.Join("=", parts.Skip(1))
                    End If
                End If

                If String.IsNullOrEmpty(targetRange) Then
                    targetRange = "C1" ' 默认目标
                End If

                Return ApplySingleFormula(ws, targetRange, formula, fillDown)
            End If

            Return False
        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyFormulaFlexible 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 应用单个公式到范围
    ''' </summary>
    Private Function ApplySingleFormula(ws As Worksheet, targetRange As String, formula As String, fillDown As Boolean) As Boolean
        Try
            ' 处理动态范围如 "C1:C" + last_row
            If targetRange.Contains("+") OrElse targetRange.Contains("last_row") Then
                Dim usedRange = ws.UsedRange
                Dim lastRow = usedRange.Row + usedRange.Rows.Count - 1
                targetRange = $"C1:C{lastRow}"
            End If

            Dim range As Range = ws.Range(targetRange)

            If Not formula.StartsWith("=") Then
                formula = "=" & formula
            End If

            If fillDown AndAlso range.Rows.Count > 1 Then
                range.Cells(1, 1).Formula = formula
                range.Cells(1, 1).AutoFill(range, XlAutoFillType.xlFillDefault)
            Else
                range.Formula = formula
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"公式已应用到 {targetRange}")
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ApplySingleFormula 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行格式化范围命令
    ''' </summary>
    Private Function ExecuteFormatRange(params As JToken) As Boolean
        Try
            Dim targetRange = params("range")?.ToString()
            If String.IsNullOrEmpty(targetRange) Then
                targetRange = params("targetRange")?.ToString()
            End If

            If String.IsNullOrEmpty(targetRange) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim range As Range = ws.Range(targetRange)

            ' 应用样式
            Dim style = params("style")?.ToString()
            Select Case style?.ToLower()
                Case "header"
                    range.Font.Bold = True
                    range.Interior.Color = RGB(68, 114, 196) ' 蓝色背景
                    range.Font.Color = RGB(255, 255, 255) ' 白色字体
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter
                Case "total"
                    range.Font.Bold = True
                    range.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlDouble
                Case "data"
                    range.Borders.LineStyle = XlLineStyle.xlContinuous
                    range.Borders.Weight = XlBorderWeight.xlThin
            End Select

            ' 应用单独的格式属性
            If params("bold")?.Value(Of Boolean)() = True Then
                range.Font.Bold = True
            End If

            If params("italic")?.Value(Of Boolean)() = True Then
                range.Font.Italic = True
            End If

            Dim fontSize = params("fontSize")?.Value(Of Integer)()
            If fontSize.HasValue AndAlso fontSize.Value > 0 Then
                range.Font.Size = fontSize.Value
            End If

            Dim bgColor = params("backgroundColor")?.ToString()
            If Not String.IsNullOrEmpty(bgColor) Then
                range.Interior.Color = ParseColor(bgColor)
            End If

            Dim fontColor = params("fontColor")?.ToString()
            If Not String.IsNullOrEmpty(fontColor) Then
                range.Font.Color = ParseColor(fontColor)
            End If

            If params("borders")?.Value(Of Boolean)() = True Then
                range.Borders.LineStyle = XlLineStyle.xlContinuous
                range.Borders.Weight = XlBorderWeight.xlThin
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"格式已应用到 {targetRange}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatRange 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行创建图表命令
    ''' </summary>
    Private Function ExecuteCreateChart(params As JToken) As Boolean
        Try
            Dim chartType = params("type")?.ToString()
            Dim dataRange = params("dataRange")?.ToString()
            Dim title = params("title")?.ToString()
            Dim position = params("position")?.ToString()

            If String.IsNullOrEmpty(dataRange) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim sourceRange As Range = ws.Range(dataRange)

            ' 确定图表类型
            Dim xlChartType As XlChartType = XlChartType.xlColumnClustered
            Select Case chartType?.ToLower()
                Case "line"
                    xlChartType = XlChartType.xlLine
                Case "pie"
                    xlChartType = XlChartType.xlPie
                Case "bar"
                    xlChartType = XlChartType.xlBarClustered
                Case "scatter"
                    xlChartType = XlChartType.xlXYScatter
                Case "area"
                    xlChartType = XlChartType.xlArea
                Case Else
                    xlChartType = XlChartType.xlColumnClustered
            End Select

            ' 确定图表位置
            Dim positionRange As Range
            If Not String.IsNullOrEmpty(position) Then
                positionRange = ws.Range(position)
            Else
                ' 默认放在数据右边
                positionRange = sourceRange.Offset(0, sourceRange.Columns.Count + 1)
            End If

            ' 创建图表
            Dim chartObj As ChartObject = ws.ChartObjects.Add(
                positionRange.Left,
                positionRange.Top,
                400,
                300)

            With chartObj.Chart
                .ChartType = xlChartType
                .SetSourceData(sourceRange)

                If Not String.IsNullOrEmpty(title) Then
                    .HasTitle = True
                    .ChartTitle.Text = title
                End If

                ' 添加图例
                Dim hasLegend = params("hasLegend")?.Value(Of Boolean)()
                .HasLegend = If(hasLegend, True, True)
            End With

            ShareRibbon.GlobalStatusStrip.ShowInfo("图表已创建")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCreateChart 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行数据清洗命令
    ''' </summary>
    Private Function ExecuteCleanData(params As JToken) As Boolean
        Try
            Dim operation = params("operation")?.ToString()
            Dim targetRange = params("range")?.ToString()

            If String.IsNullOrEmpty(targetRange) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim range As Range = ws.Range(targetRange)

            Select Case operation?.ToLower()
                Case "removeduplicates"
                    range.RemoveDuplicates(Columns:=Array.CreateInstance(GetType(Integer), range.Columns.Count))

                Case "fillempty"
                    Dim fillValue = params("fillValue")?.ToString()
                    If String.IsNullOrEmpty(fillValue) Then fillValue = "0"

                    For Each cell As Range In range.Cells
                        If cell.Value Is Nothing OrElse String.IsNullOrEmpty(cell.Value.ToString()) Then
                            cell.Value = fillValue
                        End If
                    Next

                Case "trim"
                    For Each cell As Range In range.Cells
                        If cell.Value IsNot Nothing AndAlso TypeOf cell.Value Is String Then
                            cell.Value = cell.Value.ToString().Trim()
                        End If
                    Next

                Case "replace"
                    Dim findText = params("findText")?.ToString()
                    Dim replaceText = params("replaceText")?.ToString()
                    If Not String.IsNullOrEmpty(findText) Then
                        range.Replace(findText, replaceText, XlLookAt.xlPart)
                    End If
            End Select

            ShareRibbon.GlobalStatusStrip.ShowInfo($"数据清洗已完成: {operation}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCleanData 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行数据分析命令
    ''' </summary>
    Private Function ExecuteDataAnalysis(params As JToken) As Boolean
        Try
            Dim analysisType = params("type")?.ToString()
            Dim sourceRange = params("sourceRange")?.ToString()
            Dim targetRange = params("targetRange")?.ToString()

            If String.IsNullOrEmpty(sourceRange) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim source As Range = ws.Range(sourceRange)

            Select Case analysisType?.ToLower()
                Case "summary"
                    ' 生成基本统计摘要
                    Return GenerateSummary(source, targetRange)

                Case "pivot"
                    ' 创建透视表
                    Return CreatePivotTable(source, targetRange, params)

                Case "groupby"
                    ' 分组汇总
                    Return GroupByAnalysis(source, targetRange, params)

                Case "ranking"
                    ' 排名分析
                    Return RankingAnalysis(source, targetRange, params)
            End Select

            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteDataAnalysis 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行数据转换命令
    ''' </summary>
    Private Function ExecuteTransformData(params As JToken) As Boolean
        Try
            Dim operation = params("operation")?.ToString()
            Dim sourceRange = params("sourceRange")?.ToString()
            Dim targetRange = params("targetRange")?.ToString()

            If String.IsNullOrEmpty(sourceRange) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim source As Range = ws.Range(sourceRange)

            Select Case operation?.ToLower()
                Case "transpose"
                    ' 转置数据
                    Dim target As Range = ws.Range(If(targetRange, "A1"))
                    source.Copy()
                    target.PasteSpecial(Paste:=XlPasteType.xlPasteAll, Operation:=XlPasteSpecialOperation.xlPasteSpecialOperationNone, Transpose:=True)
                    _excelApp.CutCopyMode = False

                Case "split"
                    ' 拆分列
                    Dim delimiter = params("delimiter")?.ToString()
                    If Not String.IsNullOrEmpty(delimiter) Then
                        source.TextToColumns(Destination:=source, DataType:=XlTextParsingType.xlDelimited, Other:=True, OtherChar:=delimiter)
                    End If

                Case "merge"
                    ' 合并列
                    Return MergeColumns(source, targetRange, params)
            End Select

            ShareRibbon.GlobalStatusStrip.ShowInfo($"数据转换已完成: {operation}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteTransformData 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行生成报表命令
    ''' </summary>
    Private Function ExecuteGenerateReport(params As JToken) As Boolean
        Try
            Dim reportType = params("type")?.ToString()
            Dim sourceRange = params("sourceRange")?.ToString()
            Dim targetSheet = params("targetSheet")?.ToString()
            Dim title = params("title")?.ToString()
            Dim includeChart = If(params("includeChart")?.Value(Of Boolean)(), False)

            If String.IsNullOrEmpty(sourceRange) Then
                Return False
            End If

            Dim sourceWs As Worksheet = _excelApp.ActiveSheet
            Dim source As Range = sourceWs.Range(sourceRange)

            ' 创建或获取目标工作表
            Dim targetWs As Worksheet
            If Not String.IsNullOrEmpty(targetSheet) Then
                Try
                    targetWs = _excelApp.Worksheets(targetSheet)
                Catch
                    targetWs = _excelApp.Worksheets.Add()
                    targetWs.Name = targetSheet
                End Try
            Else
                targetWs = _excelApp.Worksheets.Add()
                targetWs.Name = "报表_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
            End If

            ' 添加标题
            If Not String.IsNullOrEmpty(title) Then
                targetWs.Range("A1").Value = title
                targetWs.Range("A1").Font.Size = 16
                targetWs.Range("A1").Font.Bold = True
            End If

            ' 复制数据
            Dim dataStartRow = If(String.IsNullOrEmpty(title), 1, 3)
            source.Copy(targetWs.Range($"A{dataStartRow}"))

            ' 格式化表头
            Dim headerRange = targetWs.Range($"A{dataStartRow}").Resize(1, source.Columns.Count)
            headerRange.Font.Bold = True
            headerRange.Interior.Color = RGB(68, 114, 196)
            headerRange.Font.Color = RGB(255, 255, 255)

            ' 添加边框
            Dim dataRange = targetWs.Range($"A{dataStartRow}").Resize(source.Rows.Count, source.Columns.Count)
            dataRange.Borders.LineStyle = XlLineStyle.xlContinuous
            dataRange.Borders.Weight = XlBorderWeight.xlThin

            ' 自动调整列宽
            targetWs.Columns.AutoFit()

            ' 如果需要图表
            If includeChart Then
                ExecuteCreateChart(New JObject From {
                    {"dataRange", $"A{dataStartRow}:" & ShareRibbon.ExcelContextService.GetExcelColumnName(source.Columns.Count) & (dataStartRow + source.Rows.Count - 1)},
                    {"position", $"A{dataStartRow + source.Rows.Count + 2}"},
                    {"title", title}
                })
            End If

            targetWs.Activate()
            ShareRibbon.GlobalStatusStrip.ShowInfo($"报表已生成: {targetWs.Name}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteGenerateReport 出错: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 解析颜色字符串
    ''' </summary>
    Private Function ParseColor(colorStr As String) As Integer
        Try
            If colorStr.StartsWith("#") Then
                colorStr = colorStr.Substring(1)
            End If

            If colorStr.Length = 6 Then
                Dim r = Convert.ToInt32(colorStr.Substring(0, 2), 16)
                Dim g = Convert.ToInt32(colorStr.Substring(2, 2), 16)
                Dim b = Convert.ToInt32(colorStr.Substring(4, 2), 16)
                Return RGB(r, g, b)
            End If
        Catch
        End Try

        Return RGB(255, 255, 255)
    End Function

    ''' <summary>
    ''' 生成统计摘要
    ''' </summary>
    Private Function GenerateSummary(source As Range, targetRange As String) As Boolean
        Try
            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim target As Range

            If Not String.IsNullOrEmpty(targetRange) Then
                target = ws.Range(targetRange)
            Else
                target = source.Offset(0, source.Columns.Count + 2)
            End If

            ' 添加摘要标题
            target.Value = "数据摘要"
            target.Font.Bold = True

            ' 统计信息
            target.Offset(1, 0).Value = "行数"
            target.Offset(1, 1).Value = source.Rows.Count
            target.Offset(2, 0).Value = "列数"
            target.Offset(2, 1).Value = source.Columns.Count

            ' 对数值列计算统计
            Dim row = 3
            For col = 1 To source.Columns.Count
                Dim colRange = source.Columns(col)
                If _excelApp.WorksheetFunction.IsNumber(colRange.Cells(2, 1).Value) Then
                    target.Offset(row, 0).Value = $"列{col}合计"
                    target.Offset(row, 1).Formula = $"=SUM({colRange.Address})"
                    row += 1
                    target.Offset(row, 0).Value = $"列{col}平均"
                    target.Offset(row, 1).Formula = $"=AVERAGE({colRange.Address})"
                    row += 1
                End If
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo("统计摘要已生成")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"GenerateSummary 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 创建透视表
    ''' </summary>
    Private Function CreatePivotTable(source As Range, targetRange As String, params As JToken) As Boolean
        ' 透视表创建比较复杂，这里提供基本实现框架
        ShareRibbon.GlobalStatusStrip.ShowWarning("透视表功能正在开发中，请使用VBA代码")
        Return False
    End Function

    ''' <summary>
    ''' 分组汇总分析
    ''' </summary>
    Private Function GroupByAnalysis(source As Range, targetRange As String, params As JToken) As Boolean
        ShareRibbon.GlobalStatusStrip.ShowWarning("分组汇总功能正在开发中，请使用VBA代码")
        Return False
    End Function

    ''' <summary>
    ''' 排名分析
    ''' </summary>
    Private Function RankingAnalysis(source As Range, targetRange As String, params As JToken) As Boolean
        ShareRibbon.GlobalStatusStrip.ShowWarning("排名分析功能正在开发中，请使用VBA代码")
        Return False
    End Function

    ''' <summary>
    ''' 合并列
    ''' </summary>
    Private Function MergeColumns(source As Range, targetRange As String, params As JToken) As Boolean
        Try
            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim delimiter = If(params("delimiter")?.ToString(), " ")
            Dim target As Range

            If Not String.IsNullOrEmpty(targetRange) Then
                target = ws.Range(targetRange)
            Else
                target = source.Offset(0, source.Columns.Count + 1)
            End If

            ' 构建CONCAT公式
            For row = 1 To source.Rows.Count
                Dim formula As New System.Text.StringBuilder("=CONCAT(")
                For col = 1 To source.Columns.Count
                    If col > 1 Then
                        formula.Append($",""{delimiter}"",")
                    End If
                    formula.Append(source.Cells(row, col).Address)
                Next
                formula.Append(")")

                target.Offset(row - 1, 0).Formula = formula.ToString()
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo("列合并已完成")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"MergeColumns 出错: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

End Class
