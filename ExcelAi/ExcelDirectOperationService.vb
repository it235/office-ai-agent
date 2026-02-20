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
                ' === 基础操作 ===
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
                ' === 数据操作 ===
                Case "sortdata", "sort"
                    Return ExecuteSortData(params)
                Case "filterdata", "filter"
                    Return ExecuteFilterData(params)
                Case "removeduplicates"
                    Return ExecuteRemoveDuplicates(params)
                Case "conditionalformat"
                    Return ExecuteConditionalFormat(params)
                Case "mergecells", "merge"
                    Return ExecuteMergeCells(params)
                Case "autofit"
                    Return ExecuteAutoFit(params)
                Case "findreplace"
                    Return ExecuteFindReplace(params)
                Case "createpivottable", "pivot"
                    Return ExecuteCreatePivotTable(params)
                ' === 工作表操作 ===
                Case "createsheet"
                    Return ExecuteCreateSheet(params)
                Case "deletesheet"
                    Return ExecuteDeleteSheet(params)
                Case "renamesheet"
                    Return ExecuteRenameSheet(params)
                Case "copysheet"
                    Return ExecuteCopySheet(params)
                ' === 高级功能 ===
                Case "insertrowcol"
                    Return ExecuteInsertRowCol(params)
                Case "deleterowcol"
                    Return ExecuteDeleteRowCol(params)
                Case "hiderowcol"
                    Return ExecuteHideRowCol(params)
                Case "protectsheet"
                    Return ExecuteProtectSheet(params)
                ' === VBA回退 ===
                Case "executevba", "vba"
                    Return ExecuteVBA(params)
                ' === 旧版兼容 ===
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
                ShareRibbon.GlobalStatusStrip.ShowWarning("FormatRange: 缺少 range 参数")
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim range As Range = Nothing
            Dim chartObj As ChartObject = Nothing

            ' 先检查是否是图表对象（如 "Chart 1"）
            Try
                chartObj = ws.ChartObjects(targetRange)
            Catch
                chartObj = Nothing
            End Try

            If chartObj IsNot Nothing Then
                ' 是图表对象，格式化图表
                Return ExecuteFormatChart(chartObj, params)
            Else
                ' 是单元格范围
                Try
                    range = ws.Range(targetRange)
                Catch ex As Exception
                    ShareRibbon.GlobalStatusStrip.ShowWarning($"FormatRange: 无法找到范围 '{targetRange}': {ex.Message}")
                    Return False
                End Try

                ' 应用样式
                Dim style = params("style")?.ToString()
                Select Case style?.ToLower()
                    Case "header"
                        range.Font.Bold = True
                        range.Interior.Color = RGB(68, 114, 196)
                        range.Font.Color = RGB(255, 255, 255)
                        range.HorizontalAlignment = XlHAlign.xlHAlignCenter
                    Case "total"
                        range.Font.Bold = True
                        range.Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlDouble
                    Case "data"
                        range.Borders.LineStyle = XlLineStyle.xlContinuous
                        range.Borders.Weight = XlBorderWeight.xlThin
                End Select

                ' 应用单独的格式属性
                Dim boldParam = params("bold")
                If boldParam IsNot Nothing AndAlso boldParam.Value(Of Boolean)() = True Then
                    range.Font.Bold = True
                End If

                Dim italicParam = params("italic")
                If italicParam IsNot Nothing AndAlso italicParam.Value(Of Boolean)() = True Then
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

                Dim bordersParam = params("borders")
                If bordersParam IsNot Nothing Then
                    Dim bordersValue = bordersParam.ToString().ToLower()
                    If bordersValue = "true" OrElse bordersValue = "all" Then
                        range.Borders.LineStyle = XlLineStyle.xlContinuous
                        range.Borders.Weight = XlBorderWeight.xlThin
                    ElseIf bordersValue = "outline" Then
                        range.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin)
                    ElseIf bordersValue = "none" OrElse bordersValue = "false" Then
                        range.Borders.LineStyle = XlLineStyle.xlLineStyleNone
                    End If
                End If

                ShareRibbon.GlobalStatusStrip.ShowInfo($"格式已应用到 {targetRange}")
                Return True
            End If

        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatRange 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"格式化失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行格式化图表命令
    ''' </summary>
    Private Function ExecuteFormatChart(chartObj As ChartObject, params As JToken) As Boolean
        Try
            With chartObj.Chart
                ' 获取样式参数
                Dim styleParams = params("style")
                If styleParams IsNot Nothing AndAlso styleParams.Type = JTokenType.Object Then
                    ' 从 style 对象中读取属性
                    Dim fontSize = styleParams("fontSize")?.Value(Of Integer)()
                    Dim fontColor = styleParams("fontColor")?.ToString()
                    Dim bgColor = styleParams("backgroundColor")?.ToString()
                    Dim boldParam = styleParams("bold")

                    ' 格式化图表标题
                    If .HasTitle Then
                        If fontSize.HasValue AndAlso fontSize.Value > 0 Then
                            .ChartTitle.Font.Size = fontSize.Value
                        End If
                        If Not String.IsNullOrEmpty(fontColor) Then
                            .ChartTitle.Font.Color = ParseColor(fontColor)
                        End If
                        If boldParam IsNot Nothing AndAlso boldParam.Value(Of Boolean)() = True Then
                            .ChartTitle.Font.Bold = True
                        End If
                    End If

                    ' 格式化坐标轴
                    Try
                        If .Axes IsNot Nothing Then
                            For Each axis In .Axes
                                If fontSize.HasValue AndAlso fontSize.Value > 0 Then
                                    axis.TickLabels.Font.Size = fontSize.Value
                                End If
                                If Not String.IsNullOrEmpty(fontColor) Then
                                    axis.TickLabels.Font.Color = ParseColor(fontColor)
                                End If
                            Next
                        End If
                    Catch
                    End Try

                    ' 格式化图例
                    If .HasLegend Then
                        If fontSize.HasValue AndAlso fontSize.Value > 0 Then
                            .Legend.Font.Size = fontSize.Value
                        End If
                        If Not String.IsNullOrEmpty(fontColor) Then
                            .Legend.Font.Color = ParseColor(fontColor)
                        End If
                    End If

                    ' 格式化图表区背景
                    If Not String.IsNullOrEmpty(bgColor) Then
                        .ChartArea.Interior.Color = ParseColor(bgColor)
                    End If
                Else
                    ' 从根参数中读取属性（向后兼容）
                    Dim fontSize = params("fontSize")?.Value(Of Integer)()
                    Dim fontColor = params("fontColor")?.ToString()
                    Dim bgColor = params("backgroundColor")?.ToString()
                    Dim boldParam = params("bold")

                    ' 格式化图表标题
                    If .HasTitle Then
                        If fontSize.HasValue AndAlso fontSize.Value > 0 Then
                            .ChartTitle.Font.Size = fontSize.Value
                        End If
                        If Not String.IsNullOrEmpty(fontColor) Then
                            .ChartTitle.Font.Color = ParseColor(fontColor)
                        End If
                        If boldParam IsNot Nothing AndAlso boldParam.Value(Of Boolean)() = True Then
                            .ChartTitle.Font.Bold = True
                        End If
                    End If

                    ' 格式化坐标轴
                    Try
                        If .Axes IsNot Nothing Then
                            For Each axis In .Axes
                                If fontSize.HasValue AndAlso fontSize.Value > 0 Then
                                    axis.TickLabels.Font.Size = fontSize.Value
                                End If
                                If Not String.IsNullOrEmpty(fontColor) Then
                                    axis.TickLabels.Font.Color = ParseColor(fontColor)
                                End If
                            Next
                        End If
                    Catch
                    End Try

                    ' 格式化图例
                    If .HasLegend Then
                        If fontSize.HasValue AndAlso fontSize.Value > 0 Then
                            .Legend.Font.Size = fontSize.Value
                        End If
                        If Not String.IsNullOrEmpty(fontColor) Then
                            .Legend.Font.Color = ParseColor(fontColor)
                        End If
                    End If

                    ' 格式化图表区背景
                    If Not String.IsNullOrEmpty(bgColor) Then
                        .ChartArea.Interior.Color = ParseColor(bgColor)
                    End If
                End If
            End With

            ShareRibbon.GlobalStatusStrip.ShowInfo($"图表格式已应用")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatChart 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"格式化图表失败: {ex.Message}")
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
            Dim categoryAxisParam = params("categoryAxis")
            Dim seriesNamesParam = params("seriesNames")
            Dim legendPosition = params("legendPosition")?.ToString()
            Dim plotBy = params("plotBy")?.ToString()

            If String.IsNullOrEmpty(dataRange) Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("CreateChart: 缺少 dataRange 参数")
                Return False
            End If

            ' 解析数据范围
            Dim ws As Worksheet = Nothing
            Dim rangeAddress As String = ""
            If Not ParseExcelRange(dataRange, ws, rangeAddress) Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("CreateChart: 无法解析数据范围")
                Return False
            End If

            Dim sourceRange As Range = ws.Range(rangeAddress)

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
                Try
                    positionRange = ws.Range(position)
                Catch
                    positionRange = sourceRange.Offset(0, sourceRange.Columns.Count + 1)
                End Try
            Else
                positionRange = sourceRange.Offset(0, sourceRange.Columns.Count + 1)
            End If

            ' 创建图表
            Dim chartObj As ChartObject = ws.ChartObjects.Add(
                positionRange.Left,
                positionRange.Top,
                450,
                320)

            With chartObj.Chart
                .ChartType = xlChartType

                ' 设置数据源和绘图方向
                If Not String.IsNullOrEmpty(plotBy) AndAlso plotBy.ToLower() = "row" Then
                    .SetSourceData(sourceRange, XlRowCol.xlRows)
                Else
                    .SetSourceData(sourceRange, XlRowCol.xlColumns)
                End If

                ' 设置图表标题
                If Not String.IsNullOrEmpty(title) Then
                    .HasTitle = True
                    .ChartTitle.Text = title
                End If

                ' 设置系列名称（支持数组和范围两种格式）
                If seriesNamesParam IsNot Nothing Then
                    If seriesNamesParam.Type = JTokenType.Array Then
                        ' 数组格式
                        Dim names = seriesNamesParam.ToObject(Of List(Of String))()
                        For i As Integer = 1 To Math.Min(.SeriesCollection.Count, names.Count)
                            Try
                                .SeriesCollection(i).Name = names(i - 1)
                            Catch ex As Exception
                                Debug.WriteLine($"设置系列 {i} 名称失败: {ex.Message}")
                            End Try
                        Next
                    Else
                        ' 范围格式（如 "GDP_Sheet1!A2:A{lastRow}"）
                        Dim seriesNamesRangeStr = seriesNamesParam.ToString()
                        If Not String.IsNullOrEmpty(seriesNamesRangeStr) Then
                            Dim namesWs As Worksheet = Nothing
                            Dim namesAddress As String = ""
                            If ParseExcelRange(seriesNamesRangeStr, namesWs, namesAddress) Then
                                Try
                                    Dim namesRange As Range = namesWs.Range(namesAddress)
                                    For i As Integer = 1 To Math.Min(.SeriesCollection.Count, namesRange.Rows.Count)
                                        Try
                                            Dim nameVal = namesRange.Cells(i, 1).Value
                                            If nameVal IsNot Nothing Then
                                                .SeriesCollection(i).Name = nameVal.ToString()
                                            End If
                                        Catch ex As Exception
                                            Debug.WriteLine($"设置系列 {i} 名称失败: {ex.Message}")
                                        End Try
                                    Next
                                Catch ex As Exception
                                    Debug.WriteLine($"解析系列名称范围失败: {ex.Message}")
                                End Try
                            End If
                        End If
                    End If
                End If

                ' 设置分类轴标签（支持数组和范围两种格式）
                If categoryAxisParam IsNot Nothing Then
                    If categoryAxisParam.Type = JTokenType.Array Then
                        ' 数组格式
                        Dim labels = categoryAxisParam.ToObject(Of List(Of String))()
                        If .SeriesCollection.Count > 0 Then
                            Try
                                .SeriesCollection(1).XValues = labels.ToArray()
                            Catch ex As Exception
                                Debug.WriteLine($"设置分类轴标签失败: {ex.Message}")
                            End Try
                        End If
                    Else
                        ' 范围格式
                        Dim categoryAxisStr = categoryAxisParam.ToString()
                        If Not String.IsNullOrEmpty(categoryAxisStr) Then
                            Dim catWs As Worksheet = Nothing
                            Dim catAddress As String = ""
                            If ParseExcelRange(categoryAxisStr, catWs, catAddress) Then
                                Try
                                    Dim catRange As Range = catWs.Range(catAddress)
                                    If .SeriesCollection.Count > 0 Then
                                        Try
                                            .SeriesCollection(1).XValues = catRange
                                        Catch ex As Exception
                                            Debug.WriteLine($"设置分类轴标签失败: {ex.Message}")
                                        End Try
                                    End If
                                Catch ex As Exception
                                    Debug.WriteLine($"解析分类轴范围失败: {ex.Message}")
                                End Try
                            End If
                        End If
                    End If
                End If

                ' 设置图例
                .HasLegend = True
                If Not String.IsNullOrEmpty(legendPosition) Then
                    Select Case legendPosition.ToLower()
                        Case "right"
                            .Legend.Position = XlLegendPosition.xlLegendPositionRight
                        Case "left"
                            .Legend.Position = XlLegendPosition.xlLegendPositionLeft
                        Case "top"
                            .Legend.Position = XlLegendPosition.xlLegendPositionTop
                        Case "bottom"
                            .Legend.Position = XlLegendPosition.xlLegendPositionBottom
                        Case "corner"
                            .Legend.Position = XlLegendPosition.xlLegendPositionCorner
                    End Select
                End If
            End With

            ShareRibbon.GlobalStatusStrip.ShowInfo("图表已创建")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCreateChart 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"创建图表失败: {ex.Message}")
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

#Region "新增命令执行方法"

    ''' <summary>
    ''' 执行数据排序命令
    ''' </summary>
    Private Function ExecuteSortData(params As JToken) As Boolean
        Try
            Dim range = params("range")?.ToString()
            Dim sortColumn = params("sortColumn")?.Value(Of Integer)()
            Dim order = If(params("order")?.ToString()?.ToLower() = "desc", XlSortOrder.xlDescending, XlSortOrder.xlAscending)
            Dim hasHeader = If(params("hasHeader")?.Value(Of Boolean)(), True)

            If String.IsNullOrEmpty(range) OrElse Not sortColumn.HasValue Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim dataRange As Range = ws.Range(range)

            dataRange.Sort(
                Key1:=dataRange.Columns(sortColumn.Value),
                Order1:=order,
                Header:=If(hasHeader, XlYesNoGuess.xlYes, XlYesNoGuess.xlNo))

            ShareRibbon.GlobalStatusStrip.ShowInfo($"数据排序完成: {range}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteSortData 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"排序失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行数据筛选命令
    ''' </summary>
    Private Function ExecuteFilterData(params As JToken) As Boolean
        Try
            Dim ws As Worksheet = _excelApp.ActiveSheet

            ' 清除筛选
            Dim clearFilter = params("clearFilter")
            If clearFilter IsNot Nothing AndAlso clearFilter.Value(Of Boolean)() = True Then
                If ws.AutoFilterMode Then
                    ws.AutoFilterMode = False
                End If
                ShareRibbon.GlobalStatusStrip.ShowInfo("筛选已清除")
                Return True
            End If

            Dim range = params("range")?.ToString()
            Dim column = params("column")?.Value(Of Integer)()
            Dim criteria = params("criteria")?.ToString()

            If String.IsNullOrEmpty(range) OrElse Not column.HasValue Then
                Return False
            End If

            Dim dataRange As Range = ws.Range(range)

            dataRange.AutoFilter(Field:=column.Value, Criteria1:=criteria)

            ShareRibbon.GlobalStatusStrip.ShowInfo($"筛选已应用: 列{column} {criteria}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteFilterData 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"筛选失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行删除重复项命令
    ''' </summary>
    Private Function ExecuteRemoveDuplicates(params As JToken) As Boolean
        Try
            Dim range = params("range")?.ToString()
            Dim hasHeader = If(params("hasHeader")?.Value(Of Boolean)(), True)

            If String.IsNullOrEmpty(range) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim dataRange As Range = ws.Range(range)

            ' 默认检查所有列
            Dim columnsArray = params("columns")
            Dim cols As Object
            If columnsArray IsNot Nothing AndAlso columnsArray.Type = JTokenType.Array Then
                cols = columnsArray.ToObject(Of Integer())()
            Else
                ' 所有列
                Dim colCount = dataRange.Columns.Count
                Dim allCols(colCount - 1) As Integer
                For i = 0 To colCount - 1
                    allCols(i) = i + 1
                Next
                cols = allCols
            End If

            dataRange.RemoveDuplicates(Columns:=cols, Header:=If(hasHeader, XlYesNoGuess.xlYes, XlYesNoGuess.xlNo))

            ShareRibbon.GlobalStatusStrip.ShowInfo($"重复项已删除: {range}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteRemoveDuplicates 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"删除重复项失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行条件格式命令
    ''' </summary>
    Private Function ExecuteConditionalFormat(params As JToken) As Boolean
        Try
            Dim range = params("range")?.ToString()
            Dim rule = params("rule")?.ToString()?.ToLower()
            Dim condition = params("condition")?.ToString()
            Dim color = params("color")?.ToString()

            If String.IsNullOrEmpty(range) OrElse String.IsNullOrEmpty(rule) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim dataRange As Range = ws.Range(range)

            Select Case rule
                Case "highlight"
                    ' 突出显示单元格规则
                    Dim fc = dataRange.FormatConditions.Add(
                        Type:=XlFormatConditionType.xlCellValue,
                        Operator:=XlFormatConditionOperator.xlGreater,
                        Formula1:=If(String.IsNullOrEmpty(condition), "0", condition))
                    fc.Interior.Color = If(String.IsNullOrEmpty(color), RGB(255, 199, 206), ParseColor(color))

                Case "databar"
                    ' 数据条
                    dataRange.FormatConditions.AddDatabar()

                Case "colorscale"
                    ' 色阶
                    dataRange.FormatConditions.AddColorScale(ColorScaleType:=3)

                Case "iconset"
                    ' 图标集
                    dataRange.FormatConditions.AddIconSetCondition()
            End Select

            ShareRibbon.GlobalStatusStrip.ShowInfo($"条件格式已应用: {range}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteConditionalFormat 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"条件格式失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行合并单元格命令
    ''' </summary>
    Private Function ExecuteMergeCells(params As JToken) As Boolean
        Try
            Dim range = params("range")?.ToString()
            Dim unmergeParam = params("unmerge")
            Dim unmerge As Boolean = If(unmergeParam IsNot Nothing, unmergeParam.Value(Of Boolean)(), False)

            If String.IsNullOrEmpty(range) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim dataRange As Range = ws.Range(range)

            If unmerge Then
                dataRange.UnMerge()
                ShareRibbon.GlobalStatusStrip.ShowInfo($"已取消合并: {range}")
            Else
                dataRange.Merge()
                ShareRibbon.GlobalStatusStrip.ShowInfo($"已合并单元格: {range}")
            End If

            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteMergeCells 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"合并单元格失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行自动调整列宽/行高命令
    ''' </summary>
    Private Function ExecuteAutoFit(params As JToken) As Boolean
        Try
            Dim range = params("range")?.ToString()
            Dim fitType = If(params("type")?.ToString()?.ToLower(), "columns")

            If String.IsNullOrEmpty(range) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim dataRange As Range = ws.Range(range)

            Select Case fitType
                Case "columns"
                    dataRange.Columns.AutoFit()
                Case "rows"
                    dataRange.Rows.AutoFit()
                Case "both"
                    dataRange.Columns.AutoFit()
                    dataRange.Rows.AutoFit()
            End Select

            ShareRibbon.GlobalStatusStrip.ShowInfo($"自动调整完成: {range}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteAutoFit 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"自动调整失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行查找替换命令
    ''' </summary>
    Private Function ExecuteFindReplace(params As JToken) As Boolean
        Try
            Dim range = params("range")?.ToString()
            Dim findText = params("find")?.ToString()
            Dim replaceText = params("replace")?.ToString()
            Dim matchCase = If(params("matchCase")?.Value(Of Boolean)(), False)
            Dim matchEntireCell = If(params("matchEntireCell")?.Value(Of Boolean)(), False)

            If String.IsNullOrEmpty(findText) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim searchRange As Range

            If String.IsNullOrEmpty(range) OrElse range.ToLower() = "all" Then
                searchRange = ws.UsedRange
            Else
                searchRange = ws.Range(range)
            End If

            searchRange.Replace(
                What:=findText,
                Replacement:=If(replaceText, ""),
                LookAt:=If(matchEntireCell, XlLookAt.xlWhole, XlLookAt.xlPart),
                MatchCase:=matchCase)

            ShareRibbon.GlobalStatusStrip.ShowInfo($"查找替换完成: {findText} -> {replaceText}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteFindReplace 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"查找替换失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行创建数据透视表命令
    ''' </summary>
    Private Function ExecuteCreatePivotTable(params As JToken) As Boolean
        Try
            Dim sourceRange = params("sourceRange")?.ToString()
            Dim targetCell = params("targetCell")?.ToString()
            Dim rowFields = params("rowFields")
            Dim valueFields = params("valueFields")
            Dim columnFields = params("columnFields")

            If String.IsNullOrEmpty(sourceRange) OrElse String.IsNullOrEmpty(targetCell) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet
            Dim source As Range = ws.Range(sourceRange)

            ' 创建新工作表放置透视表
            Dim pivotWs As Worksheet = _excelApp.Worksheets.Add()
            pivotWs.Name = "透视表_" & DateTime.Now.ToString("HHmmss")

            Dim pivotCache As PivotCache = _excelApp.ActiveWorkbook.PivotCaches.Create(
                SourceType:=XlPivotTableSourceType.xlDatabase,
                SourceData:=source)

            Dim pivotTable As PivotTable = pivotCache.CreatePivotTable(
                TableDestination:=pivotWs.Range("A3"),
                TableName:="PivotTable1")

            ' 添加行字段
            If rowFields IsNot Nothing Then
                For Each field In rowFields
                    Dim fieldName = field.ToString()
                    pivotTable.PivotFields(fieldName).Orientation = XlPivotFieldOrientation.xlRowField
                Next
            End If

            ' 添加值字段
            If valueFields IsNot Nothing Then
                For Each field In valueFields
                    Dim fieldName = field.ToString()
                    pivotTable.AddDataField(pivotTable.PivotFields(fieldName), , XlConsolidationFunction.xlSum)
                Next
            End If

            ' 添加列字段
            If columnFields IsNot Nothing Then
                For Each field In columnFields
                    Dim fieldName = field.ToString()
                    pivotTable.PivotFields(fieldName).Orientation = XlPivotFieldOrientation.xlColumnField
                Next
            End If

            pivotWs.Activate()
            ShareRibbon.GlobalStatusStrip.ShowInfo($"透视表已创建: {pivotWs.Name}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCreatePivotTable 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"创建透视表失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行创建工作表命令
    ''' </summary>
    Private Function ExecuteCreateSheet(params As JToken) As Boolean
        Try
            Dim name = params("name")?.ToString()
            Dim position = params("position")?.ToString()?.ToLower()
            Dim referenceSheet = params("referenceSheet")?.ToString()

            If String.IsNullOrEmpty(name) Then
                Return False
            End If

            Dim newSheet As Worksheet

            If Not String.IsNullOrEmpty(referenceSheet) AndAlso Not String.IsNullOrEmpty(position) Then
                Dim refWs As Worksheet = _excelApp.Worksheets(referenceSheet)
                If position = "before" Then
                    newSheet = _excelApp.Worksheets.Add(Before:=refWs)
                Else
                    newSheet = _excelApp.Worksheets.Add(After:=refWs)
                End If
            Else
                newSheet = _excelApp.Worksheets.Add()
            End If

            newSheet.Name = name

            ShareRibbon.GlobalStatusStrip.ShowInfo($"工作表已创建: {name}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCreateSheet 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"创建工作表失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行删除工作表命令
    ''' </summary>
    Private Function ExecuteDeleteSheet(params As JToken) As Boolean
        Try
            Dim name = params("name")?.ToString()

            If String.IsNullOrEmpty(name) Then
                Return False
            End If

            _excelApp.DisplayAlerts = False
            _excelApp.Worksheets(name).Delete()
            _excelApp.DisplayAlerts = True

            ShareRibbon.GlobalStatusStrip.ShowInfo($"工作表已删除: {name}")
            Return True

        Catch ex As Exception
            _excelApp.DisplayAlerts = True
            Debug.WriteLine($"ExecuteDeleteSheet 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"删除工作表失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行重命名工作表命令
    ''' </summary>
    Private Function ExecuteRenameSheet(params As JToken) As Boolean
        Try
            Dim oldName = params("oldName")?.ToString()
            Dim newName = params("newName")?.ToString()

            If String.IsNullOrEmpty(oldName) OrElse String.IsNullOrEmpty(newName) Then
                Return False
            End If

            _excelApp.Worksheets(oldName).Name = newName

            ShareRibbon.GlobalStatusStrip.ShowInfo($"工作表已重命名: {oldName} -> {newName}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteRenameSheet 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"重命名工作表失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行复制工作表命令
    ''' </summary>
    Private Function ExecuteCopySheet(params As JToken) As Boolean
        Try
            Dim sourceName = params("sourceName")?.ToString()
            Dim newName = params("newName")?.ToString()

            If String.IsNullOrEmpty(sourceName) OrElse String.IsNullOrEmpty(newName) Then
                Return False
            End If

            Dim sourceWs As Worksheet = _excelApp.Worksheets(sourceName)
            sourceWs.Copy(After:=sourceWs)

            ' 复制后的工作表是活动工作表
            Dim newWs As Worksheet = _excelApp.ActiveSheet
            newWs.Name = newName

            ShareRibbon.GlobalStatusStrip.ShowInfo($"工作表已复制: {sourceName} -> {newName}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteCopySheet 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"复制工作表失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行插入行/列命令
    ''' </summary>
    Private Function ExecuteInsertRowCol(params As JToken) As Boolean
        Try
            Dim type = params("type")?.ToString()?.ToLower()
            Dim position = params("position")?.ToString()
            Dim count = If(params("count")?.Value(Of Integer)(), 1)

            If String.IsNullOrEmpty(type) OrElse String.IsNullOrEmpty(position) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet

            For i = 1 To count
                If type = "row" Then
                    Dim rowNum = Integer.Parse(position)
                    ws.Rows(rowNum).Insert(Shift:=XlInsertShiftDirection.xlShiftDown)
                Else
                    ws.Columns(position).Insert(Shift:=XlInsertShiftDirection.xlShiftToRight)
                End If
            Next

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已插入 {count} {If(type = "row", "行", "列")}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertRowCol 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"插入失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行删除行/列命令
    ''' </summary>
    Private Function ExecuteDeleteRowCol(params As JToken) As Boolean
        Try
            Dim type = params("type")?.ToString()?.ToLower()
            Dim position = params("position")?.ToString()
            Dim count = If(params("count")?.Value(Of Integer)(), 1)

            If String.IsNullOrEmpty(type) OrElse String.IsNullOrEmpty(position) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet

            If type = "row" Then
                Dim rowNum = Integer.Parse(position)
                ws.Rows($"{rowNum}:{rowNum + count - 1}").Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
            Else
                ' 计算列范围
                ws.Columns($"{position}:{GetColumnOffset(position, count - 1)}").Delete(Shift:=XlDeleteShiftDirection.xlShiftToLeft)
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已删除 {count} {If(type = "row", "行", "列")}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteDeleteRowCol 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"删除失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行隐藏/显示行/列命令
    ''' </summary>
    Private Function ExecuteHideRowCol(params As JToken) As Boolean
        Try
            Dim type = params("type")?.ToString()?.ToLower()
            Dim position = params("position")?.ToString()
            Dim unhide = If(params("unhide")?.Value(Of Boolean)(), False)

            If String.IsNullOrEmpty(type) OrElse String.IsNullOrEmpty(position) Then
                Return False
            End If

            Dim ws As Worksheet = _excelApp.ActiveSheet

            If type = "row" Then
                Dim rowNum = Integer.Parse(position)
                ws.Rows(rowNum).Hidden = Not unhide
            Else
                ws.Columns(position).Hidden = Not unhide
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已{If(unhide, "显示", "隐藏")}{If(type = "row", "行", "列")}: {position}")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteHideRowCol 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"隐藏/显示失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行保护工作表命令
    ''' </summary>
    Private Function ExecuteProtectSheet(params As JToken) As Boolean
        Try
            Dim sheetName = params("sheetName")?.ToString()
            Dim password = params("password")?.ToString()
            Dim unprotect = If(params("unprotect")?.Value(Of Boolean)(), False)

            Dim ws As Worksheet
            If String.IsNullOrEmpty(sheetName) Then
                ws = _excelApp.ActiveSheet
            Else
                ws = _excelApp.Worksheets(sheetName)
            End If

            If unprotect Then
                ws.Unprotect(Password:=password)
                ShareRibbon.GlobalStatusStrip.ShowInfo($"工作表保护已取消: {ws.Name}")
            Else
                ws.Protect(Password:=password)
                ShareRibbon.GlobalStatusStrip.ShowInfo($"工作表已保护: {ws.Name}")
            End If

            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteProtectSheet 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"保护工作表失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行VBA代码命令（自动回退机制）
    ''' </summary>
    Private Function ExecuteVBA(params As JToken) As Boolean
        Try
            Dim code = params("code")?.ToString()

            If String.IsNullOrEmpty(code) Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("ExecuteVBA: 缺少code参数")
                Return False
            End If

            ' 处理转义字符
            code = code.Replace("\n", vbCrLf).Replace("\t", vbTab).Replace("\""", """")

            ' 获取VBProject
            Dim vbProj As Microsoft.Vbe.Interop.VBProject = Nothing
            Try
                vbProj = _excelApp.VBE.ActiveVBProject
            Catch ex As Exception
                ShareRibbon.GlobalStatusStrip.ShowWarning("无法访问VBA项目，请在信任中心设置中启用'信任对VBA项目对象模型的访问'")
                Return False
            End Try

            If vbProj Is Nothing Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("无法获取VBA项目")
                Return False
            End If

            Dim vbComp As Microsoft.Vbe.Interop.VBComponent = Nothing
            Dim tempModuleName As String = "TempMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

            Try
                ' 创建临时模块
                vbComp = vbProj.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule)
                vbComp.Name = tempModuleName

                ' 检查代码是否已包含Sub/Function定义
                Dim hasProcedure = System.Text.RegularExpressions.Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", System.Text.RegularExpressions.RegexOptions.Multiline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                If hasProcedure Then
                    vbComp.CodeModule.AddFromString(code)
                    ' 查找第一个过程名
                    Dim procMatch = System.Text.RegularExpressions.Regex.Match(code, "^\s*(Sub|Function)\s+(\w+)", System.Text.RegularExpressions.RegexOptions.Multiline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    If procMatch.Success Then
                        Dim procName = procMatch.Groups(2).Value
                        _excelApp.Run(tempModuleName & "." & procName)
                    End If
                Else
                    ' 包装为Sub
                    Dim wrappedCode = "Sub Auto_Run()" & vbCrLf & code & vbCrLf & "End Sub"
                    vbComp.CodeModule.AddFromString(wrappedCode)
                    _excelApp.Run(tempModuleName & ".Auto_Run")
                End If

                ShareRibbon.GlobalStatusStrip.ShowInfo("VBA代码执行成功")
                Return True

            Catch ex As Exception
                ShareRibbon.GlobalStatusStrip.ShowWarning($"VBA执行失败: {ex.Message}")
                Return False
            Finally
                ' 删除临时模块
                Try
                    If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                        vbProj.VBComponents.Remove(vbComp)
                    End If
                Catch
                End Try
            End Try

        Catch ex As Exception
            Debug.WriteLine($"ExecuteVBA 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"VBA执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取列偏移后的列字母
    ''' </summary>
    Private Function GetColumnOffset(column As String, offset As Integer) As String
        Dim colNum = 0
        For Each c In column.ToUpper()
            colNum = colNum * 26 + (Asc(c) - Asc("A"c) + 1)
        Next
        colNum += offset

        Dim result = ""
        While colNum > 0
            colNum -= 1
            result = Chr(Asc("A"c) + (colNum Mod 26)) & result
            colNum \= 26
        End While
        Return result
    End Function

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 解析Excel范围（支持工作表名和{lastRow}占位符）
    ''' </summary>
    ''' <param name="rangeStr">范围字符串</param>
    ''' <param name="ws">输出：工作表对象</param>
    ''' <param name="rangeAddress">输出：范围地址</param>
    ''' <returns>是否成功</returns>
    Private Function ParseExcelRange(rangeStr As String, ByRef ws As Worksheet, ByRef rangeAddress As String) As Boolean
        Try
            If String.IsNullOrEmpty(rangeStr) Then
                Return False
            End If

            ws = _excelApp.ActiveSheet
            rangeAddress = rangeStr

            ' 解析工作表名
            If rangeStr.Contains("!") Then
                Dim parts = rangeStr.Split("!"c)
                Dim sheetName = parts(0).Trim("'"c)
                rangeAddress = parts(1)

                Try
                    ws = _excelApp.Worksheets(sheetName)
                Catch
                    ' 工作表不存在，尝试使用活动工作表
                    ws = _excelApp.ActiveSheet
                End Try
            End If

            ' 替换 {lastRow} 占位符
            If rangeAddress.Contains("{lastRow}") Then
                Dim usedRange = ws.UsedRange
                Dim lastRow = usedRange.Row + usedRange.Rows.Count - 1
                rangeAddress = rangeAddress.Replace("{lastRow}", lastRow.ToString())
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ParseExcelRange 出错: {ex.Message}")
            Return False
        End Try
    End Function

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
