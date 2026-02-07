' ShareRibbon\Controls\Services\ExcelContextService.vb
' Excel上下文服务：处理Excel数据的格式化和摘要生成
' 注意：此服务处理Excel返回的1-based Object(,)数组数据

Imports System.Diagnostics
Imports System.Text

''' <summary>
''' 数据结构信息
''' </summary>
Public Class DataStructureInfo
    ''' <summary>
    ''' 是否有表头
    ''' </summary>
    Public Property HasHeader As Boolean = False

    ''' <summary>
    ''' 表头行
    ''' </summary>
    Public Property HeaderRow As Integer = 1

    ''' <summary>
    ''' 列信息
    ''' </summary>
    Public Property Columns As List(Of ColumnInfo) = New List(Of ColumnInfo)()

    ''' <summary>
    ''' 行数
    ''' </summary>
    Public Property RowCount As Integer = 0

    ''' <summary>
    ''' 列数
    ''' </summary>
    Public Property ColumnCount As Integer = 0
End Class

''' <summary>
''' 列信息
''' </summary>
Public Class ColumnInfo
    ''' <summary>
    ''' 列索引
    ''' </summary>
    Public Property Index As Integer

    ''' <summary>
    ''' 列名（表头）
    ''' </summary>
    Public Property Name As String = ""

    ''' <summary>
    ''' 数据类型
    ''' </summary>
    Public Property DataType As String = "Unknown"

    ''' <summary>
    ''' 非空值计数
    ''' </summary>
    Public Property NonEmptyCount As Integer = 0
End Class

''' <summary>
''' Excel上下文服务（处理Excel返回的1-based数组）
''' </summary>
Public Class ExcelContextService

#Region "常量配置"

    ' Markdown表格显示的最大行数
    Private Const MARKDOWN_MAX_ROWS As Integer = 30

    ' 摘要模式的数据阈值（超过此数量使用摘要）
    Private Const SUMMARY_THRESHOLD As Integer = 100

#End Region

#Region "公共方法"

    ''' <summary>
    ''' 将数据转换为Markdown表格格式
    ''' </summary>
    ''' <param name="data">二维数据数组（Excel返回的1-based数组）</param>
    ''' <param name="hasHeader">是否包含表头</param>
    ''' <returns>Markdown格式的表格字符串</returns>
    Public Function FormatAsMarkdownTable(data As Object(,), Optional hasHeader As Boolean = True) As String
        If data Is Nothing OrElse data.Length = 0 Then
            Return "[无数据]"
        End If

        Dim sb As New StringBuilder()

        Try
            ' 获取数组的实际边界（Excel数组是1-based的）
            Dim rowStart = data.GetLowerBound(0)
            Dim rowEnd = data.GetUpperBound(0)
            Dim colStart = data.GetLowerBound(1)
            Dim colEnd = data.GetUpperBound(1)

            Dim totalRows = rowEnd - rowStart + 1
            Dim totalCols = colEnd - colStart + 1

            ' 限制显示行数
            Dim displayRowEnd = Math.Min(rowEnd, rowStart + MARKDOWN_MAX_ROWS - 1)

            ' 构建表头
            sb.Append("| ")
            For col = colStart To colEnd
                Dim cellValue = GetCellValueString(data(rowStart, col))
                sb.Append(cellValue)
                sb.Append(" | ")
            Next
            sb.AppendLine()

            ' 构建分隔行
            sb.Append("| ")
            For col = colStart To colEnd
                sb.Append("---")
                sb.Append(" | ")
            Next
            sb.AppendLine()

            ' 构建数据行
            Dim dataRowStart = If(hasHeader, rowStart + 1, rowStart)
            For row = dataRowStart To displayRowEnd
                sb.Append("| ")
                For col = colStart To colEnd
                    Dim cellValue = GetCellValueString(data(row, col))
                    sb.Append(cellValue)
                    sb.Append(" | ")
                Next
                sb.AppendLine()
            Next

            ' 如果有更多行未显示，添加提示
            If rowEnd > displayRowEnd Then
                sb.AppendLine($"... 共 {totalRows} 行，仅显示前 {displayRowEnd - rowStart + 1} 行")
            End If

        Catch ex As Exception
            Debug.WriteLine($"FormatAsMarkdownTable 出错: {ex.Message}")
            sb.AppendLine("[格式化数据时出错]")
        End Try

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 生成数据摘要（用于大数据集）
    ''' </summary>
    ''' <param name="data">二维数据数组</param>
    ''' <returns>数据摘要字符串</returns>
    Public Function GenerateDataSummary(data As Object(,)) As String
        If data Is Nothing OrElse data.Length = 0 Then
            Return "[无数据]"
        End If

        Dim sb As New StringBuilder()

        Try
            ' 获取数组边界
            Dim rowStart = data.GetLowerBound(0)
            Dim rowEnd = data.GetUpperBound(0)
            Dim colStart = data.GetLowerBound(1)
            Dim colEnd = data.GetUpperBound(1)

            Dim totalRows = rowEnd - rowStart + 1
            Dim totalCols = colEnd - colStart + 1

            sb.AppendLine("【数据摘要】")
            sb.AppendLine($"- 总行数: {totalRows}")
            sb.AppendLine($"- 总列数: {totalCols}")
            sb.AppendLine($"- 单元格总数: {totalRows * totalCols}")

            ' 分析每列的数据类型和统计信息
            sb.AppendLine()
            sb.AppendLine("【列信息】")

            Dim colIndex = 1
            For col = colStart To colEnd
                Dim columnName = GetCellValueString(data(rowStart, col))
                If String.IsNullOrEmpty(columnName) Then
                    columnName = GetExcelColumnName(colIndex)
                End If

                Dim dataType = DetectColumnDataType(data, col, rowStart, rowEnd)
                Dim stats = GetColumnStatistics(data, col, rowStart, rowEnd, dataType)

                sb.AppendLine($"- 列{colIndex} [{columnName}]: {dataType}")
                If Not String.IsNullOrEmpty(stats) Then
                    sb.AppendLine($"  {stats}")
                End If
                colIndex += 1
            Next

        Catch ex As Exception
            Debug.WriteLine($"GenerateDataSummary 出错: {ex.Message}")
        End Try

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 检测数据结构（基于数组）
    ''' </summary>
    ''' <param name="data">二维数据数组</param>
    ''' <returns>数据结构信息</returns>
    Public Function DetectDataStructure(data As Object(,)) As DataStructureInfo
        Dim info As New DataStructureInfo()

        Try
            If data Is Nothing OrElse data.Length = 0 Then
                Return info
            End If

            ' 获取数组边界
            Dim rowStart = data.GetLowerBound(0)
            Dim rowEnd = data.GetUpperBound(0)
            Dim colStart = data.GetLowerBound(1)
            Dim colEnd = data.GetUpperBound(1)

            info.RowCount = rowEnd - rowStart + 1
            info.ColumnCount = colEnd - colStart + 1

            ' 判断表头的简单规则：
            ' 1. 第一行全是文本
            ' 2. 第一行没有数字
            Dim allText = True
            Dim hasNumber = False

            For col = colStart To colEnd
                Dim value = data(rowStart, col)
                If value IsNot Nothing Then
                    If IsNumeric(value) Then
                        hasNumber = True
                        allText = False
                    ElseIf TypeOf value Is String Then
                        ' 保持allText为True
                    Else
                        allText = False
                    End If
                End If
            Next

            info.HasHeader = allText AndAlso Not hasNumber
            info.HeaderRow = If(info.HasHeader, 1, 0)

            ' 分析列信息
            Dim colIndex = 1
            For col = colStart To colEnd
                Dim colInfo As New ColumnInfo()
                colInfo.Index = colIndex

                If info.HasHeader AndAlso data(rowStart, col) IsNot Nothing Then
                    colInfo.Name = data(rowStart, col).ToString()
                Else
                    colInfo.Name = GetExcelColumnName(colIndex)
                End If

                info.Columns.Add(colInfo)
                colIndex += 1
            Next

        Catch ex As Exception
            Debug.WriteLine($"DetectDataStructure 出错: {ex.Message}")
        End Try

        Return info
    End Function

    ''' <summary>
    ''' 获取前N行数据
    ''' </summary>
    Public Function GetTopRows(data As Object(,), topN As Integer) As Object(,)
        If data Is Nothing OrElse data.Length = 0 Then
            Return Nothing
        End If

        Try
            ' 获取数组边界
            Dim rowStart = data.GetLowerBound(0)
            Dim rowEnd = data.GetUpperBound(0)
            Dim colStart = data.GetLowerBound(1)
            Dim colEnd = data.GetUpperBound(1)

            Dim totalRows = rowEnd - rowStart + 1
            Dim totalCols = colEnd - colStart + 1
            Dim actualRows = Math.Min(totalRows, topN)

            ' 创建1-based数组
            Dim result = DirectCast(Array.CreateInstance(GetType(Object), {actualRows, totalCols}, {1, 1}), Object(,))

            For row = 1 To actualRows
                For col = 1 To totalCols
                    result(row, col) = data(rowStart + row - 1, colStart + col - 1)
                Next
            Next

            Return result

        Catch ex As Exception
            Debug.WriteLine($"GetTopRows 出错: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 格式化选中内容为增强版上下文（使用预处理的数据）
    ''' </summary>
    Public Function FormatSelectionAsContext(
            data As Object(,),
            workbookName As String,
            worksheetName As String,
            rangeAddress As String) As String

        Dim sb As New StringBuilder()

        Try
            sb.AppendLine(vbCrLf & "--- 用户选中的Excel数据 ---")
            sb.AppendLine($"工作簿: {workbookName}")
            sb.AppendLine($"工作表: {worksheetName}")
            sb.AppendLine($"范围: {rangeAddress}")

            If data Is Nothing OrElse data.Length = 0 Then
                sb.AppendLine("[无数据]")
                sb.AppendLine("--- 数据引用结束 ---" & vbCrLf)
                Return sb.ToString()
            End If

            ' 检测数据结构
            Dim dataStructure = DetectDataStructure(data)

            Dim cellCount = dataStructure.RowCount * dataStructure.ColumnCount

            If cellCount <= SUMMARY_THRESHOLD Then
                ' 小数据集：完整Markdown表格
                sb.AppendLine()
                sb.AppendLine("【数据内容】")
                sb.AppendLine(FormatAsMarkdownTable(data, dataStructure.HasHeader))
            Else
                ' 大数据集：摘要 + 示例
                sb.AppendLine()
                sb.AppendLine(GenerateDataSummary(data))
                sb.AppendLine()
                sb.AppendLine("【前5行数据示例】")
                Dim topData = GetTopRows(data, 5)
                If topData IsNot Nothing Then
                    sb.AppendLine(FormatAsMarkdownTable(topData, dataStructure.HasHeader))
                End If
            End If

            sb.AppendLine("--- 数据引用结束 ---" & vbCrLf)

        Catch ex As Exception
            Debug.WriteLine($"FormatSelectionAsContext 出错: {ex.Message}")
            sb.AppendLine($"[读取数据时出错: {ex.Message}]")
        End Try

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 判断数据大小是否需要使用摘要模式
    ''' </summary>
    Public Function ShouldUseSummaryMode(cellCount As Integer) As Boolean
        Return cellCount > SUMMARY_THRESHOLD
    End Function

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 获取单元格值的字符串表示
    ''' </summary>
    Private Function GetCellValueString(value As Object) As String
        If value Is Nothing Then
            Return ""
        End If

        ' 处理日期
        If TypeOf value Is DateTime Then
            Return DirectCast(value, DateTime).ToString("yyyy-MM-dd")
        End If

        ' 处理数字（保留合理精度）
        If TypeOf value Is Double Then
            Dim d = DirectCast(value, Double)
            If d = Math.Floor(d) Then
                Return d.ToString("0")
            Else
                Return d.ToString("0.##")
            End If
        End If

        ' 处理布尔
        If TypeOf value Is Boolean Then
            Return If(DirectCast(value, Boolean), "是", "否")
        End If

        ' 其他类型直接转字符串
        Dim str = value.ToString()

        ' 转义Markdown特殊字符
        str = str.Replace("|", "\|")

        ' 限制单元格内容长度
        If str.Length > 50 Then
            str = str.Substring(0, 47) & "..."
        End If

        Return str
    End Function

    ''' <summary>
    ''' 检测列的数据类型
    ''' </summary>
    Private Function DetectColumnDataType(data As Object(,), col As Integer, rowStart As Integer, rowEnd As Integer) As String
        Dim numberCount = 0
        Dim textCount = 0
        Dim dateCount = 0
        Dim emptyCount = 0

        Dim startRow = rowStart + 1 ' 跳过表头
        Dim sampleEnd = Math.Min(rowEnd, rowStart + 19) ' 只检查前20行

        For row = startRow To sampleEnd
            Dim value = data(row, col)

            If value Is Nothing OrElse String.IsNullOrEmpty(value.ToString()) Then
                emptyCount += 1
            ElseIf TypeOf value Is Double OrElse IsNumeric(value) Then
                numberCount += 1
            ElseIf TypeOf value Is DateTime Then
                dateCount += 1
            Else
                textCount += 1
            End If
        Next

        ' 根据最多的类型判断
        If numberCount >= textCount AndAlso numberCount >= dateCount Then
            Return "数值"
        ElseIf dateCount >= textCount Then
            Return "日期"
        Else
            Return "文本"
        End If
    End Function

    ''' <summary>
    ''' 获取列统计信息
    ''' </summary>
    Private Function GetColumnStatistics(data As Object(,), col As Integer, rowStart As Integer, rowEnd As Integer, dataType As String) As String
        Try
            If dataType = "数值" Then
                Dim values As New List(Of Double)()

                For row = rowStart + 1 To rowEnd
                    Dim value = data(row, col)
                    If value IsNot Nothing AndAlso IsNumeric(value) Then
                        values.Add(CDbl(value))
                    End If
                Next

                If values.Count > 0 Then
                    Return $"最小:{values.Min():0.##}, 最大:{values.Max():0.##}, 平均:{values.Average():0.##}"
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"GetColumnStatistics 出错: {ex.Message}")
        End Try

        Return ""
    End Function

    ''' <summary>
    ''' 将列索引转换为Excel列名
    ''' </summary>
    Public Shared Function GetExcelColumnName(columnIndex As Integer) As String
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

#End Region

End Class
