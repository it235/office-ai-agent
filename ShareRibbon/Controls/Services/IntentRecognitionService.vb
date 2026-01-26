' ShareRibbon\Controls\Services\IntentRecognitionService.vb
' 意图识别服务：分析用户输入并识别操作意图

Imports System.Diagnostics
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq

''' <summary>
''' Excel操作意图类型枚举
''' </summary>
Public Enum ExcelIntentType
    DATA_ANALYSIS       ' 数据分析（统计、汇总、透视表）
    FORMULA_CALC        ' 公式计算
    CHART_GEN           ' 图表生成
    DATA_CLEANING       ' 数据清洗（去重、填充、格式化）
    REPORT_GEN          ' 报表生成
    DATA_TRANSFORMATION ' 数据转换（合并、拆分、转置）
    FORMAT_STYLE        ' 格式样式调整
    GENERAL_QUERY       ' 一般查询
End Enum

''' <summary>
''' 意图识别结果
''' </summary>
Public Class IntentResult
    ''' <summary>
    ''' 主要意图类型
    ''' </summary>
    Public Property IntentType As ExcelIntentType = ExcelIntentType.GENERAL_QUERY

    ''' <summary>
    ''' 次要意图（可能有多个操作）
    ''' </summary>
    Public Property SecondaryIntents As List(Of ExcelIntentType) = New List(Of ExcelIntentType)()

    ''' <summary>
    ''' 意图置信度 (0-1)
    ''' </summary>
    Public Property Confidence As Double = 0.5

    ''' <summary>
    ''' 响应模式
    ''' </summary>
    Public Property ResponseMode As String = ""

    ''' <summary>
    ''' 是否需要VBA代码
    ''' </summary>
    Public Property RequiresVBA As Boolean = True

    ''' <summary>
    ''' 是否可以使用直接操作命令
    ''' </summary>
    Public Property CanUseDirectCommand As Boolean = False

    ''' <summary>
    ''' 提取的关键实体（如范围、列名等）
    ''' </summary>
    Public Property ExtractedEntities As Dictionary(Of String, String) = New Dictionary(Of String, String)()
End Class

''' <summary>
''' 意图识别服务
''' </summary>
Public Class IntentRecognitionService

#Region "关键词映射"

    ' 数据分析关键词
    Private Shared ReadOnly DataAnalysisKeywords As String() = {
        "统计", "分析", "汇总", "求和", "平均", "最大", "最小", "计数",
        "透视表", "数据透视", "分组", "聚合", "占比", "百分比", "增长率",
        "趋势", "对比", "排名", "top", "前几", "后几"
    }

    ' 公式计算关键词
    Private Shared ReadOnly FormulaCalcKeywords As String() = {
        "公式", "计算", "求", "加", "减", "乘", "除", "sum", "average",
        "vlookup", "if", "countif", "sumif", "index", "match"
    }

    ' 图表生成关键词
    Private Shared ReadOnly ChartGenKeywords As String() = {
        "图表", "柱状图", "折线图", "饼图", "条形图", "散点图", "面积图",
        "chart", "graph", "可视化", "画图", "生成图", "做个图"
    }

    ' 数据清洗关键词
    Private Shared ReadOnly DataCleaningKeywords As String() = {
        "清洗", "去重", "删除重复", "填充", "空值", "缺失", "替换",
        "格式化", "规范", "trim", "清理", "整理", "修复"
    }

    ' 报表生成关键词
    Private Shared ReadOnly ReportGenKeywords As String() = {
        "报表", "报告", "表格", "生成表", "导出", "输出", "创建表",
        "周报", "月报", "日报", "汇报", "模板"
    }

    ' 数据转换关键词
    Private Shared ReadOnly DataTransformKeywords As String() = {
        "合并", "拆分", "转置", "行列转换", "连接", "vlookup", "关联",
        "join", "merge", "split", "transpose", "提取", "截取"
    }

    ' 格式样式关键词
    Private Shared ReadOnly FormatStyleKeywords As String() = {
        "格式", "样式", "颜色", "字体", "边框", "对齐", "加粗",
        "斜体", "底色", "高亮", "条件格式", "美化"
    }

#End Region

#Region "公共方法"

    ''' <summary>
    ''' 识别用户意图
    ''' </summary>
    ''' <param name="question">用户问题</param>
    ''' <param name="context">上下文信息（可选）</param>
    ''' <returns>意图识别结果</returns>
    Public Function IdentifyIntent(question As String, Optional context As JObject = Nothing) As IntentResult
        Dim result As New IntentResult()

        If String.IsNullOrWhiteSpace(question) Then
            Return result
        End If

        Dim lowerQuestion = question.ToLower()

        ' 计算各意图的匹配分数
        Dim scores As New Dictionary(Of ExcelIntentType, Double)()
        scores(ExcelIntentType.DATA_ANALYSIS) = CalculateKeywordScore(lowerQuestion, DataAnalysisKeywords)
        scores(ExcelIntentType.FORMULA_CALC) = CalculateKeywordScore(lowerQuestion, FormulaCalcKeywords)
        scores(ExcelIntentType.CHART_GEN) = CalculateKeywordScore(lowerQuestion, ChartGenKeywords)
        scores(ExcelIntentType.DATA_CLEANING) = CalculateKeywordScore(lowerQuestion, DataCleaningKeywords)
        scores(ExcelIntentType.REPORT_GEN) = CalculateKeywordScore(lowerQuestion, ReportGenKeywords)
        scores(ExcelIntentType.DATA_TRANSFORMATION) = CalculateKeywordScore(lowerQuestion, DataTransformKeywords)
        scores(ExcelIntentType.FORMAT_STYLE) = CalculateKeywordScore(lowerQuestion, FormatStyleKeywords)

        ' 找出最高分的意图
        Dim maxScore As Double = 0
        Dim maxIntent = ExcelIntentType.GENERAL_QUERY

        For Each kvp In scores
            If kvp.Value > maxScore Then
                maxScore = kvp.Value
                maxIntent = kvp.Key
            End If
        Next

        ' 设置主要意图
        If maxScore > 0.1 Then
            result.IntentType = maxIntent
            result.Confidence = Math.Min(maxScore, 1.0)
        End If

        ' 查找次要意图（分数超过0.05的其他意图）
        For Each kvp In scores
            If kvp.Key <> maxIntent AndAlso kvp.Value > 0.05 Then
                result.SecondaryIntents.Add(kvp.Key)
            End If
        Next

        ' 提取关键实体
        ExtractEntities(question, result)

        ' 判断是否可以使用直接命令
        DetermineExecutionMethod(result)

        Debug.WriteLine($"意图识别结果: {result.IntentType}, 置信度: {result.Confidence:F2}")
        Return result
    End Function

    ''' <summary>
    ''' 异步识别意图（支持LLM增强）
    ''' </summary>
    Public Async Function IdentifyIntentAsync(question As String, Optional context As JObject = Nothing) As Task(Of IntentResult)
        ' 首先使用关键词匹配快速识别
        Dim result = IdentifyIntent(question, context)

        ' 如果置信度较低，可以考虑调用LLM增强识别
        ' 目前直接返回关键词匹配结果，后续可扩展
        Return result
    End Function

    ''' <summary>
    ''' 获取优化后的系统提示词
    ''' </summary>
    Public Function GetOptimizedSystemPrompt(intent As IntentResult) As String
        Dim sb As New StringBuilder()

        ' 根据意图类型选择专用提示词
        Select Case intent.IntentType
            Case ExcelIntentType.DATA_ANALYSIS
                sb.AppendLine(GetDataAnalysisPrompt())
            Case ExcelIntentType.FORMULA_CALC
                sb.AppendLine(GetFormulaCalcPrompt())
            Case ExcelIntentType.CHART_GEN
                sb.AppendLine(GetChartGenPrompt())
            Case ExcelIntentType.DATA_CLEANING
                sb.AppendLine(GetDataCleaningPrompt())
            Case ExcelIntentType.REPORT_GEN
                sb.AppendLine(GetReportGenPrompt())
            Case ExcelIntentType.DATA_TRANSFORMATION
                sb.AppendLine(GetDataTransformPrompt())
            Case ExcelIntentType.FORMAT_STYLE
                sb.AppendLine(GetFormatStylePrompt())
            Case Else
                sb.AppendLine(GetGeneralPrompt())
        End Select

        ' 添加通用指导
        sb.AppendLine()
        sb.AppendLine("【通用要求】")
        sb.AppendLine("1. 必须返回严格有效的JSON格式")
        sb.AppendLine("2. 动态范围使用占位符 {lastRow} 而非JS表达式")
        sb.AppendLine("3. command必须是: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData")

        Return sb.ToString()
    End Function

#End Region

#Region "提示词模板"

    Private Function GetDataAnalysisPrompt() As String
        Return "你是Excel数据分析专家。必须返回严格的JSON命令。

```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""C1:C{lastRow}"", ""formula"": ""=A1+B1"", ""fillDown"": true}}
```
占位符: {lastRow}=最后一行, {lastCol}=最后一列"
    End Function

    Private Function GetFormulaCalcPrompt() As String
        Return "你是Excel公式专家。必须返回严格的JSON命令。

```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""C1:C{lastRow}"", ""formula"": ""=A1+B1"", ""fillDown"": true}}
```
占位符: {lastRow}, {lastCol}"
    End Function

    Private Function GetChartGenPrompt() As String
        Return "你是Excel图表专家。必须返回严格的JSON命令。

```json
{""command"": ""CreateChart"", ""params"": {""type"": ""Column"", ""dataRange"": ""A1:B{lastRow}"", ""title"": ""标题"", ""position"": ""E1""}}
```
type可选: Column, Line, Pie, Bar"
    End Function

    Private Function GetDataCleaningPrompt() As String
        Return "你是Excel数据清洗专家。必须返回严格的JSON命令。

```json
{""command"": ""CleanData"", ""params"": {""operation"": ""removeDuplicates"", ""range"": ""A1:D{lastRow}""}}
```
operation可选: removeDuplicates, fillEmpty, trim"
    End Function

    Private Function GetReportGenPrompt() As String
        Return "你是Excel报表专家。必须返回严格的JSON命令。

```json
{""command"": ""FormatRange"", ""params"": {""range"": ""A1:D{lastRow}"", ""style"": ""Header"", ""bold"": true}}
```"
    End Function

    Private Function GetDataTransformPrompt() As String
        Return "你是Excel数据转换专家。必须返回严格的JSON命令。

```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""D1:D{lastRow}"", ""formula"": ""=VLOOKUP(A1,Sheet2!A:B,2,FALSE)"", ""fillDown"": true}}
```"
    End Function

    Private Function GetFormatStylePrompt() As String
        Return "你是Excel格式专家。必须返回严格的JSON命令。

```json
{""command"": ""FormatRange"", ""params"": {""range"": ""A1:D10"", ""style"": ""Header"", ""bold"": true, ""borders"": true}}
```"
    End Function

    Private Function GetGeneralPrompt() As String
        Return "你是Excel助手。操作Excel必须返回JSON命令，问答则直接回答。

```json
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""C1:C{lastRow}"", ""formula"": ""=A1+B1"", ""fillDown"": true}}
```
command必须是: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData"
    End Function

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 计算关键词匹配分数
    ''' </summary>
    Private Function CalculateKeywordScore(text As String, keywords As String()) As Double
        Dim matchCount As Integer = 0
        Dim totalWeight As Double = 0

        For Each keyword In keywords
            If text.Contains(keyword.ToLower()) Then
                matchCount += 1
                ' 关键词越长，权重越高
                totalWeight += keyword.Length / 10.0
            End If
        Next

        ' 归一化分数
        If keywords.Length > 0 Then
            Return (matchCount / keywords.Length * 0.5) + (totalWeight / keywords.Length * 0.5)
        End If

        Return 0
    End Function

    ''' <summary>
    ''' 提取关键实体
    ''' </summary>
    Private Sub ExtractEntities(question As String, result As IntentResult)
        ' 提取单元格范围 (如 A1:B10, A1, Sheet1!A1:B10)
        Dim rangePattern As New Regex("([A-Za-z]+\d+)(:[A-Za-z]+\d+)?", RegexOptions.IgnoreCase)
        Dim rangeMatch = rangePattern.Match(question)
        If rangeMatch.Success Then
            result.ExtractedEntities("range") = rangeMatch.Value
        End If

        ' 提取列名 (如 A列, B列)
        Dim columnPattern As New Regex("([A-Za-z])列", RegexOptions.IgnoreCase)
        Dim columnMatch = columnPattern.Match(question)
        If columnMatch.Success Then
            result.ExtractedEntities("column") = columnMatch.Groups(1).Value.ToUpper()
        End If

        ' 提取工作表名 (如 Sheet1, 工作表1)
        Dim sheetPattern As New Regex("(Sheet\d+|工作表\d+)", RegexOptions.IgnoreCase)
        Dim sheetMatch = sheetPattern.Match(question)
        If sheetMatch.Success Then
            result.ExtractedEntities("sheet") = sheetMatch.Value
        End If

        ' 提取数字 (可能是行数、数量等)
        Dim numberPattern As New Regex("\b(\d+)\b")
        Dim numberMatch = numberPattern.Match(question)
        If numberMatch.Success Then
            result.ExtractedEntities("number") = numberMatch.Value
        End If
    End Sub

    ''' <summary>
    ''' 判断执行方式
    ''' </summary>
    Private Sub DetermineExecutionMethod(result As IntentResult)
        ' 以下意图可以使用直接命令
        Dim directCommandIntents = {
            ExcelIntentType.FORMULA_CALC,
            ExcelIntentType.FORMAT_STYLE,
            ExcelIntentType.DATA_CLEANING,
            ExcelIntentType.CHART_GEN
        }

        result.CanUseDirectCommand = directCommandIntents.Contains(result.IntentType)

        ' 复杂操作仍需要VBA
        result.RequiresVBA = Not result.CanUseDirectCommand OrElse
                            result.SecondaryIntents.Count > 1 OrElse
                            result.Confidence < 0.3
    End Sub

#End Region

End Class
