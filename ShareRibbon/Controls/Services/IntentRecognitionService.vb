' ShareRibbon\Controls\Services\IntentRecognitionService.vb
' 意图识别服务：分析用户输入并识别操作意图

Imports System.Diagnostics
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
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

    ''' <summary>
    ''' 用户友好的意图描述
    ''' </summary>
    Public Property UserFriendlyDescription As String = ""

    ''' <summary>
    ''' 执行计划步骤列表
    ''' </summary>
    Public Property ExecutionPlan As List(Of ExecutionStep) = New List(Of ExecutionStep)()

    ''' <summary>
    ''' 原始用户输入
    ''' </summary>
    Public Property OriginalInput As String = ""
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
    ''' 异步识别意图（始终使用LLM进行置信度评分）
    ''' </summary>
    Public Async Function IdentifyIntentAsync(question As String, Optional context As JObject = Nothing) As Task(Of IntentResult)
        ' 首先使用关键词匹配进行初步分类（但不使用其置信度）
        Dim result = IdentifyIntent(question, context)

        ' 始终调用LLM进行置信度评分（用户要求置信度由大模型打分）
        If Not String.IsNullOrWhiteSpace(question) Then
            Try
                Dim llmResult = Await IdentifyIntentWithLLMAsync(question, context)
                If llmResult IsNot Nothing Then
                    ' 使用LLM的置信度（这是核心改动）
                    result.Confidence = llmResult.Confidence
                    
                    ' 如果LLM的意图类型判断更可信，也使用LLM的意图
                    If llmResult.Confidence > 0.3 Then
                        result.IntentType = llmResult.IntentType
                        result.UserFriendlyDescription = llmResult.UserFriendlyDescription
                    End If
                    
                    Debug.WriteLine($"LLM意图识别结果: {result.IntentType}, 置信度: {result.Confidence:F2}")
                End If
            Catch ex As Exception
                Debug.WriteLine($"LLM意图识别失败，使用默认置信度0.5: {ex.Message}")
                ' 如果LLM调用失败，使用默认中等置信度
                result.Confidence = 0.5
            End Try
        End If

        Return result
    End Function

    ''' <summary>
    ''' 调用大模型识别意图
    ''' </summary>
    Private Async Function IdentifyIntentWithLLMAsync(question As String, context As JObject) As Task(Of IntentResult)
        Dim result As New IntentResult()
        result.OriginalInput = question

        Try
            ' 获取API配置
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
                Return result
            End If

            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)

            Dim apiUrl = cfg.url
            Dim apiKey = cfg.key
            Dim modelName = selectedModel.modelName

            ' 构建上下文信息
            Dim contextInfo As String = ""
            If context IsNot Nothing Then
                If context("sheetName") IsNot Nothing Then
                    contextInfo &= $"当前工作表: {context("sheetName")}" & vbCrLf
                End If
                If context("selectionAddress") IsNot Nothing AndAlso Not String.IsNullOrEmpty(context("selectionAddress").ToString()) Then
                    contextInfo &= $"选中区域: {context("selectionAddress")}" & vbCrLf
                End If
                If context("selection") IsNot Nothing AndAlso Not String.IsNullOrEmpty(context("selection").ToString()) Then
                    contextInfo &= $"选中内容预览:" & vbCrLf & context("selection").ToString() & vbCrLf
                End If
            End If

            ' 构建意图识别提示词
            Dim systemPrompt = GetIntentRecognitionSystemPrompt()
            Dim userMessage = $"用户问题: {question}"
            If Not String.IsNullOrEmpty(contextInfo) Then
                userMessage &= vbCrLf & vbCrLf & "当前Office上下文信息:" & vbCrLf & contextInfo
            End If

            ' 构建请求体
            Dim messages As New JArray()
            messages.Add(New JObject From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject From {{"role", "user"}, {"content", userMessage}})

            Dim requestBody As New JObject()
            requestBody("model") = modelName
            requestBody("messages") = messages
            requestBody("temperature") = 0.3
            requestBody("max_tokens") = 500
            requestBody("stream") = False

            ' 发送请求
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(30)

                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                request.Content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request)
                    If response.IsSuccessStatusCode Then
                        Dim responseContent = Await response.Content.ReadAsStringAsync()
                        result = ParseLLMIntentResponse(responseContent, question)
                    End If
                End Using
            End Using

        Catch ex As Exception
            Debug.WriteLine($"IdentifyIntentWithLLMAsync 出错: {ex.Message}")
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 获取意图识别系统提示词
    ''' </summary>
    Private Function GetIntentRecognitionSystemPrompt() As String
        Return "你是一个Office意图识别助手。分析用户的问题和上下文，识别用户想要执行的操作。

请用JSON格式返回识别结果：
```json
{
  ""intentType"": ""DATA_ANALYSIS"",
  ""confidence"": 0.85,
  ""description"": ""用户想要对数据进行统计分析"",
  ""requiresConfirmation"": false,
  ""suggestedAction"": ""直接执行数据分析""
}
```

intentType必须是以下之一:
- DATA_ANALYSIS: 数据分析（统计、汇总、透视表）
- FORMULA_CALC: 公式计算
- CHART_GEN: 图表生成
- DATA_CLEANING: 数据清洗（去重、填充）
- REPORT_GEN: 报表生成
- DATA_TRANSFORMATION: 数据转换（合并、拆分）
- FORMAT_STYLE: 格式样式调整
- GENERAL_QUERY: 一般问答（不需要操作Excel）
- UNCLEAR: 意图不明确，需要进一步询问

confidence范围0-1，表示你对识别结果的确信程度。
requiresConfirmation: 如果意图明确且操作安全，设为false；如果需要用户确认，设为true。

注意：
1. 如果用户只是打招呼或闲聊，intentType设为GENERAL_QUERY，confidence设为0.9
2. 如果用户的请求涉及数据修改但表述不清，requiresConfirmation设为true
3. 结合Office上下文信息（如选中内容）来更准确地判断意图"
    End Function

    ''' <summary>
    ''' 解析LLM返回的意图识别结果
    ''' </summary>
    Private Function ParseLLMIntentResponse(responseContent As String, originalQuestion As String) As IntentResult
        Dim result As New IntentResult()
        result.OriginalInput = originalQuestion

        Try
            Dim responseJson = JObject.Parse(responseContent)
            Dim choices = responseJson("choices")
            If choices Is Nothing OrElse choices.Count = 0 Then Return result

            Dim content = choices(0)("message")?("content")?.ToString()
            If String.IsNullOrEmpty(content) Then Return result

            ' 提取JSON部分
            Dim jsonMatch = Regex.Match(content, "\{[\s\S]*\}")
            If Not jsonMatch.Success Then Return result

            Dim intentJson = JObject.Parse(jsonMatch.Value)

            ' 解析意图类型
            Dim intentTypeStr = intentJson("intentType")?.ToString()?.ToUpper()
            Select Case intentTypeStr
                Case "DATA_ANALYSIS"
                    result.IntentType = ExcelIntentType.DATA_ANALYSIS
                Case "FORMULA_CALC"
                    result.IntentType = ExcelIntentType.FORMULA_CALC
                Case "CHART_GEN"
                    result.IntentType = ExcelIntentType.CHART_GEN
                Case "DATA_CLEANING"
                    result.IntentType = ExcelIntentType.DATA_CLEANING
                Case "REPORT_GEN"
                    result.IntentType = ExcelIntentType.REPORT_GEN
                Case "DATA_TRANSFORMATION"
                    result.IntentType = ExcelIntentType.DATA_TRANSFORMATION
                Case "FORMAT_STYLE"
                    result.IntentType = ExcelIntentType.FORMAT_STYLE
                Case "GENERAL_QUERY"
                    result.IntentType = ExcelIntentType.GENERAL_QUERY
                Case "UNCLEAR"
                    result.IntentType = ExcelIntentType.GENERAL_QUERY
                    result.Confidence = 0.3 ' 低置信度，需要确认
                Case Else
                    result.IntentType = ExcelIntentType.GENERAL_QUERY
            End Select

            ' 解析置信度
            If intentJson("confidence") IsNot Nothing Then
                result.Confidence = CDbl(intentJson("confidence"))
            End If

            ' 解析描述
            If intentJson("description") IsNot Nothing Then
                result.UserFriendlyDescription = intentJson("description").ToString()
            End If

            ' 解析是否需要确认
            If intentJson("requiresConfirmation") IsNot Nothing Then
                Dim needsConfirm = CBool(intentJson("requiresConfirmation"))
                If needsConfirm Then
                    result.Confidence = Math.Min(result.Confidence, 0.5) ' 降低置信度以触发确认
                End If
            End If

            Debug.WriteLine($"LLM意图解析: {result.IntentType}, 置信度: {result.Confidence:F2}, 描述: {result.UserFriendlyDescription}")

        Catch ex As Exception
            Debug.WriteLine($"ParseLLMIntentResponse 出错: {ex.Message}")
        End Try

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

        ' 添加严格的JSON Schema约束
        sb.AppendLine()
        sb.AppendLine(GetStrictJsonSchemaConstraint())

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 获取严格的JSON Schema约束（所有提示词共用）
    ''' </summary>
    Private Function GetStrictJsonSchemaConstraint() As String
        Return "
【JSON输出格式规范 - 必须严格遵守】

你必须且只能返回以下两种格式之一：

单命令格式：
{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""C1:C{lastRow}"", ""formula"": ""=A1+B1""}}

多命令格式：
{""commands"": [{""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""C1"", ""formula"": ""=A1+B1""}}, {""command"": ""ApplyFormula"", ""params"": {""targetRange"": ""E1"", ""formula"": ""=C1*D1""}}]}

【绝对禁止】
- 禁止使用 actions 数组
- 禁止使用 operations 数组
- 禁止省略 params 包装
- 禁止自创任何其他格式

【command类型】只能是: ApplyFormula, WriteData, FormatRange, CreateChart, CleanData

【占位符】使用 {lastRow} 表示最后一行

如果需求不明确，直接用中文回复询问用户，不要返回JSON。"
    End Function

    ''' <summary>
    ''' 生成用户友好的意图描述
    ''' </summary>
    Public Function GenerateUserFriendlyDescription(intent As IntentResult) As String
        Dim description As String

        Select Case intent.IntentType
            Case ExcelIntentType.DATA_ANALYSIS
                description = "对数据进行统计分析"
            Case ExcelIntentType.FORMULA_CALC
                description = "应用公式进行计算"
            Case ExcelIntentType.CHART_GEN
                description = "创建数据可视化图表"
            Case ExcelIntentType.DATA_CLEANING
                description = "清洗和整理数据"
            Case ExcelIntentType.REPORT_GEN
                description = "生成数据报表"
            Case ExcelIntentType.DATA_TRANSFORMATION
                description = "转换和处理数据"
            Case ExcelIntentType.FORMAT_STYLE
                description = "调整格式和样式"
            Case Else
                description = "处理您的请求"
        End Select

        ' 如果有提取到的实体，补充描述
        If intent.ExtractedEntities.ContainsKey("range") Then
            description &= $"（范围: {intent.ExtractedEntities("range")}）"
        ElseIf intent.ExtractedEntities.ContainsKey("column") Then
            description &= $"（{intent.ExtractedEntities("column")}列）"
        End If

        intent.UserFriendlyDescription = description
        Return description
    End Function

    ''' <summary>
    ''' 构建执行计划预览
    ''' </summary>
    Public Function BuildExecutionPlanPreview(intent As IntentResult) As List(Of ExecutionStep)
        Dim plan As New List(Of ExecutionStep)()

        Select Case intent.IntentType
            Case ExcelIntentType.DATA_ANALYSIS
                plan.Add(New ExecutionStep(1, "识别数据所在区域", "search"))
                plan.Add(New ExecutionStep(2, "分析数据结构和类型", "data"))
                plan.Add(New ExecutionStep(3, "执行统计计算", "formula"))
                plan.Add(New ExecutionStep(4, "输出分析结果", "data"))

            Case ExcelIntentType.FORMULA_CALC
                plan.Add(New ExecutionStep(1, "确定目标单元格", "search"))
                plan.Add(New ExecutionStep(2, "构建计算公式", "formula"))
                plan.Add(New ExecutionStep(3, "应用公式到指定范围", "formula"))

            Case ExcelIntentType.CHART_GEN
                plan.Add(New ExecutionStep(1, "识别图表数据源", "search"))
                plan.Add(New ExecutionStep(2, "选择合适的图表类型", "chart"))
                plan.Add(New ExecutionStep(3, "创建并配置图表", "chart"))
                plan.Add(New ExecutionStep(4, "调整图表位置和样式", "format"))

            Case ExcelIntentType.DATA_CLEANING
                plan.Add(New ExecutionStep(1, "扫描数据区域", "search"))
                plan.Add(New ExecutionStep(2, "识别需要清洗的内容", "data"))
                plan.Add(New ExecutionStep(3, "执行清洗操作", "clean"))
                plan.Add(New ExecutionStep(4, "验证清洗结果", "data"))

            Case ExcelIntentType.REPORT_GEN
                plan.Add(New ExecutionStep(1, "收集报表数据", "search"))
                plan.Add(New ExecutionStep(2, "设计报表结构", "data"))
                plan.Add(New ExecutionStep(3, "填充数据内容", "data"))
                plan.Add(New ExecutionStep(4, "应用报表格式", "format"))

            Case ExcelIntentType.DATA_TRANSFORMATION
                plan.Add(New ExecutionStep(1, "读取源数据", "search"))
                plan.Add(New ExecutionStep(2, "执行数据转换", "data"))
                plan.Add(New ExecutionStep(3, "输出转换结果", "data"))

            Case ExcelIntentType.FORMAT_STYLE
                plan.Add(New ExecutionStep(1, "选择目标区域", "search"))
                plan.Add(New ExecutionStep(2, "应用格式设置", "format"))

            Case Else
                plan.Add(New ExecutionStep(1, "分析您的需求", "search"))
                plan.Add(New ExecutionStep(2, "生成解决方案", "data"))
                plan.Add(New ExecutionStep(3, "执行操作", "default"))
        End Select

        ' 根据提取的实体更新步骤描述
        If intent.ExtractedEntities.ContainsKey("range") Then
            For Each execStep In plan
                If execStep.Description.Contains("区域") OrElse execStep.Description.Contains("范围") Then
                    execStep.WillModify = intent.ExtractedEntities("range")
                End If
            Next
        End If

        intent.ExecutionPlan = plan
        Return plan
    End Function

    ''' <summary>
    ''' 生成完整的意图澄清结果
    ''' </summary>
    Public Function GenerateIntentClarification(question As String, Optional context As JObject = Nothing) As IntentClarification
        Dim clarification As New IntentClarification()
        clarification.OriginalInput = question

        ' 识别意图
        Dim intent = IdentifyIntent(question, context)

        ' 生成描述
        clarification.Description = GenerateUserFriendlyDescription(intent)

        ' 构建执行计划
        clarification.ExecutionPlan = BuildExecutionPlanPreview(intent)

        ' 所有模式都需要确认
        clarification.RequiresConfirmation = True

        Return clarification
    End Function

    ''' <summary>
    ''' 将意图澄清结果转换为JSON（供前端使用）
    ''' </summary>
    Public Function IntentClarificationToJson(clarification As IntentClarification) As JObject
        Dim result As New JObject()
        result("description") = clarification.Description
        result("originalInput") = clarification.OriginalInput
        result("requiresConfirmation") = clarification.RequiresConfirmation

        Dim planArray As New JArray()
        For Each execStep In clarification.ExecutionPlan
            Dim stepObj As New JObject()
            stepObj("stepNumber") = execStep.StepNumber
            stepObj("description") = execStep.Description
            stepObj("icon") = execStep.Icon
            stepObj("willModify") = If(execStep.WillModify, "")
            stepObj("estimatedTime") = If(execStep.EstimatedTime, "1秒")
            planArray.Add(stepObj)
        Next
        result("plan") = planArray

        If clarification.ClarifyingQuestions.Count > 0 Then
            Dim questionsArray As New JArray()
            For Each q In clarification.ClarifyingQuestions
                questionsArray.Add(q)
            Next
            result("clarifyingQuestions") = questionsArray
        End If

        Return result
    End Function

#End Region

#Region "提示词模板"

    Private Function GetDataAnalysisPrompt() As String
        Return "你是Excel数据分析助手。

如果用户需求明确，返回JSON命令执行。
如果用户需求不明确，请先询问用户想要什么样的分析结果。

支持的操作: 公式计算、数据汇总、图表生成、数据清洗"
    End Function

    Private Function GetFormulaCalcPrompt() As String
        Return "你是Excel公式助手。

如果用户需求明确，返回JSON命令执行公式。
如果用户需求不明确，请先询问用户具体想计算什么。"
    End Function

    Private Function GetChartGenPrompt() As String
        Return "你是Excel图表助手。

如果用户需求明确，返回JSON命令创建图表。
如果用户需求不明确，请先询问用户想要什么类型的图表、数据范围等。"
    End Function

    Private Function GetDataCleaningPrompt() As String
        Return "你是Excel数据清洗助手。

如果用户需求明确，返回JSON命令清洗数据。
如果用户需求不明确，请先询问用户具体要做什么（去重、填充空值、去空格等）。"
    End Function

    Private Function GetReportGenPrompt() As String
        Return "你是Excel报表助手。

如果用户需求明确，返回JSON命令生成报表。
如果用户需求不明确，请先询问用户报表的具体内容和格式要求。"
    End Function

    Private Function GetDataTransformPrompt() As String
        Return "你是Excel数据转换助手。

如果用户需求明确，返回JSON命令进行数据转换。
如果用户需求不明确，请先询问用户具体的转换需求。"
    End Function

    Private Function GetFormatStylePrompt() As String
        Return "你是Excel格式化助手。

如果用户需求明确，返回JSON命令设置格式。
如果用户需求不明确，请先询问用户想要什么样的格式效果。"
    End Function

    Private Function GetGeneralPrompt() As String
        Return "你是Excel助手。

【重要原则】
1. 如果用户需求明确且可以执行，返回JSON命令
2. 如果用户需求不明确，必须先询问用户澄清：
   - 用户想对哪些数据操作？
   - 用户期望的结果是什么？
   - 涉及多个工作表时，请确认具体工作表名称
3. 对于简单问候或问答，直接用中文回复即可"
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
