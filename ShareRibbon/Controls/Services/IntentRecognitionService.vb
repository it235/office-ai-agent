' ShareRibbon\Controls\Services\IntentRecognitionService.vb
' 意图识别服务：分析用户输入并识别操作意图

Imports System.Diagnostics
Imports System.Linq
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 通用Office操作意图类型枚举 - 支持Excel/Word/PowerPoint
''' </summary>
Public Enum OfficeIntentType
    ' === 通用意图 ===
    GENERAL_QUERY       ' 一般查询/问答
    FORMAT_STYLE        ' 格式样式调整

    ' === Excel特有意图 ===
    DATA_ANALYSIS       ' 数据分析（统计、汇总、透视表）
    FORMULA_CALC        ' 公式计算
    CHART_GEN           ' 图表生成
    DATA_CLEANING       ' 数据清洗（去重、填充、格式化）
    REPORT_GEN          ' 报表生成
    DATA_TRANSFORMATION ' 数据转换（合并、拆分、转置）

    ' === Word特有意图 ===
    DOCUMENT_EDIT       ' 文档编辑（插入、删除、替换）
    TEXT_FORMAT         ' 文本格式化（字体、段落、样式）
    TABLE_OPERATION     ' 表格操作
    IMAGE_INSERT        ' 图片插入
    TOC_GENERATION      ' 目录生成
    REVIEW_COMMENT      ' 审阅批注

    ' === PowerPoint特有意图 ===
    SLIDE_CREATE        ' 创建幻灯片
    SLIDE_LAYOUT        ' 幻灯片布局
    ANIMATION_EFFECT    ' 动画效果
    TRANSITION_EFFECT   ' 切换效果
    TEMPLATE_APPLY      ' 应用模板
    SPEAKER_NOTES       ' 演讲者备注
End Enum

''' <summary>
''' Excel操作意图类型枚举（兼容旧代码）
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
    ''' 通用Office意图类型（支持Excel/Word/PowerPoint）
    ''' </summary>
    Public Property OfficeIntent As OfficeIntentType = OfficeIntentType.GENERAL_QUERY

    ''' <summary>
    ''' 主要意图类型（兼容旧代码，映射到ExcelIntentType）
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
''' 意图识别服务 - 支持Excel/Word/PowerPoint
''' </summary>
Public Class IntentRecognitionService

    ''' <summary>
    ''' 当前Office应用类型
    ''' </summary>
    Public Property AppType As OfficeApplicationType = OfficeApplicationType.Excel

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New()
        Me.AppType = OfficeApplicationType.Excel
    End Sub

    ''' <summary>
    ''' 带应用类型的构造函数
    ''' </summary>
    Public Sub New(appType As OfficeApplicationType)
        Me.AppType = appType
    End Sub

#Region "Excel关键词映射"

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

    ' 格式样式关键词（通用）
    Private Shared ReadOnly FormatStyleKeywords As String() = {
        "格式", "样式", "颜色", "字体", "边框", "对齐", "加粗",
        "斜体", "底色", "高亮", "条件格式", "美化"
    }

#End Region

#Region "Word关键词映射"

    ' 文档编辑关键词
    Private Shared ReadOnly DocumentEditKeywords As String() = {
        "插入", "删除", "替换", "复制", "粘贴", "剪切", "撤销",
        "查找", "搜索", "定位", "跳转", "选中", "全选"
    }

    ' 文本格式关键词
    Private Shared ReadOnly TextFormatKeywords As String() = {
        "字体", "字号", "加粗", "斜体", "下划线", "删除线", "上标", "下标",
        "段落", "行距", "缩进", "首行缩进", "对齐", "两端对齐"
    }

    ' 表格操作关键词
    Private Shared ReadOnly TableOperationKeywords As String() = {
        "表格", "插入表格", "删除表格", "合并单元格", "拆分单元格",
        "添加行", "添加列", "删除行", "删除列", "表格样式"
    }

    ' 图片插入关键词
    Private Shared ReadOnly ImageInsertKeywords As String() = {
        "图片", "插入图片", "添加图片", "图像", "截图", "照片",
        "调整图片", "裁剪", "旋转图片"
    }

    ' 目录生成关键词
    Private Shared ReadOnly TocGenerationKeywords As String() = {
        "目录", "生成目录", "更新目录", "插入目录", "大纲",
        "标题样式", "章节", "页码"
    }

    ' 审阅批注关键词
    Private Shared ReadOnly ReviewCommentKeywords As String() = {
        "批注", "注释", "评论", "修订", "审阅", "接受修订", "拒绝修订",
        "比较文档", "合并文档"
    }

#End Region

#Region "PowerPoint关键词映射"

    ' 幻灯片创建关键词
    Private Shared ReadOnly SlideCreateKeywords As String() = {
        "幻灯片", "新建幻灯片", "添加幻灯片", "插入幻灯片",
        "删除幻灯片", "复制幻灯片", "ppt", "演示文稿"
    }

    ' 幻灯片布局关键词
    Private Shared ReadOnly SlideLayoutKeywords As String() = {
        "布局", "版式", "标题幻灯片", "标题和内容", "空白幻灯片",
        "两栏内容", "比较", "仅标题"
    }

    ' 动画效果关键词
    Private Shared ReadOnly AnimationEffectKeywords As String() = {
        "动画", "进入动画", "退出动画", "强调动画", "路径动画",
        "淡入", "飞入", "缩放", "旋转动画", "动画顺序"
    }

    ' 切换效果关键词
    Private Shared ReadOnly TransitionEffectKeywords As String() = {
        "切换", "过渡", "幻灯片切换", "淡出", "推入", "擦除",
        "百叶窗", "棋盘", "切换时间"
    }

    ' 模板应用关键词
    Private Shared ReadOnly TemplateApplyKeywords As String() = {
        "模板", "主题", "应用模板", "设计", "配色方案",
        "背景", "更换背景", "幻灯片母版"
    }

    ' 演讲者备注关键词
    Private Shared ReadOnly SpeakerNotesKeywords As String() = {
        "备注", "演讲者备注", "笔记", "提示", "演讲稿",
        "演示者视图", "备注页"
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

        ' 根据AppType使用不同的关键词映射计算意图分数
        Select Case AppType
            Case OfficeApplicationType.Excel
                IdentifyExcelIntent(lowerQuestion, result)
            Case OfficeApplicationType.Word
                IdentifyWordIntent(lowerQuestion, result)
            Case OfficeApplicationType.PowerPoint
                IdentifyPowerPointIntent(lowerQuestion, result)
        End Select

        ' 通用格式样式意图检测
        Dim formatScore = CalculateKeywordScore(lowerQuestion, FormatStyleKeywords)
        If formatScore > result.Confidence Then
            result.IntentType = ExcelIntentType.FORMAT_STYLE
            result.OfficeIntent = OfficeIntentType.FORMAT_STYLE
            result.Confidence = Math.Min(formatScore, 1.0)
        End If

        ' 提取关键实体
        ExtractEntities(question, result)

        ' 判断是否可以使用直接命令
        DetermineExecutionMethod(result)

        Debug.WriteLine($"[{AppType}] 意图识别结果: {result.OfficeIntent}, 置信度: {result.Confidence:F2}")
        Return result
    End Function

    ''' <summary>
    ''' 识别Excel特有意图
    ''' </summary>
    Private Sub IdentifyExcelIntent(lowerQuestion As String, result As IntentResult)
        Dim scores As New Dictionary(Of OfficeIntentType, Double)()
        scores(OfficeIntentType.DATA_ANALYSIS) = CalculateKeywordScore(lowerQuestion, DataAnalysisKeywords)
        scores(OfficeIntentType.FORMULA_CALC) = CalculateKeywordScore(lowerQuestion, FormulaCalcKeywords)
        scores(OfficeIntentType.CHART_GEN) = CalculateKeywordScore(lowerQuestion, ChartGenKeywords)
        scores(OfficeIntentType.DATA_CLEANING) = CalculateKeywordScore(lowerQuestion, DataCleaningKeywords)
        scores(OfficeIntentType.REPORT_GEN) = CalculateKeywordScore(lowerQuestion, ReportGenKeywords)
        scores(OfficeIntentType.DATA_TRANSFORMATION) = CalculateKeywordScore(lowerQuestion, DataTransformKeywords)

        Dim maxScore As Double = 0
        Dim maxIntent = OfficeIntentType.GENERAL_QUERY

        For Each kvp In scores
            If kvp.Value > maxScore Then
                maxScore = kvp.Value
                maxIntent = kvp.Key
            End If
        Next

        If maxScore > 0.1 Then
            result.OfficeIntent = maxIntent
            result.IntentType = MapToExcelIntentType(maxIntent)
            result.Confidence = Math.Min(maxScore, 1.0)
        End If
    End Sub

    ''' <summary>
    ''' 识别Word特有意图
    ''' </summary>
    Private Sub IdentifyWordIntent(lowerQuestion As String, result As IntentResult)
        Dim scores As New Dictionary(Of OfficeIntentType, Double)()
        scores(OfficeIntentType.DOCUMENT_EDIT) = CalculateKeywordScore(lowerQuestion, DocumentEditKeywords)
        scores(OfficeIntentType.TEXT_FORMAT) = CalculateKeywordScore(lowerQuestion, TextFormatKeywords)
        scores(OfficeIntentType.TABLE_OPERATION) = CalculateKeywordScore(lowerQuestion, TableOperationKeywords)
        scores(OfficeIntentType.IMAGE_INSERT) = CalculateKeywordScore(lowerQuestion, ImageInsertKeywords)
        scores(OfficeIntentType.TOC_GENERATION) = CalculateKeywordScore(lowerQuestion, TocGenerationKeywords)
        scores(OfficeIntentType.REVIEW_COMMENT) = CalculateKeywordScore(lowerQuestion, ReviewCommentKeywords)

        Dim maxScore As Double = 0
        Dim maxIntent = OfficeIntentType.GENERAL_QUERY

        For Each kvp In scores
            If kvp.Value > maxScore Then
                maxScore = kvp.Value
                maxIntent = kvp.Key
            End If
        Next

        If maxScore > 0.1 Then
            result.OfficeIntent = maxIntent
            result.IntentType = ExcelIntentType.GENERAL_QUERY ' Word意图映射到通用查询
            result.Confidence = Math.Min(maxScore, 1.0)
        End If
    End Sub

    ''' <summary>
    ''' 识别PowerPoint特有意图
    ''' </summary>
    Private Sub IdentifyPowerPointIntent(lowerQuestion As String, result As IntentResult)
        Dim scores As New Dictionary(Of OfficeIntentType, Double)()
        scores(OfficeIntentType.SLIDE_CREATE) = CalculateKeywordScore(lowerQuestion, SlideCreateKeywords)
        scores(OfficeIntentType.SLIDE_LAYOUT) = CalculateKeywordScore(lowerQuestion, SlideLayoutKeywords)
        scores(OfficeIntentType.ANIMATION_EFFECT) = CalculateKeywordScore(lowerQuestion, AnimationEffectKeywords)
        scores(OfficeIntentType.TRANSITION_EFFECT) = CalculateKeywordScore(lowerQuestion, TransitionEffectKeywords)
        scores(OfficeIntentType.TEMPLATE_APPLY) = CalculateKeywordScore(lowerQuestion, TemplateApplyKeywords)
        scores(OfficeIntentType.SPEAKER_NOTES) = CalculateKeywordScore(lowerQuestion, SpeakerNotesKeywords)

        Dim maxScore As Double = 0
        Dim maxIntent = OfficeIntentType.GENERAL_QUERY

        For Each kvp In scores
            If kvp.Value > maxScore Then
                maxScore = kvp.Value
                maxIntent = kvp.Key
            End If
        Next

        If maxScore > 0.1 Then
            result.OfficeIntent = maxIntent
            result.IntentType = ExcelIntentType.GENERAL_QUERY ' PPT意图映射到通用查询
            result.Confidence = Math.Min(maxScore, 1.0)
        End If
    End Sub

    ''' <summary>
    ''' 将通用意图映射到Excel意图类型（兼容旧代码）
    ''' </summary>
    Private Function MapToExcelIntentType(intent As OfficeIntentType) As ExcelIntentType
        Select Case intent
            Case OfficeIntentType.DATA_ANALYSIS
                Return ExcelIntentType.DATA_ANALYSIS
            Case OfficeIntentType.FORMULA_CALC
                Return ExcelIntentType.FORMULA_CALC
            Case OfficeIntentType.CHART_GEN
                Return ExcelIntentType.CHART_GEN
            Case OfficeIntentType.DATA_CLEANING
                Return ExcelIntentType.DATA_CLEANING
            Case OfficeIntentType.REPORT_GEN
                Return ExcelIntentType.REPORT_GEN
            Case OfficeIntentType.DATA_TRANSFORMATION
                Return ExcelIntentType.DATA_TRANSFORMATION
            Case OfficeIntentType.FORMAT_STYLE
                Return ExcelIntentType.FORMAT_STYLE
            Case Else
                Return ExcelIntentType.GENERAL_QUERY
        End Select
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
    ''' 获取意图识别系统提示词 - 根据AppType返回不同的提示词
    ''' </summary>
    Private Function GetIntentRecognitionSystemPrompt() As String
        Select Case AppType
            Case OfficeApplicationType.Word
                Return GetWordIntentRecognitionPrompt()
            Case OfficeApplicationType.PowerPoint
                Return GetPowerPointIntentRecognitionPrompt()
            Case Else ' Excel
                Return GetExcelIntentRecognitionPrompt()
        End Select
    End Function

    ''' <summary>
    ''' Excel意图识别提示词
    ''' </summary>
    Private Function GetExcelIntentRecognitionPrompt() As String
        Return "你是一个Excel意图识别助手。分析用户的问题和上下文，识别用户想要执行的Excel操作。

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
3. 结合Excel上下文信息（如选中单元格、工作表）来更准确地判断意图"
    End Function

    ''' <summary>
    ''' Word意图识别提示词
    ''' </summary>
    Private Function GetWordIntentRecognitionPrompt() As String
        Return "你是一个Word意图识别助手。分析用户的问题和上下文，识别用户想要执行的Word操作。

请用JSON格式返回识别结果：
```json
{
  ""intentType"": ""DOCUMENT_EDIT"",
  ""confidence"": 0.85,
  ""description"": ""用户想要编辑文档内容"",
  ""requiresConfirmation"": false,
  ""suggestedAction"": ""直接执行文档编辑""
}
```

intentType必须是以下之一:
- DOCUMENT_EDIT: 文档编辑（插入、删除、替换文本）
- TEXT_FORMAT: 文本格式化（字体、段落、样式）
- TABLE_OPERATION: 表格操作（创建、编辑表格）
- IMAGE_INSERT: 图片插入和处理
- TOC_GENERATION: 目录生成和更新
- REVIEW_COMMENT: 审阅和批注
- FORMAT_STYLE: 格式样式调整
- GENERAL_QUERY: 一般问答（不需要操作Word）
- UNCLEAR: 意图不明确，需要进一步询问

confidence范围0-1，表示你对识别结果的确信程度。
requiresConfirmation: 如果意图明确且操作安全，设为false；如果需要用户确认，设为true。

注意：
1. 如果用户只是打招呼或闲聊，intentType设为GENERAL_QUERY，confidence设为0.9
2. 如果用户的请求涉及文档大幅修改但表述不清，requiresConfirmation设为true
3. 结合Word上下文信息（如选中文本、当前段落）来更准确地判断意图"
    End Function

    ''' <summary>
    ''' PowerPoint意图识别提示词
    ''' </summary>
    Private Function GetPowerPointIntentRecognitionPrompt() As String
        Return "你是一个PowerPoint意图识别助手。分析用户的问题和上下文，识别用户想要执行的PPT操作。

请用JSON格式返回识别结果：
```json
{
  ""intentType"": ""SLIDE_CREATE"",
  ""confidence"": 0.85,
  ""description"": ""用户想要创建新幻灯片"",
  ""requiresConfirmation"": false,
  ""suggestedAction"": ""直接创建幻灯片""
}
```

intentType必须是以下之一:
- SLIDE_CREATE: 创建幻灯片
- SLIDE_LAYOUT: 幻灯片布局和版式
- ANIMATION_EFFECT: 动画效果
- TRANSITION_EFFECT: 切换效果
- TEMPLATE_APPLY: 应用模板和主题
- SPEAKER_NOTES: 演讲者备注
- FORMAT_STYLE: 格式样式调整
- GENERAL_QUERY: 一般问答（不需要操作PPT）
- UNCLEAR: 意图不明确，需要进一步询问

confidence范围0-1，表示你对识别结果的确信程度。
requiresConfirmation: 如果意图明确且操作安全，设为false；如果需要用户确认，设为true。

注意：
1. 如果用户只是打招呼或闲聊，intentType设为GENERAL_QUERY，confidence设为0.9
2. 如果用户的请求涉及幻灯片大幅修改但表述不清，requiresConfirmation设为true
3. 结合PowerPoint上下文信息（如当前幻灯片、选中对象）来更准确地判断意图"
    End Function

    ''' <summary>
    ''' 解析LLM返回的意图识别结果 - 支持不同Office应用
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

            ' 解析意图类型 - 根据AppType处理不同的意图
            Dim intentTypeStr = intentJson("intentType")?.ToString()?.ToUpper()
            ParseIntentTypeByApp(intentTypeStr, result)

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

            Debug.WriteLine($"[{AppType}] LLM意图解析: {result.OfficeIntent}, 置信度: {result.Confidence:F2}, 描述: {result.UserFriendlyDescription}")

        Catch ex As Exception
            Debug.WriteLine($"ParseLLMIntentResponse 出错: {ex.Message}")
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 根据AppType解析意图类型
    ''' </summary>
    Private Sub ParseIntentTypeByApp(intentTypeStr As String, result As IntentResult)
        ' 通用意图
        Select Case intentTypeStr
            Case "FORMAT_STYLE"
                result.OfficeIntent = OfficeIntentType.FORMAT_STYLE
                result.IntentType = ExcelIntentType.FORMAT_STYLE
                Return
            Case "GENERAL_QUERY"
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
                result.IntentType = ExcelIntentType.GENERAL_QUERY
                Return
            Case "UNCLEAR"
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
                result.IntentType = ExcelIntentType.GENERAL_QUERY
                result.Confidence = 0.3
                Return
        End Select

        ' 根据AppType解析特定意图
        Select Case AppType
            Case OfficeApplicationType.Excel
                ParseExcelIntentType(intentTypeStr, result)
            Case OfficeApplicationType.Word
                ParseWordIntentType(intentTypeStr, result)
            Case OfficeApplicationType.PowerPoint
                ParsePowerPointIntentType(intentTypeStr, result)
        End Select
    End Sub

    ''' <summary>
    ''' 解析Excel意图类型
    ''' </summary>
    Private Sub ParseExcelIntentType(intentTypeStr As String, result As IntentResult)
        Select Case intentTypeStr
            Case "DATA_ANALYSIS"
                result.OfficeIntent = OfficeIntentType.DATA_ANALYSIS
                result.IntentType = ExcelIntentType.DATA_ANALYSIS
            Case "FORMULA_CALC"
                result.OfficeIntent = OfficeIntentType.FORMULA_CALC
                result.IntentType = ExcelIntentType.FORMULA_CALC
            Case "CHART_GEN"
                result.OfficeIntent = OfficeIntentType.CHART_GEN
                result.IntentType = ExcelIntentType.CHART_GEN
            Case "DATA_CLEANING"
                result.OfficeIntent = OfficeIntentType.DATA_CLEANING
                result.IntentType = ExcelIntentType.DATA_CLEANING
            Case "REPORT_GEN"
                result.OfficeIntent = OfficeIntentType.REPORT_GEN
                result.IntentType = ExcelIntentType.REPORT_GEN
            Case "DATA_TRANSFORMATION"
                result.OfficeIntent = OfficeIntentType.DATA_TRANSFORMATION
                result.IntentType = ExcelIntentType.DATA_TRANSFORMATION
            Case Else
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
                result.IntentType = ExcelIntentType.GENERAL_QUERY
        End Select
    End Sub

    ''' <summary>
    ''' 解析Word意图类型
    ''' </summary>
    Private Sub ParseWordIntentType(intentTypeStr As String, result As IntentResult)
        Select Case intentTypeStr
            Case "DOCUMENT_EDIT"
                result.OfficeIntent = OfficeIntentType.DOCUMENT_EDIT
            Case "TEXT_FORMAT"
                result.OfficeIntent = OfficeIntentType.TEXT_FORMAT
            Case "TABLE_OPERATION"
                result.OfficeIntent = OfficeIntentType.TABLE_OPERATION
            Case "IMAGE_INSERT"
                result.OfficeIntent = OfficeIntentType.IMAGE_INSERT
            Case "TOC_GENERATION"
                result.OfficeIntent = OfficeIntentType.TOC_GENERATION
            Case "REVIEW_COMMENT"
                result.OfficeIntent = OfficeIntentType.REVIEW_COMMENT
            Case Else
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
        End Select
        result.IntentType = ExcelIntentType.GENERAL_QUERY ' Word意图映射到通用
    End Sub

    ''' <summary>
    ''' 解析PowerPoint意图类型
    ''' </summary>
    Private Sub ParsePowerPointIntentType(intentTypeStr As String, result As IntentResult)
        Select Case intentTypeStr
            Case "SLIDE_CREATE"
                result.OfficeIntent = OfficeIntentType.SLIDE_CREATE
            Case "SLIDE_LAYOUT"
                result.OfficeIntent = OfficeIntentType.SLIDE_LAYOUT
            Case "ANIMATION_EFFECT"
                result.OfficeIntent = OfficeIntentType.ANIMATION_EFFECT
            Case "TRANSITION_EFFECT"
                result.OfficeIntent = OfficeIntentType.TRANSITION_EFFECT
            Case "TEMPLATE_APPLY"
                result.OfficeIntent = OfficeIntentType.TEMPLATE_APPLY
            Case "SPEAKER_NOTES"
                result.OfficeIntent = OfficeIntentType.SPEAKER_NOTES
            Case Else
                result.OfficeIntent = OfficeIntentType.GENERAL_QUERY
        End Select
        result.IntentType = ExcelIntentType.GENERAL_QUERY ' PPT意图映射到通用
    End Sub

    ''' <summary>
    ''' 获取优化后的系统提示词 - 使用PromptManager统一管理
    ''' </summary>
    Public Function GetOptimizedSystemPrompt(intent As IntentResult) As String
        ' 使用PromptManager获取组合后的提示词
        Dim context As New PromptContext With {
            .ApplicationType = AppType.ToString(),
            .IntentResult = intent,
            .FunctionMode = String.Empty
        }

        Return PromptManager.Instance.GetCombinedPrompt(context)
    End Function

    ''' <summary>
    ''' 根据Excel意图获取提示词
    ''' </summary>
    Private Function GetExcelPromptByIntent(intentType As ExcelIntentType) As String
        Select Case intentType
            Case ExcelIntentType.DATA_ANALYSIS
                Return GetDataAnalysisPrompt()
            Case ExcelIntentType.FORMULA_CALC
                Return GetFormulaCalcPrompt()
            Case ExcelIntentType.CHART_GEN
                Return GetChartGenPrompt()
            Case ExcelIntentType.DATA_CLEANING
                Return GetDataCleaningPrompt()
            Case ExcelIntentType.REPORT_GEN
                Return GetReportGenPrompt()
            Case ExcelIntentType.DATA_TRANSFORMATION
                Return GetDataTransformPrompt()
            Case ExcelIntentType.FORMAT_STYLE
                Return GetFormatStylePrompt()
            Case Else
                Return GetGeneralPrompt()
        End Select
    End Function

    ''' <summary>
    ''' 根据Word意图获取提示词
    ''' </summary>
    Private Function GetWordPromptByIntent(intentType As OfficeIntentType) As String
        Select Case intentType
            Case OfficeIntentType.DOCUMENT_EDIT
                Return GetWordDocumentEditPrompt()
            Case OfficeIntentType.TEXT_FORMAT
                Return GetWordTextFormatPrompt()
            Case OfficeIntentType.TABLE_OPERATION
                Return GetWordTableOperationPrompt()
            Case OfficeIntentType.IMAGE_INSERT
                Return GetWordImageInsertPrompt()
            Case OfficeIntentType.TOC_GENERATION
                Return GetWordTocGenerationPrompt()
            Case OfficeIntentType.REVIEW_COMMENT
                Return GetWordReviewCommentPrompt()
            Case OfficeIntentType.FORMAT_STYLE
                Return GetWordFormatStylePrompt()
            Case Else
                Return GetWordGeneralPrompt()
        End Select
    End Function

    ''' <summary>
    ''' 根据PowerPoint意图获取提示词
    ''' </summary>
    Private Function GetPowerPointPromptByIntent(intentType As OfficeIntentType) As String
        Select Case intentType
            Case OfficeIntentType.SLIDE_CREATE
                Return GetPptSlideCreatePrompt()
            Case OfficeIntentType.SLIDE_LAYOUT
                Return GetPptSlideLayoutPrompt()
            Case OfficeIntentType.ANIMATION_EFFECT
                Return GetPptAnimationEffectPrompt()
            Case OfficeIntentType.TRANSITION_EFFECT
                Return GetPptTransitionEffectPrompt()
            Case OfficeIntentType.TEMPLATE_APPLY
                Return GetPptTemplateApplyPrompt()
            Case OfficeIntentType.SPEAKER_NOTES
                Return GetPptSpeakerNotesPrompt()
            Case OfficeIntentType.FORMAT_STYLE
                Return GetPptFormatStylePrompt()
            Case Else
                Return GetPptGeneralPrompt()
        End Select
    End Function

    ''' <summary>
    ''' 获取严格的JSON Schema约束 - 根据AppType返回不同约束
    ''' </summary>
    Private Function GetStrictJsonSchemaConstraint() As String
        Select Case AppType
            Case OfficeApplicationType.Word
                Return GetWordJsonSchemaConstraint()
            Case OfficeApplicationType.PowerPoint
                Return GetPptJsonSchemaConstraint()
            Case Else ' Excel
                Return GetExcelJsonSchemaConstraint()
        End Select
    End Function

    ''' <summary>
    ''' Excel专用JSON Schema约束
    ''' </summary>
    Private Function GetExcelJsonSchemaConstraint() As String
        Return "
【Excel JSON输出格式规范 - 必须严格遵守】

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

【Excel command类型 - 只能使用以下5种】
1. ApplyFormula - 应用公式
2. WriteData - 写入数据
3. FormatRange - 格式化范围
4. CreateChart - 创建图表
5. CleanData - 清洗数据

【占位符】使用 {lastRow} 表示最后一行

如果需求不明确，直接用中文回复询问用户。"
    End Function

    ''' <summary>
    ''' Word专用JSON Schema约束
    ''' </summary>
    Private Function GetWordJsonSchemaConstraint() As String
        Return "
【Word JSON输出格式规范 - 必须严格遵守】

你必须且只能返回以下两种格式之一：

单命令格式：
{""command"": ""InsertText"", ""params"": {""text"": ""内容"", ""position"": ""cursor""}}

多命令格式：
{""commands"": [{""command"": ""InsertText"", ""params"": {""text"": ""内容1""}}, {""command"": ""FormatText"", ""params"": {""bold"": true}}]}

【绝对禁止】
- 禁止使用 actions 数组
- 禁止使用 operations 数组
- 禁止省略 params 包装
- 禁止自创任何其他格式
- 禁止使用Excel命令(WriteData, ApplyFormula等)
- 禁止使用PPT命令(InsertSlide, CreateSlides等)

【Word command类型 - 只能使用以下7种】
1. InsertText - 插入文本
   params: {text, position(cursor/start/end)}
2. FormatText - 格式化文本
   params: {bold, italic, underline, fontSize, fontName, color}
3. ReplaceText - 替换文本
   params: {findText, replaceWith, replaceAll}
4. InsertTable - 插入表格
   params: {rows, columns, data(二维数组)}
5. ApplyStyle - 应用样式
   params: {styleName(Heading1/Heading2/Normal等)}
6. GenerateTOC - 生成目录
   params: {position(start/cursor), levels(1-9), includePageNumbers}
7. BeautifyDocument - 美化文档
   params: {theme{h1,h2,h3,body}, margins{top,bottom,left,right}}

如果需求不明确，直接用中文回复询问用户。"
    End Function

    ''' <summary>
    ''' PowerPoint专用JSON Schema约束
    ''' </summary>
    Private Function GetPptJsonSchemaConstraint() As String
        Return "
【PowerPoint JSON输出格式规范 - 必须严格遵守】

你必须且只能返回以下两种格式之一：

单命令格式：
{""command"": ""InsertSlide"", ""params"": {""title"": ""标题"", ""content"": ""内容""}}

多命令格式：
{""commands"": [{""command"": ""InsertSlide"", ""params"": {""title"": ""标题1""}}, {""command"": ""AddAnimation"", ""params"": {""effect"": ""fadeIn""}}]}

【绝对禁止】
- 禁止使用 actions 数组
- 禁止使用 operations 数组
- 禁止省略 params 包装
- 禁止自创任何其他格式
- 禁止使用Excel命令(WriteData, ApplyFormula等)
- 禁止使用Word命令(GenerateTOC, BeautifyDocument等)

【PowerPoint command类型 - 只能使用以下9种】
1. InsertSlide - 插入单页幻灯片
   params: {position(end/start/指定位置), title, content, layout}
2. CreateSlides - 批量创建多页幻灯片(推荐)
   params: {slides数组[{title, content, layout}]}
3. InsertText - 插入文本到幻灯片
   params: {text, slideIndex(-1当前/0第一页)}
4. InsertShape - 插入形状
   params: {shapeType, text, position}
5. FormatSlide - 格式化幻灯片
   params: {slideIndex, background, titleStyle}
6. InsertTable - 插入表格到幻灯片
   params: {rows, columns, data, slideIndex}
7. AddAnimation - 添加动画效果
   params: {slideIndex(-1当前), effect(fadeIn/flyIn/zoom等), targetShapes(all/title/content)}
8. ApplyTransition - 应用切换效果
   params: {scope(all/current), transitionType(fade/push/wipe等), duration}
9. BeautifySlides - 美化幻灯片
   params: {scope(all/current), theme{background, titleFont, bodyFont}}

如果需求不明确，直接用中文回复询问用户。"
    End Function

    ''' <summary>
    ''' 生成用户友好的意图描述 - 根据AppType返回不同描述
    ''' </summary>
    Public Function GenerateUserFriendlyDescription(intent As IntentResult) As String
        Dim description As String

        Select Case AppType
            Case OfficeApplicationType.Word
                description = GetWordIntentDescription(intent.OfficeIntent)
            Case OfficeApplicationType.PowerPoint
                description = GetPowerPointIntentDescription(intent.OfficeIntent)
            Case Else ' Excel
                description = GetExcelIntentDescription(intent.IntentType)
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
    ''' 获取Excel意图描述
    ''' </summary>
    Private Function GetExcelIntentDescription(intentType As ExcelIntentType) As String
        Select Case intentType
            Case ExcelIntentType.DATA_ANALYSIS
                Return "对数据进行统计分析"
            Case ExcelIntentType.FORMULA_CALC
                Return "应用公式进行计算"
            Case ExcelIntentType.CHART_GEN
                Return "创建数据可视化图表"
            Case ExcelIntentType.DATA_CLEANING
                Return "清洗和整理数据"
            Case ExcelIntentType.REPORT_GEN
                Return "生成数据报表"
            Case ExcelIntentType.DATA_TRANSFORMATION
                Return "转换和处理数据"
            Case ExcelIntentType.FORMAT_STYLE
                Return "调整格式和样式"
            Case Else
                Return "处理您的Excel请求"
        End Select
    End Function

    ''' <summary>
    ''' 获取Word意图描述
    ''' </summary>
    Private Function GetWordIntentDescription(intentType As OfficeIntentType) As String
        Select Case intentType
            Case OfficeIntentType.DOCUMENT_EDIT
                Return "编辑文档内容"
            Case OfficeIntentType.TEXT_FORMAT
                Return "格式化文本样式"
            Case OfficeIntentType.TABLE_OPERATION
                Return "操作文档表格"
            Case OfficeIntentType.IMAGE_INSERT
                Return "插入和处理图片"
            Case OfficeIntentType.TOC_GENERATION
                Return "生成或更新目录"
            Case OfficeIntentType.REVIEW_COMMENT
                Return "添加审阅批注"
            Case OfficeIntentType.FORMAT_STYLE
                Return "调整文档格式"
            Case Else
                Return "处理您的Word请求"
        End Select
    End Function

    ''' <summary>
    ''' 获取PowerPoint意图描述
    ''' </summary>
    Private Function GetPowerPointIntentDescription(intentType As OfficeIntentType) As String
        Select Case intentType
            Case OfficeIntentType.SLIDE_CREATE
                Return "创建新幻灯片"
            Case OfficeIntentType.SLIDE_LAYOUT
                Return "调整幻灯片布局"
            Case OfficeIntentType.ANIMATION_EFFECT
                Return "添加动画效果"
            Case OfficeIntentType.TRANSITION_EFFECT
                Return "设置切换效果"
            Case OfficeIntentType.TEMPLATE_APPLY
                Return "应用模板主题"
            Case OfficeIntentType.SPEAKER_NOTES
                Return "编辑演讲者备注"
            Case OfficeIntentType.FORMAT_STYLE
                Return "调整幻灯片格式"
            Case Else
                Return "处理您的PPT请求"
        End Select
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

#Region "Word提示词模板"

    Private Function GetWordDocumentEditPrompt() As String
        Return "你是Word文档编辑助手。

如果用户需求明确，返回JSON命令执行文档编辑操作。
如果用户需求不明确，请先询问用户想要编辑什么内容。

支持的操作: 插入文本、删除文本、替换文本、查找定位"
    End Function

    Private Function GetWordTextFormatPrompt() As String
        Return "你是Word文本格式化助手。

如果用户需求明确，返回JSON命令设置文本格式。
如果用户需求不明确，请先询问用户想要什么样的格式效果（字体、字号、颜色、段落等）。"
    End Function

    Private Function GetWordTableOperationPrompt() As String
        Return "你是Word表格操作助手。

如果用户需求明确，返回JSON命令操作表格。
如果用户需求不明确，请先询问用户想要创建、编辑还是删除表格，以及具体的行列数。"
    End Function

    Private Function GetWordImageInsertPrompt() As String
        Return "你是Word图片处理助手。

如果用户需求明确，返回JSON命令插入或处理图片。
如果用户需求不明确，请先询问用户图片来源和插入位置。"
    End Function

    Private Function GetWordTocGenerationPrompt() As String
        Return "你是Word目录生成助手。

【支持的JSON命令】
- GenerateTOC: 生成目录

【GenerateTOC命令格式】
```json
{""command"": ""GenerateTOC"", ""params"": {""position"": ""start"", ""levels"": 3, ""includePageNumbers"": true}}
```

【参数说明】
- position: start(文档开头) 或 cursor(光标位置)
- levels: 目录层级(1-9)
- includePageNumbers: 是否显示页码

如果用户需求明确（如'生成目录'、'添加目录'），直接返回GenerateTOC命令。
如果用户需求不明确，请先询问：目录放在开头还是当前位置？显示几级标题？"
    End Function

    Private Function GetWordReviewCommentPrompt() As String
        Return "你是Word审阅批注助手。

如果用户需求明确，返回JSON命令添加批注或处理修订。
如果用户需求不明确，请先询问用户想要添加什么类型的批注。"
    End Function

    Private Function GetWordFormatStylePrompt() As String
        Return "你是Word格式样式助手。

【支持的JSON命令】
- BeautifyDocument: 美化文档（应用统一样式、字体、页边距）
- ApplyStyle: 应用单个样式

【BeautifyDocument命令格式】
```json
{""command"": ""BeautifyDocument"", ""params"": {""theme"": {""h1"": {""font"": ""微软雅黑"", ""size"": 22, ""bold"": true}, ""h2"": {""font"": ""微软雅黑"", ""size"": 18, ""bold"": true}, ""body"": {""font"": ""宋体"", ""size"": 12, ""lineSpacing"": 1.5}}, ""margins"": {""top"": 2.5, ""bottom"": 2.5, ""left"": 3.0, ""right"": 3.0}}}
```

【参数说明】
- theme.h1/h2/h3: 各级标题样式
- theme.body: 正文样式(含lineSpacing行间距)
- margins: 页边距(单位:厘米)

当用户说'美化文档'、'统一格式'、'调整样式'时，返回BeautifyDocument命令。
如果用户需求不明确，请先询问：想要什么字体？行间距多少？页边距要求？"
    End Function

    Private Function GetWordGeneralPrompt() As String
        Return "你是Word助手。

【重要原则】
1. 如果用户需求明确且可以执行，一定要返回可解析成code区的JSON代码，而不是普通文本
2. 如果用户需求不明确，必须先询问用户澄清：
   - 用户想对文档哪部分操作？
   - 用户期望的结果是什么？
3. 对于简单问候或问答，直接用中文回复即可"
    End Function

#End Region

#Region "PowerPoint提示词模板"

    Private Function GetPptSlideCreatePrompt() As String
        Return "你是PowerPoint幻灯片创建助手。

【支持的JSON命令】
- InsertSlide: 创建单页幻灯片
- CreateSlides: 批量创建多页幻灯片（推荐用于创建多页）

【CreateSlides命令格式（批量创建）】
```json
{""command"": ""CreateSlides"", ""params"": {""slides"": [{""title"": ""标题1"", ""content"": ""内容1"", ""layout"": ""titleAndContent""}, {""title"": ""标题2"", ""content"": ""内容2""}]}}
```

【InsertSlide命令格式（单页）】
```json
{""command"": ""InsertSlide"", ""params"": {""position"": ""end"", ""title"": ""标题"", ""content"": ""内容""}}
```

【layout可选值】
- title: 仅标题
- titleAndContent: 标题和内容（默认）
- twoContent: 两栏内容
- blank: 空白

当用户说'生成10页PPT'、'创建关于AI的演示文稿'时，使用CreateSlides命令批量创建。
当用户说'添加一页'、'新建幻灯片'时，使用InsertSlide命令创建单页。
如果用户需求不明确，请先询问：需要几页？每页的标题和内容？"
    End Function

    Private Function GetPptSlideLayoutPrompt() As String
        Return "你是PowerPoint布局助手。

如果用户需求明确，返回JSON命令调整幻灯片布局。
如果用户需求不明确，请先询问用户想要什么样的版式。"
    End Function

    Private Function GetPptAnimationEffectPrompt() As String
        Return "你是PowerPoint动画效果助手。

【支持的JSON命令】
- AddAnimation: 为幻灯片元素添加动画

【AddAnimation命令格式】
```json
{""command"": ""AddAnimation"", ""params"": {""slideIndex"": -1, ""effect"": ""fadeIn"", ""targetShapes"": ""all""}}
```

【参数说明】
- slideIndex: -1表示当前幻灯片，0表示第一页
- effect: fadeIn(淡入), flyIn(飞入), zoom(缩放), wipe(擦除), appear(出现), float(浮动)
- targetShapes: all(所有元素), title(仅标题), content(仅内容)

当用户说'添加动画'、'让元素淡入'时，直接返回AddAnimation命令。
如果用户需求不明确，请先询问：要添加什么效果？应用到哪些元素？"
    End Function

    Private Function GetPptTransitionEffectPrompt() As String
        Return "你是PowerPoint切换效果助手。

【支持的JSON命令】
- ApplyTransition: 设置幻灯片切换效果

【ApplyTransition命令格式】
```json
{""command"": ""ApplyTransition"", ""params"": {""scope"": ""all"", ""transitionType"": ""fade"", ""duration"": 1.0}}
```

【参数说明】
- scope: all(所有幻灯片), current(当前幻灯片)
- transitionType: fade(淡出), push(推入), wipe(擦除), split(拆分), reveal(显示), random(随机)
- duration: 切换时间(秒)

当用户说'添加切换效果'、'设置幻灯片过渡'时，直接返回ApplyTransition命令。
如果用户需求不明确，请先询问：应用到所有幻灯片还是当前页？要什么效果？"
    End Function

    Private Function GetPptTemplateApplyPrompt() As String
        Return "你是PowerPoint模板主题助手。

【支持的JSON命令】
- BeautifySlides: 美化幻灯片（应用统一主题、字体、配色）

【BeautifySlides命令格式】
```json
{""command"": ""BeautifySlides"", ""params"": {""scope"": ""all"", ""theme"": {""background"": ""#F5F5F5"", ""titleFont"": {""name"": ""微软雅黑"", ""size"": 28, ""color"": ""#333333""}, ""bodyFont"": {""name"": ""微软雅黑"", ""size"": 18, ""color"": ""#666666""}}}}
```

【参数说明】
- scope: all(所有幻灯片), current(当前幻灯片)
- theme.background: 背景颜色(十六进制)
- theme.titleFont: 标题字体设置
- theme.bodyFont: 正文字体设置

当用户说'美化PPT'、'统一风格'、'应用主题'时，返回BeautifySlides命令。
如果用户需求不明确，请先询问：想要什么配色？字体有什么要求？"
    End Function

    Private Function GetPptSpeakerNotesPrompt() As String
        Return "你是PowerPoint演讲者备注助手。

如果用户需求明确，返回JSON命令编辑演讲者备注。
如果用户需求不明确，请先询问用户想要添加什么内容到备注。"
    End Function

    Private Function GetPptFormatStylePrompt() As String
        Return "你是PowerPoint格式样式助手。

如果用户需求明确，返回JSON命令设置幻灯片格式。
如果用户需求不明确，请先询问用户想要什么样的格式效果。"
    End Function

    Private Function GetPptGeneralPrompt() As String
        Return "你是PowerPoint助手。

【重要原则】
1. 如果用户需求明确且可以执行，返回JSON命令
2. 如果用户需求不明确，必须先询问用户澄清：
   - 用户想对哪张幻灯片操作？
   - 用户期望的结果是什么？
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

#Region "上下文相关性检查"

    ''' <summary>
    ''' 异步检查新问题是否与历史对话相关（是追问或继续之前的话题）
    ''' </summary>
    ''' <param name="newQuestion">新问题</param>
    ''' <param name="historyMessages">历史消息列表</param>
    ''' <returns>True表示相关（追问），False表示无关（新话题）</returns>
    Public Async Function IsFollowUpQuestionAsync(newQuestion As String, historyMessages As List(Of HistoryMessage)) As Task(Of Boolean)
        ' 如果没有历史记录，则认为是新话题
        If historyMessages Is Nothing OrElse historyMessages.Count < 2 Then
            Return False
        End If

        ' 获取最近的对话上下文（排除system消息，最多取最近4条）
        Dim filteredHistory = historyMessages.Where(Function(m) m.role <> "system").ToList()
        Dim takeCount = Math.Min(4, filteredHistory.Count)
        Dim recentHistory = filteredHistory.Skip(filteredHistory.Count - takeCount).ToList()

        If recentHistory.Count = 0 Then
            Return False
        End If

        Try
            ' 获取API配置
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
                Return False
            End If

            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)

            Dim apiUrl = cfg.url
            Dim apiKey = cfg.key
            Dim modelName = selectedModel.modelName

            ' 构建上下文摘要
            Dim contextSummary As New StringBuilder()
            For Each msg In recentHistory
                Dim roleLabel = If(msg.role = "user", "用户", "AI")
                Dim contentPreview = If(msg.content?.Length > 200, msg.content.Substring(0, 200) & "...", msg.content)
                contextSummary.AppendLine($"{roleLabel}: {contentPreview}")
            Next

            ' 构建判断提示词
            Dim systemPrompt = "你是一个对话上下文分析助手。判断用户的新问题是否与之前的对话相关。

只返回JSON格式：
```json
{""isFollowUp"": true, ""reason"": ""简短原因""}
```

判断标准：
- isFollowUp=true: 新问题是对之前话题的追问、补充、澄清或继续
- isFollowUp=false: 新问题是全新的话题，与之前对话无关

示例：
- 之前讨论Excel公式，新问题""还有其他方法吗"" → true
- 之前讨论Excel公式，新问题""帮我画个图表"" → false (新话题)
- 之前讨论数据分析，新问题""这个结果不对"" → true (追问)"

            Dim userMessage = $"之前的对话：
{contextSummary}

新问题：{newQuestion}

判断新问题是否与之前对话相关？"

            ' 构建请求
            Dim messages As New JArray()
            messages.Add(New JObject From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject From {{"role", "user"}, {"content", userMessage}})

            Dim requestBody As New JObject()
            requestBody("model") = modelName
            requestBody("messages") = messages
            requestBody("temperature") = 0.2
            requestBody("max_tokens") = 100
            requestBody("stream") = False

            ' 发送请求
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(15)

                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                request.Content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request)
                    If response.IsSuccessStatusCode Then
                        Dim responseContent = Await response.Content.ReadAsStringAsync()
                        Return ParseFollowUpResponse(responseContent)
                    End If
                End Using
            End Using

        Catch ex As Exception
            Debug.WriteLine($"IsFollowUpQuestionAsync 出错: {ex.Message}")
        End Try

        ' 默认认为可能相关（避免误判导致每次都弹框）
        Return True
    End Function

    ''' <summary>
    ''' 解析追问判断的响应
    ''' </summary>
    Private Function ParseFollowUpResponse(responseContent As String) As Boolean
        Try
            Dim responseJson = JObject.Parse(responseContent)
            Dim choices = responseJson("choices")
            If choices Is Nothing OrElse choices.Count = 0 Then Return True

            Dim content = choices(0)("message")?("content")?.ToString()
            If String.IsNullOrEmpty(content) Then Return True

            ' 提取JSON部分
            Dim jsonMatch = Regex.Match(content, "\{[\s\S]*\}")
            If Not jsonMatch.Success Then Return True

            Dim resultJson = JObject.Parse(jsonMatch.Value)
            Dim isFollowUpToken = resultJson("isFollowUp")
            If isFollowUpToken Is Nothing Then Return True

            Dim isFollowUp As Boolean = isFollowUpToken.Value(Of Boolean)()

            Debug.WriteLine($"追问判断结果: isFollowUp={isFollowUp}, 原因={resultJson("reason")}")
            Return isFollowUp

        Catch ex As Exception
            Debug.WriteLine($"ParseFollowUpResponse 出错: {ex.Message}")
            Return True ' 默认认为相关
        End Try
    End Function

#End Region

End Class
