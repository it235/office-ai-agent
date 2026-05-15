' ShareRibbon\Services\Reformat\SmartFormattingOrchestrator.vb
' 智能排版编排器 - 连接 DocumentAnalyzer、FormattingKnowledgeEngine、SemanticRenderingEngine
' 支撑速排/对话/克隆三种排版模式。本编排器不直接调用AI，只做数据准备和规则判断。

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json.Linq

' ============================================================
'  枚举
' ============================================================

''' <summary>用户排版意图类型</summary>
Public Enum IntentType
    ''' <summary>自动排版 - 系统自动检测并推荐最佳标准</summary>
    AutoFormat
    ''' <summary>按指定标准排版 - 用户明确指定排版标准</summary>
    StandardFormat
    ''' <summary>样式克隆 - 照范文格式排版</summary>
    StyleClone
    ''' <summary>具体微调 - 只修改某项格式（如"标题再大一点"）</summary>
    SpecificTweak
    ''' <summary>格式清洗 - 清理混乱空白和格式污染</summary>
    FormatCleanup
End Enum

' ============================================================
'  预览方案（JSON序列化后发送到前端卡片）
' ============================================================

''' <summary>排版预览方案 - 发送到前端展示的完整方案数据</summary>
Public Class ReformatPreviewPlan
    ''' <summary>检测到的文档类型</summary>
    Public Property DetectedType As DocumentType = DocumentType.Unknown
    ''' <summary>类型识别置信度（0-1）</summary>
    Public Property TypeConfidence As Double = 0.0
    ''' <summary>推荐标准名称（如"GB/T 9704-2012"）</summary>
    Public Property StandardName As String = ""
    ''' <summary>标准描述</summary>
    Public Property StandardDescription As String = ""
    ''' <summary>即将发生的格式变更列表</summary>
    Public Property Changes As List(Of FormatChange)
    ''' <summary>总段落数</summary>
    Public Property TotalParagraphs As Integer = 0
    ''' <summary>总样式变更数</summary>
    Public Property TotalStyleChanges As Integer = 0
    ''' <summary>将要应用的语义样式映射</summary>
    Public Property SemanticMapping As SemanticStyleMapping
    ''' <summary>页面设置</summary>
    Public Property PageSettings As PageConfig
    ''' <summary>是否需要AI语义标注（True=Changes中的标签为初步结果，需AI标注后才能渲染）</summary>
    Public Property NeedsAITagging As Boolean = True
    ''' <summary>段落类型列表("text"/"image"/"table"/"formula")，用于渲染时跳过非文本段落</summary>
    Public Property ParagraphTypes As List(Of String) = Nothing

    Public Sub New()
        Changes = New List(Of FormatChange)()
        SemanticMapping = New SemanticStyleMapping()
        PageSettings = New PageConfig()
    End Sub

    ' -- consumer-compat aliases --
    Public Property DocumentTypeName As String = ""
    Public ReadOnly Property TotalChanges As Integer
        Get
            Return TotalStyleChanges
        End Get
    End Property
    Public ReadOnly Property SectionCount As Integer
        Get
            Return If(Changes IsNot Nothing, Changes.Count, 0)
        End Get
    End Property

    Public Function ToPreviewJson() As JObject
        Dim json As New JObject()
        json("docType") = DocumentTypeName
        json("confidence") = TypeConfidence
        json("standard") = StandardName
        json("totalChanges") = TotalChanges
        json("sectionCount") = SectionCount
        Dim changesArray As New JArray()
        If Changes IsNot Nothing Then
            For Each c In Changes
                Dim item As New JObject()
                item("section") = c.ChangeDescription
                item("tagId") = c.NewTag
                item("description") = c.ChangeDescription
                item("count") = 1
                changesArray.Add(item)
            Next
        End If
        json("changes") = changesArray
        Return json
    End Function
End Class

''' <summary>单处格式变更项（一个段落的格式变化描述）</summary>
Public Class FormatChange
    ''' <summary>段落索引</summary>
    Public Property ParagraphIndex As Integer = -1
    ''' <summary>段落文本预览（前50字）</summary>
    Public Property ParagraphPreview As String = ""
    ''' <summary>变更前的语义标签</summary>
    Public Property OldTag As String = ""
    ''' <summary>变更后的语义标签</summary>
    Public Property NewTag As String = ""
    ''' <summary>新字体描述（如"仿宋_GB2312 16pt"）</summary>
    Public Property NewFont As String = ""
    ''' <summary>新对齐方式（如"居中"、"两端对齐"）</summary>
    Public Property NewAlignment As String = ""
    ''' <summary>新缩进描述（如"首行缩进2字符"）</summary>
    Public Property NewIndent As String = ""
    ''' <summary>变更描述（如"宋体→仿宋 16pt, 居中→两端对齐"）</summary>
    Public Property ChangeDescription As String = ""
End Class

' ============================================================
'  用户意图
' ============================================================

''' <summary>从自然语言中解析出的用户排版意图</summary>
Public Class FormatIntent
    ''' <summary>目标文档类型</summary>
    Public Property TargetDocumentType As DocumentType = DocumentType.Unknown
    ''' <summary>排版意图类型</summary>
    Public Property IntentType As IntentType = IntentType.AutoFormat
    ''' <summary>具体格式要求列表（如"标题用红色", "行距改成1.5"）</summary>
    Public Property SpecificRequests As List(Of String)
    ''' <summary>目标标准名称（如"GB/T 9704-2012"）</summary>
    Public Property TargetStandardName As String = ""

    Public Sub New()
        SpecificRequests = New List(Of String)()
    End Sub
End Class

' ============================================================
'  格式差异（克隆对比用）
' ============================================================

''' <summary>源文档与范文之间的完整格式差异分析</summary>
Public Class FormatDiff
    ''' <summary>差异总数</summary>
    Public Property TotalDifferences As Integer = 0
    ''' <summary>差异明细列表</summary>
    Public Property Differences As List(Of FormatDifference)
    ''' <summary>范文标准名称</summary>
    Public Property MirrorStandardName As String = ""
    ''' <summary>范文分析结果</summary>
    Public Property MirrorAnalysis As DocumentAnalysisResult

    Public Sub New()
        Differences = New List(Of FormatDifference)()
        MirrorAnalysis = New DocumentAnalysisResult()
    End Sub
End Class

''' <summary>单处格式差异项——某段落当前格式与范文对应格式的对比</summary>
Public Class FormatDifference
    ''' <summary>段落索引（-1 表示全局差异）</summary>
    Public Property ParagraphIndex As Integer = -1
    ''' <summary>段落文本预览（前50字）</summary>
    Public Property ParagraphPreview As String = ""
    ''' <summary>源文档字体描述</summary>
    Public Property SourceFont As String = ""
    ''' <summary>范文字体描述</summary>
    Public Property MirrorFont As String = ""
    ''' <summary>差异文字描述</summary>
    Public Property DifferenceDescription As String = ""
End Class

' ============================================================
'  对话式微调上下文
' ============================================================

''' <summary>对话式排版的微调状态上下文，跟踪当前分析/映射/标准和对话历史</summary>
Public Class RefinementContext
    ''' <summary>当前文档分析结果</summary>
    Public Property CurrentAnalysis As DocumentAnalysisResult
    ''' <summary>当前使用的样式映射（可能已通过"标题再大一点"等指令修改）</summary>
    Public Property CurrentMapping As SemanticStyleMapping
    ''' <summary>当前使用的排版标准</summary>
    Public Property CurrentStandard As FormattingStandard
    ''' <summary>当前预览方案</summary>
    Public Property CurrentPreviewPlan As ReformatPreviewPlan
    ''' <summary>对话历史（用户微调命令列表，用于"再大一点"等上下文敏感指令）</summary>
    Public Property ConversationHistory As List(Of String)
    ''' <summary>是否已经应用到文档</summary>
    Public Property IsApplied As Boolean = False
    ''' <summary>消费者兼容别名：最后一次预览方案</summary>
    Public Property LastPlan As ReformatPreviewPlan
        Get
            Return CurrentPreviewPlan
        End Get
        Set(value As ReformatPreviewPlan)
            CurrentPreviewPlan = value
        End Set
    End Property

    Public Sub New()
        CurrentAnalysis = New DocumentAnalysisResult()
        CurrentMapping = New SemanticStyleMapping()
        CurrentStandard = Nothing
        CurrentPreviewPlan = Nothing
        ConversationHistory = New List(Of String)()
    End Sub

    ''' <summary>一次性更新上下文中的分析/映射/标准和预览方案</summary>
    Public Sub Update(analysis As DocumentAnalysisResult,
                     mapping As SemanticStyleMapping,
                     standard As FormattingStandard,
                     previewPlan As ReformatPreviewPlan)
        CurrentAnalysis = analysis
        CurrentMapping = mapping
        CurrentStandard = standard
        CurrentPreviewPlan = previewPlan
    End Sub

    ''' <summary>记录一条用户微调命令到对话历史</summary>
    Public Sub AddConversation(userCommand As String)
        ConversationHistory.Add(userCommand)
    End Sub

    ''' <summary>清空所有上下文状态</summary>
    Public Sub Clear()
        CurrentAnalysis = New DocumentAnalysisResult()
        CurrentMapping = New SemanticStyleMapping()
        CurrentStandard = Nothing
        CurrentPreviewPlan = Nothing
        ConversationHistory.Clear()
    End Sub
End Class

' ============================================================
'  主编排器
' ============================================================

''' <summary>
''' 智能排版编排器 —— 中央调度器。
''' 上层调用方（ChatFormatterAgent / ReformatService）编排完整流程：
'''   DocumentAnalyzer → FormattingKnowledgeEngine → 本编排器 → SemanticRenderingEngine
''' 本编排器不直接调用 AI，只做数据准备和规则判断。
''' </summary>
Public Class SmartFormattingOrchestrator

    Private ReadOnly _analyzer As DocumentAnalyzer
    Private ReadOnly _knowledgeEngine As FormattingKnowledgeEngine
    Private ReadOnly _refinementContext As RefinementContext

    ' ---- 标准名称关键词映射（用于 ParseUserIntent） ----
    Private Shared ReadOnly StandardKeywords As Dictionary(Of String, String) =
        New Dictionary(Of String, String) From {
            {"公文", "GB/T 9704-2012"},
            {"GB/T 9704", "GB/T 9704-2012"},
            {"国标", "GB/T 9704-2012"},
            {"党政", "GB/T 9704-2012"},
            {"学术", "学术论文通用格式"},
            {"论文", "学术论文通用格式"},
            {"参考文献", "GB/T 7714-2015"},
            {"GB/T 7714", "GB/T 7714-2015"},
            {"商务", "商务报告通用规范"},
            {"报告", "商务报告通用规范"},
            {"商业", "商务报告通用规范"}
        }

    ''' <summary>获取当前对话式微调的上下文状态</summary>
    Public ReadOnly Property RefinementContext As RefinementContext
        Get
            Return _refinementContext
        End Get
    End Property

    ''' <summary>使用默认 DocumentAnalyzer 和 FormattingKnowledgeEngine 构造</summary>
    Public Sub New()
        _analyzer = New DocumentAnalyzer()
        _knowledgeEngine = New FormattingKnowledgeEngine()
        _refinementContext = New RefinementContext()
    End Sub

    ''' <summary>注入自定义分析器和知识引擎</summary>
    Public Sub New(docAnalyzer As DocumentAnalyzer, knowledgeEngine As FormattingKnowledgeEngine)
        _analyzer = docAnalyzer
        _knowledgeEngine = knowledgeEngine
        _refinementContext = New RefinementContext()
    End Sub

    ' ============================================================
    '  模式A：速排（一键分析 → 推荐 → 预览）
    ' ============================================================

    ''' <summary>
    ''' 分析文档并推荐排版方案。
    ''' 内部流程：Analyze → DetectType → GetStandard → GeneratePreviewPlan
    ''' </summary>
    ''' <param name="paragraphTexts">文档各段落的纯文本列表</param>
    Public Function AnalyzeAndRecommend(paragraphTexts As List(Of String)) As ReformatPreviewPlan
        If paragraphTexts Is Nothing OrElse paragraphTexts.Count = 0 Then
            Return New ReformatPreviewPlan()
        End If

        ' 1. 分析文档类型和结构
        Dim analysis = _analyzer.Analyze(paragraphTexts)

        ' 2. 获取匹配标准
        Dim standard = _knowledgeEngine.GetStandardForDocumentType(analysis.DocumentType)

        If standard Is Nothing Then
            Return New ReformatPreviewPlan With {
                .DetectedType = analysis.DocumentType,
                .TypeConfidence = analysis.Confidence,
                .TotalParagraphs = analysis.ParagraphCount
            }
        End If

        ' 3. 生成预览方案
        Dim plan = GeneratePreviewPlan(analysis, standard, paragraphTexts)

        ' 4. 保存到对话式微调上下文
        _refinementContext.Update(analysis, plan.SemanticMapping, standard, plan)

        Return plan
    End Function

    ''' <summary>
    ''' 分析文档并推荐排版方案（增强版，接受Word富文本格式信息）
    ''' </summary>
    ''' <param name="paragraphTexts">文档各段落的纯文本列表</param>
    ''' <param name="paragraphStyles">段落样式名称列表</param>
    ''' <param name="paragraphFontSizes">段落字号列表（pt）</param>
    ''' <param name="paragraphIsBold">段落是否加粗列表</param>
    Public Function AnalyzeAndRecommend(paragraphTexts As List(Of String),
                                        paragraphStyles As List(Of String),
                                        paragraphFontSizes As List(Of Single),
                                        paragraphIsBold As List(Of Boolean)) As ReformatPreviewPlan
        If paragraphTexts Is Nothing OrElse paragraphTexts.Count = 0 Then
            Return New ReformatPreviewPlan()
        End If

        ' 使用增强版分析器
        Dim analysis = _analyzer.Analyze(paragraphTexts, paragraphStyles, paragraphFontSizes, paragraphIsBold)

        Dim standard = _knowledgeEngine.GetStandardForDocumentType(analysis.DocumentType)

        If standard Is Nothing Then
            Return New ReformatPreviewPlan With {
                .DetectedType = analysis.DocumentType,
                .TypeConfidence = analysis.Confidence,
                .TotalParagraphs = analysis.ParagraphCount
            }
        End If

        Dim plan = GeneratePreviewPlan(analysis, standard, paragraphTexts)
        _refinementContext.Update(analysis, plan.SemanticMapping, standard, plan)
        Return plan
    End Function

    ' ============================================================
    '  模式B：对话式排版 —— 自然语言指令解析
    ' ============================================================

    ''' <summary>
    ''' 解析用户自然语言排版指令。
    ''' 纯规则实现（不调用AI），返回结构化的 FormatIntent。
    ''' 调用方（ChatFormatterAgent）可进一步用 AI 增强解析结果。
    ''' </summary>
    ''' <param name="userMessage">用户输入的自然语言消息</param>
    ''' <param name="analysis">当前文档分析结果（用于上下文推断）</param>
    Public Function ParseUserIntent(userMessage As String, analysis As DocumentAnalysisResult) As FormatIntent
        Dim intent As New FormatIntent()
        If String.IsNullOrWhiteSpace(userMessage) Then Return intent

        Dim message = userMessage.Trim()

        ' ---- 1. 识别意图类型 ----

        ' 克隆意图
        If ContainsAny(message, {"克隆", "照这个", "范文", "模仿", "参照", "格式克隆"}) Then
            intent.IntentType = IntentType.StyleClone
            Return intent
        End If

        ' 清洗意图
        If ContainsAny(message, {"清洗", "清理", "清除格式", "去除格式", "格式清理"}) Then
            intent.IntentType = IntentType.FormatCleanup
            Return intent
        End If

        ' 标准指定意图（如 "按公文标准"、"用GB/T 9704"）
        For Each kvp In StandardKeywords
            If message.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0 Then
                intent.TargetStandardName = kvp.Value
                intent.IntentType = IntentType.StandardFormat
                Exit For
            End If
        Next

        ' 微调意图（"大一点"、"红色"、"加粗" 等）
        If Not HasTweakIntent(message) Then
            ' 未匹配到其他目的时默认为自动排版
            If intent.IntentType = IntentType.AutoFormat AndAlso
               ContainsAny(message, {"排版", "格式", "段落", "字体", "行距", "整理"}) Then
                intent.IntentType = IntentType.AutoFormat
            End If
        Else
            intent.IntentType = IntentType.SpecificTweak
        End If

        ' ---- 2. 提取具体格式请求 ----

        ' 字号调整
        If Regex.IsMatch(message, "(大|小).{0,4}(点|一些|一点)") Then
            intent.SpecificRequests.Add("adjust_font_size")
        End If
        If Regex.IsMatch(message, "(标题|正文).{0,4}(大|小)") Then
            intent.SpecificRequests.Add("adjust_font_size")
        End If

        ' 颜色
        Dim colors = {"红色", "蓝色", "黑色", "绿色", "白色"}
        For Each color In colors
            If message.Contains(color) Then
                intent.SpecificRequests.Add($"color_{color}")
            End If
        Next

        ' 行距
        Dim lineSpacingMatch = Regex.Match(message, "行距.{0,4}([\d.]+)")
        If lineSpacingMatch.Success Then
            intent.SpecificRequests.Add($"line_spacing_{lineSpacingMatch.Groups(1).Value}")
        End If

        ' 对齐
        If ContainsAny(message, {"居中", "居中对齐"}) Then
            intent.SpecificRequests.Add("align_center")
        ElseIf ContainsAny(message, {"左对齐", "靠左"}) Then
            intent.SpecificRequests.Add("align_left")
        ElseIf ContainsAny(message, {"右对齐", "靠右"}) Then
            intent.SpecificRequests.Add("align_right")
        ElseIf ContainsAny(message, {"两端对齐"}) Then
            intent.SpecificRequests.Add("align_justify")
        End If

        ' 加粗
        If ContainsAny(message, {"加粗", "粗体", "粗一点"}) Then
            intent.SpecificRequests.Add("bold")
        End If

        ' 缩进
        If message.Contains("缩进") Then
            Dim indentMatch = Regex.Match(message, "缩进.{0,4}([\d.]+)")
            If indentMatch.Success Then
                intent.SpecificRequests.Add($"indent_{indentMatch.Groups(1).Value}")
            Else
                intent.SpecificRequests.Add("indent")
            End If
        End If

        ' ---- 3. 识别目标文档类型 ----

        If ContainsAny(message, {"公文", "通知", "决定", "批复", "请示", "函"}) Then
            intent.TargetDocumentType = DocumentType.OfficialDocument
        ElseIf ContainsAny(message, {"论文", "学术", "期刊", "学报"}) Then
            intent.TargetDocumentType = DocumentType.AcademicPaper
        ElseIf ContainsAny(message, {"报告", "商务", "商业", "汇报", "总结"}) Then
            intent.TargetDocumentType = DocumentType.BusinessReport
        ElseIf ContainsAny(message, {"合同", "协议", "合约"}) Then
            intent.TargetDocumentType = DocumentType.Contract
        ElseIf ContainsAny(message, {"简历", "履历"}) Then
            intent.TargetDocumentType = DocumentType.[Resume]
        Else
            ' 从标准名称反推文档类型
            If Not String.IsNullOrEmpty(intent.TargetStandardName) Then
                Dim standard = _knowledgeEngine.GetStandardByName(intent.TargetStandardName)
                If standard IsNot Nothing AndAlso standard.ApplicableDocumentTypes.Count > 0 Then
                    Dim parsed As DocumentType
                    If [Enum].TryParse(Of DocumentType)(standard.ApplicableDocumentTypes(0), parsed) Then
                        intent.TargetDocumentType = parsed
                    End If
                End If
            End If
        End If

        Return intent
    End Function

    ' ============================================================
    '  预览方案生成
    ' ============================================================

    ''' <summary>
    ''' 根据文档分析结果和排版标准生成完整的预览方案。
    ''' 遍历段落结构，为标题和正文分别生成 FormatChange 条目。
    ''' </summary>
    ''' <param name="analysis">DocumentAnalyzer 的分析结果</param>
    ''' <param name="standard">目标排版标准</param>
    ''' <param name="paragraphTexts">原始段落文本（可选，用于生成段落预览）</param>
    Public Function GeneratePreviewPlan(
        analysis As DocumentAnalysisResult,
        standard As FormattingStandard,
        Optional paragraphTexts As List(Of String) = Nothing) As ReformatPreviewPlan

        Dim plan As New ReformatPreviewPlan()
        If analysis Is Nothing OrElse standard Is Nothing Then Return plan

        plan.DetectedType = analysis.DocumentType
        plan.TypeConfidence = analysis.Confidence
        plan.StandardName = standard.Name
        plan.StandardDescription = standard.Description
        plan.TotalParagraphs = analysis.ParagraphCount
        plan.SemanticMapping = standard.SemanticMapping
        plan.PageSettings = standard.SemanticMapping.PageConfig

        ' --- 标题变更 ---
        Dim headingChanges = BuildHeadingChanges(analysis, standard, paragraphTexts)
        plan.Changes.AddRange(headingChanges)

        ' --- 正文变更（跳过已处理的标题段落） ---
        Dim bodyChanges = BuildBodyChanges(analysis, standard, paragraphTexts)
        plan.Changes.AddRange(bodyChanges)

        ' --- 去重（同一段落不出现两次） ---
        Dim seenIndices As New HashSet(Of Integer)()
        Dim distinctChanges As New List(Of FormatChange)()
        For Each ch In plan.Changes
            If seenIndices.Add(ch.ParagraphIndex) Then
                distinctChanges.Add(ch)
            End If
        Next
        plan.Changes = distinctChanges

        plan.TotalStyleChanges = plan.Changes.Count

        Return plan
    End Function

    ' ============================================================
    '  微调处理
    ' ============================================================

    ''' <summary>
    ''' 应用用户的自然语言微调指令（如"标题再大一点"、"正文用红色"）。
    ''' 直接修改 RefinementContext 中的 SemanticStyleMapping，然后返回更新后的预览方案。
    ''' </summary>
    ''' <param name="refinementCommand">用户微调指令</param>
    Public Function ApplyRefinement(refinementCommand As String) As ReformatPreviewPlan
        If String.IsNullOrWhiteSpace(refinementCommand) Then
            Return _refinementContext.CurrentPreviewPlan
        End If

        ' 记录命令到历史
        _refinementContext.AddConversation(refinementCommand)

        Dim command = refinementCommand.Trim()
        Dim mapping = _refinementContext.CurrentMapping
        If mapping Is Nothing Then Return Nothing

        ' 1. 确定要修改的目标标签
        Dim targetTags = GetTargetTags(command, mapping)

        ' 2. 应用微调
        ApplyTweakToTags(command, targetTags, _refinementContext)

        ' 3. 更新时间戳
        mapping.LastModified = DateTime.Now

        ' 4. 重新生成预览方案
        If _refinementContext.CurrentAnalysis IsNot Nothing AndAlso
           _refinementContext.CurrentStandard IsNot Nothing Then

            Dim updatedPlan = GeneratePreviewPlan(
                _refinementContext.CurrentAnalysis,
                _refinementContext.CurrentStandard)

            ' 覆盖为标准映射（微调后的映射才是实际要用的）
            updatedPlan.SemanticMapping = mapping

            _refinementContext.CurrentPreviewPlan = updatedPlan
            Return updatedPlan
        End If

        Return _refinementContext.CurrentPreviewPlan
    End Function

    ' ============================================================
    '  模式C：范文克隆
    ' ============================================================

    ''' <summary>
    ''' 比较源文档与范文的结构和段落差异，返回格式差异分析。
    ''' 源文档和范文需已通过 DocumentAnalyzer 完成分析。
    ''' 实际格式提取和映射生成由 FormatMirrorService 完成。
    ''' </summary>
    ''' <param name="sourceAnalysis">源文档分析结果</param>
    ''' <param name="mirrorAnalysis">范文分析结果</param>
    Public Function CompareWithMirror(
        sourceAnalysis As DocumentAnalysisResult,
        mirrorAnalysis As DocumentAnalysisResult) As FormatDiff

        Dim diff As New FormatDiff()
        If sourceAnalysis Is Nothing OrElse mirrorAnalysis Is Nothing Then Return diff

        diff.MirrorAnalysis = mirrorAnalysis

        ' 段落总数对比
        If sourceAnalysis.ParagraphCount <> mirrorAnalysis.ParagraphCount Then
            diff.TotalDifferences += 1
            diff.Differences.Add(New FormatDifference With {
                .ParagraphIndex = -1,
                .ParagraphPreview = "全局",
                .SourceFont = $"共{sourceAnalysis.ParagraphCount}段",
                .MirrorFont = $"共{mirrorAnalysis.ParagraphCount}段",
                .DifferenceDescription = $"段落数不一致：源文档{sourceAnalysis.ParagraphCount}段，范文{mirrorAnalysis.ParagraphCount}段"
            })
        End If

        ' 文档类型对比
        If sourceAnalysis.DocumentType <> mirrorAnalysis.DocumentType Then
            diff.TotalDifferences += 1
            diff.Differences.Add(New FormatDifference With {
                .ParagraphIndex = -1,
                .ParagraphPreview = "文档类型",
                .SourceFont = GetDocumentTypeDisplayName(sourceAnalysis.DocumentType),
                .MirrorFont = GetDocumentTypeDisplayName(mirrorAnalysis.DocumentType),
                .DifferenceDescription = $"文档类型不一致：{GetDocumentTypeDisplayName(sourceAnalysis.DocumentType)} vs {GetDocumentTypeDisplayName(mirrorAnalysis.DocumentType)}"
            })
        End If

        ' 标题结构对比
        If sourceAnalysis.DocStructure IsNot Nothing AndAlso mirrorAnalysis.DocStructure IsNot Nothing Then
            Dim srcHeadingCount = sourceAnalysis.DocStructure.Headings.Count
            Dim mirHeadingCount = mirrorAnalysis.DocStructure.Headings.Count

            If srcHeadingCount <> mirHeadingCount Then
                diff.TotalDifferences += 1
                diff.Differences.Add(New FormatDifference With {
                    .ParagraphIndex = -1,
                    .ParagraphPreview = "标题结构",
                    .SourceFont = $"{srcHeadingCount}个标题",
                    .MirrorFont = $"{mirHeadingCount}个标题",
                    .DifferenceDescription = $"标题数量不一致：源文档{srcHeadingCount}个，范文{mirHeadingCount}个"
                })
            End If

            ' 逐标题对比
            Dim maxHeadingCount = Math.Min(srcHeadingCount, mirHeadingCount)
            For i = 0 To maxHeadingCount - 1
                Dim src = sourceAnalysis.DocStructure.Headings(i)
                Dim mir = mirrorAnalysis.DocStructure.Headings(i)
                If src.Level <> mir.Level Then
                    diff.TotalDifferences += 1
                    diff.Differences.Add(New FormatDifference With {
                        .ParagraphIndex = src.ParagraphIndex,
                        .ParagraphPreview = TruncateText(src.Text, 50),
                        .SourceFont = $"H{src.Level}  {If(src.IsNumbered, "编号", "无编号")}",
                        .MirrorFont = $"H{mir.Level}  {If(mir.IsNumbered, "编号", "无编号")}",
                        .DifferenceDescription = $"标题层级不一致：源文档H{src.Level}，范文H{mir.Level}"
                    })
                End If
            Next
        End If

        ' 格式问题对比
        If sourceAnalysis.FormattingProblems.Count > 0 AndAlso
           mirrorAnalysis.FormattingProblems.Count = 0 Then
            diff.TotalDifferences += 1
            diff.Differences.Add(New FormatDifference With {
                .ParagraphIndex = -1,
                .ParagraphPreview = "格式质量",
                .SourceFont = $"{sourceAnalysis.FormattingProblems.Count}个问题",
                .MirrorFont = "无问题",
                .DifferenceDescription = "源文档存在格式问题，范文无问题"
            })
        End If

        Return diff
    End Function

    ''' <summary>消费者兼容方法：是否有活动的排版上下文</summary>
    Public Function HasActiveContext() As Boolean
        Return _refinementContext.CurrentPreviewPlan IsNot Nothing AndAlso Not _refinementContext.IsApplied
    End Function

    ''' <summary>消费者兼容方法：一键速排</summary>
    Public Async Function QuickReformatAsync(paragraphs As List(Of String),
                                              wordParagraphs As List(Of Object)) As Task(Of ReformatPreviewPlan)
        Dim plan = AnalyzeAndRecommend(paragraphs)
        _refinementContext.CurrentPreviewPlan = plan
        _refinementContext.IsApplied = False
        plan.NeedsAITagging = True
        plan.DocumentTypeName = GetDocumentTypeName(plan.DetectedType)
        Return plan
    End Function

    ''' <summary>消费者兼容方法：对话式排版（意图驱动）</summary>
    Public Async Function ChatReformatAsync(userMessage As String,
                                             paragraphs As List(Of String),
                                             wordParagraphs As List(Of Object)) As Task(Of ReformatPreviewPlan)

        ' 解析用户意图
        Dim analysis = If(_refinementContext.CurrentAnalysis, New DocumentAnalysisResult())
        Dim intent = ParseUserIntent(userMessage, analysis)

        Dim plan As ReformatPreviewPlan

        Select Case intent.IntentType
            Case IntentType.SpecificTweak
                ' 微调模式：在当前映射基础上做增量修改，不需要重新AI标注
                plan = ApplyRefinement(userMessage)
                If plan Is Nothing Then
                    ' 没有活动上下文，降级为自动排版
                    plan = AnalyzeAndRecommend(paragraphs)
                End If

            Case IntentType.StandardFormat
                ' 指定标准排版：用用户指定的标准（如"按公文排版"→GB/T 9704-2012）
                Dim standard As FormattingStandard = Nothing
                If Not String.IsNullOrEmpty(intent.TargetStandardName) Then
                    standard = _knowledgeEngine.GetStandardByName(intent.TargetStandardName)
                End If

                If standard IsNot Nothing Then
                    ' 先做文档分析获取结构，再用指定标准生成预览
                    Dim docAnalysis = _analyzer.Analyze(paragraphs)
                    plan = GeneratePreviewPlan(docAnalysis, standard, paragraphs)
                    ' 用户明确指定了标准时，用用户意图覆盖分析器猜测的文档类型
                    If intent.TargetDocumentType <> DocumentType.Unknown Then
                        plan.DetectedType = intent.TargetDocumentType
                    End If
                    plan.DocumentTypeName = standard.Name
                    _refinementContext.Update(docAnalysis, plan.SemanticMapping, standard, plan)
                Else
                    ' 标准未找到，降级为自动排版
                    plan = AnalyzeAndRecommend(paragraphs)
                End If

            Case IntentType.StyleClone
                ' 格式克隆：由上层HandleMirrorFormat处理，此处降级为自动排版
                plan = AnalyzeAndRecommend(paragraphs)

            Case IntentType.FormatCleanup
                ' 格式清洗：先自动分析，清洗逻辑在渲染时处理
                plan = AnalyzeAndRecommend(paragraphs)

            Case Else
                ' AutoFormat 或 Unknown：自动检测文档类型并推荐标准
                plan = AnalyzeAndRecommend(paragraphs)
        End Select

        _refinementContext.CurrentPreviewPlan = plan
        _refinementContext.IsApplied = False
        plan.NeedsAITagging = True
        plan.DocumentTypeName = GetDocumentTypeName(plan.DetectedType)
        Return plan
    End Function

    ''' <summary>消费者兼容方法：对话式排版（意图驱动，增强版）</summary>
    Public Async Function ChatReformatAsync(userMessage As String,
                                             paragraphs As List(Of String),
                                             wordParagraphs As List(Of Object),
                                             paragraphStyles As List(Of String),
                                             paragraphFontSizes As List(Of Single),
                                             paragraphIsBold As List(Of Boolean)) As Task(Of ReformatPreviewPlan)

        ' 解析用户意图
        Dim analysis = If(_refinementContext.CurrentAnalysis, New DocumentAnalysisResult())
        Dim intent = ParseUserIntent(userMessage, analysis)

        Dim plan As ReformatPreviewPlan

        Select Case intent.IntentType
            Case IntentType.SpecificTweak
                plan = ApplyRefinement(userMessage)
                If plan Is Nothing Then
                    plan = AnalyzeAndRecommend(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
                End If

            Case IntentType.StandardFormat
                Dim standard As FormattingStandard = Nothing
                If Not String.IsNullOrEmpty(intent.TargetStandardName) Then
                    standard = _knowledgeEngine.GetStandardByName(intent.TargetStandardName)
                End If
                If standard IsNot Nothing Then
                    Dim docAnalysis = _analyzer.Analyze(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
                    plan = GeneratePreviewPlan(docAnalysis, standard, paragraphs)
                    ' 用户明确指定了标准时，用用户意图覆盖分析器猜测的文档类型
                    If intent.TargetDocumentType <> DocumentType.Unknown Then
                        plan.DetectedType = intent.TargetDocumentType
                    End If
                    plan.DocumentTypeName = standard.Name
                    _refinementContext.Update(docAnalysis, plan.SemanticMapping, standard, plan)
                Else
                    plan = AnalyzeAndRecommend(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
                End If

            Case Else
                plan = AnalyzeAndRecommend(paragraphs, paragraphStyles, paragraphFontSizes, paragraphIsBold)
        End Select

        _refinementContext.CurrentPreviewPlan = plan
        _refinementContext.IsApplied = False
        plan.NeedsAITagging = True
        plan.DocumentTypeName = GetDocumentTypeName(plan.DetectedType)
        Return plan
    End Function

    ''' <summary>获取文档类型的中文名称</summary>
    Private Function GetDocumentTypeName(docType As DocumentType) As String
        Select Case docType
            Case DocumentType.OfficialDocument : Return "行政公文"
            Case DocumentType.AcademicPaper : Return "学术论文"
            Case DocumentType.BusinessReport : Return "商业报告"
            Case DocumentType.Contract : Return "合同协议"
            Case DocumentType.[Resume] : Return "个人简历"
            Case DocumentType.GeneralDocument : Return "通用文档"
            Case Else : Return "未知"
        End Select
    End Function

    Public Sub ResetRefinement()
        _refinementContext.Clear()
    End Sub

    ' ============================================================
    '  内部辅助 —— 生成变更
    ' ============================================================

    ''' <summary>根据标题结构生成格式变更条目</summary>
    Private Function BuildHeadingChanges(
        analysis As DocumentAnalysisResult,
        standard As FormattingStandard,
        paragraphTexts As List(Of String)) As List(Of FormatChange)

        Dim changes As New List(Of FormatChange)()
        If analysis?.DocStructure Is Nothing Then Return changes

        For Each heading In analysis.DocStructure.Headings
            Dim change As New FormatChange()
            change.ParagraphIndex = heading.ParagraphIndex
            change.ParagraphPreview = TruncateText(heading.Text, 50)
            change.OldTag = $"heading.{heading.Level}"
            change.NewTag = GetStandardTagForHeading(heading.Level, standard)

            Dim semanticTag = standard?.SemanticMapping?.FindTag(change.NewTag)
            If semanticTag IsNot Nothing Then
                ResolveChangeFromTag(change, semanticTag)
            End If

            changes.Add(change)
        Next

        Return changes
    End Function

    ''' <summary>根据正文段落列表生成格式变更条目</summary>
    Private Function BuildBodyChanges(
        analysis As DocumentAnalysisResult,
        standard As FormattingStandard,
        paragraphTexts As List(Of String)) As List(Of FormatChange)

        Dim changes As New List(Of FormatChange)()
        If analysis?.DocStructure Is Nothing OrElse paragraphTexts Is Nothing Then Return changes

        ' 收集所有非正文段落索引（标题/列表/表格）
        Dim excludeIndices As New HashSet(Of Integer)()
        For Each h In analysis.DocStructure.Headings
            excludeIndices.Add(h.ParagraphIndex)
        Next
        For Each idx In analysis.DocStructure.ListParagraphIndices
            excludeIndices.Add(idx)
        Next
        For Each idx In analysis.DocStructure.TableParagraphIndices
            excludeIndices.Add(idx)
        Next

        ' 每个正文段落尝试映射到 body.normal
        Dim bodyTag = standard?.SemanticMapping?.FindTag("body.normal")

        For i = 0 To paragraphTexts.Count - 1
            Dim text = paragraphTexts(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For
            If excludeIndices.Contains(i) Then Continue For

            Dim change As New FormatChange()
            change.ParagraphIndex = i
            change.ParagraphPreview = TruncateText(text, 50)
            change.OldTag = "body.normal"
            change.NewTag = "" ' 待AI语义标注确定

            If bodyTag IsNot Nothing Then
                ResolveChangeFromTag(change, bodyTag)
            End If

            changes.Add(change)
        Next

        Return changes
    End Function

    ''' <summary>将语义标签的格式信息填入 FormatChange</summary>
    Private Shared Sub ResolveChangeFromTag(change As FormatChange, tag As SemanticTag)
        If tag Is Nothing Then Return

        change.NewFont = FormatFontDescription(tag.Font)
        change.NewAlignment = GetAlignmentDisplayName(tag.Paragraph?.Alignment)

        If tag.Paragraph IsNot Nothing AndAlso tag.Paragraph.FirstLineIndent > 0 Then
            change.NewIndent = $"首行缩进{tag.Paragraph.FirstLineIndent}字符"
        End If

        change.ChangeDescription = BuildFormatDescription(tag)
    End Sub

    ''' <summary>获取标准中对应标题级别的最佳语义标签ID，逐级回退</summary>
    Private Shared Function GetStandardTagForHeading(level As Integer, standard As FormattingStandard) As String
        If standard?.SemanticMapping Is Nothing Then Return $"heading.{level}"

        Dim tagId = $"heading.{level}"
        If standard.SemanticMapping.FindTag(tagId) IsNot Nothing Then
            Return tagId
        End If

        ' 逐级回退
        For fallback = level - 1 To 1 Step -1
            Dim fallbackId = $"heading.{fallback}"
            If standard.SemanticMapping.FindTag(fallbackId) IsNot Nothing Then
                Return fallbackId
            End If
        Next

        Return "body.normal"
    End Function

    ' ============================================================
    '  内部辅助 —— 微调
    ' ============================================================

    ''' <summary>根据命令语义确定要修改的目标标签列表</summary>
    Private Shared Function GetTargetTags(command As String, mapping As SemanticStyleMapping) As List(Of SemanticTag)
        Dim tags As New List(Of SemanticTag)()
        If mapping Is Nothing Then Return tags

        Dim hasHeading = ContainsAny(command, {"标题", "题目", "章", "节"})
        Dim hasBody = ContainsAny(command, {"正文", "内容", "段落", "文字"})
        Dim hasAll = ContainsAny(command, {"全部", "所有", "整体", "全局"})

        If hasAll OrElse (Not hasHeading AndAlso Not hasBody) Then
            tags.AddRange(mapping.SemanticTags)
        Else
            If hasHeading Then
                tags.AddRange(mapping.SemanticTags.Where(Function(t) t.TagId.StartsWith("heading") OrElse
                                                                     t.TagId.StartsWith("title")))
            End If
            If hasBody Then
                tags.AddRange(mapping.SemanticTags.Where(Function(t) t.TagId.StartsWith("body")))
            End If
        End If

        ' 无命中时退回到全部标签
        If tags.Count = 0 Then
            tags.AddRange(mapping.SemanticTags)
        End If

        Return tags
    End Function

    ''' <summary>对目标标签应用实际的格式数值修改</summary>
    Private Shared Sub ApplyTweakToTags(command As String, targetTags As List(Of SemanticTag), context As RefinementContext)
        If String.IsNullOrWhiteSpace(command) OrElse targetTags Is Nothing Then Return

        ' 判断是否为重复操作（"再大一点" → 沿用上一次操作）
        Dim isRepeat = ContainsAny(command, {"再", "更", "还", "继续", "进一步"})

        For Each tag In targetTags
            ' ---- 字号 ----
            If Regex.IsMatch(command, "(大|增大|加大|放大).{0,6}(点|一些|一点|字号|字体)") OrElse
               (isRepeat AndAlso Regex.IsMatch(command, "大.{0,3}(点|一些)")) Then
                tag.Font.FontSize = Math.Min(tag.Font.FontSize + 1, 72)
            End If

            If Regex.IsMatch(command, "(小|减小|缩小|调小).{0,6}(点|一些|一点|字号|字体)") OrElse
               (isRepeat AndAlso Regex.IsMatch(command, "小.{0,3}(点|一些)")) Then
                tag.Font.FontSize = Math.Max(tag.Font.FontSize - 1, 8)
            End If

            ' 直接指定字号
            Dim sizeMatch = Regex.Match(command, "(\d+)\s*(pt|磅|号)")
            If sizeMatch.Success Then
                tag.Font.FontSize = Double.Parse(sizeMatch.Groups(1).Value)
            End If

            ' ---- 颜色 ----
            If command.Contains("红色") Then
                tag.Color.FontColor = "#C00000"
            ElseIf command.Contains("蓝色") Then
                tag.Color.FontColor = "#2E5090"
            ElseIf command.Contains("黑色") Then
                tag.Color.FontColor = "#000000"
            ElseIf command.Contains("绿色") Then
                tag.Color.FontColor = "#008000"
            End If

            ' ---- 加粗 ----
            If ContainsAny(command, {"加粗", "粗体", "粗一点"}) Then
                tag.Font.Bold = True
            End If
            If ContainsAny(command, {"取消加粗", "不加粗", "细体"}) Then
                tag.Font.Bold = False
            End If

            ' ---- 对齐 ----
            If ContainsAny(command, {"居中", "居中对齐"}) Then
                tag.Paragraph.Alignment = "center"
            ElseIf ContainsAny(command, {"左对齐", "靠左"}) Then
                tag.Paragraph.Alignment = "left"
            ElseIf ContainsAny(command, {"右对齐", "靠右"}) Then
                tag.Paragraph.Alignment = "right"
            ElseIf ContainsAny(command, {"两端对齐"}) Then
                tag.Paragraph.Alignment = "justify"
            End If

            ' ---- 行距 ----
            Dim lineSpacingMatch = Regex.Match(command, "行距.{0,4}([\d.]+)")
            If lineSpacingMatch.Success Then
                Dim newSpacing As Double
                If Double.TryParse(lineSpacingMatch.Groups(1).Value, newSpacing) Then
                    tag.Paragraph.LineSpacing = Math.Max(0.5, Math.Min(newSpacing, 3.0))
                End If
            End If

            ' ---- 缩进 ----
            If command.Contains("缩进") Then
                Dim indentMatch = Regex.Match(command, "缩进.{0,4}([\d.]+)")
                If indentMatch.Success Then
                    Dim newIndent As Double
                    If Double.TryParse(indentMatch.Groups(1).Value, newIndent) Then
                        tag.Paragraph.FirstLineIndent = newIndent
                    End If
                Else
                    tag.Paragraph.FirstLineIndent = 2
                End If
            End If

            ' ---- 字体 ----
            If command.Contains("宋体") Then
                tag.Font.FontNameCN = "宋体"
            ElseIf command.Contains("仿宋") Then
                tag.Font.FontNameCN = "仿宋_GB2312"
            ElseIf command.Contains("黑体") Then
                tag.Font.FontNameCN = "黑体"
            ElseIf command.Contains("楷体") Then
                tag.Font.FontNameCN = "楷体_GB2312"
            ElseIf command.Contains("微软雅黑") Then
                tag.Font.FontNameCN = "微软雅黑"
            End If
        Next
    End Sub

    ' ============================================================
    '  内部辅助 —— 格式化与工具
    ' ============================================================

    ''' <summary>根据语义标签构建格式变更的文字描述</summary>
    Private Shared Function BuildFormatDescription(tag As SemanticTag) As String
        Dim parts As New List(Of String)()

        If tag.Font IsNot Nothing Then
            parts.Add(FormatFontDescription(tag.Font))
        End If

        If tag.Paragraph IsNot Nothing Then
            Dim alignName = GetAlignmentDisplayName(tag.Paragraph.Alignment)
            If Not String.IsNullOrEmpty(alignName) Then
                parts.Add(alignName)
            End If
            If tag.Paragraph.FirstLineIndent > 0 Then
                parts.Add($"首行缩进{tag.Paragraph.FirstLineIndent}字符")
            End If
            If tag.Paragraph.LineSpacing > 0 AndAlso
               Math.Abs(tag.Paragraph.LineSpacing - 1.5) > 0.01 Then
                parts.Add($"行距{tag.Paragraph.LineSpacing}")
            End If
        End If

        Return String.Join(", ", parts)
    End Function

    ''' <summary>格式化字体描述字符串</summary>
    Private Shared Function FormatFontDescription(font As FontConfig) As String
        If font Is Nothing Then Return ""
        Dim parts As New List(Of String)()

        If Not String.IsNullOrEmpty(font.FontNameCN) Then
            parts.Add(font.FontNameCN)
        End If
        If font.FontSize > 0 Then
            parts.Add($"{font.FontSize}pt")
        End If
        If font.Bold Then
            parts.Add("加粗")
        End If
        If font.Italic Then
            parts.Add("斜体")
        End If

        Return String.Join(" ", parts)
    End Function

    ''' <summary>获取对齐方式的中文显示名称</summary>
    Private Shared Function GetAlignmentDisplayName(alignment As String) As String
        If String.IsNullOrEmpty(alignment) Then Return ""
        Select Case alignment.ToLower()
            Case "center" : Return "居中"
            Case "right" : Return "右对齐"
            Case "justify" : Return "两端对齐"
            Case Else : Return "左对齐"
        End Select
    End Function

    ''' <summary>截断文本并在末尾加省略号</summary>
    Private Shared Function TruncateText(text As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "…"
    End Function

    ''' <summary>获取文档类型的中文名称</summary>
    Private Shared Function GetDocumentTypeDisplayName(docType As DocumentType) As String
        Select Case docType
            Case DocumentType.OfficialDocument : Return "公文"
            Case DocumentType.AcademicPaper : Return "学术论文"
            Case DocumentType.BusinessReport : Return "商业报告"
            Case DocumentType.Contract : Return "合同"
            Case DocumentType.[Resume] : Return "简历"
            Case DocumentType.GeneralDocument : Return "通用文档"
            Case Else : Return "未知"
        End Select
    End Function

    ''' <summary>检查文本是否包含任一关键词（忽略大小写）</summary>
    Private Shared Function ContainsAny(text As String, keywords As String()) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        For Each kw In keywords
            If text.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>判断消息是否表达微调意图</summary>
    Private Shared Function HasTweakIntent(message As String) As Boolean
        Return Regex.IsMatch(message, "(再|更|调|改|设|换).{0,4}(大|小|颜色|字体|行距|对齐|加粗|缩进)") OrElse
               Regex.IsMatch(message, "(大|小|颜色|字体|行距|对齐|加粗|缩进).{0,4}(点|一些|一点)") OrElse
               ContainsAny(message, {"红色", "蓝色", "黑色", "居中", "加粗"})
    End Function

End Class
