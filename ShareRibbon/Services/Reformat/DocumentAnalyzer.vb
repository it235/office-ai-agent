' ShareRibbon\Services\Reformat\DocumentAnalyzer.vb
' 文档类型识别、结构解析、格式问题检测

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Linq
Imports System.Diagnostics

''' <summary>
''' 文档类型枚举
''' </summary>
Public Enum DocumentType
    Unknown = 0
    ''' <summary>公文（党政机关公文）</summary>
    OfficialDocument
    ''' <summary>学术论文</summary>
    AcademicPaper
    ''' <summary>商业报告</summary>
    BusinessReport
    ''' <summary>合同协议</summary>
    Contract
    ''' <summary>个人简历</summary>
    [Resume]
    ''' <summary>通用文档</summary>
    GeneralDocument
End Enum

''' <summary>
''' 格式问题严重程度
''' </summary>
Public Enum ProblemSeverity
    ''' <summary>提示</summary>
    Suggestion
    ''' <summary>警告</summary>
    Warning
    ''' <summary>严重错误</summary>
    [Error]
End Enum

''' <summary>
''' 文档分析结果
''' </summary>
Public Class DocumentAnalysisResult
    ''' <summary>识别出的文档类型</summary>
    Public Property DocumentType As DocumentType = DocumentType.Unknown
    ''' <summary>类型识别置信度 0-1</summary>
    Public Property Confidence As Double = 0.0
    ''' <summary>文档结构信息</summary>
    Public Property DocStructure As DocumentStructure
    ''' <summary>检测到的格式问题列表</summary>
    Public Property FormattingProblems As List(Of FormattingProblem)
    ''' <summary>推荐的模板ID列表</summary>
    Public Property RecommendedTemplateIds As List(Of String)
    ''' <summary>段落总数</summary>
    Public Property ParagraphCount As Integer = 0
    ''' <summary>是否包含目录</summary>
    Public Property HasTableOfContents As Boolean = False
    ''' <summary>分析用时（毫秒）</summary>
    Public Property AnalysisTimeMs As Long = 0
    ''' <summary>是否使用了Word富文本信息增强分析</summary>
    Public Property EnhancedWithWordFormatting As Boolean = False

    Public Sub New()
        DocStructure = New DocumentStructure()
        FormattingProblems = New List(Of FormattingProblem)()
        RecommendedTemplateIds = New List(Of String)()
    End Sub
End Class

''' <summary>
''' 文档结构信息
''' </summary>
Public Class DocumentStructure
    ''' <summary>标题层级列表，每项为 {level, text, paragraphIndex}</summary>
    Public Property Headings As List(Of HeadingInfo)
    ''' <summary>正文段落索引范围</summary>
    Public Property BodyParagraphRanges As List(Of ParagraphRange)
    ''' <summary>列表数量</summary>
    Public Property ListCount As Integer = 0
    ''' <summary>表格数量</summary>
    Public Property TableCount As Integer = 0
    ''' <summary>图片数量</summary>
    Public Property ImageCount As Integer = 0
    ''' <summary>列表项所在的段落索引</summary>
    Public Property ListParagraphIndices As List(Of Integer)
    ''' <summary>表格所在的段落索引</summary>
    Public Property TableParagraphIndices As List(Of Integer)
    ''' <summary>总字符数</summary>
    Public Property TotalCharCount As Integer = 0
    ''' <summary>平均段落长度（字符数）</summary>
    Public Property AverageParagraphLength As Double = 0.0

    Public Sub New()
        Headings = New List(Of HeadingInfo)()
        BodyParagraphRanges = New List(Of ParagraphRange)()
        ListParagraphIndices = New List(Of Integer)()
        TableParagraphIndices = New List(Of Integer)()
    End Sub
End Class

''' <summary>
''' 标题信息
''' </summary>
Public Class HeadingInfo
    ''' <summary>标题级别（1-6）</summary>
    Public Property Level As Integer = 1
    ''' <summary>标题文本</summary>
    Public Property Text As String = ""
    ''' <summary>在段落列表中的索引</summary>
    Public Property ParagraphIndex As Integer = -1
    ''' <summary>是否为编号标题（如"一、""1.1"）</summary>
    Public Property IsNumbered As Boolean = False
End Class

''' <summary>
''' 段落索引范围
''' </summary>
Public Class ParagraphRange
    ''' <summary>起始索引（含）</summary>
    Public Property StartIndex As Integer = -1
    ''' <summary>结束索引（含）</summary>
    Public Property EndIndex As Integer = -1
End Class

''' <summary>
''' 格式问题
''' </summary>
Public Class FormattingProblem
    ''' <summary>问题描述</summary>
    Public Property Description As String = ""
    ''' <summary>严重程度</summary>
    Public Property Severity As ProblemSeverity = ProblemSeverity.Suggestion
    ''' <summary>问题所在的段落索引（-1表示全局问题）</summary>
    Public Property ParagraphIndex As Integer = -1
    ''' <summary>建议的修复方式</summary>
    Public Property SuggestedFix As String = ""
    ''' <summary>问题类别：font/spacing/style/structure</summary>
    Public Property Category As String = ""
End Class

''' <summary>
''' 文档分析器 - 负责文档类型识别、结构解析和格式问题检测
''' </summary>
Public Class DocumentAnalyzer

    ' ---- 公文关键词 ----
    Private Shared ReadOnly OfficialDocKeywords As String() = {
        "发文机关", "发文字号", "通知", "决定", "批复", "请示", "报告",
        "函", "纪要", "印发", "转发", "抄送", "主送", "主题词",
        "〔", "〕", "机密", "特急", "急件", "签发人", "会签"
    }

    ' ---- 学术论文关键词 ----
    Private Shared ReadOnly AcademicKeywords As String() = {
        "摘要", "关键词", "参考文献", "Abstract", "引言", "绪论",
        "结论", "致谢", "附录", "文献综述", "研究方法", "实验",
        "结果分析", "讨论", "数据来源", "基金项目", "作者简介",
        "DOI", "中图分类号"
    }

    ' ---- 商业报告关键词 ----
    Private Shared ReadOnly BusinessKeywords As String() = {
        "项目", "季度", "年度", "汇报", "总结", "分析", "预算",
        "营收", "利润", "增长率", "市场份额", "KPI", "指标",
        "目标", "战略", "方案", "建议", "风险评估", "里程碑",
        "交付物", "干系人", "ROI"
    }

    ' ---- 合同关键词 ----
    Private Shared ReadOnly ContractKeywords As String() = {
        "甲方", "乙方", "丙方", "合同", "协议", "条款", "签署",
        "盖章", "生效", "违约", "赔偿", "保密", "仲裁", "管辖",
        "权利义务", "不可抗力", "争议解决", "定金", "首款",
        "尾款", "服务期限", "知识产权"
    }

    ' ---- 简历关键词 ----
    Private Shared ReadOnly ResumeKeywords As String() = {
        "工作经历", "教育背景", "专业技能", "项目经验", "自我评价",
        "个人简介", "职业技能", "实习经历", "证书", "语言能力",
        "兴趣爱好", "求职意向", "期望薪资", "学历"
    }

    ' ---- 标题编号模式 ----
    Private Shared ReadOnly HeadingNumberPatterns As Regex() = {
        New Regex("^一[、.．]", RegexOptions.Compiled),                     ' 一、
        New Regex("^二[、.．]", RegexOptions.Compiled),                     ' 二、
        New Regex("^三[、.．]", RegexOptions.Compiled),                     ' 三、
        New Regex("^四[、.．]", RegexOptions.Compiled),                     ' 四、
        New Regex("^五[、.．]", RegexOptions.Compiled),                     ' 五、
        New Regex("^六[、.．]", RegexOptions.Compiled),                     ' 六、
        New Regex("^七[、.．]", RegexOptions.Compiled),                     ' 七、
        New Regex("^八[、.．]", RegexOptions.Compiled),                     ' 八、
        New Regex("^九[、.．]", RegexOptions.Compiled),                     ' 九、
        New Regex("^十[、.．]", RegexOptions.Compiled),                     ' 十、
        New Regex("^第[一二三四五六七八九十]+[、.．]", RegexOptions.Compiled), ' 第一、
        New Regex("^\d+[.．、]", RegexOptions.Compiled),                     ' 1.  1、
        New Regex("^\d+\.\d+", RegexOptions.Compiled),                      ' 1.1
        New Regex("^\d+\.\d+\.\d+", RegexOptions.Compiled),                 ' 1.1.1
        New Regex("^[（(]\d+[）)]", RegexOptions.Compiled),                  ' (1) （1）
        New Regex("^[①-⑩]", RegexOptions.Compiled)                          ' ①-⑩
    }

    Private ReadOnly _textAnalyzer As Func(Of String, String, Task(Of String))

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <param name="textAnalyzer">可选的外部AI文本分析委托，用于复杂文档类型的LLM辅助判断</param>
    Public Sub New(Optional textAnalyzer As Func(Of String, String, Task(Of String)) = Nothing)
        _textAnalyzer = textAnalyzer
    End Sub

    ''' <summary>
    ''' 对段落列表执行完整分析
    ''' </summary>
    ''' <param name="paragraphs">文档段落列表</param>
    ''' <param name="fullText">可选全文文本</param>
    Public Function Analyze(paragraphs As List(Of String),
                            Optional fullText As String = Nothing) As DocumentAnalysisResult
        Dim sw As New Stopwatch()
        sw.Restart()

        If paragraphs Is Nothing OrElse paragraphs.Count = 0 Then
            Return New DocumentAnalysisResult()
        End If

        If String.IsNullOrEmpty(fullText) Then
            fullText = String.Join(vbCrLf, paragraphs)
        End If

        Dim result As New DocumentAnalysisResult()
        result.ParagraphCount = paragraphs.Count

        ' 1. 类型识别
        Dim typeResult = DetectDocumentType(paragraphs, fullText)
        result.DocumentType = typeResult.Item1
        result.Confidence = typeResult.Item2

        ' 2. 结构解析
        result.DocStructure = AnalyzeStructure(paragraphs)
        result.HasTableOfContents = DetectTableOfContents(paragraphs)

        ' 3. 格式问题检测
        result.FormattingProblems = DetectFormattingProblems(paragraphs)

        ' 4. 推荐模板
        result.RecommendedTemplateIds = GetRecommendedTemplateIds(result.DocumentType, result.DocStructure)

        sw.Stop()
        result.AnalysisTimeMs = sw.ElapsedMilliseconds

        Return result
    End Function

    ''' <summary>
    ''' 对段落列表执行完整分析（增强版，接受Word富文本格式信息）
    ''' </summary>
    ''' <param name="paragraphs">文档段落文本列表</param>
    ''' <param name="paragraphStyles">段落样式名称列表（如"标题 1","Heading 1","正文"）</param>
    ''' <param name="paragraphFontSizes">段落字号列表（pt）</param>
    ''' <param name="paragraphIsBold">段落是否加粗列表</param>
    ''' <param name="fullText">可选全文文本</param>
    Public Function Analyze(paragraphs As List(Of String),
                            paragraphStyles As List(Of String),
                            paragraphFontSizes As List(Of Single),
                            paragraphIsBold As List(Of Boolean),
                            Optional fullText As String = Nothing) As DocumentAnalysisResult

        ' 先用纯文本分析获取基础结果
        Dim result = Analyze(paragraphs, fullText)
        If paragraphStyles Is Nothing OrElse paragraphs.Count = 0 Then Return result

        ' 用Word样式名补充标题检测
        Dim existingHeadingIndices As New HashSet(Of Integer)
        For Each h In result.DocStructure.Headings
            existingHeadingIndices.Add(h.ParagraphIndex)
        Next

        For i = 0 To Math.Min(paragraphs.Count - 1, paragraphStyles.Count - 1) - 1
            If existingHeadingIndices.Contains(i) Then Continue For
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            ' 从样式名提取标题级别
            Dim styleLevel = ParseHeadingLevelFromStyle(paragraphStyles(i))
            If styleLevel > 0 Then
                result.DocStructure.Headings.Add(New HeadingInfo With {
                    .Level = styleLevel,
                    .Text = text.Trim(),
                    .ParagraphIndex = i,
                    .IsNumbered = IsNumberedHeading(text)
                })
                existingHeadingIndices.Add(i)
                Continue For
            End If

            ' 从字号+加粗推断标题（大字+加粗+短段落=标题）
            If paragraphFontSizes IsNot Nothing AndAlso i < paragraphFontSizes.Count AndAlso
               paragraphIsBold IsNot Nothing AndAlso i < paragraphIsBold.Count Then
                Dim fontSize = paragraphFontSizes(i)
                Dim isBold = paragraphIsBold(i)
                Dim trimmedLen = text.Trim().Length

                If fontSize > 16 AndAlso isBold AndAlso trimmedLen > 0 AndAlso trimmedLen <= 80 Then
                    ' 根据字号推断级别：>=22pt→1级，>=18pt→2级，>16pt→3级
                    Dim inferredLevel = If(fontSize >= 22, 1, If(fontSize >= 18, 2, 3))
                    result.DocStructure.Headings.Add(New HeadingInfo With {
                        .Level = inferredLevel,
                        .Text = text.Trim(),
                        .ParagraphIndex = i,
                        .IsNumbered = IsNumberedHeading(text)
                    })
                    existingHeadingIndices.Add(i)
                End If
            End If
        Next

        ' 重新排序标题列表
        result.DocStructure.Headings = result.DocStructure.Headings.OrderBy(Function(h) h.ParagraphIndex).ToList()

        result.EnhancedWithWordFormatting = True
        Return result
    End Function

    ''' <summary>
    ''' 从Word样式名解析标题级别（如"标题 1"→1, "Heading 2"→2, "标题 3"→3）
    ''' </summary>
    Public Shared Function ParseHeadingLevelFromStyle(styleName As String) As Integer
        If String.IsNullOrEmpty(styleName) Then Return 0

        ' 中文样式名：标题 1, 标题 2, 标题 1 Char
        Dim cnMatch = System.Text.RegularExpressions.Regex.Match(styleName, "标题\s*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If cnMatch.Success Then Return Integer.Parse(cnMatch.Groups(1).Value)

        ' 英文样式名：Heading 1, Heading 2
        Dim enMatch = System.Text.RegularExpressions.Regex.Match(styleName, "Heading\s*(\d+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If enMatch.Success Then Return Integer.Parse(enMatch.Groups(1).Value)

        Return 0
    End Function

    ''' <summary>
    ''' 使用LLM辅助分析复杂文档（当规则判断置信度低时调用）
    ''' </summary>
    ''' <param name="paragraphs">文档段落列表</param>
    ''' <param name="fullText">全文文本</param>
    Public Async Function AnalyzeWithLLM(paragraphs As List(Of String),
                                         Optional fullText As String = Nothing) As Task(Of DocumentAnalysisResult)
        Dim result = Analyze(paragraphs, fullText)

        ' 如果置信度足够高，直接返回规则分析结果
        If result.Confidence >= 0.7 AndAlso _textAnalyzer Is Nothing Then
            Return result
        End If

        ' 如果置信度低且有AI分析委托，使用LLM辅助
        If _textAnalyzer IsNot Nothing AndAlso result.Confidence < 0.7 Then
            Try
                Dim sampleText = BuildAnalyzerPrompt(paragraphs)
                Dim llmResponse = Await _textAnalyzer("document_analysis", sampleText)

                ' 解析LLM返回的类型判断
                Dim llmType = ParseLlmTypeResponse(llmResponse)
                If llmType.Item1 <> DocumentType.Unknown Then
                    result.DocumentType = llmType.Item1
                    result.Confidence = Math.Max(result.Confidence, llmType.Item2)
                End If

                ' 如果LLM提供了额外的格式问题，合并到结果中
                Dim llmProblems = ParseLlmProblems(llmResponse)
                If llmProblems IsNot Nothing AndAlso llmProblems.Count > 0 Then
                    result.FormattingProblems.AddRange(llmProblems)
                End If
            Catch ex As Exception
                ' LLM调用失败时使用规则分析结果
            End Try
        End If

        Return result
    End Function

    ' ============================================================
    '  类型识别
    ' ============================================================

    ''' <summary>
    ''' 基于关键词和结构特征检测文档类型
    ''' </summary>
    Private Function DetectDocumentType(paragraphs As List(Of String),
                                        fullText As String) As Tuple(Of DocumentType, Double)
        Dim scores As New Dictionary(Of DocumentType, Double)

        scores(DocumentType.OfficialDocument) = ScoreOfficialDocument(paragraphs, fullText)
        scores(DocumentType.AcademicPaper) = ScoreAcademicPaper(paragraphs, fullText)
        scores(DocumentType.BusinessReport) = ScoreBusinessReport(paragraphs, fullText)
        scores(DocumentType.Contract) = ScoreContract(paragraphs, fullText)
        scores(DocumentType.[Resume]) = ScoreResume(paragraphs)

        ' 找到最高分
        Dim bestType = DocumentType.Unknown
        Dim bestScore = 0.0

        For Each kvp In scores
            If kvp.Value > bestScore Then
                bestScore = kvp.Value
                bestType = kvp.Key
            End If
        Next

        ' 如果分数低于阈值，降级为通用或未知
        If bestScore < 0.15 Then
            ' 检查是否有足够的正文内容
            Dim hasContent = paragraphs.Any(Function(p) p.Trim().Length > 20)
            If hasContent Then
                bestType = DocumentType.GeneralDocument
                bestScore = 0.5
            Else
                bestType = DocumentType.Unknown
                bestScore = 0.0
            End If
        End If

        Return Tuple.Create(bestType, bestScore)
    End Function

    ''' <summary>
    ''' 计算公文得分
    ''' </summary>
    Private Function ScoreOfficialDocument(paragraphs As List(Of String),
                                           fullText As String) As Double
        Dim score As Double = 0.0
        Dim matchedCount = 0

        ' 关键词匹配
        For Each keyword In OfficialDocKeywords
            If fullText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                matchedCount += 1
                score += 0.08
            End If
        Next

        ' 发文字号模式：〔2024〕或 〔2025〕等
        If Regex.IsMatch(fullText, "〔\d{4}〕") Then
            score += 0.3
        End If

        ' 红色横线特征（公文常见分隔线）
        If paragraphs.Any(Function(p) p.Trim() = "──" OrElse
                                       p.Trim().All(Function(c) c = "─"c OrElse c = "—"c) AndAlso
                                       p.Trim().Length >= 5) Then
            score += 0.15
        End If

        ' 模板头特征：顶格主送机关
        If paragraphs.Any(Function(p) p.Trim().StartsWith("各") AndAlso
                                       (p.Contains("单位") OrElse p.Contains("部门") OrElse
                                        p.Contains("机构") OrElse p.Contains("处室"))) Then
            score += 0.15
        End If

        ' 落款特征
        If paragraphs.Any(Function(p) p.Trim().EndsWith("办公室") OrElse
                                       p.Trim().EndsWith("局") OrElse
                                       p.Trim().EndsWith("委员会") OrElse
                                       p.Trim().EndsWith("部")) AndAlso
           paragraphs.Any(Function(p) Regex.IsMatch(p.Trim(), "^\d{4}年\d{1,2}月\d{1,2}日$")) Then
            score += 0.2
        End If

        ' 关键词覆盖度修正
        Dim keywordRatio = matchedCount / OfficialDocKeywords.Length
        If keywordRatio > 0.3 Then
            score += 0.15
        End If

        Return Math.Min(score, 1.0)
    End Function

    ''' <summary>
    ''' 计算学术论文得分
    ''' </summary>
    Private Function ScoreAcademicPaper(paragraphs As List(Of String),
                                        fullText As String) As Double
        Dim score As Double = 0.0
        Dim matchedCount = 0

        For Each keyword In AcademicKeywords
            If fullText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                matchedCount += 1
                score += 0.08
            End If
        Next

        ' 摘要特征：段落以"摘要"开头
        If paragraphs.Any(Function(p) p.Trim().StartsWith("摘要") OrElse
                                       p.Trim().StartsWith("摘　要")) Then
            score += 0.15
        End If

        ' 关键词特征
        If paragraphs.Any(Function(p) p.Trim().StartsWith("关键词") OrElse
                                       p.Trim().StartsWith("关键字")) Then
            score += 0.15
        End If

        ' 参考文献特征
        Dim refLines = paragraphs.Where(Function(p) Regex.IsMatch(p.Trim(), "^\[\d+\]")).ToList()
        If refLines.Count >= 3 Then
            score += 0.25
        End If

        ' 标题编号特征（多级编号如 1.1, 1.1.1）
        Dim numberedHeadingCount = 0
        For Each p In paragraphs
            If Regex.IsMatch(p.Trim(), "^\d+\.\d+\.?\d*") Then
                numberedHeadingCount += 1
            End If
        Next
        If numberedHeadingCount >= 3 Then
            score += 0.2
        End If

        ' 中英文摘要并存
        If fullText.IndexOf("Abstract", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
           fullText.IndexOf("摘要") >= 0 Then
            score += 0.15
        End If

        ' 关键词覆盖度修正
        Dim keywordRatio = matchedCount / AcademicKeywords.Length
        If keywordRatio > 0.25 Then
            score += 0.1
        End If

        Return Math.Min(score, 1.0)
    End Function

    ''' <summary>
    ''' 计算商业报告得分
    ''' </summary>
    Private Function ScoreBusinessReport(paragraphs As List(Of String),
                                         fullText As String) As Double
        Dim score As Double = 0.0
        Dim matchedCount = 0

        For Each keyword In BusinessKeywords
            If fullText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                matchedCount += 1
                score += 0.06
            End If
        Next

        ' 数据表格特征（含百分比、数字、万元等）
        Dim dataLineCount = 0
        For Each p In paragraphs
            If Regex.IsMatch(p, "\d+%") AndAlso (p.Contains("同比") OrElse p.Contains("环比") OrElse
                                                 p.Contains("增长") OrElse p.Contains("下降")) Then
                dataLineCount += 1
            End If
        Next
        If dataLineCount >= 3 Then
            score += 0.2
        End If

        ' 报告标题特征
        If paragraphs.Any(Function(p) (p.Contains("报告") OrElse p.Contains("汇报")) AndAlso
                                       (p.Contains("年度") OrElse p.Contains("季度") OrElse
                                        p.Contains("月度") OrElse p.Contains("工作") OrElse
                                        p.Contains("项目"))) Then
            score += 0.15
        End If

        ' 目录特征
        If paragraphs.Any(Function(p) p.Trim() = "目录" OrElse p.Trim() = "CONTENTS" OrElse
                                       p.Trim() = "目　录") Then
            score += 0.1
        End If

        Return Math.Min(score, 1.0)
    End Function

    ''' <summary>
    ''' 计算合同得分
    ''' </summary>
    Private Function ScoreContract(paragraphs As List(Of String),
                                   fullText As String) As Double
        Dim score As Double = 0.0
        Dim matchedCount = 0

        For Each keyword In ContractKeywords
            If fullText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                matchedCount += 1
                score += 0.08
            End If
        Next

        ' 甲方乙方同时出现
        If fullText.IndexOf("甲方") >= 0 AndAlso fullText.IndexOf("乙方") >= 0 Then
            score += 0.25
        End If

        ' 条款式结构：第X条
        Dim clauseCount = 0
        For Each p In paragraphs
            If Regex.IsMatch(p.Trim(), "^第[一二三四五六七八九十百千]+条") Then
                clauseCount += 1
            End If
        Next
        If clauseCount >= 3 Then
            score += 0.3
        End If

        ' 签署区域特征
        If paragraphs.Any(Function(p) p.Trim().Contains("盖章") OrElse
                                       p.Trim().Contains("签字") OrElse
                                       p.Trim().Contains("签署")) Then
            score += 0.1
        End If

        Return Math.Min(score, 1.0)
    End Function

    ''' <summary>
    ''' 计算简历得分
    ''' </summary>
    Private Function ScoreResume(paragraphs As List(Of String)) As Double
        Dim score As Double = 0.0
        Dim matchedCount = 0

        For Each keyword In ResumeKeywords
            If paragraphs.Any(Function(p) p.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0) Then
                matchedCount += 1
                score += 0.1
            End If
        Next

        ' 总段落数少（简历通常较短）
        If paragraphs.Count >= 5 AndAlso paragraphs.Count <= 30 Then
            score += 0.1
        End If

        ' 第一行可能是姓名
        If paragraphs.Count > 0 Then
            Dim firstLine = paragraphs(0).Trim()
            If firstLine.Length <= 10 AndAlso Not firstLine.Contains(" ") AndAlso
               Not firstLine.Contains("　") AndAlso firstLine.Length > 0 Then
                ' 姓名行通常短且无标点
                If Not firstLine.Any(Function(c) Char.IsPunctuation(c) AndAlso c <> "·") Then
                    score += 0.1
                End If
            End If
        End If

        ' 联系方式特征
        If paragraphs.Any(Function(p) Regex.IsMatch(p, "1[3-9]\d{9}") OrElse  ' 手机号
                                       Regex.IsMatch(p, "[\w\.-]+@[\w\.-]+\.\w+")) Then  ' 邮箱
            score += 0.1
        End If

        ' 分隔线特征（简历常用分隔线）
        If paragraphs.Any(Function(p) p.Trim().All(Function(c) c = "─"c OrElse c = "—"c OrElse
                                                              c = "-"c OrElse c = "="c) AndAlso
                                       p.Trim().Length >= 3) Then
            score += 0.05
        End If

        Return Math.Min(score, 1.0)
    End Function

    ' ============================================================
    '  结构解析
    ' ============================================================

    ''' <summary>
    ''' 解析文档结构——标题层级、正文范围、列表和表格
    ''' </summary>
    Private Function AnalyzeStructure(paragraphs As List(Of String)) As DocumentStructure
        Dim docStructure As New DocumentStructure()

        If paragraphs Is Nothing OrElse paragraphs.Count = 0 Then
            Return docStructure
        End If

        ' 计算总字符数和平均段落长度
        docStructure.TotalCharCount = paragraphs.Sum(Function(p) p.Length)
        docStructure.AverageParagraphLength = If(paragraphs.Count > 0,
            docStructure.TotalCharCount / paragraphs.Count, 0.0)

        Dim headingIndex = 0

        For i = 0 To paragraphs.Count - 1
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            ' 检测标题
            Dim headingLevel = DetectHeadingLevel(text, i, paragraphs)
            If headingLevel > 0 Then
                Dim isNumbered = IsNumberedHeading(text)
                docStructure.Headings.Add(New HeadingInfo With {
                    .Level = headingLevel,
                    .Text = text.Trim(),
                    .ParagraphIndex = i,
                    .IsNumbered = isNumbered
                })
                headingIndex += 1
            End If

            ' 检测列表项
            If IsListItem(text) Then
                docStructure.ListParagraphIndices.Add(i)
                docStructure.ListCount += 1
            End If

            ' 检测表格（简化检测：含制表符或连续空格分隔的结构化内容）
            If IsTableLine(text) Then
                docStructure.TableParagraphIndices.Add(i)
                docStructure.TableCount += 1
            End If
        Next

        ' 识别正文区域（非标题、非列表的连续段落）
        Dim bodyStart = -1
        For i = 0 To paragraphs.Count - 1
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            Dim isHeading = docStructure.Headings.Any(Function(h) h.ParagraphIndex = i)
            Dim isList = docStructure.ListParagraphIndices.Contains(i)
            Dim isTable = docStructure.TableParagraphIndices.Contains(i)

            If Not isHeading AndAlso Not isList AndAlso Not isTable Then
                If bodyStart = -1 Then
                    bodyStart = i
                End If
            Else
                If bodyStart >= 0 Then
                    docStructure.BodyParagraphRanges.Add(New ParagraphRange With {
                        .StartIndex = bodyStart,
                        .EndIndex = i - 1
                    })
                    bodyStart = -1
                End If
            End If
        Next

        ' 最后一个正文区域
        If bodyStart >= 0 Then
            docStructure.BodyParagraphRanges.Add(New ParagraphRange With {
                .StartIndex = bodyStart,
                .EndIndex = paragraphs.Count - 1
            })
        End If

        Return docStructure
    End Function

    ''' <summary>
    ''' 检测段落是否为标题并返回级别
    ''' </summary>
    Private Function DetectHeadingLevel(text As String,
                                        index As Integer,
                                        paragraphs As List(Of String)) As Integer
        Dim trimmed = text.Trim()

        ' 空行或太短
        If String.IsNullOrEmpty(trimmed) OrElse trimmed.Length <= 1 Then
            Return 0
        End If

        ' 太长的文本不太可能是标题
        If trimmed.Length > 100 Then
            Return 0
        End If

        ' 一级标题：大编号（一、二、三、第一章等）
        If Regex.IsMatch(trimmed, "^第[一二三四五六七八九十百千]+[章节篇]") Then
            Return 1
        End If
        If Regex.IsMatch(trimmed, "^(一|二|三|四|五|六|七|八|九|十)[、.．]") AndAlso
           Not Regex.IsMatch(trimmed, "^[一二三四五六七八九十]+[、.．].*[的得地]") Then
            Return 1
        End If
        If Regex.IsMatch(trimmed, "^(前言|引言|绪论|摘要|Abstract|参考文献|附录|致谢|后记|目录)") Then
            Return 1
        End If

        ' 二级标题：1.1  1.2  等
        If Regex.IsMatch(trimmed, "^\d+\.\d+\s") OrElse
           Regex.IsMatch(trimmed, "^\d+\.\d+$") Then
            Return 2
        End If

        ' 三级标题：1.1.1
        If Regex.IsMatch(trimmed, "^\d+\.\d+\.\d+") Then
            Return 3
        End If

        ' 短句标题（15字以内，无句号结尾，位于正文段前）
        If trimmed.Length <= 30 AndAlso
           Not trimmed.EndsWith("。") AndAlso
           Not trimmed.EndsWith("；") AndAlso
           Not trimmed.EndsWith("，") AndAlso
           Not trimmed.EndsWith("、") Then
            ' 相邻段落非空时，该行若较短可能是标题
            If trimmed.Length <= 20 AndAlso
               index > 0 AndAlso String.IsNullOrWhiteSpace(paragraphs(index - 1)) AndAlso
               index < paragraphs.Count - 1 AndAlso
               Not String.IsNullOrWhiteSpace(paragraphs(index + 1)) Then
                Return 3
            End If
        End If

        Return 0
    End Function

    ''' <summary>
    ''' 判断是否为编号标题
    ''' </summary>
    Private Function IsNumberedHeading(text As String) As Boolean
        Dim trimmed = text.Trim()
        For Each pattern In HeadingNumberPatterns
            If pattern.IsMatch(trimmed) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' 判断是否为列表项
    ''' </summary>
    Private Function IsListItem(text As String) As Boolean
        Dim trimmed = text.Trim()
        If String.IsNullOrEmpty(trimmed) Then Return False

        ' 编号列表：1. 2. (1) ① 等
        If Regex.IsMatch(trimmed, "^\d+[.．、)]\s") Then Return True
        If Regex.IsMatch(trimmed, "^[（(]\d+[）)]\s") Then Return True
        If Regex.IsMatch(trimmed, "^[①-⑩]") Then Return True
        If Regex.IsMatch(trimmed, "^[•·●◆※]\s") Then Return True
        If Regex.IsMatch(trimmed, "^[a-zA-Z][.．)]\s") Then Return True

        ' 破折号或星号开头的列表项
        If Regex.IsMatch(trimmed, "^-{1,2}\s") Then Return True
        If Regex.IsMatch(trimmed, "^\*\s") Then Return True

        Return False
    End Function

    ''' <summary>
    ''' 判断是否为表格行
    ''' </summary>
    Private Function IsTableLine(text As String) As Boolean
        Dim trimmed = text.Trim()
        If String.IsNullOrEmpty(trimmed) Then Return False

        ' 含制表符
        If trimmed.Contains(vbTab) Then Return True

        ' 含多个竖线（Markdown表格或文本表格）
        Dim pipeCount = trimmed.Count(Function(c) c = "|"c)
        If pipeCount >= 2 Then Return True

        ' 连续空格分隔的结构化内容（至少3列）
        If Regex.IsMatch(trimmed, "\S+\s{3,}\S+\s{3,}\S+") Then Return True

        Return False
    End Function

    ''' <summary>
    ''' 检测是否包含目录
    ''' </summary>
    Private Function DetectTableOfContents(paragraphs As List(Of String)) As Boolean
        ' 显式目录标题
        If paragraphs.Any(Function(p) p.Trim() = "目录" OrElse
                                       p.Trim() = "目　录" OrElse
                                       p.Trim() = "CONTENTS" OrElse
                                       p.Trim() = "目次" OrElse
                                       p.Trim() = "TOC") Then
            Return True
        End If

        ' 目录特征：连续多行包含"…"或"..."连接标题和页码
        Dim dotLeaderCount = 0
        For i = 0 To paragraphs.Count - 1
            If paragraphs(i).Contains("…") OrElse
               paragraphs(i).Contains("...") OrElse
               Regex.IsMatch(paragraphs(i), "\w+[\.]{3,}\d+\s*$") OrElse
               Regex.IsMatch(paragraphs(i), "\w+\s+[\.]{2,}\s+\d+\s*$") Then
                dotLeaderCount += 1
            End If
        Next

        Return dotLeaderCount >= 3
    End Function

    ' ============================================================
    '  格式问题检测
    ' ============================================================

    ''' <summary>
    ''' 检测文档中的格式问题
    ''' </summary>
    Private Function DetectFormattingProblems(paragraphs As List(Of String)) As List(Of FormattingProblem)
        Dim problems As New List(Of FormattingProblem)()

        ' 此方法检测纯文本层面的格式问题，
        ' 字体/字号等富文本格式需要外部传入富文本信息
        ' 此处做文本结构层面的问题检测

        ' 1. 标题样式缺失检测
        Dim hasNumberedHeading = False
        Dim hasProperHeading = False
        For i = 0 To paragraphs.Count - 1
            If IsNumberedHeading(paragraphs(i)) Then
                hasNumberedHeading = True
                ' 检查编号标题前面是否有空行
                If i > 0 AndAlso Not String.IsNullOrWhiteSpace(paragraphs(i - 1)) Then
                    problems.Add(New FormattingProblem With {
                        .Description = $"标题「{paragraphs(i).Trim()}」前缺少空行，标题应前后留白",
                        .Severity = ProblemSeverity.Suggestion,
                        .ParagraphIndex = i,
                        .Category = "spacing",
                        .SuggestedFix = "在标题前插入空行"
                    })
                End If
            End If
        Next

        ' 2. 检测正文缩进问题（段落开头含多余空格）
        For i = 0 To paragraphs.Count - 1
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            ' 检测半角空格开头的段落（不规范缩进）
            If text.StartsWith(" ") AndAlso text.Length > 1 Then
                Dim spaceCount = text.TakeWhile(Function(c) c = " "c).Count()
                If spaceCount > 0 AndAlso spaceCount Mod 2 <> 0 Then
                    problems.Add(New FormattingProblem With {
                        .Description = $"段落含有不规范缩进（{spaceCount}个半角空格），建议使用首行缩进2字符",
                        .Severity = ProblemSeverity.Warning,
                        .ParagraphIndex = i,
                        .Category = "spacing",
                        .SuggestedFix = "删除行首空格，设置段落首行缩进2字符"
                    })
                    Continue For
                End If
            End If
        Next

        ' 3. 检测中英文混排间距问题
        For i = 0 To paragraphs.Count - 1
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            ' 中文后接英文字母/数字但无空格
            If Regex.IsMatch(text, "[一-鿿][a-zA-Z0-9]") Then
                problems.Add(New FormattingProblem With {
                    .Description = "中英文之间缺少空格，影响阅读体验",
                    .Severity = ProblemSeverity.Suggestion,
                    .ParagraphIndex = i,
                    .Category = "spacing",
                    .SuggestedFix = "在中英文之间添加半角空格"
                })
            End If
        Next

        ' 4. 检测全角半角混用（常见于中英文混排）
        For i = 0 To paragraphs.Count - 1
            Dim text = paragraphs(i)
            If String.IsNullOrWhiteSpace(text) Then Continue For

            ' 英文内容中出现全角字母
            If Regex.IsMatch(text, "[！-～]") AndAlso
               Regex.IsMatch(text, "[a-zA-Z]{3,}") Then
                problems.Add(New FormattingProblem With {
                    .Description = "检测到全角符号与英文混用，建议统一为半角符号",
                    .Severity = ProblemSeverity.Warning,
                    .ParagraphIndex = i,
                    .Category = "style",
                    .SuggestedFix = "将全角英文字母和数字转换为半角"
                })
                Exit For
            End If
        Next

        ' 5. 检测过长的段落（缺少分段）
        For i = 0 To paragraphs.Count - 1
            If paragraphs(i).Length > 500 Then
                problems.Add(New FormattingProblem With {
                    .Description = $"第{i + 1}段过长（{paragraphs(i).Length}字符），建议适当分段",
                    .Severity = ProblemSeverity.Suggestion,
                    .ParagraphIndex = i,
                    .Category = "structure",
                    .SuggestedFix = "在适当位置断句分段"
                })
            End If
        Next

        ' 6. 检测连续的短段落（可能应为列表）
        Dim consecutiveShortCount = 0
        For i = 0 To paragraphs.Count - 1
            Dim text = paragraphs(i).Trim()
            If text.Length > 0 AndAlso text.Length < 30 AndAlso
               Not IsListItem(text) AndAlso
               Not paragraphs(i).Contains("：") AndAlso
               Not paragraphs(i).Contains(":") Then
                consecutiveShortCount += 1
            Else
                If consecutiveShortCount >= 3 Then
                    problems.Add(New FormattingProblem With {
                        .Description = $"检测到连续{consecutiveShortCount}个短段落，可能应为列表格式",
                        .Severity = ProblemSeverity.Suggestion,
                        .ParagraphIndex = i - consecutiveShortCount,
                        .Category = "structure",
                        .SuggestedFix = "考虑转换为项目符号列表"
                    })
                End If
                consecutiveShortCount = 0
            End If
        Next

        Return problems
    End Function

    ' ============================================================
    '  模板推荐
    ' ============================================================

    ''' <summary>
    ''' 根据文档类型和结构推荐模板ID
    ''' </summary>
    Private Function GetRecommendedTemplateIds(docType As DocumentType,
                                               docStructure As DocumentStructure) As List(Of String)
        Dim ids As New List(Of String)()

        Select Case docType
            Case DocumentType.OfficialDocument
                ids.Add("gbt9704-2012")
                ids.Add("official-document-simple")
            Case DocumentType.AcademicPaper
                ids.Add("academic-paper-generic")
                If docStructure.Headings.Count >= 3 Then
                    ids.Add("academic-paper-structured")
                End If
            Case DocumentType.BusinessReport
                ids.Add("business-report-generic")
                If docStructure.TableCount > 0 Then
                    ids.Add("business-report-data")
                End If
            Case DocumentType.Contract
                ids.Add("contract-generic")
            Case DocumentType.[Resume]
                ids.Add("resume-generic")
                ids.Add("resume-modern")
            Case Else
                ids.Add("general-document")
        End Select

        Return ids
    End Function

    ' ============================================================
    '  LLM 辅助
    ' ============================================================

    ''' <summary>
    ''' 构建LLM分析用的提示词
    ''' </summary>
    Private Function BuildAnalyzerPrompt(paragraphs As List(Of String)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("请分析以下文档内容，判断文档类型。可选类型：")
        sb.AppendLine("- 公文（党政机关公文）")
        sb.AppendLine("- 学术论文")
        sb.AppendLine("- 商业报告")
        sb.AppendLine("- 合同协议")
        sb.AppendLine("- 个人简历")
        sb.AppendLine("- 通用文档")
        sb.AppendLine()
        sb.AppendLine("请以JSON格式返回分析结果，格式：")
        sb.AppendLine("{""type"": ""公文/论文/报告/合同/简历/通用"", ""confidence"": 0.95}")
        sb.AppendLine()
        sb.AppendLine("文档内容（前50段采样）：")
        sb.AppendLine()

        Dim sampleCount = Math.Min(paragraphs.Count, 50)
        For i = 0 To sampleCount - 1
            Dim text = paragraphs(i).Trim()
            If Not String.IsNullOrEmpty(text) Then
                sb.AppendLine($"[{i + 1}] {text}")
            End If
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 解析LLM返回的类型判断
    ''' </summary>
    Private Function ParseLlmTypeResponse(response As String) As Tuple(Of DocumentType, Double)
        Try
            Dim json = Newtonsoft.Json.Linq.JObject.Parse(response)
            Dim typeStr = If(json("type")?.ToString(), "")
            Dim confidenceObj = json("confidence")?.ToObject(Of Double)()
            Dim confidence = If(confidenceObj, 0.0)

            Dim docType As DocumentType
            Select Case typeStr
                Case "公文"
                    docType = DocumentType.OfficialDocument
                Case "论文"
                    docType = DocumentType.AcademicPaper
                Case "报告"
                    docType = DocumentType.BusinessReport
                Case "合同"
                    docType = DocumentType.Contract
                Case "简历"
                    docType = DocumentType.[Resume]
                Case "通用"
                    docType = DocumentType.GeneralDocument
                Case Else
                    docType = DocumentType.Unknown
            End Select

            Return Tuple.Create(docType, confidence)
        Catch ex As Exception
            Return Tuple.Create(DocumentType.Unknown, 0.0)
        End Try
    End Function

    ''' <summary>
    ''' 解析LLM返回的额外格式问题
    ''' </summary>
    Private Function ParseLlmProblems(response As String) As List(Of FormattingProblem)
        Try
            Dim json = Newtonsoft.Json.Linq.JObject.Parse(response)
            If json("problems") IsNot Nothing Then
                Dim problems = json("problems").ToObject(Of List(Of FormattingProblem))()
                Return problems
            End If
        Catch ex As Exception
        End Try
        Return Nothing
    End Function

End Class
