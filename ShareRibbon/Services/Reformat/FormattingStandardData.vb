' ShareRibbon\Services\Reformat\FormattingStandardData.vb
' 内置排版标准数据

Imports System.Collections.Generic

''' <summary>
''' 内置排版标准数据提供器 - 创建各排版标准的 SemanticStyleMapping 数据
''' </summary>
Public Module FormattingStandardData

    ''' <summary>
    ''' 获取所有内置排版标准
    ''' </summary>
    Public Function GetAllBuiltInStandards() As List(Of FormattingStandard)
        Return New List(Of FormattingStandard) From {
            CreateGbt9704Standard(),
            CreateGbt7714Standard(),
            CreateAcademicStandard(),
            CreateBusinessStandard()
        }
    End Function

#Region "GB/T 9704-2012 党政机关公文格式"

    ''' <summary>
    ''' 创建 GB/T 9704-2012 党政机关公文格式标准
    ''' </summary>
    Private Function CreateGbt9704Standard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "gbt-9704-2012"
        standard.Name = "GB/T 9704-2012"
        standard.Description = "党政机关公文格式国家标准，适用于各级党政机关制发的公文。包含页面设置、发文机关标志、发文字号、标题、正文、附件、署名、成文日期、页码等完整格式规范。"
        standard.ApplicableDocumentTypes.Add(DocumentType.OfficialDocument.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateGbt9704Mapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建 GB/T 9704-2012 的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateGbt9704Mapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "gbt-9704-2012"
        mapping.Name = "GB/T 9704-2012 党政机关公文格式"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' ============ 页面设置 ============
        mapping.PageConfig.Margins = New MarginsConfig(3.7, 3.5, 2.8, 2.6)

        ' ============ 语义标签 ============
        Dim tags = mapping.SemanticTags

        ' 1. 发文机关标志
        Dim headerOrg As New SemanticTag("header.org", "发文机关标志", "header", 2, "发文机关名称或规范化简称")
        headerOrg.Font = New FontConfig("方正小标宋简体", "", 22, True)
        headerOrg.Paragraph = New ParagraphConfig("center")
        headerOrg.Color = New ColorConfig("#C00000")
        tags.Add(headerOrg)

        ' 2. 红色分隔线
        Dim headerSep As New SemanticTag("header.separator", "红色分隔线", "header", 2, "发文字号下方的红色分隔横线")
        headerSep.Color = New ColorConfig("#C00000")
        tags.Add(headerSep)

        ' 3. 发文字号
        Dim headerRefno As New SemanticTag("header.refno", "发文字号", "header", 2, "发文字号，含机关代字、年份和顺序号，如""×发〔2026〕×号""")
        headerRefno.Font = New FontConfig("仿宋_GB2312", "", 16)
        headerRefno.Paragraph = New ParagraphConfig("center")
        tags.Add(headerRefno)

        ' 4. 文件标题
        Dim titleMain As New SemanticTag("title.main", "文件标题", "title", 2, "公文主标题，居中排列，多行时用菱形排法")
        titleMain.Font = New FontConfig("方正小标宋简体", "", 22, True)
        titleMain.Paragraph = New ParagraphConfig("center")
        titleMain.Paragraph.SpaceAfter = 1
        tags.Add(titleMain)

        ' 5. 主送机关
        Dim titleRecipient As New SemanticTag("title.recipient", "主送机关", "title", 2, "主送机关名称，左对齐顶格，多个用顿号或逗号分隔")
        titleRecipient.Font = New FontConfig("仿宋_GB2312", "", 16)
        titleRecipient.Paragraph = New ParagraphConfig("left")
        tags.Add(titleRecipient)

        ' 6. 一级标题（黑体）
        Dim heading1 As New SemanticTag("heading.1", "一级标题", "", 2, "公文一级标题，如""一、""""二、""等")
        heading1.Font = New FontConfig("黑体", "", 16, True)
        heading1.Paragraph = New ParagraphConfig("left")
        heading1.Paragraph.SpaceBefore = 0.5
        heading1.Paragraph.SpaceAfter = 0.5
        tags.Add(heading1)

        ' 7. 二级标题（楷体）
        Dim heading2 As New SemanticTag("heading.2", "二级标题", "", 2, "公文二级标题，如""(一)""""(二)""等")
        heading2.Font = New FontConfig("楷体_GB2312", "", 16, True)
        heading2.Paragraph = New ParagraphConfig("left")
        tags.Add(heading2)

        ' 8. 三级标题（仿宋加粗）
        Dim heading3 As New SemanticTag("heading.3", "三级标题", "", 2, "公文三级标题编号格式")
        heading3.Font = New FontConfig("仿宋_GB2312", "", 16, True)
        heading3.Paragraph = New ParagraphConfig("left")
        tags.Add(heading3)

        ' 9. 正文
        Dim bodyNormal As New SemanticTag("body.normal", "正文", "body", 2, "公文正文段落，每面22行每行28字")
        bodyNormal.Font = New FontConfig("仿宋_GB2312", "", 16)
        bodyNormal.Paragraph = New ParagraphConfig("justify", 2, 1.875)
        tags.Add(bodyNormal)

        ' 10. 附件说明
        Dim bodyAttachment As New SemanticTag("body.attachment", "附件说明", "body", 2, "附件说明，如""附件：1. ×××""")
        bodyAttachment.Font = New FontConfig("仿宋_GB2312", "", 16)
        bodyAttachment.Paragraph = New ParagraphConfig("left")
        bodyAttachment.Paragraph.SpaceBefore = 1
        tags.Add(bodyAttachment)

        ' 11. 发文机关署名
        Dim footerSignature As New SemanticTag("footer.signature", "发文机关署名", "footer", 2, "发文机关署名，位于成文日期之上")
        footerSignature.Font = New FontConfig("仿宋_GB2312", "", 16)
        footerSignature.Paragraph = New ParagraphConfig("right")
        tags.Add(footerSignature)

        ' 12. 成文日期
        Dim footerDate As New SemanticTag("footer.date", "成文日期", "footer", 2, "成文日期，使用阿拉伯数字，不编虚位")
        footerDate.Font = New FontConfig("仿宋_GB2312", "", 16)
        footerDate.Paragraph = New ParagraphConfig("right")
        tags.Add(footerDate)

        ' 13. 页码
        Dim footerPage As New SemanticTag("footer.page", "页码", "footer", 2, "公文页码，4号半角宋体")
        footerPage.Font = New FontConfig("宋体", "", 10)
        footerPage.Paragraph = New ParagraphConfig("center")
        tags.Add(footerPage)

        ' ============ 版式骨架 ============
        Dim layout As New LayoutConfig()

        Dim el1 As New LayoutElement()
        el1.Name = "发文机关标志"
        el1.ElementType = "text"
        el1.SortOrder = 1
        el1.Required = True
        el1.PlaceholderContent = "{{发文机关}}文件"
        layout.Elements.Add(el1)

        Dim el2 As New LayoutElement()
        el2.Name = "红色横线"
        el2.ElementType = "redLine"
        el2.SortOrder = 2
        el2.Required = True
        el2.Color = New ColorConfig("#C00000")
        el2.SpecialProps("lineWidth") = "2pt"
        layout.Elements.Add(el2)

        Dim el3 As New LayoutElement()
        el3.Name = "发文字号"
        el3.ElementType = "text"
        el3.SortOrder = 3
        el3.Required = True
        el3.PlaceholderContent = "×发〔2026〕×号"
        layout.Elements.Add(el3)

        Dim el4 As New LayoutElement()
        el4.Name = "文件标题"
        el4.ElementType = "text"
        el4.SortOrder = 4
        el4.Required = True
        layout.Elements.Add(el4)

        Dim el5 As New LayoutElement()
        el5.Name = "主送机关"
        el5.ElementType = "text"
        el5.SortOrder = 5
        el5.Required = True
        el5.PlaceholderContent = "各有关单位："
        layout.Elements.Add(el5)

        mapping.LayoutSkeleton = layout

        Return mapping
    End Function

#End Region

#Region "GB/T 7714-2015 参考文献著录规则"

    ''' <summary>
    ''' 创建 GB/T 7714-2015 参考文献著录规则标准
    ''' </summary>
    Private Function CreateGbt7714Standard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "gbt-7714-2015"
        standard.Name = "GB/T 7714-2015"
        standard.Description = "信息与文献 参考文献著录规则国家标准，规定了学术论文、著作中参考文献的著录项目、著录顺序、著录符号、著录用文字及著录格式。"
        standard.ApplicableDocumentTypes.Add(DocumentType.AcademicPaper.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateGbt7714Mapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建 GB/T 7714-2015 的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateGbt7714Mapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "gbt-7714-2015"
        mapping.Name = "GB/T 7714-2015 参考文献著录规则"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' 语义标签
        Dim tags = mapping.SemanticTags

        ' 参考文献标题
        Dim titleRef As New SemanticTag("title.references", "参考文献标题", "title", 2, "参考文献章节标题，如""参考文献""")
        titleRef.Font = New FontConfig("黑体", "", 14, True)
        titleRef.Paragraph = New ParagraphConfig("left")
        titleRef.Paragraph.SpaceBefore = 1
        titleRef.Paragraph.SpaceAfter = 0.5
        tags.Add(titleRef)

        ' 参考文献条目
        Dim bodyRef As New SemanticTag("body.reference", "参考文献条目", "body", 2, "参考文献列表中的单条文献记录")
        bodyRef.Font = New FontConfig("宋体", "Times New Roman", 9)
        bodyRef.Paragraph = New ParagraphConfig("justify")
        bodyRef.Paragraph.LineSpacing = 1.25
        tags.Add(bodyRef)

        Return mapping
    End Function

#End Region

#Region "学术论文通用格式"

    ''' <summary>
    ''' 创建学术论文通用格式标准（基于 GB/T 7713.1）
    ''' </summary>
    Private Function CreateAcademicStandard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "academic-general"
        standard.Name = "学术论文通用格式"
        standard.Description = "适用于学术论文、学位论文的通用格式规范。包含标题、摘要、关键词、章节标题、正文、参考文献等要素的标准格式。"
        standard.ApplicableDocumentTypes.Add(DocumentType.AcademicPaper.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateAcademicMapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建学术论文通用格式的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateAcademicMapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "academic-general"
        mapping.Name = "学术论文通用格式"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' 页面设置
        mapping.PageConfig.Margins = New MarginsConfig(2.54, 2.54, 3.18, 3.18)

        ' 语义标签
        Dim tags = mapping.SemanticTags

        ' 论文标题
        Dim titleMain As New SemanticTag("title.main", "论文标题", "title", 2, "学术论文主标题")
        titleMain.Font = New FontConfig("黑体", "Times New Roman", 18, True)
        titleMain.Paragraph = New ParagraphConfig("center", 0, 1.5)
        titleMain.Paragraph.SpaceBefore = 2
        titleMain.Paragraph.SpaceAfter = 1
        tags.Add(titleMain)

        ' 摘要标题
        Dim titleAbstract As New SemanticTag("title.abstract", "摘要标题", "title", 2, """摘要""标题行")
        titleAbstract.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        titleAbstract.Paragraph = New ParagraphConfig("left")
        titleAbstract.Paragraph.SpaceBefore = 1
        titleAbstract.Paragraph.SpaceAfter = 0.5
        tags.Add(titleAbstract)

        ' 摘要正文
        Dim bodyAbstract As New SemanticTag("body.abstract", "摘要正文", "body", 2, "摘要内容段落")
        bodyAbstract.Font = New FontConfig("宋体", "Times New Roman", 12)
        bodyAbstract.Paragraph = New ParagraphConfig("justify", 2, 1.5)
        tags.Add(bodyAbstract)

        ' 关键词标题
        Dim titleKeywords As New SemanticTag("title.keywords", "关键词标题", "title", 2, """关键词""标识行")
        titleKeywords.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        titleKeywords.Paragraph = New ParagraphConfig("left")
        titleKeywords.Paragraph.SpaceBefore = 0.5
        tags.Add(titleKeywords)

        ' 关键词
        Dim bodyKeywords As New SemanticTag("body.keywords", "关键词", "body", 2, "关键词列表，分号分隔")
        bodyKeywords.Font = New FontConfig("宋体", "Times New Roman", 12)
        bodyKeywords.Paragraph = New ParagraphConfig("left")
        bodyKeywords.Paragraph.SpaceAfter = 1
        tags.Add(bodyKeywords)

        ' 一级标题（章）
        Dim heading1 As New SemanticTag("heading.1", "一级标题", "", 2, "论文章标题，如""第1章 引言""")
        heading1.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        heading1.Paragraph = New ParagraphConfig("left")
        heading1.Paragraph.SpaceBefore = 1
        heading1.Paragraph.SpaceAfter = 0.5
        tags.Add(heading1)

        ' 二级标题（节）
        Dim heading2 As New SemanticTag("heading.2", "二级标题", "", 2, "论文节标题，如""1.1 研究背景""")
        heading2.Font = New FontConfig("黑体", "Times New Roman", 12, True)
        heading2.Paragraph = New ParagraphConfig("left")
        heading2.Paragraph.SpaceBefore = 0.5
        heading2.Paragraph.SpaceAfter = 0.25
        tags.Add(heading2)

        ' 正文
        Dim bodyNormal As New SemanticTag("body.normal", "正文", "body", 2, "学术论文正文段落")
        bodyNormal.Font = New FontConfig("宋体", "Times New Roman", 12)
        bodyNormal.Paragraph = New ParagraphConfig("justify", 2, 1.5)
        tags.Add(bodyNormal)

        ' 参考文献标题
        Dim titleRef As New SemanticTag("title.references", "参考文献标题", "title", 2, """参考文献""章节标题")
        titleRef.Font = New FontConfig("黑体", "Times New Roman", 14, True)
        titleRef.Paragraph = New ParagraphConfig("left")
        titleRef.Paragraph.SpaceBefore = 1
        titleRef.Paragraph.SpaceAfter = 0.5
        tags.Add(titleRef)

        ' 参考文献条目
        Dim bodyRef As New SemanticTag("body.reference", "参考文献条目", "body", 2, "参考文献列表中的单条记录")
        bodyRef.Font = New FontConfig("宋体", "Times New Roman", 10)
        bodyRef.Paragraph = New ParagraphConfig("justify")
        bodyRef.Paragraph.LineSpacing = 1.25
        tags.Add(bodyRef)

        ' 页码
        Dim footerPage As New SemanticTag("footer.page", "页码", "footer", 2, "论文页码")
        footerPage.Font = New FontConfig("宋体", "Times New Roman", 10)
        footerPage.Paragraph = New ParagraphConfig("center")
        tags.Add(footerPage)

        Return mapping
    End Function

#End Region

#Region "商务报告通用规范"

    ''' <summary>
    ''' 创建商务报告通用规范标准
    ''' </summary>
    Private Function CreateBusinessStandard() As FormattingStandard
        Dim standard As New FormattingStandard()
        standard.Id = "business-report"
        standard.Name = "商务报告通用规范"
        standard.Description = "适用于商务报告、商业计划书、工作总结等商务文档的通用格式规范。采用现代商务风格排版。"
        standard.ApplicableDocumentTypes.Add(DocumentType.BusinessReport.ToString())
        standard.IsBuiltIn = True
        standard.SemanticMapping = CreateBusinessMapping()
        Return standard
    End Function

    ''' <summary>
    ''' 创建商务报告通用规范的 SemanticStyleMapping
    ''' </summary>
    Private Function CreateBusinessMapping() As SemanticStyleMapping
        Dim mapping As New SemanticStyleMapping()
        mapping.Id = "business-report"
        mapping.Name = "商务报告通用规范"
        mapping.SourceType = SemanticMappingSourceType.FromStyleGuide

        ' 页面设置（商务报告通常使用稍宽的页边距）
        mapping.PageConfig.Margins = New MarginsConfig(2.54, 2.54, 3.18, 3.18)

        ' 语义标签
        Dim tags = mapping.SemanticTags

        ' 报告标题
        Dim titleMain As New SemanticTag("title.main", "报告标题", "title", 2, "商务报告主标题")
        titleMain.Font = New FontConfig("微软雅黑", "Arial", 20, True)
        titleMain.Paragraph = New ParagraphConfig("center", 0, 1.5)
        titleMain.Paragraph.SpaceBefore = 3
        titleMain.Paragraph.SpaceAfter = 1
        tags.Add(titleMain)

        ' 一级标题
        Dim heading1 As New SemanticTag("heading.1", "一级标题", "", 2, "报告章节标题")
        heading1.Font = New FontConfig("微软雅黑", "Arial", 16, True)
        heading1.Paragraph = New ParagraphConfig("left")
        heading1.Paragraph.SpaceBefore = 1.5
        heading1.Paragraph.SpaceAfter = 0.5
        tags.Add(heading1)

        ' 二级标题
        Dim heading2 As New SemanticTag("heading.2", "二级标题", "", 2, "报告子章节标题")
        heading2.Font = New FontConfig("微软雅黑", "Arial", 14, True)
        heading2.Paragraph = New ParagraphConfig("left")
        heading2.Paragraph.SpaceBefore = 1
        heading2.Paragraph.SpaceAfter = 0.25
        tags.Add(heading2)

        ' 正文
        Dim bodyNormal As New SemanticTag("body.normal", "正文", "body", 2, "商务报告正文段落")
        bodyNormal.Font = New FontConfig("微软雅黑", "Arial", 11)
        bodyNormal.Paragraph = New ParagraphConfig("justify", 0, 1.25)
        tags.Add(bodyNormal)

        ' 摘要/总结段落
        Dim bodySummary As New SemanticTag("body.summary", "摘要", "body", 2, "报告摘要或执行摘要段落")
        bodySummary.Font = New FontConfig("微软雅黑", "Arial", 11)
        bodySummary.Paragraph = New ParagraphConfig("justify", 0, 1.25)
        bodySummary.Paragraph.SpaceBefore = 1
        bodySummary.Paragraph.SpaceAfter = 1
        tags.Add(bodySummary)

        ' 页码
        Dim footerPage As New SemanticTag("footer.page", "页码", "footer", 2, "报告页码")
        footerPage.Font = New FontConfig("微软雅黑", "Arial", 9)
        footerPage.Paragraph = New ParagraphConfig("center")
        tags.Add(footerPage)

        Return mapping
    End Function

#End Region

End Module
