' ShareRibbon\Services\Reformat\FormattingKnowledgeEngine.vb
' 排版知识引擎 - 管理内置和用户自定义的排版标准

' 注意：DocumentType 枚举在 DocumentAnalyzer.vb 中定义，此文件引用之
' 注意：DocumentAnalysisResult 也在 DocumentAnalyzer.vb 中定义

''' <summary>
''' 排版标准 - 定义一种文档类型的完整格式规范
''' </summary>
Public Class FormattingStandard
    ''' <summary>标准唯一标识</summary>
    Public Property Id As String = Guid.NewGuid().ToString()

    ''' <summary>标准名称（如 "GB/T 9704-2012"）</summary>
    Public Property Name As String = ""

    ''' <summary>标准描述</summary>
    Public Property Description As String = ""

    ''' <summary>适用的文档类型列表（DocumentType枚举值的名称）</summary>
    Public Property ApplicableDocumentTypes As List(Of String)

    ''' <summary>语义排版映射（核心格式数据）</summary>
    Public Property SemanticMapping As SemanticStyleMapping

    ''' <summary>是否为内置标准</summary>
    Public Property IsBuiltIn As Boolean = False

    ''' <summary>是否激活</summary>
    Public Property IsActive As Boolean = True

    Public Sub New()
        ApplicableDocumentTypes = New List(Of String)()
        SemanticMapping = New SemanticStyleMapping()
    End Sub

    Public Sub New(name As String, description As String)
        Me.Name = name
        Me.Description = description
        ApplicableDocumentTypes = New List(Of String)()
        SemanticMapping = New SemanticStyleMapping()
    End Sub

    Public Overrides Function ToString() As String
        Return If(String.IsNullOrEmpty(Name), "(未命名)", Name)
    End Function
End Class

''' <summary>
''' 排版知识引擎 - 管理内置和用户自定义的排版标准
''' 提供标准检索、标签规则解释等功能
''' </summary>
Public Class FormattingKnowledgeEngine
    Private ReadOnly _standards As New List(Of FormattingStandard)()
    Private ReadOnly _ruleExplanations As Dictionary(Of String, String)

    Public Sub New()
        ' 从内置数据加载所有标准
        Dim builtInStandards = FormattingStandardData.GetAllBuiltInStandards()
        _standards.AddRange(builtInStandards)

        ' 初始化规则解释
        _ruleExplanations = New Dictionary(Of String, String) From {
            {"title", "标题区域：文档标题的格式定义，包括字体、字号、对齐方式等"},
            {"title.main", "文档主标题：使用较大字号居中加粗显示，是文档最主要的标题"},
            {"title.recipient", "主送机关：公文的主送机关名称，左对齐顶格排列"},
            {"title.abstract", "摘要标题：用于标识摘要区域的标题，与正文区分"},
            {"title.keywords", "关键词标题：用于标识关键词区域的标题"},
            {"header", "页眉区域：公文页眉部分的格式定义"},
            {"header.org", "发文机关标志：使用方正小标宋简体22pt加粗，红色(#C00000)居中显示"},
            {"header.refno", "发文字号：使用仿宋_GB2312 16pt居中排列，包含机关代字、年份和顺序号"},
            {"header.separator", "红色分隔线：位于发文字号下方的红色横线，线宽2pt，颜色#C00000"},
            {"body", "正文区域：文档正文的基础格式定义"},
            {"body.normal", "正文段落：使用仿宋_GB2312 16pt，两端对齐，首行缩进2字符，行距28磅"},
            {"body.attachment", "附件说明：用于标注文档附件，正文下空1行排列，左对齐"},
            {"body.abstract", "摘要正文：学术论文摘要内容的格式定义，通常比正文紧凑"},
            {"body.keywords", "关键词：学术论文关键词的格式定义，位于摘要之后"},
            {"body.reference", "参考文献条目：参考文献列表中单条记录的格式定义"},
            {"body.summary", "总结段落：用于报告或文章的摘要总结部分，与正文有所区分"},
            {"heading", "标题层级：文档章节标题的格式定义，按层级使用不同字体和字号"},
            {"heading.1", "一级标题：公文的一级章节标题，使用黑体16pt加粗，段前段后各0.5行"},
            {"heading.2", "二级标题：公文的二级章节标题，使用楷体_GB2312 16pt加粗"},
            {"heading.3", "三级标题：公文的三级章节标题，使用仿宋_GB2312 16pt加粗"},
            {"footer", "页脚区域：文档页脚部分的格式定义"},
            {"footer.signature", "发文机关署名：位于成文日期之上的发文机关署名，右对齐"},
            {"footer.date", "成文日期：使用阿拉伯数字表示的公文成文日期，右对齐，不编虚位"},
            {"footer.page", "页码：文档页脚处页码，通常居中排列"}
        }
    End Sub

    ''' <summary>
    ''' 获取适用于指定文档类型的标准
    ''' </summary>
    Public Function GetStandardForDocumentType(docType As DocumentType) As FormattingStandard
        Dim typeName = docType.ToString()
        Return _standards.FirstOrDefault(Function(s) s.ApplicableDocumentTypes.Contains(typeName) AndAlso s.IsActive)
    End Function

    ''' <summary>
    ''' 根据名称获取标准
    ''' </summary>
    Public Function GetStandardByName(name As String) As FormattingStandard
        If String.IsNullOrEmpty(name) Then Return Nothing
        Return _standards.FirstOrDefault(Function(s) s.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
    End Function

    ''' <summary>
    ''' 获取所有已注册的标准
    ''' </summary>
    Public Function GetAllStandards() As List(Of FormattingStandard)
        Return _standards.ToList()
    End Function

    ''' <summary>
    ''' 获取所有已激活的标准
    ''' </summary>
    Public Function GetActiveStandards() As List(Of FormattingStandard)
        Return _standards.Where(Function(s) s.IsActive).ToList()
    End Function

    ''' <summary>
    ''' 注册新的排版标准（用户自定义）
    ''' </summary>
    Public Sub RegisterStandard(standard As FormattingStandard)
        If standard Is Nothing Then Return
        ' 避免重复注册同名标准
        Dim existing = GetStandardByName(standard.Name)
        If existing IsNot Nothing Then
            _standards.Remove(existing)
        End If
        standard.IsBuiltIn = False
        _standards.Add(standard)
    End Sub

    ''' <summary>
    ''' 取消注册指定标准
    ''' </summary>
    Public Function UnregisterStandard(standardId As String) As Boolean
        Dim standard = _standards.FirstOrDefault(Function(s) s.Id = standardId)
        If standard IsNot Nothing AndAlso Not standard.IsBuiltIn Then
            Return _standards.Remove(standard)
        End If
        Return False
    End Function

    ''' <summary>
    ''' 解释指定语义标签的排版规则
    ''' </summary>
    Public Function ExplainRule(tagId As String) As String
        If String.IsNullOrEmpty(tagId) Then Return "未指定标签。"

        ' 优先返回精确解释
        If _ruleExplanations.ContainsKey(tagId) Then
            Return _ruleExplanations(tagId)
        End If

        ' 尝试返回父级标签的解释
        Dim parentId = SemanticTagRegistry.GetParentTag(tagId)
        If Not String.IsNullOrEmpty(parentId) AndAlso _ruleExplanations.ContainsKey(parentId) Then
            Return $"{_ruleExplanations(parentId)}（{tagId} 是该大类下的具体细分标签）"
        End If

        Return $"标签「{tagId}」暂无详细解释，请参考所选排版标准的具体格式定义。"
    End Function

    ''' <summary>
    ''' 获取适用于指定文档类型的标准（支持字符串类型名称）
    ''' </summary>
    Public Function GetStandardForDocumentType(docTypeName As String) As FormattingStandard
        If String.IsNullOrEmpty(docTypeName) Then Return Nothing
        Return _standards.FirstOrDefault(Function(s) s.ApplicableDocumentTypes.Contains(docTypeName) AndAlso s.IsActive)
    End Function
End Class
