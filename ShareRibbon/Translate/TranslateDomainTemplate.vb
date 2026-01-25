Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 翻译领域模板 - 支持不同专业领域的翻译配置
''' </summary>
Public Class TranslateDomainTemplate
    ''' <summary>领域名称</summary>
    Public Property Name As String = ""

    ''' <summary>领域描述</summary>
    Public Property Description As String = ""

    ''' <summary>系统提示词模板</summary>
    Public Property SystemPrompt As String = ""

    ''' <summary>是否为内置模板</summary>
    Public Property IsBuiltIn As Boolean = False

    ''' <summary>专业术语列表（可选）</summary>
    Public Property Glossary As Dictionary(Of String, String) = New Dictionary(Of String, String)()

    Public Sub New()
    End Sub

    Public Sub New(name As String, description As String, systemPrompt As String, Optional isBuiltIn As Boolean = False)
        Me.Name = name
        Me.Description = description
        Me.SystemPrompt = systemPrompt
        Me.IsBuiltIn = isBuiltIn
    End Sub
End Class

''' <summary>
''' 翻译领域模板管理器
''' </summary>
Public Class TranslateDomainManager
    Private Shared ReadOnly fileName As String = "translate_domains.json"
    Private Shared ReadOnly filePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                                              ConfigSettings.OfficeAiAppDataFolder, fileName)

    Private Shared _templates As List(Of TranslateDomainTemplate)

    ''' <summary>
    ''' 获取所有领域模板
    ''' </summary>
    Public Shared ReadOnly Property Templates As List(Of TranslateDomainTemplate)
        Get
            If _templates Is Nothing Then
                Load()
            End If
            Return _templates
        End Get
    End Property

    ''' <summary>
    ''' 获取内置领域模板
    ''' </summary>
    Private Shared Function GetBuiltInTemplates() As List(Of TranslateDomainTemplate)
        Return New List(Of TranslateDomainTemplate) From {
            New TranslateDomainTemplate(
                "通用",
                "通用翻译，适用于日常文档",
                "你是一个专业的翻译专家。请准确翻译以下内容，保持原文的格式、语气和风格。翻译时注意：
1. 保持段落结构不变
2. 专有名词可保留原文或在括号中注明
3. 数字、日期格式保持一致
4. 语句通顺自然，符合目标语言习惯",
                True
            ),
            New TranslateDomainTemplate(
                "金融财经",
                "金融、财务、投资领域专业翻译",
                "你是金融财经领域的专业翻译专家。请翻译以下内容，注意：
1. 金融术语使用标准译法（如：equity-股权/权益，derivative-衍生品，hedge-对冲）
2. 财务报表科目使用规范会计术语
3. 数字、货币、百分比格式保持精确
4. 公司名称、机构名称首次出现时可附原文
5. 保持专业严谨的语气
常见术语对照：ROI-投资回报率，P/E ratio-市盈率，liquidity-流动性，leverage-杠杆",
                True
            ),
            New TranslateDomainTemplate(
                "工程技术",
                "工程、技术、制造领域专业翻译",
                "你是工程技术领域的专业翻译专家。请翻译以下内容，注意：
1. 技术术语使用行业标准译法
2. 计量单位保持原样或转换为目标语言习惯（如英制/公制）
3. 公式、规格参数保持准确
4. 型号、标准编号（如ISO、GB）保留原文
5. 保持技术文档的严谨性和准确性
6. 操作步骤、流程描述清晰明确",
                True
            ),
            New TranslateDomainTemplate(
                "法律合同",
                "法律、合同、法规领域专业翻译",
                "你是法律领域的专业翻译专家。请翻译以下内容，注意：
1. 法律术语使用规范译法（如：jurisdiction-管辖权，liability-责任/义务，indemnify-赔偿）
2. 合同条款结构保持完整
3. 日期、金额、当事人名称准确无误
4. 法律条文引用格式规范
5. 保持法律文书的严谨性和权威性
6. 不随意省略或添加内容",
                True
            ),
            New TranslateDomainTemplate(
                "医学医疗",
                "医学、医疗、制药领域专业翻译",
                "你是医学医疗领域的专业翻译专家。请翻译以下内容，注意：
1. 医学术语使用标准译法，必要时附拉丁文/英文原名
2. 药品名称使用通用名，可附商品名
3. 剂量、用法、频次准确无误
4. 解剖学术语、病理学术语规范
5. 临床表现、诊断、治疗方案翻译准确
6. 保持医学文献的专业性和严谨性",
                True
            ),
            New TranslateDomainTemplate(
                "学术论文",
                "学术论文、研究报告翻译",
                "你是学术领域的专业翻译专家。请翻译以下内容，注意：
1. 学术术语使用学科规范译法
2. 摘要、关键词、正文结构保持完整
3. 引用、参考文献格式保持一致
4. 图表标题、注释翻译准确
5. 研究方法、数据分析描述清晰
6. 保持学术文体的严谨性和客观性",
                True
            ),
            New TranslateDomainTemplate(
                "商务文书",
                "商务邮件、报告、提案翻译",
                "你是商务领域的专业翻译专家。请翻译以下内容，注意：
1. 保持商务文书的专业性和礼貌性
2. 公司名称、职位、部门名称规范
3. 商务术语使用得体
4. 语气正式但不失亲和
5. 日期、数字、联系方式格式规范
6. 邮件格式、称呼、落款保持完整",
                True
            ),
            New TranslateDomainTemplate(
                "文学创意",
                "文学作品、创意内容翻译",
                "你是文学翻译专家。请翻译以下内容，注意：
1. 保留原文的文学性和艺术性
2. 修辞手法、比喻、典故适当本地化
3. 人物对话保持性格特征
4. 文化元素合理转换
5. 韵律、节奏尽可能保留
6. 译文流畅优美，富有表现力",
                True
            )
        }
    End Function

    ''' <summary>
    ''' 加载领域模板
    ''' </summary>
    Public Shared Sub Load()
        Try
            If File.Exists(filePath) Then
                Dim json As String = File.ReadAllText(filePath)
                _templates = JsonConvert.DeserializeObject(Of List(Of TranslateDomainTemplate))(json)
            End If
        Catch
            _templates = Nothing
        End Try

        ' 确保内置模板存在
        If _templates Is Nothing Then
            _templates = New List(Of TranslateDomainTemplate)()
        End If

        Dim builtInTemplates = GetBuiltInTemplates()
        For Each builtIn In builtInTemplates
            Dim existing = _templates.FirstOrDefault(Function(t) t.Name = builtIn.Name AndAlso t.IsBuiltIn)
            If existing Is Nothing Then
                _templates.Insert(0, builtIn)
            Else
                ' 更新内置模板的提示词
                existing.Description = builtIn.Description
                existing.SystemPrompt = builtIn.SystemPrompt
            End If
        Next

        Save()
    End Sub

    ''' <summary>
    ''' 保存领域模板
    ''' </summary>
    Public Shared Sub Save()
        Try
            Dim dir = Path.GetDirectoryName(filePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
            Dim json As String = JsonConvert.SerializeObject(_templates, Formatting.Indented)
            File.WriteAllText(filePath, json)
        Catch
            ' 忽略保存错误
        End Try
    End Sub

    ''' <summary>
    ''' 添加自定义领域模板
    ''' </summary>
    Public Shared Sub AddTemplate(template As TranslateDomainTemplate)
        template.IsBuiltIn = False
        _templates.Add(template)
        Save()
    End Sub

    ''' <summary>
    ''' 删除领域模板（仅可删除非内置模板）
    ''' </summary>
    Public Shared Function RemoveTemplate(name As String) As Boolean
        Dim template = _templates.FirstOrDefault(Function(t) t.Name = name AndAlso Not t.IsBuiltIn)
        If template IsNot Nothing Then
            _templates.Remove(template)
            Save()
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' 根据名称获取模板
    ''' </summary>
    Public Shared Function GetTemplate(name As String) As TranslateDomainTemplate
        Return _templates.FirstOrDefault(Function(t) t.Name = name)
    End Function
End Class
