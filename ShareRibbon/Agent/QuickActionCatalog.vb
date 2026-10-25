' ShareRibbon\Agent\QuickActionCatalog.vb
' E4-01 起：宿主原生入口（选区工具条 / E4-02 右键菜单 / E4-03 快捷键）的唯一动作清单，
' 并承载 E3-01/E3-02 引导向导的每端示例任务——两处话术只在这一处维护，不在宿主项目里各留一张需要人工同步的表。
' E6-01 起再加一类：三端 Ribbon 按钮的 Agent 请求话术（下面 Ribbon* 那一段），入口项目不再自带文案。
' 入口动作只是"把一句话交给同一条 Agent 主链"，不承载执行能力，因此不进 ToolRegistry、不进宿主 capability 清单；
' 能对应已登记 capability 的动作必须填 CapabilityId，由 smoke 校验其真实存在，防止"入口说得到、能力做不到"。

Namespace Agent

    ''' <summary>
    ''' 一个原生入口动作：Label 是按钮文案，RequestText 是发给 Agent 的自然语言请求。
    ''' </summary>
    Public Class QuickAction

        Public Property Id As String = ""

        Public Property Label As String = ""

        Public Property RequestText As String = ""

        ''' <summary>可选：对应的宿主 capability id（如 word.proofread）；Nothing 表示纯模型任务。</summary>
        Public Property CapabilityId As String
    End Class

    ''' <summary>
    ''' 一条示例任务。Text 是向导上的文案；RunRequestText 非空表示这条可以"一键跑"（E3-02），
    ''' 点下去就是把这句话交给同一条 Agent 主链，而不是只填进输入框让用户自己按发送。
    ''' </summary>
    Public Class SampleTask

        Public Property Text As String = ""

        ''' <summary>可选：一键示例真正发给主链的请求；Nothing 表示这条只是填进输入框的建议。</summary>
        Public Property RunRequestText As String

        Public ReadOnly Property Runnable As Boolean
            Get
                Return Not String.IsNullOrWhiteSpace(RunRequestText)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' 按宿主返回入口动作与示例任务。未接入的宿主返回空表，调用方据此不出入口而不是猜文案。
    ''' </summary>
    Public NotInheritable Class QuickActionCatalog

        Public Const WordAppType As String = "Word"
        Public Const ExcelAppType As String = "Excel"
        Public Const PowerPointAppType As String = "PowerPoint"

        Private Shared ReadOnly _wordActions As IReadOnlyList(Of QuickAction) = New QuickAction() {
            New QuickAction With {
                .Id = "word.selection.rewrite",
                .Label = "改写选区",
                .RequestText = "把当前选中的内容改写成更清晰、简洁的书面表达，保持原意和格式，只替换当前选区文本。"
            },
            New QuickAction With {
                .Id = "word.selection.summarize",
                .Label = "生成摘要",
                .RequestText = "为当前选中的内容生成一段摘要（不超过 3 句），只输出摘要文本，不要改动文档。"
            },
            New QuickAction With {
                .Id = "word.selection.proofread",
                .Label = "校对选区",
                .CapabilityId = "word.proofread",
                .RequestText = "校对当前选中的内容，修正错别字、标点和明显语法问题，只处理当前选区。"
            }
        }

        Private Shared ReadOnly _emptyActions As IReadOnlyList(Of QuickAction) = New QuickAction() {}

        ' E4-02 右键菜单：Excel/PPT 侧还没有宿主 capability 注册表（见开放待办 #36），
        ' 因此这两个宿主的动作一律不声明 CapabilityId——宁可少声称，也不要"入口说得到、能力做不到"。
        Private Shared ReadOnly _excelActions As IReadOnlyList(Of QuickAction) = New QuickAction() {
            New QuickAction With {
                .Id = "excel.selection.explain",
                .Label = "AI 解读选区",
                .RequestText = "读取当前选中的单元格区域，说明它是什么数据、有哪些异常和值得注意的结论，不要修改工作表内容。"
            },
            New QuickAction With {
                .Id = "excel.selection.formula",
                .Label = "AI 生成公式",
                .RequestText = "为当前选中的单元格区域判断需要哪些计算，先说明公式含义与写入位置，再按低风险方式把公式写到相邻空白列。"
            },
            New QuickAction With {
                .Id = "excel.selection.format",
                .Label = "AI 整理格式",
                .RequestText = "把当前选中的单元格区域整理成清晰的表格样式：表头、对齐、数字格式、列宽和边框，只处理当前选区。"
            }
        }

        Private Shared ReadOnly _powerPointActions As IReadOnlyList(Of QuickAction) = New QuickAction() {
            New QuickAction With {
                .Id = "ppt.selection.rewrite",
                .Label = "AI 改写选中文字",
                .RequestText = "把当前选中的文本改写成更适合演示的书面表达，保持原意与原有字号、字体格式，只替换当前选中文本。"
            },
            New QuickAction With {
                .Id = "ppt.slide.summarize",
                .Label = "AI 提炼本页要点",
                .RequestText = "为当前幻灯片提炼不超过 3 条要点，只输出要点文本，不要改动演示文稿。"
            },
            New QuickAction With {
                .Id = "ppt.slide.proofread",
                .Label = "AI 校对本页",
                .RequestText = "校对当前幻灯片里的文字，修正错别字、标点和明显语法问题，只处理当前幻灯片。"
            }
        }

        ' E6-01：三端 Ribbon 按钮的话术。以前每条请求写死在各宿主 Ribbon1.vb 的 Click 里，
        ' 于是"入口话术"这件事实有了第二个持有者——改一句话术要在四个文件里考古，而没有任何检查会响。
        ' 现在它们与原生入口动作住在同一个类里（仍然是一份清单），但**不进 ActionsFor**：
        ' ActionsFor 喂的是选区浮层/右键菜单/命令面板，把这些按钮文案一并塞进去会当场改变面板行数，
        ' 那是产品可见的改动，属于 E6-07 与开放决策 D9（这些按钮大多是 Visible=False 的隐藏实现）要定的事。
        ' 三端 StartAgentFromRibbonAsync 的形参已改成 QuickAction，传字符串不再能编译，
        ' "入口不再持有第二份话术"因此由编译器而不是由约定保证；Label 取 Ribbon 设计器里的真实按钮文字。
        Public Shared ReadOnly WordRibbonAnalyze As QuickAction = New QuickAction With {
            .Id = "word.ribbon.analyze",
            .Label = "数据分析",
            .RequestText = "请自动分析当前选中文档内容或整篇 Word 文档，提取结构、关键数据、风险点、摘要结论和可执行改进建议；如包含表格或数字，请优先进行结构化分析。"
        }

        Public Shared ReadOnly ExcelRibbonAnalyze As QuickAction = New QuickAction With {
            .Id = "excel.ribbon.analyze",
            .Label = "数据分析",
            .RequestText = "请作为 Excel 数据分析智能体读取当前选区或当前工作表，自动判断数据结构并执行合适的数据分析：统计摘要、分组汇总、排名、公式计算、透视表、图表或报表生成。优先使用已注册 Excel 工具和匹配的 Skills，执行后观察结果并说明已完成的修改。"
        }

        Public Shared ReadOnly ExcelRibbonWebCapture As QuickAction = New QuickAction With {
            .Id = "excel.ribbon.web-capture",
            .Label = "抓取网页",
            .RequestText = "请根据当前工作簿上下文自动处理网页内容采集需求；如果用户已提供网址或网页文本，请提取结构化数据并写入适合的工作表区域。"
        }

        Public Shared ReadOnly ExcelRibbonProofread As QuickAction = New QuickAction With {
            .Id = "excel.ribbon.proofread",
            .Label = "校对",
            .RequestText = "请自动校对当前选中单元格区域，检查错别字、标点、术语一致性和明显的数据文本问题，并在适合时直接给出修改建议或执行低风险修正。"
        }

        Public Shared ReadOnly ExcelRibbonReformat As QuickAction = New QuickAction With {
            .Id = "excel.ribbon.reformat",
            .Label = "排版",
            .RequestText = "请自动分析当前选中单元格区域或当前工作表，并进行专业表格排版：标题、表头、列宽、数字格式、边框、对齐、重点高亮和可读性优化。"
        }

        Public Shared ReadOnly ExcelRibbonContinuation As QuickAction = New QuickAction With {
            .Id = "excel.ribbon.continuation",
            .Label = "续写",
            .RequestText = "请根据当前选中单元格或相邻数据自动续写/补全内容，保持已有字段、语气、格式和数据规律一致。"
        }

        Public Shared ReadOnly ExcelRibbonTemplateFormat As QuickAction = New QuickAction With {
            .Id = "excel.ribbon.template-format",
            .Label = "模板排版",
            .RequestText = "请根据当前工作表内容自动识别表格类型，并套用适合的专业模板化排版方案；如果有可用 Skills 或模板，请自动选择最匹配的方案执行。"
        }

        Public Shared ReadOnly PptRibbonWebCapture As QuickAction = New QuickAction With {
            .Id = "ppt.ribbon.web-capture",
            .Label = "抓取网页",
            .RequestText = "请根据当前演示文稿上下文自动处理网页内容采集需求；如果用户已提供网址或网页文本，请提取适合制作幻灯片的结构化要点、标题和页面内容。"
        }

        Public Shared ReadOnly PptRibbonAnalyze As QuickAction = New QuickAction With {
            .Id = "ppt.ribbon.analyze",
            .Label = "数据分析",
            .RequestText = "请自动分析当前演示文稿或选中幻灯片中的数据、图表和表格，提炼关键结论、异常点和适合放入演示的洞察；如可执行，请直接生成或优化对应页面内容。"
        }

        Public Shared ReadOnly PptRibbonProofread As QuickAction = New QuickAction With {
            .Id = "ppt.ribbon.proofread",
            .Label = "校对",
            .RequestText = "请自动校对当前选中幻灯片或整份演示文稿，检查错别字、标点、术语一致性、标题层级、表达简洁度和演示语气，并在适合时直接执行低风险修正。"
        }

        ' 引导向导的示例任务（E3-01 原有文案原样迁入）。每个宿主第一条是一键示例（E3-02）：
        ' 请求里同时写清"当前文档有内容怎么办"和"文档还空着怎么办"，由能读上下文的主链自己判断，
        ' 入口不猜用户有没有选区、也不为了猜这一层去调 COM。三处都必须带"预览"，与计划卡的风险确认同口径。
        Private Shared ReadOnly _wordSamples As IReadOnlyList(Of SampleTask) = {
            New SampleTask With {
                .Text = "一键示例：把一段话改写成书面表达",
                .RunRequestText = "先看当前文档有没有可用内容：有选区就改写当前选区，没有选区就改写第一段，把表达改得更清晰、简洁、正式，保持原意和原有格式；如果文档还是空的，先插入一段 4 句左右的内置示例文字，再改写这段示例。只改这一处，写入前先给我预览。"
            },
            New SampleTask With {
                .Text = "检查全文的编号和标题层级是否连贯"
            },
            New SampleTask With {
                .Text = "为当前文档生成一段摘要"
            }
        }

        Private Shared ReadOnly _excelSamples As IReadOnlyList(Of SampleTask) = {
            New SampleTask With {
                .Text = "一键示例：给当前数据加一列同比增长率",
                .RunRequestText = "先看当前工作表有没有可用数据：有就用当前选区，没有选区就取正在使用的表格区域，在右侧相邻空白列新增一列同比增长率公式并说明公式含义；如果工作表是空的，先写入一张内置示例数据表（月份、销售额两列，6 行），再在它上面加这一列公式。只改这一列，写入前先给我预览。"
            },
            New SampleTask With {
                .Text = "读取当前选区，找出异常值并说明原因"
            },
            New SampleTask With {
                .Text = "把选区整理成一张柱状图"
            }
        }

        Private Shared ReadOnly _pptSamples As IReadOnlyList(Of SampleTask) = {
            New SampleTask With {
                .Text = "一键示例：提炼本页要点（空文档就先做一页示例）",
                .RunRequestText = "先看当前演示文稿有没有内容：有就为当前页提炼不超过 3 条要点，只输出要点文本，不改动演示文稿；如果是空演示文稿，先用一个内置示例主题（例如本季度业务汇报）生成一页标题加三条要点，沿用模板现有字体和版式，写入前先给我预览。"
            },
            New SampleTask With {
                .Text = "统一全部标题的字体和字号"
            },
            New SampleTask With {
                .Text = "检查哪一页文字过密并给出拆分建议"
            }
        }

        Private Sub New()
        End Sub

        ''' <summary>该宿主的原生入口动作（选区浮层与右键菜单共用这一份）；未接入的宿主返回空表。</summary>
        Public Shared Function ActionsFor(appType As String) As IReadOnlyList(Of QuickAction)
            Select Case If(appType, "").Trim().ToLowerInvariant()
                Case "word"
                    Return _wordActions
                Case "excel"
                    Return _excelActions
                Case "powerpoint"
                    Return _powerPointActions
                Case Else
                    Return _emptyActions
            End Select
        End Function

        ''' <summary>该宿主引导向导展示的示例任务；第一条是一键示例，其余只填进输入框。</summary>
        Public Shared Function SampleTasks(appType As String) As IReadOnlyList(Of SampleTask)
            Select Case If(appType, "").Trim().ToLowerInvariant()
                Case "excel"
                    Return _excelSamples
                Case "powerpoint"
                    Return _pptSamples
                Case Else
                    Return _wordSamples
            End Select
        End Function
    End Class

End Namespace
