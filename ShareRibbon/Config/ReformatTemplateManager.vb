' ShareRibbon\Config\ReformatTemplateManager.vb
' 排版模板管理器（单例模式）

Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 排版模板管理器（单例模式）
''' </summary>
Public Class ReformatTemplateManager
    Private Shared _instance As ReformatTemplateManager
    Private _templates As List(Of ReformatTemplate)
    Private ReadOnly _configPath As String

    ''' <summary>获取单例实例</summary>
    Public Shared ReadOnly Property Instance As ReformatTemplateManager
        Get
            If _instance Is Nothing Then
                _instance = New ReformatTemplateManager()
            End If
            Return _instance
        End Get
    End Property

    ''' <summary>获取所有模板</summary>
    Public ReadOnly Property Templates As List(Of ReformatTemplate)
        Get
            Return _templates
        End Get
    End Property

    Private Sub New()
        _configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "reformat_templates.json")
        LoadTemplates()
    End Sub

    ''' <summary>加载模板配置</summary>
    Private Sub LoadTemplates()
        _templates = New List(Of ReformatTemplate)()

        If File.Exists(_configPath) Then
            Try
                Dim json = File.ReadAllText(_configPath, Text.Encoding.UTF8)
                Dim loadedTemplates = JsonConvert.DeserializeObject(Of List(Of ReformatTemplate))(json)

                ' 合并预置模板和用户模板
                MergePresetsAndUserTemplates(loadedTemplates)
            Catch ex As Exception
                Debug.WriteLine($"加载模板配置失败: {ex.Message}")
                LoadPresetTemplates()
            End Try
        Else
            ' 首次使用，加载预置模板
            LoadPresetTemplates()
            SaveTemplates()
        End If
    End Sub

    ''' <summary>合并预置模板和用户模板</summary>
    Private Sub MergePresetsAndUserTemplates(userTemplates As List(Of ReformatTemplate))
        If userTemplates Is Nothing Then
            LoadPresetTemplates()
            Return
        End If

        ' 先加载预置模板
        Dim presets = GetPresetTemplates()
        For Each preset In presets
            Dim existing = userTemplates.FirstOrDefault(Function(t) t.Id = preset.Id)
            If existing IsNot Nothing Then
                ' 保留用户修改，但标记为预置
                existing.IsPreset = True
                _templates.Add(existing)
            Else
                _templates.Add(preset)
            End If
        Next

        ' 添加用户自定义模板
        For Each userTemplate In userTemplates
            If Not userTemplate.IsPreset AndAlso Not _templates.Any(Function(t) t.Id = userTemplate.Id) Then
                _templates.Add(userTemplate)
            End If
        Next
    End Sub

    ''' <summary>加载预置模板</summary>
    Private Sub LoadPresetTemplates()
        _templates.AddRange(GetPresetTemplates())
    End Sub

    ''' <summary>保存模板配置</summary>
    Public Sub SaveTemplates()
        Try
            Dim dir = Path.GetDirectoryName(_configPath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json = JsonConvert.SerializeObject(_templates, Formatting.Indented)
            File.WriteAllText(_configPath, json, Text.Encoding.UTF8)
        Catch ex As Exception
            Debug.WriteLine($"保存模板配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>添加模板</summary>
    Public Sub AddTemplate(template As ReformatTemplate)
        template.Id = Guid.NewGuid().ToString()
        template.CreatedAt = DateTime.Now
        template.LastModified = DateTime.Now
        template.IsPreset = False
        _templates.Add(template)
        SaveTemplates()
    End Sub

    ''' <summary>更新模板</summary>
    Public Sub UpdateTemplate(template As ReformatTemplate)
        Dim existing = _templates.FirstOrDefault(Function(t) t.Id = template.Id)
        If existing IsNot Nothing Then
            template.LastModified = DateTime.Now
            Dim index = _templates.IndexOf(existing)
            _templates(index) = template
            SaveTemplates()
        End If
    End Sub

    ''' <summary>删除模板</summary>
    Public Function DeleteTemplate(templateId As String) As Boolean
        Dim template = _templates.FirstOrDefault(Function(t) t.Id = templateId)
        If template IsNot Nothing Then
            If template.IsPreset Then
                Return False ' 预置模板不可删除
            End If
            _templates.Remove(template)
            SaveTemplates()
            Return True
        End If
        Return False
    End Function

    ''' <summary>复制模板</summary>
    Public Function DuplicateTemplate(templateId As String, newName As String) As ReformatTemplate
        Dim original = _templates.FirstOrDefault(Function(t) t.Id = templateId)
        If original IsNot Nothing Then
            ' 深拷贝
            Dim json = JsonConvert.SerializeObject(original)
            Dim duplicate = JsonConvert.DeserializeObject(Of ReformatTemplate)(json)
            duplicate.Id = Guid.NewGuid().ToString()
            duplicate.Name = If(String.IsNullOrEmpty(newName), original.Name & " (副本)", newName)
            duplicate.IsPreset = False
            duplicate.CreatedAt = DateTime.Now
            duplicate.LastModified = DateTime.Now
            _templates.Add(duplicate)
            SaveTemplates()
            Return duplicate
        End If
        Return Nothing
    End Function

    ''' <summary>导出模板到文件</summary>
    Public Function ExportTemplate(templateId As String, filePath As String) As Boolean
        Try
            Dim template = _templates.FirstOrDefault(Function(t) t.Id = templateId)
            If template IsNot Nothing Then
                Dim json = JsonConvert.SerializeObject(template, Formatting.Indented)
                File.WriteAllText(filePath, json, Text.Encoding.UTF8)
                Return True
            End If
        Catch ex As Exception
            Debug.WriteLine($"导出模板失败: {ex.Message}")
        End Try
        Return False
    End Function

    ''' <summary>从文件导入模板</summary>
    Public Function ImportTemplate(filePath As String) As ReformatTemplate
        Try
            Dim json = File.ReadAllText(filePath, Text.Encoding.UTF8)
            Dim template = JsonConvert.DeserializeObject(Of ReformatTemplate)(json)
            template.Id = Guid.NewGuid().ToString() ' 生成新ID
            template.IsPreset = False
            template.CreatedAt = DateTime.Now
            template.LastModified = DateTime.Now
            _templates.Add(template)
            SaveTemplates()
            Return template
        Catch ex As Exception
            Debug.WriteLine($"导入模板失败: {ex.Message}")
        End Try
        Return Nothing
    End Function

    ''' <summary>根据ID获取模板</summary>
    Public Function GetTemplateById(templateId As String) As ReformatTemplate
        Return _templates.FirstOrDefault(Function(t) t.Id = templateId)
    End Function

    ''' <summary>按分类获取模板</summary>
    Public Function GetTemplatesByCategory(category As String) As List(Of ReformatTemplate)
        If String.IsNullOrEmpty(category) OrElse category = "全部" Then
            Return _templates.ToList()
        End If
        Return _templates.Where(Function(t) t.Category = category).ToList()
    End Function

    ''' <summary>获取所有分类</summary>
    Public Function GetAllCategories() As List(Of String)
        Dim categories = _templates.Select(Function(t) t.Category).Distinct().ToList()
        categories.Insert(0, "全部")
        Return categories
    End Function

    ''' <summary>刷新模板列表（重新从文件加载）</summary>
    Public Sub Refresh()
        LoadTemplates()
    End Sub

#Region "预置模板"

    ''' <summary>获取预置模板列表</summary>
    Private Function GetPresetTemplates() As List(Of ReformatTemplate)
        Dim presets As New List(Of ReformatTemplate)()

        ' 预置模板1：通用公文模板
        presets.Add(CreateGeneralOfficialTemplate())

        ' 预置模板2：行政公文模板
        presets.Add(CreateAdministrativeTemplate())

        ' 预置模板3：学术论文模板
        presets.Add(CreateAcademicTemplate())

        ' 预置模板4：商务报告模板
        presets.Add(CreateBusinessReportTemplate())

        Return presets
    End Function

    ''' <summary>创建通用公文模板</summary>
    Private Function CreateGeneralOfficialTemplate() As ReformatTemplate
        Dim template As New ReformatTemplate With {
            .Id = "preset-general-official",
            .Name = "通用公文模板",
            .Description = "适用于一般行政公文、通知、函件等",
            .Category = "通用",
            .TargetApp = "Word",
            .IsPreset = True,
            .AiGuidance = "这是标准的党政机关公文格式模板，请严格按照《党政机关公文格式》(GB/T 9704-2012)标准执行。"
        }

        ' 版式配置
        template.Layout = New LayoutConfig()
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "发文机关",
            .ElementType = "text",
            .DefaultValue = "[机关名称]",
            .Required = True,
            .SortOrder = 1,
            .Font = New FontConfig("方正小标宋简体", "Arial", 22, True),
            .Paragraph = New ParagraphConfig("center", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "红色横线",
            .ElementType = "redLine",
            .DefaultValue = "",
            .Required = True,
            .SortOrder = 2,
            .Font = New FontConfig(),
            .Paragraph = New ParagraphConfig("center"),
            .Color = New ColorConfig("#FF0000"),
            .SpecialProps = New Dictionary(Of String, String) From {
                {"lineWidth", "2pt"},
                {"lineColor", "#FF0000"}
            }
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "发文字号",
            .ElementType = "text",
            .DefaultValue = "〔2024〕X号",
            .Required = True,
            .SortOrder = 3,
            .Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16),
            .Paragraph = New ParagraphConfig("center", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "文件标题",
            .ElementType = "text",
            .DefaultValue = "",
            .Required = True,
            .SortOrder = 4,
            .Font = New FontConfig("方正小标宋简体", "Arial", 22, True),
            .Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceAfter = 1},
            .Color = New ColorConfig("#000000")
        })

        ' 正文样式
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "正文",
            .MatchCondition = "默认正文段落",
            .SortOrder = 1,
            .Font = New FontConfig("仿宋_GB2312", "Times New Roman", 16),
            .Paragraph = New ParagraphConfig("justify", 2, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "一级标题",
            .MatchCondition = "包含'一、'或'（一）'",
            .SortOrder = 2,
            .Font = New FontConfig("黑体", "Arial", 16, True),
            .Paragraph = New ParagraphConfig("left", 0, 1.5) With {.SpaceBefore = 0.5, .SpaceAfter = 0.5},
            .Color = New ColorConfig("#000000")
        })

        ' 页面设置
        template.PageSettings = New PageConfig With {
            .Margins = New MarginsConfig(3.7, 3.5, 2.8, 2.6),
            .Header = New HeaderFooterConfig(False),
            .Footer = New HeaderFooterConfig(False),
            .PageNumber = New PageNumberConfig(True, "footer", "center", "第{page}页")
        }

        Return template
    End Function

    ''' <summary>创建行政公文模板</summary>
    Private Function CreateAdministrativeTemplate() As ReformatTemplate
        Dim template As New ReformatTemplate With {
            .Id = "preset-administrative",
            .Name = "行政公文模板",
            .Description = "适用于政府公文、批复、决定等正式文件",
            .Category = "行政",
            .TargetApp = "Word",
            .IsPreset = True,
            .AiGuidance = "严格遵循党政机关公文格式国家标准GB/T 9704-2012。注意版记、附件说明、成文日期等要素的位置。"
        }

        ' 版式配置
        template.Layout = New LayoutConfig()
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "发文机关标志",
            .ElementType = "text",
            .DefaultValue = "[机关全称]",
            .Required = True,
            .SortOrder = 1,
            .Font = New FontConfig("方正小标宋简体", "Arial", 22, True),
            .Paragraph = New ParagraphConfig("center", 0, 1.5),
            .Color = New ColorConfig("#C00000")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "分隔线",
            .ElementType = "redLine",
            .DefaultValue = "",
            .Required = True,
            .SortOrder = 2,
            .Font = New FontConfig(),
            .Paragraph = New ParagraphConfig("center"),
            .Color = New ColorConfig("#C00000"),
            .SpecialProps = New Dictionary(Of String, String) From {
                {"lineWidth", "3pt"},
                {"lineColor", "#C00000"}
            }
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "发文字号",
            .ElementType = "text",
            .DefaultValue = "X政发〔2024〕X号",
            .Required = True,
            .SortOrder = 3,
            .Font = New FontConfig("仿宋", "Times New Roman", 16),
            .Paragraph = New ParagraphConfig("center", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "文件标题",
            .ElementType = "text",
            .DefaultValue = "关于XXXX的通知",
            .Required = True,
            .SortOrder = 4,
            .Font = New FontConfig("方正小标宋简体", "Arial", 22, True),
            .Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceAfter = 1.5},
            .Color = New ColorConfig("#000000")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "主送机关",
            .ElementType = "text",
            .DefaultValue = "各相关单位：",
            .Required = True,
            .SortOrder = 5,
            .Font = New FontConfig("仿宋", "Times New Roman", 16),
            .Paragraph = New ParagraphConfig("left", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })

        ' 正文样式
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "正文",
            .MatchCondition = "默认正文段落",
            .SortOrder = 1,
            .Font = New FontConfig("仿宋", "Times New Roman", 16),
            .Paragraph = New ParagraphConfig("justify", 2, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "大标题",
            .MatchCondition = "包含'一、'",
            .SortOrder = 2,
            .Font = New FontConfig("黑体", "Arial", 16, True),
            .Paragraph = New ParagraphConfig("left", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })

        ' 页面设置
        template.PageSettings = New PageConfig With {
            .Margins = New MarginsConfig(3.7, 3.5, 2.8, 2.6),
            .Header = New HeaderFooterConfig(False),
            .Footer = New HeaderFooterConfig(False),
            .PageNumber = New PageNumberConfig(True, "footer", "center", "—{page}—")
        }

        Return template
    End Function

    ''' <summary>创建学术论文模板</summary>
    Private Function CreateAcademicTemplate() As ReformatTemplate
        Dim template As New ReformatTemplate With {
            .Id = "preset-academic",
            .Name = "学术论文模板",
            .Description = "适用于学术期刊投稿、毕业论文等",
            .Category = "学术",
            .TargetApp = "Word",
            .IsPreset = True,
            .AiGuidance = "学术论文标准格式。注意区分摘要、关键词、引言、正文、参考文献等部分。参考文献采用GB/T 7714-2015格式。"
        }

        ' 版式配置
        template.Layout = New LayoutConfig()
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "论文标题",
            .ElementType = "text",
            .DefaultValue = "论文标题",
            .Required = True,
            .SortOrder = 1,
            .Font = New FontConfig("黑体", "Arial", 18, True),
            .Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceAfter = 1},
            .Color = New ColorConfig("#000000")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "作者信息",
            .ElementType = "text",
            .DefaultValue = "作者姓名（单位）",
            .Required = True,
            .SortOrder = 2,
            .Font = New FontConfig("宋体", "Times New Roman", 12),
            .Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceAfter = 1},
            .Color = New ColorConfig("#000000")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "摘要标题",
            .ElementType = "text",
            .DefaultValue = "摘要：",
            .Required = True,
            .SortOrder = 3,
            .Font = New FontConfig("黑体", "Arial", 12, True),
            .Paragraph = New ParagraphConfig("left", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })

        ' 正文样式
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "摘要正文",
            .MatchCondition = "摘要后的段落",
            .SortOrder = 1,
            .Font = New FontConfig("宋体", "Times New Roman", 10.5),
            .Paragraph = New ParagraphConfig("justify", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "一级标题",
            .MatchCondition = "包含'1 '或'一、'开头的短段落",
            .SortOrder = 2,
            .Font = New FontConfig("黑体", "Arial", 14, True),
            .Paragraph = New ParagraphConfig("left", 0, 1.5) With {.SpaceBefore = 0.5, .SpaceAfter = 0.5},
            .Color = New ColorConfig("#000000")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "二级标题",
            .MatchCondition = "包含'1.1 '或'（一）'开头的短段落",
            .SortOrder = 3,
            .Font = New FontConfig("黑体", "Arial", 12, True),
            .Paragraph = New ParagraphConfig("left", 0, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "正文",
            .MatchCondition = "默认正文段落",
            .SortOrder = 4,
            .Font = New FontConfig("宋体", "Times New Roman", 12),
            .Paragraph = New ParagraphConfig("justify", 2, 1.5),
            .Color = New ColorConfig("#000000")
        })

        ' 页面设置
        template.PageSettings = New PageConfig With {
            .Margins = New MarginsConfig(2.54, 2.54, 3.18, 3.18),
            .Header = New HeaderFooterConfig(False),
            .Footer = New HeaderFooterConfig(False),
            .PageNumber = New PageNumberConfig(True, "footer", "center", "{page}")
        }

        Return template
    End Function

    ''' <summary>创建商务报告模板</summary>
    Private Function CreateBusinessReportTemplate() As ReformatTemplate
        Dim template As New ReformatTemplate With {
            .Id = "preset-business",
            .Name = "商务报告模板",
            .Description = "适用于商务报告、项目提案、可行性研究等",
            .Category = "商务",
            .TargetApp = "Word",
            .IsPreset = True,
            .AiGuidance = "现代商务报告风格。注重视觉层次，使用微软雅黑字体，标题采用蓝色系。支持图表、数据表格等商务元素。"
        }

        ' 版式配置
        template.Layout = New LayoutConfig()
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "报告标题",
            .ElementType = "text",
            .DefaultValue = "商务报告",
            .Required = True,
            .SortOrder = 1,
            .Font = New FontConfig("微软雅黑", "Arial", 22, True),
            .Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceAfter = 2},
            .Color = New ColorConfig("#2E5090")
        })
        template.Layout.Elements.Add(New LayoutElement With {
            .Name = "副标题",
            .ElementType = "text",
            .DefaultValue = "子标题",
            .Required = False,
            .SortOrder = 2,
            .Font = New FontConfig("微软雅黑", "Arial", 16),
            .Paragraph = New ParagraphConfig("center", 0, 1.5) With {.SpaceAfter = 1.5},
            .Color = New ColorConfig("#666666")
        })

        ' 正文样式
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "一级标题",
            .MatchCondition = "章节标题或包含数字编号的标题",
            .SortOrder = 1,
            .Font = New FontConfig("微软雅黑", "Arial", 18, True),
            .Paragraph = New ParagraphConfig("left", 0, 1.5) With {.SpaceBefore = 1, .SpaceAfter = 0.5},
            .Color = New ColorConfig("#2E5090")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "二级标题",
            .MatchCondition = "二级标题或包含次级编号",
            .SortOrder = 2,
            .Font = New FontConfig("微软雅黑", "Arial", 14, True),
            .Paragraph = New ParagraphConfig("left", 0, 1.5) With {.SpaceBefore = 0.5, .SpaceAfter = 0.5},
            .Color = New ColorConfig("#000000")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "正文",
            .MatchCondition = "默认正文段落",
            .SortOrder = 3,
            .Font = New FontConfig("微软雅黑", "Arial", 11),
            .Paragraph = New ParagraphConfig("justify", 2, 1.5),
            .Color = New ColorConfig("#000000")
        })
        template.BodyStyles.Add(New StyleRule With {
            .RuleName = "要点列表",
            .MatchCondition = "包含'·'或'•'",
            .SortOrder = 4,
            .Font = New FontConfig("微软雅黑", "Arial", 11),
            .Paragraph = New ParagraphConfig("left", 0, 1.5) With {.LeftIndent = 1},
            .Color = New ColorConfig("#000000")
        })

        ' 页面设置
        template.PageSettings = New PageConfig With {
            .Margins = New MarginsConfig(2.5, 2.5, 2.5, 2.5),
            .Header = New HeaderFooterConfig(True, "商务报告", "right"),
            .Footer = New HeaderFooterConfig(False),
            .PageNumber = New PageNumberConfig(True, "footer", "right", "第{page}页 共{total}页")
        }

        Return template
    End Function

#End Region

End Class
