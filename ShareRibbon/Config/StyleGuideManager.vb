' ShareRibbon\Config\StyleGuideManager.vb
' 排版规范管理器（单例模式）

Imports System.IO
Imports Newtonsoft.Json

''' <summary>
''' 排版规范管理器（单例模式）
''' </summary>
Public Class StyleGuideManager
    Private Shared _instance As StyleGuideManager
    Private _styleGuides As List(Of StyleGuideResource)
    Private ReadOnly _configPath As String

    ''' <summary>获取单例实例</summary>
    Public Shared ReadOnly Property Instance As StyleGuideManager
        Get
            If _instance Is Nothing Then
                _instance = New StyleGuideManager()
            End If
            Return _instance
        End Get
    End Property

    ''' <summary>获取所有规范</summary>
    Public ReadOnly Property StyleGuides As List(Of StyleGuideResource)
        Get
            Return _styleGuides
        End Get
    End Property

    Private Sub New()
        _configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            "styleguides.json")
        LoadStyleGuides()
    End Sub

    ''' <summary>加载规范配置</summary>
    Private Sub LoadStyleGuides()
        _styleGuides = New List(Of StyleGuideResource)()

        If File.Exists(_configPath) Then
            Try
                Dim json = File.ReadAllText(_configPath, Text.Encoding.UTF8)
                Dim loadedGuides = JsonConvert.DeserializeObject(Of List(Of StyleGuideResource))(json)

                ' 合并预置规范和用户规范
                MergePresetsAndUserGuides(loadedGuides)
            Catch ex As Exception
                Debug.WriteLine($"加载规范配置失败: {ex.Message}")
                LoadPresetStyleGuides()
            End Try
        Else
            ' 首次使用，加载预置规范
            LoadPresetStyleGuides()
            SaveStyleGuides()
        End If
    End Sub

    ''' <summary>合并预置规范和用户规范</summary>
    Private Sub MergePresetsAndUserGuides(userGuides As List(Of StyleGuideResource))
        If userGuides Is Nothing OrElse userGuides.Count = 0 Then
            LoadPresetStyleGuides()
            Return
        End If

        ' 获取预置规范ID列表
        Dim presets = GetPresetStyleGuides()
        Dim presetIds = presets.Select(Function(p) p.Id).ToHashSet()

        ' 先处理预置规范：如果用户数据中有同ID的（可能被修改过），使用用户版本；否则使用默认预置
        For Each preset In presets
            Dim userVersion = userGuides.FirstOrDefault(Function(g) g.Id = preset.Id)
            If userVersion IsNot Nothing Then
                ' 使用用户版本（可能被修改过），确保标记为预置
                userVersion.IsPreset = True
                _styleGuides.Add(userVersion)
            Else
                ' 用户数据中没有此预置规范（可能是新增的预置），添加默认版本
                _styleGuides.Add(preset)
            End If
        Next

        ' 再添加用户自定义规范（非预置的）
        For Each userGuide In userGuides
            ' 跳过预置规范ID（已在上面处理过）
            If presetIds.Contains(userGuide.Id) Then
                Continue For
            End If
            ' 添加用户自定义规范
            If Not _styleGuides.Any(Function(g) g.Id = userGuide.Id) Then
                userGuide.IsPreset = False ' 确保用户规范不被标记为预置
                _styleGuides.Add(userGuide)
            End If
        Next
    End Sub

    ''' <summary>加载预置规范</summary>
    Private Sub LoadPresetStyleGuides()
        _styleGuides.AddRange(GetPresetStyleGuides())
    End Sub

    ''' <summary>保存规范配置</summary>
    Public Sub SaveStyleGuides()
        Try
            Dim dir = Path.GetDirectoryName(_configPath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json = JsonConvert.SerializeObject(_styleGuides, Formatting.Indented)
            File.WriteAllText(_configPath, json, Text.Encoding.UTF8)
        Catch ex As Exception
            Debug.WriteLine($"保存规范配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>添加规范</summary>
    Public Sub AddStyleGuide(guide As StyleGuideResource)
        If String.IsNullOrEmpty(guide.Id) Then
            guide.Id = Guid.NewGuid().ToString()
        End If
        guide.CreatedAt = DateTime.Now
        guide.LastModified = DateTime.Now
        guide.IsPreset = False
        _styleGuides.Add(guide)
        SaveStyleGuides()
    End Sub

    ''' <summary>更新规范</summary>
    Public Sub UpdateStyleGuide(guide As StyleGuideResource)
        Dim existing = _styleGuides.FirstOrDefault(Function(g) g.Id = guide.Id)
        If existing IsNot Nothing Then
            guide.LastModified = DateTime.Now
            Dim index = _styleGuides.IndexOf(existing)
            _styleGuides(index) = guide
            SaveStyleGuides()
        End If
    End Sub

    ''' <summary>删除规范</summary>
    Public Function DeleteStyleGuide(guideId As String) As Boolean
        Dim guide = _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
        If guide IsNot Nothing Then
            If guide.IsPreset Then
                Return False ' 预置规范不可删除
            End If
            _styleGuides.Remove(guide)
            SaveStyleGuides()
            Return True
        End If
        Return False
    End Function

    ''' <summary>复制规范</summary>
    Public Function DuplicateStyleGuide(guideId As String, newName As String) As StyleGuideResource
        Dim original = _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
        If original IsNot Nothing Then
            ' 深拷贝
            Dim json = JsonConvert.SerializeObject(original)
            Dim duplicate = JsonConvert.DeserializeObject(Of StyleGuideResource)(json)
            duplicate.Id = Guid.NewGuid().ToString()
            duplicate.Name = If(String.IsNullOrEmpty(newName), original.Name & " (副本)", newName)
            duplicate.IsPreset = False
            duplicate.CreatedAt = DateTime.Now
            duplicate.LastModified = DateTime.Now
            _styleGuides.Add(duplicate)
            SaveStyleGuides()
            Return duplicate
        End If
        Return Nothing
    End Function

    ''' <summary>导出规范到文件</summary>
    Public Function ExportStyleGuide(guideId As String, filePath As String) As Boolean
        Try
            Dim guide = _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
            If guide IsNot Nothing Then
                ' 导出为原始格式（txt/md）
                Dim extension = If(String.IsNullOrEmpty(guide.SourceFileExtension), ".md", guide.SourceFileExtension)
                Dim actualPath = If(filePath.EndsWith(extension), filePath, filePath & extension)
                File.WriteAllText(actualPath, guide.GuideContent, Text.Encoding.UTF8)
                Return True
            End If
        Catch ex As Exception
            Debug.WriteLine($"导出规范失败: {ex.Message}")
        End Try
        Return False
    End Function

    ''' <summary>从文件导入规范</summary>
    Public Function ImportStyleGuide(filePath As String) As StyleGuideResource
        Try
            Dim content = File.ReadAllText(filePath, Text.Encoding.UTF8)
            Dim guide As New StyleGuideResource()
            guide.Id = Guid.NewGuid().ToString()
            guide.Name = Path.GetFileNameWithoutExtension(filePath)
            guide.GuideContent = content
            guide.SourceFileName = Path.GetFileName(filePath)
            guide.SourceFileExtension = Path.GetExtension(filePath)
            guide.FileEncoding = "UTF-8"
            guide.Category = "通用"
            guide.IsPreset = False
            guide.CreatedAt = DateTime.Now
            guide.LastModified = DateTime.Now
            _styleGuides.Add(guide)
            SaveStyleGuides()
            Return guide
        Catch ex As Exception
            Debug.WriteLine($"导入规范失败: {ex.Message}")
        End Try
        Return Nothing
    End Function

    ''' <summary>根据ID获取规范</summary>
    Public Function GetStyleGuideById(guideId As String) As StyleGuideResource
        Return _styleGuides.FirstOrDefault(Function(g) g.Id = guideId)
    End Function

    ''' <summary>按分类获取规范</summary>
    Public Function GetStyleGuidesByCategory(category As String) As List(Of StyleGuideResource)
        If String.IsNullOrEmpty(category) OrElse category = "全部" Then
            Return _styleGuides.ToList()
        End If
        Return _styleGuides.Where(Function(g) g.Category = category).ToList()
    End Function

    ''' <summary>获取所有分类</summary>
    Public Function GetAllCategories() As List(Of String)
        Dim categories = _styleGuides.Select(Function(g) g.Category).Distinct().ToList()
        categories.Insert(0, "全部")
        Return categories
    End Function

    ''' <summary>获取所有规范（兼容方法名）</summary>
    Public Function GetAllStyleGuides() As List(Of StyleGuideResource)
        Return _styleGuides.ToList()
    End Function

    ''' <summary>刷新规范列表（重新从文件加载）</summary>
    Public Sub Refresh()
        LoadStyleGuides()
    End Sub

#Region "预置规范"

    ''' <summary>获取预置规范列表</summary>
    Private Function GetPresetStyleGuides() As List(Of StyleGuideResource)
        Dim presets As New List(Of StyleGuideResource)()

        ' 预置规范1：GB/T 9704-2012 党政机关公文格式规范
        presets.Add(CreateOfficialDocumentGuide())

        ' 预置规范2：学术论文排版规范
        presets.Add(CreateAcademicPaperGuide())

        ' 预置规范3：商务报告排版规范
        presets.Add(CreateBusinessReportGuide())

        Return presets
    End Function

    ''' <summary>创建党政机关公文格式规范</summary>
    Private Function CreateOfficialDocumentGuide() As StyleGuideResource
        Dim guide As New StyleGuideResource With {
            .Id = "preset-guide-official",
            .Name = "党政机关公文格式规范",
            .Description = "根据GB/T 9704-2012标准整理的党政机关公文格式规范",
            .Category = "行政",
            .TargetApp = "Word",
            .IsPreset = True,
            .SourceFileName = "GB_T_9704-2012.md",
            .SourceFileExtension = ".md",
            .FileEncoding = "UTF-8"
        }

        guide.GuideContent = "# 党政机关公文格式规范（GB/T 9704-2012）

## 一、页面设置

### 1. 纸张规格
- **尺寸**：A4（210mm × 297mm）
- **方向**：纵向

### 2. 页边距
- **上边距**：37mm（约3.7cm）
- **下边距**：35mm（约3.5cm）
- **左边距**：28mm（约2.8cm）
- **右边距**：26mm（约2.6cm）

### 3. 版心尺寸
- **宽度**：156mm
- **高度**：225mm

## 二、字体要求

### 1. 发文机关标志
- **字体**：方正小标宋简体
- **字号**：由发文机关自定，但应醒目美观
- **颜色**：红色（一般为RGB: 192, 0, 0）

### 2. 发文字号
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **对齐**：居中

### 3. 标题
- **字体**：方正小标宋简体
- **字号**：二号（22pt）
- **对齐**：居中
- **加粗**：是

### 4. 主送机关
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **对齐**：左对齐，顶格

### 5. 正文
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **对齐**：两端对齐
- **首行缩进**：2字符

### 6. 一级标题
- **字体**：黑体
- **字号**：三号（16pt）
- **编号格式**：一、二、三...

### 7. 二级标题
- **字体**：楷体_GB2312
- **字号**：三号（16pt）
- **编号格式**：（一）（二）（三）...

### 8. 三级标题
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）
- **编号格式**：1. 2. 3. ...
- **加粗**：是

## 三、段落格式

### 1. 行距
- **正文行距**：固定值28磅或1.5倍行距
- **段前间距**：0
- **段后间距**：0

### 2. 缩进
- **首行缩进**：2字符（约0.85cm）
- **左缩进**：0
- **右缩进**：0

## 四、特殊元素

### 1. 红色分隔线
- **位置**：发文机关标志下方
- **宽度**：与版心等宽（156mm）
- **粗细**：约2pt
- **颜色**：红色

### 2. 页码
- **位置**：页脚居中
- **格式**：—X—（如：—1—）
- **字体**：宋体
- **字号**：四号（14pt）

### 3. 成文日期
- **格式**：XXXX年X月X日
- **位置**：正文下方右侧
- **字体**：仿宋_GB2312
- **字号**：三号（16pt）

## 五、注意事项

1. 公文用纸一般采用80克以上白色胶版纸或复印纸
2. 公文如需标注紧急程度，应在公文首页左上角标注
3. 附件说明位于正文下方、成文日期上方
4. 印章应端正、居中，上不压正文、下要骑年盖月
"

        guide.Tags = New List(Of String) From {"公文", "行政", "GB/T 9704", "政府文件"}

        Return guide
    End Function

    ''' <summary>创建学术论文排版规范</summary>
    Private Function CreateAcademicPaperGuide() As StyleGuideResource
        Dim guide As New StyleGuideResource With {
            .Id = "preset-guide-academic",
            .Name = "学术论文排版规范",
            .Description = "通用学术论文排版规范，参考GB/T 7713标准",
            .Category = "学术",
            .TargetApp = "Word",
            .IsPreset = True,
            .SourceFileName = "academic_paper_guide.md",
            .SourceFileExtension = ".md",
            .FileEncoding = "UTF-8"
        }

        guide.GuideContent = "# 学术论文排版规范

## 一、页面设置

### 1. 纸张
- **尺寸**：A4（210mm × 297mm）
- **方向**：纵向

### 2. 页边距
- **上边距**：2.54cm（1英寸）
- **下边距**：2.54cm（1英寸）
- **左边距**：3.18cm（1.25英寸）
- **右边距**：3.18cm（1.25英寸）

## 二、标题层级

### 1. 论文题目
- **字体**：黑体
- **字号**：小二号（18pt）
- **对齐**：居中
- **加粗**：是
- **段后**：1行

### 2. 作者信息
- **字体**：宋体
- **字号**：小四号（12pt）
- **对齐**：居中
- **格式**：姓名（单位，城市 邮编）

### 3. 摘要
- **标题字体**：黑体
- **标题字号**：五号（10.5pt）
- **正文字体**：宋体
- **正文字号**：五号（10.5pt）
- **行距**：1.5倍
- **关键词**：3-5个，用分号分隔

### 4. 一级标题
- **字体**：黑体
- **字号**：四号（14pt）
- **对齐**：左对齐
- **加粗**：是
- **编号**：1、2、3...或一、二、三...
- **段前**：0.5行
- **段后**：0.5行

### 5. 二级标题
- **字体**：黑体
- **字号**：小四号（12pt）
- **对齐**：左对齐
- **加粗**：是
- **编号**：1.1、1.2...或（一）（二）...

### 6. 三级标题
- **字体**：宋体
- **字号**：小四号（12pt）
- **对齐**：左对齐
- **加粗**：是
- **编号**：1.1.1、1.1.2...

## 三、正文格式

### 1. 字体字号
- **中文字体**：宋体
- **英文字体**：Times New Roman
- **字号**：小四号（12pt）

### 2. 段落格式
- **对齐方式**：两端对齐
- **首行缩进**：2字符
- **行距**：1.5倍行距
- **段前段后**：0

## 四、图表格式

### 1. 图片
- **图号**：图1、图2...（居中）
- **图题**：位于图片下方，居中
- **字体**：宋体，五号

### 2. 表格
- **表号**：表1、表2...（居中）
- **表题**：位于表格上方，居中
- **字体**：宋体，五号
- **表头**：加粗

## 五、参考文献

### 1. 格式（GB/T 7714-2015）
- **字体**：宋体
- **字号**：五号（10.5pt）
- **行距**：单倍行距
- **悬挂缩进**：2字符

### 2. 常见类型格式
- **期刊**：[序号] 作者.题名[J].刊名,年,卷(期):起止页码.
- **专著**：[序号] 作者.书名[M].版次.出版地:出版者,出版年:起止页码.
- **论文集**：[序号] 作者.题名[C]//论文集名.出版地:出版者,出版年:起止页码.
- **学位论文**：[序号] 作者.题名[D].保存地:保存单位,年份.
- **网络文献**：[序号] 作者.题名[EB/OL].(发布日期)[引用日期].网址.

## 六、页码设置

- **位置**：页脚居中
- **格式**：阿拉伯数字
- **字体**：宋体
- **字号**：五号
- **首页**：可不显示页码
"

        guide.Tags = New List(Of String) From {"学术", "论文", "期刊", "毕业论文"}

        Return guide
    End Function

    ''' <summary>创建商务报告排版规范</summary>
    Private Function CreateBusinessReportGuide() As StyleGuideResource
        Dim guide As New StyleGuideResource With {
            .Id = "preset-guide-business",
            .Name = "商务报告排版规范",
            .Description = "现代商务报告排版规范，适用于项目报告、分析报告等",
            .Category = "商务",
            .TargetApp = "Word",
            .IsPreset = True,
            .SourceFileName = "business_report_guide.md",
            .SourceFileExtension = ".md",
            .FileEncoding = "UTF-8"
        }

        guide.GuideContent = "# 商务报告排版规范

## 一、页面设置

### 1. 纸张
- **尺寸**：A4
- **方向**：纵向（特殊情况可用横向）

### 2. 页边距
- **上边距**：2.5cm
- **下边距**：2.5cm
- **左边距**：2.5cm
- **右边距**：2.5cm

## 二、字体规范

### 1. 主标题（报告封面）
- **字体**：微软雅黑
- **字号**：小一号（24pt）
- **颜色**：深蓝色（#2E5090）
- **对齐**：居中
- **加粗**：是

### 2. 副标题
- **字体**：微软雅黑
- **字号**：三号（16pt）
- **颜色**：灰色（#666666）
- **对齐**：居中

### 3. 章节标题（一级）
- **字体**：微软雅黑
- **字号**：小二号（18pt）
- **颜色**：深蓝色（#2E5090）
- **对齐**：左对齐
- **加粗**：是
- **段前**：1行
- **段后**：0.5行

### 4. 小节标题（二级）
- **字体**：微软雅黑
- **字号**：四号（14pt）
- **颜色**：黑色
- **对齐**：左对齐
- **加粗**：是
- **段前**：0.5行
- **段后**：0.5行

### 5. 正文
- **字体**：微软雅黑
- **字号**：小四号（12pt）
- **颜色**：黑色（#333333）
- **对齐**：两端对齐
- **首行缩进**：2字符
- **行距**：1.5倍

## 三、特殊元素

### 1. 要点列表
- **符号**：实心圆点（•）或短横线（-）
- **缩进**：1cm
- **行距**：1.2倍
- **颜色**：可使用品牌色标记重点

### 2. 数据表格
- **表头背景**：浅蓝色（#E6F0FA）
- **边框**：浅灰色（#CCCCCC），1pt
- **单元格内边距**：上下5pt，左右8pt
- **表头字体**：微软雅黑，加粗
- **数据字体**：微软雅黑，常规

### 3. 图表说明
- **位置**：图表下方
- **字体**：微软雅黑
- **字号**：五号（10.5pt）
- **颜色**：灰色
- **对齐**：居中
- **格式**：图X：图表标题

### 4. 引用框
- **背景色**：浅灰色（#F5F5F5）
- **边框**：左侧4pt实线，品牌色
- **内边距**：15pt
- **字体**：微软雅黑，斜体

## 四、页眉页脚

### 1. 页眉
- **内容**：报告名称或公司名称
- **字体**：微软雅黑
- **字号**：小五号（9pt）
- **颜色**：灰色
- **位置**：右对齐
- **分隔线**：细线（0.5pt）

### 2. 页脚
- **内容**：页码
- **格式**：第X页 共X页
- **字体**：微软雅黑
- **字号**：小五号（9pt）
- **位置**：居中或右对齐

## 五、颜色规范

### 推荐配色方案
- **主色**：#2E5090（深蓝）
- **辅色**：#4A90D9（亮蓝）
- **强调色**：#F5A623（橙色）
- **正文色**：#333333（深灰）
- **次要文字**：#666666（中灰）
- **背景色**：#FFFFFF（白色）

## 六、注意事项

1. 保持整体风格统一，不宜使用超过3种主要颜色
2. 重要数据使用图表可视化呈现
3. 每页内容不宜过于拥挤，适当留白
4. 文字与图表的比例建议为6:4
5. 关键结论使用加粗或颜色突出显示
"

        guide.Tags = New List(Of String) From {"商务", "报告", "企业", "项目"}

        Return guide
    End Function

#End Region

End Class
