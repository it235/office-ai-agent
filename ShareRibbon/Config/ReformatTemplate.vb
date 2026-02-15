' ShareRibbon\Config\ReformatTemplate.vb
' 排版模板数据模型

Imports Newtonsoft.Json

''' <summary>
''' 排版模板数据模型
''' </summary>
Public Class ReformatTemplate
    ''' <summary>模板ID（唯一标识）</summary>
    Public Property Id As String = Guid.NewGuid().ToString()

    ''' <summary>模板名称</summary>
    Public Property Name As String = ""

    ''' <summary>模板描述</summary>
    Public Property Description As String = ""

    ''' <summary>模板类型（通用、行政、学术、商务等）</summary>
    Public Property Category As String = "通用"

    ''' <summary>适用的Office应用（Word, PowerPoint, 或两者）</summary>
    Public Property TargetApp As String = "Word"

    ''' <summary>是否为预置模板（预置模板不可删除）</summary>
    Public Property IsPreset As Boolean = False

    ''' <summary>创建时间</summary>
    Public Property CreatedAt As DateTime = DateTime.Now

    ''' <summary>最后修改时间</summary>
    Public Property LastModified As DateTime = DateTime.Now

    ''' <summary>版式（骨架）配置</summary>
    Public Property Layout As LayoutConfig

    ''' <summary>正文样式规则集合</summary>
    Public Property BodyStyles As List(Of StyleRule)

    ''' <summary>页面设置</summary>
    Public Property PageSettings As PageConfig

    ''' <summary>AI模板补充说明（给AI的额外上下文）</summary>
    Public Property AiGuidance As String = ""

    ''' <summary>预览缩略图Base64（可选）</summary>
    Public Property ThumbnailBase64 As String = ""

    ''' <summary>模板来源类型</summary>
    Public Property TemplateSource As TemplateSourceType = TemplateSourceType.Manual

    ''' <summary>源文件名（AI提取时记录）</summary>
    Public Property SourceFileName As String = ""

    ''' <summary>源文件内容摘要（用于AI分析时的参考）</summary>
    Public Property SourceFileContent As String = ""

    Public Sub New()
        Layout = New LayoutConfig()
        BodyStyles = New List(Of StyleRule)()
        PageSettings = New PageConfig()
    End Sub
End Class

''' <summary>
''' 版式配置（文档骨架）
''' </summary>
Public Class LayoutConfig
    ''' <summary>骨架元素集合</summary>
    Public Property Elements As List(Of LayoutElement)

    Public Sub New()
        Elements = New List(Of LayoutElement)()
    End Sub
End Class

''' <summary>
''' 骨架元素（如发文机关、标题、红线等）
''' </summary>
Public Class LayoutElement
    ''' <summary>元素名称（发文机关、发文字号、文件标题、红色横线等）</summary>
    Public Property Name As String = ""

    ''' <summary>元素类型（text, redLine, separator）</summary>
    Public Property ElementType As String = "text"

    ''' <summary>默认值</summary>
    Public Property DefaultValue As String = ""

    ''' <summary>是否必需</summary>
    Public Property Required As Boolean = True

    ''' <summary>排序顺序</summary>
    Public Property SortOrder As Integer = 0

    ''' <summary>字体配置</summary>
    Public Property Font As FontConfig

    ''' <summary>段落配置</summary>
    Public Property Paragraph As ParagraphConfig

    ''' <summary>颜色配置</summary>
    Public Property Color As ColorConfig

    ''' <summary>特殊属性（如红线的线宽、高度等）</summary>
    Public Property SpecialProps As Dictionary(Of String, String)
    
    ''' <summary>占位符内容模板（支持{{}}变量替换）</summary>
    Public Property PlaceholderContent As String = "{{content}}"

    Public Sub New()
        Font = New FontConfig()
        Paragraph = New ParagraphConfig()
        Color = New ColorConfig()
        SpecialProps = New Dictionary(Of String, String)()
    End Sub
    
    Public Overrides Function ToString() As String
        Return If(String.IsNullOrEmpty(Name), "(未命名)", Name)
    End Function
End Class

''' <summary>
''' 正文样式规则
''' </summary>
Public Class StyleRule
    ''' <summary>规则名称（一级标题、二级标题、正文等）</summary>
    Public Property RuleName As String = ""

    ''' <summary>匹配条件描述（如"字号>20pt"、"包含'第X章'"）</summary>
    Public Property MatchCondition As String = ""

    ''' <summary>排序顺序（用于规则匹配优先级）</summary>
    Public Property SortOrder As Integer = 0

    ''' <summary>字体配置</summary>
    Public Property Font As FontConfig

    ''' <summary>段落配置</summary>
    Public Property Paragraph As ParagraphConfig

    ''' <summary>颜色配置</summary>
    Public Property Color As ColorConfig

    Public Sub New()
        Font = New FontConfig()
        Paragraph = New ParagraphConfig()
        Color = New ColorConfig()
    End Sub
    
    Public Overrides Function ToString() As String
        Return If(String.IsNullOrEmpty(RuleName), "(未命名)", RuleName)
    End Function
End Class

''' <summary>
''' 字体配置
''' </summary>
Public Class FontConfig
    ''' <summary>中文字体</summary>
    Public Property FontNameCN As String = "宋体"

    ''' <summary>英文字体</summary>
    Public Property FontNameEN As String = "Times New Roman"

    ''' <summary>字号（pt）</summary>
    Public Property FontSize As Double = 12

    ''' <summary>是否加粗</summary>
    Public Property Bold As Boolean = False

    ''' <summary>是否斜体</summary>
    Public Property Italic As Boolean = False

    ''' <summary>下划线</summary>
    Public Property Underline As Boolean = False

    Public Sub New()
    End Sub

    Public Sub New(fontNameCN As String, fontNameEN As String, fontSize As Double, Optional bold As Boolean = False)
        Me.FontNameCN = fontNameCN
        Me.FontNameEN = fontNameEN
        Me.FontSize = fontSize
        Me.Bold = bold
    End Sub
End Class

''' <summary>
''' 段落配置
''' </summary>
Public Class ParagraphConfig
    ''' <summary>对齐方式（left, center, right, justify）</summary>
    Public Property Alignment As String = "left"

    ''' <summary>首行缩进（字符数）</summary>
    Public Property FirstLineIndent As Double = 0

    ''' <summary>左缩进（cm）</summary>
    Public Property LeftIndent As Double = 0

    ''' <summary>右缩进（cm）</summary>
    Public Property RightIndent As Double = 0

    ''' <summary>段前间距（行）</summary>
    Public Property SpaceBefore As Double = 0

    ''' <summary>段后间距（行）</summary>
    Public Property SpaceAfter As Double = 0

    ''' <summary>行距（1.0, 1.5, 2.0等）</summary>
    Public Property LineSpacing As Double = 1.5

    Public Sub New()
    End Sub

    Public Sub New(alignment As String, Optional firstLineIndent As Double = 0, Optional lineSpacing As Double = 1.5)
        Me.Alignment = alignment
        Me.FirstLineIndent = firstLineIndent
        Me.LineSpacing = lineSpacing
    End Sub
End Class

''' <summary>
''' 颜色配置
''' </summary>
Public Class ColorConfig
    ''' <summary>字体颜色（RGB十六进制，如"#000000"）</summary>
    Public Property FontColor As String = "#000000"

    ''' <summary>背景色（可选）</summary>
    Public Property BackgroundColor As String = ""

    Public Sub New()
    End Sub

    Public Sub New(fontColor As String)
        Me.FontColor = fontColor
    End Sub
End Class

''' <summary>
''' 页面设置
''' </summary>
Public Class PageConfig
    ''' <summary>页边距（上、下、左、右，单位cm）</summary>
    Public Property Margins As MarginsConfig

    ''' <summary>页眉配置</summary>
    Public Property Header As HeaderFooterConfig

    ''' <summary>页脚配置</summary>
    Public Property Footer As HeaderFooterConfig

    ''' <summary>页码显示配置</summary>
    Public Property PageNumber As PageNumberConfig

    Public Sub New()
        Margins = New MarginsConfig()
        Header = New HeaderFooterConfig()
        Footer = New HeaderFooterConfig()
        PageNumber = New PageNumberConfig()
    End Sub
End Class

''' <summary>
''' 页边距配置
''' </summary>
Public Class MarginsConfig
    Public Property Top As Double = 2.54
    Public Property Bottom As Double = 2.54
    Public Property Left As Double = 3.18
    Public Property Right As Double = 3.18

    Public Sub New()
    End Sub

    Public Sub New(top As Double, bottom As Double, left As Double, right As Double)
        Me.Top = top
        Me.Bottom = bottom
        Me.Left = left
        Me.Right = right
    End Sub
End Class

''' <summary>
''' 页眉页脚配置
''' </summary>
Public Class HeaderFooterConfig
    Public Property Enabled As Boolean = False
    Public Property Content As String = ""
    Public Property Alignment As String = "center"

    Public Sub New()
    End Sub

    Public Sub New(enabled As Boolean, Optional content As String = "", Optional alignment As String = "center")
        Me.Enabled = enabled
        Me.Content = content
        Me.Alignment = alignment
    End Sub
End Class

''' <summary>
''' 页码配置
''' </summary>
Public Class PageNumberConfig
    Public Property Enabled As Boolean = True
    Public Property Position As String = "footer" ' header/footer
    Public Property Alignment As String = "center"
    Public Property Format As String = "第{page}页 共{total}页"

    Public Sub New()
    End Sub

    Public Sub New(enabled As Boolean, Optional position As String = "footer", Optional alignment As String = "center", Optional format As String = "第{page}页")
        Me.Enabled = enabled
        Me.Position = position
        Me.Alignment = alignment
        Me.Format = format
    End Sub
End Class
