' ShareRibbon\Config\StyleGuideResource.vb
' 排版规范资源数据模型

Imports Newtonsoft.Json

''' <summary>
''' 排版规范资源数据模型
''' 用于存储txt/markdown格式的排版规范说明文档
''' </summary>
Public Class StyleGuideResource
    ''' <summary>规范ID（唯一标识）</summary>
    Public Property Id As String = Guid.NewGuid().ToString()

    ''' <summary>规范名称</summary>
    Public Property Name As String = ""

    ''' <summary>规范描述</summary>
    Public Property Description As String = ""

    ''' <summary>规范分类（通用、行政、学术、商务等）</summary>
    Public Property Category As String = "通用"

    ''' <summary>是否为预置规范（预置规范不可删除）</summary>
    Public Property IsPreset As Boolean = False

    ''' <summary>创建时间</summary>
    Public Property CreatedAt As DateTime = DateTime.Now

    ''' <summary>最后修改时间</summary>
    Public Property LastModified As DateTime = DateTime.Now

    ''' <summary>规范文本内容（markdown/txt原文）</summary>
    Public Property GuideContent As String = ""

    ''' <summary>源文件名</summary>
    Public Property SourceFileName As String = ""

    ''' <summary>源文件扩展名（.txt/.md）</summary>
    Public Property SourceFileExtension As String = ""

    ''' <summary>文件编码</summary>
    Public Property FileEncoding As String = "UTF-8"

    ''' <summary>适用的Office应用（Word, PowerPoint, 或两者）</summary>
    Public Property TargetApp As String = "Word"

    ''' <summary>标签列表（用于搜索和分类）</summary>
    Public Property Tags As List(Of String)

    Public Sub New()
        Tags = New List(Of String)()
    End Sub

    ''' <summary>
    ''' 获取内容摘要（前100个字符）
    ''' </summary>
    Public ReadOnly Property ContentSummary As String
        Get
            If String.IsNullOrEmpty(GuideContent) Then Return ""
            Dim maxLength = Math.Min(100, GuideContent.Length)
            Dim summary = GuideContent.Substring(0, maxLength).Replace(vbCrLf, " ").Replace(vbLf, " ")
            If GuideContent.Length > 100 Then summary &= "..."
            Return summary
        End Get
    End Property

    ''' <summary>
    ''' 获取内容字数
    ''' </summary>
    Public ReadOnly Property ContentLength As Integer
        Get
            Return If(String.IsNullOrEmpty(GuideContent), 0, GuideContent.Length)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return If(String.IsNullOrEmpty(Name), "(未命名规范)", Name)
    End Function
End Class

''' <summary>
''' 资源类型枚举
''' </summary>
Public Enum ReformatResourceType
    ''' <summary>排版模板（结构化JSON配置）</summary>
    Template = 0
    ''' <summary>排版规范（文本规范说明）</summary>
    StyleGuide = 1
End Enum

''' <summary>
''' 模板来源类型枚举
''' </summary>
Public Enum TemplateSourceType
    ''' <summary>手动创建</summary>
    Manual
    ''' <summary>从文档解析</summary>
    Parsed
    ''' <summary>从规格生成</summary>
    FromSpec
    ''' <summary>预置模板</summary>
    Preset
End Enum
