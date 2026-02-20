' ShareRibbon\Services\Reformat\DocumentParserService.vb
' 文档智能解析服务

Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json

''' <summary>
''' 文档解析结果
''' </summary>
Public Class DocumentParseResult
    ''' <summary>是否成功</summary>
    Public Property Success As Boolean = False
    ''' <summary>解析出的文档元素</summary>
    Public Property Elements As List(Of ParsedDocumentElement)
    ''' <summary>解析出的样式</summary>
    Public Property Styles As List(Of ParsedStyle)
    ''' <summary>页面设置</summary>
    Public Property PageSettings As PageConfig
    ''' <summary>错误信息</summary>
    Public Property ErrorMessage As String = ""
    ''' <summary>原始文件内容（摘要）</summary>
    Public Property RawContentSummary As String = ""

    Public Sub New()
        Elements = New List(Of ParsedDocumentElement)()
        Styles = New List(Of ParsedStyle)()
        PageSettings = New PageConfig()
    End Sub
End Class

''' <summary>
''' 解析出的文档元素
''' </summary>
Public Class ParsedDocumentElement
    ''' <summary>元素名称</summary>
    Public Property Name As String = ""
    ''' <summary>元素类型</summary>
    Public Property ElementType As String = "text"
    ''' <summary>元素内容</summary>
    Public Property Content As String = ""
    ''' <summary>在文档中的位置</summary>
    Public Property Position As Integer = 0
    ''' <summary>关联的样式ID</summary>
    Public Property StyleId As String = ""
    ''' <summary>是否为标题</summary>
    Public Property IsHeading As Boolean = False
    ''' <summary>标题级别（1-6）</summary>
    Public Property HeadingLevel As Integer = 0
    ''' <summary>特殊属性</summary>
    Public Property SpecialProps As Dictionary(Of String, String)

    Public Sub New()
        SpecialProps = New Dictionary(Of String, String)()
    End Sub
End Class

''' <summary>
''' 解析出的样式
''' </summary>
Public Class ParsedStyle
    ''' <summary>样式ID</summary>
    Public Property StyleId As String = ""
    ''' <summary>样式名称</summary>
    Public Property StyleName As String = ""
    ''' <summary>字体配置</summary>
    Public Property Font As FontConfig
    ''' <summary>段落配置</summary>
    Public Property Paragraph As ParagraphConfig
    ''' <summary>颜色配置</summary>
    Public Property Color As ColorConfig
    ''' <summary>使用次数</summary>
    Public Property UsageCount As Integer = 0
    ''' <summary>是否为内置样式</summary>
    Public Property IsBuiltIn As Boolean = False

    Public Sub New()
        Font = New FontConfig()
        Paragraph = New ParagraphConfig()
        Color = New ColorConfig()
    End Sub
End Class

''' <summary>
''' 文档智能解析服务基类
''' </summary>
Public MustInherit Class DocumentParserService
    ''' <summary>
    ''' 从文件路径解析文档
    ''' </summary>
    Public MustOverride Function ParseDocument(filePath As String) As DocumentParseResult

    ''' <summary>
    ''' 从当前活动文档解析
    ''' </summary>
    Public MustOverride Function ParseActiveDocument() As DocumentParseResult

    ''' <summary>
    ''' 将解析结果转换为模板
    ''' </summary>
    Public Function ConvertToTemplate(parseResult As DocumentParseResult, templateName As String) As ReformatTemplate
        Dim template As New ReformatTemplate With {
            .Name = templateName,
            .Description = "从文档解析生成的模板",
            .Category = "自定义",
            .TemplateSource = TemplateSourceType.Parsed,
            .SourceFileContent = parseResult.RawContentSummary,
            .PageSettings = parseResult.PageSettings
        }

        Dim sortOrder = 1
        For Each element In parseResult.Elements
            Dim layoutElement As New LayoutElement With {
                .Name = element.Name,
                .ElementType = element.ElementType,
                .DefaultValue = element.Content,
                .Required = True,
                .SortOrder = sortOrder,
                .SpecialProps = element.SpecialProps
            }

            Dim matchingStyle = parseResult.Styles.FirstOrDefault(Function(s) s.StyleId = element.StyleId)
            If matchingStyle IsNot Nothing Then
                layoutElement.Font = matchingStyle.Font
                layoutElement.Paragraph = matchingStyle.Paragraph
                layoutElement.Color = matchingStyle.Color
            End If

            template.Layout.Elements.Add(layoutElement)
            sortOrder += 1
        Next

        Dim styleSortOrder = 1
        For Each style In parseResult.Styles.Where(Function(s) Not s.IsBuiltIn OrElse s.UsageCount > 0)
            Dim rule As New StyleRule With {
                .RuleName = style.StyleName,
                .MatchCondition = $"使用 {style.StyleName} 样式的段落",
                .SortOrder = styleSortOrder,
                .Font = style.Font,
                .Paragraph = style.Paragraph,
                .Color = style.Color
            }
            template.BodyStyles.Add(rule)
            styleSortOrder += 1
        Next

        Return template
    End Function

    ''' <summary>
    ''' 生成元素名称
    ''' </summary>
    Protected Function GenerateElementName(element As ParsedDocumentElement, index As Integer) As String
        If Not String.IsNullOrEmpty(element.Name) Then
            Return element.Name
        End If

        If element.IsHeading Then
            Return $"标题 {element.HeadingLevel}"
        End If

        Select Case element.ElementType
            Case "text"
                Return $"文本元素 {index}"
            Case "table"
                Return $"表格 {index}"
            Case "image"
                Return $"图片 {index}"
            Case "redLine"
                Return "红色横线"
            Case "separator"
                Return "分隔线"
            Case Else
                Return $"元素 {index}"
        End Select
    End Function
End Class
