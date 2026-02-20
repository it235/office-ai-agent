' ShareRibbon\Services\Reformat\FormatPreviewService.vb
' 排版预览和应用服务

Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports Newtonsoft.Json

''' <summary>
''' 排版预览结果
''' </summary>
Public Class FormatPreviewResult
    ''' <summary>是否成功</summary>
    Public Property Success As Boolean = False
    ''' <summary>预览HTML内容</summary>
    Public Property PreviewHtml As String = ""
    ''' <summary>应用前的内容快照</summary>
    Public Property BeforeSnapshot As String = ""
    ''' <summary>应用后的内容快照</summary>
    Public Property AfterSnapshot As String = ""
    ''' <summary>修改统计</summary>
    Public Property ChangeStats As FormatChangeStats
    ''' <summary>错误信息</summary>
    Public Property ErrorMessage As String = ""

    Public Sub New()
        ChangeStats = New FormatChangeStats()
    End Sub
End Class

''' <summary>
''' 排版修改统计
''' </summary>
Public Class FormatChangeStats
    ''' <summary>修改的段落数</summary>
    Public Property ParagraphsModified As Integer = 0
    ''' <summary>修改的样式数</summary>
    Public Property StylesModified As Integer = 0
    ''' <summary>添加的元素数</summary>
    Public Property ElementsAdded As Integer = 0
    ''' <summary>删除的元素数</summary>
    Public Property ElementsRemoved As Integer = 0

    Public Sub New()
    End Sub
End Class

''' <summary>
''' 排版应用服务基类
''' </summary>
Public MustInherit Class FormatApplicationService
    ''' <summary>
    ''' 生成排版预览
    ''' </summary>
    Public MustOverride Function GeneratePreview(template As ReformatTemplate) As FormatPreviewResult

    ''' <summary>
    ''' 应用排版模板到当前文档
    ''' </summary>
    Public MustOverride Function ApplyTemplate(template As ReformatTemplate) As Boolean

    ''' <summary>
    ''' 应用排版模板到当前文档（带确认）
    ''' </summary>
    Public Function ApplyTemplateWithConfirmation(template As ReformatTemplate) As Boolean
        Dim preview = GeneratePreview(template)
        If Not preview.Success Then
            MessageBox.Show($"生成预览失败：{preview.ErrorMessage}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        Dim confirmMsg = $"即将应用模板 '{template.Name}'{vbCrLf}" &
                         $"预计修改：{preview.ChangeStats.ParagraphsModified} 个段落，{preview.ChangeStats.StylesModified} 个样式" &
                         $"{vbCrLf}{vbCrLf}是否继续？"

        Dim result = MessageBox.Show(confirmMsg, "确认应用", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result <> DialogResult.OK Then
            Return False
        End If

        Return ApplyTemplate(template)
    End Function

    ''' <summary>
    ''' 生成HTML预览
    ''' </summary>
    Protected Function GenerateHtmlPreview(template As ReformatTemplate, contentPreview As String) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html>")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset='utf-8'>")
        sb.AppendLine("<title>排版预览</title>")
        sb.AppendLine("<style>")
        sb.AppendLine("body { font-family: 'Microsoft YaHei', sans-serif; padding: 20px; background: #f5f5f5; }")
        sb.AppendLine(".preview-container { background: white; padding: 40px; max-width: 800px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }")
        sb.AppendLine(".element { margin: 10px 0; padding: 5px; border-left: 3px solid #4CAF50; }")
        sb.AppendLine(".element-header { font-size: 12px; color: #666; margin-bottom: 5px; }")
        sb.AppendLine(".style-info { font-size: 11px; color: #999; margin-top: 5px; }")
        sb.AppendLine("</style>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")
        sb.AppendLine("<div class='preview-container'>")

        sb.AppendLine($"<h2>模板：{template.Name}</h2>")
        sb.AppendLine($"<p style='color:#666;'>{template.Description}</p>")
        sb.AppendLine("<hr>")

        sb.AppendLine("<h3>版式元素：</h3>")
        For Each element In template.Layout.Elements.OrderBy(Function(e) e.SortOrder)
            sb.AppendLine("<div class='element'>")
            sb.AppendLine($"<div class='element-header'>【{element.Name}】({element.ElementType})</div>")
            sb.AppendLine($"<div style='{GetInlineStyle(element.Font, element.Paragraph, element.Color)}'>")
            sb.AppendLine(If(String.IsNullOrEmpty(element.DefaultValue), "(占位符)", element.DefaultValue))
            sb.AppendLine("</div>")
            sb.AppendLine($"<div class='style-info'>字体：{element.Font.FontNameCN} {element.Font.FontSize}pt | 对齐：{element.Paragraph.Alignment}</div>")
            sb.AppendLine("</div>")
        Next

        sb.AppendLine("<h3>样式规则：</h3>")
        For Each rule In template.BodyStyles.OrderBy(Function(r) r.SortOrder)
            sb.AppendLine("<div class='element' style='border-left-color:#2196F3;'>")
            sb.AppendLine($"<div class='element-header'>【{rule.RuleName}】</div>")
            sb.AppendLine($"<div style='{GetInlineStyle(rule.Font, rule.Paragraph, rule.Color)}'>")
            sb.AppendLine("这是应用该样式的示例文本")
            sb.AppendLine("</div>")
            sb.AppendLine($"<div class='style-info'>匹配条件：{If(String.IsNullOrEmpty(rule.MatchCondition), "无", rule.MatchCondition)}</div>")
            sb.AppendLine("</div>")
        Next

        If Not String.IsNullOrEmpty(contentPreview) Then
            sb.AppendLine("<h3>内容预览：</h3>")
            sb.AppendLine("<div style='background:#f9f9f9;padding:15px;border:1px solid #eee;'>")
            sb.AppendLine(contentPreview)
            sb.AppendLine("</div>")
        End If

        sb.AppendLine("</div>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 获取内联样式
    ''' </summary>
    Private Function GetInlineStyle(font As FontConfig, para As ParagraphConfig, color As ColorConfig) As String
        Dim styles As New List(Of String)()

        If font IsNot Nothing Then
            styles.Add($"font-family: '{font.FontNameCN}', '{font.FontNameEN}', sans-serif")
            styles.Add($"font-size: {font.FontSize}pt")
            If font.Bold Then styles.Add("font-weight: bold")
            If font.Italic Then styles.Add("font-style: italic")
            If font.Underline Then styles.Add("text-decoration: underline")
        End If

        If para IsNot Nothing Then
            Select Case para.Alignment
                Case "center" : styles.Add("text-align: center")
                Case "right" : styles.Add("text-align: right")
                Case "justify" : styles.Add("text-align: justify")
                Case Else : styles.Add("text-align: left")
            End Select
            styles.Add($"line-height: {para.LineSpacing}")
            If para.FirstLineIndent > 0 Then
                styles.Add($"text-indent: {para.FirstLineIndent}em")
            End If
        End If

        If color IsNot Nothing AndAlso Not String.IsNullOrEmpty(color.FontColor) Then
            styles.Add($"color: {color.FontColor}")
        End If

        Return String.Join("; ", styles)
    End Function
End Class
