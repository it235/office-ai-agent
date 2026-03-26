' ShareRibbon\Services\Reformat\FormatMirrorService.vb
' 格式克隆：分析现有文档/选区的实际格式，提取规则，让AI生成SemanticStyleMapping

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Microsoft.Office.Interop.Word

''' <summary>从文档段落提取的格式信息</summary>
Public Class ExtractedParagraphFormat
    Public Property StyleName As String = ""
    Public Property SampleText As String = ""
    Public Property FontNameCN As String = ""
    Public Property FontNameEN As String = ""
    Public Property FontSize As Double = 12
    Public Property Bold As Boolean = False
    Public Property Italic As Boolean = False
    Public Property AlignmentStr As String = "left"   ' left/center/right/justify
    Public Property FirstLineIndentCm As Double = 0
    Public Property LineSpacingPt As Double = 0
    Public Property SpaceBeforePt As Double = 0
    Public Property SpaceAfterPt As Double = 0
    Public Property OccurrenceCount As Integer = 1
End Class

Public Class FormatMirrorService

    Private Const MaxSamplesToExtract As Integer = 80   ' 最多采样段落数
    Private Const PointsPerCm As Double = 28.35

    ''' <summary>
    ''' 从 Word 文档中提取段落格式信息（全文或当前选区）
    ''' </summary>
    Public Shared Function ExtractFormattingFromDocument(
        wordApp As Application,
        selectionOnly As Boolean) As List(Of ExtractedParagraphFormat)

        Dim result As New List(Of ExtractedParagraphFormat)()
        Dim styleMap As New Dictionary(Of String, ExtractedParagraphFormat)()

        If wordApp Is Nothing OrElse wordApp.Documents.Count = 0 Then Return result

        Try
            Dim doc = wordApp.ActiveDocument
            Dim paragraphsToScan As IEnumerable(Of Paragraph)

            If selectionOnly AndAlso wordApp.Selection IsNot Nothing AndAlso
               wordApp.Selection.Range.Text.Trim().Length > 0 Then
                paragraphsToScan = wordApp.Selection.Range.Paragraphs.Cast(Of Paragraph)()
            Else
                paragraphsToScan = doc.Paragraphs.Cast(Of Paragraph)()
            End If

            Dim count As Integer = 0
            For Each p As Paragraph In paragraphsToScan
                If count >= MaxSamplesToExtract Then Exit For
                Dim txt = p.Range.Text.Replace(Chr(13), "").Replace(Chr(7), "").Trim()
                If String.IsNullOrWhiteSpace(txt) Then Continue For

                count += 1
                Dim styleName = p.Style.ToString()
                Dim key = styleName & "|" &
                          p.Range.Font.Size.ToString("F1") & "|" &
                          p.Range.Font.Bold.ToString() & "|" &
                          p.Alignment.ToString()

                If styleMap.ContainsKey(key) Then
                    styleMap(key).OccurrenceCount += 1
                    Continue For
                End If

                Dim fmt As New ExtractedParagraphFormat()
                fmt.StyleName = styleName
                fmt.SampleText = If(txt.Length > 60, txt.Substring(0, 60) & "…", txt)

                ' 字体
                Try
                    fmt.FontNameCN = If(String.IsNullOrEmpty(p.Range.Font.NameFarEast), "", p.Range.Font.NameFarEast)
                    fmt.FontNameEN = If(String.IsNullOrEmpty(p.Range.Font.Name), "", p.Range.Font.Name)
                    Dim sz = p.Range.Font.Size
                    fmt.FontSize = If(sz <= 0, 12, sz)
                    fmt.Bold = (p.Range.Font.Bold = CInt(True)) OrElse p.Range.Font.Bold = -1
                    fmt.Italic = (p.Range.Font.Italic = CInt(True)) OrElse p.Range.Font.Italic = -1
                Catch
                End Try

                ' 段落
                Try
                    Select Case p.Alignment
                        Case WdParagraphAlignment.wdAlignParagraphCenter : fmt.AlignmentStr = "center"
                        Case WdParagraphAlignment.wdAlignParagraphRight : fmt.AlignmentStr = "right"
                        Case WdParagraphAlignment.wdAlignParagraphJustify : fmt.AlignmentStr = "justify"
                        Case Else : fmt.AlignmentStr = "left"
                    End Select
                    Dim fi = p.FirstLineIndent
                    fmt.FirstLineIndentCm = If(fi > 0, Math.Round(fi / PointsPerCm, 2), 0)
                    fmt.LineSpacingPt = Math.Round(p.LineSpacing, 1)
                    fmt.SpaceBeforePt = Math.Round(p.SpaceBefore, 1)
                    fmt.SpaceAfterPt = Math.Round(p.SpaceAfter, 1)
                Catch
                End Try

                styleMap(key) = fmt
            Next

            ' 按出现次数排序，频率高的排前面
            result = styleMap.Values.OrderByDescending(Function(f) f.OccurrenceCount).ToList()

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"[FormatMirrorService] 提取失败: {ex.Message}")
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 构建让 AI 将提取的格式规则转换为 SemanticStyleMapping 的提示词
    ''' </summary>
    Public Shared Function BuildClonePrompt(extracted As List(Of ExtractedParagraphFormat)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("你是排版专家。根据以下从真实文档中提取的段落格式信息，生成一个SemanticStyleMapping JSON。")
        sb.AppendLine()
        sb.AppendLine("【可用语义标签】")
        sb.AppendLine("title.1（一级标题）、title.2（二级标题）、title.3（三级标题）")
        sb.AppendLine("body.normal（正文）、body.emphasis（强调正文）")
        sb.AppendLine("list.ordered（有序列表）、list.unordered（无序列表）")
        sb.AppendLine("quote（引用/摘要）、caption（图表说明）")
        sb.AppendLine()
        sb.AppendLine("【从文档提取的格式规则（按出现频率排序）】")

        For Each f In extracted.Take(20)
            sb.AppendLine($"- 样式名:{f.StyleName} | 出现{f.OccurrenceCount}次 | 样本:「{f.SampleText}」")
            sb.AppendLine($"  字体: CN={f.FontNameCN} EN={f.FontNameEN} 大小={f.FontSize}pt Bold={f.Bold} Italic={f.Italic}")
            sb.AppendLine($"  段落: 对齐={f.AlignmentStr} 首行={f.FirstLineIndentCm}cm 行距={f.LineSpacingPt}pt 前={f.SpaceBeforePt}pt 后={f.SpaceAfterPt}pt")
        Next

        sb.AppendLine()
        sb.AppendLine("【要求】")
        sb.AppendLine("1. 将每种格式映射到最合适的语义标签（body.normal 必须有）")
        sb.AppendLine("2. 仅返回如下JSON结构，不要解释：")
        sb.AppendLine("{""name"":""克隆格式"",""tags"":[{""tagId"":""title.1"",""font"":{""fontNameCN"":""...""")
        sb.AppendLine(",""fontNameEN"":""..."",""fontSize"":16,""bold"":true},""paragraph"":{""alignment"":""center""}}]}")
        sb.AppendLine("字段与 StyleGuideConverter 的输出格式完全相同。")

        Return sb.ToString()
    End Function

End Class
