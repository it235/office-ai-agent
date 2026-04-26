' ShareRibbon\Services\Reformat\FormatMirrorService.vb
' 格式克隆：分析现有文档/选区的实际格式，提取规则，让AI生成SemanticStyleMapping

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
' 注意：此服务专门用于 Word，使用 Object 类型和后期绑定以避免 ShareRibbon 直接依赖 Word Interop

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
    ''' <param name="wordApp">Word Application 对象（使用 Object 类型避免直接依赖 Word Interop）</param>
    ''' <param name="selectionOnly">是否仅处理选区</param>
    Public Shared Function ExtractFormattingFromDocument(
        wordApp As Object,
        selectionOnly As Boolean) As List(Of ExtractedParagraphFormat)

        Dim result As New List(Of ExtractedParagraphFormat)()
        Dim styleMap As New Dictionary(Of String, ExtractedParagraphFormat)()

        If wordApp Is Nothing Then Return result

        Try
            ' 使用动态绑定访问 Word 对象模型
            Dim documents As Object = wordApp.[GetType]().InvokeMember("Documents", Reflection.BindingFlags.GetProperty, Nothing, wordApp, Nothing)
            Dim docCount As Integer = CInt(documents.[GetType]().InvokeMember("Count", Reflection.BindingFlags.GetProperty, Nothing, documents, Nothing))
            If docCount = 0 Then Return result

            Dim doc As Object = wordApp.[GetType]().InvokeMember("ActiveDocument", Reflection.BindingFlags.GetProperty, Nothing, wordApp, Nothing)
            Dim paragraphs As Object = doc.[GetType]().InvokeMember("Paragraphs", Reflection.BindingFlags.GetProperty, Nothing, doc, Nothing)
            Dim paraCount As Integer = CInt(paragraphs.[GetType]().InvokeMember("Count", Reflection.BindingFlags.GetProperty, Nothing, paragraphs, Nothing))

            Dim count As Integer = 0
            For i As Integer = 1 To Math.Min(paraCount, MaxSamplesToExtract)
                Dim p As Object = paragraphs.[GetType]().InvokeMember("Item", Reflection.BindingFlags.GetProperty, Nothing, paragraphs, New Object() {i})
                Dim rangeObj As Object = p.[GetType]().InvokeMember("Range", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                Dim txt As String = CStr(rangeObj.[GetType]().InvokeMember("Text", Reflection.BindingFlags.GetProperty, Nothing, rangeObj, Nothing))
                txt = txt.Replace(Chr(13), "").Replace(Chr(7), "").Trim()
                If String.IsNullOrWhiteSpace(txt) Then Continue For

                count += 1
                Dim styleObj As Object = p.[GetType]().InvokeMember("Style", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                Dim styleName As String = styleObj.ToString()
                Dim fontObj As Object = rangeObj.[GetType]().InvokeMember("Font", Reflection.BindingFlags.GetProperty, Nothing, rangeObj, Nothing)
                Dim fontSize As Object = fontObj.[GetType]().InvokeMember("Size", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                Dim fontBold As Object = fontObj.[GetType]().InvokeMember("Bold", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                Dim alignment As Object = p.[GetType]().InvokeMember("Alignment", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)

                Dim key As String = styleName & "|" & fontSize.ToString() & "|" & fontBold.ToString() & "|" & alignment.ToString()

                If styleMap.ContainsKey(key) Then
                    styleMap(key).OccurrenceCount += 1
                    Continue For
                End If

                Dim fmt As New ExtractedParagraphFormat()
                fmt.StyleName = styleName
                fmt.SampleText = If(txt.Length > 60, txt.Substring(0, 60) & "…", txt)

                ' 字体信息
                Try
                    Dim nameFarEast As Object = fontObj.[GetType]().InvokeMember("NameFarEast", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                    Dim fontName As Object = fontObj.[GetType]().InvokeMember("Name", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                    fmt.FontNameCN = If(nameFarEast Is Nothing, "", nameFarEast.ToString())
                    fmt.FontNameEN = If(fontName Is Nothing, "", fontName.ToString())
                    fmt.FontSize = If(fontSize Is Nothing, 12, Convert.ToDouble(fontSize))
                    fmt.Bold = Convert.ToInt32(fontBold) = -1 OrElse Convert.ToInt32(fontBold) = 1
                    Dim fontItalic As Object = fontObj.[GetType]().InvokeMember("Italic", Reflection.BindingFlags.GetProperty, Nothing, fontObj, Nothing)
                    fmt.Italic = Convert.ToInt32(fontItalic) = -1 OrElse Convert.ToInt32(fontItalic) = 1
                Catch
                End Try

                ' 段落信息
                Try
                    Dim alignVal As Integer = Convert.ToInt32(alignment)
                    ' WdParagraphAlignment: 0=left, 1=center, 2=right, 3=justify
                    Select Case alignVal
                        Case 1 : fmt.AlignmentStr = "center"
                        Case 2 : fmt.AlignmentStr = "right"
                        Case 3 : fmt.AlignmentStr = "justify"
                        Case Else : fmt.AlignmentStr = "left"
                    End Select
                    Dim firstLineIndent As Object = p.[GetType]().InvokeMember("FirstLineIndent", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                    Dim lineSpacing As Object = p.[GetType]().InvokeMember("LineSpacing", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                    Dim spaceBefore As Object = p.[GetType]().InvokeMember("SpaceBefore", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)
                    Dim spaceAfter As Object = p.[GetType]().InvokeMember("SpaceAfter", Reflection.BindingFlags.GetProperty, Nothing, p, Nothing)

                    fmt.FirstLineIndentCm = If(firstLineIndent IsNot Nothing, Math.Round(Convert.ToDouble(firstLineIndent) / PointsPerCm, 2), 0)
                    fmt.LineSpacingPt = If(lineSpacing IsNot Nothing, Math.Round(Convert.ToDouble(lineSpacing), 1), 0)
                    fmt.SpaceBeforePt = If(spaceBefore IsNot Nothing, Math.Round(Convert.ToDouble(spaceBefore), 1), 0)
                    fmt.SpaceAfterPt = If(spaceAfter IsNot Nothing, Math.Round(Convert.ToDouble(spaceAfter), 1), 0)
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
        sb.AppendLine("{""name"":""克隆格式"",""semanticTags"":[{""tagId"":""title.1"",""font"":{""fontNameCN"":""..."",")
        sb.AppendLine("""fontNameEN"":""..."",""fontSize"":16,""bold"":true},""paragraph"":{""alignment"":""center""}}]}")
        sb.AppendLine("字段与 StyleGuideConverter 的输出格式完全相同（使用 semanticTags 字段）。")

        Return sb.ToString()
    End Function

End Class
