Imports System.Diagnostics
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports ShareRibbon
Imports Word = Microsoft.Office.Interop.Word

''' <summary>
''' Word文档翻译服务 - 支持一键翻译和沉浸式翻译
''' </summary>
Public Class WordDocumentTranslateService
    Inherits DocumentTranslateService

    Private _wordApp As Word.Application
    Private _document As Document
    Private _selectionRanges As List(Of Range)

    Public Sub New(wordApp As Word.Application)
        MyBase.New()
        _wordApp = wordApp
        _document = wordApp.ActiveDocument
    End Sub

    ''' <summary>
    ''' 获取文档所有段落
    ''' </summary>
    Public Overrides Function GetAllParagraphs() As List(Of String)
        Dim paragraphs As New List(Of String)()

        If _document Is Nothing Then Return paragraphs

        For Each para As Paragraph In _document.Paragraphs
            Dim text = para.Range.Text
            ' 移除段落标记
            text = text.TrimEnd(ChrW(13), ChrW(10), ChrW(7))
            If Not String.IsNullOrWhiteSpace(text) Then
                paragraphs.Add(text)
            Else
                ' 保留空段落位置
                paragraphs.Add("")
            End If
        Next

        Return paragraphs
    End Function

    ''' <summary>
    ''' 获取选中的段落
    ''' </summary>
    Public Overrides Function GetSelectedParagraphs() As List(Of String)
        Dim paragraphs As New List(Of String)()
        _selectionRanges = New List(Of Range)()

        If _wordApp.Selection Is Nothing Then Return paragraphs

        Dim selRange = _wordApp.Selection.Range

        If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(selRange.Text) Then
            Return paragraphs
        End If

        For Each para As Paragraph In selRange.Paragraphs
            Dim text = para.Range.Text
            text = text.TrimEnd(ChrW(13), ChrW(10), ChrW(7))
            If Not String.IsNullOrWhiteSpace(text) Then
                paragraphs.Add(text)
                _selectionRanges.Add(para.Range)
            Else
                paragraphs.Add("")
                _selectionRanges.Add(para.Range)
            End If
        Next

        Return paragraphs
    End Function

    ''' <summary>
    ''' 应用翻译结果到整个文档
    ''' </summary>
    Public Overrides Sub ApplyTranslation(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        If results Is Nothing OrElse results.Count = 0 Then Return

        Select Case outputMode
            Case TranslateOutputMode.Replace
                ApplyReplaceMode(results, False)
            Case TranslateOutputMode.Immersive
                ApplyImmersiveMode(results, False)
            Case TranslateOutputMode.NewDocument
                ApplyToNewDocument(results)
            Case TranslateOutputMode.SidePanel
                ' 侧栏模式由调用者处理
        End Select
    End Sub

    ''' <summary>
    ''' 应用翻译结果到选中区域
    ''' </summary>
    Public Overrides Sub ApplyTranslationToSelection(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        If results Is Nothing OrElse results.Count = 0 Then Return

        Select Case outputMode
            Case TranslateOutputMode.Replace
                ApplyReplaceMode(results, True)
            Case TranslateOutputMode.Immersive
                ApplyImmersiveMode(results, True)
            Case TranslateOutputMode.NewDocument
                ApplyToNewDocument(results)
            Case TranslateOutputMode.SidePanel
                ' 侧栏模式由调用者处理
        End Select
    End Sub

    ''' <summary>
    ''' 替换模式 - 直接替换原文
    ''' </summary>
    Private Sub ApplyReplaceMode(results As List(Of TranslateParagraphResult), isSelection As Boolean)
        Try
            _document.Application.ScreenUpdating = False
            _document.Application.UndoRecord.StartCustomRecord("AI翻译")

            If isSelection AndAlso _selectionRanges IsNot Nothing Then
                ' 替换选中区域
                For i = results.Count - 1 To 0 Step -1
                    Dim result = results(i)
                    If result.Success AndAlso i < _selectionRanges.Count Then
                        Dim range = _selectionRanges(i)
                        If Not String.IsNullOrEmpty(result.TranslatedText) Then
                            range.Text = result.TranslatedText & vbCr
                        End If
                    End If
                Next
            Else
                ' 替换整个文档
                Dim paras = _document.Paragraphs
                For i = Math.Min(results.Count, paras.Count) - 1 To 0 Step -1
                    Dim result = results(i)
                    If result.Success AndAlso Not String.IsNullOrEmpty(result.TranslatedText) Then
                        Dim para = paras(i + 1)
                        para.Range.Text = result.TranslatedText & vbCr
                    End If
                Next
            End If

            _document.Application.UndoRecord.EndCustomRecord()
        Catch ex As Exception
            MessageBox.Show("应用翻译结果时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _document.Application.ScreenUpdating = True
        End Try
    End Sub

    ''' <summary>
    ''' 沉浸式翻译模式 - 原文+译文并行显示
    ''' </summary>
    Private Sub ApplyImmersiveMode(results As List(Of TranslateParagraphResult), isSelection As Boolean)
        Try
            _document.Application.ScreenUpdating = False
            _document.Application.UndoRecord.StartCustomRecord("AI沉浸式翻译")

            Dim settings = TranslateSettings.Load()

            If isSelection AndAlso _selectionRanges IsNot Nothing Then
                ' 在选中区域后插入译文
                For i = results.Count - 1 To 0 Step -1
                    Dim result = results(i)
                    If result.Success AndAlso i < _selectionRanges.Count AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                        Dim originalRange = _selectionRanges(i)
                        InsertImmersiveTranslation(originalRange, result.TranslatedText, settings)
                    End If
                Next
            Else
                ' 在每个段落后插入译文
                Dim paras = _document.Paragraphs
                For i = Math.Min(results.Count, paras.Count) - 1 To 0 Step -1
                    Dim result = results(i)
                    If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                        Dim para = paras(i + 1)
                        InsertImmersiveTranslation(para.Range, result.TranslatedText, settings)
                    End If
                Next
            End If

            _document.Application.UndoRecord.EndCustomRecord()
        Catch ex As Exception
            MessageBox.Show("应用沉浸式翻译时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _document.Application.ScreenUpdating = True
        End Try
    End Sub

    ''' <summary>
    ''' 插入沉浸式翻译（在原文段落后插入译文）
    ''' </summary>
    Private Sub InsertImmersiveTranslation(originalRange As Range, translatedText As String, settings As TranslateSettings)
        Try
            ' 获取原段落的末尾位置（不包含段落标记）
            Dim paraEnd = originalRange.End - 1
            If paraEnd < originalRange.Start Then paraEnd = originalRange.Start

            ' 在段落末尾插入换行和译文
            Dim insertPoint = _document.Range(paraEnd, paraEnd)
            Dim textToInsert = vbCr & translatedText

            insertPoint.InsertAfter(textToInsert)

            ' 选择新插入的译文范围（跳过换行符）
            Dim translatedStart = paraEnd + 1
            Dim translatedEnd = translatedStart + translatedText.Length
            
            ' 确保范围在文档边界内
            If translatedEnd > _document.Content.End Then
                translatedEnd = _document.Content.End
            End If
            
            If translatedStart >= translatedEnd Then Return

            Dim translatedRange = _document.Range(translatedStart, translatedEnd)

            ' 只有在不保持原文格式时才设置样式
            If Not settings.PreserveFormatting Then
                With translatedRange.Font
                    ' 设置颜色
                    Try
                        Dim colorHex = settings.ImmersiveTranslationColor.TrimStart("#"c)
                        If colorHex.Length >= 6 Then
                            Dim r = Convert.ToInt32(colorHex.Substring(0, 2), 16)
                            Dim g = Convert.ToInt32(colorHex.Substring(2, 2), 16)
                            Dim b = Convert.ToInt32(colorHex.Substring(4, 2), 16)
                            .Color = CType(RGB(r, g, b), WdColor)
                        End If
                    Catch
                        ' 颜色设置失败时忽略
                    End Try

                    ' 设置斜体
                    If settings.ImmersiveTranslationItalic Then
                        .Italic = -1
                    End If

                    ' 设置字号比例
                    If settings.ImmersiveTranslationFontScale <> 1.0 AndAlso settings.ImmersiveTranslationFontScale > 0 Then
                        Dim originalSize = originalRange.Font.Size
                        If originalSize > 0 Then
                            .Size = CSng(originalSize * settings.ImmersiveTranslationFontScale)
                        End If
                    End If
                End With

                ' 设置段落缩进
                Try
                    translatedRange.ParagraphFormat.LeftIndent = originalRange.ParagraphFormat.LeftIndent + _wordApp.InchesToPoints(0.25)
                Catch
                End Try
            End If
        Catch ex As Exception
            ' 单个段落插入失败时继续处理其他段落
            Debug.WriteLine("InsertImmersiveTranslation error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 创建新文档并写入翻译结果（保留原文档格式）
    ''' </summary>
    Private Sub ApplyToNewDocument(results As List(Of TranslateParagraphResult))
        Try
            ' 复制原文档到新文档（保留所有格式）
            _document.Content.Copy()
            Dim newDoc = _wordApp.Documents.Add()
            newDoc.Content.Paste()

            ' 替换每个段落的文本内容（保留格式）
            Dim paras = newDoc.Paragraphs
            For i = 0 To Math.Min(results.Count, paras.Count) - 1
                Dim result = results(i)
                If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                    Dim para = paras(i + 1)
                    ' 仅替换文本，不改变格式
                    Dim paraRange = para.Range
                    Dim originalEnd = paraRange.End
                    
                    ' 保存原始格式
                    Dim fontName = paraRange.Font.Name
                    Dim fontSize = paraRange.Font.Size
                    Dim fontBold = paraRange.Font.Bold
                    Dim fontItalic = paraRange.Font.Italic
                    Dim fontColor = paraRange.Font.Color
                    Dim paraAlignment = para.Alignment
                    Dim firstLineIndent = para.FirstLineIndent
                    Dim leftIndent = para.LeftIndent
                    Dim rightIndent = para.RightIndent
                    Dim lineSpacing = para.LineSpacing
                    
                    ' 替换文本
                    paraRange.Text = result.TranslatedText & vbCr
                    
                    ' 恢复格式
                    paraRange.Font.Name = fontName
                    paraRange.Font.Size = fontSize
                    paraRange.Font.Bold = fontBold
                    paraRange.Font.Italic = fontItalic
                    paraRange.Font.Color = fontColor
                    para.Alignment = paraAlignment
                    para.FirstLineIndent = firstLineIndent
                    para.LeftIndent = leftIndent
                    para.RightIndent = rightIndent
                    para.LineSpacing = lineSpacing
                End If
            Next

            newDoc.Activate()
        Catch ex As Exception
            MessageBox.Show("创建新文档时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 生成翻译结果的格式化文本（用于侧栏显示）
    ''' </summary>
    Public Function FormatResultsForDisplay(results As List(Of TranslateParagraphResult), showOriginal As Boolean) As String
        Dim sb As New StringBuilder()

        For Each result In results
            If showOriginal Then
                sb.AppendLine("【原文】")
                sb.AppendLine(result.OriginalText)
                sb.AppendLine()
                sb.AppendLine("【译文】")
            End If

            If result.Success Then
                sb.AppendLine(result.TranslatedText)
            Else
                sb.AppendLine($"[翻译失败: {result.ErrorMessage}]")
                sb.AppendLine(result.OriginalText)
            End If

            If showOriginal Then
                sb.AppendLine(New String("-"c, 40))
            End If
            sb.AppendLine()
        Next

        Return sb.ToString()
    End Function
End Class
