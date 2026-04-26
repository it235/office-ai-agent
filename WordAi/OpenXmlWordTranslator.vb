Imports System.Diagnostics
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports ShareRibbon
Imports Word = Microsoft.Office.Interop.Word

''' <summary>
''' Word 翻译器 - 使用 COM 定位段落，改进的写入方式避免 InsertXML 限制
''' 核心优势：
'''   1. 保留所有文档元素（表格、图片、公式）
'''   2. 段落索引稳定（倒序处理）
'''   3. 使用 Range.Text 替换，避免 InsertXML 的 XML 结构限制
''' </summary>
Public Class OpenXmlWordTranslator

    ''' <summary>
    ''' 翻译文本块
    ''' </summary>
    Public Class TranslationBlock
        Public Property BlockIndex As Integer
        Public Property ParagraphIndex As Integer
        Public Property OriginalText As String
        Public Property IsTableCell As Boolean
        Public Property TableRow As Integer
        Public Property TableCol As Integer
        Public Property OriginalXml As String
    End Class

    ''' <summary>
    ''' 扫描整个文档，返回所有可翻译文本块
    ''' </summary>
    Public Function ScanDocument(document As Word.Document) As List(Of TranslationBlock)
        Return ScanInternal(document, Nothing)
    End Function

    ''' <summary>
    ''' 扫描选中区域，返回选中范围内的可翻译文本块
    ''' </summary>
    Public Function ScanSelection(document As Word.Document, selection As Word.Selection) As List(Of TranslationBlock)
        If selection Is Nothing OrElse selection.Range Is Nothing Then
            Return New List(Of TranslationBlock)()
        End If
        Return ScanInternal(document, selection.Range)
    End Function

    ''' <summary>
    ''' 内部扫描逻辑
    ''' </summary>
    Private Function ScanInternal(document As Word.Document, selRange As Word.Range) As List(Of TranslationBlock)
        Dim blocks As New List(Of TranslationBlock)()
        Dim index As Integer = 0

        If document Is Nothing Then Return blocks

        ' 收集表格位置，避免重复处理
        Dim processedTables As New HashSet(Of Integer)()
        For Each tbl As Word.Table In document.Tables
            Try
                processedTables.Add(tbl.Range.Start)
            Catch
            End Try
        Next

        Dim paraIndex As Integer = 1
        For Each para As Word.Paragraph In document.Paragraphs
            Try
                ' 如果指定了选择范围，检查段落是否在范围内
                If selRange IsNot Nothing Then
                    Dim paraStart = para.Range.Start
                    Dim paraEnd = para.Range.End
                    If Not (paraStart < selRange.End AndAlso paraEnd > selRange.Start) Then
                        paraIndex += 1
                        Continue For
                    End If
                End If

                Dim isInTable As Boolean = False
                Dim tableRef As Word.Table = Nothing

                Try
                    isInTable = CBool(para.Range.Information(Word.WdInformation.wdWithInTable))
                    If isInTable Then
                        tableRef = para.Range.Tables(1)
                    End If
                Catch
                End Try

                If isInTable AndAlso tableRef IsNot Nothing Then
                    Dim tableStart = tableRef.Range.Start
                    If processedTables.Contains(tableStart) Then
                        Dim rowCount = tableRef.Rows.Count
                        Dim colCount = tableRef.Columns.Count
                        For rowIdx = 1 To rowCount
                            For colIdx = 1 To colCount
                                Try
                                    Dim cell = tableRef.Cell(rowIdx, colIdx)
                                    Dim cellText = cell.Range.Text
                                    If cellText IsNot Nothing Then
                                        cellText = cellText.TrimEnd(ChrW(7), ChrW(13), ChrW(10))
                                    End If

                                    Dim cellPara = cell.Range.Paragraphs(1)

                                    blocks.Add(New TranslationBlock With {
                                        .BlockIndex = index,
                                        .ParagraphIndex = paraIndex,
                                        .OriginalText = If(String.IsNullOrWhiteSpace(cellText), "", cellText),
                                        .IsTableCell = True,
                                        .TableRow = rowIdx,
                                        .TableCol = colIdx,
                                        .OriginalXml = cellPara.Range.XML
                                    })
                                    index += 1
                                Catch cellEx As Exception
                                    Debug.WriteLine($"Scan cell ({rowIdx},{colIdx}) error: {cellEx.Message}")
                                End Try
                            Next
                        Next
                    End If
                Else
                    ' 普通段落
                    Dim paraText As String = ""
                    Try
                        paraText = para.Range.Text
                        If paraText IsNot Nothing Then
                            paraText = paraText.TrimEnd(ChrW(13), ChrW(10))
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"Get paragraph text error: {ex.Message}")
                    End Try

                    If Not String.IsNullOrWhiteSpace(paraText) Then
                        blocks.Add(New TranslationBlock With {
                            .BlockIndex = index,
                            .ParagraphIndex = paraIndex,
                            .OriginalText = paraText,
                            .IsTableCell = False,
                            .OriginalXml = para.Range.XML
                        })
                        index += 1
                    End If
                End If
            Catch paraEx As Exception
                Debug.WriteLine($"Scan paragraph {paraIndex} error: {paraEx.Message}")
            End Try

            paraIndex += 1
        Next

        Return blocks
    End Function

    ''' <summary>
    ''' 应用翻译结果到文档
    ''' </summary>
    Public Sub ApplyTranslation(document As Word.Document,
                                blocks As List(Of TranslationBlock),
                                results As List(Of TranslateParagraphResult),
                                outputMode As TranslateOutputMode,
                                settings As TranslateSettings)
        If blocks Is Nothing OrElse results Is Nothing OrElse blocks.Count = 0 Then
            Return
        End If

        Dim resultDict = results.ToDictionary(Function(r) r.Index)
        Dim sortedBlocks = blocks.OrderByDescending(Function(b) b.BlockIndex).ToList()

        Try
            document.Application.ScreenUpdating = False
            document.Application.UndoRecord.StartCustomRecord("AI翻译")

            Select Case outputMode
                Case TranslateOutputMode.Replace
                    ApplyReplaceMode(document, sortedBlocks, resultDict, settings)
                Case TranslateOutputMode.Immersive
                    ApplyImmersiveMode(document, sortedBlocks, resultDict, settings)
                Case Else
                    Debug.WriteLine($"[OpenXmlTranslator] Unsupported output mode: {outputMode}")
            End Select

            document.Application.UndoRecord.EndCustomRecord()
        Catch ex As Exception
            Debug.WriteLine($"[OpenXmlTranslator] ApplyTranslation error: {ex.Message}")
            MessageBox.Show("应用翻译结果时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            document.Application.ScreenUpdating = True
        End Try
    End Sub

#Region "替换模式"

    ''' <summary>
    ''' 替换模式：使用 Range.Text 直接替换文本，保留图片/公式等内嵌对象
    ''' </summary>
    Private Sub ApplyReplaceMode(document As Word.Document,
                                 sortedBlocks As List(Of TranslationBlock),
                                 resultDict As Dictionary(Of Integer, TranslateParagraphResult),
                                 settings As TranslateSettings)
        For Each block In sortedBlocks
            If Not resultDict.ContainsKey(block.BlockIndex) Then Continue For

            Dim result = resultDict(block.BlockIndex)
            If Not result.Success OrElse String.IsNullOrWhiteSpace(result.TranslatedText) Then Continue For

            Try
                Dim para = document.Paragraphs(block.ParagraphIndex)
                If para Is Nothing Then Continue For

                Dim paraRange = para.Range
                If paraRange Is Nothing Then Continue For

                ' 清理翻译文本中的换行符
                Dim cleanText = result.TranslatedText.Replace(vbCr, " ").Replace(vbLf, " ")

                ' 检查是否有内嵌对象
                Dim hasObjects As Boolean = False
                Try
                    hasObjects = paraRange.InlineShapes.Count > 0 OrElse paraRange.OMaths.Count > 0
                Catch
                    hasObjects = False
                End Try

                If hasObjects Then
                    ' 有内嵌对象时，使用替换逻辑保留对象
                    ReplaceTextPreservingObjects(paraRange, cleanText, document)
                Else
                    ' 没有内嵌对象，直接替换
                    paraRange.Text = cleanText
                End If

            Catch ex As Exception
                Debug.WriteLine($"[OpenXmlTranslator] Replace block {block.BlockIndex} error: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' 替换文本同时保留内嵌对象（图片、公式等）
    ''' </summary>
    Private Sub ReplaceTextPreservingObjects(range As Word.Range, newText As String, document As Word.Document)
        Try
            If range Is Nothing Then Return

            Dim rangeStart = range.Start
            Dim rangeEnd = range.End

            ' 检查是否有对象
            Dim hasObjects As Boolean = False
            Try
                hasObjects = range.InlineShapes.Count > 0 OrElse range.OMaths.Count > 0
            Catch
                hasObjects = False
            End Try

            If Not hasObjects Then
                range.Text = newText
                Return
            End If

            ' 收集所有对象的位置
            Dim objRanges As New List(Of Tuple(Of Integer, Integer))()

            Try
                For Each shape As Word.InlineShape In range.InlineShapes
                    If shape.Range.Start >= rangeStart AndAlso shape.Range.End <= rangeEnd Then
                        objRanges.Add(Tuple.Create(shape.Range.Start, shape.Range.End))
                    End If
                Next
            Catch
            End Try

            Try
                For Each omath As Word.OMath In range.OMaths
                    If omath.Range.Start >= rangeStart AndAlso omath.Range.End <= rangeEnd Then
                        objRanges.Add(Tuple.Create(omath.Range.Start, omath.Range.End))
                    End If
                Next
            Catch
            End Try

            ' 去重并排序
            objRanges = objRanges.Where(Function(r) r.Item2 > r.Item1).Distinct().OrderBy(Function(r) r.Item1).ToList()

            If objRanges.Count = 0 Then
                range.Text = newText
                Return
            End If

            ' 确定内容结束位置（排除段落结束标记）
            Dim contentEnd = rangeEnd
            If rangeEnd > rangeStart Then
                contentEnd = rangeEnd - 1
            End If

            ' 从后往前删除对象之间的文本
            Dim textEnd = contentEnd
            For i = objRanges.Count - 1 To 0 Step -1
                Dim objStart = objRanges(i).Item1
                Dim objEnd = objRanges(i).Item2

                If textEnd > objEnd Then
                    Try
                        Dim delRange = document.Range(objEnd, textEnd)
                        delRange.Text = ""
                    Catch delEx As Exception
                        Debug.WriteLine($"[OpenXmlTranslator] Delete text after object failed: {delEx.Message}")
                    End Try
                End If
                textEnd = objStart
            Next

            ' 删除第一个对象之前的文本
            If textEnd > rangeStart Then
                Try
                    Dim delRange = document.Range(rangeStart, textEnd)
                    delRange.Text = ""
                Catch delEx As Exception
                    Debug.WriteLine($"[OpenXmlTranslator] Delete text before object failed: {delEx.Message}")
                End Try
            End If

            ' 在开头插入译文
            Try
                Dim insertRange = document.Range(rangeStart, rangeStart)
                insertRange.InsertBefore(newText)
            Catch insertEx As Exception
                Debug.WriteLine($"[OpenXmlTranslator] Insert translation failed: {insertEx.Message}")
            End Try

        Catch ex As Exception
            Debug.WriteLine($"[OpenXmlTranslator] ReplaceTextPreservingObjects error: {ex.Message}")
            ' 最后的后备方案
            Try
                range.Text = newText
            Catch
            End Try
        End Try
    End Sub

#End Region

#Region "沉浸式模式"

    ''' <summary>
    ''' 沉浸式模式：在原文段落后插入译文段落
    ''' </summary>
    Private Sub ApplyImmersiveMode(document As Word.Document,
                                   sortedBlocks As List(Of TranslationBlock),
                                   resultDict As Dictionary(Of Integer, TranslateParagraphResult),
                                   settings As TranslateSettings)
        For Each block In sortedBlocks
            If Not resultDict.ContainsKey(block.BlockIndex) Then Continue For

            Dim result = resultDict(block.BlockIndex)
            If Not result.Success OrElse String.IsNullOrWhiteSpace(result.TranslatedText) Then Continue For

            Try
                Dim para = document.Paragraphs(block.ParagraphIndex)
                If para Is Nothing Then Continue For

                ' 在原文段落后插入新段落
                Dim insertRange = para.Range
                insertRange.InsertParagraphAfter()

                ' 获取新插入的段落
                Dim newPara As Word.Paragraph = Nothing
                Try
                    newPara = para.Next
                Catch
                End Try

                If newPara Is Nothing Then
                    Try
                        newPara = document.Paragraphs(block.ParagraphIndex + 1)
                    Catch
                    End Try
                End If

                If newPara IsNot Nothing Then
                    ' 设置译文文本
                    Dim cleanText = result.TranslatedText.Replace(vbCr, " ").Replace(vbLf, " ")
                    Dim newParaRange = newPara.Range
                    Dim contentEnd = newParaRange.End
                    If contentEnd > newParaRange.Start Then
                        contentEnd -= 1
                    End If
                    Dim textRange = document.Range(newParaRange.Start, contentEnd)
                    textRange.Text = cleanText

                    ' 应用沉浸式样式
                    ApplyImmersiveStyleToRange(newParaRange, settings)
                End If

            Catch ex As Exception
                Debug.WriteLine($"[OpenXmlTranslator] Immersive block {block.BlockIndex} error: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' 应用沉浸式样式到段落
    ''' </summary>
    Private Sub ApplyImmersiveStyleToRange(targetRange As Word.Range, settings As TranslateSettings)
        Try
            If settings.PreserveFormatting Then Return

            ' 颜色
            Try
                Dim colorHex = settings.ImmersiveTranslationColor.TrimStart("#"c)
                If colorHex.Length >= 6 Then
                    Dim r = Convert.ToInt32(colorHex.Substring(0, 2), 16)
                    Dim g = Convert.ToInt32(colorHex.Substring(2, 2), 16)
                    Dim b = Convert.ToInt32(colorHex.Substring(4, 2), 16)
                    targetRange.Font.Color = CType(Word.WdColor.wdColorAutomatic + RGB(r, g, b), Word.WdColor)
                End If
            Catch
            End Try

            ' 斜体
            If settings.ImmersiveTranslationItalic Then
                targetRange.Font.Italic = -1
            End If

            ' 字号
            If settings.ImmersiveTranslationFontScale <> 1.0 AndAlso settings.ImmersiveTranslationFontScale > 0 Then
                Dim origSize = targetRange.Font.Size
                If origSize > 0 Then
                    targetRange.Font.Size = CSng(origSize * settings.ImmersiveTranslationFontScale)
                End If
            End If

            ' 缩进（0.25英寸 = 18磅）
            targetRange.ParagraphFormat.LeftIndent = targetRange.ParagraphFormat.LeftIndent + 18

        Catch ex As Exception
            Debug.WriteLine($"[OpenXmlTranslator] ApplyImmersiveStyle error: {ex.Message}")
        End Try
    End Sub

#End Region

End Class
