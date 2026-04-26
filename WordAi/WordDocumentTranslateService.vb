Imports System.Diagnostics
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports ShareRibbon
Imports Word = Microsoft.Office.Interop.Word

''' <summary>
''' 翻译单元类型
''' </summary>
Public Enum TranslationUnitType
    NormalText      ' 普通文本
    TableContent    ' 表格内容
End Enum

''' <summary>
''' 翻译单元：可以是普通文本或表格
''' </summary>
Public Class TranslationUnit
    Public Property UnitType As TranslationUnitType
    Public Property Text As String                  ' 文本内容（表格以tab和\n格式化）
    Public Property ParagraphIndex As Integer       ' 文档中的段落索引（1-based）
    Public Property StartPos As Integer             ' 收集时的文档开始位置
    Public Property TableRef As Table               ' 如果是表格，保存表格引用
    Public Property RowCount As Integer             ' 表格行数
    Public Property ColumnCount As Integer          ' 表格列数
End Class

''' <summary>
''' Word文档翻译服务 - 支持一键翻译和沉浸式翻译
''' </summary>
Public Class WordDocumentTranslateService
    Inherits DocumentTranslateService

    Private _wordApp As Word.Application
    Private _document As Document
    Private _selectionParagraphIndices As List(Of Integer)  ' 选中区域段落的文档索引（1-based）
    Private _translationUnits As List(Of TranslationUnit)  ' 翻译单元列表

    ' OpenXML 级翻译器（新版）
    Private _openXmlTranslator As OpenXmlWordTranslator
    Private _translationBlocks As List(Of OpenXmlWordTranslator.TranslationBlock)

    Public Sub New(wordApp As Word.Application)
        MyBase.New()
        _wordApp = wordApp
        _document = wordApp.ActiveDocument
    End Sub

    ''' <summary>
    ''' 获取文档所有段落（使用 OpenXML 级扫描）
    ''' </summary>
    Public Overrides Function GetAllParagraphs() As List(Of String)
        _openXmlTranslator = New OpenXmlWordTranslator()
        _translationBlocks = _openXmlTranslator.ScanDocument(_document)

        ' 同时保留旧的 TranslationUnit 结构（兼容性）
        _translationUnits = New List(Of TranslationUnit)()
        For Each block In _translationBlocks
            _translationUnits.Add(New TranslationUnit With {
                .UnitType = If(block.IsTableCell, TranslationUnitType.TableContent, TranslationUnitType.NormalText),
                .Text = block.OriginalText,
                .ParagraphIndex = block.ParagraphIndex,
                .StartPos = block.BlockIndex
            })
        Next

        Return _translationBlocks.Select(Function(b) b.OriginalText).ToList()
    End Function

    ''' <summary>
    ''' 提取表格内容为格式化文本（行\n列\t）
    ''' </summary>
    Private Function ExtractTableContent(table As Table, paraIndex As Integer) As TranslationUnit
        Dim sb As New StringBuilder()
        Dim rowCount = table.Rows.Count
        Dim colCount = table.Columns.Count

        Try
            For rowIdx = 1 To rowCount
                For colIdx = 1 To colCount
                    Try
                        Dim cell = table.Cell(rowIdx, colIdx)
                        Dim cellText = cell.Range.Text

                        If cellText Is Nothing Then cellText = ""

                        ' 移除单元格结束符
                        cellText = cellText.TrimEnd(ChrW(7), ChrW(13), ChrW(10))

                        sb.Append(cellText)

                        If colIdx < colCount Then
                            sb.Append(vbTab)
                        End If
                    Catch cellEx As Exception
                        Debug.WriteLine($"Extract cell ({rowIdx},{colIdx}) error: {cellEx.Message}")
                        If colIdx < colCount Then sb.Append(vbTab)
                    End Try
                Next

                If rowIdx < rowCount Then
                    sb.AppendLine()
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"ExtractTableContent error: {ex.Message}")
        End Try

        Return New TranslationUnit With {
            .UnitType = TranslationUnitType.TableContent,
            .Text = sb.ToString(),
            .TableRef = table,
            .RowCount = rowCount,
            .ColumnCount = colCount,
            .ParagraphIndex = paraIndex,
            .StartPos = table.Range.Start
        }
    End Function

    ''' <summary>
    ''' 获取选中的段落（使用 OpenXML 级扫描）
    ''' </summary>
    Public Overrides Function GetSelectedParagraphs() As List(Of String)
        _openXmlTranslator = New OpenXmlWordTranslator()
        _translationBlocks = _openXmlTranslator.ScanSelection(_document, _wordApp.Selection)

        ' 构建选中段落索引列表（兼容性）
        _selectionParagraphIndices = _translationBlocks.Select(Function(b) b.ParagraphIndex).Distinct().ToList()

        Return _translationBlocks.Select(Function(b) b.OriginalText).ToList()
    End Function

    ''' <summary>
    ''' 应用翻译结果到整个文档（使用 OpenXML 级写入）
    ''' </summary>
    Public Overrides Sub ApplyTranslation(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        If results Is Nothing OrElse results.Count = 0 Then Return

        If outputMode = TranslateOutputMode.NewDocument Then
            ApplyToNewDocument(results)
            Return
        End If

        If outputMode = TranslateOutputMode.SidePanel Then
            ' 侧栏模式由调用方处理
            Return
        End If

        If _openXmlTranslator IsNot Nothing AndAlso _translationBlocks IsNot Nothing Then
            Dim settings = TranslateSettings.Load()
            _openXmlTranslator.ApplyTranslation(_document, _translationBlocks, results, outputMode, settings)
        End If
    End Sub

    ''' <summary>
    ''' 应用翻译结果到选中区域（使用 OpenXML 级写入）
    ''' </summary>
    Public Overrides Sub ApplyTranslationToSelection(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        If results Is Nothing OrElse results.Count = 0 Then Return

        If outputMode = TranslateOutputMode.NewDocument Then
            ApplyToNewDocument(results)
            Return
        End If

        If outputMode = TranslateOutputMode.SidePanel Then
            ' 侧栏模式由调用方处理
            Return
        End If

        If _openXmlTranslator IsNot Nothing AndAlso _translationBlocks IsNot Nothing Then
            Dim settings = TranslateSettings.Load()
            _openXmlTranslator.ApplyTranslation(_document, _translationBlocks, results, outputMode, settings)
        End If
    End Sub

    ''' <summary>
    ''' 替换模式 - 直接替换原文
    ''' </summary>
    Private Sub ApplyReplaceMode(results As List(Of TranslateParagraphResult), isSelection As Boolean)
        Try
            _document.Application.ScreenUpdating = False
            _document.Application.UndoRecord.StartCustomRecord("AI翻译")

            If isSelection AndAlso _selectionParagraphIndices IsNot Nothing Then
                ' 替换选中区域 - 通过段落索引重新获取Range
                For i = Math.Min(results.Count, _selectionParagraphIndices.Count) - 1 To 0 Step -1
                    Dim result = results(i)
                    If result.Success AndAlso Not String.IsNullOrEmpty(result.TranslatedText) Then
                        Dim paraIndex = _selectionParagraphIndices(i)
                        If paraIndex > 0 AndAlso paraIndex <= _document.Paragraphs.Count Then
                            Try
                                Dim para = _document.Paragraphs(paraIndex)
                                Dim cleanText = SanitizeParagraphText(result.TranslatedText)
                                ReplaceRangeTextPreservingObjects(para.Range, cleanText)
                            Catch ex As Exception
                                Debug.WriteLine($"[WordTranslate] Apply replace to selected paragraph {paraIndex} failed: {ex.Message}")
                            End Try
                        End If
                    End If
                Next
            Else
                ' 替换整个文档 - 使用翻译单元，通过索引重新获取Range
                If _translationUnits IsNot Nothing AndAlso _translationUnits.Count > 0 Then
                    For i = Math.Min(results.Count, _translationUnits.Count) - 1 To 0 Step -1
                        Dim result = results(i)
                        If result.Success AndAlso Not String.IsNullOrEmpty(result.TranslatedText) Then
                            Dim unit = _translationUnits(i)

                            If unit.UnitType = TranslationUnitType.TableContent Then
                                ' 表格翻译 - TableRef COM引用仍然有效
                                ApplyTableTranslation(unit.TableRef, result.TranslatedText, False)
                            Else
                                ' 普通文本翻译 - 通过段落索引重新获取Range
                                If unit.ParagraphIndex > 0 AndAlso unit.ParagraphIndex <= _document.Paragraphs.Count Then
                                    Try
                                        Dim para = _document.Paragraphs(unit.ParagraphIndex)
                                        Dim cleanText = SanitizeParagraphText(result.TranslatedText)
                                        ReplaceRangeTextPreservingObjects(para.Range, cleanText)
                                    Catch ex As Exception
                                        Debug.WriteLine($"[WordTranslate] Apply replace to paragraph {unit.ParagraphIndex} failed: {ex.Message}")
                                    End Try
                                End If
                            End If
                        End If
                    Next
                Else
                    ' 后备方案：使用原有逻辑
                    Dim paras = _document.Paragraphs
                    For i = Math.Min(results.Count, paras.Count) - 1 To 0 Step -1
                        Dim result = results(i)
                        If result.Success AndAlso Not String.IsNullOrEmpty(result.TranslatedText) Then
                            Dim para = paras(i + 1)
                            Dim cleanText = SanitizeParagraphText(result.TranslatedText)
                            ReplaceRangeTextPreservingObjects(para.Range, cleanText)
                        End If
                    Next
                End If
            End If

            _document.Application.UndoRecord.EndCustomRecord()
        Catch ex As Exception
            MessageBox.Show("应用翻译结果时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _document.Application.ScreenUpdating = True
        End Try
    End Sub

    ''' <summary>
    ''' 沉浸式表格翻译 - 复制表格并在新表格中填入翻译内容
    ''' </summary>
    Private Sub ApplyImmersiveTableTranslation(originalTable As Table, translatedText As String)
        Try
            Debug.WriteLine($"ApplyImmersiveTableTranslation - Original table rows: {originalTable.Rows.Count}, cols: {originalTable.Columns.Count}")

            ' 1. 复制原始表格到其后
            Dim tableEnd = originalTable.Range.End
            originalTable.Range.Copy()
            Dim insertPoint = _document.Range(tableEnd, tableEnd)
            insertPoint.Paste()

            ' 2. 获取新粘贴的表格
            Dim newTable As Table = Nothing
            Try
                ' 搜索新表格
                Dim searchRange = _document.Range(tableEnd, tableEnd + 300)
                Debug.WriteLine($"Searching for new table in range {tableEnd} to {tableEnd + 300}")
                Debug.WriteLine($"Found {searchRange.Tables.Count} tables in search range")

                For i = 1 To searchRange.Tables.Count
                    Dim tbl = searchRange.Tables(i)
                    Dim distance = Math.Abs(tbl.Range.Start - tableEnd)
                    Debug.WriteLine($"Table {i}: start={tbl.Range.Start}, distance={distance}")

                    If distance < 150 Then
                        newTable = tbl
                        Debug.WriteLine($"Found new table at position {tbl.Range.Start}")
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine($"Find new table error: {ex.Message}")
            End Try

            If newTable Is Nothing Then
                Debug.WriteLine("Failed to find copied table")
                Return
            End If

            Debug.WriteLine($"New table rows: {newTable.Rows.Count}, cols: {newTable.Columns.Count}")

            ' 3. 解析翻译文本并填入新表格
            ' 使用None选项保留空行和空单元格
            Dim lines = translatedText.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)

            ' 确保行列数匹配
            Dim maxRows = Math.Min(lines.Length, newTable.Rows.Count)

            Debug.WriteLine($"Processing table: {maxRows} rows, {newTable.Columns.Count} columns")
            Debug.WriteLine($"Translated text lines: {lines.Length}")

            For rowIdx = 1 To maxRows
                ' 获取行文本
                Dim line As String = ""
                If rowIdx - 1 < lines.Length Then
                    line = lines(rowIdx - 1)
                End If

                line = If(line, "").Trim()

                ' 按tab分割单元格
                Dim cells() As String
                If String.IsNullOrEmpty(line) Then
                    ' 空行，创建空单元格数组
                    ReDim cells(newTable.Columns.Count - 1)
                    For i = 0 To newTable.Columns.Count - 1
                        cells(i) = ""
                    Next
                Else
                    cells = line.Split(vbTab)
                End If

                Debug.WriteLine($"Row {rowIdx}: {cells.Length} cells, line='{line}'")

                ' 处理每一列
                For colIdx = 1 To newTable.Columns.Count
                    Try
                        Dim cell = newTable.Cell(rowIdx, colIdx)
                        Dim cellText As String = ""

                        ' 获取对应单元格的翻译文本
                        If colIdx - 1 < cells.Length Then
                            cellText = cells(colIdx - 1)
                        End If

                        cellText = If(cellText, "").Trim()

                        Debug.WriteLine($"  Cell({rowIdx},{colIdx}): '{cellText}'")

                        ' 替换单元格内容，保留对象
                        ReplaceCellTextPreservingObjects(cell.Range, cellText)
                    Catch cellEx As Exception
                        Debug.WriteLine($"Apply cell ({rowIdx},{colIdx}) translation error: {cellEx.Message}")
                    End Try
                Next
            Next

        Catch ex As Exception
            Debug.WriteLine($"ApplyImmersiveTableTranslation error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用表格翻译结果（覆盖模式）
    ''' </summary>
    Private Sub ApplyTableTranslation(originalTable As Table, translatedText As String, isCopy As Boolean)
        Try
            Debug.WriteLine($"ApplyTableTranslation - isCopy: {isCopy}, Original table rows: {originalTable.Rows.Count}, cols: {originalTable.Columns.Count}")

            Dim targetTable As Table = originalTable

            If isCopy Then
                Debug.WriteLine("Replace mode with copy: Copying table...")
                ' 覆盖模式下的复制（用于测试）
                Dim tableEnd = originalTable.Range.End
                originalTable.Range.Copy()
                Dim insertPoint = _document.Range(tableEnd, tableEnd)
                insertPoint.Paste()

                ' 获取新粘贴的表格
                Dim newTable As Table = Nothing
                Try
                    ' 尝试获取粘贴位置后的第一个表格
                    Dim searchRange = _document.Range(tableEnd, tableEnd + 200)
                    Debug.WriteLine($"Searching for new table in range {tableEnd} to {tableEnd + 200}")
                    Debug.WriteLine($"Found {searchRange.Tables.Count} tables in search range")

                    For i = 1 To searchRange.Tables.Count
                        Dim tbl = searchRange.Tables(i)
                        ' 检查表格是否在插入点附近
                        Dim distance = Math.Abs(tbl.Range.Start - tableEnd)
                        Debug.WriteLine($"Table {i}: start={tbl.Range.Start}, distance={distance}")

                        If distance < 100 Then
                            newTable = tbl
                            Debug.WriteLine($"Found matching table at position {tbl.Range.Start}")
                            Exit For
                        End If
                    Next
                Catch ex As Exception
                    Debug.WriteLine($"Find new table error: {ex.Message}")
                End Try

                If newTable IsNot Nothing Then
                    targetTable = newTable
                    Debug.WriteLine($"Successfully copied table - New table rows: {targetTable.Rows.Count}, cols: {targetTable.Columns.Count}")
                Else
                    Debug.WriteLine("Failed to get copied table, using original table")
                    targetTable = originalTable
                End If
            Else
                Debug.WriteLine("Replace mode: Using original table")
            End If

            ' 解析翻译文本（假设格式：行\n列\t）
            ' 使用None选项保留空行和空单元格
            Dim lines = translatedText.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)

            ' 确保行列数匹配
            Dim maxRows = Math.Min(lines.Length, targetTable.Rows.Count)

            Debug.WriteLine($"Processing table: {maxRows} rows, {targetTable.Columns.Count} columns")
            Debug.WriteLine($"Translated text lines: {lines.Length}")

            For rowIdx = 1 To maxRows
                ' 即使行为空也要处理，保持表格结构
                Dim line As String = ""
                If rowIdx - 1 < lines.Length Then
                    line = lines(rowIdx - 1)
                End If

                line = If(line, "").Trim()

                ' 始终处理这一行的所有列，即使line为空
                Dim cells() As String
                If String.IsNullOrEmpty(line) Then
                    ' 空行，创建对应列数的空数组
                    ReDim cells(targetTable.Columns.Count - 1)
                    For i = 0 To targetTable.Columns.Count - 1
                        cells(i) = ""
                    Next
                Else
                    ' 有内容的行，按tab分割
                    cells = line.Split(vbTab)
                End If

                ' 确保列数匹配
                Dim maxCols = Math.Min(cells.Length, targetTable.Columns.Count)

                For colIdx = 1 To targetTable.Columns.Count  ' 处理所有列
                    Try
                        Dim cell = targetTable.Cell(rowIdx, colIdx)
                        Dim cellText As String = ""

                        ' 获取对应单元格的文本
                        If colIdx - 1 < cells.Length Then
                            cellText = cells(colIdx - 1)
                        End If

                        cellText = If(cellText, "").Trim()

                        Debug.WriteLine($"Cell({rowIdx},{colIdx}): '{cellText}'")

                        ' 替换单元格内容，保留对象
                        ReplaceCellTextPreservingObjects(cell.Range, cellText)
                    Catch cellEx As Exception
                        Debug.WriteLine($"Apply cell ({rowIdx},{colIdx}) translation error: {cellEx.Message}")
                    End Try
                Next
            Next

        Catch ex As Exception
            Debug.WriteLine($"ApplyTableTranslation error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 替换范围中的文本，保留内嵌对象（图片、公式）
    ''' 采用分段删除法：从后往前删除对象之间的文本，然后插入译文
    ''' </summary>
    Private Sub ReplaceRangeTextPreservingObjects(range As Range, translatedText As String)
        Try
            If range Is Nothing Then Return

            ' 保存原始位置（后续操作会改变range的位置）
            Dim rangeStart = range.Start
            Dim rangeEnd = range.End

            ' 检查是否在表格单元格内（单元格结束符是Chr(7)，不能添加vbCr）
            Dim isInCell = False
            Try
                isInCell = range.Information(WdInformation.wdWithInTable)
            Catch
                isInCell = False
            End Try

            ' 检查是否有内嵌对象
            Dim hasObjects = False
            Try
                hasObjects = range.InlineShapes.Count > 0 OrElse range.OMaths.Count > 0
            Catch
                hasObjects = False
            End Try

            ' 没有对象：直接替换
            If Not hasObjects Then
                If isInCell Then
                    range.Text = translatedText
                Else
                    range.Text = translatedText & vbCr
                End If
                Return
            End If

            ' 有对象：收集所有对象在文档中的绝对位置
            Dim objRanges As New List(Of Tuple(Of Integer, Integer))()
            Try
                For Each shape As InlineShape In range.InlineShapes
                    If shape.Range.Start >= rangeStart AndAlso shape.Range.End <= rangeEnd Then
                        objRanges.Add(Tuple.Create(shape.Range.Start, shape.Range.End))
                    End If
                Next
            Catch
            End Try
            Try
                For Each omath As OMath In range.OMaths
                    If omath.Range.Start >= rangeStart AndAlso omath.Range.End <= rangeEnd Then
                        objRanges.Add(Tuple.Create(omath.Range.Start, omath.Range.End))
                    End If
                Next
            Catch
            End Try

            ' 去重并排序
            objRanges = objRanges _
                .Where(Function(r) r.Item2 > r.Item1) _
                .Distinct() _
                .OrderBy(Function(r) r.Item1) _
                .ToList()

            If objRanges.Count = 0 Then
                ' 实际上没有有效对象
                If isInCell Then
                    range.Text = translatedText
                Else
                    range.Text = translatedText & vbCr
                End If
                Return
            End If

            ' 确定内容结束位置（排除段落标记Chr(13)或单元格结束符Chr(7)）
            Dim contentEnd = rangeEnd
            If Not isInCell Then
                contentEnd = Math.Max(rangeStart, rangeEnd - 1)
            End If

            ' 从后往前删除对象之间的文本（这样前面的位置不会偏移）
            Dim textEnd = contentEnd
            For i = objRanges.Count - 1 To 0 Step -1
                Dim objStart = objRanges(i).Item1
                Dim objEnd = objRanges(i).Item2

                If textEnd > objEnd Then
                    Try
                        Dim delRange = _document.Range(objEnd, textEnd)
                        delRange.Text = ""
                    Catch delEx As Exception
                        Debug.WriteLine($"[WordTranslate] Delete text after object failed: {delEx.Message}")
                    End Try
                End If
                textEnd = objStart
            Next

            ' 删除第一个对象之前的文本
            If textEnd > rangeStart Then
                Try
                    Dim delRange = _document.Range(rangeStart, textEnd)
                    delRange.Text = ""
                Catch delEx As Exception
                    Debug.WriteLine($"[WordTranslate] Delete text before first object failed: {delEx.Message}")
                End Try
            End If

            ' 在段落开头插入译文
            Try
                Dim insertRange = _document.Range(rangeStart, rangeStart)
                insertRange.InsertBefore(translatedText)
            Catch insertEx As Exception
                Debug.WriteLine($"[WordTranslate] Insert translation failed: {insertEx.Message}")
            End Try

        Catch ex As Exception
            Debug.WriteLine($"[WordTranslate] ReplaceRangeTextPreservingObjects error: {ex.Message}")
            ' 最后的后备方案
            Try
                range.Text = translatedText & vbCr
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' 替换单元格文本，保留内嵌对象（图片、公式）
    ''' 采用与ReplaceRangeTextPreservingObjects相同的分段删除法
    ''' </summary>
    Private Sub ReplaceCellTextPreservingObjects(cellRange As Range, translatedText As String)
        Try
            If cellRange Is Nothing Then Return

            ' 保存原始位置
            Dim rangeStart = cellRange.Start
            Dim rangeEnd = cellRange.End

            ' 获取单元格文本（用于调试）
            Dim originalText = cellRange.Text
            If originalText Is Nothing Then originalText = ""
            originalText = originalText.TrimEnd(ChrW(7), ChrW(13), ChrW(10))

            ' 检查单元格是否有对象
            Dim hasObjects = False
            Try
                hasObjects = cellRange.InlineShapes.Count > 0 OrElse cellRange.OMaths.Count > 0
            Catch
                hasObjects = False
            End Try

            ' 调整范围以排除单元格结束符 Chr(7)
            Dim contentEnd = rangeEnd
            Try
                If cellRange.End > cellRange.Start Then
                    contentEnd = cellRange.End - 1
                    ' 如果末尾还有结束符，继续排除
                    While contentEnd > rangeStart
                        Dim lastChar = _document.Range(contentEnd - 1, contentEnd).Text
                        If lastChar = ChrW(7).ToString() OrElse lastChar = ChrW(13).ToString() Then
                            contentEnd -= 1
                        Else
                            Exit While
                        End If
                    End While
                End If
            Catch
            End Try

            Debug.WriteLine($"[WordTranslate] Cell replacement - HasObjects: {hasObjects}, Original: '{originalText}', New: '{translatedText}'")

            ' 没有对象：直接替换
            If Not hasObjects Then
                If contentEnd < rangeEnd Then
                    Dim adjustedRange = _document.Range(rangeStart, contentEnd)
                    adjustedRange.Text = translatedText
                Else
                    cellRange.Text = translatedText
                End If
                Return
            End If

            ' 有对象：收集对象位置
            Dim objRanges As New List(Of Tuple(Of Integer, Integer))()
            Try
                For Each shape As InlineShape In cellRange.InlineShapes
                    If shape.Range.Start >= rangeStart AndAlso shape.Range.End <= rangeEnd Then
                        objRanges.Add(Tuple.Create(shape.Range.Start, shape.Range.End))
                    End If
                Next
            Catch
            End Try
            Try
                For Each omath As OMath In cellRange.OMaths
                    If omath.Range.Start >= rangeStart AndAlso omath.Range.End <= rangeEnd Then
                        objRanges.Add(Tuple.Create(omath.Range.Start, omath.Range.End))
                    End If
                Next
            Catch
            End Try

            objRanges = objRanges _
                .Where(Function(r) r.Item2 > r.Item1) _
                .Distinct() _
                .OrderBy(Function(r) r.Item1) _
                .ToList()

            If objRanges.Count = 0 Then
                If contentEnd < rangeEnd Then
                    Dim adjustedRange = _document.Range(rangeStart, contentEnd)
                    adjustedRange.Text = translatedText
                Else
                    cellRange.Text = translatedText
                End If
                Return
            End If

            ' 从后往前删除对象之间的文本
            Dim textEnd = contentEnd
            For i = objRanges.Count - 1 To 0 Step -1
                Dim objStart = objRanges(i).Item1
                Dim objEnd = objRanges(i).Item2

                If textEnd > objEnd Then
                    Try
                        Dim delRange = _document.Range(objEnd, textEnd)
                        delRange.Text = ""
                    Catch delEx As Exception
                        Debug.WriteLine($"[WordTranslate] Cell delete text after object failed: {delEx.Message}")
                    End Try
                End If
                textEnd = objStart
            Next

            ' 删除第一个对象之前的文本
            If textEnd > rangeStart Then
                Try
                    Dim delRange = _document.Range(rangeStart, textEnd)
                    delRange.Text = ""
                Catch delEx As Exception
                    Debug.WriteLine($"[WordTranslate] Cell delete text before first object failed: {delEx.Message}")
                End Try
            End If

            ' 在单元格开头插入译文
            Try
                Dim insertRange = _document.Range(rangeStart, rangeStart)
                insertRange.InsertBefore(translatedText)
            Catch insertEx As Exception
                Debug.WriteLine($"[WordTranslate] Cell insert translation failed: {insertEx.Message}")
            End Try

        Catch ex As Exception
            Debug.WriteLine($"[WordTranslate] ReplaceCellTextPreservingObjects error: {ex.Message}")
            Try
                cellRange.Text = translatedText
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' 清理段落翻译结果中的换行符，防止Word创建新段落导致索引错乱
    ''' 表格单元格内的文本不调用此方法，保留多行能力
    ''' </summary>
    Private Function SanitizeParagraphText(text As String) As String
        If String.IsNullOrEmpty(text) Then Return text
        ' 将所有换行符替换为空格，避免创建新段落
        text = text.Replace(vbCrLf, " ").Replace(vbCr, " ").Replace(vbLf, " ")
        ' 合并多个连续空格
        While text.Contains("  ")
            text = text.Replace("  ", " ")
        End While
        Return text.Trim()
    End Function

    ''' <summary>
    ''' 后备方案：手动替换文本节点，保留对象
    ''' </summary>
    Private Sub ReplaceTextNodesOnly(range As Range, translatedText As String)
        Try
            ' 检查是否有对象：图片、表格、公式
            Dim hasObjects As Boolean = False
            Try
                hasObjects = range.InlineShapes.Count > 0 OrElse range.Tables.Count > 0 OrElse range.OMaths.Count > 0
            Catch
                hasObjects = False
            End Try

            If hasObjects Then
                ' 收集所有对象的位置
                Dim objectPositions As New List(Of Integer)()

                ' 添加InlineShapes（图片等）
                Try
                    For Each shape As InlineShape In range.InlineShapes
                        objectPositions.Add(shape.Range.Start - range.Start)
                    Next
                Catch
                End Try

                ' 添加Tables（表格）
                Try
                    For Each table As Table In range.Tables
                        objectPositions.Add(table.Range.Start - range.Start)
                    Next
                Catch
                End Try

                ' 添加OMaths（公式）
                Try
                    For Each omath As OMath In range.OMaths
                        objectPositions.Add(omath.Range.Start - range.Start)
                    Next
                Catch
                End Try

                ' 如果有对象，将译文插入到第一个对象之前
                If objectPositions.Count > 0 Then
                    objectPositions.Sort()
                    Dim firstObjPos = objectPositions(0)
                    Dim insertPos = range.Start + firstObjPos

                    ' 删除第一个对象之前的文本
                    If firstObjPos > 0 Then
                        Dim textRange = _document.Range(range.Start, insertPos)
                        textRange.Text = translatedText
                    Else
                        ' 对象在开头，在前面插入
                        range.InsertBefore(translatedText)
                    End If

                    ' 删除对象之后的文本
                    Dim lastObjEnd = range.Start

                    ' 检查InlineShapes
                    Try
                        For Each shape As InlineShape In range.InlineShapes
                            If shape.Range.End > lastObjEnd Then
                                lastObjEnd = shape.Range.End
                            End If
                        Next
                    Catch
                    End Try

                    ' 检查Tables
                    Try
                        For Each table As Table In range.Tables
                            If table.Range.End > lastObjEnd Then
                                lastObjEnd = table.Range.End
                            End If
                        Next
                    Catch
                    End Try

                    ' 检查OMaths
                    Try
                        For Each omath As OMath In range.OMaths
                            If omath.Range.End > lastObjEnd Then
                                lastObjEnd = omath.Range.End
                            End If
                        Next
                    Catch
                    End Try

                    If lastObjEnd < range.End - 1 Then
                        Dim afterTextRange = _document.Range(lastObjEnd, range.End - 1)
                        afterTextRange.Text = ""
                    End If
                End If
            Else
                ' 没有对象，直接替换
                range.Text = translatedText & vbCr
            End If
        Catch ex As Exception
            Debug.WriteLine($"ReplaceTextNodesOnly error: {ex.Message}")
            ' 最后的后备
            range.Text = translatedText & vbCr
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

            If isSelection AndAlso _selectionParagraphIndices IsNot Nothing Then
                ' 在选中区域后插入译文 - 通过索引重新获取Range
                For i = Math.Min(results.Count, _selectionParagraphIndices.Count) - 1 To 0 Step -1
                    Dim result = results(i)
                    If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                        Dim paraIndex = _selectionParagraphIndices(i)
                        If paraIndex > 0 AndAlso paraIndex <= _document.Paragraphs.Count Then
                            Try
                                Dim para = _document.Paragraphs(paraIndex)
                                InsertImmersiveTranslation(para.Range, result.TranslatedText, settings)
                            Catch ex As Exception
                                Debug.WriteLine($"[WordTranslate] Immersive mode selected paragraph {paraIndex} error: {ex.Message}")
                            End Try
                        End If
                    End If
                Next
            Else
                ' 使用翻译单元处理整个文档
                If _translationUnits IsNot Nothing AndAlso _translationUnits.Count > 0 Then
                    ' 收集所有需要处理的单元（从前向后处理，因为沉浸式模式是在原文后插入）
                    Dim unitsToProcess As New List(Of Tuple(Of Integer, TranslationUnit, TranslateParagraphResult))

                    For i = 0 To Math.Min(results.Count, _translationUnits.Count) - 1
                        Dim result = results(i)
                        If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                            unitsToProcess.Add(Tuple.Create(i, _translationUnits(i), result))
                        End If
                    Next

                    ' 从后向前处理，避免前面插入新段落导致后续段落索引偏移
                    For idx = unitsToProcess.Count - 1 To 0 Step -1
                        Try
                            Dim item = unitsToProcess(idx)
                            Dim index = item.Item1
                            Dim unit = item.Item2
                            Dim result = item.Item3

                            If unit.UnitType = TranslationUnitType.TableContent Then
                                ' 表格翻译 - 沉浸式模式
                                If unit.TableRef IsNot Nothing Then
                                    Try
                                        ' 沉浸式模式：复制表格，然后在新表格中填入翻译内容
                                        ApplyImmersiveTableTranslation(unit.TableRef, result.TranslatedText)
                                    Catch tableEx As Exception
                                        Debug.WriteLine($"Table translation error at index {index}: {tableEx.Message}")
                                    End Try
                                End If
                            Else
                                ' 普通文本翻译 - 通过段落索引重新获取Range
                                If unit.ParagraphIndex > 0 AndAlso unit.ParagraphIndex <= _document.Paragraphs.Count Then
                                    Try
                                        Dim para = _document.Paragraphs(unit.ParagraphIndex)
                                        InsertImmersiveTranslation(para.Range, result.TranslatedText, settings)
                                    Catch rangeEx As Exception
                                        Debug.WriteLine($"Range translation error at index {index}: {rangeEx.Message}")
                                    End Try
                                End If
                            End If
                        Catch unitEx As Exception
                            Debug.WriteLine($"Unit processing error: {unitEx.Message}")
                        End Try
                    Next
                Else
                    ' 后备方案：使用原有逻辑
                    Dim paras = _document.Paragraphs
                    For i = Math.Min(results.Count, paras.Count) - 1 To 0 Step -1
                        Dim result = results(i)
                        If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                            Dim para = paras(i + 1)
                            InsertImmersiveTranslation(para.Range, result.TranslatedText, settings)
                        End If
                    Next
                End If
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
    ''' 采用 InsertParagraphAfter + 格式复制，避免 Copy+Paste 的剪贴板依赖和 COM 异常
    ''' </summary>
    Private Sub InsertImmersiveTranslation(originalRange As Range, translatedText As String, settings As TranslateSettings)
        Try
            ' 检查Range是否有效
            If originalRange Is Nothing Then
                Debug.WriteLine("InsertImmersiveTranslation: originalRange is Nothing")
                Return
            End If

            ' 尝试获取Range文本，检查是否有效
            Try
                Dim testText = originalRange.Text
            Catch ex As Exception
                Debug.WriteLine($"InsertImmersiveTranslation: Range is invalid - {ex.Message}")
                Return
            End Try

            ' 保存原文格式信息（用于后续复制）
            Dim origFontName As String = ""
            Dim origFontSize As Single = 0
            Dim origFontBold As Integer = 0
            Dim origFontItalic As Integer = 0
            Dim origLeftIndent As Single = 0
            Dim origRightIndent As Single = 0
            Dim origFirstLineIndent As Single = 0
            Dim origAlignment As WdParagraphAlignment = WdParagraphAlignment.wdAlignParagraphLeft
            Dim origLineSpacing As Single = 0

            Try
                origFontName = originalRange.Font.Name
                origFontSize = originalRange.Font.Size
                origFontBold = originalRange.Font.Bold
                origFontItalic = originalRange.Font.Italic
                origLeftIndent = originalRange.ParagraphFormat.LeftIndent
                origRightIndent = originalRange.ParagraphFormat.RightIndent
                origFirstLineIndent = originalRange.ParagraphFormat.FirstLineIndent
                origAlignment = originalRange.ParagraphFormat.Alignment
                origLineSpacing = originalRange.ParagraphFormat.LineSpacing
            Catch fmtEx As Exception
                Debug.WriteLine($"[WordTranslate] 保存原文格式失败: {fmtEx.Message}")
            End Try

            ' 在原文段落后插入新段落（不依赖剪贴板）
            Dim paraEnd = originalRange.End
            Dim insertRange = _document.Range(paraEnd, paraEnd)
            insertRange.InsertParagraphAfter()

            ' 获取新插入的段落
            ' InsertParagraphAfter 在 paraEnd 处创建新段落，直接按位置获取即可
            Dim newPara = _document.Range(paraEnd, paraEnd).Paragraphs(1)
            If newPara Is Nothing Then Return

            Dim translatedRange = newPara.Range
            If translatedRange Is Nothing Then Return

            ' 移除新段落自带的段落结束符，写入翻译文本
            Try
                Dim contentEnd = translatedRange.End
                If translatedRange.End > translatedRange.Start Then
                    contentEnd = translatedRange.End - 1
                End If
                Dim textRange = _document.Range(translatedRange.Start, contentEnd)
                textRange.Text = translatedText
            Catch txtEx As Exception
                Debug.WriteLine($"[WordTranslate] 设置译文文本失败: {txtEx.Message}")
                translatedRange.Text = translatedText
            End Try

            ' 重新获取范围（文本修改后 Range 会变化）
            translatedRange = newPara.Range

            ' 复制原文格式到新段落
            Try
                If Not String.IsNullOrEmpty(origFontName) Then translatedRange.Font.Name = origFontName
                If origFontSize > 0 Then translatedRange.Font.Size = origFontSize
                translatedRange.Font.Bold = origFontBold
                translatedRange.Font.Italic = origFontItalic
                translatedRange.ParagraphFormat.LeftIndent = origLeftIndent
                translatedRange.ParagraphFormat.RightIndent = origRightIndent
                translatedRange.ParagraphFormat.FirstLineIndent = origFirstLineIndent
                translatedRange.ParagraphFormat.Alignment = origAlignment
                If origLineSpacing > 0 Then translatedRange.ParagraphFormat.LineSpacing = origLineSpacing
            Catch fmtEx As Exception
                Debug.WriteLine($"[WordTranslate] 复制格式失败: {fmtEx.Message}")
            End Try

            ' 重新获取范围（格式修改后 Range 会变化）
            translatedRange = newPara.Range

            ' 只有在不保持原文格式时才设置沉浸式样式
            If Not settings.PreserveFormatting Then
                ' 对整个段落应用样式
                Try
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
                            If origFontSize > 0 Then
                                .Size = CSng(origFontSize * settings.ImmersiveTranslationFontScale)
                            End If
                        End If
                    End With
                Catch fontEx As Exception
                    Debug.WriteLine($"Apply font style error: {fontEx.Message}")
                End Try

                ' 设置段落缩进（在原文缩进基础上额外缩进）
                Try
                    translatedRange.ParagraphFormat.LeftIndent = origLeftIndent + _wordApp.InchesToPoints(0.25)
                Catch indentEx As Exception
                    Debug.WriteLine($"Apply indent error: {indentEx.Message}")
                End Try
            End If

        Catch ex As Exception
            ' 单个段落插入失败时继续处理其他段落
            Debug.WriteLine("InsertImmersiveTranslation error: " & ex.Message)
            ' 最后的后备方案：尝试直接在原文后添加文本
            Try
                Dim endPos = originalRange.End
                Dim insertRange = _document.Range(endPos, endPos)
                insertRange.InsertAfter(vbCrLf & translatedText & vbCrLf)
            Catch finalEx As Exception
                Debug.WriteLine($"InsertImmersiveTranslation final fallback failed: {finalEx.Message}")
            End Try
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
