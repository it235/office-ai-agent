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
    Public Property Range As Range                  ' 原始范围
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
    Private _selectionRanges As List(Of Range)
    Private _translationUnits As List(Of TranslationUnit)  ' 新增：翻译单元列表

    Public Sub New(wordApp As Word.Application)
        MyBase.New()
        _wordApp = wordApp
        _document = wordApp.ActiveDocument
    End Sub

    ''' <summary>
    ''' 获取文档所有段落（支持表格切分）
    ''' </summary>
    Public Overrides Function GetAllParagraphs() As List(Of String)
        Dim paragraphs As New List(Of String)()
        _translationUnits = New List(Of TranslationUnit)()

        If _document Is Nothing Then Return paragraphs

        ' 遍历文档的Story Ranges（包括主文档、表格等）
        Dim mainStory = _document.StoryRanges(WdStoryType.wdMainTextStory)

        ' 先收集所有表格的位置
        Dim tablePositions As New Dictionary(Of Integer, Table)()
        For Each tbl As Table In _document.Tables
            tablePositions(tbl.Range.Start) = tbl
        Next

        ' 遍历段落，检测表格
        For Each para As Paragraph In _document.Paragraphs
            Try
                ' 检查该段落是否在表格内
                Dim isInTable As Boolean = False
                Dim tableRef As Table = Nothing

                Try
                    If para.Range.Tables.Count > 0 Then
                        isInTable = True
                        tableRef = para.Range.Tables(1)
                    End If
                Catch
                End Try

                If isInTable AndAlso tableRef IsNot Nothing Then
                    ' 该段落在表格内，检查是否已经处理过这个表格
                    Dim tableKey = tableRef.Range.Start
                    If tablePositions.ContainsKey(tableKey) Then
                        ' 第一次遇到这个表格，提取表格内容
                        Dim tableUnit = ExtractTableContent(tableRef)
                        _translationUnits.Add(tableUnit)
                        paragraphs.Add(tableUnit.Text)

                        ' 标记为已处理
                        tablePositions.Remove(tableKey)
                    End If
                    ' 跳过表格内的其他段落
                Else
                    ' 普通段落
                    Dim text = para.Range.Text
                    If text Is Nothing Then text = ""

                    text = text.TrimEnd(ChrW(13), ChrW(10), ChrW(7))

                    Dim unit As New TranslationUnit With {
                        .UnitType = TranslationUnitType.NormalText,
                        .Text = If(String.IsNullOrWhiteSpace(text), "", text),
                        .Range = para.Range
                    }

                    _translationUnits.Add(unit)
                    paragraphs.Add(unit.Text)
                End If

            Catch ex As Exception
                Debug.WriteLine($"Process paragraph error: {ex.Message}")
                ' 出错时添加空段落
                paragraphs.Add("")
                _translationUnits.Add(New TranslationUnit With {
                    .UnitType = TranslationUnitType.NormalText,
                    .Text = ""
                })
            End Try
        Next

        Return paragraphs
    End Function

    ''' <summary>
    ''' 提取表格内容为格式化文本（行\n列\t）
    ''' </summary>
    Private Function ExtractTableContent(table As Table) As TranslationUnit
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
            .Range = table.Range
        }
    End Function

    ''' <summary>
    ''' 获取选中的段落
    ''' </summary>
    Public Overrides Function GetSelectedParagraphs() As List(Of String)
        Dim paragraphs As New List(Of String)()
        _selectionRanges = New List(Of Range)()

        If _wordApp.Selection Is Nothing Then Return paragraphs

        Dim selRange = _wordApp.Selection.Range

        ' 检查selRange是否为Nothing或其Text属性为Nothing
        If selRange Is Nothing Then Return paragraphs

        Dim selText = selRange.Text
        If selText Is Nothing OrElse String.IsNullOrWhiteSpace(selText) Then
            Return paragraphs
        End If

        For Each para As Paragraph In selRange.Paragraphs
            Dim text = para.Range.Text

            ' 防止text为Nothing
            If text Is Nothing Then
                text = ""
            End If

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
                            ReplaceRangeTextPreservingObjects(range, result.TranslatedText)
                        End If
                    End If
                Next
            Else
                ' 替换整个文档 - 使用翻译单元
                If _translationUnits IsNot Nothing AndAlso _translationUnits.Count > 0 Then
                    For i = Math.Min(results.Count, _translationUnits.Count) - 1 To 0 Step -1
                        Dim result = results(i)
                        If result.Success AndAlso Not String.IsNullOrEmpty(result.TranslatedText) Then
                            Dim unit = _translationUnits(i)

                            If unit.UnitType = TranslationUnitType.TableContent Then
                                ' 表格翻译 - 直接替换单元格内容
                                ApplyTableTranslation(unit.TableRef, result.TranslatedText, False)
                            Else
                                ' 普通文本翻译
                                If unit.Range IsNot Nothing Then
                                    ReplaceRangeTextPreservingObjects(unit.Range, result.TranslatedText)
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
                            ReplaceRangeTextPreservingObjects(para.Range, result.TranslatedText)
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
    ''' 替换范围中的文本，保留内嵌对象（图片、表格、公式）
    ''' 注意：表格会在GetAllParagraphs中被切分，所以这里不应该遇到表格
    ''' </summary>
    Private Sub ReplaceRangeTextPreservingObjects(range As Range, translatedText As String)
        Try
            ' 获取原始文本（排除段落符）
            Dim originalText = range.Text

            ' 检查originalText是否为Nothing或空
            If originalText Is Nothing Then
                originalText = ""
            End If

            If originalText.EndsWith(vbCr) Then
                originalText = originalText.Substring(0, originalText.Length - 1)
            End If
            If originalText.EndsWith(vbLf) Then
                originalText = originalText.Substring(0, originalText.Length - 1)
            End If

            ' 检查是否有内嵌对象：图片、公式
            Dim hasObjects As Boolean = False
            Try
                hasObjects = range.InlineShapes.Count > 0 OrElse range.OMaths.Count > 0
            Catch
                ' 如果检查失败，假设有对象（保守处理）
                hasObjects = True
            End Try

            ' 如果没有内嵌对象，直接替换（更快）
            If Not hasObjects Then
                range.Text = translatedText & vbCr
                Return
            End If

            ' 有对象时，使用Find.Execute替换文本部分
            If Not String.IsNullOrWhiteSpace(originalText) Then
                With range.Find
                    .ClearFormatting()
                    .Replacement.ClearFormatting()
                    .Text = originalText
                    .Replacement.Text = translatedText
                    .Forward = True
                    .Wrap = WdFindWrap.wdFindStop
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False

                    ' 执行替换
                    Dim replaced = .Execute(Replace:=WdReplace.wdReplaceOne)

                    ' 如果Find替换失败（可能因为文本不匹配），使用后备方案
                    If Not replaced Then
                        Debug.WriteLine($"Find.Replace failed, using fallback method")
                        ' 备用方案：手动替换文本节点
                        ReplaceTextNodesOnly(range, translatedText)
                    End If
                End With
            Else
                ' 原文为空，直接在开头插入译文
                range.InsertBefore(translatedText)
            End If

        Catch ex As Exception
            Debug.WriteLine($"ReplaceRangeTextPreservingObjects error: {ex.Message}")
            ' 如果所有方法都失败，最后的后备方案
            Try
                range.Text = translatedText & vbCr
            Catch
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' 处理包含表格的范围翻译
    ''' 策略：将表格内容提取为单独的翻译单元，逐个翻译
    ''' </summary>

    ''' <summary>
    ''' 替换单元格文本，保留内嵌对象（图片、公式）
    ''' </summary>
    Private Sub ReplaceCellTextPreservingObjects(cellRange As Range, translatedText As String)
        Try
            ' 获取单元格文本
            Dim originalText = cellRange.Text
            If originalText Is Nothing Then originalText = ""

            ' 移除单元格结束符 (Chr(7) 是单元格结束符, Chr(13) 是段落符)
            originalText = originalText.TrimEnd(ChrW(7), ChrW(13), ChrW(10))

            ' 检查单元格是否有对象
            Dim hasObjects As Boolean = False
            Try
                hasObjects = cellRange.InlineShapes.Count > 0 OrElse cellRange.OMaths.Count > 0
            Catch
                hasObjects = False
            End Try

            ' 调整cellRange以排除单元格结束符
            Dim adjustedRange As Range = Nothing
            Try
                If cellRange.End > cellRange.Start Then
                    adjustedRange = _document.Range(cellRange.Start, cellRange.End - 1)
                    ' 再次检查是否还有结束符
                    While adjustedRange.End > adjustedRange.Start AndAlso
                          (adjustedRange.Characters.Last.Text = ChrW(7) OrElse
                           adjustedRange.Characters.Last.Text = ChrW(13))
                        adjustedRange.End = adjustedRange.End - 1
                    End While
                Else
                    adjustedRange = cellRange
                End If
            Catch
                adjustedRange = cellRange
            End Try

            Debug.WriteLine($"Cell replacement - HasObjects: {hasObjects}, Original: '{originalText}', New: '{translatedText}'")

            If Not hasObjects Then
                ' 没有对象，直接替换
                adjustedRange.Text = translatedText
                Return
            End If

            ' 有对象，使用Find替换
            If Not String.IsNullOrWhiteSpace(originalText) Then
                Try
                    With adjustedRange.Find
                        .ClearFormatting()
                        .Replacement.ClearFormatting()
                        .Text = originalText
                        .Replacement.Text = translatedText
                        .Forward = True
                        .Wrap = WdFindWrap.wdFindStop
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False

                        ' 执行替换
                        Dim replaced = .Execute(Replace:=WdReplace.wdReplaceOne)
                        Debug.WriteLine($"Find.Execute result: {replaced}")
                    End With
                Catch findEx As Exception
                    Debug.WriteLine($"Find.Execute failed: {findEx.Message}")
                    ' 备用方案
                    adjustedRange.Text = translatedText
                End Try
            Else
                ' 原文为空，直接设置译文
                adjustedRange.Text = translatedText
            End If

        Catch ex As Exception
            Debug.WriteLine($"ReplaceCellTextPreservingObjects error: {ex.Message}")
            ' 最后的后备方案
            Try
                ' 尝试直接替换，但这可能会丢失对象
                cellRange.Text = translatedText
            Catch
            End Try
        End Try
    End Sub

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
                ' 使用翻译单元处理整个文档
                If _translationUnits IsNot Nothing AndAlso _translationUnits.Count > 0 Then
                    ' 收集所有需要处理的单元
                    Dim unitsToProcess As New List(Of Tuple(Of Integer, TranslationUnit, TranslateParagraphResult))

                    For i = 0 To Math.Min(results.Count, _translationUnits.Count) - 1
                        Dim result = results(i)
                        If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                            unitsToProcess.Add(Tuple.Create(i, _translationUnits(i), result))
                        End If
                    Next

                    ' 从前向后处理，避免Range失效
                    For Each item In unitsToProcess
                        Try
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
                                ' 普通文本翻译
                                If unit.Range IsNot Nothing Then
                                    Try
                                        InsertImmersiveTranslation(unit.Range, result.TranslatedText, settings)
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
            ' 获取原段落的末尾位置
            Dim paraEnd = originalRange.End

            ' 复制原段落到其后面（保留所有内容和格式）
            Try
                originalRange.Copy()
            Catch copyEx As Exception
                Debug.WriteLine($"InsertImmersiveTranslation: Copy failed - {copyEx.Message}")
                Return
            End Try

            Dim insertPoint = _document.Range(paraEnd, paraEnd)
            Try
                insertPoint.Paste()
            Catch pasteEx As Exception
                Debug.WriteLine($"InsertImmersiveTranslation: Paste failed - {pasteEx.Message}")
                Return
            End Try
            ' 获取新粘贴的段落范围
            Dim translatedStart = paraEnd
            Dim newParagraph = _document.Range(translatedStart, translatedStart).Paragraphs(1)
            If newParagraph Is Nothing Then Return

            Dim translatedRange = newParagraph.Range
            If translatedRange Is Nothing Then Return

            ' 使用Find.Execute来替换文本，保留内嵌对象
            ' 获取原始文本（排除段落符）
            Dim originalText = ""
            Try
                originalText = originalRange.Text

                ' 检查originalText是否为Nothing或空
                If originalText Is Nothing Then
                    originalText = ""
                End If

                If originalText.EndsWith(vbCr) Then
                    originalText = originalText.Substring(0, originalText.Length - 1)
                End If
                If originalText.EndsWith(vbLf) Then
                    originalText = originalText.Substring(0, originalText.Length - 1)
                End If
            Catch rangeEx As Exception
                Debug.WriteLine($"InsertImmersiveTranslation: Failed to get original text - {rangeEx.Message}")
                originalText = ""
            End Try

            ' 如果原文本为空，直接设置译文
            If String.IsNullOrWhiteSpace(originalText) Then
                ' 删除段落符之前的内容
                Dim textEnd = translatedRange.End - 1
                If textEnd > translatedRange.Start Then
                    Dim textRange = _document.Range(translatedRange.Start, textEnd)
                    textRange.Text = translatedText
                End If
            Else
                ' 使用Find替换文本部分，保留对象
                Try
                    With translatedRange.Find
                        .ClearFormatting()
                        .Replacement.ClearFormatting()
                        .Text = originalText
                        .Replacement.Text = translatedText
                        .Forward = True
                        .Wrap = WdFindWrap.wdFindStop
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False

                        ' 执行替换
                        .Execute(Replace:=WdReplace.wdReplaceOne)
                    End With
                Catch findEx As Exception
                    Debug.WriteLine($"InsertImmersiveTranslation: Find.Execute failed - {findEx.Message}")
                    ' 备用方案：直接设置文本
                    Try
                        translatedRange.Text = translatedText & vbCr
                    Catch directEx As Exception
                        Debug.WriteLine($"InsertImmersiveTranslation: Direct text set failed - {directEx.Message}")
                    End Try
                End Try
            End If

            ' 重新获取段落范围（因为文本可能变化）
            newParagraph = _document.Range(translatedStart, translatedStart).Paragraphs(1)
            If newParagraph Is Nothing Then Return
            translatedRange = newParagraph.Range

            ' 只有在不保持原文格式时才设置样式
            If Not settings.PreserveFormatting Then
                ' 对整个段落应用样式（不影响对象）
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
                            Dim originalSize = originalRange.Font.Size
                            If originalSize > 0 Then
                                .Size = CSng(originalSize * settings.ImmersiveTranslationFontScale)
                            End If
                        End If
                    End With
                Catch fontEx As Exception
                    Debug.WriteLine($"Apply font style error: {fontEx.Message}")
                End Try

                ' 设置段落缩进
                Try
                    translatedRange.ParagraphFormat.LeftIndent = originalRange.ParagraphFormat.LeftIndent + _wordApp.InchesToPoints(0.25)
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
                insertRange.InsertAfter(translatedText & vbCrLf)
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
