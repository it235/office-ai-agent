Imports System.Diagnostics
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

''' <summary>
''' PowerPoint幻灯片项（用于跟踪文本位置）
''' </summary>
Public Class SlideTextItem
    Public Property SlideIndex As Integer
    Public Property ShapeIndex As Integer
    Public Property ShapeName As String
    Public Property Text As String
    Public Property Shape As Shape
End Class

''' <summary>
''' PowerPoint文档翻译服务 - 支持全幻灯片翻译
''' </summary>
Public Class PowerPointDocumentTranslateService
    Inherits DocumentTranslateService

    Private _pptApp As PowerPoint.Application
    Private _presentation As Presentation
    Private _textItems As List(Of SlideTextItem)

    Public Sub New(pptApp As PowerPoint.Application)
        MyBase.New()
        _pptApp = pptApp
        _presentation = pptApp.ActivePresentation
        _textItems = New List(Of SlideTextItem)()
    End Sub

    ''' <summary>
    ''' 获取所有幻灯片的所有文本
    ''' </summary>
    Public Overrides Function GetAllParagraphs() As List(Of String)
        Dim texts As New List(Of String)()
        _textItems.Clear()

        If _presentation Is Nothing Then Return texts

        For slideIdx = 1 To _presentation.Slides.Count
            Dim slide = _presentation.Slides(slideIdx)
            ExtractTextFromSlide(slide, slideIdx, texts)
        Next

        Return texts
    End Function

    ''' <summary>
    ''' 从单个幻灯片提取文本
    ''' </summary>
    Private Sub ExtractTextFromSlide(slide As Slide, slideIndex As Integer, texts As List(Of String))
        For shapeIdx = 1 To slide.Shapes.Count
            Dim shape = slide.Shapes(shapeIdx)
            ExtractTextFromShape(shape, slideIndex, shapeIdx, texts)
        Next
    End Sub

    ''' <summary>
    ''' 从形状中提取文本
    ''' </summary>
    Private Sub ExtractTextFromShape(shape As Shape, slideIndex As Integer, shapeIndex As Integer, texts As List(Of String))
        Try
            ' 处理组合形状
            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoGroup Then
                For i = 1 To shape.GroupItems.Count
                    ExtractTextFromShape(shape.GroupItems(i), slideIndex, shapeIndex, texts)
                Next
                Return
            End If

            ' 处理表格
            If shape.HasTable Then
                Dim table = shape.Table
                For row = 1 To table.Rows.Count
                    For col = 1 To table.Columns.Count
                        Dim cell = table.Cell(row, col)
                        If cell.Shape.HasTextFrame Then
                            Dim text = cell.Shape.TextFrame.TextRange.Text
                            If Not String.IsNullOrWhiteSpace(text) Then
                                texts.Add(text.Trim())
                                _textItems.Add(New SlideTextItem() With {
                                    .SlideIndex = slideIndex,
                                    .ShapeIndex = shapeIndex,
                                    .ShapeName = $"Table({row},{col})",
                                    .Text = text.Trim(),
                                    .Shape = cell.Shape
                                })
                            End If
                        End If
                    Next
                Next
                Return
            End If

            ' 处理普通文本框
            If shape.HasTextFrame Then
                Dim textFrame = shape.TextFrame
                If textFrame.HasText Then
                    Dim text = textFrame.TextRange.Text
                    If Not String.IsNullOrWhiteSpace(text) Then
                        texts.Add(text.Trim())
                        _textItems.Add(New SlideTextItem() With {
                            .SlideIndex = slideIndex,
                            .ShapeIndex = shapeIndex,
                            .ShapeName = shape.Name,
                            .Text = text.Trim(),
                            .Shape = shape
                        })
                    End If
                End If
            End If
        Catch
            ' 忽略无法访问的形状
        End Try
    End Sub

    ''' <summary>
    ''' 获取选中的文本
    ''' </summary>
    Public Overrides Function GetSelectedParagraphs() As List(Of String)
        Dim texts As New List(Of String)()
        _textItems.Clear()

        Try
            Dim sel = _pptApp.ActiveWindow.Selection

            Select Case sel.Type
                Case PpSelectionType.ppSelectionText
                    ' 选中了文本
                    Dim text = sel.TextRange.Text
                    If Not String.IsNullOrWhiteSpace(text) Then
                        texts.Add(text.Trim())
                        _textItems.Add(New SlideTextItem() With {
                            .SlideIndex = sel.SlideRange(1).SlideIndex,
                            .ShapeIndex = 0,
                            .ShapeName = "Selection",
                            .Text = text.Trim(),
                            .Shape = sel.ShapeRange(1)
                        })
                    End If

                Case PpSelectionType.ppSelectionShapes
                    ' 选中了形状
                    For i = 1 To sel.ShapeRange.Count
                        Dim shape = sel.ShapeRange(i)
                        Dim slideIdx = sel.SlideRange(1).SlideIndex
                        ExtractTextFromShapeForSelection(shape, slideIdx, i, texts)
                    Next

                Case PpSelectionType.ppSelectionSlides
                    ' 选中了幻灯片
                    For i = 1 To sel.SlideRange.Count
                        Dim slide = sel.SlideRange(i)
                        ExtractTextFromSlide(slide, slide.SlideIndex, texts)
                    Next
            End Select
        Catch
            ' 无选区时返回空列表
        End Try

        Return texts
    End Function

    ''' <summary>
    ''' 从选中的形状提取文本
    ''' </summary>
    Private Sub ExtractTextFromShapeForSelection(shape As Shape, slideIndex As Integer, shapeIndex As Integer, texts As List(Of String))
        Try
            If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                Dim text = shape.TextFrame.TextRange.Text
                If Not String.IsNullOrWhiteSpace(text) Then
                    texts.Add(text.Trim())
                    _textItems.Add(New SlideTextItem() With {
                        .SlideIndex = slideIndex,
                        .ShapeIndex = shapeIndex,
                        .ShapeName = shape.Name,
                        .Text = text.Trim(),
                        .Shape = shape
                    })
                End If
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 应用翻译结果到所有幻灯片
    ''' </summary>
    Public Overrides Sub ApplyTranslation(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        If results Is Nothing OrElse results.Count = 0 Then Return

        Select Case outputMode
            Case TranslateOutputMode.Replace
                ApplyReplaceMode(results)
            Case TranslateOutputMode.Immersive
                ApplyImmersiveMode(results)
            Case TranslateOutputMode.NewDocument
                ApplyToNewPresentation(results)
            Case TranslateOutputMode.SidePanel
                ' 侧栏模式由调用者处理
        End Select
    End Sub

    ''' <summary>
    ''' 应用翻译结果到选中内容
    ''' </summary>
    Public Overrides Sub ApplyTranslationToSelection(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)
        ApplyTranslation(results, outputMode)
    End Sub

    ''' <summary>
    ''' 替换模式 - 直接替换原文
    ''' </summary>
    Private Sub ApplyReplaceMode(results As List(Of TranslateParagraphResult))
        Try
            For i = 0 To Math.Min(results.Count, _textItems.Count) - 1
                Dim result = results(i)
                Dim item = _textItems(i)

                If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                    Try
                        If item.Shape IsNot Nothing AndAlso item.Shape.HasTextFrame Then
                            item.Shape.TextFrame.TextRange.Text = result.TranslatedText
                        End If
                    Catch
                        ' 忽略无法修改的形状
                    End Try
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("应用翻译结果时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 沉浸式翻译模式 - 在每页后面复制一页，替换为译文
    ''' </summary>
    Private Sub ApplyImmersiveMode(results As List(Of TranslateParagraphResult))
        Try
            Dim settings = TranslateSettings.Load()

            ' 按幻灯片分组结果
            Dim slideResults As New Dictionary(Of Integer, List(Of Tuple(Of SlideTextItem, TranslateParagraphResult)))()

            For i = 0 To Math.Min(results.Count, _textItems.Count) - 1
                Dim item = _textItems(i)
                Dim result = results(i)

                If Not slideResults.ContainsKey(item.SlideIndex) Then
                    slideResults(item.SlideIndex) = New List(Of Tuple(Of SlideTextItem, TranslateParagraphResult))()
                End If
                slideResults(item.SlideIndex).Add(Tuple.Create(item, result))
            Next

            ' 从后往前处理幻灯片，避免索引变化问题
            Dim slideIndices = slideResults.Keys.OrderByDescending(Function(x) x).ToList()

            For Each slideIdx In slideIndices
                Try
                    Dim originalSlide = _presentation.Slides(slideIdx)

                    ' 使用Duplicate方法完整复制幻灯片（包括所有背景格式）
                    Dim newSlide As Slide = Nothing
                    Try
                        ' Duplicate会在原幻灯片后面创建完全相同的副本
                        newSlide = originalSlide.Duplicate()(1)
                        
                        ' 将新幻灯片移动到正确的位置
                        If newSlide.SlideIndex <> slideIdx + 1 Then
                            newSlide.MoveTo(slideIdx + 1)
                        End If
                    Catch dupEx As Exception
                        Debug.WriteLine($"Duplicate幻灯片 {slideIdx} 失败: {dupEx.Message}")
                        ' 如果Duplicate失败，尝试使用Copy/Paste作为后备方案
                        Try
                            originalSlide.Copy()
                            newSlide = _presentation.Slides.Paste(slideIdx + 1)(1)
                            
                            ' 手动复制背景格式
                            If originalSlide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse Then
                                newSlide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse
                                If originalSlide.Background.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    newSlide.Background.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                                    Select Case originalSlide.Background.Fill.Type
                                        Case Microsoft.Office.Core.MsoFillType.msoFillSolid
                                            newSlide.Background.Fill.Solid()
                                            newSlide.Background.Fill.ForeColor.RGB = originalSlide.Background.Fill.ForeColor.RGB
                                        Case Microsoft.Office.Core.MsoFillType.msoFillPatterned
                                            newSlide.Background.Fill.Patterned(originalSlide.Background.Fill.Pattern)
                                            newSlide.Background.Fill.ForeColor.RGB = originalSlide.Background.Fill.ForeColor.RGB
                                            newSlide.Background.Fill.BackColor.RGB = originalSlide.Background.Fill.BackColor.RGB
                                        Case Microsoft.Office.Core.MsoFillType.msoFillGradient
                                            ' 渐变填充需要复制更多属性
                                            ' 由于渐变属性复杂，这里依赖Paste的自动复制
                                        Case Microsoft.Office.Core.MsoFillType.msoFillTextured
                                            ' 纹理填充
                                        Case Microsoft.Office.Core.MsoFillType.msoFillPicture
                                            ' 图片填充
                                    End Select
                                End If
                            Else
                                newSlide.FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoTrue
                            End If
                        Catch copyEx As Exception
                            Debug.WriteLine($"Copy/Paste幻灯片 {slideIdx} 也失败: {copyEx.Message}")
                            Continue For
                        End Try
                    End Try

                    If newSlide Is Nothing Then Continue For

                    ' 在新幻灯片上替换为译文
                    Dim slideItems = slideResults(slideIdx)
                    ApplyTranslationToSlide(newSlide, slideItems, settings)
                Catch ex As Exception
                    Debug.WriteLine($"处理幻灯片 {slideIdx} 时出错: {ex.Message}")
                End Try
            Next

        Catch ex As Exception
            MessageBox.Show("应用沉浸式翻译时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 在指定幻灯片上应用翻译
    ''' </summary>
    Private Sub ApplyTranslationToSlide(slide As Slide, items As List(Of Tuple(Of SlideTextItem, TranslateParagraphResult)), settings As TranslateSettings)
        ' 收集新幻灯片上的文本项
        Dim newTextItems As New List(Of SlideTextItem)()
        For shapeIdx = 1 To slide.Shapes.Count
            Dim shape = slide.Shapes(shapeIdx)
            CollectTextItemsFromShape(shape, slide.SlideIndex, shapeIdx, newTextItems)
        Next

        ' 按原始顺序匹配并替换
        Dim matchIndex = 0
        For Each itemPair In items
            Dim originalItem = itemPair.Item1
            Dim result = itemPair.Item2

            If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                ' 尝试在新幻灯片上找到对应的文本框
                If matchIndex < newTextItems.Count Then
                    Dim newItem = newTextItems(matchIndex)
                    Try
                        If newItem.Shape IsNot Nothing AndAlso newItem.Shape.HasTextFrame Then
                            newItem.Shape.TextFrame.TextRange.Text = result.TranslatedText

                            ' 如果不保持原文格式，设置自定义样式
                            If Not settings.PreserveFormatting Then
                                Dim textRange = newItem.Shape.TextFrame.TextRange
                                Try
                                    Dim colorHex = settings.ImmersiveTranslationColor.TrimStart("#"c)
                                    If colorHex.Length >= 6 Then
                                        Dim r = Convert.ToInt32(colorHex.Substring(0, 2), 16)
                                        Dim g = Convert.ToInt32(colorHex.Substring(2, 2), 16)
                                        Dim b = Convert.ToInt32(colorHex.Substring(4, 2), 16)
                                        textRange.Font.Color.RGB = RGB(r, g, b)
                                    End If
                                Catch
                                End Try

                                If settings.ImmersiveTranslationItalic Then
                                    textRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue
                                End If
                            End If
                        End If
                    Catch
                    End Try
                End If
            End If
            matchIndex += 1
        Next
    End Sub

    ''' <summary>
    ''' 创建新演示文稿并写入翻译结果
    ''' </summary>
    Private Sub ApplyToNewPresentation(results As List(Of TranslateParagraphResult))
        Try
            ' 复制当前演示文稿
            Dim newPres = _pptApp.Presentations.Add()

            ' 复制所有幻灯片
            For slideIdx = 1 To _presentation.Slides.Count
                _presentation.Slides(slideIdx).Copy()
                newPres.Slides.Paste()
            Next

            ' 应用翻译到新演示文稿
            Dim newService As New PowerPointDocumentTranslateService(_pptApp)
            newService._presentation = newPres
            newService._textItems = _textItems

            ' 重新收集新演示文稿的文本项
            Dim newTextItems As New List(Of SlideTextItem)()
            For slideIdx = 1 To newPres.Slides.Count
                Dim slide = newPres.Slides(slideIdx)
                CollectTextItemsFromSlide(slide, slideIdx, newTextItems)
            Next

            ' 应用翻译
            For i = 0 To Math.Min(results.Count, newTextItems.Count) - 1
                Dim result = results(i)
                Dim item = newTextItems(i)

                If result.Success AndAlso Not String.IsNullOrWhiteSpace(result.TranslatedText) Then
                    Try
                        If item.Shape IsNot Nothing AndAlso item.Shape.HasTextFrame Then
                            item.Shape.TextFrame.TextRange.Text = result.TranslatedText
                        End If
                    Catch
                    End Try
                End If
            Next

            newPres.Windows(1).Activate()
        Catch ex As Exception
            MessageBox.Show("创建新演示文稿时出错：" & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 收集幻灯片中的文本项
    ''' </summary>
    Private Sub CollectTextItemsFromSlide(slide As Slide, slideIndex As Integer, items As List(Of SlideTextItem))
        For shapeIdx = 1 To slide.Shapes.Count
            Dim shape = slide.Shapes(shapeIdx)
            CollectTextItemsFromShape(shape, slideIndex, shapeIdx, items)
        Next
    End Sub

    ''' <summary>
    ''' 从形状收集文本项
    ''' </summary>
    Private Sub CollectTextItemsFromShape(shape As Shape, slideIndex As Integer, shapeIndex As Integer, items As List(Of SlideTextItem))
        Try
            If shape.Type = Microsoft.Office.Core.MsoShapeType.msoGroup Then
                For i = 1 To shape.GroupItems.Count
                    CollectTextItemsFromShape(shape.GroupItems(i), slideIndex, shapeIndex, items)
                Next
                Return
            End If

            If shape.HasTable Then
                Dim table = shape.Table
                For row = 1 To table.Rows.Count
                    For col = 1 To table.Columns.Count
                        Dim cell = table.Cell(row, col)
                        If cell.Shape.HasTextFrame AndAlso cell.Shape.TextFrame.HasText Then
                            items.Add(New SlideTextItem() With {
                                .SlideIndex = slideIndex,
                                .ShapeIndex = shapeIndex,
                                .ShapeName = $"Table({row},{col})",
                                .Text = cell.Shape.TextFrame.TextRange.Text,
                                .Shape = cell.Shape
                            })
                        End If
                    Next
                Next
                Return
            End If

            If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                items.Add(New SlideTextItem() With {
                    .SlideIndex = slideIndex,
                    .ShapeIndex = shapeIndex,
                    .ShapeName = shape.Name,
                    .Text = shape.TextFrame.TextRange.Text,
                    .Shape = shape
                })
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 生成翻译结果的格式化文本（用于侧栏显示）
    ''' </summary>
    Public Function FormatResultsForDisplay(results As List(Of TranslateParagraphResult), showOriginal As Boolean) As String
        Dim sb As New StringBuilder()
        Dim currentSlide = -1

        For i = 0 To Math.Min(results.Count, _textItems.Count) - 1
            Dim result = results(i)
            Dim item = _textItems(i)

            ' 添加幻灯片分隔
            If item.SlideIndex <> currentSlide Then
                currentSlide = item.SlideIndex
                sb.AppendLine()
                sb.AppendLine($"=== 幻灯片 {currentSlide} ===")
                sb.AppendLine()
            End If

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

            sb.AppendLine()
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 获取幻灯片统计信息
    ''' </summary>
    Public Function GetStatistics() As String
        Dim slideCount = _presentation.Slides.Count
        Dim textCount = _textItems.Count

        Return $"共 {slideCount} 张幻灯片，{textCount} 个文本块"
    End Function
End Class
