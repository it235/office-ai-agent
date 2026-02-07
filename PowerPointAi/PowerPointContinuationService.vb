Imports System.Diagnostics
Imports System.Text
Imports Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

''' <summary>
''' PowerPoint幻灯片续写服务 - 获取光标上下文并插入续写内容
''' </summary>
Public Class PowerPointContinuationService
    Inherits ContinuationService

    Private _pptApp As PowerPoint.Application
    Private _presentation As Presentation
    Private _currentSlide As Slide
    Private _currentShape As Shape
    Private _currentTextRange As TextRange

    Public Sub New(pptApp As PowerPoint.Application)
        _pptApp = pptApp
        _presentation = pptApp.ActivePresentation
    End Sub

    ''' <summary>
    ''' 检查是否可以进行续写
    ''' </summary>
    Public Overrides Function CanContinue() As Boolean
        If _pptApp Is Nothing OrElse _presentation Is Nothing Then
            Return False
        End If

        Try
            Dim sel = _pptApp.ActiveWindow.Selection
            Return sel IsNot Nothing
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 获取光标位置的上下文
    ''' </summary>
    ''' <param name="paragraphsBefore">光标前要获取的段落数</param>
    ''' <param name="paragraphsAfter">光标后要获取的段落数</param>
    Public Overrides Function GetCursorContext(paragraphsBefore As Integer, paragraphsAfter As Integer) As ContinuationContext
        Try
            Dim sel = _pptApp.ActiveWindow.Selection
            If sel Is Nothing Then Return Nothing

            Dim context As New ContinuationContext()

            ' 获取当前幻灯片
            Try
                _currentSlide = _pptApp.ActiveWindow.View.Slide
                context.CursorPosition = _currentSlide.SlideIndex
            Catch
                If _presentation.Slides.Count > 0 Then
                    _currentSlide = _presentation.Slides(1)
                    context.CursorPosition = 1
                Else
                    Return Nothing
                End If
            End Try

            context.DocumentPath = If(_presentation.Path, "")

            ' 判断选择类型并获取上下文
            Select Case sel.Type
                Case PpSelectionType.ppSelectionText
                    ' 在文本框中编辑
                    Return GetTextSelectionContext(sel, context, paragraphsBefore, paragraphsAfter)

                Case PpSelectionType.ppSelectionShapes
                    ' 选中了形状
                    Return GetShapeSelectionContext(sel, context)

                Case Else
                    ' 其他情况，获取当前幻灯片内容作为上下文
                    Return GetSlideContext(context)
            End Select

        Catch ex As Exception
            Debug.WriteLine($"GetCursorContext 出错: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 获取文本选择的上下文
    ''' </summary>
    Private Function GetTextSelectionContext(sel As Selection, context As ContinuationContext, paragraphsBefore As Integer, paragraphsAfter As Integer) As ContinuationContext
        Try
            _currentTextRange = sel.TextRange
            _currentShape = sel.ShapeRange(1)

            ' 获取整个文本框的文本
            Dim fullText As String = ""
            If _currentShape.HasTextFrame AndAlso _currentShape.TextFrame.HasText Then
                fullText = _currentShape.TextFrame.TextRange.Text
            End If

            ' 获取光标位置（选中文本的起始位置）
            Dim cursorPos = _currentTextRange.Start
            context.CursorOffsetInParagraph = cursorPos

            ' 将文本按段落分割（PPT中用vbCr分隔段落）
            Dim paragraphs = fullText.Split(New Char() {ChrW(13)}, StringSplitOptions.None)

            ' 找到光标所在段落
            Dim currentParaIndex = 0
            Dim charCount = 0
            For i = 0 To paragraphs.Length - 1
                charCount += paragraphs(i).Length + 1 ' +1 for the paragraph mark
                If charCount >= cursorPos Then
                    currentParaIndex = i
                    Exit For
                End If
            Next

            ' 获取当前段落
            If currentParaIndex < paragraphs.Length Then
                context.CurrentParagraphText = paragraphs(currentParaIndex).Trim()
            End If

            ' 获取前面的段落
            Dim beforeBuilder As New StringBuilder()
            Dim startIdx = Math.Max(0, currentParaIndex - paragraphsBefore)
            For i = startIdx To currentParaIndex - 1
                If Not String.IsNullOrWhiteSpace(paragraphs(i)) Then
                    beforeBuilder.AppendLine(paragraphs(i).Trim())
                End If
            Next
            context.ContextBefore = beforeBuilder.ToString().TrimEnd()

            ' 获取后面的段落
            Dim afterBuilder As New StringBuilder()
            Dim endIdx = Math.Min(paragraphs.Length - 1, currentParaIndex + paragraphsAfter)
            For i = currentParaIndex + 1 To endIdx
                If Not String.IsNullOrWhiteSpace(paragraphs(i)) Then
                    afterBuilder.AppendLine(paragraphs(i).Trim())
                End If
            Next
            context.ContextAfter = afterBuilder.ToString().TrimEnd()

            ' 判断位置类型
            If currentParaIndex = 0 Then
                context.PositionType = CursorPositionType.DocumentStart
            ElseIf currentParaIndex >= paragraphs.Length - 1 Then
                context.PositionType = CursorPositionType.DocumentEnd
            Else
                context.PositionType = CursorPositionType.DocumentMiddle
            End If

            Return context
        Catch ex As Exception
            Debug.WriteLine($"GetTextSelectionContext 出错: {ex.Message}")
            Return GetSlideContext(context)
        End Try
    End Function

    ''' <summary>
    ''' 获取形状选择的上下文
    ''' </summary>
    Private Function GetShapeSelectionContext(sel As Selection, context As ContinuationContext) As ContinuationContext
        Try
            _currentShape = sel.ShapeRange(1)

            ' 获取选中形状的文本
            If _currentShape.HasTextFrame AndAlso _currentShape.TextFrame.HasText Then
                context.CurrentParagraphText = _currentShape.TextFrame.TextRange.Text.Trim()
            End If

            ' 获取同一幻灯片上其他形状的文本作为上下文
            Dim beforeBuilder As New StringBuilder()
            Dim afterBuilder As New StringBuilder()
            Dim foundCurrent = False

            For i = 1 To _currentSlide.Shapes.Count
                Dim shape = _currentSlide.Shapes(i)
                If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                    Dim text = shape.TextFrame.TextRange.Text.Trim()
                    If Not String.IsNullOrWhiteSpace(text) Then
                        If shape.Id = _currentShape.Id Then
                            foundCurrent = True
                        ElseIf Not foundCurrent Then
                            beforeBuilder.AppendLine(text)
                        Else
                            afterBuilder.AppendLine(text)
                        End If
                    End If
                End If
            Next

            context.ContextBefore = beforeBuilder.ToString().TrimEnd()
            context.ContextAfter = afterBuilder.ToString().TrimEnd()
            context.PositionType = CursorPositionType.DocumentMiddle

            Return context
        Catch ex As Exception
            Debug.WriteLine($"GetShapeSelectionContext 出错: {ex.Message}")
            Return GetSlideContext(context)
        End Try
    End Function

    ''' <summary>
    ''' 获取整个幻灯片的上下文
    ''' </summary>
    Private Function GetSlideContext(context As ContinuationContext) As ContinuationContext
        Try
            Dim slideContent As New StringBuilder()

            ' 获取当前幻灯片所有文本
            For i = 1 To _currentSlide.Shapes.Count
                Dim shape = _currentSlide.Shapes(i)
                If shape.HasTextFrame AndAlso shape.TextFrame.HasText Then
                    Dim text = shape.TextFrame.TextRange.Text.Trim()
                    If Not String.IsNullOrWhiteSpace(text) Then
                        slideContent.AppendLine(text)
                    End If
                End If
            Next

            context.ContextBefore = slideContent.ToString().TrimEnd()
            context.PositionType = CursorPositionType.DocumentEnd

            Return context
        Catch ex As Exception
            Debug.WriteLine($"GetSlideContext 出错: {ex.Message}")
            Return context
        End Try
    End Function

    ''' <summary>
    ''' 插入续写内容到PowerPoint
    ''' </summary>
    Public Overrides Sub InsertContinuation(content As String, insertPosition As InsertPosition)
        If String.IsNullOrWhiteSpace(content) Then Return

        Try
            Select Case insertPosition
                Case ShareRibbon.InsertPosition.AtCursor
                    InsertAtCursor(content)
                Case ShareRibbon.InsertPosition.DocumentStart
                    InsertToFirstSlide(content)
                Case ShareRibbon.InsertPosition.DocumentEnd
                    InsertToLastSlide(content)
                Case ShareRibbon.InsertPosition.AfterParagraph
                    InsertAfterParagraph(content)
                Case ShareRibbon.InsertPosition.NewParagraph
                    InsertAsNewTextBox(content)
                Case Else
                    InsertAtCursor(content)
            End Select

            OnContinuationCompleted(New ContinuationResult() With {
                .Content = content,
                .Success = True
            })

        Catch ex As Exception
            Debug.WriteLine($"InsertContinuation 出错: {ex.Message}")
            OnContinuationCompleted(New ContinuationResult() With {
                .Success = False,
                .ErrorMessage = ex.Message
            })
        End Try
    End Sub

    ''' <summary>
    ''' 插入到首页
    ''' </summary>
    Private Sub InsertToFirstSlide(content As String)
        Try
            If _presentation.Slides.Count = 0 Then
                InsertAsNewTextBox(content)
                Return
            End If
            
            Dim firstSlide = _presentation.Slides(1)
            InsertTextBoxToSlide(firstSlide, content)
        Catch ex As Exception
            Debug.WriteLine($"InsertToFirstSlide 出错: {ex.Message}")
            InsertAsNewTextBox(content)
        End Try
    End Sub

    ''' <summary>
    ''' 插入到末页
    ''' </summary>
    Private Sub InsertToLastSlide(content As String)
        Try
            If _presentation.Slides.Count = 0 Then
                InsertAsNewTextBox(content)
                Return
            End If
            
            Dim lastSlide = _presentation.Slides(_presentation.Slides.Count)
            InsertTextBoxToSlide(lastSlide, content)
        Catch ex As Exception
            Debug.WriteLine($"InsertToLastSlide 出错: {ex.Message}")
            InsertAsNewTextBox(content)
        End Try
    End Sub

    ''' <summary>
    ''' 在指定幻灯片上插入文本框
    ''' </summary>
    Private Sub InsertTextBoxToSlide(slide As Slide, content As String)
        Try
            Dim slideWidth = _presentation.PageSetup.SlideWidth
            Dim slideHeight = _presentation.PageSetup.SlideHeight

            Dim left As Single = 50
            Dim top As Single = slideHeight - 150
            Dim width As Single = slideWidth - 100
            Dim height As Single = 100

            Dim textBox = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, width, height)

            textBox.TextFrame.TextRange.Text = content
            textBox.TextFrame.TextRange.Font.Size = 14
        Catch ex As Exception
            Debug.WriteLine($"InsertTextBoxToSlide 出错: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 在光标位置插入
    ''' </summary>
    Private Sub InsertAtCursor(content As String)
        Try
            If _currentTextRange IsNot Nothing Then
                ' 如果有当前文本范围，在其后插入
                _currentTextRange.InsertAfter(content)
            ElseIf _currentShape IsNot Nothing AndAlso _currentShape.HasTextFrame Then
                ' 在形状文本末尾追加
                Dim textRange = _currentShape.TextFrame.TextRange
                textRange.InsertAfter(vbCr & content)
            Else
                ' 创建新文本框
                InsertAsNewTextBox(content)
            End If
        Catch ex As Exception
            Debug.WriteLine($"InsertAtCursor 出错: {ex.Message}")
            InsertAsNewTextBox(content)
        End Try
    End Sub

    ''' <summary>
    ''' 在段落后插入
    ''' </summary>
    Private Sub InsertAfterParagraph(content As String)
        Try
            If _currentShape IsNot Nothing AndAlso _currentShape.HasTextFrame Then
                Dim textRange = _currentShape.TextFrame.TextRange
                textRange.InsertAfter(vbCr & content)
            Else
                InsertAsNewTextBox(content)
            End If
        Catch ex As Exception
            Debug.WriteLine($"InsertAfterParagraph 出错: {ex.Message}")
            InsertAsNewTextBox(content)
        End Try
    End Sub

    ''' <summary>
    ''' 创建新文本框插入
    ''' </summary>
    Private Sub InsertAsNewTextBox(content As String)
        Try
            If _currentSlide Is Nothing Then Return

            ' 在幻灯片底部创建新文本框
            Dim slideWidth = _presentation.PageSetup.SlideWidth
            Dim slideHeight = _presentation.PageSetup.SlideHeight

            Dim left As Single = 50
            Dim top As Single = slideHeight - 150
            Dim width As Single = slideWidth - 100
            Dim height As Single = 100

            Dim textBox = _currentSlide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, width, height)

            textBox.TextFrame.TextRange.Text = content
            textBox.TextFrame.TextRange.Font.Size = 14

        Catch ex As Exception
            Debug.WriteLine($"InsertAsNewTextBox 出错: {ex.Message}")
            Throw
        End Try
    End Sub
End Class
