Imports System.Diagnostics
Imports System.Text
Imports Microsoft.Office.Interop.Word
Imports ShareRibbon
Imports Word = Microsoft.Office.Interop.Word

''' <summary>
''' Word文档续写服务 - 获取光标上下文并插入续写内容
''' </summary>
Public Class WordContinuationService
    Inherits ContinuationService

    Private _wordApp As Word.Application
    Private _document As Document
    Private _cursorRange As Range

    Public Sub New(wordApp As Word.Application)
        _wordApp = wordApp
        _document = wordApp.ActiveDocument
    End Sub

    ''' <summary>
    ''' 检查是否可以进行续写
    ''' </summary>
    Public Overrides Function CanContinue() As Boolean
        If _wordApp Is Nothing OrElse _document Is Nothing Then
            Return False
        End If

        Try
            Dim sel = _wordApp.Selection
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
            Dim sel = _wordApp.Selection
            If sel Is Nothing Then Return Nothing

            ' 保存当前光标位置
            _cursorRange = sel.Range.Duplicate
            Dim cursorPos = sel.Start

            Dim context As New ContinuationContext()
            context.CursorPosition = cursorPos
            context.DocumentPath = If(_document.Path, "")

            ' 获取当前段落
            Dim currentPara As Paragraph = Nothing
            Dim currentParaIndex As Integer = 0

            Try
                currentPara = sel.Range.Paragraphs(1)
                ' 找到当前段落在文档中的索引
                For i = 1 To _document.Paragraphs.Count
                    If _document.Paragraphs(i).Range.Start = currentPara.Range.Start Then
                        currentParaIndex = i
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine($"获取当前段落失败: {ex.Message}")
            End Try

            ' 获取当前段落文本
            If currentPara IsNot Nothing Then
                Dim paraText = currentPara.Range.Text
                paraText = paraText.TrimEnd(ChrW(13), ChrW(10), ChrW(7))
                context.CurrentParagraphText = paraText

                ' 计算光标在段落中的偏移
                context.CursorOffsetInParagraph = cursorPos - currentPara.Range.Start
            End If

            ' 判断光标位置类型
            Dim totalParas = _document.Paragraphs.Count
            If currentParaIndex <= 1 AndAlso context.CursorOffsetInParagraph <= 1 Then
                context.PositionType = CursorPositionType.DocumentStart
            ElseIf currentParaIndex >= totalParas AndAlso
                   context.CursorOffsetInParagraph >= Len(context.CurrentParagraphText) - 1 Then
                context.PositionType = CursorPositionType.DocumentEnd
            Else
                context.PositionType = CursorPositionType.DocumentMiddle
            End If

            ' 获取光标前的段落
            Dim beforeBuilder As New StringBuilder()
            Dim startIndex = Math.Max(1, currentParaIndex - paragraphsBefore)
            For i = startIndex To currentParaIndex - 1
                If i >= 1 AndAlso i <= totalParas Then
                    Dim paraText = _document.Paragraphs(i).Range.Text
                    paraText = paraText.TrimEnd(ChrW(13), ChrW(10), ChrW(7))
                    If Not String.IsNullOrWhiteSpace(paraText) Then
                        beforeBuilder.AppendLine(paraText)
                    End If
                End If
            Next
            context.ContextBefore = beforeBuilder.ToString().TrimEnd()

            ' 获取光标后的段落
            Dim afterBuilder As New StringBuilder()
            Dim endIndex = Math.Min(totalParas, currentParaIndex + paragraphsAfter)
            For i = currentParaIndex + 1 To endIndex
                If i >= 1 AndAlso i <= totalParas Then
                    Dim paraText = _document.Paragraphs(i).Range.Text
                    paraText = paraText.TrimEnd(ChrW(13), ChrW(10), ChrW(7))
                    If Not String.IsNullOrWhiteSpace(paraText) Then
                        afterBuilder.AppendLine(paraText)
                    End If
                End If
            Next
            context.ContextAfter = afterBuilder.ToString().TrimEnd()

            Return context
        Catch ex As Exception
            Debug.WriteLine($"GetCursorContext 出错: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 插入续写内容到文档
    ''' </summary>
    Public Overrides Sub InsertContinuation(content As String, insertPosition As InsertPosition)
        If String.IsNullOrWhiteSpace(content) Then Return

        Try
            _document.Application.ScreenUpdating = False
            _document.Application.UndoRecord.StartCustomRecord("AI续写")

            Dim sel = _wordApp.Selection
            Dim insertRange As Range

            Select Case insertPosition
                Case ShareRibbon.InsertPosition.AtCursor
                    ' 在光标位置插入
                    If _cursorRange IsNot Nothing Then
                        insertRange = _cursorRange.Duplicate
                    Else
                        insertRange = sel.Range.Duplicate
                    End If
                    insertRange.Collapse(WdCollapseDirection.wdCollapseEnd)
                    insertRange.Text = content

                Case ShareRibbon.InsertPosition.DocumentStart
                    ' 在文档开头插入
                    insertRange = _document.Range(0, 0)
                    insertRange.Text = content & vbCr

                Case ShareRibbon.InsertPosition.DocumentEnd
                    ' 在文档结尾插入
                    insertRange = _document.Range(_document.Content.End - 1, _document.Content.End - 1)
                    insertRange.Text = vbCr & content

                Case ShareRibbon.InsertPosition.AfterParagraph
                    ' 在当前段落后插入
                    Dim currentPara = sel.Range.Paragraphs(1)
                    insertRange = currentPara.Range.Duplicate
                    insertRange.Collapse(WdCollapseDirection.wdCollapseEnd)
                    insertRange.Text = vbCr & content

                Case ShareRibbon.InsertPosition.NewParagraph
                    ' 新建段落插入
                    insertRange = sel.Range.Duplicate
                    insertRange.Collapse(WdCollapseDirection.wdCollapseEnd)
                    insertRange.Text = vbCr & vbCr & content

                Case Else
                    ' 默认在光标位置插入
                    insertRange = sel.Range.Duplicate
                    insertRange.Collapse(WdCollapseDirection.wdCollapseEnd)
                    insertRange.Text = content
            End Select

            ' 将光标移动到插入内容的末尾
            sel.SetRange(insertRange.End, insertRange.End)

            _document.Application.UndoRecord.EndCustomRecord()

            ' 触发完成事件
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
        Finally
            _document.Application.ScreenUpdating = True
        End Try
    End Sub
End Class
