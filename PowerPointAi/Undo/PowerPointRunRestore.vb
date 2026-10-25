Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

Namespace PowerPointUndo

    ''' <summary>
    ''' PowerPoint 没有"一条命令 = 一个撤销记录"的事务边界，撤销本轮只能按还原快照重放：
    ''' 执行前抓下目标页每个文本形状的文字和备注，撤销时逐格写回，再用同一套签名核对是否真回到执行前。
    ''' 这条路只覆盖文字与备注：形状增删、位置、字体、填充、动画任一变化都不产生快照，
    ''' 因为写回 Text 还原不了它们，硬承诺"撤销本轮"等于骗用户。
    ''' 采集失败、体积超限、形状对不上时一律返回 Nothing 或拒绝重放，让 UI 不给"撤销本轮"按钮。
    ''' </summary>
    Public Class PowerPointRunRestore

        Public Const RestoreKind As String = "powerpoint_slide_text"

        ''' <summary>单页文本采集上限，只为挡住超长文稿；最终能否落库仍由共享合同的体积闸门决定。</summary>
        Private Const MaxTextChars As Integer = 20000

        ''' <summary>抓取目标页的执行前文本状态；任何一处读不到就整体作废，不留半份快照。</summary>
        Public Shared Function Capture(app As Object, slideIndex As Integer) As JObject
            If slideIndex <= 0 Then Return Nothing
            Dim presentation As Object = Nothing
            Dim slides As Object = Nothing
            Dim slide As Object = Nothing
            Try
                presentation = app.ActivePresentation
                If presentation Is Nothing Then Return Nothing
                slides = presentation.Slides
                Dim slideCount = CInt(slides.Count)
                If slideIndex > slideCount Then Return Nothing
                slide = slides.Item(slideIndex)

                Dim state = CollectTextState(slide)
                If state Is Nothing Then Return Nothing
                state("presentation") = If(TryCast(presentation.Name, String), "")
                state("slideIndex") = slideIndex
                state("slideCount") = slideCount
                If EstimateChars(state) > MaxTextChars Then Return Nothing
                Return state
            Catch ex As Exception
                Debug.WriteLine($"[PowerPoint] 还原快照采集失败: {ex.Message}")
                Return Nothing
            Finally
                ComObjectHelper.ReleaseComObject(slide)
                ComObjectHelper.ReleaseComObject(slides)
                ComObjectHelper.ReleaseComObject(presentation)
            End Try
        End Function

        ''' <summary>
        ''' 组装还原节点：调用方已经确认改动只落在文本/备注通道上（形状数、动画、格式签名都没动）。
        ''' 补上执行后的签名，重放时用它区分"本轮留下的状态"和"用户后来又改的状态"。
        ''' </summary>
        Public Shared Function BuildNode(before As JObject, after As JObject, textOnlyChange As Boolean) As JObject
            If Not textOnlyChange Then Return Nothing
            If before Is Nothing OrElse after Is Nothing Then Return Nothing
            Dim beforeSignature = If(before("signature")?.ToString(), "")
            Dim afterSignature = If(after("signature")?.ToString(), "")
            If beforeSignature.Length = 0 OrElse afterSignature.Length = 0 Then Return Nothing
            If String.Equals(beforeSignature, afterSignature, StringComparison.Ordinal) Then Return Nothing
            before("afterSignature") = afterSignature
            Return ShareRibbon.Agent.AgentRunChangeLog.BuildRestoreNode(RestoreKind, before)
        End Function

        ''' <summary>
        ''' 按快照把目标页文字写回执行前。返回空字符串表示已还原并核对通过；
        ''' 其余任何返回值都是"没有还原"的原因，调用方据此如实告诉用户。
        ''' </summary>
        Public Shared Function Replay(app As Object, state As JObject) As String
            If state Is Nothing Then Return "还原快照缺失，已拒绝自动回退。"
            Dim presentation As Object = Nothing
            Dim slides As Object = Nothing
            Dim slide As Object = Nothing
            Try
                Dim slideIndex = ReadInteger(state, "slideIndex")
                If slideIndex <= 0 Then Return "还原快照没有记录页码，已拒绝自动回退。"

                presentation = app.ActivePresentation
                If presentation Is Nothing Then Return "没有打开的演示文稿，无法还原本轮改动。"
                Dim expectedPresentation = If(state("presentation")?.ToString(), "")
                If expectedPresentation.Length > 0 AndAlso
                   Not String.Equals(If(TryCast(presentation.Name, String), ""), expectedPresentation, StringComparison.Ordinal) Then
                    Return $"当前演示文稿不是本轮改动所在的「{expectedPresentation}」，已拒绝还原。"
                End If

                slides = presentation.Slides
                If slideIndex > CInt(slides.Count) Then
                    Return $"第 {slideIndex} 页已不在演示文稿里（现在只剩 {CInt(slides.Count)} 页），无法按快照还原。"
                End If
                slide = slides.Item(slideIndex)

                If CInt(slide.Shapes.Count) <> ReadInteger(state, "shapeCount") Then
                    Return $"第 {slideIndex} 页的形状数量已与本轮记录不一致，写回文字会落到别的形状上，已拒绝还原。"
                End If

                Dim beforeSignature = If(state("signature")?.ToString(), "")
                Dim expectedSignature = If(state("afterSignature")?.ToString(), "")
                Dim current = ReadSignature(slide)
                If current Is Nothing Then Return "无法读取该页当前内容，已拒绝自动回退。"
                If beforeSignature.Length > 0 AndAlso String.Equals(current, beforeSignature, StringComparison.Ordinal) Then
                    Return ""
                End If
                If expectedSignature.Length > 0 AndAlso
                   Not String.Equals(current, expectedSignature, StringComparison.Ordinal) Then
                    Return $"第 {slideIndex} 页在本轮之后又有其他改动，已拒绝还原以免覆盖你的内容。"
                End If

                Dim bodyError = WriteTextShapes(slide.Shapes, TryCast(state("shapes"), JArray), "文本框")
                If bodyError.Length > 0 Then Return bodyError
                Dim notesError = WriteNotesText(slide, TryCast(state("notesShapes"), JArray))
                If notesError.Length > 0 Then Return notesError

                Dim restored = ReadSignature(slide)
                If restored Is Nothing Then Return "写回后无法核对页面内容，请对照变更清单人工确认。"
                If beforeSignature.Length > 0 AndAlso
                   Not String.Equals(restored, beforeSignature, StringComparison.Ordinal) Then
                    Return $"已写回第 {slideIndex} 页，但与执行前的记录仍不一致，请对照变更清单人工核对。"
                End If
                Return ""
            Catch ex As Exception
                Debug.WriteLine($"[PowerPoint] 还原快照重放失败: {ex.Message}")
                Return $"还原失败: {AppLogger.Redact(ex.Message)}"
            Finally
                ComObjectHelper.ReleaseComObject(slide)
                ComObjectHelper.ReleaseComObject(slides)
                ComObjectHelper.ReleaseComObject(presentation)
            End Try
        End Function

        ''' <summary>采集正文与备注两类文本形状；和签名的口径必须一致，否则核对永远对不上。</summary>
        Private Shared Function CollectTextState(slide As Object) As JObject
            Dim shapes = CaptureTextShapes(slide.Shapes)
            If shapes Is Nothing Then Return Nothing
            Dim notesPage As Object = Nothing
            Dim notes As JArray
            Try
                notesPage = slide.NotesPage
                notes = CaptureTextShapes(notesPage.Shapes)
            Catch
                notes = Nothing
            Finally
                ComObjectHelper.ReleaseComObject(notesPage)
            End Try
            If notes Is Nothing Then Return Nothing

            Dim state As New JObject From {
                {"shapeCount", CInt(slide.Shapes.Count)},
                {"shapes", shapes},
                {"notesShapes", notes}
            }
            state("signature") = Signature(shapes, notes)
            Return state
        End Function

        Private Shared Function ReadSignature(slide As Object) As String
            Dim state = CollectTextState(slide)
            If state Is Nothing Then Return Nothing
            Return If(state("signature")?.ToString(), "")
        End Function

        Private Shared Function CaptureTextShapes(container As Object) As JArray
            Dim result As New JArray()
            Try
                Dim index As Integer = 0
                For Each shape As Object In container
                    index += 1
                    If Not CType(shape.HasTextFrame, Boolean) Then Continue For
                    Dim text As String = ""
                    Dim range As Object = Nothing
                    Try
                        range = shape.TextFrame.TextRange
                        If range IsNot Nothing Then text = If(TryCast(range.Text, String), "")
                    Finally
                        ComObjectHelper.ReleaseComObject(range)
                    End Try
                    result.Add(New JObject From {
                        {"index", index},
                        {"name", If(TryCast(shape.Name, String), "")},
                        {"text", text}
                    })
                Next
                Return result
            Catch ex As Exception
                Debug.WriteLine($"[PowerPoint] 文本形状采集失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>备注页只在执行前确实有文本形状时才回写，避免凭空造出一个备注占位符。</summary>
        Private Shared Function WriteNotesText(slide As Object, recorded As JArray) As String
            If recorded Is Nothing Then Return "备注还原快照不完整，已拒绝自动回退。"
            If recorded.Count = 0 Then Return ""
            Dim notesPage As Object = Nothing
            Try
                notesPage = slide.NotesPage
                Return WriteTextShapes(notesPage.Shapes, recorded, "备注")
            Catch ex As Exception
                Return $"备注写回失败: {AppLogger.Redact(ex.Message)}"
            Finally
                ComObjectHelper.ReleaseComObject(notesPage)
            End Try
        End Function

        ''' <summary>
        ''' 逐格写回：先用序号+形状名确认还是同一个形状，不是就整体停下——按位置写回一页文字，
        ''' 撤错形状比撤不动严重得多。文字已经一致的形状跳过，免得白白重刷一遍格式。
        ''' </summary>
        Private Shared Function WriteTextShapes(container As Object, recorded As JArray, label As String) As String
            If recorded Is Nothing Then Return $"{label}还原快照不完整，已拒绝自动回退。"
            For Each entry In recorded.OfType(Of JObject)()
                Dim index = ReadInteger(entry, "index")
                Dim name = If(entry("name")?.ToString(), "")
                If index <= 0 Then Return $"{label}还原快照缺少形状序号，已拒绝自动回退。"
                Dim expected = If(entry("text")?.ToString(), "")
                Dim shape As Object = Nothing
                Dim range As Object = Nothing
                Try
                    shape = container.Item(index)
                    If Not String.Equals(If(TryCast(shape.Name, String), ""), name, StringComparison.Ordinal) Then
                        Return $"第 {index} 个{label}已不是你记录的那个形状，为避免改错对象已拒绝还原。"
                    End If
                    If Not CType(shape.HasTextFrame, Boolean) Then Return $"第 {index} 个{label}已经没有文本框，已拒绝还原。"
                    range = shape.TextFrame.TextRange
                    If range IsNot Nothing AndAlso
                       Not String.Equals(If(TryCast(range.Text, String), ""), expected, StringComparison.Ordinal) Then
                        range.Text = expected
                    End If
                Catch ex As Exception
                    Return $"{label}写回失败: {AppLogger.Redact(ex.Message)}"
                Finally
                    ComObjectHelper.ReleaseComObject(range)
                    ComObjectHelper.ReleaseComObject(shape)
                End Try
            Next
            Return ""
        End Function

        Private Shared Function Signature(shapes As JArray, notes As JArray) As String
            Dim text = shapes.ToString(Formatting.None) & "|" & notes.ToString(Formatting.None)
            Dim bytes = System.Text.Encoding.UTF8.GetBytes(text)
            Using sha = System.Security.Cryptography.SHA256.Create()
                Return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Function EstimateChars(state As JObject) As Integer
            Return If(state("shapes")?.ToString(Formatting.None).Length, 0) +
                   If(state("notesShapes")?.ToString(Formatting.None).Length, 0)
        End Function

        Private Shared Function ReadInteger(obj As JObject, name As String) As Integer
            Dim token = obj?(name)
            If token Is Nothing Then Return 0
            Try
                Return token.Value(Of Integer)()
            Catch
                Return 0
            End Try
        End Function

    End Class

End Namespace
