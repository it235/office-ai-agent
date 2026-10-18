' WordAi/Services/ParagraphService.vb
' 段落服务 - 提供原子级别的段落查询和操作能力

Imports Word = Microsoft.Office.Interop.Word
Imports System.Diagnostics
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

Namespace Services

    ''' <summary>
    ''' 段落服务 - Harness 架构的原子工具
    ''' </summary>
    Public Class ParagraphService

        Private ReadOnly _app As Word.Application

        Public Sub New(app As Word.Application)
            _app = app
        End Sub

        ''' <summary>
        ''' 列出段落 - 返回段落摘要列表
        ''' </summary>
        Public Function ListParagraphs(Optional maxCount As Integer = 50) As JArray
            Try
                Dim doc As Word.Document = _app.ActiveDocument
                If doc Is Nothing Then
                    Return New JArray()
                End If

                Dim result As New JArray()
                Dim count As Integer = 0

                For Each para As Word.Paragraph In doc.Paragraphs
                    If count >= maxCount Then Exit For

                    Try
                        Dim text As String = para.Range.Text.Trim()
                        If String.IsNullOrEmpty(text) Then Continue For

                        ' 截取前30字符
                        Dim preview As String = text
                        If preview.Length > 30 Then
                            preview = preview.Substring(0, 30) & "..."
                        End If

                        ' 获取样式名称
                        Dim styleName As String = "正文"
                        Try
                            styleName = para.Style.NameLocal
                        Catch
                        End Try

                        Dim paraInfo As New JObject()
                        paraInfo("index") = count + 1
                        paraInfo("text") = preview
                        paraInfo("style") = styleName
                        paraInfo("length") = text.Length

                        result.Add(paraInfo)
                        count += 1

                    Catch ex As Exception
                        Debug.WriteLine($"[ParagraphService] 处理段落失败: {ex.Message}")
                    End Try
                Next

                Debug.WriteLine($"[ParagraphService] 列出 {result.Count} 个段落")
                Return result

            Catch ex As Exception
                Debug.WriteLine($"[ParagraphService] ListParagraphs 失败: {ex.Message}")
                Return New JArray()
            End Try
        End Function

        ''' <summary>
        ''' 获取段落详细信息
        ''' </summary>
        Public Function GetParagraphInfo(paragraphIndex As Integer) As JObject
            Try
                Dim doc As Word.Document = _app.ActiveDocument
                If doc Is Nothing OrElse paragraphIndex < 1 OrElse paragraphIndex > doc.Paragraphs.Count Then
                    Return Nothing
                End If

                Dim para As Word.Paragraph = doc.Paragraphs(paragraphIndex)
                Dim result As New JObject()

                ' 基本信息
                result("index") = paragraphIndex
                result("text") = para.Range.Text.Trim()

                ' 样式
                Try
                    result("style") = para.Style.NameLocal
                Catch
                    result("style") = "正文"
                End Try

                ' 字体信息
                Try
                    Dim font As Word.Font = para.Range.Font
                    result("fontSize") = font.Size
                    result("fontName") = font.Name
                    result("bold") = (font.Bold = -1)
                    result("italic") = (font.Italic = -1)
                    result("underline") = (font.Underline <> Word.WdUnderline.wdUnderlineNone)

                    ' 颜色
                    Dim colorName As String = "black"
                    Select Case font.Color
                        Case Word.WdColor.wdColorRed
                            colorName = "red"
                        Case Word.WdColor.wdColorBlue
                            colorName = "blue"
                        Case Word.WdColor.wdColorGreen
                            colorName = "green"
                    End Select
                    result("color") = colorName

                Catch fontEx As Exception
                    Debug.WriteLine($"[ParagraphService] 获取字体信息失败: {fontEx.Message}")
                End Try

                ' 段落格式
                Try
                    Dim alignment As String = "left"
                    Select Case para.Alignment
                        Case Word.WdParagraphAlignment.wdAlignParagraphCenter
                            alignment = "center"
                        Case Word.WdParagraphAlignment.wdAlignParagraphRight
                            alignment = "right"
                        Case Word.WdParagraphAlignment.wdAlignParagraphJustify
                            alignment = "justify"
                    End Select
                    result("alignment") = alignment

                Catch alignEx As Exception
                    Debug.WriteLine($"[ParagraphService] 获取对齐方式失败: {alignEx.Message}")
                End Try

                Debug.WriteLine($"[ParagraphService] 获取段落 {paragraphIndex} 详情")
                Return result

            Catch ex As Exception
                Debug.WriteLine($"[ParagraphService] GetParagraphInfo 失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 设置段落格式
        ''' </summary>
        Public Function SetParagraphFormat(paragraphIndex As Integer,
                                           params As JObject,
                                           Optional ByRef failure As JObject = Nothing) As Boolean
            failure = Nothing
            Try
                Dim doc As Word.Document = _app.ActiveDocument
                If doc Is Nothing Then
                    failure = BuildFailure(ExceptionClassifier.CodeDocMissing, "当前没有可操作的 Word 文档")
                    Return False
                End If
                If paragraphIndex < 1 OrElse paragraphIndex > doc.Paragraphs.Count Then
                    ' 执行器只有把"第几段不存在、现在一共几段"说出来，Loop 才有东西可修；
                    ' 光回 Boolean 会在 ToolRegistry 塌成 UNKNOWN。
                    failure = BuildFailure(ExceptionClassifier.CodeNotFound,
                                           $"第 {paragraphIndex} 段不存在：文档当前只有 {doc.Paragraphs.Count} 段")
                    Return False
                End If

                Dim para As Word.Paragraph = doc.Paragraphs(paragraphIndex)
                Dim modified As Boolean = False

                ' 字号
                If params("fontSize") IsNot Nothing Then
                    para.Range.Font.Size = params("fontSize").Value(Of Single)()
                    modified = True
                    Debug.WriteLine($"[ParagraphService] 设置字号: {params("fontSize")}")
                End If

                ' 字体
                If params("fontName") IsNot Nothing Then
                    para.Range.Font.Name = params("fontName").ToString()
                    modified = True
                    Debug.WriteLine($"[ParagraphService] 设置字体: {params("fontName")}")
                End If

                ' 加粗
                If params("bold") IsNot Nothing Then
                    para.Range.Font.Bold = If(params("bold").Value(Of Boolean)(), -1, 0)
                    modified = True
                    Debug.WriteLine($"[ParagraphService] 设置加粗: {params("bold")}")
                End If

                ' 斜体
                If params("italic") IsNot Nothing Then
                    para.Range.Font.Italic = If(params("italic").Value(Of Boolean)(), -1, 0)
                    modified = True
                    Debug.WriteLine($"[ParagraphService] 设置斜体: {params("italic")}")
                End If

                ' 颜色
                If params("color") IsNot Nothing Then
                    Dim colorStr As String = params("color").ToString().ToLower()
                    ' 未识别的颜色值不能算"已应用"，否则执行器谎报成功、观察器只能说"未检测到变化"
                    Dim appliedColor As Boolean = True
                    Select Case colorStr
                        Case "red"
                            para.Range.Font.Color = Word.WdColor.wdColorRed
                        Case "blue"
                            para.Range.Font.Color = Word.WdColor.wdColorBlue
                        Case "green"
                            para.Range.Font.Color = Word.WdColor.wdColorGreen
                        Case "black"
                            para.Range.Font.Color = Word.WdColor.wdColorBlack
                        Case Else
                            appliedColor = False
                    End Select
                    If appliedColor Then
                        modified = True
                        Debug.WriteLine($"[ParagraphService] 设置颜色: {colorStr}")
                    End If
                End If

                ' 对齐
                If params("alignment") IsNot Nothing Then
                    Dim alignStr As String = params("alignment").ToString().ToLower()
                    Dim appliedAlign As Boolean = True
                    Select Case alignStr
                        Case "left"
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                        Case "center"
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        Case "right"
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
                        Case "justify"
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
                        Case Else
                            appliedAlign = False
                    End Select
                    If appliedAlign Then
                        modified = True
                        Debug.WriteLine($"[ParagraphService] 设置对齐: {alignStr}")
                    End If
                End If

                ' 首行缩进按“字符”交给 Word 原生字符单位属性，避免各处按字号/厘米换算的第三种口径
                If params("firstLineIndent") IsNot Nothing Then
                    para.Format.CharacterUnitFirstLineIndent = params("firstLineIndent").Value(Of Single)()
                    modified = True
                    Debug.WriteLine($"[ParagraphService] 设置首行缩进: {params("firstLineIndent")} 字符")
                End If

                If params("beforeSpacing") IsNot Nothing Then
                    para.Format.SpaceBefore = params("beforeSpacing").Value(Of Single)()
                    modified = True
                    Debug.WriteLine($"[ParagraphService] 设置段前距: {params("beforeSpacing")} 磅")
                End If

                If params("afterSpacing") IsNot Nothing Then
                    para.Format.SpaceAfter = params("afterSpacing").Value(Of Single)()
                    modified = True
                    Debug.WriteLine($"[ParagraphService] 设置段后距: {params("afterSpacing")} 磅")
                End If

                If modified Then
                    Debug.WriteLine($"[ParagraphService] 段落 {paragraphIndex} 格式已更新")
                Else
                    failure = BuildFailure(ExceptionClassifier.CodeArgument,
                                           $"第 {paragraphIndex} 段没有可应用的格式属性；收到 {DescribeReceivedKeys(params)}，" &
                                           "支持 fontSize/fontName/bold/italic/color(red|blue|green|black)/alignment(left|center|right|justify)/firstLineIndent/beforeSpacing/afterSpacing")
                End If

                Return modified

            Catch ex As Exception
                Debug.WriteLine($"[ParagraphService] SetParagraphFormat 失败: {ex.Message}")
                failure = BuildFailure(ExceptionClassifier.CodeCom, $"设置第 {paragraphIndex} 段格式时 Word 报错：{ex.Message}")
                Return False
            End Try
        End Function

        Private Shared Function BuildFailure(code As String, message As String) As JObject
            Return New JObject From {{"code", code}, {"message", message}}
        End Function

        Private Shared Function DescribeReceivedKeys(params As JObject) As String
            If params Is Nothing OrElse params.Count = 0 Then Return "空参数"
            Dim names As New List(Of String)()
            For Each p In params.Properties()
                names.Add(p.Name)
            Next
            Return String.Join("/", names)
        End Function

    End Class

End Namespace
