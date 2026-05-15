' ShareRibbon\Services\Reformat\SemanticRenderingEngine.vb
' 确定性渲染引擎 - 根据语义标签应用Word格式

Imports Newtonsoft.Json.Linq

''' <summary>
''' 语义渲染引擎 - 根据语义标签确定性地应用Word格式
''' 核心类：接收AI标注结果 + SemanticStyleMapping，渲染到Word段落
''' </summary>
Public Class SemanticRenderingEngine

    ''' <summary>段落格式快照（用于撤销排版）</summary>
    Public Class ParaFormatSnapshot
        Public Property FontNameFarEast As String
        Public Property FontNameEN As String
        Public Property FontSize As Single
        Public Property Bold As Integer
        Public Property Italic As Integer
        Public Property Underline As Integer
        Public Property FontColor As Object
        Public Property Alignment As Object
        Public Property FirstLineIndent As Single
        Public Property CharacterUnitFirstLineIndent As Single
        Public Property LineSpacingRule As Object
        Public Property LineSpacing As Single
        Public Property SpaceBefore As Single
        Public Property SpaceAfter As Single
    End Class

    ''' <summary>排版格式快照集合（用于撤销排版）</summary>
    Public Class ReformatSnapshot
        Public Property ParagraphSnapshots As New List(Of ParaFormatSnapshot)
        Public Property CapturedAt As DateTime = DateTime.Now
        Public Property ParagraphCount As Integer
    End Class

    ''' <summary>
    ''' 捕获段落格式快照（在应用排版前调用）
    ''' </summary>
    Public Shared Function CaptureFormatSnapshot(paragraphs As List(Of Object), paragraphTypes As List(Of String)) As ReformatSnapshot
        Dim snapshot As New ReformatSnapshot()
        snapshot.ParagraphCount = paragraphs.Count

        For i As Integer = 0 To paragraphs.Count - 1
            If paragraphTypes IsNot Nothing AndAlso i < paragraphTypes.Count AndAlso paragraphTypes(i) <> "text" Then
                snapshot.ParagraphSnapshots.Add(Nothing)
                Continue For
            End If

            Try
                Dim para = paragraphs(i)
                Dim rng = para.Range
                Dim item As New ParaFormatSnapshot() With {
                    .FontNameFarEast = SafeGetString(rng.Font.NameFarEast),
                    .FontNameEN = SafeGetString(rng.Font.Name),
                    .FontSize = SafeGetSingle(rng.Font.Size),
                    .Bold = SafeGetInt(rng.Font.Bold),
                    .Italic = SafeGetInt(rng.Font.Italic),
                    .Underline = SafeGetInt(rng.Font.Underline),
                    .FontColor = SafeGetObject(rng.Font.Color),
                    .Alignment = SafeGetObject(rng.ParagraphFormat.Alignment),
                    .FirstLineIndent = SafeGetSingle(rng.ParagraphFormat.FirstLineIndent),
                    .CharacterUnitFirstLineIndent = SafeGetSingle(rng.ParagraphFormat.CharacterUnitFirstLineIndent),
                    .LineSpacingRule = SafeGetObject(rng.ParagraphFormat.LineSpacingRule),
                    .LineSpacing = SafeGetSingle(rng.ParagraphFormat.LineSpacing),
                    .SpaceBefore = SafeGetSingle(rng.ParagraphFormat.SpaceBefore),
                    .SpaceAfter = SafeGetSingle(rng.ParagraphFormat.SpaceAfter)
                }
                snapshot.ParagraphSnapshots.Add(item)
            Catch ex As Exception
                snapshot.ParagraphSnapshots.Add(Nothing)
                Debug.WriteLine($"CaptureFormatSnapshot 段落{i}失败: {ex.Message}")
            End Try
        Next

        Return snapshot
    End Function

    ''' <summary>
    ''' 从快照恢复段落格式（撤销排版）
    ''' </summary>
    Public Shared Function RestoreFormatSnapshot(snapshot As ReformatSnapshot, paragraphs As List(Of Object), paragraphTypes As List(Of String)) As Integer
        Dim restoredCount As Integer = 0
        If snapshot Is Nothing OrElse paragraphs Is Nothing Then Return 0

        Dim count = Math.Min(snapshot.ParagraphSnapshots.Count, paragraphs.Count)
        For i As Integer = 0 To count - 1
            Dim snap = snapshot.ParagraphSnapshots(i)
            If snap Is Nothing Then Continue For
            If paragraphTypes IsNot Nothing AndAlso i < paragraphTypes.Count AndAlso paragraphTypes(i) <> "text" Then Continue For

            Try
                Dim para = paragraphs(i)
                Dim rng = para.Range

                ' 恢复字体
                If Not String.IsNullOrEmpty(snap.FontNameFarEast) Then rng.Font.NameFarEast = snap.FontNameFarEast
                If Not String.IsNullOrEmpty(snap.FontNameEN) Then rng.Font.Name = snap.FontNameEN
                If snap.FontSize > 0 Then rng.Font.Size = snap.FontSize
                rng.Font.Bold = snap.Bold
                rng.Font.Italic = snap.Italic
                rng.Font.Underline = snap.Underline
                If snap.FontColor IsNot Nothing Then
                    Try : rng.Font.Color = snap.FontColor : Catch : End Try
                End If

                ' 恢复段落格式
                If snap.Alignment IsNot Nothing Then
                    Try : rng.ParagraphFormat.Alignment = snap.Alignment : Catch : End Try
                End If
                If snap.CharacterUnitFirstLineIndent > 0 Then
                    rng.ParagraphFormat.CharacterUnitFirstLineIndent = snap.CharacterUnitFirstLineIndent
                ElseIf snap.FirstLineIndent > 0 Then
                    rng.ParagraphFormat.FirstLineIndent = snap.FirstLineIndent
                End If
                If snap.LineSpacingRule IsNot Nothing Then
                    rng.ParagraphFormat.LineSpacingRule = snap.LineSpacingRule
                End If
                If snap.LineSpacing > 0 Then
                    rng.ParagraphFormat.LineSpacing = snap.LineSpacing
                End If
                If snap.SpaceBefore > 0 Then rng.ParagraphFormat.SpaceBefore = snap.SpaceBefore
                If snap.SpaceAfter > 0 Then rng.ParagraphFormat.SpaceAfter = snap.SpaceAfter

                restoredCount += 1
            Catch ex As Exception
                Debug.WriteLine($"RestoreFormatSnapshot 段落{i}失败: {ex.Message}")
            End Try
        Next

        Return restoredCount
    End Function

    Private Shared Function SafeGetString(val As Object) As String
        Try : Return If(val IsNot Nothing, val.ToString(), "") : Catch : Return "" : End Try
    End Function
    Private Shared Function SafeGetSingle(val As Object) As Single
        Try : Return If(val IsNot Nothing, CSng(val), 0F) : Catch : Return 0F : End Try
    End Function
    Private Shared Function SafeGetInt(val As Object) As Integer
        Try : Return If(val IsNot Nothing, CInt(val), 0) : Catch : Return 0 : End Try
    End Function
    Private Shared Function SafeGetObject(val As Object) As Object
        Try : Return val : Catch : Return Nothing : End Try
    End Function

    ''' <summary>渲染结果统计</summary>
    Public Class RenderResult
        Public Property AppliedCount As Integer = 0
        Public Property SkippedCount As Integer = 0
        Public Property TagUsage As New Dictionary(Of String, Integer)()
        Public Property Errors As New List(Of String)()
        ''' <summary>生成的DSL指令列表（含Rollback原始值，用于日志/重放/逐条撤销）</summary>
        Public Property GeneratedInstructions As List(Of Instruction)

        ''' <summary>转换为JSON（用于推送前端）</summary>
        Public Function ToJson() As JObject
            Dim result As New JObject()
            result("appliedCount") = AppliedCount
            result("skippedCount") = SkippedCount
            Dim tagsObj As New JObject()
            For Each kvp In TagUsage
                tagsObj(kvp.Key) = kvp.Value
            Next
            result("tags") = tagsObj
            Return result
        End Function
    End Class

    ''' <summary>
    ''' 应用语义排版到Word段落
    ''' </summary>
    ''' <param name="taggedParagraphs">AI标注结果: List of (paraIndex, tagId)</param>
    ''' <param name="mapping">语义样式映射</param>
    ''' <param name="wordParagraphs">Word段落对象列表</param>
    ''' <param name="paragraphTypes">段落类型列表（text/image/table/formula）</param>
    ''' <param name="wordApp">Word Application对象（用于页面设置）</param>
    ''' <param name="onProgress">进度回调 (当前序号1-based, 总数, tagId)</param>
    Public Shared Function ApplySemanticFormatting(
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping,
        wordParagraphs As List(Of Object),
        paragraphTypes As List(Of String),
        Optional wordApp As Object = Nothing,
        Optional onProgress As Action(Of Integer, Integer, String) = Nothing) As RenderResult

        Dim result As New RenderResult()

        ' 构建 tagId → SemanticTag 查找字典
        Dim tagDict As New Dictionary(Of String, SemanticTag)()
        For Each tag In mapping.SemanticTags
            If Not tagDict.ContainsKey(tag.TagId) Then
                tagDict(tag.TagId) = tag
            End If
        Next

        ' 遍历标注结果，逐段落应用格式
        Dim i As Integer = 0
        For Each tagged In taggedParagraphs
            If tagged.ParaIndex < 0 OrElse tagged.ParaIndex >= wordParagraphs.Count Then
                result.Errors.Add($"段落索引越界: {tagged.ParaIndex}")
                onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
                i += 1
                Continue For
            End If

            ' 跳过非文本段落
            If paragraphTypes IsNot Nothing AndAlso tagged.ParaIndex < paragraphTypes.Count Then
                Dim pType = paragraphTypes(tagged.ParaIndex)
                If pType <> "text" Then
                    result.SkippedCount += 1
                    onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
                    i += 1
                    Continue For
                End If
            End If

            ' 查找语义标签（精确匹配 → 父级回退）
            Dim semanticTag = FindTagWithFallback(tagged.TagId, tagDict, mapping)
            If semanticTag Is Nothing Then
                result.Errors.Add($"未找到标签: {tagged.TagId}")
                result.SkippedCount += 1
                onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
                i += 1
                Continue For
            End If

            ' 应用格式到段落
            Try
                Dim para = wordParagraphs(tagged.ParaIndex)
                Dim range = para.Range
                ApplyFormatToRange(range, semanticTag)
                result.AppliedCount += 1

                ' 记录标签使用次数
                If result.TagUsage.ContainsKey(tagged.TagId) Then
                    result.TagUsage(tagged.TagId) += 1
                Else
                    result.TagUsage(tagged.TagId) = 1
                End If
            Catch ex As Exception
                result.Errors.Add($"段落{tagged.ParaIndex}格式应用失败: {ex.Message}")
            End Try

            onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
            i += 1
        Next

        ' 应用页面设置（如果提供了Word应用对象）
        If wordApp IsNot Nothing AndAlso mapping.PageConfig IsNot Nothing Then
            Try
                ApplyPageConfig(wordApp, mapping.PageConfig)
            Catch ex As Exception
                result.Errors.Add($"页面设置应用失败: {ex.Message}")
            End Try
        End If

        Return result
    End Function

    ''' <summary>
    ''' 通过指令方式应用语义排版：为每个段落生成Instruction对象，再调用ApplyFormatToRange执行
    ''' 指令自带Rollback信息（原始值），支持逐条撤销和重放
    ''' </summary>
    Public Shared Function ApplySemanticFormattingViaInstructions(
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping,
        wordParagraphs As List(Of Object),
        paragraphTypes As List(Of String),
        Optional wordApp As Object = Nothing,
        Optional onProgress As Action(Of Integer, Integer, String) = Nothing) As RenderResult

        Dim result As New RenderResult()

        ' 构建 tagId → SemanticTag 查找字典
        Dim tagDict As New Dictionary(Of String, SemanticTag)()
        For Each tag In mapping.SemanticTags
            If Not tagDict.ContainsKey(tag.TagId) Then
                tagDict(tag.TagId) = tag
            End If
        Next

        ' 收集所有生成的指令（供外部使用：日志/重放/逐条撤销）
        result.GeneratedInstructions = New List(Of Instruction)()

        Dim i As Integer = 0
        For Each tagged In taggedParagraphs
            If tagged.ParaIndex < 0 OrElse tagged.ParaIndex >= wordParagraphs.Count Then
                result.Errors.Add($"段落索引越界: {tagged.ParaIndex}")
                onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
                i += 1 : Continue For
            End If

            ' 跳过非文本段落
            If paragraphTypes IsNot Nothing AndAlso tagged.ParaIndex < paragraphTypes.Count Then
                If paragraphTypes(tagged.ParaIndex) <> "text" Then
                    result.SkippedCount += 1
                    onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
                    i += 1 : Continue For
                End If
            End If

            ' 查找语义标签
            Dim semanticTag = FindTagWithFallback(tagged.TagId, tagDict, mapping)
            If semanticTag Is Nothing Then
                result.Errors.Add($"未找到标签: {tagged.TagId}")
                result.SkippedCount += 1
                onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
                i += 1 : Continue For
            End If

            ' 生成指令 + 通过ApplyFormatToRange执行（每条指令自带Rollback原始值）
            Try
                Dim para = wordParagraphs(tagged.ParaIndex)
                Dim instructions = BuildReformatInstructions(tagged.ParaIndex, semanticTag, para)
                For Each dslInstr In instructions
                    result.GeneratedInstructions.Add(dslInstr)
                Next
                ' 执行格式应用（复用已验证的直接COM操作）
                ApplyFormatToRange(para.Range, semanticTag)
                result.AppliedCount += 1

                If result.TagUsage.ContainsKey(tagged.TagId) Then
                    result.TagUsage(tagged.TagId) += 1
                Else
                    result.TagUsage(tagged.TagId) = 1
                End If
            Catch ex As Exception
                result.Errors.Add($"段落{tagged.ParaIndex}指令生成/执行失败: {ex.Message}")
            End Try

            onProgress?.Invoke(i + 1, taggedParagraphs.Count, tagged.TagId)
            i += 1
        Next

        ' 应用页面设置
        If wordApp IsNot Nothing AndAlso mapping.PageConfig IsNot Nothing Then
            Try
                ApplyPageConfig(wordApp, mapping.PageConfig)
            Catch ex As Exception
                result.Errors.Add($"页面设置应用失败: {ex.Message}")
            End Try
        End If

        Return result
    End Function

    ''' <summary>从语义标签构建排版指令列表</summary>
    Private Shared Function BuildReformatInstructions(paraIndex As Integer, tag As SemanticTag, para As Object) As List(Of Instruction)
        Dim instructions As New List(Of Instruction)()

        ' 段落级指令
        If tag.Paragraph IsNot Nothing Then
            Dim pInstr = BuildSetParagraphStyleInstruction(paraIndex, tag.Paragraph, para)
            If pInstr IsNot Nothing Then instructions.Add(pInstr)
        End If

        ' 字符级指令
        If tag.Font IsNot Nothing OrElse tag.Color IsNot Nothing Then
            Dim cInstr = BuildSetCharacterFormatInstruction(paraIndex, tag, para)
            If cInstr IsNot Nothing Then instructions.Add(cInstr)
        End If

        Return instructions
    End Function

    ''' <summary>构建段落样式指令</summary>
    Private Shared Function BuildSetParagraphStyleInstruction(paraIndex As Integer, paraCfg As ParagraphConfig, para As Object) As Instruction
        Dim target As New JObject()
        target("type") = "paraIndex"
        target("index") = paraIndex

        Dim params As New JObject()
        Dim rollback As New JObject()

        Try
            Dim rng = para.Range
            ' 对齐
            If Not String.IsNullOrEmpty(paraCfg.Alignment) Then
                params("alignment") = paraCfg.Alignment.ToLower()
                Try : rollback("alignment") = CInt(rng.ParagraphFormat.Alignment) : Catch : End Try
            End If
            ' 首行缩进
            If paraCfg.FirstLineIndent > 0 Then
                params("firstLineIndent") = paraCfg.FirstLineIndent
                Try : rollback("firstLineIndent") = CSng(rng.ParagraphFormat.FirstLineIndent) : Catch : End Try
                Try : rollback("charFirstLineIndent") = CSng(rng.ParagraphFormat.CharacterUnitFirstLineIndent) : Catch : End Try
            End If
            ' 行距
            If paraCfg.LineSpacing > 0 Then
                params("lineSpacing") = paraCfg.LineSpacing
                Try : rollback("lineSpacingRule") = CInt(rng.ParagraphFormat.LineSpacingRule) : Catch : End Try
                Try : rollback("lineSpacing") = CSng(rng.ParagraphFormat.LineSpacing) : Catch : End Try
            End If
            ' 段前
            If paraCfg.SpaceBefore > 0 Then
                params("spaceBefore") = paraCfg.SpaceBefore
                Try : rollback("spaceBefore") = CSng(rng.ParagraphFormat.SpaceBefore) : Catch : End Try
            End If
            ' 段后
            If paraCfg.SpaceAfter > 0 Then
                params("spaceAfter") = paraCfg.SpaceAfter
                Try : rollback("spaceAfter") = CSng(rng.ParagraphFormat.SpaceAfter) : Catch : End Try
            End If
        Catch ex As Exception
            Debug.WriteLine($"BuildSetParagraphStyleInstruction 捕获快照失败: {ex.Message}")
        End Try

        Dim instr As New Instruction("setParagraphStyle", params, Nothing)
        instr.Target = target
        instr.Rollback = rollback
        Return instr
    End Function

    ''' <summary>构建字符格式指令</summary>
    Private Shared Function BuildSetCharacterFormatInstruction(paraIndex As Integer, tag As SemanticTag, para As Object) As Instruction
        Dim target As New JObject()
        target("type") = "paraIndex"
        target("index") = paraIndex

        Dim params As New JObject()
        Dim rollback As New JObject()

        Try
            Dim rng = para.Range
            If tag.Font IsNot Nothing Then
                If Not String.IsNullOrEmpty(tag.Font.FontNameCN) Then
                    params("fontNameCN") = tag.Font.FontNameCN
                    Try : rollback("fontNameFarEast") = If(rng.Font.NameFarEast?.ToString(), "") : Catch : End Try
                End If
                If Not String.IsNullOrEmpty(tag.Font.FontNameEN) Then
                    params("fontNameEN") = tag.Font.FontNameEN
                    Try : rollback("fontName") = If(rng.Font.Name?.ToString(), "") : Catch : End Try
                End If
                If tag.Font.FontSize > 0 Then
                    params("fontSize") = tag.Font.FontSize
                    Try : rollback("fontSize") = CSng(rng.Font.Size) : Catch : End Try
                End If
                params("bold") = tag.Font.Bold
                Try : rollback("bold") = CInt(rng.Font.Bold) : Catch : End Try
                params("italic") = tag.Font.Italic
                Try : rollback("italic") = CInt(rng.Font.Italic) : Catch : End Try
                params("underline") = tag.Font.Underline
                Try : rollback("underline") = CInt(rng.Font.Underline) : Catch : End Try
            End If
            If tag.Color IsNot Nothing AndAlso Not String.IsNullOrEmpty(tag.Color.FontColor) Then
                params("fontColor") = tag.Color.FontColor
                Try : rollback("fontColor") = CInt(rng.Font.Color) : Catch : End Try
            End If
        Catch ex As Exception
            Debug.WriteLine($"BuildSetCharacterFormatInstruction 捕获快照失败: {ex.Message}")
        End Try

        Dim instr As New Instruction("setCharacterFormat", params, Nothing)
        instr.Target = target
        instr.Rollback = rollback
        Return instr
    End Function

    ''' <summary>查找标签（精确匹配 → 父级回退）</summary>
    Private Shared Function FindTagWithFallback(
        tagId As String,
        tagDict As Dictionary(Of String, SemanticTag),
        mapping As SemanticStyleMapping) As SemanticTag

        ' 精确匹配
        If tagDict.ContainsKey(tagId) Then Return tagDict(tagId)

        ' 父级回退
        Dim parentId = SemanticTagRegistry.GetParentTag(tagId)
        If Not String.IsNullOrEmpty(parentId) AndAlso tagDict.ContainsKey(parentId) Then
            Return tagDict(parentId)
        End If

        ' 通过mapping的FindTag方法
        Return mapping.FindTag(tagId)
    End Function

    ''' <summary>
    ''' 将语义标签的格式应用到Word Range
    ''' </summary>
    Public Shared Sub ApplyFormatToRange(targetRange As Object, tag As SemanticTag)
        If targetRange Is Nothing OrElse tag Is Nothing Then Return

        ' 应用字体
        ApplyFontConfig(targetRange, tag.Font)

        ' 应用段落格式
        ApplyParagraphConfig(targetRange, tag.Paragraph)

        ' 应用颜色
        ApplyColorConfig(targetRange, tag.Color)
    End Sub

    ''' <summary>应用字体配置</summary>
    Private Shared Sub ApplyFontConfig(targetRange As Object, font As FontConfig)
        If font Is Nothing Then Return

        Try
            ' 中文字体
            If Not String.IsNullOrEmpty(font.FontNameCN) Then
                targetRange.Font.NameFarEast = font.FontNameCN
            End If

            ' 英文字体
            If Not String.IsNullOrEmpty(font.FontNameEN) Then
                targetRange.Font.Name = font.FontNameEN
            End If

            ' 字号
            If font.FontSize > 0 Then
                targetRange.Font.Size = CSng(font.FontSize)
            End If

            ' 加粗
            targetRange.Font.Bold = If(font.Bold, -1, 0)

            ' 斜体
            targetRange.Font.Italic = If(font.Italic, -1, 0)

            ' 下划线
            If font.Underline Then
                targetRange.Font.Underline = 1 ' wdUnderlineSingle
            Else
                targetRange.Font.Underline = 0 ' wdUnderlineNone
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用字体配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>应用段落配置</summary>
    Private Shared Sub ApplyParagraphConfig(targetRange As Object, para As ParagraphConfig)
        If para Is Nothing Then Return

        Try
            ' 对齐方式
            If Not String.IsNullOrEmpty(para.Alignment) Then
                Select Case para.Alignment.ToLower()
                    Case "center"
                        targetRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                    Case "right"
                        targetRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
                    Case "justify"
                        targetRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
                    Case Else
                        targetRange.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                End Select
            End If

            ' 首行缩进
            If para.FirstLineIndent > 0 Then
                Try
                    targetRange.ParagraphFormat.CharacterUnitFirstLineIndent = CSng(para.FirstLineIndent)
                Catch
                    ' 回退: 使用磅值（1字符约10.5磅）
                    targetRange.ParagraphFormat.FirstLineIndent = CSng(para.FirstLineIndent * 10.5)
                End Try
            End If

            ' 行距
            If para.LineSpacing > 0 Then
                If para.LineSpacing = 1.0 Then
                    targetRange.ParagraphFormat.LineSpacingRule = 0 ' wdLineSpaceSingle
                ElseIf para.LineSpacing = 1.5 Then
                    targetRange.ParagraphFormat.LineSpacingRule = 1 ' wdLineSpace1pt5
                ElseIf para.LineSpacing = 2.0 Then
                    targetRange.ParagraphFormat.LineSpacingRule = 2 ' wdLineSpaceDouble
                Else
                    targetRange.ParagraphFormat.LineSpacingRule = 5 ' wdLineSpaceMultiple
                    targetRange.ParagraphFormat.LineSpacing = CSng(12 * para.LineSpacing)
                End If
            End If

            ' 段前间距
            If para.SpaceBefore > 0 Then
                targetRange.ParagraphFormat.SpaceBefore = CSng(para.SpaceBefore * 12) ' 行→磅
            End If

            ' 段后间距
            If para.SpaceAfter > 0 Then
                targetRange.ParagraphFormat.SpaceAfter = CSng(para.SpaceAfter * 12) ' 行→磅
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用段落配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>应用颜色配置</summary>
    Private Shared Sub ApplyColorConfig(targetRange As Object, color As ColorConfig)
        If color Is Nothing Then Return

        Try
            If Not String.IsNullOrEmpty(color.FontColor) AndAlso color.FontColor <> "#000000" Then
                Dim clr = System.Drawing.ColorTranslator.FromHtml(color.FontColor)
                targetRange.Font.Color = System.Drawing.ColorTranslator.ToOle(clr)
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用颜色配置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>应用页面设置</summary>
    Private Shared Sub ApplyPageConfig(wordApp As Object, config As PageConfig)
        If config Is Nothing Then Return

        Try
            Dim doc = wordApp.ActiveDocument
            Dim pageSetup = doc.PageSetup

            ' 页边距（cm → 磅，1cm = 28.35磅）
            If config.Margins IsNot Nothing Then
                Dim cmToPt As Double = 28.35
                pageSetup.TopMargin = config.Margins.Top * cmToPt
                pageSetup.BottomMargin = config.Margins.Bottom * cmToPt
                pageSetup.LeftMargin = config.Margins.Left * cmToPt
                pageSetup.RightMargin = config.Margins.Right * cmToPt
            End If
        Catch ex As Exception
            Debug.WriteLine($"应用页面设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 排版后校验 - 检查渲染结果是否匹配预期
    ''' </summary>
    Public Shared Function ValidateRenderedDocument(
        wordParagraphs As List(Of Object),
        taggedParagraphs As List(Of TaggedParagraph),
        mapping As SemanticStyleMapping) As List(Of String)

        Dim deviations As New List(Of String)()

        For Each tagged In taggedParagraphs
            If tagged.ParaIndex < 0 OrElse tagged.ParaIndex >= wordParagraphs.Count Then Continue For

            Dim expectedTag = mapping.FindTag(tagged.TagId)
            If expectedTag Is Nothing Then Continue For

            Try
                Dim para = wordParagraphs(tagged.ParaIndex)
                Dim range = para.Range

                ' 检查字号
                If expectedTag.Font.FontSize > 0 Then
                    Dim actualSize As Double = CDbl(range.Font.Size)
                    If Math.Abs(actualSize - expectedTag.Font.FontSize) > 0.5 Then
                        deviations.Add($"段落{tagged.ParaIndex}: 字号偏差 期望{expectedTag.Font.FontSize}pt 实际{actualSize}pt")
                        ' 自动修正
                        range.Font.Size = CSng(expectedTag.Font.FontSize)
                    End If
                End If
            Catch ex As Exception
                ' 校验失败不影响主流程
                Debug.WriteLine($"校验段落{tagged.ParaIndex}失败: {ex.Message}")
            End Try
        Next

        Return deviations
    End Function
End Class

''' <summary>
''' AI标注的段落结构
''' </summary>
Public Class TaggedParagraph
    ''' <summary>段落索引</summary>
    Public Property ParaIndex As Integer

    ''' <summary>语义标签ID</summary>
    Public Property TagId As String = ""

    Public Sub New()
    End Sub

    Public Sub New(paraIndex As Integer, tagId As String)
        Me.ParaIndex = paraIndex
        Me.TagId = tagId
    End Sub
End Class
