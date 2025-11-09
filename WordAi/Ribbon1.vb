' WordAi\Ribbon1.vb
Imports System.Diagnostics
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Xml
Imports AngleSharp
Imports Microsoft.Office.Tools.Ribbon
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon  ' 添加此引用

Public Class Ribbon1
    Inherits BaseOfficeRibbon

    Protected Overrides Async Sub ChatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Async Sub WebResearchButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowDataCaptureTaskPane()
    End Sub
    Protected Overrides Sub SpotlightButton_Click(sender As Object, e As RibbonControlEventArgs)
        'Globals.ThisAddIn.ShowChatTaskPane()
    End Sub
    Protected Overrides Sub DataAnalysisButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' Word 特定的数据分析逻辑
        MessageBox.Show("Word数据分析功能正在开发中...")
    End Sub

    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Word", OfficeApplicationType.Word)
    End Function

    Protected Overrides Sub DeepseekButton_Click(sender As Object, e As RibbonControlEventArgs)
        Globals.ThisAddIn.ShowDeepseekTaskPane()
    End Sub
    Protected Overrides Sub BatchDataGenButton_Click(sender As Object, e As RibbonControlEventArgs)
    End Sub

    Protected Overrides Sub MCPButton_Click(sender As Object, e As RibbonControlEventArgs)
        ' 创建并显示MCP配置表单
        Dim mcpConfigForm As New MCPConfigForm()
        If mcpConfigForm.ShowDialog() = DialogResult.OK Then
            ' 在需要时可以集成到ChatControl调用MCP服务
        End If
    End Sub
    ' Proofread 按钮 — 直接使用 ThisAddIn.chatControl
    Protected Overrides Async Sub ProofreadButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            Dim selText As String = String.Empty

            Try
                If wordApp IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                    selText = If(wordApp.Selection.Range IsNot Nothing, wordApp.Selection.Range.Text, String.Empty)
                End If
            Catch ex As Exception
                selText = String.Empty
            End Try

            Dim useWholeDoc As Boolean = False
            If String.IsNullOrWhiteSpace(selText) Then
                Dim result = MessageBox.Show("当前未选中文本，是否要对整个文档进行校对？", "确认校对范围", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result <> DialogResult.Yes Then
                    Return
                End If
                useWholeDoc = True
            End If

            Dim targetText As String = If(useWholeDoc, Globals.ThisAddIn.Application.ActiveDocument.Content.Text, selText)

            ' 打开侧栏（CreateChatTaskPane 内已保证单例）
            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(250)

            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                MessageBox.Show("无法获取聊天控件实例，请确认 Chat 面板已打开。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 前端提示
            Try
                Dim responseUuid As String = Guid.NewGuid().ToString()
                Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName
                Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
                Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCreate)
                Dim js = $"appendRenderer('{responseUuid}','正在向模型发起校对请求，请耐心等待');"
                Await chatCtrl.ExecuteJavaScriptAsyncJS(js)
            Catch ex As Exception
                Debug.WriteLine("ExecuteJavaScriptAsyncJS 调用失败: " & ex.Message)
            End Try

            ' 构建提示词
            Dim sb As New StringBuilder()
            sb.AppendLine("你是严格的Word校对助手。请基于下方原文找出所有需要修正的错字、错标点或需插入的换行。")
            sb.AppendLine("必须且仅返回一个 JSON 数组，数组项格式如下（严格按此字段名）：")
            sb.AppendLine("[{")
            sb.AppendLine("  ""index"": 0,")
            sb.AppendLine("  ""action"": ""replace"",          // insert|delete|replace")
            sb.AppendLine("  ""matchText"": ""原文片段"",      // 要定位并替换/删除/插入的片段，必填（insert 时可为空）")
            sb.AppendLine("  ""contextBefore"": ""片段前若干字符（可空）"",")
            sb.AppendLine("  ""contextAfter"": ""片段后若干字符（可空）"",")
            sb.AppendLine("  ""replaceWith"": ""替换为的文字（insert 或 replace 用）"",")
            sb.AppendLine("  ""rule"": ""错词约束|错标点约束|换行约束"",")
            sb.AppendLine("  ""note"": ""可选：给人的简短说明''")
            sb.AppendLine("}]")
            sb.AppendLine()
            sb.AppendLine("说明：")
            sb.AppendLine("- 使用 contextBefore/contextAfter 提供上下文锚点以便准确定位（尽量提供 6-30 字符）。")
            sb.AppendLine("- 不要输出任何非 JSON 的内容。")
            targetText = "以下为需要修订的原文，请开始你的工作：" & targetText
            Await chatCtrl.Send(targetText, sb.ToString(), False, "proofread")
        Catch ex As Exception
            MessageBox.Show("执行校对过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Protected Overrides Async Sub ReformatButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim wordApp = Globals.ThisAddIn.Application
            Dim selText As String = String.Empty

            Try
                If wordApp IsNot Nothing AndAlso wordApp.Selection IsNot Nothing Then
                    selText = If(wordApp.Selection.Range IsNot Nothing, wordApp.Selection.Range.Text, String.Empty)
                End If
            Catch
                selText = String.Empty
            End Try

            Dim useWholeDoc As Boolean = False
            If String.IsNullOrWhiteSpace(selText) Then
                Dim result = MessageBox.Show("当前未选中文本，是否要对整个文档进行排版/重构？", "确认排版范围", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result <> DialogResult.Yes Then
                    Return
                End If
                useWholeDoc = True
            End If

            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim baseRange = If(useWholeDoc, doc.Content, wordApp.Selection.Range)

            ' 将文档分块（段落、表格、图片、公式）
            Dim blocks As New Newtonsoft.Json.Linq.JArray()
            Dim blockIndex As Integer = 0

            For Each p In baseRange.Paragraphs
                Dim r = p.Range
                ' 表格独立为块
                If r.Tables.Count > 0 Then
                    For i = 1 To r.Tables.Count
                        Dim t = r.Tables(i)
                        Dim tRange = t.Range
                        Dim tblObj As New Newtonsoft.Json.Linq.JObject()
                        tblObj("id") = "blk_" & blockIndex
                        tblObj("type") = "table"
                        tblObj("text") = "" ' 可选简述
                        Try
                            tblObj("wordOpenXml") = tRange.WordOpenXML
                        Catch
                        End Try
                        blocks.Add(tblObj)
                        blockIndex += 1
                    Next
                ElseIf r.InlineShapes.Count > 0 OrElse r.ShapeRange.Count > 0 Then
                    Dim imgObj As New Newtonsoft.Json.Linq.JObject()
                    imgObj("id") = "blk_" & blockIndex
                    imgObj("type") = "image"
                    ' 安全转换为字符串，避免 COM 类型映射异常
                    imgObj("text") = If(r.Text IsNot Nothing, r.Text.ToString(), String.Empty)
                    blocks.Add(imgObj)
                    blockIndex += 1
                ElseIf r.OMaths IsNot Nothing AndAlso r.OMaths.Count > 0 Then
                    Dim eqObj As New Newtonsoft.Json.Linq.JObject()
                    eqObj("id") = "blk_" & blockIndex
                    eqObj("type") = "equation"
                    eqObj("text") = If(r.Text IsNot Nothing, r.Text.ToString(), String.Empty)
                    Try
                        eqObj("wordOpenXml") = r.WordOpenXML
                    Catch
                    End Try
                    blocks.Add(eqObj)
                    blockIndex += 1
                Else
                    Dim paraObj As New Newtonsoft.Json.Linq.JObject()
                    paraObj("id") = "blk_" & blockIndex
                    paraObj("type") = "paragraph"
                    paraObj("text") = If(r.Text IsNot Nothing, r.Text.ToString(), String.Empty)
                    blocks.Add(paraObj)
                    blockIndex += 1
                End If
            Next

            ' 构建可控的系统提示（增加 DocumentFormat.OpenXml.Packaging 可用动作说明）
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("你是专业且严格的文档排版助手。输入为 blocks 数组（每个块包含 id, type, text, 可选 wordOpenXml/tableData 等）。")
            sb.AppendLine("请严格且只返回一个 JSON 对象，格式如下（不要输出其它任何文本）：")
            GetPrompt1(sb)
            sb.AppendLine("{")
            sb.AppendLine("  ""documentPlan"": [ /* 每项：{""blockId"":..., ""action"":..., ""attributes"": {...}, ""note"":...} */ ],")
            sb.AppendLine("  ""previewHtmlMap"": { /* 可选：键为 blockId，值为 HTML 用于前端预览 */ }")
            sb.AppendLine("}")

            sb.AppendLine()
            sb.AppendLine(blocks.ToString(Newtonsoft.Json.Formatting.None))

            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(250)

            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                MessageBox.Show("无法获取聊天控件实例，请确认 Chat 面板已打开。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 发送（responseMode 使用 reformat）
            Await chatCtrl.Send("请基于提供的 blocks 输出 documentPlan 与 previewHtmlMap（严格JSON格式）。", sb.ToString(), False, "reformat")

        Catch ex As Exception
            MessageBox.Show("执行排版过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetPrompt1(sb As System.Text.StringBuilder) As String
        sb.AppendLine("注意：我们在插件侧使用 DocumentFormat.OpenXml.Packaging 库来执行具体的回写/格式化操作。")
        sb.AppendLine("请只使用下列允许的高层 action 名称，复杂或精细的修改应通过 attributes.openXmlActions 字段以数组形式声明，openXmlActions 中每项为：")
        sb.AppendLine("{ ""verb"": ""<openxml_verb>"", ""params"": { ... } }")
        sb.AppendLine()
        sb.AppendLine("Allowed high-level actions（必须使用其中之一）：")
        'sb.AppendLine("- replaceText / replace                : 使用 attributes.wordOpenXml 替换目标块，注意要严格最新wordOpenXml替换格式，即定位一定要准确")
        sb.AppendLine("- insert                               : 在目标块前插入 attributes.text 或 attributes.previewHtml")
        sb.AppendLine("- delete                               : 删除目标块")
        sb.AppendLine("- format                               : 对文本段落应用字体/对齐/样式等（避免直接修改表格/图片/公式）")
        sb.AppendLine("- skip                                 : 明确跳过该块")
        sb.AppendLine()
        sb.AppendLine("当需要更精细地利用 DocumentFormat.OpenXml.Packaging 时，请在 attributes 中使用 openXmlActions，例如：")
        sb.AppendLine("""attributes"": {")
        sb.AppendLine("  ""openXmlActions"": [")
        sb.AppendLine("    { ""verb"": ""insertParagraph"", ""params"": { ""text"": ""新的段落"", ""paragraphProps"":{""alignment"":""right""} } },")
        sb.AppendLine("    { ""verb"": ""setRunProperties"", ""params"": { ""fontName"":""仿宋"", ""fontSize"":12, ""bold"":true } },")
        sb.AppendLine("    { ""verb"": ""insertTable"", ""params"": { ""rows"": [[\""A\"",\""B\""],[\""C\"",\""D\""]], ""tableProps"":{...} } },")
        sb.AppendLine("    { ""verb"": ""insertImage"", ""params"": { ""imageUrl"": ""https://...jpg"", ""width"":200, ""height"":100 } },")
        sb.AppendLine("    { ""verb"": ""setParagraphSpacing"", ""params"": { ""before"":6, ""after"":6, ""line"": 360 } }")
        sb.AppendLine("  ]")
        sb.AppendLine("}")
        sb.AppendLine()
        sb.AppendLine("建议模型在生成 openXmlActions 时优先使用下列常见 OpenXML 操作 verb（插件侧会把 verb 映射到 DocumentFormat.OpenXml.Packaging + OpenXml SDK 的实现）：")
        'sb.AppendLine("- findAndReplace (基于 filePath/tagName/newText 精确定位并用 WordProcessingML 替换匹配范围)")
        'sb.AppendLine("- replaceWithWordOpenXml (直接替换指定 range 为给定 WordProcessingML 片段)")
        sb.AppendLine("- insertParagraph / insertRun / insertText")
        sb.AppendLine("- setRunProperties (font name/size/bold/italic/underline/color)")
        sb.AppendLine("- setParagraphProperties (alignment, indentation, spacing, numberingReference)")
        sb.AppendLine("- applyStyle / createStyle")
        sb.AppendLine("- insertTable / setTableProperties / insertTableRow / mergeTableCells / splitTableCell / setTableCellProperties")
        sb.AppendLine("- insertImage (通过关系添加图片并生成 Drawing 元素)")
        sb.AppendLine("- insertEquation (OfficeMath 元素，或插入 wordOpenXml 的公式片段)")
        sb.AppendLine("- insertHyperlink / setBookmark / insertContentControl (Sdt)")
        sb.AppendLine("- insertHeader / insertFooter / setSectionProperties (page size/margins/columns)")
        sb.AppendLine("- insertPageBreak / insertSectionBreak")
        sb.AppendLine("- addFootnote / addEndnote")
        sb.AppendLine("- setTableCellShading / setTableBorders / setRunHighlight")
        sb.AppendLine()
        sb.AppendLine("说明与约束：")
        sb.AppendLine("- 对于表格/图片/公式类块：优先返回 attributes.wordOpenXml 或 openXmlActions 能直接生成 WordProcessingML 的操作，插件将以 OpenXml API 应用。")
        sb.AppendLine("- 对于复杂结构（表格、公式、嵌入对象），若无法用 OpenXML 可靠回写，可在 note 中提示并返回 skip。")
        sb.AppendLine("- 严格返回 JSON，openXmlActions 中 verb 名称和值须为字符串或数字/布尔或数组，不要包含注释或解释文本。")
        Return sb.ToString

    End Function

    Private Function GetPrompt2(sb As System.Text.StringBuilder) As String
        sb.AppendLine("注意：我们在插件侧使用 DocumentFormat.OpenXml.Packaging（OpenXml SDK）来执行具体的回写/格式化操作。")
        sb.AppendLine("为了使回写安全且可精确定位，请严格遵守下列要求：")
        sb.AppendLine("1) 对于 action 为 replace / replaceText：")
        sb.AppendLine("   - 优先返回 ""wordOpenXml"" 字段（WordProcessingML 片段，表示要写回的目标结构），插件将把该片段写入目标位置；")
        sb.AppendLine("   - 或者返回 ""attributes.openXmlActions""（数组），每项为 { ""verb"": <verb>, ""params"": {...} }，插件会把 verb 映射到 OpenXml SDK 的实现并执行；")
        sb.AppendLine("   - 如果无法提供 wordOpenXml 或 openXmlActions，必须提供可用于定位的锚点（matchText 与 contextBefore/contextAfter 或 start 偏移），否则请返回 action = \""skip\"" 并在 note 中说明原因。")
        sb.AppendLine("2) openXmlActions 推荐 verb（示例，插件侧会映射执行）：")
        sb.AppendLine("   - ""findAndReplace"" : { ""matchText"": \""...\ "", ""contextBefore"": \""...\ "", ""contextAfter"": \""...\ "", ""replaceWithWordOpenXml"": \"" < w: Document> ...</w:Document>\ "" }")
        sb.AppendLine("   - ""replaceWithWordOpenXml"" : 直接替换指定定位范围为提供的 WordProcessingML 片段")
        sb.AppendLine("   - ""insertParagraph"" / ""insertTable"" / ""insertImage"" / ""setRunProperties"" / ""setParagraphProperties"" 等（详见下方 verb 列表）")
        sb.AppendLine("3) 对于表格/图片/公式等复杂对象：")
        sb.AppendLine("   - 强制优先提供 wordOpenXml 或 openXmlActions 能直接生成 WordProcessingML 的操作；")
        sb.AppendLine("   - 若模型无法返回可靠 OpenXML 指令，则返回 action = \""skip\"" 并在 note 中说明，避免前端盲写回破坏文档结构。")
        sb.AppendLine()
        sb.AppendLine("示例（严格 JSON，替换用 openXmlActions）：")
        sb.AppendLine("{")
        sb.AppendLine("  ""blockId"": ""blk_3"",")
        sb.AppendLine("  ""action"": ""replace"",")
        sb.AppendLine("  ""attributes"": {")
        sb.AppendLine("    ""openXmlActions"": [")
        sb.AppendLine("      {")
        sb.AppendLine("        ""verb"": ""findAndReplace"",")
        sb.AppendLine("        ""params"": {")
        sb.AppendLine("          ""matchText"": ""原文片段"",")
        sb.AppendLine("          ""contextBefore"": ""片段前 8-20 字符"",")
        sb.AppendLine("          ""contextAfter"": ""片段后 8-20 字符"",")
        sb.AppendLine("          ""replaceWithWordOpenXml"": ""<w:document>...替换内容的 WordProcessingML ...</w:document>""")
        sb.AppendLine("        }")
        sb.AppendLine("      }")
        sb.AppendLine("    ]")
        sb.AppendLine("  },")
        sb.AppendLine("  ""note"": ""优先用 OpenXML 精确替换，否则跳过""")
        sb.AppendLine("}")
        sb.AppendLine()
        sb.AppendLine("示例（严格 JSON，替换直接提供 wordOpenXml）：")
        sb.AppendLine("{ ""blockId"": ""blk_3"", ""action"": ""replace"", ""wordOpenXml"": ""<w:document>...完整或片段WordProcessingML...</w:document>"" }")
        sb.AppendLine()
        sb.AppendLine("下列 verb 为推荐映射（插件会实现这些 verb）：")
        sb.AppendLine("- findAndReplace (基于 matchText/context 精确定位并用 WordProcessingML 替换匹配范围)")
        sb.AppendLine("- replaceWithWordOpenXml (直接替换指定 range 为给定 WordProcessingML 片段)")
        sb.AppendLine("- insertParagraph / insertRun / insertText")
        sb.AppendLine("- setRunProperties / setParagraphProperties / applyStyle")
        sb.AppendLine("- insertTable / setTableProperties / insertTableRow / mergeTableCells")
        sb.AppendLine("- insertImage (通过关系添加图片并生成 Drawing 元素)")
        sb.AppendLine("- insertEquation / insertHyperlink / insertContentControl")
        sb.AppendLine("- insertHeader / insertFooter / setSectionProperties")
        sb.AppendLine()
        sb.AppendLine("重要提醒：如果模型返回无法精确定位的 free-text 替换（仅提供 replaceWith 而无锚点或 OpenXML），插件会拒绝执行并要求模型补充定位信息或返回 skip。")
        sb.AppendLine()
        Return sb.ToString
    End Function

End Class
