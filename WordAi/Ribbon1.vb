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

    ' Proofread 按钮 — 校对功能（简化版：仅校对选中内容，使用段落索引定位）
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

            ' 必须先选中内容
            If String.IsNullOrWhiteSpace(selText) Then
                MessageBox.Show("请先选中需要校对的文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selRange = wordApp.Selection.Range

            ' 按段落分割选中内容，构建带索引的文本
            Dim paragraphs As New List(Of String)()
            Dim paraIndex As Integer = 0
            Dim sb As New StringBuilder()
            sb.AppendLine("以下是需要校对的内容（按段落编号）：")

            For Each p In selRange.Paragraphs
                Dim paraText As String = If(p.Range.Text IsNot Nothing, p.Range.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)
                If Not String.IsNullOrWhiteSpace(paraText) Then
                    paragraphs.Add(paraText)
                    sb.AppendLine($"[段落{paraIndex}] {paraText}")
                    paraIndex += 1
                End If
            Next

            If paragraphs.Count = 0 Then
                MessageBox.Show("选中的内容没有有效段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 打开侧栏
            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(250)

            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                MessageBox.Show("无法获取聊天控件实例，请确认 Chat 面板已打开。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 构建前端提示
            buildHtmlHint(chatCtrl, "正在向模型发起校对请求，请耐心等待")

            ' 构建简化的提示词
            Dim systemPrompt As New StringBuilder()
            systemPrompt.AppendLine("你是Word内容校对助手。请检查以下内容中的错字、错标点或不当换行，并给出修正建议。")
            systemPrompt.AppendLine("必须且仅返回一个严格的JSON数组，格式如下：")
            systemPrompt.AppendLine("[{")
            systemPrompt.AppendLine("  ""paraIndex"": 0,")
            systemPrompt.AppendLine("  ""original"": ""原文片段"",")
            systemPrompt.AppendLine("  ""corrected"": ""修正后的文字"",")
            systemPrompt.AppendLine("  ""reason"": ""简短说明修正原因""")
            systemPrompt.AppendLine("}]")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("注意：")
            systemPrompt.AppendLine("- paraIndex是段落编号，从0开始")
            systemPrompt.AppendLine("- 如果没有需要修正的内容，返回空数组[]")
            systemPrompt.AppendLine("- 不要输出任何非JSON内容")

            Await chatCtrl.Send(sb.ToString(), systemPrompt.ToString(), False, "proofread")
        Catch ex As Exception
            MessageBox.Show("执行校对过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub buildHtmlHint(chatCtrl As ChatControl, displayContent As String)

        Try
            Dim responseUuid As String = Guid.NewGuid().ToString()
            Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName
            Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(jsCreate)
            Dim js = $"appendRenderer('{responseUuid}','{displayContent}');"
            Await chatCtrl.ExecuteJavaScriptAsyncJS(js)
        Catch ex As Exception
            Debug.WriteLine("ExecuteJavaScriptAsyncJS 调用失败: " & ex.Message)
        End Try
    End Sub

    ' 排版功能（简化版：按段落获取WordOpenXML并替换）
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

            ' 必须先选中内容
            If String.IsNullOrWhiteSpace(selText) Then
                MessageBox.Show("请先选中需要排版的文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selRange = wordApp.Selection.Range

            ' 按段落分割，构建简化的blocks数组
            Dim blocks As New Newtonsoft.Json.Linq.JArray()
            Dim paraIndex As Integer = 0

            For Each p In selRange.Paragraphs
                Dim r = p.Range
                Dim paraText As String = If(r.Text IsNot Nothing, r.Text.ToString().TrimEnd(vbCr, vbLf), String.Empty)

                If Not String.IsNullOrWhiteSpace(paraText) Then
                    Dim paraObj As New Newtonsoft.Json.Linq.JObject()
                    paraObj("paraIndex") = paraIndex
                    paraObj("text") = paraText
                    ' 不再传递 WordOpenXML，只传递文本

                    blocks.Add(paraObj)
                    paraIndex += 1
                End If
            Next

            If blocks.Count = 0 Then
                MessageBox.Show("选中的内容没有有效段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Globals.ThisAddIn.ShowChatTaskPane()
            Await Task.Delay(250)

            Dim chatCtrl = Globals.ThisAddIn.chatControl
            If chatCtrl Is Nothing Then
                MessageBox.Show("无法获取聊天控件实例，请确认 Chat 面板已打开。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 构建前端提示
            buildHtmlHint(chatCtrl, "正在向模型发起排版请求，请耐心等待")

            ' 构建简化的系统提示 - 返回格式化属性而非XML
            Dim systemPrompt As New System.Text.StringBuilder()
            systemPrompt.AppendLine("你是Word排版助手。我提供文档段落，请帮我优化排版。")
            systemPrompt.AppendLine("排版规则：")
            systemPrompt.AppendLine("1. 中文字体使用宋体，英文使用Times New Roman")
            systemPrompt.AppendLine("2. 正文字号12pt（小四），标题根据级别设置（如16pt/14pt）")
            systemPrompt.AppendLine("3. 段落首行缩进2字符")
            systemPrompt.AppendLine("4. 行距1.5倍")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("必须且仅返回一个严格的JSON数组，格式如下：")
            systemPrompt.AppendLine("[{")
            systemPrompt.AppendLine("  ""paraIndex"": 0,")
            systemPrompt.AppendLine("  ""formatting"": {")
            systemPrompt.AppendLine("    ""fontNameCN"": ""宋体"",")
            systemPrompt.AppendLine("    ""fontNameEN"": ""Times New Roman"",")
            systemPrompt.AppendLine("    ""fontSize"": 12,")
            systemPrompt.AppendLine("    ""bold"": false,")
            systemPrompt.AppendLine("    ""alignment"": ""left"",")
            systemPrompt.AppendLine("    ""firstLineIndent"": 2,")
            systemPrompt.AppendLine("    ""lineSpacing"": 1.5")
            systemPrompt.AppendLine("  },")
            systemPrompt.AppendLine("  ""previewText"": ""格式化后的纯文本预览"",")
            systemPrompt.AppendLine("  ""changes"": ""简述做了哪些格式修改""")
            systemPrompt.AppendLine("}]")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("formatting字段说明：")
            systemPrompt.AppendLine("- fontNameCN: 中文字体名称")
            systemPrompt.AppendLine("- fontNameEN: 英文字体名称")
            systemPrompt.AppendLine("- fontSize: 字号(pt)")
            systemPrompt.AppendLine("- bold: 是否加粗(true/false)")
            systemPrompt.AppendLine("- alignment: 对齐方式(left/center/right/justify)")
            systemPrompt.AppendLine("- firstLineIndent: 首行缩进字符数(0表示无缩进)")
            systemPrompt.AppendLine("- lineSpacing: 行距倍数(1/1.5/2等)")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("注意：")
            systemPrompt.AppendLine("- paraIndex是段落编号，与输入对应")
            systemPrompt.AppendLine("- 如果段落无需修改，可以不包含该段落")
            systemPrompt.AppendLine("- 不要输出任何非JSON内容")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("以下是需要排版的段落：")
            systemPrompt.AppendLine(blocks.ToString(Newtonsoft.Json.Formatting.Indented))

            Await chatCtrl.Send("请基于提供的段落进行排版优化。", systemPrompt.ToString(), False, "reformat")

        Catch ex As Exception
            MessageBox.Show("执行排版过程出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
