Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Math
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Reflection.Emit
Imports System.Text
Imports System.Text.JSON
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Windows.Forms
Imports System.Windows.Forms.ListBox
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Tab
Imports Markdig
Imports Microsoft.Vbe.Interop
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon
Imports DocumentFormat.OpenXml.Packaging
Imports HtmlToOpenXml

Public Class ChatControl
    Inherits BaseChatControl


    Private sheetContentItems As New Dictionary(Of String, Tuple(Of System.Windows.Forms.Label, System.Windows.Forms.Button))


    Public Sub New()
        ' 此调用是设计师所必需的。
        InitializeComponent()

        ' 确保WebView2控件可以正常交互
        ChatBrowser.BringToFront()

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)

        ' 订阅Word的SelectionChange 事件
        ' 帮我补全word选择的内容事件
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, AddressOf GetSelectionContent
    End Sub

    '获取选中的内容
    Protected Overrides Sub GetSelectionContent(target As Object)
        Try
            If Not Me.Visible OrElse Not selectedCellChecked Then
                Return
            End If

            ' 转换为 Word.Selection 对象
            Dim selection = TryCast(Globals.ThisAddIn.Application.Selection, Microsoft.Office.Interop.Word.Selection)
            If selection Is Nothing Then
                Return
            End If

            ' 检查是否有实际选中内容（通过比较Start和End位置）
            If selection.Start = selection.End Then
                ' 光标在单一位置，没有选中内容，清除之前的选中显示
                ClearSelectedContentBySheetName("Word文档")
                Return
            End If

            ' 获取选中内容的详细信息
            Dim content As String = String.Empty

            ' 检查是否选中了表格
            If selection.Tables.Count > 0 Then
                ' 如果选中的是表格
                Dim table = selection.Tables(1)
                Dim sb As New StringBuilder()

                ' 遍历表格内容
                For row As Integer = 1 To table.Rows.Count
                    For col As Integer = 1 To table.Columns.Count
                        sb.Append(table.Cell(row, col).Range.Text.TrimEnd(ChrW(13), ChrW(7)))
                        If col < table.Columns.Count Then sb.Append(vbTab)
                    Next
                    sb.AppendLine()
                Next
                content = sb.ToString()

            ElseIf selection.InlineShapes.Count > 0 OrElse selection.ShapeRange.Count > 0 Then
                ' 如果选中的是图片或形状
                content = "[图片或形状]"
            Else
                ' 普通文本选择
                content = selection.Text
            End If

            If Not String.IsNullOrEmpty(content) Then
                ' 添加到选中内容列表
                AddSelectedContentItem(
                "Word文档",  ' 使用文档名称作为标识
                If(selection.Tables.Count > 0,
                   "[表格内容]",
                   content.Substring(0, Math.Min(content.Length, 50)) & If(content.Length > 50, "...", ""))
            )
            Else
                ClearSelectedContentBySheetName("Word文档")
            End If

        Catch ex As Exception
            Debug.WriteLine($"获取Word选中内容时出错: {ex.Message}")
        End Try
    End Sub


    ' 获取选中内容的详细信息
    Private Function GetSelectionDetails(selection As Microsoft.Office.Interop.Word.Selection) As String
        Dim details As New StringBuilder()

        ' 添加基本信息
        details.AppendLine($"开始位置: {selection.Start}")
        details.AppendLine($"结束位置: {selection.End}")
        details.AppendLine($"字符数: {selection.Characters.Count}")

        ' 如果是表格，添加表格信息
        If selection.Tables.Count > 0 Then
            Dim table = selection.Tables(1)
            details.AppendLine($"表格大小: {table.Rows.Count}行 x {table.Columns.Count}列")
        End If

        Return details.ToString()
    End Function

    ' 初始化时注入基础 HTML 结构
    Private Async Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 初始化 WebView2
        Await InitializeWebView2()
        InitializeWebView2Script()
    End Sub


    ' 返回应用信息
    Protected Overrides Function GetApplication() As ApplicationInfo
        Return New ApplicationInfo("Word", OfficeApplicationType.Word)
    End Function

    ' 返回Office应用类型
    Protected Overrides Function GetOfficeAppType() As String
        Return "Word"
    End Function

    ' 返回 Word Application 对象
    Protected Overrides Function GetOfficeApplicationObject() As Object
        Return Globals.ThisAddIn.Application
    End Function

    ' 返回当前文档的 VBProject（可能为 Nothing）
    Protected Overrides Function GetVBProject() As VBProject
        Try
            Return Globals.ThisAddIn.Application.ActiveDocument.VBProject
        Catch
            Return Nothing
        End Try
    End Function

    ' 预览运行：展示代码并询问是否继续（返回 True 执行）
    Protected Overrides Function RunCodePreview(vbaCode As String, preview As Boolean) As Boolean
        If Not preview Then Return True
        Dim prompt As String = "预览将要执行的 VBA 代码，是否继续？" & vbCrLf & "----" & vbCrLf & vbaCode
        Return (MessageBox.Show(prompt, "VBA 预览", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes)
    End Function

    ' 真正执行宏（通过 Application.Run 调用模块.过程）
    Protected Overrides Function RunCode(vbaCode As String) As Object
        Try
            Globals.ThisAddIn.Application.Run(vbaCode)
        Catch ex As Exception
            MessageBox.Show("执行宏失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

    ' 将要发送到 LLM 的消息委托到底层 Send 方法（异步）
    Protected Overrides Sub SendChatMessage(message As String)
        Task.Run(Async Function()
                     Await Send(message, "", True, "")
                 End Function)
    End Sub

    ''' <summary>
    ''' 使用意图识别结果发送聊天消息（重写基类方法）
    ''' </summary>
    Protected Overrides Sub SendChatMessageWithIntent(message As String, intent As IntentResult)
        If intent IsNot Nothing AndAlso intent.Confidence > 0.2 Then
            Dim optimizedPrompt = IntentService.GetOptimizedSystemPrompt(intent)
            Debug.WriteLine($"Word使用意图优化提示词: {intent.IntentType}, 置信度: {intent.Confidence:F2}")

            Task.Run(Async Function()
                         Await Send(message, optimizedPrompt, True, "")
                     End Function)
        Else
            ' 回退到普通发送
            SendChatMessage(message)
        End If
    End Sub

    ' 解析 Word 文件为文本（用于 file 引用）
    Protected Overrides Function ParseFile(filePath As String) As FileContentResult
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim doc = app.Documents.Open(FileName:=filePath, ReadOnly:=True, Visible:=False)
            Dim txt = doc.Content.Text
            doc.Close(False)
            Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Word",
                .ParsedContent = txt,
                .RawData = Nothing
            }
        Catch ex As Exception
            Return New FileContentResult With {
                .FileName = Path.GetFileName(filePath),
                .FileType = "Word",
                .ParsedContent = $"[解析文档失败: {ex.Message}]"
            }
        End Try
    End Function

    ' 返回当前文档所在目录（未保存返回空字符串）
    Protected Overrides Function GetCurrentWorkingDirectory() As String
        Try
            Dim p = Globals.ThisAddIn.Application.ActiveDocument.Path
            If String.IsNullOrEmpty(p) Then Return String.Empty
            Return p
        Catch
            Return String.Empty
        End Try
    End Function

    ' 将当前选区内容附加到提示，并记录 PendingSelectionInfo 供写回使用
    Protected Overrides Function AppendCurrentSelectedContent(message As String) As String
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            Dim txt As String = If(sel IsNot Nothing AndAlso sel.Range IsNot Nothing, sel.Range.Text, String.Empty)

            Dim info As New SelectionInfo()
            info.DocumentPath = If(Globals.ThisAddIn.Application.ActiveDocument.Path, "")
            info.SelectedText = txt
            Try
                info.StartPos = sel.Range.Start
                info.EndPos = sel.Range.End
            Catch
                info.StartPos = 0
                info.EndPos = 0
            End Try

            PendingSelectionInfo = info

            If String.IsNullOrWhiteSpace(txt) Then
                Return message
            Else
                Return message & vbCrLf & vbCrLf & txt
            End If
        Catch
            Return message
        End Try
    End Function


    ' 修订、审阅功能（简化版：使用段落索引定位）
    Protected Overrides Sub HandleApplyRevisionSegment(jsonDoc As JObject)
        Try
            ' 期望收到字段： uuid, paraIndex, original, corrected
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim paraIndex As Integer = If(jsonDoc("paraIndex") IsNot Nothing, CInt(jsonDoc("paraIndex")), -1)
            Dim original As String = If(jsonDoc("original") IsNot Nothing, jsonDoc("original").ToString(), String.Empty)
            Dim corrected As String = If(jsonDoc("corrected") IsNot Nothing, jsonDoc("corrected").ToString(), String.Empty)

            If paraIndex < 0 Then
                GlobalStatusStrip.ShowWarning("缺少 paraIndex 参数")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("校对功能仅在 Word 环境下支持")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try
            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象")
                Return
            End If

            Dim doc = officeApp.ActiveDocument
            Dim selRange = officeApp.Selection.Range

            ' 使用选中范围内的段落索引定位
            If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(selRange.Text) Then
                GlobalStatusStrip.ShowWarning("请先选中需要校对的内容")
                Return
            End If

            ' 获取选中范围内的段落
            Dim paragraphs = selRange.Paragraphs
            If paraIndex >= paragraphs.Count Then
                GlobalStatusStrip.ShowWarning($"段落索引 {paraIndex} 超出范围")
                Return
            End If

            ' 定位目标段落（段落索引从1开始）
            Dim targetPara = paragraphs(paraIndex + 1)
            Dim targetRange = targetPara.Range

            ' 在目标段落中查找并替换原文
            If Not String.IsNullOrEmpty(original) Then
                Dim paraText As String = targetRange.Text
                Dim startPos As Integer = paraText.IndexOf(original, StringComparison.Ordinal)
                If startPos >= 0 Then
                    ' 创建精确的替换范围
                    Dim replaceRange = doc.Range(targetRange.Start + startPos, targetRange.Start + startPos + original.Length)

                    ' 开启审阅模式
                    Try
                        doc.TrackRevisions = True
                    Catch
                    End Try

                    ' 执行替换
                    replaceRange.Text = corrected
                    GlobalStatusStrip.ShowInfo($"已替换段落 {paraIndex} 中的内容（审阅模式）")
                Else
                    GlobalStatusStrip.ShowWarning($"在段落 {paraIndex} 中未找到原文：{original}")
                End If
            Else
                GlobalStatusStrip.ShowWarning("缺少原文内容")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionSegment 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("校对写回异常: " & ex.Message)
        End Try
    End Sub

    ' 新增：在 Range 插入 WordProcessingML（OpenXML）片段
    Private Function InsertOpenXmlIntoRange(openXml As String, targetRange As Object) As Boolean
        Try
            If String.IsNullOrEmpty(openXml) OrElse targetRange Is Nothing Then Return False

            ' Word Range.InsertXML 需要完整的 WordProcessingML 文档结构
            ' 如果传入的只是片段（如 <w:p>），需要包装成完整结构
            Dim wrappedXml As String = WrapXmlFragment(openXml)

            Try
                Debug.Print("InsertOpenXmlIntoRange: " & wrappedXml.Substring(0, Math.Min(500, wrappedXml.Length)))
                targetRange.InsertXML(wrappedXml)
                Return True
            Catch ex As Exception
                Debug.WriteLine("InsertOpenXmlIntoRange: InsertXML 失败: " & ex.Message)
                ' 回退：尝试直接设置文本
                Try
                    Dim plainText As String = ExtractTextFromXml(openXml)
                    If Not String.IsNullOrEmpty(plainText) Then
                        targetRange.Text = plainText
                        Return True
                    End If
                Catch
                End Try
                Return False
            End Try
        Catch ex As Exception
            Debug.WriteLine("InsertOpenXmlIntoRange 出错: " & ex.Message)
            Return False
        End Try
    End Function

    ' 将 OpenXML 片段包装成完整的 WordProcessingML 文档
    Private Function WrapXmlFragment(fragment As String) As String
        If String.IsNullOrEmpty(fragment) Then Return String.Empty

        ' 检查是否已经是完整的文档结构
        If fragment.Contains("<w:document") OrElse fragment.Contains("<pkg:package") Then
            Return fragment
        End If

        ' 定义命名空间
        Const wNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        Const rNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        ' 包装成完整的 WordProcessingML 文档
        Dim sb As New StringBuilder()
        sb.Append("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
        sb.Append($"<w:document xmlns:w=""{wNs}"" xmlns:r=""{rNs}"">")
        sb.Append("<w:body>")
        sb.Append(fragment)
        sb.Append("</w:body>")
        sb.Append("</w:document>")

        Return sb.ToString()
    End Function

    ' 从 OpenXML 片段中提取纯文本（作为回退方案）
    Private Function ExtractTextFromXml(xml As String) As String
        Try
            If String.IsNullOrEmpty(xml) Then Return String.Empty
            ' 简单的正则提取 <w:t> 标签内容
            Dim matches = System.Text.RegularExpressions.Regex.Matches(xml, "<w:t[^>]*>([^<]*)</w:t>")
            Dim result As New StringBuilder()
            For Each m As System.Text.RegularExpressions.Match In matches
                If m.Groups.Count > 1 Then
                    result.Append(m.Groups(1).Value)
                End If
            Next
            Return result.ToString()
        Catch
            Return String.Empty
        End Try
    End Function

    ' applyRevision
    Protected Overrides Sub HandleApplyRevisionAll(jsonDoc As JObject)
        Try
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim newContent As String = If(jsonDoc("newContent") IsNot Nothing, jsonDoc("newContent").ToString(), String.Empty)

            If String.IsNullOrWhiteSpace(newContent) Then
                GlobalStatusStrip.ShowWarning("没有接收到写回的新内容")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("写回操作仅在 Word 环境下支持（默认实现）")
                Return
            End If

            ' 使用 GetOfficeApplicationObject 获取宿主 Word Application 对象（子类需实现）
            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try

            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象，写回失败")
                Return
            End If

            Try
                ' 在审阅模式下写回：先开启 TrackRevisions，再执行删除/插入以产生审阅记录
                Dim doc = officeApp.ActiveDocument
                Dim selRange = officeApp.Selection.Range
                Dim useRange = Nothing

                If selRange IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(selRange.Text) Then
                    useRange = selRange
                Else
                    useRange = doc.Content
                End If

                ' 开启审阅模式
                Try
                    doc.TrackRevisions = True
                Catch
                    ' 忽略，如果宿主不支持
                End Try

                ' 删除原文本（此操作会被记录为删除），然后插入新文本（被记录为插入）
                useRange.Delete()
                useRange.InsertAfter(newContent)

                GlobalStatusStrip.ShowInfo("写回已完成（审阅模式）。请在审阅中查看修改。")
            Catch ex As Exception
                Debug.WriteLine("写回失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("写回失败: " & ex.Message)
            End Try

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionAll 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("写回操作异常")
        End Try
    End Sub

    Protected Overrides Sub HandleApplyRevisionAccept(jsonDoc As JObject)
        Try
            ' 期望 { type:'applyRevisionAccept', responseUuid:..., globalIndex: n }
            Dim responseUuid As String = If(jsonDoc("responseUuid") IsNot Nothing, jsonDoc("responseUuid").ToString(), If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty))
            Dim globalIndex As Integer = If(jsonDoc("globalIndex") IsNot Nothing, CInt(jsonDoc("globalIndex")), -1)

            If globalIndex < 0 Then
                GlobalStatusStrip.ShowWarning("applyRevisionAccept: 缺少 globalIndex")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("接受单个修订仅在 Word 环境下支持（默认实现）")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try

            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象，接受修订失败")
                Return
            End If

            Try
                Dim doc = officeApp.ActiveDocument
                ' Word Revisions 集合是 1 基的；尝试保护性调用
                If globalIndex >= 1 And globalIndex <= doc.Revisions.Count Then
                    doc.Revisions(globalIndex).Accept()
                    GlobalStatusStrip.ShowInfo($"已接受修订 #{globalIndex}")
                Else
                    GlobalStatusStrip.ShowWarning("指定的修订索引超出范围或不存在")
                End If
            Catch ex As Exception
                Debug.WriteLine("接受修订失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("接受修订失败: " & ex.Message)
            End Try
        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionAccept 出错: {ex.Message}")
        End Try
    End Sub

    Protected Overrides Sub CheckAndCompleteProcessingHook(_finalUuid As String, allPlainMarkdownBuffer As StringBuilder)

        ' 如果此次会话绑定了选区信息，则发送对比预览（原文 vs AI 输出）
        Try
            ' 使用 response->request 的映射查找对应的选区信息（修正原有逻辑中使用 _finalUuid 直接查找的错误）
            Dim requestId As String = Nothing
            If _responseToRequestMap.ContainsKey(_finalUuid) Then
                requestId = _responseToRequestMap(_finalUuid)
            End If

            Dim mode As String = ""
            If _responseModeMap.ContainsKey(_finalUuid) Then
                mode = _responseModeMap(_finalUuid)
            End If

            ' 如果是排版重构动作，则触发 showComparison
            If _responseSelectionMap.ContainsKey(_finalUuid) AndAlso String.Equals(mode, "reformat", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim selInfo = _responseSelectionMap(_finalUuid)
                    Dim originalText As String = If(selInfo?.SelectedText, "")
                    Dim aiFinal As String = allPlainMarkdownBuffer.ToString()

                    Dim js As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(aiFinal)});"
                    ExecuteJavaScriptAsyncJS(js)
                Catch ex As Exception
                    Debug.WriteLine("尝试解析并发送 comparison 时出错: " & ex.Message)
                End Try
            End If

            ' 如果是审阅修订动作，解析并展示 revisions（前端会处理）
            If String.Equals(mode, "proofread", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim aiText As String = allPlainMarkdownBuffer.ToString()
                    Dim revisions As JArray = TryExtractJsonArrayFromText(aiText)
                    If revisions IsNot Nothing AndAlso revisions.Count > 0 Then
                        _revisionsMap(_finalUuid) = revisions
                        Dim jsRev As String = $"showRevisions('{_finalUuid}', {revisions.ToString(Formatting.None)});"
                        ExecuteJavaScriptAsyncJS(jsRev)
                    End If
                Catch ex As Exception
                    Debug.WriteLine("尝试解析并发送 revisions 时出错: " & ex.Message)
                End Try
            End If

            ' 解析并发送文档计划或预览 HTML 给前端，作为唯一内容
            If String.Equals(mode, "documentPlan", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(mode, "previewHtml", StringComparison.OrdinalIgnoreCase) Then
                Try
                    ' 尝试直接解析 JSON 对象（可能是 documentPlan 数组 / previewHtml / previewHtmlMap / 单个 planItem）
                    Dim rawText As String = allPlainMarkdownBuffer.ToString()
                    Dim jsonObj As JObject = TryExtractJsonObjectFromText(rawText)

                    If jsonObj IsNot Nothing Then
                        ' 如果后端/模型仅返回单个 planItem（键为 planItem），将其包装为 documentPlan 数组以便前端统一处理
                        Dim sendObj As JObject = Nothing
                        If jsonObj("planItem") IsNot Nothing Then
                            Dim arr As New JArray()
                            arr.Add(jsonObj("planItem"))
                            sendObj = New JObject()
                            sendObj("documentPlan") = arr
                            ' 若同时包含 previewHtmlMap，保留之
                            If jsonObj("previewHtmlMap") IsNot Nothing Then
                                sendObj("previewHtmlMap") = jsonObj("previewHtmlMap")
                            End If
                            ' 若 planItem 自身已包含 previewHtmlMap（极少见），合并也可按需处理
                        Else
                            ' 直接使用解析到的对象：可能含 documentPlan、previewHtml、previewHtmlMap 等
                            sendObj = jsonObj
                        End If

                        ' 获取原始选区文本（若存在）
                        Dim originalText As String = ""
                        If _responseSelectionMap.ContainsKey(_finalUuid) Then
                            Dim selInfo = _responseSelectionMap(_finalUuid)
                            originalText = If(selInfo?.SelectedText, "")
                        End If

                        ' 将整个对象序列化为字符串后传给前端的 showComparison，前端会解析 previewHtmlMap 或 documentPlan
                        Dim payload As String = sendObj.ToString(Formatting.None)
                        Dim jsPlan As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(payload)});"
                        ExecuteJavaScriptAsyncJS(jsPlan)
                    Else
                        ' 退回尝试解析为 JSON 数组（旧版可能只返回数组）
                        Dim planArr As JArray = TryExtractJsonArrayFromText(rawText)
                        If planArr IsNot Nothing AndAlso planArr.Count > 0 Then
                            Dim wrapper As New JObject()
                            wrapper("documentPlan") = planArr

                            Dim originalText As String = ""
                            If _responseSelectionMap.ContainsKey(_finalUuid) Then
                                Dim selInfo = _responseSelectionMap(_finalUuid)
                                originalText = If(selInfo?.SelectedText, "")
                            End If

                            Dim payload As String = wrapper.ToString(Formatting.None)
                            Dim jsPlan As String = $"showComparison('{_finalUuid}', {JsonConvert.SerializeObject(originalText)}, {JsonConvert.SerializeObject(payload)});"
                            ExecuteJavaScriptAsyncJS(jsPlan)
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine("处理 documentPlan/previewHtml 失败: " & ex.Message)
                End Try
            End If

        Catch ex As Exception
            Debug.WriteLine($"发送对比预览失败: {ex.Message}")
        End Try

        ' 调用基类处理续写模式
        MyBase.CheckAndCompleteProcessingHook(_finalUuid, allPlainMarkdownBuffer)
    End Sub

    ' 提取文本中第一个 JSON 数组（严格数组格式），返回 JArray 或 Nothing
    Private Function TryExtractJsonArrayFromText(text As String) As JArray
        Try
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            ' 尝试用正则抽取第一个 [...] 数组块（Singleline 允许跨行）
            Dim m As Match = Regex.Match(text, "\[.*\]", RegexOptions.Singleline)
            If m.Success Then
                Dim jsonCandidate As String = m.Value.Trim()
                ' 验证并解析为 JArray
                Try
                    Dim arr As JArray = JArray.Parse(jsonCandidate)
                    Return arr
                Catch ex As Exception
                    Debug.WriteLine("解析 JSON 数组失败: " & ex.Message)
                    Return Nothing
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("TryExtractJsonArrayFromText 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' 提取文本中第一个 JSON 对象（如 {"documentPlan":..., "previewHtml":...}），返回 JObject 或 Nothing
    Private Function TryExtractJsonObjectFromText(text As String) As JObject
        Try
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            ' 尝试用正则抽取第一个 { ... } 对象块（Singleline 允许跨行）
            Dim m As Match = Regex.Match(text, "\{[\s\S]*\}", RegexOptions.Singleline)
            If m.Success Then
                Dim jsonCandidate As String = m.Value.Trim()
                ' 验证并解析为 JObject
                Try
                    Dim obj As JObject = JObject.Parse(jsonCandidate)
                    Return obj
                Catch ex As Exception
                    Debug.WriteLine("解析 JSON 对象失败: " & ex.Message)
                    Return Nothing
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("TryExtractJsonObjectFromText 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' 排版功能（使用Word对象模型应用格式化属性）
    Protected Overrides Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
        Try
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)

            ' 期望收到字段： paraIndex, formatting (对象)
            Dim paraIndex As Integer = If(jsonDoc("paraIndex") IsNot Nothing, CInt(jsonDoc("paraIndex")), -1)
            Dim formatting As JObject = Nothing
            If jsonDoc("formatting") IsNot Nothing Then
                formatting = DirectCast(jsonDoc("formatting"), JObject)
            End If

            If paraIndex < 0 Then
                GlobalStatusStrip.ShowWarning("缺少 paraIndex 参数")
                Return
            End If

            If formatting Is Nothing Then
                GlobalStatusStrip.ShowWarning("缺少 formatting 参数")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("排版功能仅在 Word 环境下支持")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try
            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象")
                Return
            End If

            Dim doc = officeApp.ActiveDocument
            Dim selRange = officeApp.Selection.Range

            If selRange Is Nothing OrElse String.IsNullOrWhiteSpace(selRange.Text) Then
                GlobalStatusStrip.ShowWarning("请先选中需要排版的内容")
                Return
            End If

            ' 获取选中范围内的段落
            Dim paragraphs = selRange.Paragraphs
            If paraIndex >= paragraphs.Count Then
                GlobalStatusStrip.ShowWarning($"段落索引 {paraIndex} 超出范围")
                Return
            End If

            ' 定位目标段落
            Dim targetPara = paragraphs(paraIndex + 1)
            Dim targetRange = targetPara.Range

            ' 使用Word对象模型应用格式化
            Try
                ApplyFormattingToRange(targetRange, formatting)
                GlobalStatusStrip.ShowInfo($"已应用段落 {paraIndex} 的排版")
            Catch ex As Exception
                Debug.WriteLine("排版写回失败: " & ex.Message)
                GlobalStatusStrip.ShowWarning("排版写回失败: " & ex.Message)
            End Try

        Catch ex As Exception
            Debug.WriteLine("HandleApplyDocumentPlanItem 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("排版应用出错: " & ex.Message)
        End Try
    End Sub

    ' 使用Word对象模型应用格式化属性
    Private Sub ApplyFormattingToRange(targetRange As Object, formatting As JObject)
        Try
            ' 字体名称（中文）
            If formatting("fontNameCN") IsNot Nothing Then
                Dim fontNameCN As String = formatting("fontNameCN").ToString()
                If Not String.IsNullOrEmpty(fontNameCN) Then
                    Try
                        targetRange.Font.NameFarEast = fontNameCN
                    Catch
                        ' 某些 Word 版本可能不支持 NameFarEast
                    End Try
                End If
            End If

            ' 字体名称（英文/西文）
            If formatting("fontNameEN") IsNot Nothing Then
                Dim fontNameEN As String = formatting("fontNameEN").ToString()
                If Not String.IsNullOrEmpty(fontNameEN) Then
                    Try
                        targetRange.Font.Name = fontNameEN
                    Catch
                    End Try
                End If
            End If

            ' 字号
            If formatting("fontSize") IsNot Nothing Then
                Dim fontSize As Single = 0
                Single.TryParse(formatting("fontSize").ToString(), fontSize)
                If fontSize > 0 Then
                    Try
                        targetRange.Font.Size = fontSize
                    Catch
                    End Try
                End If
            End If

            ' 加粗
            If formatting("bold") IsNot Nothing Then
                Try
                    Dim bold As Boolean = formatting("bold").ToObject(Of Boolean)()
                    targetRange.Font.Bold = If(bold, -1, 0) ' Word: -1 = True, 0 = False
                Catch
                End Try
            End If

            ' 对齐方式
            If formatting("alignment") IsNot Nothing Then
                Dim alignment As String = formatting("alignment").ToString().ToLower()
                Try
                    Select Case alignment
                        Case "left"
                            targetRange.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                        Case "center"
                            targetRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                        Case "right"
                            targetRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
                        Case "justify", "justified"
                            targetRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
                    End Select
                Catch
                End Try
            End If

            ' 首行缩进（字符数）
            If formatting("firstLineIndent") IsNot Nothing Then
                Dim indent As Single = 0
                Single.TryParse(formatting("firstLineIndent").ToString(), indent)
                If indent > 0 Then
                    Try
                        ' CharacterUnitFirstLineIndent 以字符为单位
                        targetRange.ParagraphFormat.CharacterUnitFirstLineIndent = indent
                    Catch
                        ' 回退：使用磅值（1字符约=10.5磅 for 五号字）
                        Try
                            targetRange.ParagraphFormat.FirstLineIndent = indent * 10.5
                        Catch
                        End Try
                    End Try
                End If
            End If

            ' 行距
            If formatting("lineSpacing") IsNot Nothing Then
                Dim lineSpacing As Single = 0
                Single.TryParse(formatting("lineSpacing").ToString(), lineSpacing)
                If lineSpacing > 0 Then
                    Try
                        ' LineSpacingRule: 0=wdLineSpaceSingle, 1=wdLineSpace1pt5, 2=wdLineSpaceDouble, 5=wdLineSpaceMultiple
                        If lineSpacing = 1.0 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 0 ' wdLineSpaceSingle
                        ElseIf lineSpacing = 1.5 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 1 ' wdLineSpace1pt5
                        ElseIf lineSpacing = 2.0 Then
                            targetRange.ParagraphFormat.LineSpacingRule = 2 ' wdLineSpaceDouble
                        Else
                            ' 使用多倍行距
                            targetRange.ParagraphFormat.LineSpacingRule = 5 ' wdLineSpaceMultiple
                            targetRange.ParagraphFormat.LineSpacing = 12 * lineSpacing ' 12磅 * 倍数
                        End If
                    Catch
                    End Try
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine("ApplyFormattingToRange 出错: " & ex.Message)
            Throw
        End Try
    End Sub

    ' 辅助：由纯文本生成最简单的 WordProcessingML OpenXML 片段（每个换行生成一个段落）
    Private Function BuildOpenXmlFromText(text As String) As String
        Try
            If String.IsNullOrEmpty(text) Then Return String.Empty
            Dim ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            Dim sb As New StringBuilder()
            sb.Append($"<w:document xmlns:w=""{ns}""><w:body>")
            Dim lines = text.Replace(vbCrLf, vbLf).Split(New Char() {vbLf})
            For Each line In lines
                Dim escaped = line
                escaped = escaped.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                ' 保留前后空格
                sb.Append($"<w:p><w:r><w:t xml:space=""preserve"">{escaped}</w:t></w:r></w:p>")
            Next
            sb.Append("</w:body></w:document>")
            Return sb.ToString()
        Catch ex As Exception
            Debug.WriteLine("BuildOpenXmlFromText 出错: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    Protected Overrides Function CaptureCurrentSelectionInfo(mode As String) As SelectionInfo
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            Dim txt As String = If(sel IsNot Nothing AndAlso sel.Range IsNot Nothing, sel.Range.Text, String.Empty)
            If String.IsNullOrEmpty(txt) Then
                Return Nothing
            End If

            Dim info As New SelectionInfo()
            info.SelectedText = txt
            info.DocumentPath = Globals.ThisAddIn.Application.ActiveDocument.FullName

            Try
                info.StartPos = sel.Range.Start
                info.EndPos = sel.Range.End
            Catch
                info.StartPos = 0
                info.EndPos = 0
            End Try

            Return info
        Catch
            Return Nothing
        End Try
    End Function

    ' ========== 续写功能 ==========

    Private _continuationService As WordContinuationService
    Private _cachedContinuationContext As ContinuationContext ' 缓存续写上下文，用于多轮续写

    ''' <summary>
    ''' 触发续写 - 获取光标上下文并发送AI请求
    ''' </summary>
    Protected Overrides Sub HandleTriggerContinuation(jsonDoc As JObject)
        Try
            ' 提取参数
            Dim style As String = ""
            Dim isContinuationMode As Boolean = False

            If jsonDoc IsNot Nothing Then
                If jsonDoc("style") IsNot Nothing Then
                    style = jsonDoc("style").ToString()
                End If
                If jsonDoc("isContinuationMode") IsNot Nothing Then
                    isContinuationMode = jsonDoc("isContinuationMode").ToObject(Of Boolean)()
                End If
            End If

            ' 初始化续写服务
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 检查是否可以续写
            If Not _continuationService.CanContinue() Then
                GlobalStatusStrip.ShowWarning("无法获取文档信息，请确保文档已打开")
                Return
            End If

            Dim context As ContinuationContext

            ' 如果是续写模式的后续请求，并且有缓存的上下文，则复用
            If isContinuationMode AndAlso _cachedContinuationContext IsNot Nothing Then
                ' 多轮续写：使用缓存的上下文，但style作为新的调整要求
                context = _cachedContinuationContext
                GlobalStatusStrip.ShowInfo("继续续写...")
            Else
                ' 首次续写或非续写模式：重新获取上下文
                context = _continuationService.GetCursorContext(3, 3)
                If context Is Nothing Then
                    GlobalStatusStrip.ShowWarning("无法获取文档上下文")
                    Return
                End If
                ' 缓存上下文
                _cachedContinuationContext = context
                GlobalStatusStrip.ShowInfo("正在分析上下文并生成续写内容...")
            End If

            ' 发送续写请求（带上风格参数）
            SendContinuationRequest(context, style)

        Catch ex As Exception
            Debug.WriteLine($"HandleTriggerContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"触发续写时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用续写结果到Word文档
    ''' </summary>
    Protected Overrides Sub HandleApplyContinuation(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("续写内容为空")
                Return
            End If

            ' 确保续写服务已初始化
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor
            End Select

            ' 插入续写内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("续写内容已插入文档")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeContinuationActions('{uuid}');")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyContinuation 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"插入续写内容时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用模板渲染结果到Word文档
    ''' </summary>
    Protected Overrides Sub HandleApplyTemplateContent(jsonDoc As JObject)
        Try
            Dim content As String = If(jsonDoc("content") IsNot Nothing, jsonDoc("content").ToString(), String.Empty)
            Dim positionStr As String = If(jsonDoc("position") IsNot Nothing, jsonDoc("position").ToString(), "current")

            If String.IsNullOrWhiteSpace(content) Then
                GlobalStatusStrip.ShowWarning("模板内容为空")
                Return
            End If

            ' 确保续写服务已初始化（复用其插入逻辑）
            If _continuationService Is Nothing Then
                _continuationService = New WordContinuationService(Globals.ThisAddIn.Application)
            End If

            ' 根据position参数确定插入位置
            Dim insertPos As ShareRibbon.InsertPosition
            Select Case positionStr.ToLower()
                Case "start"
                    insertPos = ShareRibbon.InsertPosition.DocumentStart
                Case "end"
                    insertPos = ShareRibbon.InsertPosition.DocumentEnd
                Case Else ' "current" 或默认
                    insertPos = ShareRibbon.InsertPosition.AtCursor
            End Select

            ' 插入模板内容
            _continuationService.InsertContinuation(content, insertPos)

            GlobalStatusStrip.ShowInfo("模板内容已插入文档")

            ' 通知前端移除操作按钮
            Dim uuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            If Not String.IsNullOrEmpty(uuid) Then
                ExecuteJavaScriptAsyncJS($"removeTemplateActions('{uuid}');")
            End If

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyTemplateContent 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning($"插入模板内容时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取当前Word上下文快照（用于自动补全）
    ''' </summary>
    Protected Overrides Function GetContextSnapshot() As JObject
        Dim snapshot As New JObject()
        snapshot("appType") = "Word"

        Try
            Dim selection = Globals.ThisAddIn.Application.Selection
            If selection IsNot Nothing AndAlso selection.Start <> selection.End Then
                ' 有选中内容
                Dim selText = selection.Text
                If Not String.IsNullOrEmpty(selText) AndAlso selText.Length > 500 Then
                    selText = selText.Substring(0, 500) & "..."
                End If
                snapshot("selection") = If(selText, "")
            Else
                snapshot("selection") = ""
            End If

            ' 获取文档标题
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            If doc IsNot Nothing Then
                snapshot("documentName") = If(doc.Name, "")
            End If

        Catch ex As Exception
            Debug.WriteLine($"GetContextSnapshot 出错: {ex.Message}")
        End Try

        Return snapshot
    End Function

    ''' <summary>
    ''' 重写保存设置方法，同步更新Word补全管理器状态
    ''' </summary>
    Protected Overrides Sub HandleSaveSettings(jsonDoc As JObject)
        MyBase.HandleSaveSettings(jsonDoc)
        
        ' 同步更新Word补全管理器的启用状态
        Try
            Dim enableAutocomplete As Boolean = If(jsonDoc("enableAutocomplete")?.Value(Of Boolean)(), False)
            WordCompletionManager.Instance.Enabled = enableAutocomplete
            Debug.WriteLine($"[WordChatControl] 补全设置已同步: Enabled={enableAutocomplete}")
        Catch ex As Exception
            Debug.WriteLine($"[WordChatControl] 同步补全设置失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 执行JSON命令（重写基类方法）- 带严格验证
    ''' </summary>
    Protected Overrides Function ExecuteJsonCommand(jsonCode As String, preview As Boolean) As Boolean
        Try
            ' 使用严格的结构验证
            Dim errorMessage As String = ""
            Dim normalizedJson As JToken = Nothing
            
            If Not WordJsonCommandSchema.ValidateJsonStructure(jsonCode, errorMessage, normalizedJson) Then
                ' 格式验证失败
                Debug.WriteLine($"Word JSON格式验证失败: {errorMessage}")
                Debug.WriteLine($"原始JSON: {jsonCode.Substring(0, Math.Min(200, jsonCode.Length))}...")
                
                ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式不符合规范: {errorMessage}")
                Return False
            End If
            
            ' 验证通过，根据类型执行
            If normalizedJson.Type = JTokenType.Object Then
                Dim jsonObj = CType(normalizedJson, JObject)
                
                ' 命令数组格式
                If jsonObj("commands") IsNot Nothing Then
                    Return ExecuteWordCommandsArray(jsonObj("commands"), jsonCode, preview)
                End If
                
                ' 单命令格式
                Return ExecuteWordSingleCommand(jsonObj, jsonCode, preview)
            End If
            
            ShareRibbon.GlobalStatusStrip.ShowWarning("无效的JSON格式")
            Return False

        Catch ex As Newtonsoft.Json.JsonReaderException
            ShareRibbon.GlobalStatusStrip.ShowWarning($"JSON格式无效: {ex.Message}")
            Return False
        Catch ex As Exception
            ShareRibbon.GlobalStatusStrip.ShowWarning($"执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行Word命令数组
    ''' </summary>
    Private Function ExecuteWordCommandsArray(commandsArray As JToken, originalJson As String, preview As Boolean) As Boolean
        Try
            Dim commands = CType(commandsArray, JArray)
            If commands.Count = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowWarning("命令数组为空")
                Return False
            End If

            ' 预览所有命令
            If preview Then
                Dim previewMsg As New StringBuilder()
                previewMsg.AppendLine($"即将执行 {commands.Count} 个Word命令:")
                previewMsg.AppendLine()

                Dim cmdIndex = 1
                For Each cmd In commands
                    If cmd.Type = JTokenType.Object Then
                        Dim cmdObj = CType(cmd, JObject)
                        Dim cmdName = cmdObj("command")?.ToString()
                        Dim content = cmdObj("params")?("content")?.ToString()
                        
                        previewMsg.AppendLine($"{cmdIndex}. {cmdName}")
                        If Not String.IsNullOrEmpty(content) Then
                            previewMsg.AppendLine($"   内容: {content.Substring(0, Math.Min(50, content.Length))}...")
                        End If
                        previewMsg.AppendLine()
                        cmdIndex += 1
                    End If
                Next

                previewMsg.AppendLine("是否继续执行？")

                If MessageBox.Show(previewMsg.ToString(), "Word批量命令预览", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) <> DialogResult.OK Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 执行所有命令
            Dim successCount = 0
            Dim failCount = 0

            For Each cmd In commands
                If cmd.Type = JTokenType.Object Then
                    Dim cmdObj = CType(cmd, JObject)
                    If ExecuteWordCommand(cmdObj) Then
                        successCount += 1
                    Else
                        failCount += 1
                    End If
                End If
            Next

            If failCount = 0 Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"所有 {successCount} 个命令执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"执行完成: {successCount} 成功, {failCount} 失败")
            End If

            Return failCount = 0

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommandsArray 出错: {ex.Message}")
            ShareRibbon.GlobalStatusStrip.ShowWarning($"批量执行失败: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行单个Word命令
    ''' </summary>
    Private Function ExecuteWordSingleCommand(commandJson As JObject, processedJson As String, preview As Boolean) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            
            ' 预览
            If preview Then
                Dim params = commandJson("params")
                Dim content = params?("content")?.ToString()

                Dim previewMsg = $"即将执行 Word 命令:{vbCrLf}{vbCrLf}" &
                                $"命令: {command}{vbCrLf}" &
                                If(Not String.IsNullOrEmpty(content), $"内容: {content.Substring(0, Math.Min(100, content.Length))}...{vbCrLf}", "") &
                                $"{vbCrLf}是否继续执行？"

                If MessageBox.Show(previewMsg, "Word命令预览", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) <> DialogResult.OK Then
                    ExecuteJavaScriptAsyncJS("handleExecutionCancelled('')")
                    Return True
                End If
            End If

            ' 执行命令
            Dim success = ExecuteWordCommand(commandJson)

            If success Then
                ShareRibbon.GlobalStatusStrip.ShowInfo($"命令 '{command}' 执行成功")
            Else
                ShareRibbon.GlobalStatusStrip.ShowWarning($"命令 '{command}' 执行失败")
            End If

            Return success

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordSingleCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 执行具体的Word命令
    ''' </summary>
    Private Function ExecuteWordCommand(commandJson As JObject) As Boolean
        Try
            Dim command = commandJson("command")?.ToString()
            Dim params = commandJson("params")
            
            Dim doc = Globals.ThisAddIn.Application.ActiveDocument
            Dim selection = Globals.ThisAddIn.Application.Selection

            Select Case command.ToLower()
                Case "inserttext"
                    Return ExecuteInsertText(params, selection)
                Case "formattext"
                    Return ExecuteFormatText(params, selection)
                Case "replacetext"
                    Return ExecuteReplaceText(params, doc)
                Case "inserttable"
                    Return ExecuteInsertTable(params, selection)
                Case "applystyle"
                    Return ExecuteApplyStyle(params, selection)
                Case "generatetoc"
                    Return ExecuteGenerateTOC(params, doc)
                Case "beautifydocument"
                    Return ExecuteBeautifyDocument(params, doc)
                Case Else
                    Debug.WriteLine($"不支持的Word命令: {command}")
                    Return False
            End Select

        Catch ex As Exception
            Debug.WriteLine($"ExecuteWordCommand 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertText(params As JToken, selection As Object) As Boolean
        Try
            Dim content = params("content")?.ToString()
            Dim position = If(params("position")?.ToString(), "cursor")

            Select Case position.ToLower()
                Case "start"
                    Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0).InsertBefore(content)
                Case "end"
                    Globals.ThisAddIn.Application.ActiveDocument.Content.InsertAfter(content)
                Case Else ' cursor
                    selection.TypeText(content)
            End Select

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteFormatText(params As JToken, selection As Object) As Boolean
        Try
            Dim range = selection.Range

            If params("bold") IsNot Nothing Then
                range.Font.Bold = If(params("bold").Value(Of Boolean)(), -1, 0)
            End If

            If params("italic") IsNot Nothing Then
                range.Font.Italic = If(params("italic").Value(Of Boolean)(), -1, 0)
            End If

            If params("underline") IsNot Nothing Then
                range.Font.Underline = If(params("underline").Value(Of Boolean)(), 1, 0)
            End If

            If params("fontSize") IsNot Nothing Then
                range.Font.Size = params("fontSize").Value(Of Single)()
            End If

            If params("fontName") IsNot Nothing Then
                range.Font.Name = params("fontName").ToString()
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteFormatText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteReplaceText(params As JToken, doc As Object) As Boolean
        Try
            Dim find = params("find")?.ToString()
            Dim replace = If(params("replace")?.ToString(), "")
            Dim matchCase = If(params("matchCase")?.Value(Of Boolean)(), False)

            Dim findObj = doc.Content.Find
            findObj.ClearFormatting()
            findObj.Replacement.ClearFormatting()
            findObj.Text = find
            findObj.Replacement.Text = replace
            findObj.Forward = True
            findObj.Wrap = 1 ' wdFindContinue
            findObj.MatchCase = matchCase
            findObj.Execute(Replace:=2) ' wdReplaceAll

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteReplaceText 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteInsertTable(params As JToken, selection As Object) As Boolean
        Try
            Dim rows = params("rows")?.Value(Of Integer)()
            Dim cols = params("cols")?.Value(Of Integer)()

            If rows <= 0 OrElse cols <= 0 Then Return False

            Dim table = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
                selection.Range, rows, cols)

            ' 如果有data，填充表格
            Dim data = params("data")
            If data IsNot Nothing AndAlso data.Type = JTokenType.Array Then
                Dim dataArr = CType(data, JArray)
                Dim x As Integer = dataArr.Count - 1
                Dim x2 As Integer = rows - 1
                For rowIdx = 0 To Math.Min(x, x2)
                    Dim rowData = dataArr(rowIdx)
                    If rowData.Type = JTokenType.Array Then
                        Dim rowArr = CType(rowData, JArray)
                        Dim y As Integer = rowArr.Count - 1
                        Dim y1 As Integer = cols - 1
                        For colIdx = 0 To Math.Min(y, y1)
                            table.Cell(rowIdx + 1, colIdx + 1).Range.Text = rowArr(colIdx).ToString()
                        Next
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteInsertTable 出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function ExecuteApplyStyle(params As JToken, selection As Object) As Boolean
        Try
            Dim styleName = params("styleName")?.ToString()
            If String.IsNullOrEmpty(styleName) Then Return False

            selection.Style = styleName
            Return True
        Catch ex As Exception
            Debug.WriteLine($"ExecuteApplyStyle 出错: {ex.Message}")
            Return False
        End Try
    End Function

#Region "高级Word命令实现"

    ''' <summary>
    ''' 生成目录
    ''' </summary>
    Private Function ExecuteGenerateTOC(params As JToken, doc As Object) As Boolean
        Try
            Dim position = If(params("position")?.ToString(), "start")
            Dim levels = If(params("levels")?.Value(Of Integer)(), 3)
            Dim includePageNumbers = If(params("includePageNumbers")?.Value(Of Boolean)(), True)

            ' 确定插入位置
            Dim range As Object
            If position.ToLower() = "start" Then
                range = doc.Range(0, 0)
            Else
                range = Globals.ThisAddIn.Application.Selection.Range
            End If

            ' 删除已有目录
            For Each toc In doc.TablesOfContents
                toc.Delete()
            Next

            ' 插入新目录
            Dim newToc = doc.TablesOfContents.Add(
                Range:=range,
                UseHeadingStyles:=True,
                UpperHeadingLevel:=1,
                LowerHeadingLevel:=levels,
                IncludePageNumbers:=includePageNumbers
            )

            ' 更新目录
            newToc.Update()

            ShareRibbon.GlobalStatusStrip.ShowInfo($"已生成{levels}级目录")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteGenerateTOC 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 文档美化
    ''' </summary>
    Private Function ExecuteBeautifyDocument(params As JToken, doc As Object) As Boolean
        Try
            Dim theme = params("theme")
            Dim margins = params("margins")

            ' 应用页边距
            If margins IsNot Nothing Then
                ApplyMargins(doc, margins)
            End If

            ' 应用主题样式
            If theme IsNot Nothing Then
                ApplyThemeStyles(doc, theme)
            End If

            ShareRibbon.GlobalStatusStrip.ShowInfo("文档美化完成")
            Return True

        Catch ex As Exception
            Debug.WriteLine($"ExecuteBeautifyDocument 出错: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 应用页边距
    ''' </summary>
    Private Sub ApplyMargins(doc As Object, margins As JToken)
        Try
            Dim pageSetup = doc.PageSetup
            
            ' 单位转换: 厘米 -> 磅 (1cm = 28.35磅)
            Const cmToPoints As Single = 28.35F

            If margins("top") IsNot Nothing Then
                pageSetup.TopMargin = margins("top").Value(Of Single)() * cmToPoints
            End If
            If margins("bottom") IsNot Nothing Then
                pageSetup.BottomMargin = margins("bottom").Value(Of Single)() * cmToPoints
            End If
            If margins("left") IsNot Nothing Then
                pageSetup.LeftMargin = margins("left").Value(Of Single)() * cmToPoints
            End If
            If margins("right") IsNot Nothing Then
                pageSetup.RightMargin = margins("right").Value(Of Single)() * cmToPoints
            End If
        Catch ex As Exception
            Debug.WriteLine($"ApplyMargins 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用主题样式到文档
    ''' </summary>
    Private Sub ApplyThemeStyles(doc As Object, theme As JToken)
        Try
            ' 应用标题1样式
            Dim h1Theme = theme("h1")
            If h1Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 1", h1Theme)
            End If

            ' 应用标题2样式
            Dim h2Theme = theme("h2")
            If h2Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 2", h2Theme)
            End If

            ' 应用标题3样式
            Dim h3Theme = theme("h3")
            If h3Theme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "标题 3", h3Theme)
            End If

            ' 应用正文样式
            Dim bodyTheme = theme("body")
            If bodyTheme IsNot Nothing Then
                ApplyStyleFromTheme(doc, "正文", bodyTheme)
                
                ' 应用行间距到所有段落
                If bodyTheme("lineSpacing") IsNot Nothing Then
                    Dim lineSpacing = bodyTheme("lineSpacing").Value(Of Single)()
                    For Each para In doc.Paragraphs
                        Try
                            para.LineSpacingRule = 5 ' wdLineSpaceMultiple
                            para.LineSpacing = 12 * lineSpacing ' 12磅 * 倍数
                        Catch
                        End Try
                    Next
                End If
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyThemeStyles 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用样式设置
    ''' </summary>
    Private Sub ApplyStyleFromTheme(doc As Object, styleName As String, themeSettings As JToken)
        Try
            Dim style = doc.Styles(styleName)
            
            If themeSettings("font") IsNot Nothing Then
                style.Font.Name = themeSettings("font").ToString()
            End If
            If themeSettings("size") IsNot Nothing Then
                style.Font.Size = themeSettings("size").Value(Of Single)()
            End If
            If themeSettings("color") IsNot Nothing Then
                Dim colorStr = themeSettings("color").ToString()
                Dim color = System.Drawing.ColorTranslator.FromHtml(colorStr)
                style.Font.Color = System.Drawing.ColorTranslator.ToOle(color)
            End If
            If themeSettings("bold") IsNot Nothing Then
                style.Font.Bold = If(themeSettings("bold").Value(Of Boolean)(), -1, 0)
            End If
            If themeSettings("italic") IsNot Nothing Then
                style.Font.Italic = If(themeSettings("italic").Value(Of Boolean)(), -1, 0)
            End If

        Catch ex As Exception
            Debug.WriteLine($"ApplyStyleFromTheme ({styleName}) 出错: {ex.Message}")
        End Try
    End Sub

#End Region

End Class
