Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
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


    ' 修订、审阅功能（由前端触发）
    Protected Overrides Sub HandleApplyRevisionSegment(jsonDoc As JObject)
        Debug.Print(123)
        Try
            ' 期望收到字段： uuid, index, action, matchText, contextBefore, contextAfter, replaceWith, start(optional)
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)
            Dim indexVal As Integer = If(jsonDoc("index") IsNot Nothing, CInt(jsonDoc("index")), -1)
            Dim action As String = If(jsonDoc("action") IsNot Nothing, jsonDoc("action").ToString().ToLower(), String.Empty)
            Dim matchText As String = If(jsonDoc("matchText") IsNot Nothing, jsonDoc("matchText").ToString(), String.Empty)
            Dim contextBefore As String = If(jsonDoc("contextBefore") IsNot Nothing, jsonDoc("contextBefore").ToString(), String.Empty)
            Dim contextAfter As String = If(jsonDoc("contextAfter") IsNot Nothing, jsonDoc("contextAfter").ToString(), String.Empty)
            Dim replaceWith As String = If(jsonDoc("replaceWith") IsNot Nothing, jsonDoc("replaceWith").ToString(), String.Empty)
            Dim startIdx As Integer = -1
            If jsonDoc("start") IsNot Nothing Then Integer.TryParse(jsonDoc("start").ToString(), startIdx)

            If indexVal < 0 Then
                GlobalStatusStrip.ShowWarning("缺少 index 参数")
                Return
            End If
            If String.IsNullOrWhiteSpace(action) Then
                GlobalStatusStrip.ShowWarning("缺少 action 参数")
                Return
            End If

            Dim appInfo As ApplicationInfo = GetApplication()
            If appInfo Is Nothing OrElse appInfo.Type <> OfficeApplicationType.Word Then
                GlobalStatusStrip.ShowWarning("逐段写回仅在 Word 环境下支持（默认实现）")
                Return
            End If

            Dim officeApp As Object = Nothing
            Try
                officeApp = GetOfficeApplicationObject()
            Catch ex As Exception
                Debug.WriteLine("获取 Office 应用对象失败: " & ex.Message)
            End Try
            If officeApp Is Nothing Then
                GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象，逐段写回失败")
                Return
            End If

            Dim doc = officeApp.ActiveDocument
            Dim selRange = officeApp.Selection.Range
            Dim useRange = If(selRange IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(selRange.Text), selRange, doc.Content)
            Dim fullText As String = useRange.Text

            ' 使用类级函数定位 Range（优先 match+context，然后 start，最后模糊匹配）
            Dim targetRange As Object = FindRangeByAnchor(doc, useRange, fullText, matchText, contextBefore, contextAfter, startIdx)
            If targetRange Is Nothing Then
                GlobalStatusStrip.ShowWarning($"未能定位到第 {indexVal} 项的目标文本，操作已取消。")
                Return
            End If

            ' 开启审阅模式再执行（以产生修订）
            Try
                doc.TrackRevisions = True
            Catch
            End Try
            Debug.Print(action)
            Select Case action
                Case "replace"
                    Try
                        targetRange.Text = replaceWith
                        GlobalStatusStrip.ShowInfo($"已替换第 {indexVal} 项（审阅模式）")
                    Catch ex As Exception
                        Debug.WriteLine("替换失败: " & ex.Message)
                        GlobalStatusStrip.ShowWarning("替换失败: " & ex.Message)
                    End Try
                Case "insert"
                    Try
                        ' 插入到目标位置之前（若 matchText 为空且提供 start 可基于 start 插入）
                        Dim insertRange = targetRange
                        If String.IsNullOrEmpty(matchText) AndAlso startIdx >= 0 Then
                            insertRange = doc.Range(useRange.Start + startIdx, useRange.Start + startIdx)
                        End If
                        insertRange.InsertBefore(replaceWith)
                        GlobalStatusStrip.ShowInfo($"已插入第 {indexVal} 项（审阅模式）")
                    Catch ex As Exception
                        Debug.WriteLine("插入失败: " & ex.Message)
                        GlobalStatusStrip.ShowWarning("插入失败: " & ex.Message)
                    End Try
                Case "delete"
                    Try
                        targetRange.Delete()
                        GlobalStatusStrip.ShowInfo($"已删除第 {indexVal} 项（审阅模式）")
                    Catch ex As Exception
                        Debug.WriteLine("删除失败: " & ex.Message)
                        GlobalStatusStrip.ShowWarning("删除失败: " & ex.Message)
                    End Try
                Case Else
                    GlobalStatusStrip.ShowWarning("未知 action 类型: " & action)
            End Select

        Catch ex As Exception
            Debug.WriteLine($"HandleApplyRevisionSegment 出错: {ex.Message}")
            GlobalStatusStrip.ShowWarning("逐段写回异常: " & ex.Message)
        End Try
    End Sub

    ' 新增：在 Range 插入 WordProcessingML（OpenXML）片段
    Private Function InsertOpenXmlIntoRange(openXml As String, targetRange As Object) As Boolean
        Try
            If String.IsNullOrEmpty(openXml) OrElse targetRange Is Nothing Then Return False
            ' Word Range 有 InsertXML 方法，可直接插入 WordProcessingML 片段
            Try
                Debug.Print(openXml)
                targetRange.InsertXML(openXml)
                Return True
            Catch ex As Exception
                Debug.WriteLine("InsertOpenXmlIntoRange: InsertXML 失败: " & ex.Message)
                Return False
            End Try
        Catch ex As Exception
            Debug.WriteLine("InsertOpenXmlIntoRange 出错: " & ex.Message)
            Return False
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

    ' 排版重构新增：后端逐项应用 documentPlan 项（增强版）
    Protected Overrides Sub HandleApplyDocumentPlanItem(jsonDoc As JObject)
        Try
            ' 支持传单项 planItem 或整个 plan 数组（plan/documentPlan）
            Dim responseUuid As String = If(jsonDoc("uuid") IsNot Nothing, jsonDoc("uuid").ToString(), String.Empty)

            ' 单项处理（兼容旧接口）
            Dim itemObj As JObject = Nothing
            If jsonDoc("planItem") IsNot Nothing Then
                itemObj = DirectCast(jsonDoc("planItem"), JObject)
            ElseIf jsonDoc("planItemId") IsNot Nothing AndAlso _revisionsMap.ContainsKey(responseUuid) Then
                Dim arr = _revisionsMap(responseUuid)
                Dim pid = jsonDoc("planItemId").ToString()
                For Each it In arr
                    If it("id") IsNot Nothing AndAlso it("id").ToString() = pid Then
                        itemObj = DirectCast(it, JObject)
                        Exit For
                    End If
                Next
            End If

            If itemObj Is Nothing Then
                GlobalStatusStrip.ShowWarning("未提供可应用的 planItem")
                Return
            End If

            ' 应用单项
            ApplySinglePlanItem(responseUuid, itemObj)

        Catch ex As Exception
            Debug.WriteLine("HandleApplyDocumentPlanItem 错误: " & ex.Message)
            GlobalStatusStrip.ShowWarning("应用 plan 项出错: " & ex.Message)
        End Try
    End Sub

    ' 批量应用整个 documentPlan（按序执行）
    Private Sub ApplyDocumentPlanItems(responseUuid As String, planArr As JArray)
        If planArr Is Nothing Then Return
        ' 确保在 UI 线程上执行 Word 操作
        For Each it In planArr
            Try
                ApplySinglePlanItem(responseUuid, DirectCast(it, JObject))
                Threading.Thread.Sleep(120) ' 简短延迟避免太快
            Catch ex As Exception
                Debug.WriteLine("ApplyDocumentPlanItems 单项失败: " & ex.Message)
            End Try
        Next
        GlobalStatusStrip.ShowInfo("已尝试应用所有 documentPlan 项")
    End Sub


    ' 应用单个 planItem（增强版，支持 attributes.openXmlActions 的通用调度）
    Private Sub ApplySinglePlanItem(responseUuid As String, itemObj As JObject)
        If itemObj Is Nothing Then Return

        Dim blockId As String = If(itemObj("blockId") IsNot Nothing, itemObj("blockId").ToString(), String.Empty)
        Dim action As String = If(itemObj("action") IsNot Nothing, itemObj("action").ToString().ToLower(), String.Empty)
        Dim attributes As JObject = If(itemObj("attributes") IsNot Nothing, DirectCast(itemObj("attributes"), JObject), Nothing)
        Dim textContent As String = If(attributes IsNot Nothing AndAlso attributes("text") IsNot Nothing, attributes("text").ToString(), "")
        Dim itemWordOpenXml As String = If(itemObj("wordOpenXml") IsNot Nothing, itemObj("wordOpenXml").ToString(), If(attributes IsNot Nothing AndAlso attributes("wordOpenXml") IsNot Nothing, attributes("wordOpenXml").ToString(), ""))

        Dim officeApp = GetOfficeApplicationObject()
        If officeApp Is Nothing Then
            GlobalStatusStrip.ShowWarning("无法获取 Word 应用对象")
            Return
        End If
        Dim doc = officeApp.ActiveDocument
        Dim selRange = officeApp.Selection.Range
        Dim useRange = If(selRange IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(selRange.Text), selRange, doc.Content)

        Try
            ' 定位目标块（按 ReformatButton_Click 的分块逻辑重建 mapping）
            Dim targetRange As Object = Nothing
            If Not String.IsNullOrEmpty(blockId) Then
                targetRange = FindRangeByBlockId(doc, useRange, blockId)
            End If

            ' 未找到目标块则退化处理：replaceText -> 替换 useRange；insert -> 文档末尾插入；其他则末尾插入
            If targetRange Is Nothing Then
                If String.Equals(action, "replacetext", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(action, "replace", StringComparison.OrdinalIgnoreCase) Then
                    targetRange = useRange
                Else
                    targetRange = doc.Content
                    targetRange.Collapse(0) ' wdCollapseEnd = 0
                End If
            End If

            ' 如果提供了 attributes.openXmlActions，则优先以通用 openXml 调度器执行（由模型输出 verb 指令）
            If attributes IsNot Nothing AndAlso attributes("openXmlActions") IsNot Nothing Then
                Dim oa = attributes("openXmlActions")
                Try
                    If HandleOpenXmlActions(oa, doc, targetRange, blockId) Then
                        GlobalStatusStrip.ShowInfo($"已通过 openXmlActions 应用 block '{blockId}'")
                        Return
                    End If
                Catch ex As Exception
                    Debug.WriteLine("HandleOpenXmlActions 异常: " & ex.Message)
                    ' 继续回退其他策略
                End Try
            End If

            ' 检测目标是否为复杂对象（表格/图片/公式），某些动作需谨慎处理
            Dim isTable As Boolean = False
            Dim isImage As Boolean = False
            Dim isEquation As Boolean = False
            Try
                isTable = (targetRange.Tables IsNot Nothing AndAlso targetRange.Tables.Count > 0)
            Catch : End Try
            Try
                isImage = ((targetRange.InlineShapes IsNot Nothing AndAlso targetRange.InlineShapes.Count > 0) OrElse (targetRange.ShapeRange IsNot Nothing AndAlso targetRange.ShapeRange.Count > 0))
            Catch : End Try
            Try
                isEquation = (targetRange.OMaths IsNot Nothing AndAlso targetRange.OMaths.Count > 0)
            Catch : End Try

            Select Case action
                Case "format"
                    ' 只对文本段落进行格式化，避免破坏表格/图片/公式
                    If isTable OrElse isImage OrElse isEquation Then
                        GlobalStatusStrip.ShowWarning($"跳过格式化块 '{blockId}'：目标块为表格/图片/公式，避免破坏原生对象。")
                        Return
                    End If

                    Try
                        If attributes IsNot Nothing Then
                            ' 字体大小
                            If attributes("fontSize") IsNot Nothing Then
                                Dim fs As Integer = 0
                                Integer.TryParse(attributes("fontSize").ToString(), fs)
                                If fs > 0 Then targetRange.Font.Size = fs
                            End If
                            ' 粗体/斜体/下划线 等
                            If attributes("bold") IsNot Nothing Then
                                Dim b As Boolean = attributes("bold").ToObject(Of Boolean)()
                                targetRange.Font.Bold = b
                            End If
                            If attributes("italic") IsNot Nothing Then
                                Dim it As Boolean = attributes("italic").ToObject(Of Boolean)()
                                targetRange.Font.Italic = it
                            End If
                            If attributes("underline") IsNot Nothing Then
                                Dim u As Boolean = attributes("underline").ToObject(Of Boolean)()
                                targetRange.Font.Underline = If(u, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse)
                            End If
                            ' 对齐方式： left/center/right/justify
                            If attributes("alignment") IsNot Nothing Then
                                Dim a As String = attributes("alignment").ToString().ToLower()
                                Select Case a
                                    Case "left"
                                        targetRange.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                                    Case "center"
                                        targetRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                                    Case "right"
                                        targetRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
                                    Case "justify", "justified"
                                        targetRange.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
                                End Select
                            End If
                            ' 字体名称
                            If attributes("fontName") IsNot Nothing Then
                                Dim fn As String = attributes("fontName").ToString()
                                If Not String.IsNullOrEmpty(fn) Then targetRange.Font.Name = fn
                            End If
                        End If
                        GlobalStatusStrip.ShowInfo($"已对 block '{blockId}' 应用格式设置")
                    Catch ex As Exception
                        Debug.WriteLine("Format 处理失败: " & ex.Message)
                        GlobalStatusStrip.ShowWarning("格式化操作失败: " & ex.Message)
                    End Try

                Case "replacetext", "replace"
                    ' 优先使用提供的 wordOpenXml（保持表格/公式/格式）
                    'If Not String.IsNullOrEmpty(itemWordOpenXml) Then
                    '    Try
                    '        ' 清除原块内容（若是占位或段落），然后插入 OpenXML
                    '        Try
                    '            targetRange.Text = ""
                    '        Catch : End Try
                    '        If InsertOpenXmlIntoRange(itemWordOpenXml, targetRange) Then
                    '            GlobalStatusStrip.ShowInfo($"已通过 OpenXML 回写 block '{blockId}'")
                    '            Return
                    '        End If
                    '    Catch ex As Exception
                    '        Debug.WriteLine("使用 itemWordOpenXml 回写失败: " & ex.Message)
                    '    End Try
                    'End If

                    ' 如果目标是复杂对象但没有 OpenXML，谨慎处理：尽量不删除表格/公式/图片
                    If isTable OrElse isImage OrElse isEquation Then
                        GlobalStatusStrip.ShowWarning($"跳过替换块 '{blockId}'：目标块为表格/图片/公式，且未提供 wordOpenXml。")
                        Return
                    End If

                    ' 普通文本替换：用 OpenXML（更安全）或直接替换文本
                    Dim openXml As String = BuildOpenXmlFromText(textContent)
                    If Not String.IsNullOrEmpty(openXml) Then
                        Try
                            targetRange.Text = ""
                            If InsertOpenXmlIntoRange(openXml, targetRange) Then
                                GlobalStatusStrip.ShowInfo($"已替换 block '{blockId}'（使用 OpenXML）")
                                Return
                            End If
                        Catch ex As Exception
                            Debug.WriteLine("InsertOpenXmlIntoRange 失败: " & ex.Message)
                        End Try
                    End If

                    ' 回退：直接替换纯文本
                    Try
                        targetRange.Text = textContent
                        GlobalStatusStrip.ShowInfo($"已替换 block '{blockId}' 文本（纯文本回退）")
                    Catch ex As Exception
                        Debug.WriteLine("纯文本替换失败: " & ex.Message)
                        GlobalStatusStrip.ShowWarning("替换失败: " & ex.Message)
                    End Try

                Case "insert"
                    If Not String.IsNullOrEmpty(textContent) Then
                        targetRange.InsertBefore(textContent)
                        GlobalStatusStrip.ShowInfo($"已在 block '{blockId}' 前插入文本")
                    End If

                Case "delete"
                    Try
                        ' 删除块（对于复杂对象也允许删除）
                        targetRange.Delete()
                        GlobalStatusStrip.ShowInfo($"已删除 block '{blockId}'")
                    Catch ex As Exception
                        Debug.WriteLine("删除失败: " & ex.Message)
                        GlobalStatusStrip.ShowWarning("删除失败: " & ex.Message)
                    End Try

                Case "insert_table", "table"
                    HandleInsertTable(itemObj, doc, useRange)

                Case "insert_image", "image"
                    HandleInsertImage(itemObj, doc, useRange)

                Case "promoteheading", "maketitle", "makeheading"
                    HandleMakeHeading(itemObj, doc, useRange)

                Case "demoteheading"
                    HandleDemoteHeading(itemObj, doc, useRange)

                Case "makelist", "makebullet", "makenumber"
                    HandleMakeList(itemObj, doc, useRange)

                Case "skip"
                    ' 有意跳过
                    GlobalStatusStrip.ShowInfo($"跳过 block '{blockId}'")

                Case Else
                    ' 默认尝试替换文本（兼容）
                    If Not String.IsNullOrEmpty(textContent) Then
                        Try
                            targetRange.Text = textContent
                            GlobalStatusStrip.ShowInfo($"已对 block '{blockId}' 应用默认替换")
                        Catch ex As Exception
                            Debug.WriteLine("默认替换失败: " & ex.Message)
                        End Try
                    Else
                        GlobalStatusStrip.ShowWarning($"未知 action '{action}'，未执行任何操作")
                    End If
            End Select

        Catch ex As Exception
            Debug.WriteLine("ApplySinglePlanItem (enhanced) 出错: " & ex.Message)
            GlobalStatusStrip.ShowWarning("应用 plan 项失败: " & ex.Message)
        End Try
    End Sub

    ' 通用 openXmlActions 调度器（接收模型返回的 verb 列表，优先以安全方式执行）

    ' 扩展 HandleOpenXmlActions：新增对 findAndReplace / replaceWithWordOpenXml 的支持（优先通过临时 docx + InsertDocxIntoRange 安全写回）
    Private Function HandleOpenXmlActions(openXmlActionsToken As JToken, doc As Object, targetRange As Object, blockId As String) As Boolean
        Try
            If openXmlActionsToken Is Nothing Then Return False
            Dim acted As Boolean = False

            Dim actions As JArray = Nothing
            If openXmlActionsToken.Type = JTokenType.Array Then
                actions = DirectCast(openXmlActionsToken, JArray)
            Else
                actions = New JArray(openXmlActionsToken)
            End If

            For Each a In actions
                If a("verb") Is Nothing Then Continue For
                Dim verb = a("verb").ToString().ToLower()
                Dim paramsToken = If(a("params") IsNot Nothing, a("params"), Nothing)

                Select Case verb
                    Case "findandreplace", "find_and_replace"
                        If paramsToken Is Nothing Then Continue For
                        Dim matchText = If(paramsToken("matchText") IsNot Nothing, paramsToken("matchText").ToString(), "")
                        Dim contextBefore = If(paramsToken("contextBefore") IsNot Nothing, paramsToken("contextBefore").ToString(), "")
                        Dim contextAfter = If(paramsToken("contextAfter") IsNot Nothing, paramsToken("contextAfter").ToString(), "")
                        Dim startIdx As Integer = -1
                        If paramsToken("start") IsNot Nothing Then Integer.TryParse(paramsToken("start").ToString(), startIdx)

                        ' 在 targetRange 范围内查找子范围（优先使用 FindRangeByAnchor2）
                        Dim subRange As Object = Nothing
                        Try
                            subRange = FindRangeByAnchor2(doc, targetRange, targetRange.Text, matchText, contextBefore, contextAfter, startIdx)
                        Catch ex As Exception
                            Debug.WriteLine("FindRangeByAnchor2 failed: " & ex.Message)
                        End Try
                        If subRange Is Nothing Then
                            subRange = FindRangeByAnchor(doc, targetRange, targetRange.Text, matchText, contextBefore, contextAfter, startIdx)
                        End If

                        If subRange Is Nothing Then
                            Debug.WriteLine("findAndReplace: 未能定位到子范围，跳过此 action")
                            Continue For
                        End If

                        ' 优先使用 replaceWithWordOpenXml；其次使用 replaceWithText
                        If paramsToken("replaceWithWordOpenXml") IsNot Nothing Then
                            Dim fragXml = paramsToken("replaceWithWordOpenXml").ToString()
                            Dim tmpDoc = CreateDocxFromWordOpenXml(fragXml)
                            If Not String.IsNullOrEmpty(tmpDoc) Then
                                Try
                                    Dim insertPos As Integer = subRange.Start
                                    subRange.Delete()
                                    Dim insRange = doc.Range(insertPos, insertPos)
                                    If InsertDocxIntoRange(tmpDoc, insRange) Then
                                        acted = True
                                    End If
                                Catch ex As Exception
                                    Debug.WriteLine("findAndReplace replaceWithWordOpenXml 失败: " & ex.Message)
                                Finally
                                    Try : File.Delete(tmpDoc) : Catch : End Try
                                End Try
                            End If
                        ElseIf paramsToken("replaceWithText") IsNot Nothing Then
                            Try
                                Dim txt = paramsToken("replaceWithText").ToString()
                                subRange.Text = txt
                                acted = True
                            Catch ex As Exception
                                Debug.WriteLine("findAndReplace replaceWithText 失败: " & ex.Message)
                            End Try
                        End If

                    Case "replacewithwordopenxml", "replace_with_wordopenxml"
                        If paramsToken Is Nothing Then Continue For
                        Dim frag = If(paramsToken("wordOpenXml") IsNot Nothing, paramsToken("wordOpenXml").ToString(), If(paramsToken("fragment") IsNot Nothing, paramsToken("fragment").ToString(), ""))
                        If String.IsNullOrWhiteSpace(frag) Then Continue For
                        Dim tmp = CreateDocxFromWordOpenXml(frag)
                        If String.IsNullOrEmpty(tmp) Then Continue For
                        Try
                            ' 如果能在 targetRange 内找到更精确子范围，则先替换子范围，否则替换整个 targetRange
                            Dim replaceRange As Object = targetRange
                            Dim matchText = If(paramsToken("matchText") IsNot Nothing, paramsToken("matchText").ToString(), "")
                            Dim contextBefore = If(paramsToken("contextBefore") IsNot Nothing, paramsToken("contextBefore").ToString(), "")
                            Dim contextAfter = If(paramsToken("contextAfter") IsNot Nothing, paramsToken("contextAfter").ToString(), "")
                            Dim startIdx As Integer = -1
                            If paramsToken("start") IsNot Nothing Then Integer.TryParse(paramsToken("start").ToString(), startIdx)

                            Dim subR As Object = Nothing
                            Try
                                subR = FindRangeByAnchor2(doc, targetRange, targetRange.Text, matchText, contextBefore, contextAfter, startIdx)
                            Catch : End Try
                            If subR IsNot Nothing Then replaceRange = subR

                            Dim pos = replaceRange.Start
                            replaceRange.Delete()
                            Dim insRange = doc.Range(pos, pos)
                            If InsertDocxIntoRange(tmp, insRange) Then acted = True
                        Catch ex As Exception
                            Debug.WriteLine("replaceWithWordOpenXml 执行失败: " & ex.Message)
                        Finally
                            Try : File.Delete(tmp) : Catch : End Try
                        End Try

                    Case "insertparagraph", "insert_paragraph"
                        If paramsToken IsNot Nothing AndAlso paramsToken("text") IsNot Nothing Then
                            Dim txt = paramsToken("text").ToString()
                            Dim xml = BuildOpenXmlFromText(txt)
                            If Not String.IsNullOrEmpty(xml) Then
                                Try
                                    If InsertOpenXmlIntoRange(xml, targetRange) Then acted = True
                                Catch ex As Exception
                                    Debug.WriteLine("insertParagraph InsertOpenXmlIntoRange 失败: " & ex.Message)
                                End Try
                            End If
                        End If

                    Case "setrunproperties", "set_run_properties"
                        If paramsToken IsNot Nothing Then
                            Try
                                If paramsToken("fontName") IsNot Nothing Then targetRange.Font.Name = paramsToken("fontName").ToString()
                                If paramsToken("fontSize") IsNot Nothing Then
                                    Dim fs As Integer = 0
                                    Integer.TryParse(paramsToken("fontSize").ToString(), fs)
                                    If fs > 0 Then targetRange.Font.Size = fs
                                End If
                                If paramsToken("bold") IsNot Nothing Then targetRange.Font.Bold = paramsToken("bold").ToObject(Of Boolean)()
                                If paramsToken("italic") IsNot Nothing Then targetRange.Font.Italic = paramsToken("italic").ToObject(Of Boolean)()
                                If paramsToken("color") IsNot Nothing Then
                                    Try
                                        Dim colStr = paramsToken("color").ToString().TrimStart("#"c)
                                        If colStr.Length = 6 Then
                                            Dim r = Convert.ToInt32(colStr.Substring(0, 2), 16)
                                            Dim g = Convert.ToInt32(colStr.Substring(2, 2), 16)
                                            Dim b = Convert.ToInt32(colStr.Substring(4, 2), 16)
                                            targetRange.Font.Color = (r + (g << 8) + (b << 16))
                                        End If
                                    Catch : End Try
                                End If
                                acted = True
                            Catch ex As Exception
                                Debug.WriteLine("setRunProperties 失败: " & ex.Message)
                            End Try
                        End If

                    Case "setparagraphproperties", "set_paragraph_properties"
                        If paramsToken IsNot Nothing Then
                            Try
                                If paramsToken("alignment") IsNot Nothing Then
                                    Dim b = paramsToken("alignment").ToString().ToLower()
                                    Select Case b
                                        Case "left"
                                            targetRange.ParagraphFormat.Alignment = 0
                                        Case "center"
                                            targetRange.ParagraphFormat.Alignment = 1
                                        Case "right"
                                            targetRange.ParagraphFormat.Alignment = 2
                                        Case "justify"
                                            targetRange.ParagraphFormat.Alignment = 3
                                    End Select
                                End If
                                If paramsToken("before") IsNot Nothing Then
                                    Dim beforeVal As Integer = 0
                                    Integer.TryParse(paramsToken("before").ToString(), beforeVal)
                                    If beforeVal >= 0 Then targetRange.ParagraphFormat.SpaceBefore = beforeVal
                                End If
                                If paramsToken("after") IsNot Nothing Then
                                    Dim afterVal As Integer = 0
                                    Integer.TryParse(paramsToken("after").ToString(), afterVal)
                                    If afterVal >= 0 Then targetRange.ParagraphFormat.SpaceAfter = afterVal
                                End If
                                If paramsToken("line") IsNot Nothing Then
                                    Dim lineVal As Integer = 0
                                    Integer.TryParse(paramsToken("line").ToString(), lineVal)
                                    If lineVal > 0 Then targetRange.ParagraphFormat.LineSpacing = lineVal
                                End If
                                acted = True
                            Catch ex As Exception
                                Debug.WriteLine("setParagraphProperties 失败: " & ex.Message)
                            End Try
                        End If

                    Case "inserttable", "insert_table"
                        If paramsToken IsNot Nothing AndAlso paramsToken("rows") IsNot Nothing Then
                            Try
                                Dim rowsToken = paramsToken("rows")
                                If rowsToken.Type = JTokenType.Array Then
                                    Dim rowsList As New List(Of String())
                                    For Each r In DirectCast(rowsToken, JArray)
                                        If r.Type = JTokenType.Array Then
                                            Dim cols = DirectCast(r, JArray).Select(Function(c) c.ToString()).ToArray()
                                            rowsList.Add(cols)
                                        End If
                                    Next
                                    Dim xml = BuildTableOpenXml(rowsList)
                                    If Not String.IsNullOrEmpty(xml) Then
                                        If InsertOpenXmlIntoRange(xml, targetRange) Then acted = True
                                    End If
                                End If
                            Catch ex As Exception
                                Debug.WriteLine("insertTable 失败: " & ex.Message)
                            End Try
                        End If

                    Case "insertimage", "insert_image"
                        If paramsToken IsNot Nothing Then
                            Try
                                Dim imageUrl As String = If(paramsToken("imageUrl") IsNot Nothing, paramsToken("imageUrl").ToString(), If(paramsToken("src") IsNot Nothing, paramsToken("src").ToString(), ""))
                                If Not String.IsNullOrEmpty(imageUrl) Then
                                    Dim tmp = DownloadUrlToTempFile(imageUrl)
                                    If Not String.IsNullOrEmpty(tmp) Then
                                        Try
                                            targetRange.InlineShapes.AddPicture(tmp, False, True)
                                            acted = True
                                        Catch ex As Exception
                                            Debug.WriteLine("插入图片（InlineShapes）失败: " & ex.Message)
                                        End Try
                                        Try : File.Delete(tmp) : Catch : End Try
                                    End If
                                End If
                            Catch ex As Exception
                                Debug.WriteLine("insertImage 处理失败: " & ex.Message)
                            End Try
                        End If

                    Case "insertpagebreak", "insert_page_break"
                        Try
                            targetRange.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
                            acted = True
                        Catch ex As Exception
                            Debug.WriteLine("insertPageBreak 失败: " & ex.Message)
                        End Try

                    Case Else
                        Debug.WriteLine($"未知 openXml verb: {verb}（跳过）")
                End Select
            Next

            Return acted
        Catch ex As Exception
            Debug.WriteLine("HandleOpenXmlActions 出错: " & ex.Message)
            Return False
        End Try
    End Function

    ' 简单构造 OpenXML 表格片段（用于 InsertOpenXmlIntoRange 回写）
    Private Function BuildTableOpenXml(rows As List(Of String())) As String
        Try
            If rows Is Nothing OrElse rows.Count = 0 Then Return String.Empty
            Dim ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            Dim sb As New StringBuilder()
            sb.Append($"<w:tbl xmlns:w=""{ns}"">")
            ' 简单表格属性（可扩展）
            sb.Append("<w:tblPr><w:tblBorders>" &
                      "<w:top w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""000000""/>" &
                      "<w:left w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""000000""/>" &
                      "<w:bottom w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""000000""/>" &
                      "<w:right w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""000000""/>" &
                      "<w:insideH w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""000000""/>" &
                      "<w:insideV w:val=""single"" w:sz=""4"" w:space=""0"" w:color=""000000""/>" &
                      "</w:tblBorders></w:tblPr>")
            For Each r In rows
                sb.Append("<w:tr>")
                For Each c In r
                    Dim escaped = If(c, "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    sb.Append($"<w:tc><w:p><w:r><w:t xml:space=""preserve"">{escaped}</w:t></w:r></w:p></w:tc>")
                Next
                sb.Append("</w:tr>")
            Next
            sb.Append("</w:tbl>")
            ' 包装为完整文档片段，以兼容 InsertXML
            Dim docNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            Dim wrapper As New StringBuilder()
            wrapper.Append($"<w:document xmlns:w=""{docNs}""><w:body>")
            wrapper.Append(sb.ToString())
            wrapper.Append("</w:body></w:document>")
            Return wrapper.ToString()
        Catch ex As Exception
            Debug.WriteLine("BuildTableOpenXml 出错: " & ex.Message)
            Return String.Empty
        End Try
    End Function


    ' 新：按 ReformatButton_Click 的分块策略通过 blockId 定位 Range（返回具体 Range 或 Nothing）
    Private Function FindRangeByBlockId(doc As Object, baseRange As Object, blockId As String) As Object
        Try
            If String.IsNullOrEmpty(blockId) OrElse baseRange Is Nothing Then Return Nothing

            Dim blockIndex As Integer = 0

            For Each p In baseRange.Paragraphs
                Dim r = p.Range
                ' 表格：将每个表格视为独立块（与 ReformatButton_Click 保持一致）
                If r.Tables IsNot Nothing AndAlso r.Tables.Count > 0 Then
                    For i As Integer = 1 To r.Tables.Count
                        Dim t = r.Tables(i)
                        Dim tRange = t.Range
                        If $"blk_{blockIndex}" = blockId Then Return tRange
                        blockIndex += 1
                    Next
                ElseIf (r.InlineShapes IsNot Nothing AndAlso r.InlineShapes.Count > 0) OrElse (r.ShapeRange IsNot Nothing AndAlso r.ShapeRange.Count > 0) Then
                    ' 图片 / 形状 块
                    If $"blk_{blockIndex}" = blockId Then
                        Return r
                    End If
                    blockIndex += 1
                ElseIf (r.OMaths IsNot Nothing AndAlso r.OMaths.Count > 0) Then
                    ' 公式
                    If $"blk_{blockIndex}" = blockId Then
                        Return r
                    End If
                    blockIndex += 1
                Else
                    ' 常规段落
                    If $"blk_{blockIndex}" = blockId Then
                        Return r
                    End If
                    blockIndex += 1
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine("FindRangeByBlockId 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function


    ' 新实现：按 blockId 或任意锚文本定位 Range
    ' 说明：本实现向后兼容原有 matchText 用法，但优先把 matchText 当作 blockId（如 'bl_3'）。
    ' 查找策略（按序）：
    '  1) 在 fullText 中查找常见占位格式： <!--bl_3-->、[[bl_3]]、{bl_3}、(bl_3) 或 直文本 'bl_3'；
    '  2) 若未找到且 matchText 非空，则尝试进行不区分大小写的 IndexOf 搜索（legacy 支持）；
    '  3) 若提供 startIdx 则基于 startIdx 返回对应 Range；
    '  4) 未命中则返回 Nothing（调用方可退化为末尾插入）。
    Private Function FindRangeByAnchor2(doc As Object, useRange As Object, fullText As String, matchText As String, contextBefore As String, contextAfter As String, startIdx As Integer) As Object
        Try
            If String.IsNullOrEmpty(fullText) Then Return Nothing

            If Not String.IsNullOrEmpty(matchText) Then
                Dim patterns As New List(Of String) From {
                    $"<!--{matchText}-->",
                    $"[[{matchText}]]",
                    $"{{{matchText}}}",
                    $"({matchText})",
                    matchText
                }

                For Each pat In patterns
                    Dim idx = fullText.IndexOf(pat, StringComparison.OrdinalIgnoreCase)
                    If idx >= 0 Then
                        Dim docStart = useRange.Start + idx
                        Dim docEnd = Math.Min(useRange.Start + idx + pat.Length, useRange.End)
                        Return doc.Range(docStart, docEnd)
                    End If
                Next

                ' 兼容老的直接 matchText 内容查找（当 matchText 实际为一段文本时）
                Dim idx2 = fullText.IndexOf(matchText, StringComparison.OrdinalIgnoreCase)
                If idx2 >= 0 Then
                    Dim docStart = useRange.Start + idx2
                    Dim docEnd = Math.Min(useRange.Start + idx2 + matchText.Length, useRange.End)
                    Return doc.Range(docStart, docEnd)
                End If
            End If

            ' 如果提供 start 索引则使用（边界检查）
            If startIdx >= 0 AndAlso startIdx < fullText.Length Then
                Dim s = useRange.Start + startIdx
                Dim e = Math.Min(useRange.Start + startIdx + 1, useRange.End)
                Return doc.Range(s, e)
            End If

        Catch ex As Exception
            Debug.WriteLine("FindRangeByAnchor (block-based) 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

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

    ' -------------------- 各类 handler（可按需扩展） --------------------

    Private Function HandleMakeHeading(itemObj As JObject, doc As Object, useRange As Object) As Boolean
        Try
            Dim matchText = If(itemObj("matchText") IsNot Nothing, itemObj("matchText").ToString(), "")
            Dim contextBefore = If(itemObj("contextBefore") IsNot Nothing, itemObj("contextBefore").ToString(), "")
            Dim contextAfter = If(itemObj("contextAfter") IsNot Nothing, itemObj("contextAfter").ToString(), "")
            Dim attributes = If(itemObj("attributes") IsNot Nothing, DirectCast(itemObj("attributes"), JObject), Nothing)
            Dim level As Integer = 1
            If attributes IsNot Nothing AndAlso attributes("headingLevel") IsNot Nothing Then Integer.TryParse(attributes("headingLevel").ToString(), level)
            Dim targetRange = FindRangeByAnchor(doc, useRange, useRange.Text, matchText, contextBefore, contextAfter, -1)
            If targetRange Is Nothing Then Return False
            Dim styleName As String = "Heading " & Math.Max(1, Math.Min(9, level)).ToString()
            targetRange.Style = styleName
            GlobalStatusStrip.ShowInfo($"已设置为 {styleName}")
            Return True
        Catch ex As Exception
            Debug.WriteLine("HandleMakeHeading error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function HandleDemoteHeading(itemObj As JObject, doc As Object, useRange As Object) As Boolean
        ' 这里简单实现为降低 heading 级别
        Try
            Dim attributes = If(itemObj("attributes") IsNot Nothing, DirectCast(itemObj("attributes"), JObject), Nothing)
            Dim matchText = If(itemObj("matchText") IsNot Nothing, itemObj("matchText").ToString(), "")
            Dim targetRange = FindRangeByAnchor(doc, useRange, useRange.Text, matchText, If(itemObj("contextBefore")?.ToString(), ""), If(itemObj("contextAfter")?.ToString(), ""), -1)
            If targetRange Is Nothing Then Return False
            Dim currentStyle = targetRange.Style
            ' 若样式为 Heading N 则降一级
            Dim m = System.Text.RegularExpressions.Regex.Match(currentStyle.ToString(), "Heading\s*(\d+)")
            If m.Success Then
                Dim lvl = Integer.Parse(m.Groups(1).Value)
                lvl = Math.Min(9, lvl + 1)
                targetRange.Style = "Heading " & lvl.ToString()
                GlobalStatusStrip.ShowInfo($"标题降级为 Heading {lvl}")
                Return True
            End If
            Return False
        Catch ex As Exception
            Debug.WriteLine("HandleDemoteHeading error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function HandleMakeList(itemObj As JObject, doc As Object, useRange As Object) As Boolean
        Try
            Dim matchText = If(itemObj("matchText") IsNot Nothing, itemObj("matchText").ToString(), "")
            Dim attributes = If(itemObj("attributes") IsNot Nothing, DirectCast(itemObj("attributes"), JObject), Nothing)
            Dim targetRange = FindRangeByAnchor(doc, useRange, useRange.Text, matchText, If(itemObj("contextBefore")?.ToString(), ""), If(itemObj("contextAfter")?.ToString(), ""), -1)
            If targetRange Is Nothing Then Return False
            Dim listType = If(attributes IsNot Nothing AndAlso attributes("listType") IsNot Nothing, attributes("listType").ToString().ToLower(), "bullet")
            If listType = "number" Then
                targetRange.ListFormat.ApplyNumberDefault()
            Else
                targetRange.ListFormat.ApplyBulletDefault()
            End If
            GlobalStatusStrip.ShowInfo("已转为列表")
            Return True
        Catch ex As Exception
            Debug.WriteLine("HandleMakeList error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function HandleReplace(itemObj As JObject, doc As Object, useRange As Object) As Boolean
        Try
            Dim matchText = If(itemObj("matchText") IsNot Nothing, itemObj("matchText").ToString(), "")
            Dim replaceWith = If(itemObj("replaceWith") IsNot Nothing, itemObj("replaceWith").ToString(), "")
            Dim contextBefore = If(itemObj("contextBefore") IsNot Nothing, itemObj("contextBefore").ToString(), "")
            Dim contextAfter = If(itemObj("contextAfter") IsNot Nothing, itemObj("contextAfter").ToString(), "")
            Dim startIdx As Integer = -1
            Dim tr = FindRangeByAnchor(doc, useRange, useRange.Text, matchText, contextBefore, contextAfter, startIdx)
            If tr Is Nothing Then Return False
            tr.Text = replaceWith
            GlobalStatusStrip.ShowInfo("已替换文本")
            Return True
        Catch ex As Exception
            Debug.WriteLine("HandleReplace error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function HandleInsert(itemObj As JObject, doc As Object, useRange As Object) As Boolean
        Try
            ' 插入纯文本或 previewHtml（如果有）
            Dim insertHtml = If(itemObj("previewHtml") IsNot Nothing, itemObj("previewHtml").ToString(), "")
            Dim insertText = If(itemObj("replaceWith") IsNot Nothing, itemObj("replaceWith").ToString(), "")
            Dim matchText = If(itemObj("matchText") IsNot Nothing, itemObj("matchText").ToString(), "")
            Dim contextBefore = If(itemObj("contextBefore") IsNot Nothing, itemObj("contextBefore").ToString(), "")
            Dim contextAfter = If(itemObj("contextAfter") IsNot Nothing, itemObj("contextAfter").ToString(), "")
            Dim tr = FindRangeByAnchor(doc, useRange, useRange.Text, matchText, contextBefore, contextAfter, -1)
            If tr Is Nothing Then
                ' 在文档末尾插入
                tr = doc.Content
                tr.Collapse(0) ' wdCollapseEnd=0
            End If

            If Not String.IsNullOrEmpty(insertHtml) Then
                Return PasteHtmlToRange(insertHtml, tr)
            ElseIf Not String.IsNullOrEmpty(insertText) Then
                tr.InsertBefore(insertText)
                Return True
            End If
            Return False
        Catch ex As Exception
            Debug.WriteLine("HandleInsert error: " & ex.Message)
            Return False
        End Try
    End Function

    'Private Function HandleInsertHtml(itemObj As JObject, doc As Object, useRange As Object) As Boolean
    '    ' 直接调用 HandleInsert（复用）
    '    Return HandleInsert(itemObj, doc, useRange)
    'End Function

    Private Function HandleInsertImage(itemObj As JObject, doc As Object, useRange As Object) As Boolean
        Try
            Dim imageUrl = If(itemObj("attributes") IsNot Nothing AndAlso itemObj("attributes")("src") IsNot Nothing, itemObj("attributes")("src").ToString(), "")
            If String.IsNullOrEmpty(imageUrl) Then Return False
            Dim matchText = If(itemObj("matchText") IsNot Nothing, itemObj("matchText").ToString(), "")
            Dim contextBefore = If(itemObj("contextBefore") IsNot Nothing, itemObj("contextBefore").ToString(), "")
            Dim tr = FindRangeByAnchor(doc, useRange, useRange.Text, matchText, contextBefore, If(itemObj("contextAfter")?.ToString(), ""), -1)
            If tr Is Nothing Then tr = doc.Content : tr.Collapse(0)
            Dim tmpFile = DownloadUrlToTempFile(imageUrl)
            If String.IsNullOrEmpty(tmpFile) Then Return False
            tr.InlineShapes.AddPicture(tmpFile, False, True)
            GlobalStatusStrip.ShowInfo("已插入图片")
            Return True
        Catch ex As Exception
            Debug.WriteLine("HandleInsertImage error: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function HandleInsertTable(itemObj As JObject, doc As Object, useRange As Object) As Boolean
        Try
            ' 支持 attributes.tableData 为二维数组或 CSV 字符串
            Dim tableDataToken = If(itemObj("attributes") IsNot Nothing, itemObj("attributes")("tableData"), Nothing)
            Dim rows As New List(Of String())
            If tableDataToken Is Nothing Then
                Return False
            End If

            If tableDataToken.Type = JTokenType.String Then
                ' 解析 CSV（简单）
                Dim csv = tableDataToken.ToString()
                For Each line In csv.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    Dim cols = line.Split(New Char() {vbTab, ","c})
                    rows.Add(cols)
                Next
            ElseIf tableDataToken.Type = JTokenType.Array Then
                For Each r In DirectCast(tableDataToken, JArray)
                    If r.Type = JTokenType.Array Then
                        Dim listcols = New List(Of String)()
                        For Each c In DirectCast(r, JArray)
                            listcols.Add(c.ToString())
                        Next
                        rows.Add(listcols.ToArray())
                    End If
                Next
            End If

            If rows.Count = 0 Then Return False
            Dim rCount = rows.Count
            Dim cCount = rows(0).Length
            Dim matchText = If(itemObj("matchText") IsNot Nothing, itemObj("matchText").ToString(), "")
            Dim tr = FindRangeByAnchor(doc, useRange, useRange.Text, matchText, If(itemObj("contextBefore")?.ToString(), ""), If(itemObj("contextAfter")?.ToString(), ""), -1)
            If tr Is Nothing Then tr = doc.Content : tr.Collapse(0)
            Dim table = doc.Tables.Add(tr, rCount, cCount)
            For i = 0 To rCount - 1
                For j = 0 To cCount - 1
                    table.Cell(i + 1, j + 1).Range.Text = rows(i)(j)
                Next
            Next
            GlobalStatusStrip.ShowInfo("已插入表格")
            Return True
        Catch ex As Exception
            Debug.WriteLine("HandleInsertTable error: " & ex.Message)
            Return False
        End Try
    End Function

    ' -------------------- 工具函数 --------------------

    ' 将 HTML 粘贴到指定 Range（通过剪贴板 + PasteSpecial wdPasteHTML）
    Private Function PasteHtmlToRange(html As String, targetRange As Object) As Boolean
        Try
            If String.IsNullOrEmpty(html) OrElse targetRange Is Nothing Then Return False

            ' 构造 IDataObject 并设置 HTML
            Dim dataObj As New DataObject()
            dataObj.SetData(DataFormats.Html, html)
            ' 需要在 STA 线程上设置剪贴板
            Try
                Clipboard.SetDataObject(dataObj, True)
            Catch ex As Exception
                Debug.WriteLine("Clipboard.SetDataObject 失败: " & ex.Message)
                ' 回退：直接插入纯文本
                targetRange.InsertBefore(StripHtmlTags(html))
                Return True
            End Try

            ' 粘贴 HTML（wdPasteHTML = 10）
            targetRange.PasteSpecial(DataType:=10)
            Return True
        Catch ex As Exception
            Debug.WriteLine("PasteHtmlToRange 错误: " & ex.Message)
            Return False
        End Try
    End Function

    ' 从 URL 下载到临时文件（图片）
    Private Function DownloadUrlToTempFile(url As String) As String
        Try
            Using client As New System.Net.Http.HttpClient()
                Dim bytes = client.GetByteArrayAsync(url).Result
                Dim ext = Path.GetExtension(New Uri(url).LocalPath)
                If String.IsNullOrEmpty(ext) Then ext = ".img"
                Dim tmp = Path.Combine(Path.GetTempPath(), "oa_img_" & Guid.NewGuid().ToString() & ext)
                File.WriteAllBytes(tmp, bytes)
                Return tmp
            End Using
        Catch ex As Exception
            Debug.WriteLine("DownloadUrlToTempFile 失败: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    ' 简单的 HTML -> 纯文本（回退）
    Private Function StripHtmlTags(html As String) As String
        If String.IsNullOrEmpty(html) Then Return ""
        Return System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", "")
    End Function


    ' 类级辅助：通过 matchText + context 定位 Range，回退到 start 或模糊匹配
    Private Function FindRangeByAnchor(doc As Object, useRange As Object, fullText As String, matchText As String, contextBefore As String, contextAfter As String, startIdx As Integer) As Object
        Try
            ' 1) 精确 matchText 查找（忽略大小写），并验证前后 context
            If Not String.IsNullOrEmpty(matchText) Then
                Dim startPos As Integer = 0
                Do
                    Dim idx As Integer = fullText.IndexOf(matchText, startPos, StringComparison.OrdinalIgnoreCase)
                    If idx < 0 Then Exit Do
                    Dim ok As Boolean = True
                    If Not String.IsNullOrEmpty(contextBefore) Then
                        Dim beforeStart = Math.Max(0, idx - contextBefore.Length)
                        Dim actualBefore = fullText.Substring(beforeStart, Math.Min(contextBefore.Length, idx - beforeStart))
                        If Not actualBefore.EndsWith(contextBefore.Trim(), StringComparison.OrdinalIgnoreCase) Then ok = False
                    End If
                    If ok AndAlso Not String.IsNullOrEmpty(contextAfter) Then
                        Dim afterIdx = idx + matchText.Length
                        If afterIdx <= fullText.Length - 1 Then
                            Dim actualAfterLen = Math.Min(contextAfter.Length, fullText.Length - afterIdx)
                            Dim actualAfter = fullText.Substring(afterIdx, actualAfterLen)
                            If Not actualAfter.StartsWith(contextAfter.Trim(), StringComparison.OrdinalIgnoreCase) Then ok = False
                        Else
                            ok = False
                        End If
                    End If
                    If ok Then
                        Dim docStart = useRange.Start + idx
                        Dim docEnd = docStart + matchText.Length
                        Return doc.Range(docStart, docEnd)
                    End If
                    startPos = idx + 1
                Loop
            End If

            ' 2) 如果 composite context 提供，尝试整体匹配
            If Not String.IsNullOrEmpty(contextBefore) AndAlso Not String.IsNullOrEmpty(contextAfter) AndAlso Not String.IsNullOrEmpty(matchText) Then
                Dim composite = contextBefore.Trim() & matchText & contextAfter.Trim()
                Dim idx2 = fullText.IndexOf(composite, StringComparison.OrdinalIgnoreCase)
                If idx2 >= 0 Then
                    Dim docStart = useRange.Start + idx2 + contextBefore.Trim().Length
                    Dim docEnd = docStart + matchText.Length
                    Return doc.Range(docStart, docEnd)
                End If
            End If

            ' 3) 若提供 start 索引则使用（边界检查）
            If startIdx >= 0 AndAlso startIdx < fullText.Length Then
                Dim s = useRange.Start + startIdx
                Dim e = Math.Min(useRange.Start + startIdx + Math.Max(1, If(matchText?.Length, 1)), useRange.End)
                Return doc.Range(s, e)
            End If

            ' 4) 退化为模糊匹配：遍历段落，使用最短编辑距离作为回退
            Dim bestRange As Object = Nothing
            Dim bestDist As Integer = Integer.MaxValue
            For Each p In useRange.Paragraphs
                Dim pt = p.Range.Text
                If String.IsNullOrWhiteSpace(pt) Then Continue For
                Dim window = Math.Min(pt.Length, Math.Max(1, If(matchText?.Length, 20)))
                For i As Integer = 0 To Math.Max(0, pt.Length - window)
                    Dim subx = pt.Substring(i, Math.Min(window, pt.Length - i))
                    Dim dist = LevenshteinDistance(subx, If(matchText, subx))
                    If dist < bestDist Then
                        bestDist = dist
                        Dim docStart = p.Range.Start + i
                        Dim docEnd = Math.Min(docStart + subx.Length, p.Range.End)
                        bestRange = doc.Range(docStart, docEnd)
                    End If
                Next
            Next
            If bestRange IsNot Nothing AndAlso bestDist <= Math.Max(1, Math.Min(10, If(matchText?.Length, 10) \ 3)) Then
                Return bestRange
            End If

        Catch ex As Exception
            Debug.WriteLine("FindRangeByAnchor 出错: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' Levenshtein 编辑距离（类级辅助）
    Private Function LevenshteinDistance(s As String, t As String) As Integer
        If s Is Nothing Then s = ""
        If t Is Nothing Then t = ""
        Dim n = s.Length
        Dim m = t.Length
        If n = 0 Then Return m
        If m = 0 Then Return n
        Dim d(n + 1, m + 1) As Integer
        For i = 0 To n
            d(i, 0) = i
        Next
        For j = 0 To m
            d(0, j) = j
        Next
        For i = 1 To n
            For j = 1 To m
                Dim cost = If(s(i - 1) = t(j - 1), 0, 1)
                d(i, j) = Math.Min(Math.Min(d(i - 1, j) + 1, d(i, j - 1) + 1), d(i - 1, j - 1) + cost)
            Next
        Next
        Return d(n, m)
    End Function

    ' 新：在 Range 插入 docx 文件内容（用于 HtmlToOpenXml 生成的临时 docx）
    Private Function InsertDocxIntoRange(tempDocxPath As String, targetRange As Object) As Boolean
        Try
            If String.IsNullOrEmpty(tempDocxPath) OrElse targetRange Is Nothing Then Return False
            ' 插入文件（InsertFile 会将 docx 内容插入）
            targetRange.InsertFile(tempDocxPath)
            Return True
        Catch ex As Exception
            Debug.WriteLine("InsertDocxIntoRange 失败: " & ex.Message)
            Return False
        End Try
    End Function



    Protected Overrides Function CaptureCurrentSelectionInfo(mode As String) As SelectionInfo
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            Dim txt As String = If(sel IsNot Nothing AndAlso sel.Range IsNot Nothing, sel.Range.Text, String.Empty)
            If String.IsNullOrEmpty(txt) Then
                If String.Equals(mode, "reformat", StringComparison.OrdinalIgnoreCase) Or String.Equals(mode, "proofread", StringComparison.OrdinalIgnoreCase) Then
                    ' 如果未选中，并且是重构排版或审阅功能，则获取所有内容
                    Dim doc = Globals.ThisAddIn.Application.ActiveDocument
                    txt = doc.Content.Text
                End If
            End If

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

            Return info
        Catch
        End Try
    End Function

    Private Function CreateDocxFromWordOpenXml(fragmentXml As String) As String
        Try
            If String.IsNullOrWhiteSpace(fragmentXml) Then Return String.Empty
            Dim tmpPath = Path.Combine(Path.GetTempPath(), $"oa_fragment_{Guid.NewGuid():N}.docx")

            Using wordDoc = WordprocessingDocument.Create(tmpPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
                Dim mainPart = wordDoc.AddMainDocumentPart()
                mainPart.Document = New DocumentFormat.OpenXml.Wordprocessing.Document(New DocumentFormat.OpenXml.Wordprocessing.Body())

                Dim inner As String = fragmentXml
                Try
                    ' 尝试提取 <w:body> 的内容以兼容完整文档或片段
                    Dim m = System.Text.RegularExpressions.Regex.Match(fragmentXml, "<w:body[^>]*>([\s\S]*?)</w:body>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    If m.Success Then
                        inner = m.Groups(1).Value
                    Else
                        Dim m2 = System.Text.RegularExpressions.Regex.Match(fragmentXml, "<w:document[^>]*>([\s\S]*?)</w:document>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                        If m2.Success Then
                            Dim mm = System.Text.RegularExpressions.Regex.Match(m2.Groups(1).Value, "<w:body[^>]*>([\s\S]*?)</w:body>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                            If mm.Success Then inner = mm.Groups(1).Value
                        End If
                    End If
                Catch
                End Try

                Try
                    ' 直接设置 Body.InnerXml（fragment 必须为合法 WordProcessingML 的子节点集合）
                    mainPart.Document.Body.InnerXml = inner
                    mainPart.Document.Save()
                Catch ex As Exception
                    Debug.WriteLine("CreateDocxFromWordOpenXml: 设置 InnerXml 失败: " & ex.Message)
                    ' 回退：把 fragment 当成纯文本段落写入
                    mainPart.Document.Body = New DocumentFormat.OpenXml.Wordprocessing.Body()
                    Dim p As New DocumentFormat.OpenXml.Wordprocessing.Paragraph()
                    Dim r As New DocumentFormat.OpenXml.Wordprocessing.Run()
                    r.AppendChild(New DocumentFormat.OpenXml.Wordprocessing.Text(fragmentXml))
                    p.AppendChild(r)
                    mainPart.Document.Body.AppendChild(p)
                    mainPart.Document.Save()
                End Try
            End Using

            Return tmpPath
        Catch ex As Exception
            Debug.WriteLine("CreateDocxFromWordOpenXml 失败: " & ex.Message)
            Return String.Empty
        End Try
    End Function


End Class

