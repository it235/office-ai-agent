Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 自动补全服务：处理输入补全请求、FIM/Chat 两种补全模式、补全历史记录
''' </summary>
Public Class AutocompleteService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _getContextSnapshot As Func(Of JObject)
    Private ReadOnly _getAppType As Func(Of String)

    Public Sub New(
        executeScript As Func(Of String, Task),
        getContextSnapshot As Func(Of JObject),
        getAppType As Func(Of String))

        _executeScript = executeScript
        _getContextSnapshot = getContextSnapshot
        _getAppType = getAppType
    End Sub

    ''' <summary>
    ''' 处理自动补全请求
    ''' </summary>
    Public Async Function HandleRequestCompletion(jsonDoc As JObject) As Task
        Try
            If Not ChatSettings.EnableAutocomplete Then Return

            Dim inputText As String = If(jsonDoc("input")?.ToString(), "")
            Dim timestamp As Long = If(jsonDoc("timestamp")?.Value(Of Long)(), 0)

            If String.IsNullOrWhiteSpace(inputText) OrElse inputText.Length < 2 Then Return

            Dim contextSnapshot = _getContextSnapshot()
            Dim completions = Await RequestCompletionsFromLLM(inputText, contextSnapshot)

            Dim resultJson As New JObject()
            resultJson("completions") = JArray.FromObject(completions)
            resultJson("timestamp") = timestamp

            Await _executeScript($"showCompletions({resultJson.ToString(Newtonsoft.Json.Formatting.None)});")

        Catch ex As Exception
            Debug.WriteLine($"HandleRequestCompletion 出错: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 处理补全采纳记录
    ''' </summary>
    Public Sub HandleAcceptCompletion(jsonDoc As JObject)
        Try
            Dim inputText As String = If(jsonDoc("input")?.ToString(), "")
            Dim completion As String = If(jsonDoc("completion")?.ToString(), "")
            Dim context As String = If(jsonDoc("context")?.ToString(), "")
            RecordCompletionHistory(inputText, completion, context)
        Catch ex As Exception
            Debug.WriteLine($"HandleAcceptCompletion 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 为意图识别/规划阶段丰富上下文：内容区引用摘要 + RAG 相关记忆
    ''' </summary>
    Public Sub EnrichContextForIntent(snapshot As JObject,
                                      question As String,
                                      filePaths As List(Of String),
                                      selectedContents As List(Of SendMessageReferenceContentItem))
        If snapshot Is Nothing Then Return
        Dim refParts As New List(Of String)()
        If filePaths IsNot Nothing AndAlso filePaths.Count > 0 Then
            refParts.Add($"用户引用了 {filePaths.Count} 个文件")
        End If
        If selectedContents IsNot Nothing AndAlso selectedContents.Count > 0 Then
            refParts.Add($"{selectedContents.Count} 段选中内容")
            For i = 0 To Math.Min(selectedContents.Count - 1, 4)
                Dim item = selectedContents(i)
                Dim desc = If(String.IsNullOrEmpty(item.sheetName), item.address, $"{item.sheetName}: {item.address}")
                If desc.Length > 60 Then desc = desc.Substring(0, 57) & "..."
                refParts.Add($"  - {desc}")
            Next
        End If
        If refParts.Count > 0 Then
            snapshot("referenceSummary") = String.Join("；" & vbCrLf, refParts)
        End If
        If Not String.IsNullOrWhiteSpace(question) Then
            Try
                Dim memories = MemoryService.GetRelevantMemories(question, 2, Nothing, Nothing, _getAppType())
                If memories IsNot Nothing AndAlso memories.Count > 0 Then
                    Dim lines As New List(Of String)()
                    For Each m In memories
                        Dim c = If(m.Content, "").Trim()
                        If c.Length > 200 Then c = c.Substring(0, 197) & "..."
                        If Not String.IsNullOrEmpty(c) Then lines.Add(c)
                    Next
                    If lines.Count > 0 Then snapshot("ragSnippets") = String.Join(vbCrLf & "---" & vbCrLf, lines)
                End If
            Catch ex As Exception
                Debug.WriteLine($"EnrichContextForIntent RAG: {ex.Message}")
            End Try
        End If
    End Sub

    ''' <summary>
    ''' 调用大模型获取补全建议
    ''' </summary>
    Private Async Function RequestCompletionsFromLLM(inputText As String, contextSnapshot As JObject) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        Try
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then Return completions

            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)

            Dim apiUrl = cfg.url
            Dim apiKey = cfg.key

            Dim useFimMode = selectedModel.fimSupported AndAlso Not String.IsNullOrEmpty(selectedModel.fimUrl)

            If useFimMode Then
                completions = Await RequestCompletionsWithFIM(inputText, contextSnapshot, selectedModel, apiKey)
            Else
                completions = Await RequestCompletionsWithChat(inputText, contextSnapshot, cfg, selectedModel, apiKey)
            End If
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsFromLLM 出错: {ex.Message}")
        End Try
        Return completions
    End Function

    ''' <summary>
    ''' 使用 FIM (Fill-In-the-Middle) API 获取补全
    ''' </summary>
    Private Async Function RequestCompletionsWithFIM(inputText As String, contextSnapshot As JObject,
                                                      model As ConfigManager.ConfigItemModel, apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        Try
            Dim fimUrl = model.fimUrl
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("prompt") = inputText
            requestObj("suffix") = ""
            requestObj("max_tokens") = 50
            requestObj("temperature") = 0.3
            requestObj("stream") = False
            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(10)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(fimUrl, content)
                response.EnsureSuccessStatusCode()
                Dim responseBody = Await response.Content.ReadAsStringAsync()
                Dim jObj = JObject.Parse(responseBody)
                Dim text = jObj("choices")?(0)?("text")?.ToString()
                If Not String.IsNullOrWhiteSpace(text) Then
                    text = text.Trim().Split({vbCr, vbLf, vbCrLf}, StringSplitOptions.RemoveEmptyEntries)(0)
                    If text.Length <= 50 Then completions.Add(text)
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsWithFIM 出错: {ex.Message}")
        End Try
        Return completions
    End Function

    ''' <summary>
    ''' 使用 Chat Completion API 获取补全
    ''' </summary>
    Private Async Function RequestCompletionsWithChat(inputText As String, contextSnapshot As JObject,
                                                       cfg As ConfigManager.ConfigItem, model As ConfigManager.ConfigItemModel,
                                                       apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        Try
            Dim apiUrl = cfg.url
            Dim modelName = model.modelName
            Dim appType = If(contextSnapshot("appType")?.ToString(), "Office")
            Dim selectionText = If(contextSnapshot("selection")?.ToString(), "")
            Dim systemPrompt = GetCompletionSystemPrompt(appType)

            Dim userContent As New StringBuilder()
            userContent.AppendLine($"当前应用: {appType}")
            userContent.AppendLine($"用户已输入: ""{inputText}""")
            If Not String.IsNullOrWhiteSpace(selectionText) Then
                userContent.AppendLine($"选中内容: ""{selectionText.Substring(0, Math.Min(200, selectionText.Length))}""")
            End If
            If contextSnapshot("sheetName") IsNot Nothing Then
                userContent.AppendLine($"当前工作表: {contextSnapshot("sheetName")}")
            End If
            If contextSnapshot("slideIndex") IsNot Nothing Then
                userContent.AppendLine($"当前幻灯片: 第{contextSnapshot("slideIndex")}页")
            End If
            userContent.AppendLine()
            userContent.AppendLine("请给出补全建议（JSON格式）。")

            Dim requestObj As New JObject()
            requestObj("model") = modelName
            requestObj("stream") = False
            requestObj("temperature") = 0.3
            Dim messages As New JArray()
            messages.Add(New JObject() From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject() From {{"role", "user"}, {"content", userContent.ToString()}})
            requestObj("messages") = messages
            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(10)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()
                Dim responseBody = Await response.Content.ReadAsStringAsync()

                Dim jObj As JObject = Nothing
                Try
                    jObj = JObject.Parse(responseBody)
                Catch apiParseEx As Exception
                    Debug.WriteLine($"解析API响应失败: {apiParseEx.Message}")
                    Return completions
                End Try

                Dim msg As String = Nothing
                Try
                    msg = jObj("choices")?(0)?("message")?("content")?.ToString()
                Catch
                    msg = jObj("message")?.ToString()
                End Try

                If Not String.IsNullOrEmpty(msg) Then
                    Try
                        Dim cleanedMsg = msg.Trim()
                        If cleanedMsg.StartsWith("```") Then
                            Dim firstNewLine = cleanedMsg.IndexOf(vbLf)
                            If firstNewLine > 0 Then cleanedMsg = cleanedMsg.Substring(firstNewLine + 1)
                        End If
                        If cleanedMsg.EndsWith("```") Then
                            cleanedMsg = cleanedMsg.Substring(0, cleanedMsg.Length - 3)
                        End If
                        cleanedMsg = cleanedMsg.Trim()
                        Dim jsonStart = cleanedMsg.IndexOf("{")
                        Dim jsonEnd = cleanedMsg.LastIndexOf("}")
                        If jsonStart >= 0 AndAlso jsonEnd > jsonStart Then
                            cleanedMsg = cleanedMsg.Substring(jsonStart, jsonEnd - jsonStart + 1)
                        End If
                        Dim resultObj = JObject.Parse(cleanedMsg)
                        Dim completionsArray = resultObj("completions")
                        If completionsArray IsNot Nothing Then
                            For Each item In completionsArray
                                Dim c = item.ToString().Trim()
                                If Not String.IsNullOrWhiteSpace(c) Then completions.Add(c)
                            Next
                        End If
                    Catch parseEx As Exception
                        Debug.WriteLine($"解析补全JSON失败: {parseEx.Message}")
                        If Not String.IsNullOrWhiteSpace(msg) AndAlso msg.Length < 50 Then
                            completions.Add(msg.Trim())
                        End If
                    End Try
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionsWithChat 出错: {ex.Message}")
        End Try
        Return completions
    End Function

    ''' <summary>
    ''' 记录补全历史
    ''' </summary>
    Private Sub RecordCompletionHistory(inputText As String, completion As String, context As String)
        Try
            Dim historyPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                ConfigSettings.OfficeAiAppDataFolder,
                "autocomplete_history.json")

            Dim history As JObject
            If File.Exists(historyPath) Then
                Dim json = File.ReadAllText(historyPath)
                history = JObject.Parse(json)
            Else
                history = New JObject()
                history("version") = 1
                history("history") = New JArray()
            End If

            Dim historyArray = CType(history("history"), JArray)
            Dim existingItem = historyArray.FirstOrDefault(Function(item)
                                                               Return item("input")?.ToString() = inputText AndAlso
                                                                      item("completion")?.ToString() = completion
                                                           End Function)
            If existingItem IsNot Nothing Then
                existingItem("count") = existingItem("count").Value(Of Integer)() + 1
                existingItem("lastUsed") = DateTime.UtcNow.ToString("o")
            Else
                Dim newItem As New JObject()
                newItem("input") = inputText
                newItem("completion") = completion
                newItem("context") = context
                newItem("count") = 1
                newItem("lastUsed") = DateTime.UtcNow.ToString("o")
                historyArray.Add(newItem)
                While historyArray.Count > 100
                    historyArray.RemoveAt(0)
                End While
            End If

            Dim dir = Path.GetDirectoryName(historyPath)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            File.WriteAllText(historyPath, history.ToString(Newtonsoft.Json.Formatting.Indented))
        Catch ex As Exception
            Debug.WriteLine($"RecordCompletionHistory 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 根据 Office 应用类型获取场景化的补全系统提示词
    ''' </summary>
    Private Function GetCompletionSystemPrompt(appType As String) As String
        Dim baseRules = "
规则：
1. 只返回补全的剩余部分，不要重复用户已输入的内容
2. 返回JSON格式: {""completions"": [""补全1"", ""补全2"", ""补全3""]}
3. 最多返回3个候选
4. 补全应简洁，通常不超过20个字"

        Select Case appType.ToLower()
            Case "excel"
                Return $"你是Excel AI助手的输入补全引擎。根据用户当前输入和Excel上下文，预测用户想要的操作。

常见Excel场景补全示例：
- ""帮我"" → ""计算这列的总和"", ""筛选重复数据"", ""生成数据透视表""
- ""把"" → ""选中区域转换为表格"", ""这列数据去重"", ""A列和B列合并""
- ""统计"" → ""每个类别的数量"", ""销售额的平均值"", ""各月份的增长率""
- ""公式"" → ""计算两列的差值"", ""查找匹配的数据"", ""条件求和""
- ""格式"" → ""设置为货币格式"", ""添加条件格式"", ""调整列宽""
- ""图表"" → ""创建柱状图"", ""生成趋势线"", ""添加数据标签""
{baseRules}"

            Case "word"
                Return $"你是Word AI助手的输入补全引擎。根据用户当前输入和Word上下文，预测用户想要的操作。

常见Word场景补全示例：
- ""帮我"" → ""润色这段文字"", ""翻译选中内容"", ""生成文章大纲""
- ""把"" → ""这段改成正式语气"", ""标题设为一级标题"", ""段落缩进调整""
- ""总结"" → ""这篇文章的要点"", ""会议纪要"", ""核心观点""
- ""扩写"" → ""这个段落"", ""详细说明这个观点"", ""增加案例论证""
- ""格式"" → ""统一段落间距"", ""添加页眉页脚"", ""设置目录样式""
- ""检查"" → ""语法错误"", ""错别字"", ""标点符号""
{baseRules}"

            Case "powerpoint"
                Return $"你是PowerPoint AI助手的输入补全引擎。根据用户当前输入和PPT上下文，预测用户想要的操作。

常见PPT场景补全示例：
- ""帮我"" → ""美化这张幻灯片"", ""生成演讲稿"", ""添加过渡动画""
- ""把"" → ""文字转换为SmartArt"", ""图片裁剪为圆形"", ""背景改为渐变色""
- ""生成"" → ""项目汇报PPT"", ""产品介绍页"", ""团队介绍页""
- ""添加"" → ""图表展示数据"", ""时间线"", ""流程图""
- ""设计"" → ""统一字体样式"", ""配色方案"", ""母版布局""
- ""总结"" → ""演示要点"", ""关键数据"", ""结论页内容""
{baseRules}"

            Case Else
                Return $"你是Office AI助手的输入补全引擎。根据用户当前输入和Office上下文，预测用户想要输入的内容。
{baseRules}
5. 考虑Office上下文（选中内容、文档类型）"
        End Select
    End Function

End Class
