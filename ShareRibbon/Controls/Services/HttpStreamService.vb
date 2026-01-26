' ShareRibbon\Controls\Services\HttpStreamService.vb
' HTTP 流式请求服务：发送请求、处理流数据、MCP 工具调用

Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web
Imports Markdig
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' HTTP 流式请求服务，负责发送请求、处理流数据和 MCP 工具调用
''' </summary>
Public Class HttpStreamService
        Private ReadOnly _stateService As ChatStateService
        Private ReadOnly _getApplication As Func(Of ApplicationInfo)
        Private ReadOnly _executeScript As Func(Of String, Task)

        ' MCP 工具调用相关
        Private ReadOnly _pendingToolCalls As New Dictionary(Of String, JObject)()
        Private ReadOnly _completedToolCalls As New List(Of JObject)()

        ' 流处理状态
        Private _mainStreamCompleted As Boolean = False
        Private _pendingMcpTasks As Integer = 0
        Private _finalUuid As String = String.Empty
        Private _currentMarkdownBuffer As New StringBuilder()

        ' 停止标志
        Public Property StopStream As Boolean = False

        ''' <summary>
        ''' 流处理完成事件
        ''' </summary>
        Public Event StreamCompleted As EventHandler(Of String)

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        Public Sub New(stateService As ChatStateService, getApplication As Func(Of ApplicationInfo), executeScript As Func(Of String, Task))
            _stateService = stateService
            _getApplication = getApplication
            _executeScript = executeScript
        End Sub

#Region "发送请求"

        ''' <summary>
        ''' 发送流式 HTTP 请求
        ''' </summary>
        Public Async Function SendStreamRequestAsync(
            apiUrl As String,
            apiKey As String,
            requestBody As String,
            originQuestion As String,
            requestUuid As String,
            addHistory As Boolean,
            responseMode As String) As Task

            ' 生成响应 UUID
            Dim responseUuid As String = Guid.NewGuid().ToString()

            ' 保存映射
            _stateService.MapResponseToRequest(responseUuid, requestUuid)
            _stateService.SetResponseMode(responseUuid, responseMode)
            _stateService.MigrateSelectionToResponse(responseUuid, requestUuid)

            _finalUuid = responseUuid
            _mainStreamCompleted = False
            _pendingMcpTasks = 0
            _stateService.ResetSessionTokens()

            Try
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

                Using client As New HttpClient()
                    client.Timeout = System.Threading.Timeout.InfiniteTimeSpan

                    Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                    request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                    request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                    Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName

                    Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                        response.EnsureSuccessStatusCode()

                        ' 创建前端聊天节
                        Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
                        Await _executeScript(jsCreate)

                        ' 设置 requestId
                        Dim jsSetMapping As String = $"(function(){{ var el = document.getElementById('chat-{responseUuid}'); if(el) el.dataset.requestId = '{requestUuid}'; }})();"
                        Await _executeScript(jsSetMapping)

                        ' 处理流
                        Dim stringBuilder As New StringBuilder()
                        Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                            Using reader As New StreamReader(responseStream, Encoding.UTF8)
                                Dim buffer(102300) As Char
                                Dim readCount As Integer

                                Do
                                    If StopStream Then
                                        _currentMarkdownBuffer.Clear()
                                        _stateService.MarkdownBuffer.Clear()
                                        Exit Do
                                    End If

                                    readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                    If readCount = 0 Then Exit Do

                                    Dim chunk As String = New String(buffer, 0, readCount)
                                    chunk = chunk.Replace("data:", "")
                                    stringBuilder.Append(chunk)

                                    If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                        Await ProcessStreamChunkAsync(stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}), responseUuid, originQuestion)
                                        stringBuilder.Clear()
                                    End If
                                Loop
                            End Using
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                Throw
            Finally
                _mainStreamCompleted = True
                FinalizeStream(addHistory)
            End Try
        End Function

        ''' <summary>
        ''' 完成流处理
        ''' </summary>
        Private Sub FinalizeStream(addHistory As Boolean)
            Dim finalTokens As Integer = _stateService.CurrentSessionTotalTokens
            If _stateService.LastTokenInfo.HasValue Then
                finalTokens += _stateService.LastTokenInfo.Value.TotalTokens
                _stateService.AddTokens(_stateService.LastTokenInfo.Value.TotalTokens)
            End If

            CheckAndCompleteProcessing()

            If addHistory Then
                _stateService.AddMessage("assistant", $"这是大模型基于用户问题的答复作为历史参考：{_stateService.MarkdownBuffer.ToString()}")
            End If

            _stateService.ClearBuffers()
            _stateService.LastTokenInfo = Nothing
        End Sub

        ''' <summary>
        ''' 检查并完成处理
        ''' </summary>
        Private Sub CheckAndCompleteProcessing()
            If _mainStreamCompleted AndAlso _pendingMcpTasks = 0 Then
                _executeScript($"processStreamComplete('{_finalUuid}',{_stateService.CurrentSessionTotalTokens});")
                RaiseEvent StreamCompleted(Me, _finalUuid)
            End If
        End Sub

#End Region

#Region "流数据处理"

        ''' <summary>
        ''' 处理流数据块
        ''' </summary>
        Private Async Function ProcessStreamChunkAsync(rawChunk As String, uuid As String, originQuestion As String) As Task
            Try
                Dim lines As String() = rawChunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                For Each line In lines
                    line = line.Trim()

                    If line = "[DONE]" Then
                        If _pendingToolCalls.Count > 0 Then
                            Await ProcessCompletedToolCallsAsync(uuid, originQuestion)
                        End If
                        Await FlushBufferAsync("content", uuid)
                        Return
                    End If

                    If line = "" Then Continue For

                    Dim jsonObj As JObject = JObject.Parse(line)

                    ' 获取 token 信息
                    Dim usage = jsonObj("usage")
                    If usage IsNot Nothing AndAlso usage.Type = JTokenType.Object Then
                        _stateService.LastTokenInfo = New TokenInfo With {
                            .PromptTokens = CInt(usage("prompt_tokens")),
                            .CompletionTokens = CInt(usage("completion_tokens")),
                            .TotalTokens = CInt(usage("total_tokens"))
                        }
                    End If

                    ' 处理推理内容
                    Dim reasoning_content As String = jsonObj("choices")(0)("delta")("reasoning_content")?.ToString()
                    If Not String.IsNullOrEmpty(reasoning_content) Then
                        _currentMarkdownBuffer.Append(reasoning_content)
                        Await FlushBufferAsync("reasoning", uuid)
                    End If

                    ' 处理正文内容
                    Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()
                    If Not String.IsNullOrEmpty(content) Then
                        _currentMarkdownBuffer.Append(content)
                        Await FlushBufferAsync("content", uuid)
                    End If

                    ' 检查工具调用
                    Dim choices = jsonObj("choices")
                    If choices IsNot Nothing AndAlso choices.Count > 0 Then
                        Dim choice = choices(0)
                        Dim delta = choice("delta")
                        Dim finishReason = choice("finish_reason")?.ToString()

                        If delta IsNot Nothing Then
                            Dim toolCalls = delta("tool_calls")
                            If toolCalls IsNot Nothing AndAlso toolCalls.Count > 0 Then
                                CollectToolCallData(toolCalls)
                            End If
                        End If

                        If finishReason = "tool_calls" Then
                            Await ProcessCompletedToolCallsAsync(uuid, originQuestion)
                        End If
                    End If
                Next
            Catch ex As Exception
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 刷新缓冲区到前端
        ''' </summary>
        Private Async Function FlushBufferAsync(contentType As String, uuid As String) As Task
            If _currentMarkdownBuffer.Length = 0 Then Return

            Dim plainContent As String = _currentMarkdownBuffer.ToString()
            Dim escapedContent = HttpUtility.JavaScriptStringEncode(_currentMarkdownBuffer.ToString())
            _currentMarkdownBuffer.Clear()

            Dim js As String
            If contentType = "reasoning" Then
                js = $"appendReasoning('{uuid}','{escapedContent}');"
            Else
                js = $"appendRenderer('{uuid}','{escapedContent}');"
                _stateService.MarkdownBuffer.Append(escapedContent)
                _stateService.PlainMarkdownBuffer.Append(plainContent)
            End If

            Await _executeScript(js)
        End Function

#End Region

#Region "MCP 工具调用"

        ''' <summary>
        ''' 收集工具调用数据
        ''' </summary>
        Private Sub CollectToolCallData(toolCalls As JArray)
            Try
                For Each toolCall In toolCalls
                    Dim toolIndex = toolCall("index")?.Value(Of Integer)()
                    Dim toolId = toolCall("id")?.ToString()
                    Dim toolKey As String = $"tool_{toolIndex}"

                    If Not _pendingToolCalls.ContainsKey(toolKey) Then
                        _pendingToolCalls(toolKey) = New JObject()
                        _pendingToolCalls(toolKey)("realId") = If(String.IsNullOrEmpty(toolId), toolKey, toolId)
                        _pendingToolCalls(toolKey)("index") = toolIndex
                        _pendingToolCalls(toolKey)("type") = toolCall("type")?.ToString()
                        _pendingToolCalls(toolKey)("function") = New JObject()
                        _pendingToolCalls(toolKey)("function")("name") = ""
                        _pendingToolCalls(toolKey)("function")("arguments") = ""
                        _pendingToolCalls(toolKey)("processed") = False
                    End If

                    Dim currentTool = _pendingToolCalls(toolKey)

                    Dim functionName = toolCall("function")("name")?.ToString()
                    If Not String.IsNullOrEmpty(functionName) Then
                        currentTool("function")("name") = functionName
                    End If

                    Dim arguments = toolCall("function")("arguments")?.ToString()
                    If Not String.IsNullOrEmpty(arguments) Then
                        Dim currentArgs = currentTool("function")("arguments").ToString()
                        currentTool("function")("arguments") = currentArgs & arguments
                    End If
                Next
            Catch ex As Exception
            End Try
        End Sub

        ''' <summary>
        ''' 处理完成的工具调用
        ''' </summary>
        Private Async Function ProcessCompletedToolCallsAsync(uuid As String, originQuestion As String) As Task
            Try
                If _pendingToolCalls.Count = 0 Then Return

                For Each kvp In _pendingToolCalls
                    Dim toolCall = kvp.Value
                    Dim toolKey = kvp.Key

                    If CBool(toolCall("processed")) Then Continue For

                    Dim toolName = toolCall("function")("name").ToString()
                    Dim argumentsStr = toolCall("function")("arguments").ToString()

                    If String.IsNullOrEmpty(toolName) Then Continue For

                    toolCall("processed") = True

                    Dim argumentsObj As JObject = Nothing
                Dim parseError As Boolean = False
                Try
                    If Not String.IsNullOrEmpty(argumentsStr) Then
                        argumentsObj = JObject.Parse(argumentsStr)
                    Else
                        argumentsObj = New JObject()
                    End If
                Catch ex As Exception
                    parseError = True
                End Try

                If parseError Then
                    _currentMarkdownBuffer.Append($"<br/>**工具调用参数解析错误：**<br/>工具名称: {toolName}<br/>")
                    Await FlushBufferAsync("content", uuid)
                    Continue For
                End If

                _currentMarkdownBuffer.Append($"<br/>**正在调用工具: {toolName}**<br/>参数: `{argumentsObj.ToString(Newtonsoft.Json.Formatting.None)}`<br/>")
                Await FlushBufferAsync("content", uuid)

                ' 获取 MCP 连接
                Dim chatSettings As New ChatSettings(_getApplication())
                Dim enabledMcpList = chatSettings.EnabledMcpList

                If enabledMcpList IsNot Nothing AndAlso enabledMcpList.Count > 0 Then
                    Dim mcpConnectionName = enabledMcpList(0)
                    Dim result = Await HandleMcpToolCallAsync(toolName, argumentsObj, mcpConnectionName)

                    If result("isError") IsNot Nothing AndAlso CBool(result("isError")) Then
                        _currentMarkdownBuffer.Append($"<br/>**工具调用失败：**<br/>")
                        Await FlushBufferAsync("content", uuid)
                    Else
                        _pendingMcpTasks += 1
                        Await SendToolResultForFormattingAsync(toolName, argumentsObj, result, uuid, originQuestion)
                    End If
                Else
                    _currentMarkdownBuffer.Append("<br/>**配置错误：**<br/>没有启用的MCP连接<br/>")
                    Await FlushBufferAsync("content", uuid)
                End If
            Next

            _pendingToolCalls.Clear()
            _completedToolCalls.Clear()
        Catch ex As Exception
        End Try
    End Function

    ''' <summary>
    ''' 处理 MCP 工具调用
    ''' </summary>
    Private Async Function HandleMcpToolCallAsync(toolName As String, arguments As JObject, mcpConnectionName As String) As Task(Of JObject)
        Try
            Dim connections = MCPConnectionManager.LoadConnections()
            Dim connection = connections.FirstOrDefault(Function(c) c.Name = mcpConnectionName AndAlso c.IsActive)

            If connection Is Nothing Then
                Return CreateErrorResponse($"MCP连接 '{mcpConnectionName}' 未找到或未启用")
            End If

            Using client As New StreamJsonRpcMCPClient()
                Await client.ConfigureAsync(connection.Url)

                Dim initResult = Await client.InitializeAsync()
                If Not initResult.Success Then
                    Return CreateErrorResponse($"初始化MCP连接失败: {initResult.ErrorMessage}")
                End If

                Dim result = Await client.CallToolAsync(toolName, arguments)

                If result.IsError Then
                    Return CreateErrorResponse($"调用MCP工具失败: {result.ErrorMessage}")
                End If

                Dim responseObj = New JObject()
                Dim contentArray = New JArray()

                If result.Content IsNot Nothing Then
                    For Each content In result.Content
                        Dim contentObj = New JObject()
                        contentObj("type") = content.Type
                        If Not String.IsNullOrEmpty(content.Text) Then contentObj("text") = content.Text
                        If Not String.IsNullOrEmpty(content.Data) Then contentObj("data") = content.Data
                        If Not String.IsNullOrEmpty(content.MimeType) Then contentObj("mimeType") = content.MimeType
                        contentArray.Add(contentObj)
                    Next
                End If

                responseObj("content") = contentArray
                Return responseObj
            End Using
        Catch ex As Exception
            Return CreateErrorResponse($"MCP工具调用异常: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' 发送工具结果进行格式化
    ''' </summary>
    Private Async Function SendToolResultForFormattingAsync(toolName As String, arguments As JObject, result As JObject, uuid As String, originQuestion As String) As Task
        Dim hasError As Boolean = False
        Dim errorResultJson As String = String.Empty

        Try
            Dim promptBuilder As New StringBuilder()
            promptBuilder.AppendLine($"用户的原始问题：'{originQuestion}'，但用户使用了 MCP 工具 '{toolName}'，参数为：")
            promptBuilder.AppendLine("```json")
            promptBuilder.AppendLine(arguments.ToString(Newtonsoft.Json.Formatting.Indented))
            promptBuilder.AppendLine("```")
            promptBuilder.AppendLine()
            promptBuilder.AppendLine("工具执行结果为：")
            promptBuilder.AppendLine("```json")
            promptBuilder.AppendLine(result.ToString(Newtonsoft.Json.Formatting.Indented))
            promptBuilder.AppendLine("```")
            promptBuilder.AppendLine()
            promptBuilder.AppendLine("请将上述结果整理成易于理解的格式，使用Markdown呈现。")

            Dim messagesArray = New JArray()
            Dim systemMessage = New JObject()
            systemMessage("role") = "system"
            systemMessage("content") = "你是一个帮助解释API调用结果的助手。"

            Dim userMessage = New JObject()
            userMessage("role") = "user"
            userMessage("content") = promptBuilder.ToString()

            messagesArray.Add(systemMessage)
            messagesArray.Add(userMessage)

            Dim requestObj = New JObject()
            requestObj("model") = ConfigSettings.ModelName
            requestObj("messages") = messagesArray
            requestObj("stream") = True

            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)

            Using client As New HttpClient()
                client.Timeout = System.Threading.Timeout.InfiniteTimeSpan

                Dim request As New HttpRequestMessage(HttpMethod.Post, ConfigSettings.ApiUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", ConfigSettings.ApiKey)
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()

                    Dim formattedBuilder As New StringBuilder()
                    formattedBuilder.AppendLine("<br/>**工具调用结果：**<br/>")

                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Dim stringBuilder As New StringBuilder()
                            Dim buffer(1023) As Char
                            Dim readCount As Integer

                            Do
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do

                                Dim chunk As String = New String(buffer, 0, readCount)
                                chunk = chunk.Replace("data:", "")
                                stringBuilder.Append(chunk)

                                If stringBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    Dim lines As String() = stringBuilder.ToString().Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                                    For Each line In lines
                                        line = line.Trim()
                                        If line = "[DONE]" OrElse line = "" Then Continue For

                                        Try
                                            Dim jsonObj As JObject = JObject.Parse(line)
                                            Dim usage = jsonObj("usage")
                                            If usage IsNot Nothing Then
                                                _stateService.AddTokens(CInt(usage("total_tokens")))
                                            End If

                                            Dim content As String = jsonObj("choices")(0)("delta")("content")?.ToString()
                                            If Not String.IsNullOrEmpty(content) Then
                                                formattedBuilder.Append(content)
                                            End If
                                        Catch
                                        End Try
                                    Next

                                    stringBuilder.Clear()
                                End If
                            Loop
                        End Using
                    End Using

                    _currentMarkdownBuffer.Append(formattedBuilder.ToString())
                    Await FlushBufferAsync("content", uuid)
                End Using
            End Using
        Catch ex As Exception
            hasError = True
            errorResultJson = result.ToString(Newtonsoft.Json.Formatting.Indented)
        End Try

        ' Handle error outside of Catch block (Await not allowed in Catch)
        If hasError Then
            _currentMarkdownBuffer.Append($"{vbLf}{vbLf}**工具调用结果：**{vbLf}{vbLf}```json{vbLf}{errorResultJson}{vbLf}```{vbLf}")
            Await FlushBufferAsync("content", uuid)
        End If

        ' Cleanup
        _pendingMcpTasks -= 1
        CheckAndCompleteProcessing()
    End Function

        ''' <summary>
        ''' 创建错误响应
        ''' </summary>
        Private Function CreateErrorResponse(errorMessage As String) As JObject
            Dim responseObj = New JObject()
            responseObj("isError") = True
            responseObj("errorMessage") = errorMessage
            responseObj("timestamp") = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            Return responseObj
        End Function

#End Region

#Region "请求体构建"

        ''' <summary>
        ''' 创建请求体
        ''' </summary>
        Public Function CreateRequestBody(uuid As String, question As String, systemPrompt As String, addHistory As Boolean) As String
            Dim result As String = StripQuestion(question)
            Dim messages As New List(Of String)()

            Dim systemMessage = New HistoryMessage() With {
                .Role = "system",
                .Content = systemPrompt
            }

            Dim q = New HistoryMessage() With {
                .Role = "user",
                .Content = result
            }

            If addHistory Then
                _stateService.SetSystemMessage(systemPrompt)
                _stateService.AddMessage("user", result)

                For Each message In _stateService.HistoryMessages
                    Dim safeContent As String = If(message.Content, String.Empty)
                    safeContent = safeContent.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n")
                    messages.Add($"{{""role"": ""{message.Role}"", ""content"": ""{safeContent}""}}")
                Next
            Else
                Dim tempMessages As New List(Of HistoryMessage)()
                tempMessages.Add(systemMessage)
                tempMessages.Add(q)

                For Each message In tempMessages
                    Dim safeContent As String = If(message.Content, String.Empty)
                    safeContent = safeContent.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n")
                    messages.Add($"{{""role"": ""{message.Role}"", ""content"": ""{safeContent}""}}")
                Next
            End If

            ' 添加 MCP 工具
            Dim toolsArray As JArray = BuildToolsArray()
            Dim messagesJson = String.Join(",", messages)

            If toolsArray IsNot Nothing AndAlso toolsArray.Count > 0 Then
                Dim toolsJson = toolsArray.ToString(Newtonsoft.Json.Formatting.None)
                Return $"{{""model"": ""{ConfigSettings.ModelName}"", ""tools"": {toolsJson}, ""messages"": [{messagesJson}], ""stream"": true}}"
            Else
                Return $"{{""model"": ""{ConfigSettings.ModelName}"", ""messages"": [{messagesJson}], ""stream"": true}}"
            End If
        End Function

        ''' <summary>
        ''' 构建工具数组
        ''' </summary>
        Private Function BuildToolsArray() As JArray
            Dim toolsArray As JArray = Nothing
            Dim chatSettings As New ChatSettings(_getApplication())

            If chatSettings.EnabledMcpList IsNot Nothing AndAlso chatSettings.EnabledMcpList.Count > 0 Then
                toolsArray = New JArray()
                Dim connections = MCPConnectionManager.LoadConnections()

                For Each mcpName In chatSettings.EnabledMcpList
                    Dim connection = connections.FirstOrDefault(Function(c) c.Name = mcpName AndAlso c.IsActive)
                    If connection IsNot Nothing Then
                        If connection.Tools IsNot Nothing AndAlso connection.Tools.Count > 0 Then
                            For Each toolObj In connection.Tools
                                toolsArray.Add(toolObj)
                            Next
                        End If
                    End If
                Next
            End If

            Return toolsArray
        End Function

        ''' <summary>
        ''' 转义问题字符串
        ''' </summary>
        Private Function StripQuestion(question As String) As String
            Return question.Replace("\", "\\").Replace("""", "\""").
                          Replace(vbCr, "\r").Replace(vbLf, "\n").
                          Replace(vbTab, "\t").Replace(vbBack, "\b").
                          Replace(Chr(12), "\f")
        End Function

#End Region

    End Class
