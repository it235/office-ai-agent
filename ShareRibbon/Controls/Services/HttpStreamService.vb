' ShareRibbon\Controls\Services\HttpStreamService.vb
' HTTP 流式请求服务：发送请求、处理流数据、MCP 工具调用

Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web
Imports System.Diagnostics
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

    ' ReAct 工具调用循环
    Private _toolCallIterations As Integer = 0
    Private Const MAX_TOOL_CALL_ITERATIONS As Integer = 5
    Private _originalRequestMessages As JArray = Nothing

    ' 流处理状态
    Private _mainStreamCompleted As Boolean = False
    Private _catchException As Exception
    Private _pendingMcpTasks As Integer = 0
    Private _finalUuid As String = String.Empty
    Private _currentMarkdownBuffer As New StringBuilder()

    ' 停止标志
    Public Property StopStream As Boolean = False

    ' BaseChatControl 专用回调
    Private ReadOnly _waitForRendererMap As Func(Of String, Task)
    Private ReadOnly _onStreamComplete As Action(Of String, System.Text.StringBuilder)
    Private ReadOnly _onAgentContent As Action(Of String)
    Private ReadOnly _onAgentCompleted As Action

    ''' <summary>
    ''' 流完成后保存历史的回调（由调用方在每次 SendStreamRequestAsync 前设置）
    ''' 参数：(addHistory As Boolean, originQuestion As String)
    ''' </summary>
    Public Property FinalizeCallback As Action(Of Boolean, String)

    ''' <summary>最终响应 UUID（供 BaseChatControl 通过属性代理访问）</summary>
    Public ReadOnly Property FinalUuid As String
        Get
            Return _finalUuid
        End Get
    End Property

    ''' <summary>
    ''' 流处理完成事件
    ''' </summary>
    Public Event StreamCompleted As EventHandler(Of String)

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    Public Sub New(stateService As ChatStateService,
                       getApplication As Func(Of ApplicationInfo),
                       executeScript As Func(Of String, Task),
                       Optional waitForRendererMap As Func(Of String, Task) = Nothing,
                       Optional onStreamComplete As Action(Of String, System.Text.StringBuilder) = Nothing,
                       Optional onAgentContent As Action(Of String) = Nothing,
                       Optional onAgentCompleted As Action = Nothing)
        _stateService = stateService
        _getApplication = getApplication
        _executeScript = executeScript
        _waitForRendererMap = waitForRendererMap
        _onStreamComplete = onStreamComplete
        _onAgentContent = onAgentContent
        _onAgentCompleted = onAgentCompleted
    End Sub

#Region "发送请求"

    ''' <summary>
    ''' 检测模型是否需要非流式请求
    ''' </summary>
    Private Function NeedsNonStreamingRequest(modelName As String) As Boolean
        ' OpenAI o1 系列模型不支持流式请求
        Return modelName.StartsWith("o1-", StringComparison.OrdinalIgnoreCase) OrElse
                   modelName.Equals("o1", StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' 发送非流式 HTTP 请求（用于不支持流式的模型）
    ''' </summary>
    Private Async Function SendNonStreamingRequestAsync(
            apiUrl As String,
            apiKey As String,
            requestBody As String,
            originQuestion As String,
            requestUuid As String,
            addHistory As Boolean,
            responseMode As String,
            responseUuid As String) As Task

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromMinutes(5)

                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)

                ' 检测是否是 Anthropic API
                Dim isAnthropic As Boolean = apiUrl.Contains("anthropic.com")

                ' 设置认证头
                If isAnthropic Then
                    request.Headers.Add("x-api-key", apiKey)
                    request.Headers.Add("anthropic-version", "2023-06-01")
                    requestBody = ConvertToAnthropicFormatNonStreaming(requestBody)
                Else
                    request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                End If

                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName

                ' 创建前端聊天节
                Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
                Await _executeScript(jsCreate)

                ' 等待 rendererMap 就绪
                If _waitForRendererMap IsNot Nothing Then
                    Await _waitForRendererMap(responseUuid)
                End If

                ' 设置 requestId
                Dim jsSetMapping As String = $"(function(){{ var el = document.getElementById('chat-{responseUuid}'); if(el) el.dataset.requestId = '{requestUuid}'; }})();"
                Await _executeScript(jsSetMapping)

                ' 显示"正在思考"提示
                _currentMarkdownBuffer.Append("<br/>*正在思考中...*<br/>")
                Await FlushBufferAsync("content", responseUuid)

                ' 发送非流式请求
                Using response As HttpResponseMessage = Await client.SendAsync(request)
                    response.EnsureSuccessStatusCode()

                    Dim jsonContent As String = Await response.Content.ReadAsStringAsync()
                    Await ProcessNonStreamingResponseAsync(jsonContent, responseUuid, originQuestion, isAnthropic)
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[NonStreaming] 请求失败: {ex.Message}")
            _currentMarkdownBuffer.Append($"<br/>**请求失败: {ex.Message}**<br/>")
            _catchException = ex
        Finally
            _mainStreamCompleted = True
            FinalizeStream(addHistory, originQuestion)
        End Try

        If _catchException IsNot Nothing Then
            Await FlushBufferAsync("content", responseUuid)
            Dim rethrowEx = _catchException
            _catchException = Nothing
            Throw rethrowEx
        End If
    End Function

    ''' <summary>
    ''' 转换请求体为 Anthropic 非流式格式
    ''' </summary>
    Private Function ConvertToAnthropicFormatNonStreaming(requestBody As String) As String
        Try
            Dim json = JObject.Parse(requestBody)
            Dim anthropicBody = New JObject()

            ' 模型名称
            anthropicBody("model") = json("model")

            ' max_tokens 是必需的
            anthropicBody("max_tokens") = 4096

            ' 转换 tools（Anthropic 格式不同）
            If json("tools") IsNot Nothing Then
                Dim anthropicTools = New JArray()
                For Each tool In json("tools")
                    Dim tObj = New JObject()
                    tObj("name") = tool("function")("name")?.ToString()
                    tObj("description") = tool("function")("description")?.ToString()
                    If tool("function")("parameters") IsNot Nothing Then
                        tObj("input_schema") = tool("function")("parameters")
                    End If
                    anthropicTools.Add(tObj)
                Next
                anthropicBody("tools") = anthropicTools
            End If

            ' 转换 messages
            Dim messages = json("messages")
            If messages IsNot Nothing Then
                Dim newMessages = New JArray()
                Dim systemContent As String = ""

                For Each msg In messages
                    Dim role = msg("role")?.ToString()

                    ' Anthropic 不支持 system 角色在 messages 中，需要单独设置
                    If role = "system" Then
                        systemContent = msg("content")?.ToString()
                        Continue For
                    End If

                    ' 处理 assistant 的 tool_calls
                    If role = "assistant" AndAlso msg("tool_calls") IsNot Nothing Then
                        Dim contentArr = New JArray()
                        ' 如果有文本内容先添加
                        Dim textContent = msg("content")?.ToString()
                        If Not String.IsNullOrEmpty(textContent) Then
                            contentArr.Add(New JObject From {
                                    {"type", "text"},
                                    {"text", textContent}
                                })
                        End If
                        ' 添加 tool_use blocks
                        For Each tc In msg("tool_calls")
                            contentArr.Add(New JObject From {
                                    {"type", "tool_use"},
                                    {"id", tc("id")?.ToString()},
                                    {"name", tc("function")("name")?.ToString()},
                                    {"input", If(tc("function")("arguments") IsNot Nothing,
                                        JObject.Parse(tc("function")("arguments").ToString()),
                                        New JObject())}
                                })
                        Next
                        newMessages.Add(New JObject From {
                                {"role", "assistant"},
                                {"content", contentArr}
                            })
                        Continue For
                    End If

                    ' 处理 tool role（工具结果）
                    If role = "tool" Then
                        Dim contentArr = New JArray()
                        contentArr.Add(New JObject From {
                                {"type", "tool_result"},
                                {"tool_use_id", msg("tool_call_id")?.ToString()},
                                {"content", msg("content")?.ToString()}
                            })
                        newMessages.Add(New JObject From {
                                {"role", "user"},
                                {"content", contentArr}
                            })
                        Continue For
                    End If

                    ' 普通消息
                    newMessages.Add(New JObject From {
                            {"role", role},
                            {"content", msg("content")?.ToString()}
                        })
                Next

                anthropicBody("messages") = newMessages

                ' 设置 system 提示词
                If Not String.IsNullOrEmpty(systemContent) Then
                    anthropicBody("system") = systemContent
                End If
            End If

            ' 非流式输出
            anthropicBody("stream") = False

            Return anthropicBody.ToString(Newtonsoft.Json.Formatting.None)
        Catch ex As Exception
            Debug.WriteLine($"[Anthropic] 格式转换失败: {ex.Message}")
            Return requestBody
        End Try
    End Function

    ''' <summary>
    ''' 处理非流式响应
    ''' </summary>
    Private Async Function ProcessNonStreamingResponseAsync(
            jsonContent As String,
            uuid As String,
            originQuestion As String,
            isAnthropic As Boolean) As Task

        Try
            Dim jsonObj As JObject = JObject.Parse(jsonContent)

            Dim contentText As String = ""

            If isAnthropic Then
                ' 处理 Anthropic 格式响应
                If jsonObj("content") IsNot Nothing AndAlso jsonObj("content").Type = JTokenType.Array Then
                    For Each contentItem In jsonObj("content")
                        If contentItem("type")?.ToString() = "text" Then
                            contentText &= contentItem("text")?.ToString()
                        End If
                    Next
                End If
            Else
                ' 处理 OpenAI 格式响应
                If jsonObj("choices") IsNot Nothing AndAlso jsonObj("choices").Count > 0 Then
                    Dim choice = jsonObj("choices")(0)
                    If choice("message") IsNot Nothing Then
                        contentText = choice("message")("content")?.ToString()
                    End If
                End If
            End If

            ' 获取 token 信息
            If jsonObj("usage") IsNot Nothing AndAlso jsonObj("usage").Type = JTokenType.Object Then
                _stateService.LastTokenInfo = New TokenInfo With {
                        .PromptTokens = CInt(jsonObj("usage")("prompt_tokens")),
                        .CompletionTokens = CInt(jsonObj("usage")("completion_tokens")),
                        .TotalTokens = CInt(jsonObj("usage")("total_tokens"))
                    }
            End If

            ' 输出内容
            If Not String.IsNullOrEmpty(contentText) Then
                _currentMarkdownBuffer.Append(contentText)
                _stateService.PlainMarkdownBuffer.Append(contentText)
                Await FlushBufferAsync("content", uuid)
                _onAgentContent?.Invoke(contentText)
            End If

            _onAgentCompleted?.Invoke()

        Catch ex As Exception
            Debug.WriteLine($"[NonStreaming] 处理响应失败: {ex.Message}")
            _currentMarkdownBuffer.Append($"<br/>**处理响应失败: {ex.Message}**<br/>")
            _catchException = ex
        End Try

        If _catchException IsNot Nothing Then
            Await FlushBufferAsync("content", uuid)
            _catchException = Nothing
        End If
    End Function

    ''' <summary>
    ''' 发送流式 HTTP 请求（或非流式，取决于模型）
    ''' </summary>
    Public Async Function SendStreamRequestAsync(
            apiUrl As String,
            apiKey As String,
            requestBody As String,
            originQuestion As String,
            requestUuid As String,
            addHistory As Boolean,
            responseMode As String,
            Optional responseUuid As String = Nothing) As Task

        ' 生成响应 UUID（若调用方未指定则自动生成）
        If String.IsNullOrEmpty(responseUuid) Then
            responseUuid = Guid.NewGuid().ToString()
        End If

        ' 保存映射
        _stateService.MapResponseToRequest(responseUuid, requestUuid)
        _stateService.SetResponseMode(responseUuid, responseMode)
        _stateService.MigrateSelectionToResponse(responseUuid, requestUuid)

        _finalUuid = responseUuid
        _mainStreamCompleted = False
        _pendingMcpTasks = 0
        _toolCallIterations = 0
        _originalRequestMessages = Nothing
        _stateService.ResetSessionTokens()

        ' 检测模型是否需要非流式请求
        Dim useNonStreaming As Boolean = NeedsNonStreamingRequest(ConfigSettings.ModelName)
        If useNonStreaming Then
            Debug.WriteLine($"[HttpStream] 检测到 o1 模型，使用非流式请求")
            ' 修改请求体为非流式
            Try
                Dim reqJson = JObject.Parse(requestBody)
                reqJson("stream") = False
                requestBody = reqJson.ToString(Newtonsoft.Json.Formatting.None)
            Catch
            End Try
            Await SendNonStreamingRequestAsync(apiUrl, apiKey, requestBody, originQuestion,
                                                  requestUuid, addHistory, responseMode, responseUuid)
            Return
        End If

        ' 检测是否是 Anthropic API
        Dim isAnthropic As Boolean = apiUrl.Contains("anthropic.com")

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = System.Threading.Timeout.InfiniteTimeSpan

                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)

                ' Anthropic 使用不同的认证方式
                If isAnthropic Then
                    request.Headers.Add("x-api-key", apiKey)
                    request.Headers.Add("anthropic-version", "2023-06-01")
                    requestBody = ConvertToAnthropicFormat(requestBody)
                Else
                    request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                End If

                ' 保存原始请求消息，供 ReAct 工具调用循环使用
                Try
                    Dim reqJson = JObject.Parse(requestBody)
                    If reqJson("messages") IsNot Nothing Then
                        _originalRequestMessages = CType(reqJson("messages"), JArray)
                    End If
                Catch
                End Try

                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Dim aiName As String = ConfigSettings.platform & " " & ConfigSettings.ModelName

                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()

                    ' 创建前端聊天节
                    Dim jsCreate As String = $"createChatSection('{aiName}', formatDateTime(new Date()), '{responseUuid}');"
                    Await _executeScript(jsCreate)

                    ' 等待 rendererMap 就绪（BaseChatControl 注入的回调）
                    If _waitForRendererMap IsNot Nothing Then
                        Await _waitForRendererMap(responseUuid)
                    End If

                    ' 设置 requestId
                    Dim jsSetMapping As String = $"(function(){{ var el = document.getElementById('chat-{responseUuid}'); if(el) el.dataset.requestId = '{requestUuid}'; }})();"
                    Await _executeScript(jsSetMapping)

                    ' 处理流 - 按SSE规范逐行解析
                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Do
                                If StopStream Then
                                    _currentMarkdownBuffer.Clear()
                                    _stateService.MarkdownBuffer.Clear()
                                    Exit Do
                                End If

                                Dim line As String = Await reader.ReadLineAsync()
                                If line Is Nothing Then Exit Do

                                line = line.Trim()

                                ' 跳过空行和SSE注释
                                If String.IsNullOrEmpty(line) OrElse line.StartsWith(":") Then Continue Do

                                ' 处理 [DONE] 信号
                                If line = "data: [DONE]" OrElse line = "[DONE]" Then Exit Do

                                ' Anthropic 使用 SSE 格式，但数据格式不同
                                If isAnthropic Then
                                    Dim chunk As String = ProcessAnthropicChunk(line)
                                    If Not String.IsNullOrEmpty(chunk) Then
                                        Await ProcessStreamChunkAsync(chunk, responseUuid, originQuestion)
                                    End If
                                Else
                                    ' 标准SSE: 只处理 "data: " 开头的行
                                    Dim dataContent As String = Nothing
                                    If line.StartsWith("data: ") Then
                                        dataContent = line.Substring(6)
                                    ElseIf line.StartsWith("data:") Then
                                        dataContent = line.Substring(5).TrimStart()
                                    ElseIf line.StartsWith("{") Then
                                        ' 非SSE格式，直接是JSON（某些兼容API）
                                        ' dataContent = line
                                        ' End If
                                        dataContent = line
                                    End If

                                    If Not String.IsNullOrEmpty(dataContent) AndAlso dataContent <> "[DONE]" Then
                                        Await ProcessStreamChunkAsync(dataContent, responseUuid, originQuestion)
                                    End If
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
            FinalizeStream(addHistory, originQuestion)
        End Try
    End Function

    ''' <summary>
    ''' 转换请求体为 Anthropic 格式（支持 tool_calls / tool role）
    ''' </summary>
    Private Function ConvertToAnthropicFormat(requestBody As String) As String
        Try
            Dim json = JObject.Parse(requestBody)
            Dim anthropicBody = New JObject()

            ' 模型名称
            anthropicBody("model") = json("model")

            ' max_tokens 是必需的
            anthropicBody("max_tokens") = 4096

            ' 转换 tools（Anthropic 格式不同）
            If json("tools") IsNot Nothing Then
                Dim anthropicTools = New JArray()
                For Each tool In json("tools")
                    Dim tObj = New JObject()
                    tObj("name") = tool("function")("name")?.ToString()
                    tObj("description") = tool("function")("description")?.ToString()
                    If tool("function")("parameters") IsNot Nothing Then
                        tObj("input_schema") = tool("function")("parameters")
                    End If
                    anthropicTools.Add(tObj)
                Next
                anthropicBody("tools") = anthropicTools
            End If

            ' 转换 messages
            Dim messages = json("messages")
            If messages IsNot Nothing Then
                Dim newMessages = New JArray()
                Dim systemContent As String = ""

                For Each msg In messages
                    Dim role = msg("role")?.ToString()

                    ' Anthropic 不支持 system 角色在 messages 中，需要单独设置
                    If role = "system" Then
                        systemContent = msg("content")?.ToString()
                        Continue For
                    End If

                    ' 处理 assistant 的 tool_calls
                    If role = "assistant" AndAlso msg("tool_calls") IsNot Nothing Then
                        Dim contentArr = New JArray()
                        ' 如果有文本内容先添加
                        Dim textContent = msg("content")?.ToString()
                        If Not String.IsNullOrEmpty(textContent) Then
                            contentArr.Add(New JObject From {
                                    {"type", "text"},
                                    {"text", textContent}
                                })
                        End If
                        ' 添加 tool_use blocks
                        For Each tc In msg("tool_calls")
                            contentArr.Add(New JObject From {
                                    {"type", "tool_use"},
                                    {"id", tc("id")?.ToString()},
                                    {"name", tc("function")("name")?.ToString()},
                                    {"input", If(tc("function")("arguments") IsNot Nothing,
                                        JObject.Parse(tc("function")("arguments").ToString()),
                                        New JObject())}
                                })
                        Next
                        newMessages.Add(New JObject From {
                                {"role", "assistant"},
                                {"content", contentArr}
                            })
                        Continue For
                    End If

                    ' 处理 tool role（工具结果）
                    If role = "tool" Then
                        Dim contentArr = New JArray()
                        contentArr.Add(New JObject From {
                                {"type", "tool_result"},
                                {"tool_use_id", msg("tool_call_id")?.ToString()},
                                {"content", msg("content")?.ToString()}
                            })
                        newMessages.Add(New JObject From {
                                {"role", "user"},
                                {"content", contentArr}
                            })
                        Continue For
                    End If

                    ' 普通消息
                    newMessages.Add(New JObject From {
                            {"role", role},
                            {"content", msg("content")?.ToString()}
                        })
                Next

                anthropicBody("messages") = newMessages

                ' 设置 system 提示词
                If Not String.IsNullOrEmpty(systemContent) Then
                    anthropicBody("system") = systemContent
                End If
            End If

            ' 流式输出
            anthropicBody("stream") = True

            Return anthropicBody.ToString(Newtonsoft.Json.Formatting.None)
        Catch ex As Exception
            Debug.WriteLine($"[Anthropic] 格式转换失败: {ex.Message}")
            Return requestBody
        End Try
    End Function

    ''' <summary>
    ''' 处理 Anthropic SSE 数据块，转换为 OpenAI 兼容格式
    ''' </summary>
    Private Function ProcessAnthropicChunk(chunk As String) As String
        Try
            ' Anthropic SSE 格式: event: xxx\ndata: {...}
            Dim result = New StringBuilder()
            Dim lines = chunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                If line.StartsWith("data:") Then
                    Dim dataContent = line.Substring(5).Trim()
                    If dataContent = "[DONE]" Then
                        result.AppendLine("[DONE]")
                        Continue For
                    End If

                    Try
                        Dim anthropicJson = JObject.Parse(dataContent)
                        Dim eventType = anthropicJson("type")?.ToString()

                        Select Case eventType
                            Case "content_block_delta"
                                ' 转换为 OpenAI 格式
                                Dim delta = anthropicJson("delta")
                                If delta IsNot Nothing Then
                                    Dim text = delta("text")?.ToString()
                                    If Not String.IsNullOrEmpty(text) Then
                                        Dim openaiFormat = New JObject From {
                                                {"choices", New JArray From {
                                                    New JObject From {
                                                        {"delta", New JObject From {
                                                            {"content", text}
                                                        }}
                                                    }
                                                }}
                                            }
                                        result.AppendLine(openaiFormat.ToString(Newtonsoft.Json.Formatting.None))
                                    End If
                                End If

                            Case "message_stop"
                                result.AppendLine("[DONE]")

                            Case "message_delta"
                                ' 包含 usage 信息
                                Dim usage = anthropicJson("usage")
                                If usage IsNot Nothing Then
                                    Dim openaiFormat = New JObject From {
                                            {"choices", New JArray From {
                                                New JObject From {
                                                    {"delta", New JObject()}
                                                }
                                            }},
                                            {"usage", usage}
                                        }
                                    result.AppendLine(openaiFormat.ToString(Newtonsoft.Json.Formatting.None))
                                End If
                        End Select
                    Catch
                        ' 解析失败，跳过
                    End Try
                End If
            Next

            Return result.ToString()
        Catch ex As Exception
            Return chunk.Replace("data:", "")
        End Try
    End Function

    ''' <summary>
    ''' 完成流处理
    ''' </summary>
    Private Sub FinalizeStream(addHistory As Boolean, Optional originQuestion As String = "")
        If _stateService.LastTokenInfo.HasValue Then
            _stateService.AddTokens(_stateService.LastTokenInfo.Value.TotalTokens)
        End If

        CheckAndCompleteProcessing()

        If FinalizeCallback IsNot Nothing Then
            ' BaseChatControl 负责完整的历史保存逻辑
            FinalizeCallback.Invoke(addHistory, originQuestion)
        ElseIf addHistory Then
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
            _onStreamComplete?.Invoke(_finalUuid, _stateService.PlainMarkdownBuffer)
            ' 确保 Agent 响应完成回调被触发（防止 [DONE] 消息未被正确检测的情况）
            _onAgentCompleted?.Invoke()
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
                    ' Agent 响应完成回调
                    _onAgentCompleted?.Invoke()
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
                Dim reasoning_content As String = Nothing
                If jsonObj("choices") IsNot Nothing AndAlso jsonObj("choices").Count > 0 Then
                    reasoning_content = jsonObj("choices")(0)("delta")("reasoning_content")?.ToString()
                End If

                If Not String.IsNullOrEmpty(reasoning_content) Then
                    _currentMarkdownBuffer.Append(reasoning_content)
                    Await FlushBufferAsync("reasoning", uuid)
                End If

                ' 处理正文内容
                Dim content As String = Nothing
                If jsonObj("choices") IsNot Nothing AndAlso jsonObj("choices").Count > 0 Then
                    content = jsonObj("choices")(0)("delta")("content")?.ToString()
                End If
                If Not String.IsNullOrEmpty(content) Then
                    _currentMarkdownBuffer.Append(content)
                    Await FlushBufferAsync("content", uuid)
                    ' Agent 响应收集回调
                    _onAgentContent?.Invoke(content)
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
    ''' 处理完成的工具调用：收集所有工具结果后，一次性回注模型继续推理（ReAct 循环）
    ''' </summary>
    Private Async Function ProcessCompletedToolCallsAsync(uuid As String, originQuestion As String) As Task
        Try
            If _pendingToolCalls.Count = 0 Then Return

            Dim allToolCalls As New List(Of JObject)()
            Dim allToolResults As New List(Of JObject)()

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
                    ' 尝试所有启用的MCP连接，优先使用工具所属的连接
                    Dim mcpConnectionName As String = Nothing
                    For Each connName In enabledMcpList
                        Dim connections = MCPConnectionManager.LoadConnections()
                        Dim conn = connections.FirstOrDefault(Function(c) c.Name = connName AndAlso c.IsActive)
                        If conn IsNot Nothing Then
                            Dim client As StreamJsonRpcMCPClient = Nothing
                            Try
                                client = New StreamJsonRpcMCPClient()
                                Await client.ConfigureAsync(conn.Url)
                                Dim initResult = Await client.InitializeAsync()
                                If initResult.Success Then
                                    Dim tools = Await client.ListToolsAsync()
                                    If tools IsNot Nothing AndAlso tools.Any(Function(t) t.Name = toolName) Then
                                        mcpConnectionName = connName
                                        client.Dispose()
                                        Exit For
                                    End If
                                End If
                            Catch
                            Finally
                                If client IsNot Nothing Then client.Dispose()
                            End Try
                        End If
                    Next

                    ' 如果没找到包含该工具的连接，使用第一个
                    If mcpConnectionName Is Nothing Then
                        mcpConnectionName = enabledMcpList(0)
                    End If

                    Dim result = Await HandleMcpToolCallAsync(toolName, argumentsObj, mcpConnectionName)

                    allToolCalls.Add(toolCall)
                    allToolResults.Add(result)

                    If result("isError") IsNot Nothing AndAlso CBool(result("isError")) Then
                        _currentMarkdownBuffer.Append($"<br/>**工具调用失败：**<br/>")
                        Await FlushBufferAsync("content", uuid)
                    End If
                Else
                    _currentMarkdownBuffer.Append("<br/>**配置错误：**<br/>没有启用的MCP连接<br/>")
                    Await FlushBufferAsync("content", uuid)
                End If
            Next

            ' 收集完毕后，一次性将所有工具结果回注模型（ReAct 循环）
            If allToolCalls.Count > 0 AndAlso allToolResults.Count > 0 Then
                _pendingMcpTasks += 1
                Dim isAnthropic As Boolean = ConfigSettings.ApiUrl.Contains("anthropic.com")
                Await SendToolResultForReActAsync(allToolCalls, allToolResults, uuid, originQuestion,
                                                      ConfigSettings.ApiUrl, ConfigSettings.ApiKey, isAnthropic)
            End If

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
                    For Each contentItem In result.Content
                        Dim contentObj = New JObject()
                        contentObj("type") = contentItem.Type
                        If Not String.IsNullOrEmpty(contentItem.Text) Then contentObj("text") = contentItem.Text
                        If Not String.IsNullOrEmpty(contentItem.Data) Then contentObj("data") = contentItem.Data
                        If Not String.IsNullOrEmpty(contentItem.MimeType) Then contentObj("mimeType") = contentItem.MimeType
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
    ''' ReAct 工具结果回注：将工具执行结果以 tool role 消息回注到原始对话，
    ''' 让模型继续推理，实现 Thought->Action->Observation 循环
    ''' </summary>
    Private Async Function SendToolResultForReActAsync(
            allToolCalls As List(Of JObject),
            allToolResults As List(Of JObject),
            uuid As String,
            originQuestion As String,
            apiUrl As String,
            apiKey As String,
            isAnthropic As Boolean) As Task

        If _toolCallIterations >= MAX_TOOL_CALL_ITERATIONS Then
            _currentMarkdownBuffer.Append("<br/>**工具调用已达最大迭代次数，停止循环。**<br/>")
            Await FlushBufferAsync("content", uuid)
            Return
        End If

        _toolCallIterations += 1
        Debug.WriteLine($"[ReAct] 开始第 {_toolCallIterations} 轮工具结果回注，共 {allToolCalls.Count} 个工具调用")

        Try
            ' 构建回注消息：原始消息 + assistant的tool_calls + tool结果
            Dim messagesArray As JArray

            If _originalRequestMessages IsNot Nothing Then
                messagesArray = CType(_originalRequestMessages.DeepClone(), JArray)
            Else
                ' 回退：构建最小上下文
                messagesArray = New JArray()
                Dim sysMsg = New JObject()
                sysMsg("role") = "system"
                sysMsg("content") = "你是一个智能助手，能够使用工具来帮助用户解决问题。请根据工具执行结果继续推理或给出最终回答。"
                messagesArray.Add(sysMsg)
                Dim userMsg = New JObject()
                userMsg("role") = "user"
                userMsg("content") = originQuestion
                messagesArray.Add(userMsg)
            End If

            ' 添加 assistant 的 tool_calls 消息
            Dim assistantMsg = New JObject()
            assistantMsg("role") = "assistant"
            assistantMsg("content") = Nothing
            Dim toolCallsArr = New JArray()
            For Each tc In allToolCalls
                Dim tcObj = New JObject()
                tcObj("id") = tc("realId")?.ToString()
                tcObj("type") = "function"
                Dim funcObj = New JObject()
                funcObj("name") = tc("function")("name")?.ToString()
                funcObj("arguments") = tc("function")("arguments")?.ToString()
                tcObj("function") = funcObj
                toolCallsArr.Add(tcObj)
            Next
            assistantMsg("tool_calls") = toolCallsArr
            messagesArray.Add(assistantMsg)

            ' 添加每个工具的结果（tool role）
            For i = 0 To allToolResults.Count - 1
                Dim toolResultMsg = New JObject()
                toolResultMsg("role") = "tool"
                toolResultMsg("tool_call_id") = allToolCalls(i)("realId")?.ToString()
                Dim resultContent = allToolResults(i).ToString(Newtonsoft.Json.Formatting.None)
                toolResultMsg("content") = resultContent
                messagesArray.Add(toolResultMsg)
            Next

            ' 构建请求
            Dim requestObj = New JObject()
            requestObj("model") = ConfigSettings.ModelName
            requestObj("messages") = messagesArray
            requestObj("stream") = True

            Dim toolsArray = BuildToolsArray()
            If toolsArray IsNot Nothing AndAlso toolsArray.Count > 0 Then
                requestObj("tools") = toolsArray
            End If

            Dim requestBody = requestObj.ToString(Newtonsoft.Json.Formatting.None)

            ' Anthropic 格式转换
            If isAnthropic Then
                requestBody = ConvertToAnthropicFormat(requestBody)
            End If

            _currentMarkdownBuffer.Append($"<br/>**工具执行完成，让AI继续推理...（第{_toolCallIterations}轮）**<br/>")
            Await FlushBufferAsync("content", uuid)

            ' 发送请求并流式处理
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = System.Threading.Timeout.InfiniteTimeSpan

                Dim request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                If isAnthropic Then
                    request.Headers.Add("x-api-key", apiKey)
                    request.Headers.Add("anthropic-version", "2023-06-01")
                Else
                    request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                End If
                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Using response As HttpResponseMessage = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()

                    Using responseStream As Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New StreamReader(responseStream, Encoding.UTF8)
                            Do
                                If StopStream Then
                                    _currentMarkdownBuffer.Clear()
                                    _stateService.MarkdownBuffer.Clear()
                                    Exit Do
                                End If

                                Dim line As String = Await reader.ReadLineAsync()
                                If line Is Nothing Then Exit Do
                                line = line.Trim()

                                If String.IsNullOrEmpty(line) OrElse line.StartsWith(":") Then Continue Do
                                If line = "data: [DONE]" OrElse line = "[DONE]" Then Exit Do

                                If isAnthropic Then
                                    Dim chunk As String = ProcessAnthropicChunk(line)
                                    If Not String.IsNullOrEmpty(chunk) Then
                                        Await ProcessStreamChunkAsync(chunk, uuid, originQuestion)
                                    End If
                                Else
                                    Dim dataContent As String = Nothing
                                    If line.StartsWith("data: ") Then
                                        dataContent = line.Substring(6)
                                    ElseIf line.StartsWith("data:") Then
                                        dataContent = line.Substring(5).TrimStart()
                                    ElseIf line.StartsWith("{") Then
                                        dataContent = line
                                    End If

                                    If Not String.IsNullOrEmpty(dataContent) AndAlso dataContent <> "[DONE]" Then
                                        Await ProcessStreamChunkAsync(dataContent, uuid, originQuestion)
                                    End If
                                End If
                            Loop
                        End Using
                    End Using
                End Using
            End Using

            Debug.WriteLine($"[ReAct] 第 {_toolCallIterations} 轮完成")
        Catch ex As Exception
            Debug.WriteLine($"[ReAct] 工具结果回注失败: {ex.Message}")
            _currentMarkdownBuffer.Append($"<br/>**工具结果回注失败: {ex.Message}**<br/>")
            _catchException = ex
        End Try

        If _catchException IsNot Nothing Then
            Await FlushBufferAsync("content", uuid)
            _catchException = Nothing
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
                .role = "system",
                .content = systemPrompt
            }

        Dim q = New HistoryMessage() With {
                .role = "user",
                .content = result
            }

        If addHistory Then
            _stateService.SetSystemMessage(systemPrompt)
            _stateService.AddMessage("user", result)

            For Each message In _stateService.HistoryMessages
                Dim safeContent As String = If(message.content, String.Empty)
                safeContent = safeContent.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n")
                messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{safeContent}""}}")
            Next
        Else
            Dim tempMessages As New List(Of HistoryMessage)()
            tempMessages.Add(systemMessage)
            tempMessages.Add(q)

            For Each message In tempMessages
                Dim safeContent As String = If(message.content, String.Empty)
                safeContent = safeContent.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n")
                messages.Add($"{{""role"": ""{message.role}"", ""content"": ""{safeContent}""}}")
            Next
        End If

        ' 添加 MCP 工具
        Dim toolsArray = BuildToolsArray()
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
