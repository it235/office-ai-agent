Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports StreamJsonRpc

Public Enum MCPTransportType
    SSE
    Stdio
End Enum

Public Class StreamJsonRpcMCPClient
    Implements IDisposable

    Private _jsonRpc As JsonRpc
    Private _httpClient As HttpClient
    Private _serverUrl As String
    Private _apiKey As String
    Private _isInitialized As Boolean = False
    Private _serverCapabilities As MCPServerCapabilities
    Private _transportType As MCPTransportType
    Private _stdioProcess As Process

    Public Sub New()
        _httpClient = New HttpClient()
        _httpClient.Timeout = TimeSpan.FromSeconds(30)
    End Sub

    ' 配置客户端 - 支持SSE和Stdio
    Public Async Function ConfigureAsync(serverUrl As String, Optional apiKey As String = Nothing) As Task
        _serverUrl = serverUrl
        _apiKey = apiKey

        ' 判断传输类型
        If _serverUrl.StartsWith("stdio://") Then
            _transportType = MCPTransportType.Stdio
            Await SetupStdioTransportAsync()
        Else
            _transportType = MCPTransportType.SSE
            SetupSSETransport()
        End If
    End Function
    ' 修改 SetupStdioTransportAsync 方法中的进程配置部分
    Private Async Function SetupStdioTransportAsync() As Task
        Try
            Dim options = StdioOptions.Parse(_serverUrl)

            ' 解析命令和参数
            _stdioProcess = New Process()

            ' 更详细的日志记录
            Debug.WriteLine($"Stdio 命令: {options.Command}")
            Debug.WriteLine($"Stdio 参数: {options.Arguments}")

            ' 对于 Node.js 脚本，需要特别处理
            If options.Command.EndsWith(".js") Then
                ' 直接使用 node 命令执行 js 文件
                _stdioProcess.StartInfo.FileName = "node"
                ' 直接将脚本路径作为参数传递，不再拼接参数
                _stdioProcess.StartInfo.Arguments = $"\""{options.Command}\"" {options.Arguments}"
                Debug.WriteLine($"启动 Node.js 脚本: {_stdioProcess.StartInfo.FileName} {_stdioProcess.StartInfo.Arguments}")
            ElseIf options.Command.EndsWith(".py") Then
                ' Python 脚本
                _stdioProcess.StartInfo.FileName = "python"
                _stdioProcess.StartInfo.Arguments = $"\""{options.Command}\"" {options.Arguments}"
            ElseIf options.Command.Contains("npx") Then
                ' NPX 命令 - 使用 cmd /c 启动
                _stdioProcess.StartInfo.FileName = "cmd.exe"
                _stdioProcess.StartInfo.Arguments = $"/c {options.Command} {options.Arguments}"
            Else
                ' 直接执行命令
                _stdioProcess.StartInfo.FileName = options.Command
                _stdioProcess.StartInfo.Arguments = options.Arguments
            End If

            ' 设置工作目录
            If Not String.IsNullOrEmpty(options.WorkingDirectory) Then
                _stdioProcess.StartInfo.WorkingDirectory = options.WorkingDirectory
            End If

            ' 设置环境变量
            For Each kvp In options.EnvironmentVariables
                _stdioProcess.StartInfo.EnvironmentVariables(kvp.Key) = kvp.Value
            Next

            ' 标准进程配置 - 关键修改：添加 UTF-8 编码
            _stdioProcess.StartInfo.UseShellExecute = False
            _stdioProcess.StartInfo.RedirectStandardInput = True
            _stdioProcess.StartInfo.RedirectStandardOutput = True
            _stdioProcess.StartInfo.RedirectStandardError = True
            _stdioProcess.StartInfo.CreateNoWindow = True
            _stdioProcess.StartInfo.StandardOutputEncoding = Encoding.UTF8  ' 使用 UTF-8 编码
            _stdioProcess.StartInfo.StandardErrorEncoding = Encoding.UTF8   ' 使用 UTF-8 编码

            ' 启动进程
            Debug.WriteLine($"启动进程: {_stdioProcess.StartInfo.FileName} {_stdioProcess.StartInfo.Arguments}")
            _stdioProcess.Start()

            ' 记录错误输出
            AddHandler _stdioProcess.ErrorDataReceived, Sub(sender, e)
                                                            If Not String.IsNullOrEmpty(e.Data) Then
                                                                Debug.WriteLine($"[MCP Process Error] {e.Data}")
                                                            End If
                                                        End Sub

            _stdioProcess.BeginErrorReadLine()

            ' 等待进程准备就绪
            Await Task.Delay(500)

            Debug.WriteLine("Stdio 传输设置完成")
        Catch ex As Exception
            Debug.WriteLine($"设置 stdio 传输失败: {ex.Message}")
            Debug.WriteLine($"异常详情: {ex.ToString()}")

            ' 确保在发生错误时清理资源
            If _stdioProcess IsNot Nothing AndAlso Not _stdioProcess.HasExited Then
                Try
                    _stdioProcess.Kill()
                Catch
                    ' 忽略清理错误
                End Try
            End If

            Throw
        End Try
    End Function


    ' 添加新的 Stdio 初始化方法
    Private Async Function InitializeStdioAsync(initParams As Object) As Task(Of MCPInitResponse)
        Try
            ' 构建 JSON-RPC 请求
            Dim request = New JObject()
            request("jsonrpc") = "2.0"
            request("id") = "init-" + Guid.NewGuid().ToString()
            request("method") = "initialize"
            request("params") = JObject.FromObject(initParams)

            ' 发送请求并获取响应
            Dim response = Await SendStdioRequestAsync(request)

            ' 处理初始化响应
            If response("result") IsNot Nothing Then
                Return ProcessInitResponse(response("result").ToObject(Of JObject)())
            ElseIf response("error") IsNot Nothing Then
                Dim errorMsg = response("error")("message")?.ToString()
                Return New MCPInitResponse() With {
                .Success = False,
                .ErrorMessage = errorMsg
            }
            Else
                Return New MCPInitResponse() With {
                .Success = False,
                .ErrorMessage = "Invalid response format"
            }
            End If
        Catch ex As Exception
            Debug.WriteLine($"Stdio初始化失败: {ex.Message}")
            Return New MCPInitResponse() With {
            .Success = False,
            .ErrorMessage = $"Stdio initialization failed: {ex.Message}"
        }
        End Try
    End Function


    ' 修改 SendStdioRequestAsync 方法，确保使用 UTF-8 编码
    Private Async Function SendStdioRequestAsync(request As JObject) As Task(Of JObject)
        If _stdioProcess Is Nothing OrElse _stdioProcess.HasExited Then
            Throw New InvalidOperationException("进程未启动或已退出")
        End If

        ' 添加行尾换行符，确保请求完整发送
        Dim requestString = request.ToString(Formatting.None) & Environment.NewLine

        ' 明确使用 UTF-8 编码写入
        Dim requestBytes = Encoding.UTF8.GetBytes(requestString)
        Await _stdioProcess.StandardInput.BaseStream.WriteAsync(requestBytes, 0, requestBytes.Length)
        Await _stdioProcess.StandardInput.BaseStream.FlushAsync()

        Debug.WriteLine($"发送请求: {requestString.Trim()}")

        ' 从标准输出读取响应
        Dim responseString = Await _stdioProcess.StandardOutput.ReadLineAsync()

        If String.IsNullOrEmpty(responseString) Then
            Throw New Exception("No response received from process")
        End If

        Debug.WriteLine($"收到响应: {responseString}")

        ' 解析响应
        Try
            Return JObject.Parse(responseString)
        Catch jsonEx As JsonException
            Debug.WriteLine($"JSON解析错误: {jsonEx.Message}")

            ' 尝试处理可能的非标准响应
            Dim fallbackResponse = New JObject()
            fallbackResponse("result") = New JObject()
            fallbackResponse("result")("content") = New JArray(New JObject From {
            {"type", "text"},
            {"text", responseString}
        })

            Return fallbackResponse
        End Try
    End Function

    ' 修改 InvokeMethodAsync 方法，适应新的 Stdio 实现
    Private Async Function InvokeMethodAsync(Of T)(method As String, ParamArray arguments() As Object) As Task(Of T)
        EnsureInitialized()

        Debug.WriteLine($"调用方法: {method}")

        Try
            Select Case _transportType
                Case MCPTransportType.Stdio
                    ' 使用直接的 Stdio 通信
                    Dim request = New JObject()
                    request("jsonrpc") = "2.0"
                    request("id") = Guid.NewGuid().ToString()
                    request("method") = method

                    If arguments.Length > 0 AndAlso arguments(0) IsNot Nothing Then
                        request("params") = JToken.FromObject(arguments(0))
                    End If

                    ' 发送请求
                    Debug.WriteLine($"发送Stdio请求: {request.ToString(Formatting.None)}")
                    Dim response = Await SendStdioRequestAsync(request)
                    Debug.WriteLine($"收到Stdio响应: {response.ToString(Formatting.None)}")

                    ' 检查错误
                    If response("error") IsNot Nothing Then
                        Dim errorMsg = response("error")("message")?.ToString()
                        Debug.WriteLine($"Stdio响应包含错误: {errorMsg}")
                        Throw New Exception($"服务器错误: {errorMsg}")
                    End If

                    ' 返回结果
                    Debug.WriteLine("Stdio请求成功完成")
                    Return response("result").ToObject(Of T)()
                Case Else
                    Debug.WriteLine("使用SSE方法调用")
                    Return Await InvokeSSEMethodAsync(Of T)(method, arguments)
            End Select
        Catch ex As Exception
            Debug.WriteLine($"方法调用异常: {ex.ToString()}")
            Throw
        End Try
    End Function

    ' 在 Dispose 方法中确保正确清理 Stdio 资源
    Public Sub Dispose() Implements IDisposable.Dispose
        _jsonRpc?.Dispose()
        _httpClient?.Dispose()

        If _stdioProcess IsNot Nothing Then
            Try
                ' 尝试发送退出请求
                Try
                    If Not _stdioProcess.HasExited Then
                        ' 发送 shutdown 请求
                        Dim request = New JObject()
                        request("jsonrpc") = "2.0"
                        request("id") = "shutdown"
                        request("method") = "shutdown"
                        _stdioProcess.StandardInput.WriteLine(request.ToString(Formatting.None))
                        _stdioProcess.StandardInput.Flush()

                        ' 给进程一点时间来处理
                        Thread.Sleep(500)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"发送退出请求失败: {ex.Message}")
                End Try

                ' 终止进程
                If Not _stdioProcess.HasExited Then
                    _stdioProcess.Kill()
                End If

                _stdioProcess.Dispose()
            Catch ex As Exception
                Debug.WriteLine($"清理进程资源失败: {ex.Message}")
            End Try
        End If

        _isInitialized = False
    End Sub

    ' 设置SSE传输
    Private Sub SetupSSETransport()
        _httpClient.DefaultRequestHeaders.Clear()
        _httpClient.DefaultRequestHeaders.Add("Accept", "application/json, text/event-stream")
        _httpClient.DefaultRequestHeaders.Add("Cache-Control", "no-cache")
        _httpClient.DefaultRequestHeaders.Add("User-Agent", "VSTO-MCP-Client/1.0")

        ' 检查 URL 中是否已包含 API 密钥
        Dim hasApiKeyInUrl = _serverUrl.Contains("api_key=") OrElse
                             _serverUrl.Contains("apikey=") OrElse
                             _serverUrl.Contains("key=") OrElse
                             _serverUrl.Contains("token=") OrElse
                             _serverUrl.Contains("access_token=") OrElse
                             _serverUrl.Contains("auth=") OrElse
                             _serverUrl.Contains("ak=")

        ' 只有当显式提供 API Key 且 URL 中不包含密钥时，才添加 Authorization 头
        If Not String.IsNullOrEmpty(_apiKey) AndAlso Not hasApiKeyInUrl Then
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}")
        End If
    End Sub

    Public Property InitResponse As MCPInitResponse

    ' 初始化MCP连接
    Public Async Function InitializeAsync() As Task(Of MCPInitResponse)
        Try
            Dim initParams = New With {
                .protocolVersion = "2024-11-05",
                .capabilities = New With {
                    .roots = New With {.listChanged = True},
                    .sampling = New Object(),
                    .experimental = New Object()
                },
                .clientInfo = New With {
                    .name = "VSTO-MCP-Client",
                    .version = "1.0.0"
                }
            }

            Dim response As MCPInitResponse
            Select Case _transportType
                Case MCPTransportType.Stdio
                    response = Await InitializeStdioAsync(initParams)
                Case Else
                    response = Await InitializeSSEAsync(initParams)
            End Select

            ' 保存响应
            Me.InitResponse = response
            Return response

        Catch ex As Exception
            Debug.WriteLine($"初始化MCP连接失败: {ex.Message}")
            Return New MCPInitResponse() With {
                .Success = False,
                .ErrorMessage = ex.Message
            }
        End Try
    End Function


    ' SSE初始化 - 标准MCP实现
    Private Async Function InitializeSSEAsync(initParams As Object) As Task(Of MCPInitResponse)
        Try
            ' 构建请求
            Dim requestUrl = _serverUrl

            ' 构建JSON-RPC请求对象
            Dim jsonRpcRequest = New JObject()
            jsonRpcRequest("jsonrpc") = "2.0"
            jsonRpcRequest("id") = 1
            jsonRpcRequest("method") = "initialize"
            jsonRpcRequest("params") = JObject.FromObject(initParams)

            Dim content = New StringContent(jsonRpcRequest.ToString(), Encoding.UTF8, "application/json")

            Debug.WriteLine($"初始化请求: {requestUrl}")
            Debug.WriteLine($"请求内容: {jsonRpcRequest}")

            ' 发送POST请求初始化
            Dim response = Await _httpClient.PostAsync(requestUrl, content)

            If response.IsSuccessStatusCode Then
                Dim responseContent = Await response.Content.ReadAsStringAsync()
                Debug.WriteLine($"初始化响应: {responseContent}")

                ' 解析JSON-RPC响应
                Dim jsonResponse = JObject.Parse(responseContent)

                ' 检查是否有错误
                If jsonResponse("error") IsNot Nothing Then
                    Dim errorMsg = jsonResponse("error")("message")?.ToString()
                    Throw New Exception($"Server error: {errorMsg}")
                End If

                ' 处理结果
                Dim result = jsonResponse("result")

                If result IsNot Nothing Then
                    _isInitialized = True
                    _serverCapabilities = New MCPServerCapabilities() With {
                        .Tools = result("capabilities")("tools") IsNot Nothing,
                        .Resources = result("capabilities")("resources") IsNot Nothing,
                        .Prompts = result("capabilities")("prompts") IsNot Nothing,
                        .Sampling = result("capabilities")("sampling") IsNot Nothing,
                        .Roots = result("capabilities")("roots") IsNot Nothing
                    }

                    Return New MCPInitResponse() With {
                        .Success = True,
                        .ProtocolVersion = result("protocolVersion")?.ToString(),
                        .ServerInfo = New MCPServerInfo() With {
                            .Name = result("serverInfo")("name")?.ToString(),
                            .Version = result("serverInfo")("version")?.ToString()
                        },
                        .Capabilities = _serverCapabilities
                    }
                Else
                    Throw New Exception("Invalid server response: missing result object")
                End If
            Else
                ' 如果服务器不支持POST方法，尝试使用GET进行简单的连接测试
                Dim getResponse = Await _httpClient.GetAsync(requestUrl)
                getResponse.EnsureSuccessStatusCode()

                ' 简单的连接测试成功，假定服务器支持基本功能
                _isInitialized = True
                _serverCapabilities = New MCPServerCapabilities() With {
                    .Tools = True,
                    .Resources = False,
                    .Prompts = False,
                    .Sampling = False,
                    .Roots = False
                }

                Return New MCPInitResponse() With {
                    .Success = True,
                    .ProtocolVersion = "2024-11-05",
                    .ServerInfo = New MCPServerInfo() With {
                        .Name = "MCP Server",
                        .Version = "1.0.0"
                    },
                    .Capabilities = _serverCapabilities
                }
            End If
        Catch ex As Exception
            Debug.WriteLine($"SSE初始化错误: {ex.Message}")
            Return New MCPInitResponse() With {
                .Success = False,
                .ErrorMessage = $"SSE initialization failed: {ex.Message}"
            }
        End Try
    End Function

    ' 处理初始化响应
    Private Function ProcessInitResponse(result As JObject) As MCPInitResponse
        Try
            _serverCapabilities = New MCPServerCapabilities() With {
                .Tools = result("capabilities")("tools") IsNot Nothing,
                .Resources = result("capabilities")("resources") IsNot Nothing,
                .Prompts = result("capabilities")("prompts") IsNot Nothing,
                .Sampling = result("capabilities")("sampling") IsNot Nothing,
                .Roots = result("capabilities")("roots") IsNot Nothing
            }

            _isInitialized = True

            Return New MCPInitResponse() With {
                .Success = True,
                .ProtocolVersion = result("protocolVersion")?.ToString(),
                .ServerInfo = New MCPServerInfo() With {
                    .Name = result("serverInfo")("name")?.ToString(),
                    .Version = result("serverInfo")("version")?.ToString()
                },
                .Capabilities = _serverCapabilities
            }

        Catch ex As Exception
            Throw New Exception($"Failed to process init response: {ex.Message}", ex)
        End Try
    End Function


    ' 标准JSON-RPC请求
    Private Async Function InvokeSSEMethodAsync(Of T)(method As String, arguments() As Object) As Task(Of T)
        Try
            ' 构建JSON-RPC请求
            Dim jsonRpcRequest = New JObject()
            jsonRpcRequest("jsonrpc") = "2.0"
            jsonRpcRequest("id") = Guid.NewGuid().ToString()
            jsonRpcRequest("method") = method

            ' 添加参数
            If arguments.Length > 0 AndAlso arguments(0) IsNot Nothing Then
                jsonRpcRequest("params") = JToken.FromObject(arguments(0))
            End If

            Dim json = jsonRpcRequest.ToString()
            Dim content = New StringContent(json, Encoding.UTF8, "application/json")

            Debug.WriteLine($"JSON-RPC请求: {_serverUrl}")
            Debug.WriteLine($"请求内容: {json}")

            ' 发送请求
            Dim response = Await _httpClient.PostAsync(_serverUrl, content)

            If Not response.IsSuccessStatusCode Then
                Debug.WriteLine($"HTTP 错误: {response.StatusCode} - {response.ReasonPhrase}")
                Throw New HttpRequestException($"服务器返回错误: {response.StatusCode} {response.ReasonPhrase}")
            End If

            Dim responseContent = Await response.Content.ReadAsStringAsync()
            Debug.WriteLine($"响应内容: {responseContent}")

            ' 解析 JSON-RPC 响应
            Try
                Dim jsonResponse = JObject.Parse(responseContent)

                ' 检查是否有错误
                If jsonResponse("error") IsNot Nothing Then
                    Dim errorCode = jsonResponse("error")("code")?.ToObject(Of Integer)()
                    Dim errorMsg = jsonResponse("error")("message")?.ToString()

                    Debug.WriteLine($"JSON-RPC 错误: 错误码: {errorCode}, 消息: {errorMsg}")
                    Throw New Exception($"服务器错误: {errorMsg}")
                End If

                ' 获取结果
                Dim result = jsonResponse("result")

                If result IsNot Nothing Then
                    ' 将结果转换为请求的类型
                    Return result.ToObject(Of T)()
                Else
                    Throw New Exception("无效的服务器响应: 缺少 result 对象")
                End If
            Catch jsonEx As JsonException
                Debug.WriteLine($"JSON 解析错误: {jsonEx.Message}")

                ' 尝试其他方式解析响应
                If GetType(T) Is GetType(JObject) Then
                    ' 构造一个基本响应对象
                    Dim fallbackResult = New JObject()
                    Dim contentArray = New JArray()
                    contentArray.Add(New JObject From {
                        {"type", "text"},
                        {"text", responseContent}
                    })
                    fallbackResult("content") = contentArray
                    Return CType(CObj(fallbackResult), T)
                End If

                Throw
            End Try
        Catch ex As Exception
            Debug.WriteLine($"JSON-RPC 请求失败: {ex.Message}")
            Throw
        End Try
    End Function

    ' 列出工具 - 标准MCP方法
    Public Async Function ListToolsAsync() As Task(Of List(Of MCPToolInfo))
        Try
            Dim result = Await InvokeMethodAsync(Of JObject)("tools/list", New Object())

            Dim tools = New List(Of MCPToolInfo)()
            Dim toolsArray = result("tools")

            If toolsArray IsNot Nothing Then
                For Each tool In toolsArray
                    tools.Add(New MCPToolInfo() With {
                        .Name = tool("name")?.ToString(),
                        .Description = tool("description")?.ToString(),
                        .InputSchema = tool("inputSchema")
                    })
                Next
            End If

            Debug.WriteLine($"成功获取到 {tools.Count} 个工具")
            Return tools

        Catch ex As Exception
            Debug.WriteLine($"获取工具列表失败: {ex.Message}")
            Throw New Exception($"Failed to list tools: {ex.Message}", ex)
        End Try
    End Function

    ' 调用工具 - 标准MCP方法
    Public Async Function CallToolAsync(toolName As String, arguments As Object) As Task(Of MCPToolResult)
        Try
            ' 构建标准参数
            Dim params = New With {
            .name = toolName,
            .arguments = arguments
        }

            ' 记录请求详情
            Debug.WriteLine($"调用工具 {toolName} 开始，参数: {JsonConvert.SerializeObject(arguments)}")

            Dim result = Await InvokeMethodAsync(Of JObject)("tools/call", params)

            ' 记录原始响应
            Debug.WriteLine($"工具 {toolName} 调用响应: {result.ToString()}")

            Dim content = New List(Of MCPContent)()
            If result("content") IsNot Nothing Then
                For Each item In result("content")
                    Dim contentItem = New MCPContent() With {
                    .Type = item("type")?.ToString(),
                    .Text = item("text")?.ToString(),
                    .Data = item("data")?.ToString(),
                    .MimeType = item("mimeType")?.ToString()
                }
                    content.Add(contentItem)
                    Debug.WriteLine($"添加内容项: 类型={contentItem.Type}, 文本长度={If(contentItem.Text Is Nothing, 0, contentItem.Text.Length)}")
                Next
            Else
                Debug.WriteLine("响应中未找到content字段")
            End If

            Dim isError = False
            Dim errorMessage As String = Nothing

            If result("isError") IsNot Nothing Then
                isError = result("isError").ToObject(Of Boolean)()
                Debug.WriteLine($"响应中isError字段: {isError}")
            End If

            If result("error") IsNot Nothing Then
                errorMessage = result("error").ToString()
                Debug.WriteLine($"响应中error字段: {errorMessage}")
            End If

            If result("errorMessage") IsNot Nothing Then
                errorMessage = result("errorMessage").ToString()
                Debug.WriteLine($"响应中errorMessage字段: {errorMessage}")
            End If

            Return New MCPToolResult() With {
            .IsError = isError,
            .ErrorMessage = errorMessage,
            .Content = content
        }

        Catch ex As Exception
            Debug.WriteLine($"工具调用异常: {ex.ToString()}")
            Return New MCPToolResult() With {
            .IsError = True,
            .ErrorMessage = $"调用工具时发生错误: {ex.Message}",
            .Content = New List(Of MCPContent)()
        }
        End Try
    End Function

    ' 列出资源 - 标准MCP方法
    Public Async Function ListResourcesAsync() As Task(Of List(Of MCPResourceInfo))
        Try
            Dim result = Await InvokeMethodAsync(Of JObject)("resources/list", New Object())

            Dim resources = New List(Of MCPResourceInfo)()
            Dim resourcesArray = result("resources")

            If resourcesArray IsNot Nothing Then
                For Each resource In resourcesArray
                    resources.Add(New MCPResourceInfo() With {
                        .Uri = resource("uri")?.ToString(),
                        .Name = resource("name")?.ToString(),
                        .Description = resource("description")?.ToString(),
                        .MimeType = resource("mimeType")?.ToString()
                    })
                Next
            End If

            Return resources

        Catch ex As Exception
            Debug.WriteLine($"获取资源列表失败: {ex.Message}")
            Throw New Exception($"Failed to list resources: {ex.Message}", ex)
        End Try
    End Function

    ' 读取资源 - 标准MCP方法
    Public Async Function ReadResourceAsync(uri As String) As Task(Of MCPResourceResult)
        Try
            Dim params = New With {.uri = uri}
            Dim result = Await InvokeMethodAsync(Of JObject)("resources/read", params)

            Dim contents = New List(Of MCPContent)()
            If result("contents") IsNot Nothing Then
                For Each item In result("contents")
                    contents.Add(New MCPContent() With {
                        .Uri = item("uri")?.ToString(),
                        .MimeType = item("mimeType")?.ToString(),
                        .Text = item("text")?.ToString(),
                        .Blob = item("blob")?.ToString()
                    })
                Next
            End If

            Return New MCPResourceResult() With {
                .Contents = contents
            }

        Catch ex As Exception
            Debug.WriteLine($"读取资源失败: {ex.Message}")
            Throw New Exception($"Failed to read resource: {ex.Message}", ex)
        End Try
    End Function

    ' 列出提示 - 标准MCP方法
    Public Async Function ListPromptsAsync() As Task(Of List(Of MCPPromptInfo))
        Try
            Dim result = Await InvokeMethodAsync(Of JObject)("prompts/list", New Object())

            Dim prompts = New List(Of MCPPromptInfo)()
            Dim promptsArray = result("prompts")

            If promptsArray IsNot Nothing Then
                For Each prompt In promptsArray
                    Dim args = New List(Of MCPPromptArgument)()
                    If prompt("arguments") IsNot Nothing Then
                        For Each arg In prompt("arguments")
                            args.Add(New MCPPromptArgument() With {
                                .Name = arg("name")?.ToString(),
                                .Description = arg("description")?.ToString(),
                                .Required = If(arg("required") Is Nothing, False, arg("required").ToObject(Of Boolean)())
                            })
                        Next
                    End If

                    prompts.Add(New MCPPromptInfo() With {
                        .Name = prompt("name")?.ToString(),
                        .Description = prompt("description")?.ToString(),
                        .Arguments = args
                    })
                Next
            End If

            Return prompts

        Catch ex As Exception
            Debug.WriteLine($"获取提示列表失败: {ex.Message}")
            Throw New Exception($"Failed to list prompts: {ex.Message}", ex)
        End Try
    End Function

    ' 检查是否已初始化
    Private Sub EnsureInitialized()
        If Not _isInitialized Then
            Throw New InvalidOperationException("MCP client not initialized. Call InitializeAsync first.")
        End If
    End Sub

    Public ReadOnly Property IsInitialized As Boolean
        Get
            Return _isInitialized
        End Get
    End Property

    Public ReadOnly Property ServerCapabilities As MCPServerCapabilities
        Get
            Return _serverCapabilities
        End Get
    End Property

    Public ReadOnly Property TransportType As MCPTransportType
        Get
            Return _transportType
        End Get
    End Property



End Class