Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json

Public Class LLMUtil

    ' 创建请求体
    Public Shared Function CreateRequestBody(question As String) As String
        Dim result As String = question.Replace("\", "\\").Replace("""", "\""").
                                  Replace(vbCr, "\r").Replace(vbLf, "\n").
                                  Replace(vbTab, "\t").Replace(vbBack, "\b").
                                  Replace(Chr(12), "\f")
        ' 使用从 ConfigSettings 中获取的模型名称
        Return "{""model"": """ & ConfigSettings.ModelName & """, ""messages"": [{""role"": ""user"", ""content"": """ & result & """}]}"
    End Function



    ' 创建LLM API请求体
    Public Shared Function CreateLlmRequestBody(
        prompt As String,
        modelT As String,
        systemPrompt As String,
        temperatureT As Double,
        maxTokens As Integer) As String

        Try
            ' 构建消息数组
            Dim messagesT As New List(Of Object)()

            ' 添加系统消息（如果有）
            If Not String.IsNullOrEmpty(systemPrompt) Then
                messagesT.Add(New With {
                    .role = "system",
                    .content = systemPrompt
                })
            End If

            ' 添加用户消息
            messagesT.Add(New With {
                .role = "user",
                .content = prompt
            })

            ' 构建完整请求对象
            Dim requestObj = New With {
                .model = modelT,
                .messages = messagesT,
                .temperature = temperatureT,
                .max_tokens = maxTokens,
                .stream = False  ' 关闭流式响应
            }

            ' 序列化为JSON
            Return JsonConvert.SerializeObject(requestObj)

        Catch ex As Exception
            Throw New Exception($"创建请求体时出错: {ex.Message}")
        End Try
    End Function

    Public Shared Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            SimpleLogger.LogInfo($"开始发送HTTP请求到: {apiUrl}")
            SimpleLogger.LogInfo($"请求头Authorization: Bearer {apiKey.Substring(0, Math.Min(10, apiKey.Length))}...")
            SimpleLogger.LogInfo($"请求体长度: {requestBody.Length}")

            ' 强制使用 TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim handler As New HttpClientHandler()

            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromSeconds(120) ' 设置超时时间为 120 秒
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)

                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                SimpleLogger.LogInfo("正在发送POST请求...")

                Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)

                SimpleLogger.LogInfo($"HTTP响应状态码: {response.StatusCode}")
                SimpleLogger.LogInfo($"HTTP响应原因: {response.ReasonPhrase}")

                ' 检查响应状态
                If Not response.IsSuccessStatusCode Then
                    Dim errorContent As String = Await response.Content.ReadAsStringAsync()
                    SimpleLogger.LogInfo($"HTTP错误响应内容: {errorContent}")
                    Throw New HttpRequestException($"HTTP请求失败: {response.StatusCode} - {response.ReasonPhrase}. 详细信息: {errorContent}")
                End If

                Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                SimpleLogger.LogInfo($"HTTP响应内容长度: {responseContent.Length}")
                SimpleLogger.LogInfo($"HTTP响应内容前200字符: {responseContent.Substring(0, Math.Min(200, responseContent.Length))}")

                Return responseContent
            End Using

        Catch ex As TaskCanceledException
            SimpleLogger.LogInfo($"HTTP请求超时: {ex.Message}")
            Return $"错误: 请求超时 - {ex.Message}"
        Catch ex As HttpRequestException
            SimpleLogger.LogInfo($"HTTP请求异常: {ex.Message}")
            ' 不显示MessageBox，直接返回错误信息
            Return $"错误: HTTP请求失败 - {ex.Message}"
        Catch ex As Exception
            SimpleLogger.LogInfo($"发送HTTP请求时发生未知异常: {ex.Message}")
            SimpleLogger.LogInfo($"异常类型: {ex.GetType().Name}")
            SimpleLogger.LogInfo($"异常堆栈: {ex.StackTrace}")
            Return $"错误: {ex.Message}"
        End Try
    End Function
    ' 添加同步版本的HTTP请求方法
    Public Shared Function SendHttpRequestSync(apiUrl As String, apiKey As String, requestBody As String) As String
        Try

            ' 强制使用 TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)

                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")

                ' 使用 .Result 进行同步调用
                Dim response As HttpResponseMessage = client.PostAsync(apiUrl, content).Result

                SimpleLogger.LogInfo($"HTTP响应状态码: {response.StatusCode}")

                If Not response.IsSuccessStatusCode Then
                    Dim errorContent As String = response.Content.ReadAsStringAsync().Result
                    SimpleLogger.LogInfo($"HTTP错误响应内容: {errorContent}")
                    Return $"错误: HTTP请求失败 - {response.StatusCode} {response.ReasonPhrase}"
                End If

                Dim responseContent As String = response.Content.ReadAsStringAsync().Result
                Return responseContent
            End Using

        Catch ex As AggregateException
            ' 处理 .Result 可能产生的 AggregateException
            Dim innerEx = ex.GetBaseException()
            Return $"错误: {innerEx.Message}"
        Catch ex As Exception
            SimpleLogger.LogInfo($"异常类型: {ex.GetType().Name}")
            Return $"错误: {ex.Message}"
        End Try
    End Function
End Class
