Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class LLMUtil

    ''' <summary>
    ''' 检查 WPS Office 是否正在运行。
    ''' WPS 各组件进程名：wps(Writer)、et(表格)、wpp(演示)、wpspdf(PDF)。
    ''' 注：GetProcessesByName 在 Windows 上大小写不敏感，无需重复传入大写变体。
    ''' </summary>
    Public Shared Function IsWpsActive() As Boolean
        Try
            ' Process.GetProcessesByName 返回持有 OS 句柄的 Process 对象，必须 Dispose 避免句柄泄漏
            For Each name In {"wps", "et", "wpp", "wpspdf"}
                Dim procs = Process.GetProcessesByName(name)
                Dim found = procs.Length > 0
                For Each p In procs
                    p.Dispose()
                Next
                If found Then Return True
            Next
            Return False
        Catch
            Return False
        End Try
    End Function

    ' 创建请求体
    Public Shared Function CreateRequestBody(question As String) As String
        Dim requestObj As New JObject()
        requestObj("model") = ConfigSettings.ModelName
        requestObj("messages") = New JArray() From {
            New JObject() From {
                {"role", "user"},
                {"content", If(question, String.Empty)}
            }
        }
        ReasoningRequestHelper.ApplyReasoningOptions(requestObj, ConfigSettings.ReasoningMode, ConfigSettings.ModelName, ConfigSettings.platform, ConfigSettings.ApiUrl)
        Return requestObj.ToString(Formatting.None)
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
            ' E3-05：ExcelDna 自定义函数与批量数据生成从这里出去，载荷是工作表内容。
            ' 拒绝沿用本函数既有的「返回 错误: …」约定，而不是另造第三种终态。
            If PrivacyConsentGate.RefuseQuietly("LLMUtil") Then Return "错误: " & PrivacyConsentGate.RefusalReason
            Debug.WriteLine($"开始发送HTTP请求到: {apiUrl}")
            Debug.WriteLine($"请求头Authorization: Bearer {apiKey.Substring(0, Math.Min(10, apiKey.Length))}...")
            Debug.WriteLine($"请求体长度: {requestBody.Length}")

            ' 强制使用 TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim handler As New HttpClientHandler()

            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromSeconds(120) ' 设置超时时间为 120 秒
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)

                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Debug.WriteLine("正在发送POST请求...")

                Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)

                Debug.WriteLine($"HTTP响应状态码: {response.StatusCode}")
                Debug.WriteLine($"HTTP响应原因: {response.ReasonPhrase}")

                ' 检查响应状态
                If Not response.IsSuccessStatusCode Then
                    Dim errorContent As String = Await response.Content.ReadAsStringAsync()
                    Debug.WriteLine($"HTTP错误响应内容: {errorContent}")
                    Throw New HttpRequestException($"HTTP请求失败: {response.StatusCode} - {response.ReasonPhrase}. 详细信息: {errorContent}")
                End If

                Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                Debug.WriteLine($"HTTP响应内容长度: {responseContent.Length}")
                Debug.WriteLine($"HTTP响应内容前200字符: {responseContent.Substring(0, Math.Min(200, responseContent.Length))}")

                Return responseContent
            End Using

        Catch ex As TaskCanceledException
            Debug.WriteLine($"HTTP请求超时: {ex.Message}")
            Return $"错误: 请求超时 - {ex.Message}"
        Catch ex As HttpRequestException
            Debug.WriteLine($"HTTP请求异常: {ex.Message}")
            ' 不显示MessageBox，直接返回错误信息
            Return $"错误: HTTP请求失败 - {ex.Message}"
        Catch ex As Exception
            Debug.WriteLine($"发送HTTP请求时发生未知异常: {ex.Message}")
            Debug.WriteLine($"异常类型: {ex.GetType().Name}")
            Debug.WriteLine($"异常堆栈: {ex.StackTrace}")
            Return $"错误: {ex.Message}"
        End Try
    End Function
    ''' <summary>
    ''' Sync HTTP for true sync boundaries (e.g. ExcelDna UDF).
    ''' Uses SyncOverAsync so the request does not capture WinForms SynchronizationContext.
    ''' Prefer SendHttpRequest (async) on chat/agent paths.
    ''' </summary>
    Public Shared Function SendHttpRequestSync(apiUrl As String, apiKey As String, requestBody As String) As String
        Try
            Dim result = SyncOverAsync.Run(
                Function() SendHttpRequest(apiUrl, apiKey, requestBody),
                120000)
            Return If(result, "错误: 请求超时或无响应")
        Catch ex As Exception
            Debug.WriteLine($"[LLMUtil.SendHttpRequestSync] {ex.GetType().Name}: {ex.Message}")
            Return $"错误: {ex.Message}"
        End Try
    End Function
End Class
