' ShareRibbon\Controls\Services\ModelConnectivityTester.vb
' 连接测试只有一个实现：配置窗体的"验证并保存"和首次引导的"测试连接"必须给同一个结论，
' 否则引导里显示可用、窗体里保存失败，用户只能反复重填 Key。

Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ModelConnectivityResult
    Public Property Success As Boolean
    Public Property Reason As String
    Public Property Detail As String

    Public Function ToJson() As String
        Dim payload As New JObject()
        payload("success") = Success
        payload("reason") = If(Reason, "")
        payload("detail") = If(Detail, "")
        Return payload.ToString(Formatting.None)
    End Function
End Class

Public Class ModelConnectivityTester
    Private Const TimeoutSeconds As Integer = 30

    Public Async Function TestAsync(apiUrl As String, apiKey As String, modelName As String,
                                    Optional reasoningMode As String = Nothing,
                                    Optional platformName As String = Nothing) As Task(Of ModelConnectivityResult)
        If String.IsNullOrWhiteSpace(apiUrl) OrElse String.IsNullOrWhiteSpace(modelName) Then
            Return Build(False, "invalid_config", "接口地址或模型名称为空，无法发起测试。")
        End If

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim client = HttpClientPool.GetClient(apiUrl)
            Using request As New HttpRequestMessage(HttpMethod.Post, apiUrl)
                ' 本地服务（LM Studio 等）不校验 Key，这里不能替用户把空 Key 直接判死
                If Not String.IsNullOrWhiteSpace(apiKey) Then
                    request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                End If
                Dim requestBody = New JObject() From {
                    {"model", modelName},
                    {"messages", New JArray() From {
                        New JObject() From {
                            {"role", "user"},
                            {"content", "Hi"}
                        }
                    }},
                    {"max_tokens", 5}
                }
                ReasoningRequestHelper.ApplyReasoningOptions(requestBody, reasoningMode, modelName, platformName, apiUrl)
                request.Content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")

                Using timeoutCts As New CancellationTokenSource(TimeSpan.FromSeconds(TimeoutSeconds))
                    Dim response As HttpResponseMessage = Nothing
                    Try
                        response = Await client.SendAsync(request, timeoutCts.Token)
                    Catch cancelled As TaskCanceledException
                        Return Build(False, "timeout", $"{TimeoutSeconds} 秒内没有响应，请确认服务已启动或网络可达。")
                    End Try

                    Using response
                        If response.IsSuccessStatusCode Then
                            Return Build(True, "ok", "连接正常。")
                        End If
                        Return Build(False, MapStatus(CInt(response.StatusCode)), DescribeStatus(CInt(response.StatusCode), apiUrl, modelName))
                    End Using
                End Using
            End Using
        Catch httpEx As HttpRequestException
            Return Build(False, "unreachable", $"无法连接 {apiUrl}：{httpEx.Message}")
        Catch keyEx As ArgumentException
            ' Bearer 头拒绝非法字符，多半是粘贴带了空格/换行的 Key
            Return Build(False, "invalid_key_format", "API Key 含非法字符，请重新粘贴完整 Key。")
        Catch ex As Exception
            Return Build(False, "error", ex.Message)
        End Try
    End Function

    Private Shared Function MapStatus(status As Integer) As String
        Select Case status
            Case 401, 403
                Return "rejected_key"
            Case 404
                Return "not_found"
            Case 429
                Return "rate_limited"
            Case Else
                If status >= 500 Then Return "server_error"
                Return "http_error"
        End Select
    End Function

    Private Shared Function DescribeStatus(status As Integer, apiUrl As String, modelName As String) As String
        Select Case status
            Case 401, 403
                Return "接口拒绝了当前 API Key（HTTP " & status & "），请确认 Key 是否有效或已过期。"
            Case 404
                Return $"接口返回 404（HTTP {status}），请检查地址或模型名称 {modelName} 是否存在。"
            Case 429
                Return "请求被限流（HTTP 429），稍后再试或更换额度。"
            Case Else
                If status >= 500 Then
                    Return $"服务端异常（HTTP {status}），{apiUrl} 暂时不可用。"
                End If
                Return $"请求失败（HTTP {status}）。"
        End Select
    End Function

    Private Shared Function Build(success As Boolean, reason As String, detail As String) As ModelConnectivityResult
        Return New ModelConnectivityResult With {
            .Success = success,
            .Reason = reason,
            .Detail = detail
        }
    End Function
End Class
