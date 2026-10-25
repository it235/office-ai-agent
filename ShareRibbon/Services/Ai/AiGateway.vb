' ShareRibbon\Services\Ai\AiGateway.vb
' Shared non-streaming chat gateway for internal LLM calls.

Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class AiRequestOptions
    Public Property ApiUrl As String
    Public Property ApiKey As String
    Public Property ModelName As String
    Public Property Platform As String
    Public Property ReasoningMode As String
    Public Property SystemPrompt As String
    Public Property UserPrompt As String
    Public Property Messages As JArray
    Public Property Temperature As Double?
    Public Property MaxTokens As Integer?
    Public Property TimeoutSeconds As Integer

    ''' <summary>
    ''' 调用方自己的取消来源（例如用户点了"停止"）。默认 None：
    ''' 记忆沉淀、embedding 等后台链路不传，就不会被回合取消连带掐断。
    ''' </summary>
    Public Property Cancellation As CancellationToken

    Public Sub New()
        ReasoningMode = ReasoningRequestHelper.ReasoningDefault
        TimeoutSeconds = 60
    End Sub
End Class

Public Class AiGatewayResponse
    Public Property Success As Boolean
    Public Property Content As String
    Public Property RawResponse As String
    Public Property ErrorMessage As String
    Public Property StatusCode As Integer
    ''' <summary>provider 报了 usage 才有；没报就是 Nothing，调用方不得当成 0 用量。</summary>
    Public Property Usage As TokenUsage
End Class

Public Class AiGateway
    Private Shared ReadOnly _httpClient As New HttpClient()
    Private Const AnthropicVersion As String = "2023-06-01"

    Public Shared Async Function SendChatAsync(options As AiRequestOptions) As Task(Of AiGatewayResponse)
        If options Is Nothing Then
            Return Fail("AiRequestOptions is required.")
        End If

        If String.IsNullOrWhiteSpace(options.ApiUrl) Then Return Fail("ApiUrl is required.")
        If String.IsNullOrWhiteSpace(options.ApiKey) Then Return Fail("ApiKey is required.")
        If String.IsNullOrWhiteSpace(options.ModelName) Then Return Fail("ModelName is required.")

        ' E3-05：共享请求层的出口。意图识别、记忆沉淀、AgentKernel 的内部调用都从这里出去，
        ' 在这里拦一次，比在每个调用方各拦一次可靠。理由文案归 PrivacyConsentGate 持有。
        If Not PrivacyConsentGate.IsAccepted() Then
            Debug.WriteLine("[AiGateway] refused: privacy consent not accepted")
            Return Fail(PrivacyConsentGate.RefusalReason)
        End If

        Dim isAnthropic = IsAnthropicEndpoint(options.ApiUrl)
        Dim requestBody = BuildProviderRequest(options)
        Dim timeoutSeconds = If(options.TimeoutSeconds > 0, options.TimeoutSeconds, 60)

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using request As New HttpRequestMessage(HttpMethod.Post, options.ApiUrl)
                ApplyHeaders(request, options.ApiKey, isAnthropic)
                request.Content = New StringContent(requestBody.ToString(Formatting.None), Encoding.UTF8, "application/json")

                Using cts As CancellationTokenSource = CancellationTokenSource.CreateLinkedTokenSource(options.Cancellation)
                    cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds))
                    Using response = Await _httpClient.SendAsync(request, cts.Token)
                        Dim responseText = Await response.Content.ReadAsStringAsync()
                        If Not response.IsSuccessStatusCode Then
                            Dim errorMessage = $"HTTP {CInt(response.StatusCode)} {response.ReasonPhrase}: {Truncate(responseText, 1000)}"
                            Debug.WriteLine($"[AiGateway] Request failed: {errorMessage}")
                            Return New AiGatewayResponse With {
                                .Success = False,
                                .RawResponse = responseText,
                                .ErrorMessage = errorMessage,
                                .StatusCode = CInt(response.StatusCode)
                            }
                        End If

                        Return New AiGatewayResponse With {
                            .Success = True,
                            .Content = ExtractAssistantContent(responseText, isAnthropic),
                            .RawResponse = responseText,
                            .Usage = TokenUsage.FromResponse(responseText),
                            .StatusCode = CInt(response.StatusCode)
                        }
                    End Using
                End Using
            End Using
        Catch ex As OperationCanceledException When options.Cancellation.IsCancellationRequested
            Throw
        Catch ex As TaskCanceledException
            Dim errorMessage = $"Request timed out after {timeoutSeconds} seconds."
            Debug.WriteLine($"[AiGateway] {errorMessage} {ex.Message}")
            Return Fail(errorMessage)
        Catch ex As Exception
            Debug.WriteLine($"[AiGateway] Request exception: {ex.Message}")
            Return Fail(ex.Message)
        End Try
    End Function

    Public Shared Function BuildOpenAiCompatibleRequest(options As AiRequestOptions) As JObject
        Dim requestObj As New JObject()
        requestObj("model") = options.ModelName
        requestObj("messages") = BuildMessages(options)
        requestObj("stream") = False

        If options.Temperature.HasValue Then requestObj("temperature") = options.Temperature.Value
        If options.MaxTokens.HasValue AndAlso options.MaxTokens.Value > 0 Then requestObj("max_tokens") = options.MaxTokens.Value

        ReasoningRequestHelper.ApplyReasoningOptions(
            requestObj,
            options.ReasoningMode,
            options.ModelName,
            options.Platform,
            options.ApiUrl)

        Return requestObj
    End Function

    Public Shared Function BuildProviderRequest(options As AiRequestOptions) As JObject
        If options Is Nothing Then Throw New ArgumentNullException(NameOf(options))

        Dim requestObj = BuildOpenAiCompatibleRequest(options)
        If IsAnthropicEndpoint(options.ApiUrl) Then
            Return ConvertToAnthropicRequest(requestObj)
        End If

        Return requestObj
    End Function

    Public Shared Function ExtractAssistantContent(responseText As String, Optional isAnthropic As Boolean = False) As String
        If String.IsNullOrWhiteSpace(responseText) Then Return ""

        Try
            Dim obj = JObject.Parse(responseText)

            If isAnthropic Then
                Dim anthropicText = ExtractAnthropicContent(obj)
                If Not String.IsNullOrWhiteSpace(anthropicText) Then Return anthropicText
            End If

            Dim content = TokenToString(obj.SelectToken("choices[0].message.content"))
            If Not String.IsNullOrWhiteSpace(content) Then Return content

            content = TokenToString(obj.SelectToken("choices[0].text"))
            If Not String.IsNullOrWhiteSpace(content) Then Return content

            content = TokenToString(obj("output_text"))
            If Not String.IsNullOrWhiteSpace(content) Then Return content

            content = ExtractResponsesOutputText(obj)
            If Not String.IsNullOrWhiteSpace(content) Then Return content

            content = ExtractAnthropicContent(obj)
            If Not String.IsNullOrWhiteSpace(content) Then Return content
        Catch ex As Exception
            Debug.WriteLine($"[AiGateway] Response parse failed: {ex.Message}")
        End Try

        Return responseText
    End Function

    Private Shared Function BuildMessages(options As AiRequestOptions) As JArray
        If options.Messages IsNot Nothing Then
            Return New JArray(options.Messages)
        End If

        Dim messages As New JArray()
        If Not String.IsNullOrWhiteSpace(options.SystemPrompt) Then
            messages.Add(New JObject From {
                {"role", "system"},
                {"content", options.SystemPrompt}
            })
        End If

        messages.Add(New JObject From {
            {"role", "user"},
            {"content", If(options.UserPrompt, "")}
        })
        Return messages
    End Function

    Private Shared Sub ApplyHeaders(request As HttpRequestMessage, apiKey As String, isAnthropic As Boolean)
        If isAnthropic Then
            request.Headers.Add("x-api-key", apiKey)
            request.Headers.Add("anthropic-version", AnthropicVersion)
        Else
            request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
        End If
    End Sub

    Private Shared Function IsAnthropicEndpoint(apiUrl As String) As Boolean
        Dim normalized = If(apiUrl, "").ToLowerInvariant()
        Return normalized.Contains("anthropic.com") OrElse normalized.Contains("/v1/messages")
    End Function

    Private Shared Function ConvertToAnthropicRequest(openAiRequest As JObject) As JObject
        Dim anthropicBody As New JObject()
        anthropicBody("model") = openAiRequest("model")
        anthropicBody("max_tokens") = If(openAiRequest("max_tokens"), New JValue(4096))
        anthropicBody("stream") = False

        If openAiRequest("temperature") IsNot Nothing Then anthropicBody("temperature") = openAiRequest("temperature")

        Dim messages = TryCast(openAiRequest("messages"), JArray)
        Dim convertedMessages As New JArray()
        Dim systemParts As New List(Of String)()

        If messages IsNot Nothing Then
            For Each msgToken In messages
                Dim msg = TryCast(msgToken, JObject)
                If msg Is Nothing Then Continue For

                Dim role = If(msg("role"), "").ToString()
                Dim content = msg("content")

                If String.Equals(role, "system", StringComparison.OrdinalIgnoreCase) Then
                    Dim systemText = TokenToString(content)
                    If Not String.IsNullOrWhiteSpace(systemText) Then systemParts.Add(systemText)
                    Continue For
                End If

                If String.Equals(role, "tool", StringComparison.OrdinalIgnoreCase) Then
                    convertedMessages.Add(New JObject From {
                        {"role", "user"},
                        {"content", New JArray From {
                            New JObject From {
                                {"type", "tool_result"},
                                {"tool_use_id", If(msg("tool_call_id"), "").ToString()},
                                {"content", TokenToString(content)}
                            }
                        }}
                    })
                    Continue For
                End If

                If String.Equals(role, "assistant", StringComparison.OrdinalIgnoreCase) AndAlso msg("tool_calls") IsNot Nothing Then
                    convertedMessages.Add(ConvertAssistantToolCallMessage(msg))
                    Continue For
                End If

                Dim normalizedRole = If(String.Equals(role, "assistant", StringComparison.OrdinalIgnoreCase), "assistant", "user")
                convertedMessages.Add(New JObject From {
                    {"role", normalizedRole},
                    {"content", TokenToAnthropicContent(content)}
                })
            Next
        End If

        If systemParts.Count > 0 Then anthropicBody("system") = String.Join(vbCrLf, systemParts)
        anthropicBody("messages") = convertedMessages

        If openAiRequest("tools") IsNot Nothing Then
            anthropicBody("tools") = ConvertToolsToAnthropic(TryCast(openAiRequest("tools"), JArray))
        End If

        Return anthropicBody
    End Function

    Private Shared Function ConvertAssistantToolCallMessage(msg As JObject) As JObject
        Dim contentArr As New JArray()
        Dim textContent = TokenToString(msg("content"))
        If Not String.IsNullOrWhiteSpace(textContent) Then
            contentArr.Add(New JObject From {
                {"type", "text"},
                {"text", textContent}
            })
        End If

        Dim toolCalls = TryCast(msg("tool_calls"), JArray)
        If toolCalls IsNot Nothing Then
            For Each tcToken In toolCalls
                Dim tc = TryCast(tcToken, JObject)
                If tc Is Nothing Then Continue For

                Dim functionObj = TryCast(tc("function"), JObject)
                Dim argsToken As JToken = New JObject()
                If functionObj IsNot Nothing AndAlso functionObj("arguments") IsNot Nothing Then
                    argsToken = ParseJsonOrString(functionObj("arguments").ToString())
                End If

                contentArr.Add(New JObject From {
                    {"type", "tool_use"},
                    {"id", If(tc("id"), "").ToString()},
                    {"name", If(functionObj Is Nothing, "", If(functionObj("name"), "").ToString())},
                    {"input", argsToken}
                })
            Next
        End If

        Return New JObject From {
            {"role", "assistant"},
            {"content", contentArr}
        }
    End Function

    Private Shared Function ConvertToolsToAnthropic(tools As JArray) As JArray
        Dim anthropicTools As New JArray()
        If tools Is Nothing Then Return anthropicTools

        For Each toolToken In tools
            Dim toolObj = TryCast(toolToken, JObject)
            If toolObj Is Nothing Then Continue For

            Dim functionObj = TryCast(toolObj("function"), JObject)
            If functionObj Is Nothing Then Continue For

            Dim converted As New JObject()
            converted("name") = If(functionObj("name"), "")
            converted("description") = If(functionObj("description"), "")
            If functionObj("parameters") IsNot Nothing Then converted("input_schema") = functionObj("parameters")
            anthropicTools.Add(converted)
        Next

        Return anthropicTools
    End Function

    Private Shared Function TokenToAnthropicContent(content As JToken) As JToken
        If content Is Nothing Then Return ""
        If content.Type = JTokenType.Array Then
            Dim converted As New JArray()
            For Each itemToken In DirectCast(content, JArray)
                Dim item = TryCast(itemToken, JObject)
                If item Is Nothing Then
                    converted.Add(itemToken.DeepClone())
                    Continue For
                End If

                If String.Equals(TokenToString(item("type")), "image_url", StringComparison.OrdinalIgnoreCase) Then
                    Dim imageBlock = ConvertOpenAiImageToAnthropic(item)
                    If imageBlock IsNot Nothing Then converted.Add(imageBlock)
                Else
                    converted.Add(item.DeepClone())
                End If
            Next
            Return converted
        End If
        Return TokenToString(content)
    End Function

    Private Shared Function ConvertOpenAiImageToAnthropic(item As JObject) As JObject
        Dim imageUrl = TryCast(item("image_url"), JObject)
        Dim dataUrl = If(imageUrl Is Nothing, "", TokenToString(imageUrl("url")))
        If String.IsNullOrWhiteSpace(dataUrl) OrElse
           Not dataUrl.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase) Then Return Nothing

        Dim commaIndex = dataUrl.IndexOf(","c)
        If commaIndex <= 5 OrElse commaIndex >= dataUrl.Length - 1 Then Return Nothing
        Dim header = dataUrl.Substring(5, commaIndex - 5)
        If Not header.EndsWith(";base64", StringComparison.OrdinalIgnoreCase) Then Return Nothing

        Dim mimeType = header.Substring(0, header.Length - ";base64".Length)
        If mimeType <> "image/png" AndAlso mimeType <> "image/jpeg" AndAlso
           mimeType <> "image/gif" AndAlso mimeType <> "image/webp" Then Return Nothing

        Return New JObject From {
            {"type", "image"},
            {"source", New JObject From {
                {"type", "base64"},
                {"media_type", mimeType},
                {"data", dataUrl.Substring(commaIndex + 1)}
            }}
        }
    End Function

    Private Shared Function ParseJsonOrString(value As String) As JToken
        If String.IsNullOrWhiteSpace(value) Then Return New JObject()
        Try
            Return JToken.Parse(value)
        Catch
            Return New JValue(value)
        End Try
    End Function

    Private Shared Function ExtractAnthropicContent(obj As JObject) As String
        Dim content = obj("content")
        If content Is Nothing Then Return ""

        If content.Type = JTokenType.String Then Return content.ToString()

        Dim contentArray = TryCast(content, JArray)
        If contentArray Is Nothing Then Return ""

        Dim sb As New StringBuilder()
        For Each contentItem In contentArray
            Dim itemObj = TryCast(contentItem, JObject)
            If itemObj Is Nothing Then Continue For
            If String.Equals(TokenToString(itemObj("type")), "text", StringComparison.OrdinalIgnoreCase) Then
                sb.Append(TokenToString(itemObj("text")))
            End If
        Next

        Return sb.ToString()
    End Function

    Private Shared Function ExtractResponsesOutputText(obj As JObject) As String
        Dim outputArray = TryCast(obj("output"), JArray)
        If outputArray Is Nothing Then Return ""

        Dim sb As New StringBuilder()
        For Each outputItemToken In outputArray
            Dim outputItem = TryCast(outputItemToken, JObject)
            If outputItem Is Nothing Then Continue For

            Dim contentArray = TryCast(outputItem("content"), JArray)
            If contentArray Is Nothing Then Continue For

            For Each contentItemToken In contentArray
                Dim contentItem = TryCast(contentItemToken, JObject)
                If contentItem Is Nothing Then Continue For

                Dim contentType = TokenToString(contentItem("type"))
                If String.Equals(contentType, "output_text", StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(contentType, "text", StringComparison.OrdinalIgnoreCase) Then
                    sb.Append(TokenToString(contentItem("text")))
                End If
            Next
        Next

        Return sb.ToString()
    End Function

    Private Shared Function TokenToString(token As JToken) As String
        If token Is Nothing Then Return ""
        If token.Type = JTokenType.Null Then Return ""
        Return token.ToString()
    End Function

    Private Shared Function Fail(errorMessage As String) As AiGatewayResponse
        Return New AiGatewayResponse With {
            .Success = False,
            .ErrorMessage = If(errorMessage, "")
        }
    End Function

    Private Shared Function Truncate(value As String, maxLength As Integer) As String
        If String.IsNullOrEmpty(value) OrElse value.Length <= maxLength Then Return If(value, "")
        Return value.Substring(0, maxLength)
    End Function
End Class
