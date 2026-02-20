' ShareRibbon\Services\EmbeddingService.vb
' Embedding 向量服务：调用 OpenAI Embedding API 生成文本向量
Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' Embedding 向量服务
''' </summary>
Public Class EmbeddingService

    ''' <summary>
    ''' 根据 API 端点获取默认的 Embedding 模型
    ''' </summary>
    Private Shared Function GetDefaultEmbeddingModel(apiUrl As String) As String
        If String.IsNullOrWhiteSpace(apiUrl) Then
            Return "text-embedding-3-small"
        End If

        Dim urlLower = apiUrl.ToLowerInvariant()

        ' 硅基流动 (SiliconFlow)
        If urlLower.Contains("siliconflow") Then
            Return "BAAI/bge-large-zh-v1.5"
        End If

        ' 阿里云百炼 (Qwen/DashScope)
        If urlLower.Contains("dashscope") OrElse urlLower.Contains("aliyun") Then
            Return "text-embedding-v3"
        End If

        ' 智谱清言 (GLM)
        If urlLower.Contains("bigmodel") OrElse urlLower.Contains("zhipu") Then
            Return "embedding-3"
        End If

        ' 腾讯混元
        If urlLower.Contains("hunyuan") OrElse urlLower.Contains("tencent") Then
            Return "hunyuan-embedding"
        End If

        ' 百度千帆
        If urlLower.Contains("qianfan") OrElse urlLower.Contains("baidu") Then
            Return "embedding-v1"
        End If

        ' 深度求索 (DeepSeek)
        If urlLower.Contains("deepseek") Then
            Return "deepseek-embedding"
        End If

        ' OpenAI
        If urlLower.Contains("openai") Then
            Return "text-embedding-3-small"
        End If

        ' 默认返回 OpenAI 格式
        Return "text-embedding-3-small"
    End Function

    ''' <summary>
    ''' 调用 OpenAI Embedding API 生成单个文本的向量
    ''' </summary>
    Public Shared Async Function GetEmbeddingAsync(text As String) As Task(Of Single())
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim apiUrl = ConfigSettings.ApiUrl
            Dim apiKey = ConfigSettings.ApiKey

            If String.IsNullOrWhiteSpace(apiUrl) OrElse String.IsNullOrWhiteSpace(apiKey) Then
                Debug.WriteLine("[EmbeddingService] API 配置不完整")
                Return Nothing
            End If

            ' 适配不同的 API 格式（OpenAI 格式 vs 其他）
            Dim isOpenAiFormat = apiUrl.Contains("openai.com") OrElse
                                  apiUrl.Contains("/v1/embeddings") OrElse
                                  Not apiUrl.Contains("anthropic.com")

            Dim embeddingUrl As String
            If apiUrl.Contains("/v1/chat/completions") Then
                embeddingUrl = apiUrl.Replace("/v1/chat/completions", "/v1/embeddings")
            ElseIf Not apiUrl.EndsWith("/embeddings") AndAlso Not apiUrl.Contains("/embeddings") Then
                If apiUrl.EndsWith("/") Then
                    embeddingUrl = apiUrl & "v1/embeddings"
                Else
                    embeddingUrl = apiUrl & "/v1/embeddings"
                End If
            Else
                embeddingUrl = apiUrl
            End If

            ' 根据 API 提供商选择合适的 Embedding 模型
            Dim defaultModel = GetDefaultEmbeddingModel(apiUrl)
            Dim modelToUse = If(String.IsNullOrWhiteSpace(ConfigSettings.EmbeddingModel), defaultModel, ConfigSettings.EmbeddingModel)

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(30)

                Dim request As New HttpRequestMessage(HttpMethod.Post, embeddingUrl)
                request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)

                ' 构建请求体
                Dim requestObj As New JObject()
                requestObj("model") = modelToUse
                requestObj("input") = text

                request.Content = New StringContent(requestObj.ToString(), Encoding.UTF8, "application/json")

                Debug.WriteLine($"[EmbeddingService] 调用 Embedding API: {embeddingUrl}")
                Debug.WriteLine($"[EmbeddingService] 使用模型: {modelToUse}")
                Debug.WriteLine($"[EmbeddingService] 输入文本长度: {text.Length}")

                Using response As HttpResponseMessage = Await client.SendAsync(request)
                    If Not response.IsSuccessStatusCode Then
                        Dim errorContent = Await response.Content.ReadAsStringAsync()
                        Debug.WriteLine($"[EmbeddingService] API 请求失败: {response.StatusCode} - {errorContent}")
                        Debug.WriteLine($"[EmbeddingService] 如果模型不支持，请在配置中设置正确的 EmbeddingModel")
                        Return Nothing
                    End If

                    Dim jsonContent = Await response.Content.ReadAsStringAsync()
                    Dim responseObj = JObject.Parse(jsonContent)

                    ' 解析向量
                    Dim dataArray = responseObj("data")
                    If dataArray IsNot Nothing AndAlso dataArray.Type = JTokenType.Array AndAlso dataArray.Count > 0 Then
                        Dim embeddingArray = dataArray(0)("embedding")
                        If embeddingArray IsNot Nothing AndAlso embeddingArray.Type = JTokenType.Array Then
                            Dim result As New List(Of Single)()
                            For Each valx In embeddingArray
                                result.Add(Convert.ToSingle(valx))
                            Next
                            Debug.WriteLine($"[EmbeddingService] 成功生成向量，维度: {result.Count}")
                            Return result.ToArray()
                        End If
                    End If

                    Debug.WriteLine($"[EmbeddingService] 无法解析 API 响应")
                    Return Nothing
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[EmbeddingService] 出错: {ex.Message}")
            Debug.WriteLine($"[EmbeddingService] 堆栈: {ex.StackTrace}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 计算两个向量的余弦相似度
    ''' </summary>
    Public Shared Function CosineSimilarity(v1 As Single(), v2 As Single()) As Single
        If v1 Is Nothing OrElse v2 Is Nothing OrElse v1.Length <> v2.Length Then
            Return 0.0F
        End If

        Dim dotProduct As Single = 0.0F
        Dim norm1 As Single = 0.0F
        Dim norm2 As Single = 0.0F

        For i As Integer = 0 To v1.Length - 1
            dotProduct += v1(i) * v2(i)
            norm1 += v1(i) * v1(i)
            norm2 += v2(i) * v2(i)
        Next

        If norm1 = 0.0F OrElse norm2 = 0.0F Then
            Return 0.0F
        End If

        Return dotProduct / (Math.Sqrt(norm1) * Math.Sqrt(norm2))
    End Function

    ''' <summary>
    ''' 将向量序列化为 JSON 字符串存储
    ''' </summary>
    Public Shared Function SerializeVector(vector As Single()) As String
        If vector Is Nothing OrElse vector.Length = 0 Then
            Return Nothing
        End If
        Return JsonConvert.SerializeObject(vector)
    End Function

    ''' <summary>
    ''' 从 JSON 字符串反序列化向量
    ''' </summary>
    Public Shared Function DeserializeVector(json As String) As Single()
        If String.IsNullOrWhiteSpace(json) Then
            Return Nothing
        End If
        Try
            Return JsonConvert.DeserializeObject(Of Single())(json)
        Catch
            Return Nothing
        End Try
    End Function

End Class
