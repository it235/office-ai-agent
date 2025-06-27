Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json

Public Class LLMUtil

    ' ����������
    Public Shared Function CreateRequestBody(question As String) As String
        Dim result As String = question.Replace("\", "\\").Replace("""", "\""").
                                  Replace(vbCr, "\r").Replace(vbLf, "\n").
                                  Replace(vbTab, "\t").Replace(vbBack, "\b").
                                  Replace(Chr(12), "\f")
        ' ʹ�ô� ConfigSettings �л�ȡ��ģ������
        Return "{""model"": """ & ConfigSettings.ModelName & """, ""messages"": [{""role"": ""user"", ""content"": """ & result & """}]}"
    End Function



    ' ����LLM API������
    Public Shared Function CreateLlmRequestBody(
        prompt As String,
        modelT As String,
        systemPrompt As String,
        temperatureT As Double,
        maxTokens As Integer) As String

        Try
            ' ������Ϣ����
            Dim messagesT As New List(Of Object)()

            ' ���ϵͳ��Ϣ������У�
            If Not String.IsNullOrEmpty(systemPrompt) Then
                messagesT.Add(New With {
                    .role = "system",
                    .content = systemPrompt
                })
            End If

            ' ����û���Ϣ
            messagesT.Add(New With {
                .role = "user",
                .content = prompt
            })

            ' ���������������
            Dim requestObj = New With {
                .model = modelT,
                .messages = messagesT,
                .temperature = temperatureT,
                .max_tokens = maxTokens,
                .stream = False  ' �ر���ʽ��Ӧ
            }

            ' ���л�ΪJSON
            Return JsonConvert.SerializeObject(requestObj)

        Catch ex As Exception
            Throw New Exception($"����������ʱ����: {ex.Message}")
        End Try
    End Function

    ' ���� HTTP ����
    'Public Shared Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
    '    Try
    '        ' ǿ��ʹ�� TLS 1.2
    '        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    '        Dim handler As New HttpClientHandler()
    '        Using client As New HttpClient(handler)
    '            client.Timeout = TimeSpan.FromSeconds(120) ' ���ó�ʱʱ��Ϊ 120 ��
    '            client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
    '            Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
    '            Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)
    '            response.EnsureSuccessStatusCode()
    '            Return Await response.Content.ReadAsStringAsync()
    '        End Using
    '    Catch ex As HttpRequestException
    '        MessageBox.Show("����ʧ��: " & ex.Message, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Return String.Empty
    '    End Try
    'End Function
    ' ���� HTTP ����
    Public Shared Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            Debug.WriteLine($"��ʼ����HTTP����: {apiUrl}")
            Debug.WriteLine($"����ͷAuthorization: Bearer {apiKey.Substring(0, Math.Min(10, apiKey.Length))}...")
            Debug.WriteLine($"�����峤��: {requestBody.Length}")

            ' ǿ��ʹ�� TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Dim handler As New HttpClientHandler()

            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromSeconds(120) ' ���ó�ʱʱ��Ϊ 120 ��
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)

                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Debug.WriteLine("���ڷ���POST����...")

                Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)

                Debug.WriteLine($"HTTP��Ӧ״̬��: {response.StatusCode}")
                Debug.WriteLine($"HTTP��Ӧԭ��: {response.ReasonPhrase}")

                ' �����Ӧ״̬
                If Not response.IsSuccessStatusCode Then
                    Dim errorContent As String = Await response.Content.ReadAsStringAsync()
                    Debug.WriteLine($"HTTP������Ӧ����: {errorContent}")
                    Throw New HttpRequestException($"HTTP����ʧ��: {response.StatusCode} - {response.ReasonPhrase}. ��ϸ��Ϣ: {errorContent}")
                End If

                Dim responseContent As String = Await response.Content.ReadAsStringAsync()
                Debug.WriteLine($"HTTP��Ӧ���ݳ���: {responseContent.Length}")
                Debug.WriteLine($"HTTP��Ӧ����ǰ200�ַ�: {responseContent.Substring(0, Math.Min(200, responseContent.Length))}")

                Return responseContent
            End Using

        Catch ex As TaskCanceledException
            Debug.WriteLine($"HTTP����ʱ: {ex.Message}")
            Return $"����: ����ʱ - {ex.Message}"
        Catch ex As HttpRequestException
            Debug.WriteLine($"HTTP�����쳣: {ex.Message}")
            ' ����ʾMessageBox��ֱ�ӷ��ش�����Ϣ
            Return $"����: HTTP����ʧ�� - {ex.Message}"
        Catch ex As Exception
            Debug.WriteLine($"����HTTP����ʱ����δ֪�쳣: {ex.Message}")
            Debug.WriteLine($"�쳣����: {ex.GetType().Name}")
            Debug.WriteLine($"�쳣��ջ: {ex.StackTrace}")
            Return $"����: {ex.Message}"
        End Try
    End Function
    ' ���ͬ���汾��HTTP���󷽷�
    Public Shared Function SendHttpRequestSync(apiUrl As String, apiKey As String, requestBody As String) As String
        Try

            ' ǿ��ʹ�� TLS 1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)

                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")

                ' ʹ�� .Result ����ͬ������
                Dim response As HttpResponseMessage = client.PostAsync(apiUrl, content).Result

                Debug.WriteLine($"HTTP��Ӧ״̬��: {response.StatusCode}")

                If Not response.IsSuccessStatusCode Then
                    Dim errorContent As String = response.Content.ReadAsStringAsync().Result
                    Debug.WriteLine($"HTTP������Ӧ����: {errorContent}")
                    Return $"����: HTTP����ʧ�� - {response.StatusCode} {response.ReasonPhrase}"
                End If

                Dim responseContent As String = response.Content.ReadAsStringAsync().Result
                Return responseContent
            End Using

        Catch ex As AggregateException
            ' ���� .Result ���ܲ����� AggregateException
            Dim innerEx = ex.GetBaseException()
            Return $"����: {innerEx.Message}"
        Catch ex As Exception
            Debug.WriteLine($"�쳣����: {ex.GetType().Name}")
            Return $"����: {ex.Message}"
        End Try
    End Function
End Class
