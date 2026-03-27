Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

''' <summary>
''' 批量数据生成服务：调用 AI 接口生成测试/示例数据并写入 Excel。
''' </summary>
Public Class BatchDataService

    ''' <summary>
    ''' 根据字段定义调用 AI 生成 JSON 数组文本。
    ''' 返回原始 content 字符串，调用方负责解析和写入 Excel。
    ''' </summary>
    Public Async Function GenerateBatchDataAsync(fields As List(Of FieldDefinition), rowCount As Integer) As Task(Of String)
        Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.translateSelected)
        If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
            GlobalStatusStripAll.ShowWarning("未配置 AI 平台，请在「翻译」配置中选择平台和模型")
            Return Nothing
        End If

        Dim modelName = cfg.model.FirstOrDefault(Function(m) m.translateSelected)?.modelName
        If String.IsNullOrEmpty(modelName) Then modelName = cfg.model(0).modelName

        ' 构建字段描述
        Dim fieldList = String.Join("、", fields.Select(Function(f)
                                                            If String.IsNullOrWhiteSpace(f.FieldDescription) Then
                                                                Return f.FieldName
                                                            Else
                                                                Return $"{f.FieldName}（{f.FieldDescription}）"
                                                            End If
                                                        End Function))

        Dim systemPrompt = "你是数据生成助手。只输出纯 JSON 数组，不使用 Markdown 代码块，不附加任何说明文字。"
        Dim userContent = $"请生成 {rowCount} 条随机测试数据，以 JSON 数组返回。" &
                          $"每条数据是 JSON 对象，包含以下字段：{fieldList}。" &
                          "只返回 JSON 数组，示例格式：[{{""字段名"": ""值""}}]"

        Dim requestBody = BuildRequestBody(systemPrompt, userContent, modelName)
        Dim raw = Await SendRequestAsync(cfg.url, cfg.key, requestBody)
        If String.IsNullOrEmpty(raw) Then Return Nothing

        Try
            Dim jObj = JObject.Parse(raw)
            Return jObj("choices")(0)("message")("content")?.ToString()
        Catch ex As Exception
            Debug.WriteLine($"[BatchDataService] 解析响应失败: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Private Function BuildRequestBody(systemPrompt As String, userContent As String, modelName As String) As String
        Dim esc = Function(s As String) As String
                      Return s.Replace("\", "\\") _
                               .Replace("""", "\""") _
                               .Replace(vbCr, "") _
                               .Replace(vbLf, "\n")
                  End Function
        Return $"{{""model"":""{modelName}"",""messages"":[" &
               $"{{""role"":""system"",""content"":""{esc(systemPrompt)}""}}," &
               $"{{""role"":""user"",""content"":""{esc(userContent)}""}}]," &
               $"""stream"":false}}"
    End Function

    Private Async Function SendRequestAsync(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            Debug.WriteLine($"[BatchDataService] HTTP 请求失败: {ex.Message}")
            Return String.Empty
        End Try
    End Function

End Class
