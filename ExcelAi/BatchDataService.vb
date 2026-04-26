Imports System.Diagnostics
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

''' <summary>
''' 批量数据生成服务：调用 AI 接口生成测试/示例数据并写入 Excel。
''' </summary>
Public Class BatchDataService

    ''' <summary>
    ''' 根据字段定义调用 AI 生成 JSON 数组文本。
    ''' 返回原始 content 字符串；失败时返回 Nothing（并通过状态栏告知用户）。
    ''' </summary>
    Public Async Function GenerateBatchDataAsync(fields As List(Of FieldDefinition), rowCount As Integer) As Task(Of String)
        ' 复用翻译功能的 AI 平台配置，避免要求用户单独配置第二套 key
        Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.translateSelected)
        If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
            GlobalStatusStripAll.ShowWarning("未配置 AI 平台，请在「翻译」配置中选择平台和模型")
            Return Nothing
        End If

        Dim modelName = cfg.model.FirstOrDefault(Function(m) m.translateSelected)?.modelName
        If String.IsNullOrEmpty(modelName) Then modelName = cfg.model(0).modelName

        ' 有描述时附上，AI 据此决定字段的取值范围/格式（如"中文姓名"比"姓名"生成质量更高）
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

        ' 使用 JsonConvert.SerializeObject 而非手写字符串拼接：
        ' 手写转义容易遗漏 \t、\b、\f 及 U+0000~U+001F 控制字符，导致 API 返回 400/parse error
        Dim requestBody = BuildRequestBody(systemPrompt, userContent, modelName)

        ' 复用 LLMUtil.SendHttpRequest：已在生产验证过的 HTTP 客户端逻辑，避免重复维护
        ' LLMUtil 在出错时返回以 "错误:" 开头的字符串而不是抛出异常
        Dim raw = Await LLMUtil.SendHttpRequest(cfg.url, cfg.key, requestBody)

        If String.IsNullOrEmpty(raw) OrElse raw.StartsWith("错误:") Then
            GlobalStatusStripAll.ShowWarning(If(String.IsNullOrEmpty(raw), "AI 请求失败，请检查配置", raw))
            Return Nothing
        End If

        Try
            Dim jObj = JObject.Parse(raw)
            Return jObj("choices")(0)("message")("content")?.ToString()
        Catch ex As Exception
            Debug.WriteLine($"[BatchDataService] 解析响应失败: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 使用 JsonConvert 序列化请求体，确保所有特殊字符都被正确转义。
    ''' </summary>
    Private Shared Function BuildRequestBody(systemPrompt As String, userContent As String, modelName As String) As String
        Dim requestObj = New With {
            .model = modelName,
            .messages = New Object() {
                New With {.role = "system", .content = systemPrompt},
                New With {.role = "user", .content = userContent}
            },
            .stream = False
        }
        Return JsonConvert.SerializeObject(requestObj)
    End Function

End Class
