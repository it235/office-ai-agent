Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Windows.Forms

Public MustInherit Class BaseTranslateService

    Public Sub New()
    End Sub

    ' 获取当前选中的文本
    Public MustOverride Function GetSelectedText() As String

    ' 钩子：选区变化
    Public MustOverride Sub HookSelectionChange()

    ' 钩子：右键菜单
    Public MustOverride Sub HookRightClickMenu()

    ' 选区变化时自动翻译
    Protected Sub OnSelectionChanged()
        Dim settings = TranslateSettings.Load()
        If settings.EnableSelectionTranslate Then
            Dim txt = GetSelectedText()
            If Not String.IsNullOrEmpty(txt) AndAlso txt.Length < 500 Then
                TranslateTextAsync(txt)
            End If
        End If
    End Sub

    ' 右键菜单点击翻译
    Protected Sub OnRightClickTranslate()
        Dim txt = GetSelectedText()
        If Not String.IsNullOrEmpty(txt) Then
            TranslateTextAsync(txt)
        End If
    End Sub

    ' 右下角弹窗展示（复用 ChatControl 的方法）
    Protected Sub ShowPopup(result As String)
        ShowTranslatePopup(result)
    End Sub


    ' 异步翻译入口，自动读取翻译平台和模型
    Public Async Sub TranslateTextAsync(text As String)
        Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.translateSelected)
        If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
            ShowTranslatePopup("未配置翻译平台")
            Return
        End If
        Dim modelName = cfg.model.FirstOrDefault(Function(m) m.translateSelected)?.modelName
        If String.IsNullOrEmpty(modelName) Then modelName = cfg.model(0).modelName

        Dim settings = TranslateSettings.Load()
        Dim prompt = settings.PromptText
        Dim sourceLang = settings.SourceLanguage
        Dim targetLang = settings.TargetLanguage

        Dim apiUrl = cfg.url
        Dim apiKey = cfg.key

        ' 构造请求体，复用 CreateRequestBody
        Dim userContent = $"请将以下内容从{sourceLang}翻译为{targetLang}，保留原格式：{text}"
        Dim requestBody = CreateRequestBodyForTranslate(prompt, userContent, modelName)
        GlobalStatusStripAll.ShowWarning($"正在调用大模型翻译，请等待")
        Try
            Dim result = Await SendHttpRequest(apiUrl, apiKey, requestBody)
            Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(result)
            Dim msg = jObj("choices")(0)("message")("content")?.ToString()
            ShowTranslatePopup(If(String.IsNullOrEmpty(msg), "无翻译结果", msg))
        Catch ex As Exception
            ShowTranslatePopup("翻译失败：" & ex.Message)
        End Try
    End Sub

    ' 构造翻译请求体（复用原有 CreateRequestBody 逻辑，简化为只用 system/user）
    Protected Function CreateRequestBodyForTranslate(systemPrompt As String, userContent As String, modelName As String) As String
        ' 对内容进行JSON转义，避免特殊字符导致JSON格式错误
        Dim escapedSystemPrompt = EscapeJsonString(systemPrompt)
        Dim escapedUserContent = EscapeJsonString(userContent)
        
        Dim messages As New List(Of String) From {
            $"{{""role"": ""system"", ""content"": ""{escapedSystemPrompt}""}}",
            $"{{""role"": ""user"", ""content"": ""{escapedUserContent}""}}"
        }
        Dim messagesJson = String.Join(",", messages)
        Return $"{{""model"": ""{modelName}"", ""messages"": [{messagesJson}], ""stream"": false}}"
    End Function
    
    ' JSON字符串转义
    Private Function EscapeJsonString(input As String) As String
        If String.IsNullOrEmpty(input) Then Return ""
        Return input.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n").Replace(vbTab, "\t")
    End Function

    ' 右下角弹窗展示
    Public Sub ShowTranslatePopup(result As String)
        'GlobalStatusStripAll.ShowWarning($"翻译结果:{result}")
        GlobalTips.ShowWarning($"翻译结果:{result}", 0, True)
    End Sub



    ' 共用的HTTP请求方法
    Protected Async Function SendHttpRequest(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(120)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
                Dim response As HttpResponseMessage = Await client.PostAsync(apiUrl, content)
                response.EnsureSuccessStatusCode()
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            MessageBox.Show($"请求失败: {ex.Message}")
            Return String.Empty
        End Try
    End Function
End Class