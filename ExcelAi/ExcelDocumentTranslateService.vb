Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

''' <summary>
''' Excel文档翻译服务 - 用于翻译单元格内容，不拦截右键菜单
''' </summary>
Public Class ExcelDocumentTranslateService

    ''' <summary>
    ''' 批量翻译单元格内容
    ''' </summary>
    ''' <param name="cellTexts">要翻译的文本列表</param>
    ''' <param name="cellRanges">对应的单元格范围列表</param>
    ''' <param name="settings">翻译设置</param>
    ''' <returns>翻译结果列表</returns>
    Public Async Function TranslateCellsAsync(cellTexts As List(Of String),
                                               cellRanges As List(Of Range),
                                               settings As TranslateSettings) As Task(Of List(Of String))
        Dim results As New List(Of String)()

        If cellTexts Is Nothing OrElse cellTexts.Count = 0 Then
            Return results
        End If

        ' 获取翻译配置
        Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.translateSelected)
        If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
            GlobalStatusStripAll.ShowWarning("未配置翻译平台，请先在翻译配置中选择平台和模型")
            Return results
        End If

        Dim modelName = cfg.model.FirstOrDefault(Function(m) m.translateSelected)?.modelName
        If String.IsNullOrEmpty(modelName) Then modelName = cfg.model(0).modelName

        Dim apiUrl = cfg.url
        Dim apiKey = cfg.key

        ' 获取领域提示词
        Dim domainTemplate = TranslateDomainManager.GetTemplate(settings.CurrentDomain)
        Dim systemPrompt = If(domainTemplate IsNot Nothing, domainTemplate.SystemPrompt, settings.PromptText)

        Dim sourceLang = GetLanguageName(settings.SourceLanguage)
        Dim targetLang = GetLanguageName(settings.TargetLanguage)

        ' 逐个翻译单元格
        For i = 0 To cellTexts.Count - 1
            Try
                GlobalStatusStripAll.ShowWarning($"正在翻译 {i + 1}/{cellTexts.Count}...")

                Dim text = cellTexts(i)
                If String.IsNullOrWhiteSpace(text) Then
                    results.Add(text)
                    Continue For
                End If

                Dim userContent = $"请将以下内容从{sourceLang}翻译为{targetLang}，只输出翻译结果，不要添加任何解释：

{text}"

                Dim requestBody = CreateRequestBody(systemPrompt, userContent, modelName)
                Dim response = Await SendHttpRequestAsync(apiUrl, apiKey, requestBody)
                Dim jObj = JObject.Parse(response)
                Dim translatedText = jObj("choices")(0)("message")("content")?.ToString()

                results.Add(If(String.IsNullOrEmpty(translatedText), text, translatedText))

                ' 根据输出模式应用翻译结果
                If cellRanges IsNot Nothing AndAlso i < cellRanges.Count Then
                    ApplyTranslationToCell(cellRanges(i), text, translatedText, settings.OutputMode)
                End If

                ' 控制请求频率
                If i < cellTexts.Count - 1 Then
                    Await Task.Delay(200)
                End If

            Catch ex As Exception
                results.Add(cellTexts(i))
                Debug.WriteLine($"翻译单元格 {i + 1} 失败: {ex.Message}")
            End Try
        Next

        Return results
    End Function

    ''' <summary>
    ''' 应用翻译结果到单元格
    ''' </summary>
    Private Sub ApplyTranslationToCell(cell As Range, originalText As String, translatedText As String, outputMode As TranslateOutputMode)
        Try
            Select Case outputMode
                Case TranslateOutputMode.Replace
                    ' 替换原文
                    cell.Value = translatedText

                Case TranslateOutputMode.Immersive
                    ' 沉浸式：原文+译文并行显示
                    cell.Value = originalText & vbCrLf & translatedText

                Case TranslateOutputMode.SidePanel
                    ' 仅显示在侧栏，不修改单元格
                    GlobalTips.ShowWarning($"翻译结果: {translatedText}", 0, True)

                Case TranslateOutputMode.NewDocument
                    ' 在右侧新列显示
                    Dim nextCell = cell.Offset(0, 1)
                    nextCell.Value = translatedText
            End Select
        Catch ex As Exception
            Debug.WriteLine($"应用翻译结果失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取语言名称
    ''' </summary>
    Private Function GetLanguageName(code As String) As String
        Select Case code.ToLower()
            Case "auto" : Return "原语言"
            Case "zh" : Return "中文"
            Case "en" : Return "英文"
            Case "ja" : Return "日语"
            Case "ko" : Return "韩语"
            Case "fr" : Return "法语"
            Case "de" : Return "德语"
            Case "es" : Return "西班牙语"
            Case "ru" : Return "俄语"
            Case "pt" : Return "葡萄牙语"
            Case "it" : Return "意大利语"
            Case "vi" : Return "越南语"
            Case "th" : Return "泰语"
            Case "id" : Return "印尼语"
            Case "ar" : Return "阿拉伯语"
            Case Else : Return code
        End Select
    End Function

    ''' <summary>
    ''' 创建请求体
    ''' </summary>
    Private Function CreateRequestBody(systemPrompt As String, userContent As String, modelName As String) As String
        Dim requestObj As New JObject()
        requestObj("model") = modelName
        requestObj("stream") = False

        Dim messages As New JArray()
        messages.Add(New JObject() From {{"role", "system"}, {"content", systemPrompt}})
        messages.Add(New JObject() From {{"role", "user"}, {"content", userContent}})
        requestObj("messages") = messages

        Return requestObj.ToString(Newtonsoft.Json.Formatting.None)
    End Function

    ''' <summary>
    ''' 发送HTTP请求
    ''' </summary>
    Private Async Function SendHttpRequestAsync(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(120)
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
            Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
            Dim response = Await client.PostAsync(apiUrl, content)
            response.EnsureSuccessStatusCode()
            Return Await response.Content.ReadAsStringAsync()
        End Using
    End Function

End Class
