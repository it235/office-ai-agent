Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 翻译段落结果
''' </summary>
Public Class TranslateParagraphResult
    Public Property Index As Integer
    Public Property OriginalText As String
    Public Property TranslatedText As String
    Public Property Success As Boolean = True
    Public Property ErrorMessage As String = ""
End Class

''' <summary>
''' 翻译进度事件参数
''' </summary>
Public Class TranslateProgressEventArgs
    Inherits EventArgs
    Public Property Current As Integer
    Public Property Total As Integer
    Public Property Message As String
    Public Property Percentage As Integer
        Get
            If Total = 0 Then Return 0
            Return CInt((Current / CDbl(Total)) * 100)
        End Get
        Set(value As Integer)
        End Set
    End Property
End Class

''' <summary>
''' 文档翻译服务基类 - 支持批量翻译
''' </summary>
Public MustInherit Class DocumentTranslateService

    ''' <summary>翻译进度事件</summary>
    Public Event ProgressChanged As EventHandler(Of TranslateProgressEventArgs)

    ''' <summary>翻译完成事件</summary>
    Public Event TranslationCompleted As EventHandler(Of List(Of TranslateParagraphResult))

    ''' <summary>翻译设置</summary>
    Protected Property Settings As TranslateSettings

    ''' <summary>取消令牌</summary>
    Protected Property CancellationSource As CancellationTokenSource

    Public Sub New()
        Settings = TranslateSettings.Load()
    End Sub

    ''' <summary>
    ''' 获取要翻译的所有段落/文本块
    ''' </summary>
    Public MustOverride Function GetAllParagraphs() As List(Of String)

    ''' <summary>
    ''' 应用翻译结果到文档
    ''' </summary>
    Public MustOverride Sub ApplyTranslation(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)

    ''' <summary>
    ''' 获取选中的文本段落
    ''' </summary>
    Public MustOverride Function GetSelectedParagraphs() As List(Of String)

    ''' <summary>
    ''' 应用翻译结果到选中区域
    ''' </summary>
    Public MustOverride Sub ApplyTranslationToSelection(results As List(Of TranslateParagraphResult), outputMode As TranslateOutputMode)

    ''' <summary>
    ''' 取消翻译
    ''' </summary>
    Public Sub CancelTranslation()
        If CancellationSource IsNot Nothing Then
            CancellationSource.Cancel()
        End If
    End Sub

    ''' <summary>
    ''' 翻译所有内容
    ''' </summary>
    Public Async Function TranslateAllAsync() As Task(Of List(Of TranslateParagraphResult))
        Dim paragraphs = GetAllParagraphs()
        Return Await TranslateParagraphsAsync(paragraphs)
    End Function

    ''' <summary>
    ''' 翻译选中内容
    ''' </summary>
    Public Async Function TranslateSelectionAsync() As Task(Of List(Of TranslateParagraphResult))
        Dim paragraphs = GetSelectedParagraphs()
        Return Await TranslateParagraphsAsync(paragraphs)
    End Function

    ''' <summary>
    ''' 批量翻译段落
    ''' </summary>
    Protected Async Function TranslateParagraphsAsync(paragraphs As List(Of String)) As Task(Of List(Of TranslateParagraphResult))
        Dim results As New List(Of TranslateParagraphResult)()
        CancellationSource = New CancellationTokenSource()

        If paragraphs Is Nothing OrElse paragraphs.Count = 0 Then
            Return results
        End If

        Dim total = paragraphs.Count
        ' BatchSize=0 表示整批翻译（不分批）
        Dim batchSize = If(Settings.BatchSize <= 0, total, Settings.BatchSize)
        Dim totalBatches = CInt(Math.Ceiling(total / CDbl(batchSize)))

        ' 获取翻译配置
        Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.translateSelected)
        If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
            Throw New Exception("未配置翻译平台，请先在翻译配置中选择平台和模型")
        End If

        Dim modelName = cfg.model.FirstOrDefault(Function(m) m.translateSelected)?.modelName
        If String.IsNullOrEmpty(modelName) Then modelName = cfg.model(0).modelName

        Dim apiUrl = cfg.url
        Dim apiKey = cfg.key

        ' 获取领域提示词
        Dim domainTemplate = TranslateDomainManager.GetTemplate(Settings.CurrentDomain)
        Dim systemPrompt = If(domainTemplate IsNot Nothing, domainTemplate.SystemPrompt, Settings.PromptText)

        Dim sourceLang = Settings.SourceLanguage
        Dim targetLang = Settings.TargetLanguage

        ' 总体开始提示
        RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
            .Current = 0,
            .Total = total,
            .Message = $"AI翻译启动中... 共{total}段内容，分{totalBatches}批进行"
        })

        ' 按批次翻译
        Dim currentIndex = 0
        Dim currentBatch = 0
        While currentIndex < total
            If CancellationSource.Token.IsCancellationRequested Then
                Exit While
            End If

            currentBatch += 1
            Dim batch = paragraphs.Skip(currentIndex).Take(batchSize).ToList()

            ' 批次开始提示
            RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
                .Current = currentIndex,
                .Total = total,
                .Message = $"正在准备第{currentBatch}批（共{totalBatches}批），{batch.Count}段内容..."
            })

            Dim batchResults = Await TranslateBatchAsync(batch, currentIndex, currentBatch, totalBatches, apiUrl, apiKey, modelName, systemPrompt, sourceLang, targetLang)
            results.AddRange(batchResults)

            currentIndex += batch.Count

            ' 批次完成提示
            RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
                .Current = currentIndex,
                .Total = total,
                .Message = $"已完成 {currentIndex}/{total} ({CInt(currentIndex * 100.0 / total)}%)"
            })

            ' 控制请求频率（如果还有更多批次）
            If currentIndex < total Then
                RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
                    .Current = currentIndex,
                    .Total = total,
                    .Message = $"等待发送第{currentBatch + 1}批请求..."
                })
                Await Task.Delay(CInt(1000 / Settings.MaxRequestsPerSecond))
            End If
        End While

        RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
            .Current = total,
            .Total = total,
            .Message = "AI翻译全部完成，正在写入文档..."
        })

        RaiseEvent TranslationCompleted(Me, results)
        Return results
    End Function

    ''' <summary>
    ''' 翻译一批段落
    ''' </summary>
    Private Async Function TranslateBatchAsync(batch As List(Of String), startIndex As Integer, currentBatch As Integer, totalBatches As Integer,
                                                apiUrl As String, apiKey As String, modelName As String,
                                                systemPrompt As String, sourceLang As String, targetLang As String) As Task(Of List(Of TranslateParagraphResult))
        Dim results As New List(Of TranslateParagraphResult)()

        ' 构建批量翻译请求
        Dim contentBuilder As New StringBuilder()
        For i = 0 To batch.Count - 1
            Dim text = batch(i)
            If Not String.IsNullOrWhiteSpace(text) Then
                contentBuilder.AppendLine($"[{i}] {text}")
            End If
        Next

        If contentBuilder.Length = 0 Then
            For i = 0 To batch.Count - 1
                results.Add(New TranslateParagraphResult() With {
                    .Index = startIndex + i,
                    .OriginalText = batch(i),
                    .TranslatedText = batch(i),
                    .Success = True
                })
            Next
            Return results
        End If

        Dim userContent = $"请将以下内容从{GetLanguageName(sourceLang)}翻译为{GetLanguageName(targetLang)}。每个段落以[数字]开头，请保持相同格式输出，只输出翻译结果：

{contentBuilder}"

        Dim requestBody = CreateRequestBody(systemPrompt, userContent, modelName)
        Dim batchFailed As Boolean = False
        Dim batchException As Exception = Nothing

        ' 发送请求前提示
        RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
            .Current = startIndex,
            .Total = startIndex + batch.Count,
            .Message = $"第{currentBatch}/{totalBatches}批：正在请求AI翻译，请稍候..."
        })

        Try
            Dim response = Await SendHttpRequestAsync(apiUrl, apiKey, requestBody)

            ' 收到响应后提示
            RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
                .Current = startIndex,
                .Total = startIndex + batch.Count,
                .Message = $"第{currentBatch}/{totalBatches}批：正在解析AI返回结果..."
            })

            Dim jObj = JObject.Parse(response)
            Dim msg = jObj("choices")(0)("message")("content")?.ToString()

            If String.IsNullOrEmpty(msg) Then
                Throw New Exception("翻译结果为空")
            End If

            ' 解析翻译结果（AI 可能重排返回顺序，用字典按 [N] 索引匹配）
            Dim translatedDict = ParseBatchResponse(msg, batch.Count)

            For i = 0 To batch.Count - 1
                Dim translatedText As String = batch(i)
                If translatedDict.ContainsKey(i) Then
                    translatedText = translatedDict(i)
                End If

                results.Add(New TranslateParagraphResult() With {
                    .Index = startIndex + i,
                    .OriginalText = batch(i),
                    .TranslatedText = translatedText,
                    .Success = True
                })
            Next
        Catch ex As Exception
            batchFailed = True
            batchException = ex
        End Try

        ' 批量失败时，逐个重试
        If batchFailed Then
            RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
                .Current = startIndex,
                .Total = startIndex + batch.Count,
                .Message = $"第{currentBatch}/{totalBatches}批：批量翻译遇到问题，正在逐条重试..."
            })

            For i = 0 To batch.Count - 1
                If CancellationSource IsNot Nothing AndAlso CancellationSource.Token.IsCancellationRequested Then
                    Exit For
                End If

                ' 单条重试前提示
                RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
                    .Current = startIndex + i,
                    .Total = startIndex + batch.Count,
                    .Message = $"第{currentBatch}/{totalBatches}批：正在单独翻译第{i + 1}/{batch.Count}条..."
                })

                Try
                    Dim singleResult = Await TranslateSingleAsync(batch(i), apiUrl, apiKey, modelName, systemPrompt, sourceLang, targetLang, startIndex, batch.Count, i, batch.Count)
                    results.Add(New TranslateParagraphResult() With {
                        .Index = startIndex + i,
                        .OriginalText = batch(i),
                        .TranslatedText = singleResult,
                        .Success = True
                    })
                Catch singleEx As Exception
                    results.Add(New TranslateParagraphResult() With {
                        .Index = startIndex + i,
                        .OriginalText = batch(i),
                        .TranslatedText = batch(i),
                        .Success = False,
                        .ErrorMessage = singleEx.Message
                    })
                End Try
            Next
        End If

        Return results
    End Function

    ''' <summary>
    ''' 翻译单个段落
    ''' </summary>
    Private Async Function TranslateSingleAsync(text As String, apiUrl As String, apiKey As String,
                                                 modelName As String, systemPrompt As String,
                                                 sourceLang As String, targetLang As String,
                                                 Optional batchIndex As Integer = 0,
                                                 Optional batchTotal As Integer = 0,
                                                 Optional itemIndex As Integer = 0,
                                                 Optional itemTotal As Integer = 0) As Task(Of String)
        If String.IsNullOrWhiteSpace(text) Then
            Return text
        End If

        Dim userContent = $"请将以下内容从{GetLanguageName(sourceLang)}翻译为{GetLanguageName(targetLang)}，只输出翻译结果，不要添加任何解释：

{text}"

        Dim requestBody = CreateRequestBody(systemPrompt, userContent, modelName)

        ' 显示正在请求单条翻译的进度（仅在批量重试场景下）
        If batchTotal > 0 AndAlso itemTotal > 0 Then
            RaiseEvent ProgressChanged(Me, New TranslateProgressEventArgs() With {
                .Current = batchIndex + itemIndex,
                .Total = batchIndex + itemTotal,
                .Message = $"正在请求AI翻译第{itemIndex + 1}/{itemTotal}条，请稍候..."
            })
        End If

        Dim response = Await SendHttpRequestAsync(apiUrl, apiKey, requestBody)
        Dim jObj = JObject.Parse(response)
        Return jObj("choices")(0)("message")("content")?.ToString()
    End Function

    ''' <summary>
    ''' 解析批量翻译响应
    ''' </summary>
    Private Function ParseBatchResponse(response As String, expectedCount As Integer) As Dictionary(Of Integer, String)
        Dim results As New Dictionary(Of Integer, String)()
        Dim lines = response.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
        Dim currentIndex = -1
        Dim currentText As New StringBuilder()

        For Each line In lines
            ' 检查是否是新段落开始 [数字]
            Dim match = System.Text.RegularExpressions.Regex.Match(line, "^\[(\d+)\]\s*(.*)$")
            If match.Success Then
                ' 保存前一个段落到字典（key = [N] 中的 N）
                If currentIndex >= 0 Then
                    results(currentIndex) = currentText.ToString().Trim()
                End If
                currentIndex = Integer.Parse(match.Groups(1).Value)
                currentText.Clear()
                currentText.AppendLine(match.Groups(2).Value)
            ElseIf currentIndex >= 0 Then
                currentText.AppendLine(line)
            End If
        Next

        ' 保存最后一个段落
        If currentIndex >= 0 Then
            results(currentIndex) = currentText.ToString().Trim()
        End If

        ' 如果解析失败，将整个响应放入索引 0
        If results.Count = 0 AndAlso Not String.IsNullOrEmpty(response) Then
            results(0) = response.Trim()
        End If

        Return results
    End Function

    ''' <summary>
    ''' 获取语言名称
    ''' </summary>
    Protected Function GetLanguageName(code As String) As String
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
    Protected Function CreateRequestBody(systemPrompt As String, userContent As String, modelName As String) As String
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
    Protected Async Function SendHttpRequestAsync(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Using client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(180)
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
            Dim content As New StringContent(requestBody, Encoding.UTF8, "application/json")
            Dim response = Await client.PostAsync(apiUrl, content)
            response.EnsureSuccessStatusCode()
            Return Await response.Content.ReadAsStringAsync()
        End Using
    End Function
End Class
