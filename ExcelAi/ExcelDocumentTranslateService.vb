Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

''' <summary>
''' Excel文档翻译服务 - 用于翻译单元格内容
''' 支持批量翻译优化（减少API调用次数）
''' </summary>
Public Class ExcelDocumentTranslateService

    ' 每批次最大翻译数量（避免单次请求过大）
    Private Const BATCH_SIZE As Integer = 20

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

        ' 根据输出模式预先插入行/列（使用改进后的逻辑）
        Dim insertedOffset As Integer = 0
        If cellRanges IsNot Nothing AndAlso cellRanges.Count > 0 Then
            insertedOffset = PrepareOutputSpace(cellRanges, settings.OutputMode)
        End If

        ' 批量翻译（减少API调用）
        Dim allResults = Await BatchTranslateAsync(cellTexts, systemPrompt, sourceLang, targetLang, apiUrl, apiKey, modelName)
        results.AddRange(allResults)

        ' 应用翻译结果到单元格
        If cellRanges IsNot Nothing Then
            ApplyAllTranslations(cellRanges, cellTexts, results, settings.OutputMode, insertedOffset)
        End If

        Return results
    End Function

    ''' <summary>
    ''' 批量翻译文本（使用编号格式减少API调用）
    ''' </summary>
    Private Async Function BatchTranslateAsync(texts As List(Of String),
                                                systemPrompt As String,
                                                sourceLang As String,
                                                targetLang As String,
                                                apiUrl As String,
                                                apiKey As String,
                                                modelName As String) As Task(Of List(Of String))
        Dim results As New List(Of String)()
        
        ' 分批处理
        Dim batchCount = Math.Ceiling(texts.Count / CDbl(BATCH_SIZE))
        
        For batchIndex = 0 To CInt(batchCount) - 1
            Dim startIdx = batchIndex * BATCH_SIZE
            Dim endIdx = Math.Min(startIdx + BATCH_SIZE, texts.Count)
            Dim batchTexts = texts.Skip(startIdx).Take(endIdx - startIdx).ToList()
            
            GlobalStatusStripAll.ShowWarning($"正在翻译批次 {batchIndex + 1}/{CInt(batchCount)}（共{texts.Count}个单元格）...")
            
            Try
                Dim batchResults = Await TranslateBatchAsync(batchTexts, systemPrompt, sourceLang, targetLang, apiUrl, apiKey, modelName)
                results.AddRange(batchResults)
            Catch ex As Exception
                ' 如果批量失败，回退到原文
                Debug.WriteLine($"批量翻译失败: {ex.Message}")
                For Each item In batchTexts
                    results.Add(item)
                Next
            End Try
            
            ' 批次间延迟
            If batchIndex < CInt(batchCount) - 1 Then
                Await Task.Delay(300)
            End If
        Next
        
        Return results
    End Function

    ''' <summary>
    ''' 翻译单个批次（使用编号格式）
    ''' </summary>
    Private Async Function TranslateBatchAsync(texts As List(Of String),
                                                systemPrompt As String,
                                                sourceLang As String,
                                                targetLang As String,
                                                apiUrl As String,
                                                apiKey As String,
                                                modelName As String) As Task(Of List(Of String))
        ' 构建编号格式的请求
        Dim sb As New StringBuilder()
        sb.AppendLine($"请将以下编号内容从{sourceLang}翻译为{targetLang}。")
        sb.AppendLine("严格按照相同的编号格式返回翻译结果，每行一个编号，格式为 [编号] 翻译内容")
        sb.AppendLine("不要添加任何额外解释，只返回翻译结果。")
        sb.AppendLine()
        
        For i = 0 To texts.Count - 1
            Dim cellText = texts(i)
            If String.IsNullOrWhiteSpace(cellText) Then
                sb.AppendLine($"[{i + 1}] ")
            Else
                ' 替换文本中的换行符为特殊标记，避免格式混乱
                Dim cleanText = cellText.Replace(vbCrLf, "<<BR>>").Replace(vbLf, "<<BR>>").Replace(vbCr, "<<BR>>")
                sb.AppendLine($"[{i + 1}] {cleanText}")
            End If
        Next
        
        Dim userContent = sb.ToString()
        Dim requestBody = CreateRequestBody(systemPrompt, userContent, modelName)
        Dim response = Await SendHttpRequestAsync(apiUrl, apiKey, requestBody)
        
        ' 解析响应
        Dim jObj = JObject.Parse(response)
        Dim responseText = jObj("choices")(0)("message")("content")?.ToString()
        
        ' 解析编号格式的响应
        Return ParseBatchResponse(responseText, texts)
    End Function

    ''' <summary>
    ''' 解析批量翻译响应
    ''' </summary>
    Private Function ParseBatchResponse(responseText As String, originalTexts As List(Of String)) As List(Of String)
        Dim results As New List(Of String)()
        
        ' 初始化结果列表，默认使用原文
        For Each item In originalTexts
            results.Add(item)
        Next
        
        If String.IsNullOrEmpty(responseText) Then
            Return results
        End If
        
        Try
            ' 使用正则表达式匹配 [编号] 内容 格式
            Dim pattern = "\[(\d+)\]\s*(.+?)(?=\[\d+\]|$)"
            Dim matches = Regex.Matches(responseText, pattern, RegexOptions.Singleline)
            
            For Each match As Match In matches
                Dim index = Integer.Parse(match.Groups(1).Value) - 1
                Dim translatedText = match.Groups(2).Value.Trim()
                
                ' 还原换行符
                translatedText = translatedText.Replace("<<BR>>", vbCrLf)
                
                If index >= 0 AndAlso index < results.Count Then
                    results(index) = translatedText
                End If
            Next
        Catch ex As Exception
            Debug.WriteLine($"解析批量响应失败: {ex.Message}")
        End Try
        
        Return results
    End Function

    ''' <summary>
    ''' 根据输出模式预先插入行或列（改进版：正确计算需要插入的数量）
    ''' </summary>
    ''' <returns>插入的行/列数</returns>
    Private Function PrepareOutputSpace(cellRanges As List(Of Range), outputMode As TranslateOutputMode) As Integer
        Try
            If cellRanges Is Nothing OrElse cellRanges.Count = 0 Then Return 0
            
            ' 分析选中区域的范围
            Dim minRow As Integer = Integer.MaxValue
            Dim maxRow As Integer = 0
            Dim minCol As Integer = Integer.MaxValue
            Dim maxCol As Integer = 0
            Dim worksheet As Worksheet = Nothing
            
            For Each cell In cellRanges
                If cell.Row < minRow Then minRow = cell.Row
                If cell.Row > maxRow Then maxRow = cell.Row
                If cell.Column < minCol Then minCol = cell.Column
                If cell.Column > maxCol Then maxCol = cell.Column
                If worksheet Is Nothing Then worksheet = cell.Worksheet
            Next
            
            If worksheet Is Nothing Then Return 0
            
            Select Case outputMode
                Case TranslateOutputMode.Immersive
                    ' 右侧模式：在选中区域最右边插入与选中区域列数相同的列
                    Dim colCount = maxCol - minCol + 1
                    ' 从右到左插入，避免列号变化影响
                    For i = 1 To colCount
                        Dim insertRange = worksheet.Columns(maxCol + 1)
                        insertRange.Insert(XlInsertShiftDirection.xlShiftToRight)
                    Next
                    Return colCount

                Case TranslateOutputMode.NewDocument
                    ' 下方模式：在选中区域最下方插入与选中区域行数相同的行
                    Dim rowCount = maxRow - minRow + 1
                    ' 从下到上插入，避免行号变化影响
                    For i = 1 To rowCount
                        Dim insertRange = worksheet.Rows(maxRow + 1)
                        insertRange.Insert(XlInsertShiftDirection.xlShiftDown)
                    Next
                    Return rowCount
            End Select
            
            Return 0
        Catch ex As Exception
            Debug.WriteLine($"预插入行列失败: {ex.Message}")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 应用所有翻译结果到单元格
    ''' </summary>
    Private Sub ApplyAllTranslations(cellRanges As List(Of Range),
                                      originalTexts As List(Of String),
                                      translatedTexts As List(Of String),
                                      outputMode As TranslateOutputMode,
                                      insertedOffset As Integer)
        Try
            If cellRanges Is Nothing OrElse translatedTexts Is Nothing Then Return
            
            For i = 0 To Math.Min(cellRanges.Count, translatedTexts.Count) - 1
                Dim cell = cellRanges(i)
                Dim translatedText = translatedTexts(i)
                
                Try
                    Select Case outputMode
                        Case TranslateOutputMode.Replace
                            ' 替换原文
                            cell.Value = translatedText

                        Case TranslateOutputMode.Immersive
                            ' 译文放在右侧：偏移量就是插入的列数
                            ' 例如选中A1,B1，插入了2列，A1译文放到C1(+2)，B1译文放到D1(+2)
                            Dim rightCell = cell.Offset(0, insertedOffset)
                            rightCell.Value = translatedText

                        Case TranslateOutputMode.SidePanel
                            ' 仅显示在侧栏，不修改单元格
                            ' 跳过

                        Case TranslateOutputMode.NewDocument
                            ' 译文放在下方：偏移量就是插入的行数
                            ' 例如选中A1,A2，插入了2行，A1译文放到A3(+2)，A2译文放到A4(+2)
                            Dim bottomCell = cell.Offset(insertedOffset, 0)
                            bottomCell.Value = translatedText
                    End Select
                Catch ex As Exception
                    Debug.WriteLine($"应用翻译到单元格 {i + 1} 失败: {ex.Message}")
                End Try
            Next
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
