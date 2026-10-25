Imports System.Diagnostics
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json.Linq

''' <summary>
''' Office原生补全服务 - 提供Word和PPT的AI补全功能
''' </summary>
Public Class OfficeCompletionService
    Private Shared _instance As OfficeCompletionService
    Private Shared ReadOnly _lock As New Object()
    
    ' 单例 HttpClient（避免频繁创建连接，提升性能）
    Private Shared ReadOnly _httpClient As New HttpClient() With {
        .Timeout = TimeSpan.FromSeconds(8)
    }
    
    Private _isEnabled As Boolean = False
    Private _currentCompletions As List(Of String)
    
    ' 取消令牌源（用于取消进行中的请求）
    Private _cancellationTokenSource As CancellationTokenSource
    
    ''' <summary>
    ''' 获取单例实例
    ''' </summary>
    Public Shared ReadOnly Property Instance As OfficeCompletionService
        Get
            If _instance Is Nothing Then
                SyncLock _lock
                    If _instance Is Nothing Then
                        _instance = New OfficeCompletionService()
                    End If
                End SyncLock
            End If
            Return _instance
        End Get
    End Property
    
    Private Sub New()
        _currentCompletions = New List(Of String)()
        _cancellationTokenSource = New CancellationTokenSource()
    End Sub
    
    ''' <summary>
    ''' 启用/禁用补全服务
    ''' </summary>
    Public Property Enabled As Boolean
        Get
            Return _isEnabled
        End Get
        Set(value As Boolean)
            _isEnabled = value
            If Not value Then
                CancelPendingRequest()
            End If
        End Set
    End Property
    
    ''' <summary>
    ''' 取消待处理的请求
    ''' </summary>
    Public Sub CancelPendingRequest()
        ' 取消进行中的 HTTP 请求
        If _cancellationTokenSource IsNot Nothing Then
            Try
                _cancellationTokenSource.Cancel()
            Catch
                ' 忽略取消异常
            End Try
        End If
        _cancellationTokenSource = New CancellationTokenSource()
    End Sub
    
    ''' <summary>
    ''' 调用LLM获取补全
    ''' </summary>
    Private Async Function GetCompletionsFromLLM(inputText As String, appType As String, 
                                                  Optional token As CancellationToken = Nothing) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            If token = Nothing Then token = _cancellationTokenSource.Token
            
            Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.selected)
            If cfg Is Nothing OrElse cfg.model Is Nothing OrElse cfg.model.Count = 0 Then
                Return completions
            End If
            
            Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.selected)
            If selectedModel Is Nothing Then selectedModel = cfg.model(0)
            
            Dim modelName = selectedModel.modelName
            Dim apiUrl = cfg.url
            Dim apiKey = cfg.key
            
            ' 检查是否支持FIM
            If selectedModel.fimSupported AndAlso Not String.IsNullOrEmpty(selectedModel.fimUrl) Then
                completions = Await GetCompletionsWithFIM(inputText, selectedModel, apiKey, token)
            Else
                completions = Await GetCompletionsWithChat(inputText, appType, cfg, selectedModel, apiKey, token)
            End If
            
        Catch ex As OperationCanceledException
            Debug.WriteLine("[Completion] 请求已取消")
        Catch ex As Exception
            Debug.WriteLine($"GetCompletionsFromLLM 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function
    
    ''' <summary>
    ''' 使用FIM API获取补全
    ''' </summary>
    Private Async Function GetCompletionsWithFIM(inputText As String, model As ConfigManager.ConfigItemModel, 
                                                  apiKey As String, token As CancellationToken) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        ' E3-05：补全载荷是用户当前文档内容，未同意不发；空列表就是本函数已有的「没有结果」。
        If PrivacyConsentGate.RefuseQuietly("OfficeCompletionService") Then Return completions
        
        Try
            token.ThrowIfCancellationRequested()
            
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("prompt") = inputText
            requestObj("suffix") = ""
            requestObj("max_tokens") = 60  ' 减少 token 数量以加快响应
            requestObj("temperature") = 0.3
            requestObj("stream") = False
            
            ' 使用单例 HttpClient
            Dim request As New HttpRequestMessage(HttpMethod.Post, model.fimUrl)
            request.Headers.Add("Authorization", "Bearer " & apiKey)
            request.Content = New StringContent(requestObj.ToString(), Encoding.UTF8, "application/json")
            
            Dim response = Await _httpClient.SendAsync(request, token)
            response.EnsureSuccessStatusCode()
            
            Dim responseBody = Await response.Content.ReadAsStringAsync()
            Dim jObj = JObject.Parse(responseBody)
            
            Dim text = jObj("choices")?(0)?("text")?.ToString()
            If Not String.IsNullOrWhiteSpace(text) Then
                ' 取第一行
                Dim firstLine = text.Trim().Split({vbCr, vbLf, vbCrLf}, StringSplitOptions.RemoveEmptyEntries)(0)
                If firstLine.Length <= 100 Then
                    completions.Add(firstLine)
                End If
            End If
            
        Catch ex As OperationCanceledException
            Debug.WriteLine("[FIM] 请求已取消")
        Catch ex As Exception
            Debug.WriteLine($"GetCompletionsWithFIM 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function
    
    ''' <summary>
    ''' 使用Chat API获取补全
    ''' </summary>
    Private Async Function GetCompletionsWithChat(inputText As String, appType As String,
                                                   cfg As ConfigManager.ConfigItem, model As ConfigManager.ConfigItemModel,
                                                   apiKey As String, token As CancellationToken) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        ' E3-05：chat 形态的补同同样把正文发出去，口径必须与 FIM 分支一致，否则「拦一半」形同没拦。
        If PrivacyConsentGate.RefuseQuietly("OfficeCompletionService") Then Return completions
        
        Try
            token.ThrowIfCancellationRequested()
            
            Dim systemPrompt = GetSystemPrompt(appType)
            
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("stream") = False
            requestObj("temperature") = 0.3
            
            Dim messages As New JArray()
            messages.Add(New JObject() From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject() From {{"role", "user"}, {"content", $"请补全以下文本（只返回补全部分，不要重复原文）：{vbCrLf}{inputText}"}})
            requestObj("messages") = messages
            
            ' 使用单例 HttpClient
            Dim request As New HttpRequestMessage(HttpMethod.Post, cfg.url)
            request.Headers.Add("Authorization", "Bearer " & apiKey)
            request.Content = New StringContent(requestObj.ToString(), Encoding.UTF8, "application/json")
            
            Dim response = Await _httpClient.SendAsync(request, token)
            response.EnsureSuccessStatusCode()
            
            Dim responseBody = Await response.Content.ReadAsStringAsync()
            Dim jObj = JObject.Parse(responseBody)
            
            Dim msg = jObj("choices")?(0)?("message")?("content")?.ToString()
            If Not String.IsNullOrWhiteSpace(msg) AndAlso msg.Length <= 100 Then
                completions.Add(msg.Trim())
            End If
            
        Catch ex As OperationCanceledException
            Debug.WriteLine("[Chat] 请求已取消")
        Catch ex As Exception
            Debug.WriteLine($"GetCompletionsWithChat 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function
    
    ''' <summary>
    ''' 获取系统提示词
    ''' </summary>
    Private Function GetSystemPrompt(appType As String) As String
        Select Case appType.ToLower()
            Case "word"
                Return "你是Word文档的智能补全助手。根据用户正在输入的文本，预测并补全后续内容。
规则：
1. 只返回补全的后续内容，不要重复用户已输入的部分
2. 补全应该自然流畅，符合上下文语境
3. 补全长度通常在10-50个字符
4. 保持与原文相同的语言和风格"
            
            Case "powerpoint"
                Return "你是PowerPoint演示文稿的智能补全助手。根据用户正在输入的标题或内容，预测并补全。
规则：
1. 只返回补全的后续内容，不要重复用户已输入的部分
2. 补全应简洁有力，适合演示文稿
3. 补全长度通常在10-30个字符
4. 如果是标题，保持简短精炼"
            
            Case Else
                Return "你是Office文档的智能补全助手。根据用户正在输入的文本，预测并补全后续内容。只返回补全部分，不要重复原文。"
        End Select
    End Function
    
    ''' <summary>
    ''' 直接获取补全（不带防抖，供外部调用）
    ''' </summary>
    Public Async Function GetCompletionsDirectAsync(inputText As String, appType As String, 
                                                     Optional token As CancellationToken = Nothing) As Task(Of List(Of String))
        If Not _isEnabled OrElse Not ChatSettings.EnableAutocomplete Then
            Return New List(Of String)()
        End If
        
        If token = Nothing Then token = _cancellationTokenSource.Token
        
        Dim completions = Await GetCompletionsFromLLM(inputText, appType, token)
        _currentCompletions = completions
        Return completions
    End Function
    
    ''' <summary>
    ''' 清除当前补全
    ''' </summary>
    Public Sub ClearCompletions()
        _currentCompletions.Clear()
    End Sub
End Class
