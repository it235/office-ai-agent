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
    
    Private _debounceTimer As Timer
    Private _lastInputText As String = ""
    Private _isEnabled As Boolean = False
    Private _currentCompletions As List(Of String)
    Private _completionCallback As Action(Of List(Of String), System.Drawing.Point)
    
    ' 防抖延迟（毫秒）
    Private Const DEBOUNCE_DELAY_MS As Integer = 800
    
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
    ''' 设置补全回调（用于显示补全UI）
    ''' </summary>
    Public Sub SetCompletionCallback(callback As Action(Of List(Of String), System.Drawing.Point))
        _completionCallback = callback
    End Sub
    
    ''' <summary>
    ''' 用户输入变化时调用（带防抖）
    ''' </summary>
    Public Sub OnTextChanged(inputText As String, cursorPosition As System.Drawing.Point, appType As String)
        If Not _isEnabled OrElse Not ChatSettings.EnableAutocomplete Then
            Return
        End If
        
        ' 取消之前的定时器
        CancelPendingRequest()
        
        ' 输入太短不触发
        If String.IsNullOrWhiteSpace(inputText) OrElse inputText.Length < 3 Then
            Return
        End If
        
        _lastInputText = inputText
        
        ' 设置防抖定时器
        _debounceTimer = New Timer(
            Sub(state)
                RequestCompletionAsync(inputText, cursorPosition, appType)
            End Sub,
            Nothing,
            DEBOUNCE_DELAY_MS,
            Timeout.Infinite
        )
    End Sub
    
    ''' <summary>
    ''' 取消待处理的请求
    ''' </summary>
    Public Sub CancelPendingRequest()
        If _debounceTimer IsNot Nothing Then
            _debounceTimer.Dispose()
            _debounceTimer = Nothing
        End If
    End Sub
    
    ''' <summary>
    ''' 异步请求补全
    ''' </summary>
    Private Async Sub RequestCompletionAsync(inputText As String, cursorPosition As System.Drawing.Point, appType As String)
        Try
            ' 检查输入是否已变化
            If inputText <> _lastInputText Then
                Return
            End If
            
            Dim completions = Await GetCompletionsFromLLM(inputText, appType)
            
            ' 再次检查输入是否已变化
            If inputText <> _lastInputText Then
                Return
            End If
            
            _currentCompletions = completions
            
            ' 回调显示补全
            If completions.Count > 0 AndAlso _completionCallback IsNot Nothing Then
                _completionCallback.Invoke(completions, cursorPosition)
            End If
            
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletionAsync 出错: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 调用LLM获取补全
    ''' </summary>
    Private Async Function GetCompletionsFromLLM(inputText As String, appType As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
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
                completions = Await GetCompletionsWithFIM(inputText, selectedModel, apiKey)
            Else
                completions = Await GetCompletionsWithChat(inputText, appType, cfg, selectedModel, apiKey)
            End If
            
        Catch ex As Exception
            Debug.WriteLine($"GetCompletionsFromLLM 出错: {ex.Message}")
        End Try
        
        Return completions
    End Function
    
    ''' <summary>
    ''' 使用FIM API获取补全
    ''' </summary>
    Private Async Function GetCompletionsWithFIM(inputText As String, model As ConfigManager.ConfigItemModel, 
                                                  apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("prompt") = inputText
            requestObj("suffix") = ""
            requestObj("max_tokens") = 100
            requestObj("temperature") = 0.3
            requestObj("stream") = False
            
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(10)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestObj.ToString(), Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(model.fimUrl, content)
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
            End Using
            
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
                                                   apiKey As String) As Task(Of List(Of String))
        Dim completions As New List(Of String)()
        
        Try
            Dim systemPrompt = GetSystemPrompt(appType)
            
            Dim requestObj As New JObject()
            requestObj("model") = model.modelName
            requestObj("stream") = False
            requestObj("temperature") = 0.3
            
            Dim messages As New JArray()
            messages.Add(New JObject() From {{"role", "system"}, {"content", systemPrompt}})
            messages.Add(New JObject() From {{"role", "user"}, {"content", $"请补全以下文本（只返回补全部分，不要重复原文）：{vbCrLf}{inputText}"}})
            requestObj("messages") = messages
            
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(10)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New StringContent(requestObj.ToString(), Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(cfg.url, content)
                response.EnsureSuccessStatusCode()
                
                Dim responseBody = Await response.Content.ReadAsStringAsync()
                Dim jObj = JObject.Parse(responseBody)
                
                Dim msg = jObj("choices")?(0)?("message")?("content")?.ToString()
                If Not String.IsNullOrWhiteSpace(msg) AndAlso msg.Length <= 100 Then
                    completions.Add(msg.Trim())
                End If
            End Using
            
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
    Public Async Function GetCompletionsDirectAsync(inputText As String, appType As String) As Task(Of List(Of String))
        If Not _isEnabled OrElse Not ChatSettings.EnableAutocomplete Then
            Return New List(Of String)()
        End If
        
        Dim completions = Await GetCompletionsFromLLM(inputText, appType)
        _currentCompletions = completions
        Return completions
    End Function
    
    ''' <summary>
    ''' 获取当前补全列表
    ''' </summary>
    Public Function GetCurrentCompletions() As List(Of String)
        Return _currentCompletions
    End Function
    
    ''' <summary>
    ''' 清除当前补全
    ''' </summary>
    Public Sub ClearCompletions()
        _currentCompletions.Clear()
    End Sub
End Class
