' ShareRibbon\Controls\Services\ChatStateService.vb
' 聊天状态管理服务：历史记录、选区映射、响应映射等

Imports System.Text
Imports Newtonsoft.Json.Linq

''' <summary>
''' 聊天状态管理服务，负责管理聊天历史、选区映射和响应映射
''' </summary>
Public Class ChatStateService
        ' 聊天历史记录
        Private ReadOnly _historyMessages As New List(Of HistoryMessage)()

        ' 选区映射：requestUuid -> SelectionInfo
        Private ReadOnly _selectionPendingMap As New Dictionary(Of String, SelectionInfo)()

        ' 响应到请求的映射：responseUuid -> requestUuid
        Private ReadOnly _responseToRequestMap As New Dictionary(Of String, String)()

        ' 响应到选区的映射：responseUuid -> SelectionInfo
        Private ReadOnly _responseSelectionMap As New Dictionary(Of String, SelectionInfo)()

        ' 响应模式映射：responseUuid -> mode (reformat, proofread, etc.)
        Private ReadOnly _responseModeMap As New Dictionary(Of String, String)()

        ' 修订映射：responseUuid -> JArray
        Private ReadOnly _revisionsMap As New Dictionary(Of String, JArray)()

        ' 上下文限制
        Private _contextLimit As Integer = 10

        ' Markdown 缓冲区
        Private ReadOnly _markdownBuffer As New StringBuilder()
        Private ReadOnly _plainMarkdownBuffer As New StringBuilder()

        ' Token 统计
        Private _currentSessionTotalTokens As Integer = 0
        Private _lastTokenInfo As Nullable(Of TokenInfo) = Nothing

        ' 第一个问题（用于文件命名）
        Private _firstQuestion As String = String.Empty
        Private _isFirstMessage As Boolean = True
        Private _chatHtmlFilePath As String = String.Empty

#Region "属性"

        ''' <summary>
        ''' 获取历史消息列表
        ''' </summary>
        Public ReadOnly Property HistoryMessages As List(Of HistoryMessage)
            Get
                Return _historyMessages
            End Get
        End Property

        ''' <summary>
        ''' 获取或设置上下文限制
        ''' </summary>
        Public Property ContextLimit As Integer
            Get
                Return _contextLimit
            End Get
            Set(value As Integer)
                _contextLimit = value
            End Set
        End Property

        ''' <summary>
        ''' 获取当前会话总 Token 数
        ''' </summary>
        Public Property CurrentSessionTotalTokens As Integer
            Get
                Return _currentSessionTotalTokens
            End Get
            Set(value As Integer)
                _currentSessionTotalTokens = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置最后的 Token 信息
        ''' </summary>
        Public Property LastTokenInfo As Nullable(Of TokenInfo)
            Get
                Return _lastTokenInfo
            End Get
            Set(value As Nullable(Of TokenInfo))
                _lastTokenInfo = value
            End Set
        End Property

        ''' <summary>
        ''' 获取 Markdown 缓冲区
        ''' </summary>
        Public ReadOnly Property MarkdownBuffer As StringBuilder
            Get
                Return _markdownBuffer
            End Get
        End Property

        ''' <summary>
        ''' 获取纯文本 Markdown 缓冲区
        ''' </summary>
        Public ReadOnly Property PlainMarkdownBuffer As StringBuilder
            Get
                Return _plainMarkdownBuffer
            End Get
        End Property

        ''' <summary>
        ''' 获取第一个问题
        ''' </summary>
        Public ReadOnly Property FirstQuestion As String
            Get
                Return _firstQuestion
            End Get
        End Property

        ''' <summary>
        ''' 获取修订映射
        ''' </summary>
        Public ReadOnly Property RevisionsMap As Dictionary(Of String, JArray)
            Get
                Return _revisionsMap
            End Get
        End Property

#End Region

#Region "历史管理"

        ''' <summary>
        ''' 添加消息到历史记录
        ''' </summary>
        Public Sub AddMessage(role As String, content As String)
            _historyMessages.Add(New HistoryMessage With {
                .role = role,
                .content = content,
                .Timestamp = DateTime.Now
            })
            ManageHistorySize()
        End Sub

        ''' <summary>
        ''' 添加或更新系统消息
        ''' </summary>
        Public Sub SetSystemMessage(content As String)
            Dim existingSystem = _historyMessages.FirstOrDefault(Function(m) m.role = "system")
            If existingSystem IsNot Nothing Then
                _historyMessages.Remove(existingSystem)
            End If
            _historyMessages.Insert(0, New HistoryMessage With {
                .role = "system",
                .content = content
            })
        End Sub

        ''' <summary>
        ''' 管理历史消息大小
        ''' </summary>
        Public Sub ManageHistorySize()
            ' 保留系统消息和最近的消息
            While _historyMessages.Count > _contextLimit + 2
                If _historyMessages.Count > 2 Then
                    _historyMessages.RemoveAt(2)
                End If
            End While
        End Sub

        ''' <summary>
        ''' 清空聊天历史
        ''' </summary>
        Public Sub ClearHistory()
            _historyMessages.Clear()
        End Sub

        ''' <summary>
        ''' 记录第一个问题
        ''' </summary>
        Public Sub RecordFirstQuestion(question As String)
            If _isFirstMessage AndAlso Not String.IsNullOrEmpty(question) Then
                _firstQuestion = question
                _isFirstMessage = False
                _chatHtmlFilePath = String.Empty
            End If
        End Sub

#End Region

#Region "选区映射"

        ''' <summary>
        ''' 绑定选区到请求
        ''' </summary>
        Public Sub BindSelectionToRequest(requestUuid As String, selectionInfo As SelectionInfo)
            If selectionInfo IsNot Nothing Then
                _selectionPendingMap(requestUuid) = selectionInfo
            End If
        End Sub

        ''' <summary>
        ''' 获取请求对应的选区
        ''' </summary>
        Public Function GetSelectionByRequest(requestUuid As String) As SelectionInfo
            If _selectionPendingMap.ContainsKey(requestUuid) Then
                Return _selectionPendingMap(requestUuid)
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' 移除请求的选区绑定
        ''' </summary>
        Public Sub RemoveSelectionBinding(requestUuid As String)
            If _selectionPendingMap.ContainsKey(requestUuid) Then
                _selectionPendingMap.Remove(requestUuid)
            End If
        End Sub

#End Region

#Region "响应映射"

        ''' <summary>
        ''' 建立响应到请求的映射
        ''' </summary>
        Public Sub MapResponseToRequest(responseUuid As String, requestUuid As String)
            _responseToRequestMap(responseUuid) = requestUuid
        End Sub

        ''' <summary>
        ''' 设置响应模式
        ''' </summary>
        Public Sub SetResponseMode(responseUuid As String, mode As String)
            If Not String.IsNullOrEmpty(mode) Then
                _responseModeMap(responseUuid) = mode
            End If
        End Sub

        ''' <summary>
        ''' 获取响应模式
        ''' </summary>
        Public Function GetResponseMode(responseUuid As String) As String
            If _responseModeMap.ContainsKey(responseUuid) Then
                Return _responseModeMap(responseUuid)
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' 迁移选区信息到响应映射
        ''' </summary>
        Public Sub MigrateSelectionToResponse(responseUuid As String, requestUuid As String)
            If Not String.IsNullOrEmpty(requestUuid) AndAlso _selectionPendingMap.ContainsKey(requestUuid) Then
                _responseSelectionMap(responseUuid) = _selectionPendingMap(requestUuid)
                _selectionPendingMap.Remove(requestUuid)
            End If
        End Sub

        ''' <summary>
        ''' 根据响应获取选区信息
        ''' </summary>
        Public Function GetSelectionByResponse(responseUuid As String) As SelectionInfo
            If _responseSelectionMap.ContainsKey(responseUuid) Then
                Return _responseSelectionMap(responseUuid)
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' 获取请求 UUID
        ''' </summary>
        Public Function GetRequestUuid(responseUuid As String) As String
            If _responseToRequestMap.ContainsKey(responseUuid) Then
                Return _responseToRequestMap(responseUuid)
            End If
            Return String.Empty
        End Function

#End Region

#Region "缓冲区管理"

        ''' <summary>
        ''' 清空所有缓冲区
        ''' </summary>
        Public Sub ClearBuffers()
            _markdownBuffer.Clear()
            _plainMarkdownBuffer.Clear()
        End Sub

        ''' <summary>
        ''' 重置会话 Token 计数
        ''' </summary>
        Public Sub ResetSessionTokens()
            _currentSessionTotalTokens = 0
            _lastTokenInfo = Nothing
        End Sub

        ''' <summary>
        ''' 累加 Token
        ''' </summary>
        Public Sub AddTokens(tokens As Integer)
            _currentSessionTotalTokens += tokens
        End Sub

#End Region

#Region "文件路径"

        ''' <summary>
        ''' 获取聊天 HTML 文件路径
        ''' </summary>
        Public Function GetChatHtmlFilePath() As String
            If Not String.IsNullOrEmpty(_chatHtmlFilePath) Then
                Return _chatHtmlFilePath
            End If

            Dim baseDir As String = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                ConfigSettings.OfficeAiAppDataFolder)

            Dim fileName As String
            If Not String.IsNullOrEmpty(_firstQuestion) Then
                Dim questionPrefix As String = GetFirst10Characters(_firstQuestion)
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}_{questionPrefix}.html"
            Else
                fileName = $"saved_chat_{DateTime.Now:yyyyMMdd_HHmmss}.html"
            End If

            _chatHtmlFilePath = System.IO.Path.Combine(baseDir, fileName)
            Return _chatHtmlFilePath
        End Function

        Private Function GetFirst10Characters(text As String) As String
            If String.IsNullOrEmpty(text) Then Return String.Empty

            Dim result As String = If(text.Length > 20, text.Substring(0, 20), text)
            Dim invalidChars As Char() = System.IO.Path.GetInvalidFileNameChars()

            For Each invalidChar In invalidChars
                result = result.Replace(invalidChar, "_"c)
            Next

            result = result.Replace(" ", "_").Replace(".", "_").Replace(",", "_").
                           Replace(":", "_").Replace("?", "_").Replace("!", "_")

            Return result
        End Function

#End Region

    End Class
