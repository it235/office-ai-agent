' ShareRibbon\Controls\Services\ChatStateService.vb
' 聊天状态管理服务：历史记录、选区映射、响应映射等

Imports System.Text

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

        ' 上下文限制
        Private _contextLimit As Integer = 10

        ' Markdown 缓冲区：只保留明文一份。转义是 JS 字面量通道的局部细节，
        ' 一旦成为可被历史/记忆读到的属性，就会被误当成正文落库。
        Private ReadOnly _plainMarkdownBuffer As New StringBuilder()

        ' Token 统计
        Private _currentSessionTotalTokens As Integer = 0
        Private _lastTokenInfo As Nullable(Of TokenInfo) = Nothing

        ' 第一个问题（用于文件命名）

        ' 当前会话 ID（用于持久化到 conversation 表）
        Private _currentSessionId As String = Nothing

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
        ''' 获取纯文本 Markdown 缓冲区
        ''' </summary>
        Public ReadOnly Property PlainMarkdownBuffer As StringBuilder
            Get
                Return _plainMarkdownBuffer
            End Get
        End Property

        ''' <summary>
        ''' 当前会话 ID，新建会话时生成 GUID
        ''' </summary>
        Public ReadOnly Property CurrentSessionId As String
            Get
                If String.IsNullOrEmpty(_currentSessionId) Then
                    _currentSessionId = Guid.NewGuid().ToString()
                End If
                Return _currentSessionId
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

            ' 持久化到 conversation 表（后台线程，避免阻塞UI）
            If role = "user" OrElse role = "assistant" Then
                Try
                    Dim sid = CurrentSessionId
                    Dim r = role
                    Dim c = content
                    Task.Run(Sub()
                                 Try
                                     ConversationRepository.InsertMessage(sid, r, c, False)
                                 Catch ex As Exception
                                     Debug.WriteLine("ConversationRepository.InsertMessage 后台写入失败: " & ex.Message)
                                 End Try
                             End Sub)
                Catch ex As Exception
                    Debug.WriteLine("ConversationRepository.InsertMessage 调度失败: " & ex.Message)
                End Try
            End If
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
        ''' 新建会话：生成新 session_id，清空缓冲区
        ''' </summary>
        Public Sub StartNewSession()
            _currentSessionId = Guid.NewGuid().ToString()
            _historyMessages.Clear()
            ClearBuffers()
            ResetSessionTokens()
        End Sub

        ''' <summary>
        ''' 切换到指定会话：从 conversation 表加载该会话消息并设为当前会话
        ''' </summary>
        Public Sub SwitchToSession(sessionId As String)
            If String.IsNullOrEmpty(sessionId) Then Return
            _currentSessionId = sessionId
            _historyMessages.Clear()
            ClearBuffers()
            ResetSessionTokens()
            Try
                Dim messages = ConversationRepository.GetMessagesBySession(sessionId)
                For Each dto In messages
                    Dim ts As DateTime = DateTime.Now
                    DateTime.TryParse(dto.CreateTime, ts)
                    _historyMessages.Add(New HistoryMessage With {
                        .role = dto.Role,
                        .content = dto.Content,
                        .Timestamp = ts
                    })
                Next
                ManageHistorySize()
            Catch ex As Exception
                Debug.WriteLine("SwitchToSession 加载消息失败: " & ex.Message)
            End Try
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
        ''' 暴露内部字典供 BaseChatControl 做属性代理（只读引用，同一实例）
        ''' </summary>
        Public ReadOnly Property ResponseToRequestMap As Dictionary(Of String, String)
            Get
                Return _responseToRequestMap
            End Get
        End Property

        Public ReadOnly Property ResponseModeMap As Dictionary(Of String, String)
            Get
                Return _responseModeMap
            End Get
        End Property

        Public ReadOnly Property ResponseSelectionMap As Dictionary(Of String, SelectionInfo)
            Get
                Return _responseSelectionMap
            End Get
        End Property

        Public ReadOnly Property SelectionPendingMap As Dictionary(Of String, SelectionInfo)
            Get
                Return _selectionPendingMap
            End Get
        End Property

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

    End Class
