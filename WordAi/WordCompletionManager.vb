Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms
Imports ShareRibbon
Imports WordAi.Services

''' <summary>
''' Word文档补全管理器 - 提供实时AI补全功能（使用内联灰色文本）
''' </summary>
Public Class WordCompletionManager
    Private Shared _instance As WordCompletionManager
    Private Shared ReadOnly _lock As New Object()

    Private _wordApp As Microsoft.Office.Interop.Word.Application
    Private _completionService As OfficeCompletionService
    Private _ghostTextManager As WordGhostTextManager  ' 灰色文本管理器
    Private _isEnabled As Boolean = False
    Private _debounceTimer As System.Threading.Timer
    Private _lastParagraphText As String = ""
    Private _lastFullParagraphContent As String = ""  ' 用于检测内容是否真正变化
    Private _lastParaStart As Integer = -1  ' 上次段落起始位置，用于检测是否切换了段落
    Private _uiSyncContext As SynchronizationContext  ' UI线程同步上下文
    Private _cancellationTokenSource As CancellationTokenSource

    ' 快捷键监听相关
    Private _keyPollTimer As System.Windows.Forms.Timer  ' 按键轮询定时器
    Private _lastTriggerState As Boolean = False  ' 上次快捷键触发状态

    Private Const DEBOUNCE_DELAY_MS As Integer = 800
    Private Const KEY_POLL_INTERVAL_MS As Integer = 50  ' 按键轮询间隔

    ' Win32 API 声明
    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(vKey As Integer) As Short
    End Function

    ' 虚拟键码常量
    Private Const VK_CONTROL As Integer = &H11   ' Ctrl键
    Private Const VK_MENU As Integer = &H12      ' Alt键
    Private Const VK_RETURN As Integer = &HD     ' Enter键
    Private Const VK_RIGHT As Integer = &H27     ' 右箭头键
    Private Const VK_OEM_2 As Integer = &HBF     ' / 键
    Private Const VK_OEM_PERIOD As Integer = &HBE ' . 键

    ''' <summary>
    ''' 获取单例实例
    ''' </summary>
    Public Shared ReadOnly Property Instance As WordCompletionManager
        Get
            If _instance Is Nothing Then
                SyncLock _lock
                    If _instance Is Nothing Then
                        _instance = New WordCompletionManager()
                    End If
                End SyncLock
            End If
            Return _instance
        End Get
    End Property

    Private Sub New()
        _completionService = OfficeCompletionService.Instance
        ' 捕获当前（主）线程的同步上下文
        _uiSyncContext = SynchronizationContext.Current
        If _uiSyncContext Is Nothing Then
            _uiSyncContext = New WindowsFormsSynchronizationContext()
        End If
        _cancellationTokenSource = New CancellationTokenSource()

        ' 初始化按键轮询定时器
        _keyPollTimer = New System.Windows.Forms.Timer()
        _keyPollTimer.Interval = KEY_POLL_INTERVAL_MS
        AddHandler _keyPollTimer.Tick, AddressOf OnKeyPollTick
    End Sub

    ''' <summary>
    ''' 初始化Word补全功能
    ''' </summary>
    Public Sub Initialize(wordApp As Microsoft.Office.Interop.Word.Application)
        _wordApp = wordApp

        ' 创建灰色文本管理器
        _ghostTextManager = New WordGhostTextManager(wordApp)

        ' 监听选区变化事件
        AddHandler _wordApp.WindowSelectionChange, AddressOf OnSelectionChange

        Debug.WriteLine("WordCompletionManager 已初始化（Ghost Text 模式）")
    End Sub

    ''' <summary>
    ''' 启用/禁用补全
    ''' </summary>
    Public Property Enabled As Boolean
        Get
            Return _isEnabled
        End Get
        Set(value As Boolean)
            _isEnabled = value
            _completionService.Enabled = value
            If Not value Then
                ClearGhostText()
            End If
        End Set
    End Property

    ''' <summary>
    ''' 选区变化事件处理
    ''' </summary>
    Private Sub OnSelectionChange(sel As Microsoft.Office.Interop.Word.Selection)
        Try
            If Not _isEnabled OrElse Not ChatSettings.EnableAutocomplete Then
                Return
            End If

            ' 检查是否应该保留 ghost text（光标仍在原位）
            If _ghostTextManager IsNot Nothing AndAlso _ghostTextManager.HasGhostText Then
                If Not _ghostTextManager.IsCursorAtGhostTextStart() Then
                    ' 光标已移动，同步清除 ghost text（避免Post延迟导致竞态）
                    CancelPendingOperations()
                    _ghostTextManager.ClearGhostTextDirect()
                Else
                    ' 光标仍在原位，ghost text 还在显示，不需要新的请求
                    Return
                End If
            End If

            ' 获取当前段落文本
            If sel Is Nothing OrElse sel.Range Is Nothing Then
                Return
            End If

            ' 跳过多段落选区（如 Ctrl+A 全选）
            Try
                If sel.Range.Paragraphs.Count > 1 Then
                    Return
                End If
            Catch
                ' COM异常忽略
            End Try

            ' 跳过非插入点选区（选中了文本块）
            If sel.Type <> Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP Then
                Return
            End If

            Dim currentParagraph = sel.Range.Paragraphs(1)
            If currentParagraph Is Nothing Then Return

            ' 重新读取段落内容（ghost text 清除后内容可能已变化）
            Dim paragraphText = currentParagraph.Range.Text

            If String.IsNullOrWhiteSpace(paragraphText) OrElse paragraphText.Length < 5 Then
                Return
            End If

            Dim cleanParagraphText = paragraphText.TrimEnd(vbCr, vbLf)
            Dim currentParaStart = currentParagraph.Range.Start

            ' 检测是否切换到了不同段落（光标移动到其他段落，不触发补全）
            If currentParaStart <> _lastParaStart Then
                _lastParaStart = currentParaStart
                _lastFullParagraphContent = cleanParagraphText
                _lastParagraphText = ""
                Return  ' 仅更新跟踪状态，不触发补全
            End If

            ' 同一段落内：检查内容是否真正变化
            If cleanParagraphText = _lastFullParagraphContent Then
                Return
            End If

            ' 更新段落内容记录
            _lastFullParagraphContent = cleanParagraphText

            ' 获取光标前的文本
            Dim cursorPos = sel.Range.Start
            Dim paraStart = currentParagraph.Range.Start
            Dim textBeforeCursor = ""

            If cursorPos > paraStart Then
                Dim beforeRange = _wordApp.ActiveDocument.Range(paraStart, cursorPos)
                textBeforeCursor = beforeRange.Text
            End If

            If String.IsNullOrWhiteSpace(textBeforeCursor) OrElse textBeforeCursor.Length < 3 Then
                Return
            End If

            ' 检查光标位置的文本是否变化（二次防抖）
            If textBeforeCursor = _lastParagraphText Then
                Return
            End If

            ' 取消之前的定时器和请求
            CancelPendingOperations()

            _lastParagraphText = textBeforeCursor

            ' 设置防抖定时器
            _debounceTimer = New System.Threading.Timer(
                Sub(state)
                    RequestCompletion(textBeforeCursor)
                End Sub,
                Nothing,
                DEBOUNCE_DELAY_MS,
                System.Threading.Timeout.Infinite
            )

        Catch ex As Exception
            Debug.WriteLine($"OnSelectionChange 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 取消待处理的操作
    ''' </summary>
    Private Sub CancelPendingOperations()
        ' 取消定时器
        If _debounceTimer IsNot Nothing Then
            _debounceTimer.Dispose()
            _debounceTimer = Nothing
        End If

        ' 取消进行中的请求
        If _cancellationTokenSource IsNot Nothing Then
            Try
                _cancellationTokenSource.Cancel()
            Catch
            End Try
        End If
        _cancellationTokenSource = New CancellationTokenSource()

        ' 取消服务中的请求
        _completionService.CancelPendingRequest()
    End Sub

    ''' <summary>
    ''' 请求补全
    ''' </summary>
    Private Async Sub RequestCompletion(inputText As String)
        Try
            ' 检查输入是否已变化
            If inputText <> _lastParagraphText Then
                Debug.WriteLine("[WordCompletion] 输入已变化，跳过请求")
                Return
            End If

            Debug.WriteLine($"[WordCompletion] 开始请求补全: '{inputText}'")

            ' 获取取消令牌
            Dim token = _cancellationTokenSource.Token

            ' 调用补全服务获取结果
            Dim completions = Await _completionService.GetCompletionsDirectAsync(inputText, "Word", token)

            ' 检查是否已取消
            If token.IsCancellationRequested Then
                Debug.WriteLine("[WordCompletion] 请求已取消")
                Return
            End If

            Debug.WriteLine($"[WordCompletion] 获取到 {completions.Count} 个补全建议")

            ' 再次检查输入是否已变化
            If completions.Count > 0 AndAlso inputText = _lastParagraphText Then
                ' 显示第一个补全建议为灰色文本
                _ghostTextManager.ShowGhostText(completions(0))
                ' 启动按键轮询（检测快捷键接受补全）
                StartKeyPolling()
            End If

        Catch ex As OperationCanceledException
            Debug.WriteLine("[WordCompletion] 请求已取消")
        Catch ex As Exception
            Debug.WriteLine($"RequestCompletion 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 接受当前补全（将灰色文本变为正常文本）
    ''' </summary>
    Public Sub AcceptCurrentCompletion()
        If _ghostTextManager IsNot Nothing AndAlso _ghostTextManager.HasGhostText Then
            StopKeyPolling()
            _ghostTextManager.AcceptGhostText()
            _completionService.ClearCompletions()
            Debug.WriteLine("[WordCompletion] 已接受补全")
        End If
    End Sub

    ''' <summary>
    ''' 清除灰色文本
    ''' </summary>
    Public Sub ClearGhostText()
        StopKeyPolling()
        If _ghostTextManager IsNot Nothing Then
            _ghostTextManager.ClearGhostText()
        End If
        _completionService.ClearCompletions()
    End Sub

    ''' <summary>
    ''' 启动按键轮询（检测快捷键）
    ''' </summary>
    Private Sub StartKeyPolling()
        If _uiSyncContext IsNot Nothing Then
            _uiSyncContext.Post(Sub(state)
                                    _lastTriggerState = False
                                    _keyPollTimer.Start()
                                    Debug.WriteLine("[WordCompletion] 按键轮询已启动")
                                End Sub, Nothing)
        End If
    End Sub

    ''' <summary>
    ''' 停止按键轮询
    ''' </summary>
    Private Sub StopKeyPolling()
        If _keyPollTimer IsNot Nothing Then
            _keyPollTimer.Stop()
        End If
    End Sub

    ''' <summary>
    ''' 按键轮询回调 - 检测配置的快捷键
    ''' </summary>
    Private Sub OnKeyPollTick(sender As Object, e As EventArgs)
        Try
            ' 如果没有ghost text，停止轮询
            If Not HasGhostText Then
                StopKeyPolling()
                Return
            End If

            ' 根据配置检测快捷键
            Dim isTriggered As Boolean = CheckShortcutTriggered()

            ' 检测快捷键按下（从未按到按下的边缘触发）
            If isTriggered AndAlso Not _lastTriggerState Then
                Debug.WriteLine($"[WordCompletion] 检测到快捷键 '{ChatSettings.AutocompleteShortcut}'，接受补全")
                AcceptCurrentCompletion()
            End If

            _lastTriggerState = isTriggered

        Catch ex As Exception
            Debug.WriteLine($"[WordCompletion] OnKeyPollTick 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 检查配置的快捷键是否被触发
    ''' </summary>
    Private Function CheckShortcutTriggered() As Boolean
        Dim shortcut = ChatSettings.AutocompleteShortcut

        Select Case shortcut
            Case "Ctrl+Enter"
                Return IsKeyDown(VK_CONTROL) AndAlso IsKeyDown(VK_RETURN)
            Case "Alt+/"
                Return IsKeyDown(VK_MENU) AndAlso IsKeyDown(VK_OEM_2)
            Case "右箭头 →"
                Return IsKeyDown(VK_RIGHT)
            Case "Ctrl+."
                Return IsKeyDown(VK_CONTROL) AndAlso IsKeyDown(VK_OEM_PERIOD)
            Case Else
                ' 默认使用 Ctrl+Enter
                Return IsKeyDown(VK_CONTROL) AndAlso IsKeyDown(VK_RETURN)
        End Select
    End Function

    ''' <summary>
    ''' 检查按键是否按下
    ''' </summary>
    Private Function IsKeyDown(vKey As Integer) As Boolean
        Return (GetAsyncKeyState(vKey) And &H8000) <> 0
    End Function

    ''' <summary>
    ''' 检查是否有活动的 ghost text
    ''' </summary>
    Public ReadOnly Property HasGhostText As Boolean
        Get
            Return _ghostTextManager IsNot Nothing AndAlso _ghostTextManager.HasGhostText
        End Get
    End Property

    ''' <summary>
    ''' 清理资源
    ''' </summary>
    Public Sub Dispose()
        StopKeyPolling()
        If _keyPollTimer IsNot Nothing Then
            RemoveHandler _keyPollTimer.Tick, AddressOf OnKeyPollTick
            _keyPollTimer.Dispose()
            _keyPollTimer = Nothing
        End If
        CancelPendingOperations()
        If _ghostTextManager IsNot Nothing Then
            _ghostTextManager.Dispose()
        End If
        If _wordApp IsNot Nothing Then
            RemoveHandler _wordApp.WindowSelectionChange, AddressOf OnSelectionChange
        End If
    End Sub
End Class
