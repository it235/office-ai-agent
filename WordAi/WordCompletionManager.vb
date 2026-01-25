Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms
'Imports Microsoft.Office.Interop.Word
Imports ShareRibbon

''' <summary>
''' Word文档补全管理器 - 提供实时AI补全功能
''' </summary>
Public Class WordCompletionManager
    Private Shared _instance As WordCompletionManager
    Private Shared ReadOnly _lock As New Object()

    Private _wordApp As Microsoft.Office.Interop.Word.Application
    Private _completionService As OfficeCompletionService
    Private _completionPopup As CompletionPopupForm
    Private _isEnabled As Boolean = False
    Private _debounceTimer As System.Threading.Timer
    Private _lastParagraphText As String = ""
    Private _uiSyncContext As SynchronizationContext  ' UI线程同步上下文

    Private Const DEBOUNCE_DELAY_MS As Integer = 800

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
        ' 在UI线程上创建弹窗
        _completionPopup = New CompletionPopupForm()
        AddHandler _completionPopup.CompletionAccepted, AddressOf OnCompletionAccepted
    End Sub

    ''' <summary>
    ''' 初始化Word补全功能
    ''' </summary>
    Public Sub Initialize(wordApp As Microsoft.Office.Interop.Word.Application)
        _wordApp = wordApp

        ' 监听选区变化事件
        AddHandler _wordApp.WindowSelectionChange, AddressOf OnSelectionChange

        Debug.WriteLine("WordCompletionManager 已初始化")
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
                HideCompletionPopup()
            End If
        End Set
    End Property

    ''' <summary>
    ''' 选区变化事件处理
    ''' </summary>
    Private Sub OnSelectionChange(sel As Microsoft.Office.Interop.Word.Selection)
        Try
            Debug.WriteLine($"[WordCompletion] OnSelectionChange 触发, _isEnabled={_isEnabled}, ChatSettings.EnableAutocomplete={ChatSettings.EnableAutocomplete}")

            If Not _isEnabled OrElse Not ChatSettings.EnableAutocomplete Then
                Debug.WriteLine("[WordCompletion] 补全已禁用，跳过")
                Return
            End If

            ' 取消之前的定时器
            If _debounceTimer IsNot Nothing Then
                _debounceTimer.Dispose()
                _debounceTimer = Nothing
            End If

            ' 隐藏当前补全
            HideCompletionPopup()

            ' 获取当前段落文本
            If sel Is Nothing OrElse sel.Range Is Nothing Then
                Debug.WriteLine("[WordCompletion] 选区为空")
                Return
            End If

            Dim currentParagraph = sel.Range.Paragraphs(1)
            If currentParagraph Is Nothing Then Return

            Dim paragraphText = currentParagraph.Range.Text
            Debug.WriteLine($"[WordCompletion] 段落文本长度: {If(paragraphText, "").Length}")

            If String.IsNullOrWhiteSpace(paragraphText) OrElse paragraphText.Length < 5 Then
                Return
            End If

            ' 获取光标前的文本
            Dim cursorPos = sel.Range.Start
            Dim paraStart = currentParagraph.Range.Start
            Dim textBeforeCursor = ""

            If cursorPos > paraStart Then
                Dim beforeRange = _wordApp.ActiveDocument.Range(paraStart, cursorPos)
                textBeforeCursor = beforeRange.Text
            End If

            Debug.WriteLine($"[WordCompletion] 光标前文本: '{textBeforeCursor}'")

            If String.IsNullOrWhiteSpace(textBeforeCursor) OrElse textBeforeCursor.Length < 3 Then
                Return
            End If

            _lastParagraphText = textBeforeCursor
            Debug.WriteLine($"[WordCompletion] 准备请求补全，输入: '{textBeforeCursor}'")

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

            ' 获取屏幕上的光标位置
            Dim cursorScreenPos = GetWordCursorScreenPosition()
            Debug.WriteLine($"[WordCompletion] 光标位置: {cursorScreenPos}")

            ' 直接调用补全服务获取结果（不经过防抖）
            Dim completions = Await _completionService.GetCompletionsDirectAsync(inputText, "Word")

            Debug.WriteLine($"[WordCompletion] 获取到 {completions.Count} 个补全建议")

            ' 再次检查输入是否已变化
            If completions.Count > 0 AndAlso inputText = _lastParagraphText Then
                ShowCompletionPopup(completions, cursorScreenPos)
            End If

        Catch ex As Exception
            Debug.WriteLine($"RequestCompletion 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 获取Word中光标的屏幕位置
    ''' </summary>
    Private Function GetWordCursorScreenPosition() As Point
        Try
            Dim sel = _wordApp.Selection
            If sel IsNot Nothing Then
                Dim left, top, width, height As Integer
                _wordApp.ActiveWindow.GetPoint(left, top, width, height, sel.Range)
                Return New Point(left, top + height)
            End If
        Catch ex As Exception
            Debug.WriteLine($"GetWordCursorScreenPosition 出错: {ex.Message}")
        End Try
        Return New Point(100, 100)
    End Function

    ''' <summary>
    ''' 显示补全弹窗（使用UI同步上下文，避免死锁）
    ''' </summary>
    Private Sub ShowCompletionPopup(completions As List(Of String), position As Point)
        Try
            ' 使用Post而不是Invoke，避免阻塞和死锁
            _uiSyncContext.Post(Sub(state)
                                    Try
                                        ShowCompletionPopupInternal(completions, position)
                                    Catch ex As Exception
                                        Debug.WriteLine($"ShowCompletionPopupInternal 出错: {ex.Message}")
                                    End Try
                                End Sub, Nothing)
        Catch ex As Exception
            Debug.WriteLine($"ShowCompletionPopup 出错: {ex.Message}")
        End Try
    End Sub

    Private Sub ShowCompletionPopupInternal(completions As List(Of String), position As Point)
        _completionPopup.SetCompletions(completions)
        _completionPopup.Location = position
        _completionPopup.Show()
    End Sub

    ''' <summary>
    ''' 隐藏补全弹窗（使用UI同步上下文，避免死锁）
    ''' </summary>
    Public Sub HideCompletionPopup()
        Try
            _uiSyncContext.Post(Sub(state)
                                    Try
                                        _completionPopup.Hide()
                                    Catch ex As Exception
                                        Debug.WriteLine($"Hide popup 出错: {ex.Message}")
                                    End Try
                                End Sub, Nothing)
            _completionService.ClearCompletions()
        Catch ex As Exception
            Debug.WriteLine($"HideCompletionPopup 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 接受补全
    ''' </summary>
    Private Sub OnCompletionAccepted(completion As String)
        Try
            If _wordApp IsNot Nothing AndAlso _wordApp.Selection IsNot Nothing Then
                _wordApp.Selection.TypeText(completion)
            End If
            HideCompletionPopup()
        Catch ex As Exception
            Debug.WriteLine($"OnCompletionAccepted 出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 处理Ctrl+.快捷键
    ''' </summary>
    Public Sub AcceptFirstCompletion()
        Dim completions = _completionService.GetCurrentCompletions()
        If completions.Count > 0 Then
            OnCompletionAccepted(completions(0))
        End If
    End Sub

    ''' <summary>
    ''' 清理资源
    ''' </summary>
    Public Sub Dispose()
        If _debounceTimer IsNot Nothing Then
            _debounceTimer.Dispose()
        End If
        If _completionPopup IsNot Nothing Then
            _completionPopup.Dispose()
        End If
        If _wordApp IsNot Nothing Then
            RemoveHandler _wordApp.WindowSelectionChange, AddressOf OnSelectionChange
        End If
    End Sub
End Class

''' <summary>
''' 补全建议弹窗
''' </summary>
Public Class CompletionPopupForm
    Inherits Form

    Private _listBox As ListBox
    Private _completions As List(Of String)
    Private _hintLabel As Label

    Public Event CompletionAccepted(completion As String)

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.FormBorderStyle = FormBorderStyle.None
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.StartPosition = FormStartPosition.Manual
        Me.BackColor = Color.White
        Me.Size = New Size(350, 120)

        ' 列表框
        _listBox = New ListBox()
        _listBox.Dock = DockStyle.Fill
        _listBox.Font = New Font("Microsoft YaHei", 10)
        _listBox.BorderStyle = BorderStyle.FixedSingle
        AddHandler _listBox.DoubleClick, AddressOf ListBox_DoubleClick
        AddHandler _listBox.KeyDown, AddressOf ListBox_KeyDown
        Me.Controls.Add(_listBox)

        ' 提示标签
        _hintLabel = New Label()
        _hintLabel.Dock = DockStyle.Bottom
        _hintLabel.Height = 20
        _hintLabel.Font = New Font("Microsoft YaHei", 8)
        _hintLabel.ForeColor = Color.Gray
        _hintLabel.TextAlign = ContentAlignment.MiddleCenter
        _hintLabel.BackColor = Color.FromArgb(245, 245, 245)
        Me.Controls.Add(_hintLabel)

        UpdateHintText()
    End Sub

    ''' <summary>
    ''' 更新提示标签文本
    ''' </summary>
    Private Sub UpdateHintText()
        Dim shortcut = ChatSettings.AutocompleteShortcut
        If String.IsNullOrEmpty(shortcut) Then shortcut = "Ctrl+."
        _hintLabel.Text = $"按 {shortcut} 或双击接受 | Esc 关闭"
    End Sub

    Public Sub SetCompletions(completions As List(Of String))
        _completions = completions
        _listBox.Items.Clear()
        For Each c In completions
            _listBox.Items.Add(c)
        Next
        If _listBox.Items.Count > 0 Then
            _listBox.SelectedIndex = 0
        End If
        Me.Height = Math.Min(120, 25 * completions.Count + 25)
        UpdateHintText()
    End Sub

    Private Sub ListBox_DoubleClick(sender As Object, e As EventArgs)
        AcceptSelected()
    End Sub

    Private Sub ListBox_KeyDown(sender As Object, e As KeyEventArgs)
        ' 检查是否匹配配置的快捷键或 Enter 键
        If e.KeyCode = Keys.Enter OrElse MatchesConfiguredShortcut(e) Then
            e.Handled = True
            AcceptSelected()
        ElseIf e.KeyCode = Keys.Escape Then
            e.Handled = True
            Me.Hide()
        End If
    End Sub

    ''' <summary>
    ''' 检查按键是否匹配配置的快捷键
    ''' </summary>
    Private Function MatchesConfiguredShortcut(e As KeyEventArgs) As Boolean
        Dim shortcut = ChatSettings.AutocompleteShortcut
        If String.IsNullOrEmpty(shortcut) Then shortcut = "Ctrl+."

        Dim parts = shortcut.ToLower().Split("+"c)
        Dim requireCtrl = parts.Contains("ctrl")
        Dim requireAlt = parts.Contains("alt")
        Dim requireShift = parts.Contains("shift")

        ' 获取主键
        Dim mainKey = parts.LastOrDefault()
        If mainKey Is Nothing Then Return False

        ' 检查修饰键
        If requireCtrl <> e.Control Then Return False
        If requireAlt <> e.Alt Then Return False
        If requireShift <> e.Shift Then Return False

        ' 匹配主键
        Select Case mainKey
            Case "."
                Return e.KeyCode = Keys.OemPeriod
            Case "/"
                Return e.KeyCode = Keys.OemQuestion OrElse e.KeyCode = Keys.Divide
            Case "enter"
                Return e.KeyCode = Keys.Enter
            Case "space"
                Return e.KeyCode = Keys.Space
            Case "tab"
                Return e.KeyCode = Keys.Tab
            Case Else
                ' 尝试解析字母键
                If mainKey.Length = 1 AndAlso Char.IsLetter(mainKey(0)) Then
                    Dim expectedKey = CType(System.Enum.Parse(GetType(Keys), mainKey.ToUpper()), Keys)
                    Return e.KeyCode = expectedKey
                End If
                Return False
        End Select
    End Function

    Private Sub AcceptSelected()
        If _listBox.SelectedItem IsNot Nothing Then
            RaiseEvent CompletionAccepted(_listBox.SelectedItem.ToString())
        End If
    End Sub

    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
        End Get
    End Property
End Class
