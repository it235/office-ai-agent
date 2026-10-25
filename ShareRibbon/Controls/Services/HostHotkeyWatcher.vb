' ShareRibbon\Controls\Services\HostHotkeyWatcher.vb
' E4-03 全局快捷键：三端共用一个轮询器，宿主只回答「按下了要做什么」。
' 这里不碰任何 Office 互操作类型，也不判断开关——开关在 ChatSettings，由宿主决定要不要 Start。

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

''' <summary>
''' 组合键描述：一个主键加一组修饰键。解析不出主键、或者一个修饰键都没有时返回 Nothing：
''' 单键会吃掉用户正常输入，这类键一律拒绝注册，而不是等运行时再和用户抢键盘。
''' </summary>
Public Class HostHotkeyChord

    Public Property MainKey As Integer
    Public Property Control As Boolean
    Public Property Alt As Boolean
    Public Property Shift As Boolean

    ''' <summary>必须至少带 Ctrl 或 Alt：中文输入法候选和宿主原生 accel 都在抢无修饰的键。</summary>
    Public ReadOnly Property IsSafe As Boolean
        Get
            Return MainKey <> 0 AndAlso (Control OrElse Alt)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Dim parts As New List(Of String)()
        If Control Then parts.Add("Ctrl")
        If Alt Then parts.Add("Alt")
        If Shift Then parts.Add("Shift")
        parts.Add(KeyName(MainKey))
        Return String.Join("+", parts)
    End Function

    ''' <summary>解析 "Ctrl+Alt+K" 这种写法；分隔符允许 + 与 -，大小写不敏感。</summary>
    Public Shared Function Parse(chord As String) As HostHotkeyChord
        If String.IsNullOrWhiteSpace(chord) Then Return Nothing
        Dim tokens = chord.Split("+"c, "-"c)
        Dim result As New HostHotkeyChord()
        Dim seen As Integer = 0
        For Each raw In tokens
            Dim token = raw.Trim().ToLowerInvariant()
            If token.Length = 0 Then Continue For
            seen += 1
            Select Case token
                Case "ctrl", "control"
                    result.Control = True
                Case "alt", "menu", "option"
                    result.Alt = True
                Case "shift"
                    result.Shift = True
                Case Else
                    ' 两个非修饰键（"Ctrl+Shift+A+B"）不是组合键，是写错了
                    If result.MainKey <> 0 Then Return Nothing
                    result.MainKey = KeyFromText(token)
                    If result.MainKey = 0 Then Return Nothing
            End Select
        Next
        If seen = 0 OrElse result.MainKey = 0 Then Return Nothing
        Return result
    End Function

    Private Shared Function KeyFromText(token As String) As Integer
        If token.Length = 1 Then
            Dim ch = token.Chars(0)
            If ch >= "a"c AndAlso ch <= "z"c Then Return AscW(UCase(ch))
            If ch >= "0"c AndAlso ch <= "9"c Then Return AscW(ch)
            Select Case ch
                Case "."c : Return CInt(Keys.OemPeriod)
                Case ","c : Return CInt(Keys.Oemcomma)
                Case "/"c : Return CInt(Keys.OemQuestion)
                Case ";"c : Return CInt(Keys.OemSemicolon)
                Case "'"c : Return CInt(Keys.OemQuotes)
                Case "["c : Return CInt(Keys.OemOpenBrackets)
                Case "]"c : Return CInt(Keys.OemCloseBrackets)
                Case "\"c : Return CInt(Keys.OemPipe)
                Case "-"c : Return CInt(Keys.OemMinus)
                Case "="c : Return CInt(Keys.Oemplus)
            End Select
            Return 0
        End If

        If token.StartsWith("f") AndAlso token.Length >= 2 Then
            Dim number As Integer
            If Integer.TryParse(token.Substring(1), number) AndAlso number >= 1 AndAlso number <= 12 Then
                Return CInt(Keys.F1) + number - 1
            End If
        End If

        Select Case token
            Case "enter", "return" : Return CInt(Keys.Enter)
            Case "space", "spacebar" : Return CInt(Keys.Space)
            Case "tab" : Return CInt(Keys.Tab)
            Case "esc", "escape" : Return CInt(Keys.Escape)
            Case "backspace", "bksp" : Return CInt(Keys.Back)
            Case "left" : Return CInt(Keys.Left)
            Case "right" : Return CInt(Keys.Right)
            Case "up" : Return CInt(Keys.Up)
            Case "down" : Return CInt(Keys.Down)
            Case "home" : Return CInt(Keys.Home)
            Case "end" : Return CInt(Keys.End)
            Case "pageup" : Return CInt(Keys.PageUp)
            Case "pagedown" : Return CInt(Keys.PageDown)
            Case "insert" : Return CInt(Keys.Insert)
            Case "delete", "del" : Return CInt(Keys.Delete)
            Case Else : Return 0
        End Select
    End Function

    Private Shared Function KeyName(vk As Integer) As String
        If vk >= 65 AndAlso vk <= 90 Then Return ChrW(vk).ToString()
        If vk >= 48 AndAlso vk <= 57 Then Return ChrW(vk).ToString()
        If vk >= CInt(Keys.F1) AndAlso vk <= CInt(Keys.F12) Then Return "F" & (vk - CInt(Keys.F1) + 1)
        Select Case vk
            Case CInt(Keys.Enter) : Return "Enter"
            Case CInt(Keys.Space) : Return "Space"
            Case CInt(Keys.Tab) : Return "Tab"
            Case CInt(Keys.Escape) : Return "Esc"
            Case CInt(Keys.Back) : Return "Backspace"
            Case CInt(Keys.Left) : Return "Left"
            Case CInt(Keys.Right) : Return "Right"
            Case CInt(Keys.Up) : Return "Up"
            Case CInt(Keys.Down) : Return "Down"
            Case CInt(Keys.Home) : Return "Home"
            Case CInt(Keys.End) : Return "End"
            Case CInt(Keys.PageUp) : Return "PageUp"
            Case CInt(Keys.PageDown) : Return "PageDown"
            Case CInt(Keys.Insert) : Return "Insert"
            Case CInt(Keys.Delete) : Return "Delete"
            Case CInt(Keys.Oemperiod) : Return "."
            Case CInt(Keys.Oemcomma) : Return ","
            Case CInt(Keys.OemQuestion) : Return "/"
            Case CInt(Keys.OemSemicolon) : Return ";"
            Case CInt(Keys.OemQuotes) : Return "'"
            Case CInt(Keys.OemOpenBrackets) : Return "["
            Case CInt(Keys.OemCloseBrackets) : Return "]"
            Case CInt(Keys.OemPipe) : Return "\"
            Case CInt(Keys.OemMinus) : Return "-"
            Case CInt(Keys.Oemplus) : Return "="
            Case Else : Return "?"
        End Select
    End Function

End Class

''' <summary>
''' 用 GetAsyncKeyState 轮询一个组合键的按下边沿。宿主没有可靠的按键事件面可用
''' （Word/PPT 只有选区变化事件，Excel 的 OnKey 会写进用户配置文件），所以沿用仓库里
''' 自动补全已经在用的轮询做法，并把三条风险控制住：只接受带修饰键的组合、
''' 只在宿主窗口是前台时触发、回调期间停表避免重入。
''' </summary>
Public Class HostHotkeyWatcher
    Implements IDisposable

    ''' <summary>只查 2~4 个虚拟键；这个间隔不会漏掉一次按下，也测不出 CPU 开销。</summary>
    Private Const PollIntervalMs As Integer = 80
    Private Const VK_SHIFT As Integer = 16
    Private Const VK_CONTROL As Integer = 17
    Private Const VK_MENU As Integer = 18
    Private Const KEY_DOWN_MASK As Integer = &H8000

    <DllImport("user32.dll")>
    Private Shared Function GetAsyncKeyState(virtualKey As Integer) As Short
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef processId As Integer) As Integer
    End Function

    Private ReadOnly _hostProcessId As Integer
    Private _timer As Timer
    Private _chord As HostHotkeyChord
    Private _onTrigger As Action
    Private _wasDown As Boolean
    Private _disposed As Boolean

    Public Sub New()
        _hostProcessId = Process.GetCurrentProcess().Id
    End Sub

    ''' <summary>已经在监听就返回 False：快捷键是独占资源，一个宿主只该有一个触发者。</summary>
    Public ReadOnly Property IsWatching As Boolean
        Get
            Return _timer IsNot Nothing
        End Get
    End Property

    ''' <summary>注册组合键；写法无效、没有修饰键或已有人监听都返回 False，调用方据此如实降级。</summary>
    Public Function Start(chord As String, onTrigger As Action) As Boolean
        If _disposed OrElse _timer IsNot Nothing OrElse onTrigger Is Nothing Then Return False
        Dim parsed = HostHotkeyChord.Parse(chord)
        If parsed Is Nothing OrElse Not parsed.IsSafe Then Return False
        _chord = parsed
        _onTrigger = onTrigger
        _wasDown = False
        _timer = New Timer With {.Interval = PollIntervalMs}
        AddHandler _timer.Tick, AddressOf OnPollTick
        _timer.Start()
        Return True
    End Function

    Public Sub StopWatching()
        If _timer Is Nothing Then Return
        _timer.Stop()
        RemoveHandler _timer.Tick, AddressOf OnPollTick
        _timer.Dispose()
        _timer = Nothing
        _onTrigger = Nothing
        _wasDown = False
    End Sub

    Private Sub OnPollTick(sender As Object, e As EventArgs)
        If _disposed OrElse _chord Is Nothing OrElse _onTrigger Is Nothing Then Return
        If Not IsKeyDown(_chord.MainKey) Then
            _wasDown = False
            Return
        End If
        ' 只认「从未按到按下」这一条边沿：长按和输入法候选期间的重复按下都不该再唤起一次
        If _wasDown Then Return
        _wasDown = True
        If Not ModifiersMatch(_chord) Then Return
        If Not IsHostForeground() Then Return

        ' 回调要开面板、走 COM，甚至可能进入模态；这段时间必须停表，
        ' 否则同一按键上会重入第二次 Tick，把「唤起」变成「反复唤起」。
        _timer.Stop()
        Try
            _onTrigger()
        Catch ex As Exception
            AppLogger.Debug("HostHotkeyWatcher", "快捷键回调失败", ex)
        End Try
        ' 回调里可能已经把监听关掉了（关掉设置、销毁控件），这时候必须尊重那个决定
        If Not _disposed AndAlso _timer IsNot Nothing AndAlso _onTrigger IsNot Nothing Then
            _timer.Start()
        End If
    End Sub

    Private Shared Function IsKeyDown(vk As Integer) As Boolean
        Return (CInt(GetAsyncKeyState(vk)) And KEY_DOWN_MASK) <> 0
    End Function

    ''' <summary>修饰键按精确集合比对：声明 Ctrl+Alt+K 时，再多按一个 Shift 不算数。</summary>
    Private Shared Function ModifiersMatch(chord As HostHotkeyChord) As Boolean
        Return IsKeyDown(VK_CONTROL) = chord.Control AndAlso
               IsKeyDown(VK_MENU) = chord.Alt AndAlso
               IsKeyDown(VK_SHIFT) = chord.Shift
    End Function

    ''' <summary>GetAsyncKeyState 是全局状态，所以必须自己判断焦点在不在宿主进程里，
    ''' 否则用户在别的程序按同一个组合键也会把 Office 面板拉起来。</summary>
    Private Function IsHostForeground() As Boolean
        Try
            Dim hwnd = GetForegroundWindow()
            If hwnd = IntPtr.Zero Then Return False
            Dim processId As Integer
            GetWindowThreadProcessId(hwnd, processId)
            Return processId = _hostProcessId
        Catch ex As Exception
            AppLogger.Debug("HostHotkeyWatcher", "前台窗口判断失败", ex)
            Return False
        End Try
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If _disposed Then Return
        _disposed = True
        StopWatching()
    End Sub

End Class
