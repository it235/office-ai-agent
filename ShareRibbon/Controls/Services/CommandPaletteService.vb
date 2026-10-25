' ShareRibbon\Controls\Services\CommandPaletteService.vb
' E4-03 全局快捷键 + 命令面板：一个组合键在任一宿主里唤起面板，面板里既能动词式快捷动作，也能直接写自然语言。
' 这个服务只管「唤起」和「把动作清单交给前端」：真正执行仍然走前端 startAgent -> HandleStartAgent 那一条主链，
' 与面板输入同一个模型配置闸门、同一次上下文采集、同一套预览与风险确认。

Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon.Agent

''' <summary>
''' 命令面板的后端侧：快捷键监听 + 动作清单下发。
''' </summary>
Public Class CommandPaletteService

    Private ReadOnly _executeScript As Func(Of String, Task)
    Private ReadOnly _getAppType As Func(Of String)
    Private ReadOnly _ensurePaneVisible As Action
    Private _watcher As HostHotkeyWatcher
    Private _pageReady As Boolean
    Private _openPending As Boolean

    ''' <param name="ensurePaneVisible">快捷键很可能在面板收起时按下：面板没显示，动作清单就没有落点。</param>
    Public Sub New(executeScript As Func(Of String, Task), getAppType As Func(Of String), ensurePaneVisible As Action)
        _executeScript = executeScript
        _getAppType = getAppType
        _ensurePaneVisible = ensurePaneVisible
    End Sub

    Public ReadOnly Property IsHotkeyActive As Boolean
        Get
            Return _watcher IsNot Nothing AndAlso _watcher.IsWatching
        End Get
    End Property

    ''' <summary>
    ''' HTML 就绪后由 InitializeSettings 调用：先同步一次清单，再补上此前攒下的唤起请求。
    ''' </summary>
    Public Sub PushState()
        _pageReady = True
        Send("syncCommandPalette", BuildState())
        If _openPending Then
            _openPending = False
            Send("showCommandPalette", "{}")
        End If
    End Sub

    ''' <summary>唤起面板：宿主快捷键和面板按钮共用这一个入口，两条路的表现必须一致。</summary>
    Public Sub Open()
        Try
            If _ensurePaneVisible IsNot Nothing Then _ensurePaneVisible()
        Catch ex As Exception
            AppLogger.Debug("CommandPaletteService", "展开任务窗格失败", ex)
        End Try

        If _pageReady Then
            Send("showCommandPalette", "{}")
        Else
            ' 第一次唤起时 WebView 还在导航，等 PushState 再补，免得在这里排队重试
            _openPending = True
        End If
    End Sub

    ''' <summary>
    ''' 按当前设置启停监听。设置面板里改完就调一次，不要求用户重启宿主。
    ''' </summary>
    Public Sub ApplyHotkey()
        Try
            If _watcher Is Nothing Then _watcher = New HostHotkeyWatcher()

            If Not ChatSettings.EnableHotkey Then
                _watcher.StopWatching()
                Return
            End If

            ' 组合键可能刚被改过：先停再注册，否则旧的键会继续吃掉这一次唤起
            _watcher.StopWatching()
            If Not _watcher.Start(ChatSettings.HotkeyChord, AddressOf OnHotkeyTriggered) Then
                _watcher.StopWatching()
                GlobalStatusStrip.ShowWarning(
                    $"命令面板快捷键 {ChatSettings.HotkeyChord} 无法注册：请用带 Ctrl 或 Alt 的组合键。")
            End If
        Catch ex As Exception
            AppLogger.Debug("CommandPaletteService", "注册快捷键失败", ex)
        End Try
    End Sub

    ''' <summary>面板销毁时必须停表：轮询器活着就等于快捷键在用户以为已经关掉插件之后仍然生效。</summary>
    Public Sub Shutdown()
        Try
            If _watcher IsNot Nothing Then
                _watcher.Dispose()
                _watcher = Nothing
            End If
        Catch ex As Exception
            AppLogger.Debug("CommandPaletteService", "释放快捷键监听失败", ex)
        End Try
    End Sub

    Private Sub OnHotkeyTriggered()
        Open()
    End Sub

    ''' <summary>
    ''' 动作清单只有 QuickActionCatalog 一份来源：命令面板、选区工具条、右键菜单、引导示例共用，
    ''' 这里不复制话术，前端也就不会出现一份和后端对不上的按钮表。
    ''' E6-04 起再叠一层任务包归组：先按包登记的入口顺序给出，未归组的动作仍然照旧下发并标空 packageId，
    ''' 前端据此分组显示"审阅润色"这类工作目标，而不是只看到一串动词。
    ''' </summary>
    Private Function BuildState() As String
        Dim appType = If(_getAppType(), "")
        Dim state As New JObject()
        state("appType") = appType
        state("enabled") = ChatSettings.EnableHotkey
        state("shortcut") = ChatSettings.HotkeyChord

        Dim catalog As New List(Of QuickAction)(QuickActionCatalog.ActionsFor(appType))
        Dim actions As New JArray()
        Dim packageHeaders As New JArray()
        Dim emitted As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each package In TaskPackageCatalog.PackagesFor(appType)
            Dim header As New JObject()
            header("id") = package.Id
            header("title") = package.Title
            header("deliverable") = package.Deliverable
            packageHeaders.Add(header)
            For Each actionId In package.EntryActions
                Dim matched = catalog.FirstOrDefault(Function(item) String.Equals(item.Id, actionId, StringComparison.OrdinalIgnoreCase))
                If matched Is Nothing Then Continue For
                actions.Add(BuildActionItem(matched, package.Id))
                emitted.Add(matched.Id)
            Next
        Next

        For Each action In catalog
            If emitted.Contains(action.Id) Then Continue For
            actions.Add(BuildActionItem(action, Nothing))
        Next

        state("packages") = packageHeaders
        state("actions") = actions

        Return state.ToString(Formatting.None)
    End Function

    Private Shared Function BuildActionItem(action As QuickAction, packageId As String) As JObject
        Dim item As New JObject()
        item("id") = action.Id
        item("label") = action.Label
        item("request") = action.RequestText
        If Not String.IsNullOrWhiteSpace(packageId) Then item("packageId") = packageId
        Return item
    End Function

    Private Async Function SendAsync(jsFunction As String, payloadJson As String) As Task
        If _executeScript Is Nothing Then Return
        Try
            Await _executeScript($"{jsFunction}({payloadJson});")
        Catch ex As Exception
            AppLogger.Debug("CommandPaletteService", $"推送 {jsFunction} 失败", ex)
        End Try
    End Function

    ''' <summary>同步调用方只需要「最终送达」，不等 WebView 回执行结果。</summary>
    Private Async Sub Send(jsFunction As String, payloadJson As String)
        Await SendAsync(jsFunction, payloadJson)
    End Sub

End Class
