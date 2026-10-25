' WordAi\Selection\WordSelectionToolbarController.vb
' E4-01：Word 选区工具条的编排者。只负责"什么时候出现、出现什么、点了交给谁"，
' 不读文档内容做判断、不执行任何写操作——执行交回 BaseChatControl.StartAgentFromHostEntry（整个动作，含入口 id），
' 走 OfficeHarness -> AgentKernel -> LoopEngine 这条唯一主链。

Imports System.Drawing
Imports System.Windows.Forms
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace Selection

    ''' <summary>
    ''' 选区工具条控制器。与 ghost text 的选区监听并存（各自订阅、互不调度），
    ''' 因此 Initialize/Dispose 必须幂等且对称，避免重复打开面板造成双触发。
    ''' </summary>
    Public Class WordSelectionToolbarController

        ' 验收要求"选中文本后 300ms 内出现"，留出余量给渲染与 COM 读取
        Private Const DebounceMs As Integer = 200
        Private Const MinSelectionChars As Integer = 4

        Private ReadOnly _app As Microsoft.Office.Interop.Word.Application
        Private _toolbar As WordSelectionToolbar
        Private _timer As Timer
        Private _bound As Boolean
        Private _shownStart As Integer = -1
        Private _shownEnd As Integer = -1

        Public Sub New(app As Microsoft.Office.Interop.Word.Application)
            _app = app
        End Sub

        Public Sub Initialize()
            If _bound Then Return

            _toolbar = New WordSelectionToolbar()
            AddHandler _toolbar.ActionChosen, AddressOf OnActionChosen

            _timer = New Timer()
            _timer.Interval = DebounceMs
            AddHandler _timer.Tick, AddressOf OnDebounceTick

            AddHandler _app.WindowSelectionChange, AddressOf OnWindowSelectionChange
            _bound = True
        End Sub

        ''' <summary>
        ''' 面板可见性等外部状态变化时收起浮层；不解绑事件，防止反复开合面板造成重复绑定。
        ''' </summary>
        Public Sub HideStrip()
            If Not _bound Then Return
            Try
                _timer.Stop()
                _toolbar.HideStrip()
                ForgetShownSelection()
            Catch ex As Exception
                AppLogger.Warn("WordSelectionToolbar", $"hide skipped: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Public Sub Dispose()
            If Not _bound Then Return
            _bound = False

            Try
                RemoveHandler _app.WindowSelectionChange, AddressOf OnWindowSelectionChange
                If _toolbar IsNot Nothing Then RemoveHandler _toolbar.ActionChosen, AddressOf OnActionChosen
            Catch ex As Exception
                AppLogger.Warn("WordSelectionToolbar", $"unbind failed: {AppLogger.Redact(ex.Message)}")
            End Try

            If _timer IsNot Nothing Then
                RemoveHandler _timer.Tick, AddressOf OnDebounceTick
                _timer.Dispose()
                _timer = Nothing
            End If

            If _toolbar IsNot Nothing Then
                _toolbar.Dispose()
                _toolbar = Nothing
            End If
        End Sub

        Private Sub OnWindowSelectionChange(selection As Microsoft.Office.Interop.Word.Selection)
            Try
                If Not CanShowToolbar() Then
                    HideStrip()
                    Return
                End If

                ' 连续拖动框选时只认最后一次，否则每个中间态都会重画一次浮层
                _timer.Stop()
                _timer.Start()
            Catch ex As Exception
                AppLogger.Warn("WordSelectionToolbar", $"selection handler failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub OnDebounceTick(sender As Object, e As EventArgs)
            _timer.Stop()
            Try
                ShowForCurrentSelection()
            Catch ex As Exception
                ' 浮层定位失败不该打扰写作，只留脱敏日志给真机验收查因
                AppLogger.Warn("WordSelectionToolbar", $"show skipped: {AppLogger.Redact(ex.Message)}")
                HideStrip()
            End Try
        End Sub

        Private Sub ShowForCurrentSelection()
            Dim actions = QuickActionCatalog.ActionsFor(QuickActionCatalog.WordAppType)
            If actions Is Nothing OrElse actions.Count = 0 Then
                HideStrip()
                Return
            End If

            Dim selection = _app.Selection
            Dim range = selection?.Range
            If range Is Nothing Then
                HideStrip()
                Return
            End If

            Dim start = range.Start
            Dim [end] = range.End
            Dim text = CleanSelectionText(selection.Text)
            If start = [end] OrElse text.Length < MinSelectionChars Then
                HideStrip()
                Return
            End If

            ' 同一处选区已经显示过就不再挪动位置，避免鼠标小幅抖动导致浮层跳来跳去
            If _toolbar.Visible AndAlso _shownStart = start AndAlso _shownEnd = [end] Then Return

            Dim anchor As Point
            WordSelectionToolbarGeometry.TryGetCursorPoint(anchor)

            Dim ownerRect As Rectangle
            WordSelectionToolbarGeometry.TryGetForegroundWindowRect(ownerRect)

            Dim stripSize = WordSelectionToolbarGeometry.ComputeStripSize(actions.Count)
            Dim bounds = WordSelectionToolbarGeometry.ResolveLocation(anchor, ownerRect, stripSize)
            _toolbar.ShowAt(actions, bounds)
            _shownStart = start
            _shownEnd = [end]
        End Sub

        Private Function CanShowToolbar() As Boolean
            If Not ChatSettings.EnableSelectionToolbar Then Return False

            ' WordAi 是 Application 级加载项：开始页、邮件正文、WPS 的其他窗口都能有活动窗口，
            ' 数文档个数会把"有文档但没有活动文档窗口"误判成可以出浮层，所以只认当前活动窗口的类型。
            Dim activeWindow = _app.ActiveWindow
            If activeWindow Is Nothing Then Return False
            If activeWindow.Type <> Microsoft.Office.Interop.Word.WdWindowType.wdWindowDocument Then Return False

            Dim chat = ThisAddIn.chatControl
            If chat Is Nothing OrElse Not chat.Visible Then Return False

            ' Agent 正在跑时会出现计划卡与风险审批，此时再叠一条浮层只会打断用户
            If chat.IsAgentRunActive Then Return False

            Return True
        End Function

        Private Sub OnActionChosen(action As QuickAction)
            Try
                HideStrip()

                ' 只记 id：选区原文不进日志，交给主链的时候才由上下文采集读取
                AppLogger.Info("WordSelectionToolbar", $"quick action id={action.Id}")

                ' 面板关闭时会一并收起浮层（ThisAddIn 的 VisibleChanged），所以这里能点就说明面板还在
                Dim chat = ThisAddIn.chatControl
                If chat Is Nothing Then Return
                chat.StartAgentFromHostEntry(action)
            Catch ex As Exception
                AppLogger.Warn("WordSelectionToolbar", $"action failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Sub ForgetShownSelection()
            _shownStart = -1
            _shownEnd = -1
        End Sub

        ''' <summary>
        ''' Word 的选区文本带段落标记与单元格结束符，只用来判断"值不值得出浮层"，不参与任何内容判断。
        ''' </summary>
        Private Shared Function CleanSelectionText(raw As String) As String
            If String.IsNullOrEmpty(raw) Then Return ""
            Return raw.Replace(vbCr, "").Replace(vbLf, "").Replace(ChrW(7), "").Trim()
        End Function
    End Class

End Namespace
