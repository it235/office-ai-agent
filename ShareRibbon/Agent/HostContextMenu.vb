' ShareRibbon\Agent\HostContextMenu.vb
' E4-02：宿主右键菜单入口的共享登记器。只用 OCOM 层的 Office.CommandBar，不碰任何
' Word/Excel/PowerPoint 具体类型——哪条菜栏、选区怎么探，由三个宿主各自的小适配器决定。
' 与 E4-01 同一个口径：入口只把一句话交回 BaseChatControl.StartAgentFromHostEntry，
' 执行、预览、风险确认、变更摘要、撤销都还在同一条 Agent 主链上。

Imports Office = Microsoft.Office.Core

Namespace Agent

    ''' <summary>点击那一刻的可用性判断。Reason 是"为什么现在不能用"，直接展示给用户。</summary>
    Public Class ContextMenuAvailability

        Public Property Enabled As Boolean = True

        Public Property Reason As String = ""
    End Class

    ''' <summary>
    ''' 往一条宿主上下文菜单里挂本加载项的动作项，并保证可撤销：所有控件都带 Tag 前缀，
    ''' Detach 按 Tag 删除，重复 Attach 先 Detach，因此不会出现"装了两次/卸载不干净"。
    '''
    ''' 关于降级：Office 15 PIA 的 CommandBar 根本没有事件面（BeforeRightClick 只存在于
    ''' 更早的类型库），所以没法在菜单弹出时把不合适的项置灰。这里改成点击时探测：
    ''' 菜单项始终可见可点，点了之后若当前上下文不满足，就把原因回给用户而不是硬跑一轮 Agent。
    ''' 静止的灰色比"点一下告诉你为什么不行"更容易骗人，所以不做置灰。
    ''' </summary>
    Public NotInheritable Class HostContextMenu

        ' Tag 是本加载项拥有这些控件的唯一依据，禁止改成 Caption 匹配（用户可改标题、语言也不同）
        Private Const TagPrefix As String = "aiq:"

        Private _bar As Office.CommandBar
        Private _actions As IReadOnlyList(Of QuickAction)
        Private _onRequest As Action(Of QuickAction)
        Private _availability As Func(Of ContextMenuAvailability)
        Private ReadOnly _buttons As New List(Of Office.CommandBarButton)
        Private _attached As Boolean

        Public ReadOnly Property Attached As Boolean
            Get
                Return _attached
            End Get
        End Property

        ''' <summary>挂菜单；宿主版本或 WPS 拿不到那条菜栏时返回 False，调用方不得重试成环。</summary>
        Public Function Attach(bar As Office.CommandBar,
                               actions As IReadOnlyList(Of QuickAction),
                               onRequest As Action(Of QuickAction),
                               availability As Func(Of ContextMenuAvailability)) As Boolean
            If bar Is Nothing OrElse actions Is Nothing OrElse actions.Count = 0 Then Return False
            If onRequest Is Nothing Then Return False
            If _attached Then Detach()

            Try
                ' 先认下这条菜栏：后面每一项都可能失败，失败时 Detach 必须还能把已加的项按 Tag 洗掉
                _bar = bar

                ' Office 的上下文菜单没有"分隔线控件"可加，分组靠第一项的 BeginGroup
                Dim first = True
                For Each action In actions
                    Dim button = CType(bar.Controls.Add(Type:=Office.MsoControlType.msoControlButton, Temporary:=True),
                                       Office.CommandBarButton)
                    button.Tag = TagPrefix & action.Id
                    button.Caption = action.Label
                    button.Style = Office.MsoButtonStyle.msoButtonCaption
                    button.TooltipText = action.RequestText
                    If first Then
                        first = False
                        Try
                            button.BeginGroup = True
                        Catch ex As Exception
                            ' 分组只是排版，宿主不支持时不能让整条菜单挂不上
                        End Try
                    End If
                    ' CommandBarButton 的事件挂点在 RCW 上，必须把按钮引用留在字段里，
                    ' 否则 GC 回收 RCW 后 Click 再也不会回调
                    AddHandler button.Click, AddressOf OnButtonClick
                    _buttons.Add(button)
                Next

                _actions = actions
                _onRequest = onRequest
                _availability = availability
                _attached = True
                Return True
            Catch ex As Exception
                ' 改宿主菜单属于"锦上添花"，任何失败都不能连累加载项启动或右键本身
                AppLogger.Warn("HostContextMenu", $"attach failed: {AppLogger.Redact(ex.Message)}")
                Detach()
                Return False
            End Try
        End Function

        Public Sub Detach()
            _attached = False

            For Each button In _buttons
                Try
                    RemoveHandler button.Click, AddressOf OnButtonClick
                Catch ex As Exception
                End Try
            Next
            _buttons.Clear()

            If _bar IsNot Nothing Then
                RemoveOwnControls()
            End If

            _bar = Nothing
            _actions = Nothing
            _onRequest = Nothing
            _availability = Nothing
        End Sub

        ''' <summary>倒着删：正序枚举 COM 集合时删除会让索引错位，留下洗不掉的残项。</summary>
        Private Sub RemoveOwnControls()
            Try
                For i = _bar.Controls.Count To 1 Step -1
                    Dim control = _bar.Controls.Item(i)
                    If control.Tag IsNot Nothing AndAlso control.Tag.StartsWith(TagPrefix, StringComparison.Ordinal) Then
                        control.Delete()
                    End If
                Next
            Catch ex As Exception
                AppLogger.Warn("HostContextMenu", $"detach skipped: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        ''' <summary>
        ''' 探测当前上下文能不能跑这个动作。探测本身失败时按"可用"处理：主链会重读上下文并
        ''' 给出合理回应，把入口锁死反而让右键变成噪声。
        ''' </summary>
        Private Function BlockedReason() As String
            Try
                If _availability Is Nothing Then Return Nothing
                Dim state = _availability.Invoke()
                If state Is Nothing OrElse state.Enabled Then Return Nothing
                Return If(String.IsNullOrWhiteSpace(state.Reason), "当前上下文还不适用于这个 AI 动作", state.Reason)
            Catch ex As Exception
                AppLogger.Warn("HostContextMenu", $"availability probe failed: {AppLogger.Redact(ex.Message)}")
                Return Nothing
            End Try
        End Function

        Private Sub OnButtonClick(control As Office.CommandBarButton, ByRef cancelDefault As Boolean)
            Dim action = FindAction(control.Tag)
            If action Is Nothing Then Return

            ' 只记 id：菜单项承载的原文与文档内容都不进日志
            AppLogger.Info("HostContextMenu", $"quick action id={action.Id}")

            Dim reason = BlockedReason()
            If reason IsNot Nothing Then
                ' 只写状态栏、不弹窗：用户此刻正站在右键菜单上，抢焦点的提示框比不响应更糟
                GlobalStatusStripAll.ShowProgress(reason)
                Return
            End If

            Try
                If _onRequest IsNot Nothing Then _onRequest.Invoke(action)
            Catch ex As Exception
                AppLogger.Warn("HostContextMenu", $"action failed: {AppLogger.Redact(ex.Message)}")
            End Try
        End Sub

        Private Function FindAction(tag As String) As QuickAction
            If _actions Is Nothing OrElse String.IsNullOrEmpty(tag) Then Return Nothing
            If Not tag.StartsWith(TagPrefix, StringComparison.Ordinal) Then Return Nothing

            Dim id = tag.Substring(TagPrefix.Length)
            For Each action In _actions
                If String.Equals(action.Id, id, StringComparison.Ordinal) Then Return action
            Next
            Return Nothing
        End Function
    End Class

End Namespace
