' WordAi\HostEntry\WordContextMenu.vb
' E4-02：把 Word 的入口动作挂进文本右键菜单。这一层只回答两个宿主问题——
' "挂哪条菜栏"和"此刻有没有可选区"；动作文案在 QuickActionCatalog，执行一律交回主链。
'
' 命名空间刻意叫 HostEntry 而不是 Selection：VB 的名字解析会把同名的命名空间和
' Microsoft.Office.Interop.* 里的类型混起来（PowerPoint 的 Selection 首当其冲），
' 这种冲突不是逻辑错，改都难查，所以宿主入口层一律用不会撞名的 HostEntry。

Imports Office = Microsoft.Office.Core
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace HostEntry

    ''' <summary>
    ''' Word 文本右键菜单适配器。与选区浮层（E4-01）共用同一份动作清单，
    ''' 区别只是"用户主动右键"和"选完自动出现"，因此两者的可用性口径必须一致。
    ''' </summary>
    Public Class WordContextMenu

        ' 同一条菜栏在不同 Word 版本/语言包里的名字不完全一致，按候选顺序取第一个能解析到的；
        ' WPS 走同一套 OCOM 接口。全都取不到就当作本宿主不提供这个入口，不重试、不弹错。
        Private Shared ReadOnly BarCandidates As String() = {"Text", "Text Context Menu"}

        Private ReadOnly _app As Microsoft.Office.Interop.Word.Application
        Private ReadOnly _menu As New HostContextMenu()

        Public Sub New(app As Microsoft.Office.Interop.Word.Application)
            _app = app
        End Sub

        ''' <summary>幂等挂载：反复开合面板只会挂一次，重复调用先卸再挂由共享层保证。</summary>
        Public Sub Attach(onRequest As Action(Of QuickAction))
            If _menu.Attached Then Return

            Dim bar = FindContextMenuBar()
            If bar Is Nothing Then
                AppLogger.Warn("WordContextMenu", "text context menu bar unavailable")
                Return
            End If

            _menu.Attach(bar, QuickActionCatalog.ActionsFor(QuickActionCatalog.WordAppType),
                         onRequest, AddressOf CurrentAvailability)
        End Sub

        Public Sub Detach()
            _menu.Detach()
        End Sub

        Private Function FindContextMenuBar() As Office.CommandBar
            For Each name In BarCandidates
                Try
                    Dim bar = _app.CommandBars(name)
                    If bar IsNot Nothing Then Return bar
                Catch ex As Exception
                    ' 名字不存在时 Word 直接抛异常，这是预期的探测路径，不记日志
                End Try
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Word 的 Selection.Type 比"文本长度"更能说明有没有真选区：光标停靠（wdSelectionIP）时
        ''' 任何"改写选区"类请求都无处落地，此时给出原因而不是让 Agent 空跑一轮。
        ''' </summary>
        Private Function CurrentAvailability() As ContextMenuAvailability
            Dim noSelection = New ContextMenuAvailability With {
                .Enabled = False,
                .Reason = "先选中一段文字，再用这里的 AI 动作"
            }

            Dim selection = _app.Selection
            If selection Is Nothing Then Return noSelection
            If selection.Type = Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionIP Then Return noSelection

            Return New ContextMenuAvailability With {.Enabled = True}
        End Function
    End Class

End Namespace
