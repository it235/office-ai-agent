' ExcelAi\HostEntry\ExcelContextMenu.vb
' E4-02：把 Excel 的入口动作挂进单元格右键菜单。这一层只回答两个宿主问题——
' "挂哪条菜栏"和"此刻右键点的是不是单元格区域"；动作文案在 QuickActionCatalog，执行一律交回主链。
' 命名空间用 HostEntry：Selection 会和 Microsoft.Office.Interop 里的同名类型撞车。

Imports Office = Microsoft.Office.Core
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace HostEntry

    ''' <summary>
    ''' Excel 单元格右键菜单适配器。与 Word 侧同一条规则：入口只说话，
    ''' 点下去之后仍然只有 `OfficeHarness -> AgentKernel -> LoopEngine` 这一条链。
    ''' </summary>
    Public Class ExcelContextMenu

        ' "Cell" 是 Excel 单元格右键那条 popup；老版本/中文包另有别名，按候选顺序试探。
        Private Shared ReadOnly BarCandidates As String() = {"Cell", "Range Context Menu"}

        Private ReadOnly _app As Microsoft.Office.Interop.Excel.Application
        Private ReadOnly _menu As New HostContextMenu()

        Public Sub New(app As Microsoft.Office.Interop.Excel.Application)
            _app = app
        End Sub

        Public Sub Attach(onRequest As Action(Of QuickAction))
            If _menu.Attached Then Return

            Dim bar = FindContextMenuBar()
            If bar Is Nothing Then
                AppLogger.Warn("ExcelContextMenu", "cell context menu bar unavailable")
                Return
            End If

            _menu.Attach(bar, QuickActionCatalog.ActionsFor(QuickActionCatalog.ExcelAppType),
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
                    ' 菜栏名在该版本里不存在，属于预期探测路径
                End Try
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' Excel 的"选区"只有落在 Range 上三个动作才成立：右键图表、形状或编辑对象内部时
        ''' Selection 不是 Range，此时点下去只会收到一句"请先选中单元格"，而不是让 Agent 空跑一轮。
        ''' </summary>
        Private Function CurrentAvailability() As ContextMenuAvailability
            Dim notRange = New ContextMenuAvailability With {
                .Enabled = False,
                .Reason = "这些 AI 动作作用于单元格区域，请先选中单元格"
            }

            Dim selection = _app.Selection
            If selection Is Nothing Then Return notRange

            Dim range = TryCast(selection, Microsoft.Office.Interop.Excel.Range)
            If range Is Nothing Then Return notRange

            Return New ContextMenuAvailability With {.Enabled = True}
        End Function
    End Class

End Namespace
