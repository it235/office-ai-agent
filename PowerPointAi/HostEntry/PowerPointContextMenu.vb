' PowerPointAi\HostEntry\PowerPointContextMenu.vb
' E4-02：把 PowerPoint 的入口动作挂进右键菜单。这一层只回答两个宿主问题——
' "挂哪条菜栏"和"此刻有没有可处理的选中对象"；动作文案在 QuickActionCatalog，执行一律交回主链。
' 命名空间用 HostEntry：Selection 会和 Microsoft.Office.Interop 里的同名类型撞车。

Imports Office = Microsoft.Office.Core
Imports ShareRibbon
Imports ShareRibbon.Agent

Namespace HostEntry

    ''' <summary>
    ''' PowerPoint 右键菜单适配器。PPT 对各版本右键菜栏的暴露比 Word/Excel 零散，
    ''' 取不到菜栏时本入口直接不出（不报错、不重试），其余两端不受影响。
    ''' </summary>
    Public Class PowerPointContextMenu

        ' 按候选顺序试探：文本编辑态、形状右键、对象右键。真机验收需按版本记录实际命中哪一条。
        Private Shared ReadOnly BarCandidates As String() = {"Text", "Drawing", "Object"}

        Private ReadOnly _app As Microsoft.Office.Interop.PowerPoint.Application
        Private ReadOnly _menu As New HostContextMenu()

        Public Sub New(app As Microsoft.Office.Interop.PowerPoint.Application)
            _app = app
        End Sub

        Public Sub Attach(onRequest As Action(Of QuickAction))
            If _menu.Attached Then Return

            Dim bar = FindContextMenuBar()
            If bar Is Nothing Then
                AppLogger.Warn("PowerPointContextMenu", "context menu bar unavailable")
                Return
            End If

            _menu.Attach(bar, QuickActionCatalog.ActionsFor(QuickActionCatalog.PowerPointAppType),
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
                    ' 该版本没有暴露这条菜栏，属于预期探测路径
                End Try
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' 选中形状/文本框/缩略图里的幻灯片都能落到"当前对象/当前页"上；
        ''' ppSelectionNone（刚打开、焦点在功能区）时点下去只收到一句提示，避免 Agent 空跑一轮。
        ''' </summary>
        Private Function CurrentAvailability() As ContextMenuAvailability
            Dim nothingSelected = New ContextMenuAvailability With {
                .Enabled = False,
                .Reason = "先选中一个文本框、形状或幻灯片，再用这里的 AI 动作"
            }

            Dim window = _app.ActiveWindow
            If window Is Nothing Then Return nothingSelected

            Dim selection = window.Selection
            If selection Is Nothing Then Return nothingSelected

            Select Case selection.Type
                Case PowerPoint.PpSelectionType.ppSelectionShapes,
                     PowerPoint.PpSelectionType.ppSelectionText,
                     PowerPoint.PpSelectionType.ppSelectionSlides
                    Return New ContextMenuAvailability With {.Enabled = True}
                Case Else
                    Return nothingSelected
            End Select
        End Function
    End Class

End Namespace
