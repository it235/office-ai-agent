Imports Microsoft.Office.Interop.PowerPoint
Imports ShareRibbon
Imports ShareRibbon.Controls

Public Class PowerPointTranslateService
    Inherits BaseTranslateService


    Public Sub New()
        HookSelectionChange()
        HookRightClickMenu()
    End Sub

    Public Overrides Function GetSelectedText() As String
        Try
            Dim sel = Globals.ThisAddIn.Application.ActiveWindow.Selection
            If sel.Type = PpSelectionType.ppSelectionText Then
                Return sel.TextRange.Text
            End If
        Catch
        End Try
        Return ""
    End Function

    Public Overrides Sub HookSelectionChange()
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, Sub(win)
                                                                            OnSelectionChanged()
                                                                        End Sub
    End Sub

    Public Overrides Sub HookRightClickMenu()
        ' PowerPoint 没有直接的右键事件，只能通过自定义 Ribbon 或定时检测
        ' 这里建议用 Ribbon 按钮或定时检测选区
    End Sub
End Class