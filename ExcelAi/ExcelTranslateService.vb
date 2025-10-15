Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports ShareRibbon
Imports ShareRibbon.Controls

Public Class ExcelTranslateService
    Inherits BaseTranslateService


    Public Sub New()
        HookSelectionChange()
        HookRightClickMenu()
    End Sub

    Public Overrides Function GetSelectedText() As String
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            If TypeOf sel Is Range Then
                Dim rng = CType(sel, Range)
                Return rng.Text
            End If
        Catch
        End Try
        Return ""
    End Function

    Public Overrides Sub HookSelectionChange()
        AddHandler Globals.ThisAddIn.Application.SheetSelectionChange, Sub(sh, rng)
                                                                           OnSelectionChanged()
                                                                       End Sub
    End Sub

    Public Overrides Sub HookRightClickMenu()
        AddHandler Globals.ThisAddIn.Application.SheetBeforeRightClick, Sub(sh, rng, ByRef Cancel)
                                                                            Dim txt = GetSelectedText()
                                                                            If Not String.IsNullOrEmpty(txt) Then
                                                                                Dim menu As New ContextMenuStrip()
                                                                                Dim item As New ToolStripMenuItem("翻译选中内容")
                                                                                AddHandler item.Click, Sub(s, e) OnRightClickTranslate()
                                                                                menu.Items.Add(item)
                                                                                menu.Show(Cursor.Position)
                                                                                Cancel = True
                                                                            End If
                                                                        End Sub
    End Sub
End Class