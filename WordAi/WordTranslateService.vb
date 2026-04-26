Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports ShareRibbon

Public Class WordTranslateService
    Inherits BaseTranslateService

    ' 右键菜单按钮引用，用于避免重复添加
    Private _translateButton As Microsoft.Office.Core.CommandBarButton
    Private _buttonAdded As Boolean = False

    Public Sub New()
        HookSelectionChange()
        HookRightClickMenu()
    End Sub

    Public Overrides Function GetSelectedText() As String
        Try
            Dim sel = Globals.ThisAddIn.Application.Selection
            If sel IsNot Nothing AndAlso sel.Type = WdSelectionType.wdSelectionNormal Then
                Return sel.Text.Trim()
            End If
        Catch
        End Try
        Return ""
    End Function

    Public Overrides Sub HookSelectionChange()
        AddHandler Globals.ThisAddIn.Application.WindowSelectionChange, Sub(doc)
                                                                            OnSelectionChanged()
                                                                        End Sub

    End Sub

    Public Overrides Sub HookRightClickMenu()
        ' 在 Word 原有右键菜单上追加翻译按钮，而不是覆盖原菜单
        Try
            If _buttonAdded Then Return

            ' 获取 Word 文本右键菜单
            Dim commandBar As Microsoft.Office.Core.CommandBar = Nothing
            Try
                commandBar = Globals.ThisAddIn.Application.CommandBars("Text")
            Catch
                ' 如果获取失败，尝试其他菜单名称
                commandBar = Globals.ThisAddIn.Application.CommandBars("Table Text")
            End Try

            If commandBar Is Nothing Then Return

            ' 添加分隔线和翻译按钮
            ' 先添加分隔线
            Dim separator As Microsoft.Office.Core.CommandBarControl = commandBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Before:=commandBar.Controls.Count + 1, Temporary:=True)
            separator.BeginGroup = True
            separator.Caption = "-"
            separator.Visible = True

            ' 添加翻译按钮
            _translateButton = DirectCast(commandBar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, Before:=commandBar.Controls.Count + 1, Temporary:=True), Microsoft.Office.Core.CommandBarButton)
            _translateButton.Caption = "翻译选中内容"
            _translateButton.FaceId = 0  ' 使用默认图标
            _translateButton.Visible = True
            _translateButton.Tag = "TranslateSelection"

            ' 绑定点击事件
            AddHandler _translateButton.Click, Sub(ctrl, ByRef cancelDefault)
                                                   OnRightClickTranslate()
                                               End Sub

            _buttonAdded = True
        Catch ex As Exception
            Debug.WriteLine($"HookRightClickMenu 失败: {ex.Message}")
        End Try
    End Sub

    ' 清理资源
    Public Sub Cleanup()
        Try
            If _translateButton IsNot Nothing Then
                _translateButton.Delete()
                _translateButton = Nothing
            End If
            _buttonAdded = False
        Catch
        End Try
    End Sub
End Class