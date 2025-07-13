Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports ShareRibbon

Public Class Spotlight
    ' 声明类级别变量，用于跟踪聚光灯功能状态
    Private _spotlightActive As Boolean = False
    Private WithEvents _appEvents As Excel.Application
    Private _currentWorkbook As Excel.Workbook

    ' 跟踪当前聚光灯位置
    Private _currentRow As Integer = 0
    Private _currentColumn As Integer = 0

    ' 条件格式名称
    Private Const ROW_FORMAT_NAME As String = "SpotlightRow"
    Private Const COLUMN_FORMAT_NAME As String = "SpotlightColumn"
    Private Const CELL_FORMAT_NAME As String = "SpotlightCell"

    ' 聚光灯配置 - 修改默认颜色为浅灰色
    Private _rowColor As Integer = RGB(230, 230, 230) ' 浅灰色
    Private _columnColor As Integer = RGB(230, 230, 230) ' 浅灰色
    Private _cellColor As Integer = RGB(200, 200, 200) ' 稍深的灰色，让活动单元格更明显
    Private _rowDisplay As Boolean = True ' 是否显示行高亮
    Private _columnDisplay As Boolean = True ' 是否显示列高亮

    ' 单例模式实现
    Private Shared _instance As Spotlight = Nothing

    ' 获取单例实例
    Public Shared Function GetInstance() As Spotlight
        If _instance Is Nothing Then
            _instance = New Spotlight()
        End If
        Return _instance
    End Function

    ' 私有构造函数，防止外部直接创建实例
    Private Sub New()
    End Sub

    ' 检查聚光灯是否激活
    Public ReadOnly Property IsActive As Boolean
        Get
            Return _spotlightActive
        End Get
    End Property

    ' 切换行显示
    Public Sub ToggleRowDisplay()
        _rowDisplay = Not _rowDisplay
        If _spotlightActive Then
            UpdateHighlight()
        End If
    End Sub

    ' 切换列显示
    Public Sub ToggleColumnDisplay()
        _columnDisplay = Not _columnDisplay
        If _spotlightActive Then
            UpdateHighlight()
        End If
    End Sub

    ' 设置聚光灯颜色
    Public Sub SetColors(rowColor As Integer, columnColor As Integer, cellColor As Integer)
        _rowColor = rowColor
        _columnColor = columnColor
        _cellColor = cellColor

        ' 如果聚光灯已激活，立即应用新颜色
        If _spotlightActive Then
            ApplyHighlight()
        End If
    End Sub

    ' 显示颜色选择对话框
    Public Sub ShowColorDialog()
        Try
            ' 创建颜色对话框
            Using colorDialog As New ColorDialog()
                ' 设置初始颜色
                colorDialog.Color = System.Drawing.ColorTranslator.FromOle(_rowColor)
                colorDialog.FullOpen = True ' 显示完整的颜色对话框
                colorDialog.CustomColors = New Integer() {
                    RGB(230, 230, 230), ' 浅灰色
                    RGB(255, 255, 150), ' 浅黄色
                    RGB(200, 255, 200), ' 浅绿色
                    RGB(200, 200, 255), ' 浅蓝色
                    RGB(255, 200, 200)  ' 浅红色
                }

                ' 显示对话框
                If colorDialog.ShowDialog() = DialogResult.OK Then
                    ' 用户选择了颜色，更新聚光灯颜色
                    Dim selectedColor As Integer = System.Drawing.ColorTranslator.ToOle(colorDialog.Color)

                    ' 行和列使用选择的颜色
                    _rowColor = selectedColor
                    _columnColor = selectedColor

                    ' 活动单元格使用稍深的颜色
                    Dim cellR As Integer = colorDialog.Color.R - 30
                    Dim cellG As Integer = colorDialog.Color.G - 30
                    Dim cellB As Integer = colorDialog.Color.B - 30

                    ' 确保RGB值不小于0
                    cellR = Math.Max(0, cellR)
                    cellG = Math.Max(0, cellG)
                    cellB = Math.Max(0, cellB)

                    _cellColor = RGB(cellR, cellG, cellB)

                    ' 如果聚光灯已激活，应用新颜色
                    If _spotlightActive Then
                        ApplyHighlight()
                    End If

                    GlobalStatusStripAll.ShowWarning("聚光灯颜色已更新")
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine("显示颜色对话框时出错: " & ex.Message)
            MessageBox.Show("显示颜色对话框时出错: " & ex.Message, "聚光灯颜色", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 切换聚光灯状态
    Public Function Toggle() As Boolean
        If _spotlightActive Then
            Deactivate()
            GlobalStatusStripAll.ShowWarning("聚光灯功能已关闭,双击聚光灯按钮可修改颜色")
        Else
            Activate()
            GlobalStatusStripAll.ShowWarning("聚光灯功能已开启，双击聚光灯按钮可修改颜色")
        End If

        Return _spotlightActive
    End Function

    ' 激活聚光灯功能
    Public Sub Activate()
        Try
            If _appEvents Is Nothing Then
                _appEvents = Globals.ThisAddIn.Application
            End If

            ' 保存对当前活动工作簿的引用
            _currentWorkbook = _appEvents.ActiveWorkbook

            ' 保存原始设置
            _spotlightActive = True

            ' 添加事件处理程序
            AddHandler _appEvents.SheetSelectionChange, AddressOf AppEvents_SheetSelectionChange
            AddHandler _appEvents.SheetActivate, AddressOf AppEvents_SheetActivate
            AddHandler _appEvents.SheetDeactivate, AddressOf AppEvents_SheetDeactivate
            AddHandler _appEvents.WorkbookActivate, AddressOf AppEvents_WorkbookActivate

            ' 应用高亮
            ApplyHighlight()

            ' 调试信息
            Debug.WriteLine("聚光灯功能已激活")
        Catch ex As Exception
            Debug.WriteLine("激活聚光灯功能时出错: " & ex.Message)
            MessageBox.Show("激活聚光灯功能时出错: " & ex.Message, "聚光灯功能", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' 取消激活聚光灯功能
    Public Sub Deactivate()
        Try
            If _appEvents IsNot Nothing Then
                ' 移除事件处理程序
                RemoveHandler _appEvents.SheetSelectionChange, AddressOf AppEvents_SheetSelectionChange
                RemoveHandler _appEvents.SheetActivate, AddressOf AppEvents_SheetActivate
                RemoveHandler _appEvents.SheetDeactivate, AddressOf AppEvents_SheetDeactivate
                RemoveHandler _appEvents.WorkbookActivate, AddressOf AppEvents_WorkbookActivate
            End If

            ' 移除高亮
            RemoveHighlight()

            ' 更新状态
            _spotlightActive = False
            _currentRow = 0
            _currentColumn = 0

            ' 调试信息
            Debug.WriteLine("聚光灯功能已停用")
        Catch ex As Exception
            Debug.WriteLine("取消激活聚光灯功能时出错: " & ex.Message)
        End Try
    End Sub

    ' 工作表选择更改事件处理程序
    Private Sub AppEvents_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)
        If _spotlightActive Then
            Debug.WriteLine("选择位置已更改")
            UpdateHighlight()

            ' 确保选中单元格仍然是选中状态
            Target.Select()

            ' 让Excel重新计算公式，确保单元格输入有效
            If Target.Count = 1 Then
                Try
                    Target.Calculate()
                Catch
                    ' 忽略计算错误
                End Try
            End If
        End If
    End Sub

    ' 工作表激活事件处理程序
    Private Sub AppEvents_SheetActivate(ByVal Sh As Object)
        If _spotlightActive Then
            Debug.WriteLine("工作表已激活")
            ApplyHighlight()
        End If
    End Sub

    ' 工作表停用事件处理程序
    Private Sub AppEvents_SheetDeactivate(ByVal Sh As Object)
        If _spotlightActive Then
            Debug.WriteLine("工作表已停用")
            RemoveHighlight()
        End If
    End Sub

    ' 工作簿激活事件处理程序
    Private Sub AppEvents_WorkbookActivate(ByVal Wb As Excel.Workbook)
        If _spotlightActive Then
            Debug.WriteLine("工作簿已激活")
            _currentWorkbook = Wb
            ApplyHighlight()
        End If
    End Sub

    ' 应用高亮
    Private Sub ApplyHighlight()
        Try
            ' 首先移除已有的高亮
            RemoveHighlight()

            ' 获取当前活动单元格
            Dim activeCell As Excel.Range = _appEvents.ActiveCell
            Dim activeSheet As Excel.Worksheet = _appEvents.ActiveSheet

            ' 保存当前位置
            _currentRow = activeCell.Row
            _currentColumn = activeCell.Column

            'Debug.WriteLine("当前位置: 行=" & _currentRow & ", 列=" & _currentColumn)

            ' 应用条件格式
            ApplyConditionalFormatting(activeCell, activeSheet)
        Catch ex As Exception
            Debug.WriteLine("应用高亮时出错: " & ex.Message)
        End Try
    End Sub

    ' 更新高亮
    Private Sub UpdateHighlight()
        Try
            ' 获取当前活动单元格
            Dim activeCell As Excel.Range = _appEvents.ActiveCell

            ' 如果位置变化了，重新应用高亮
            If _currentRow <> activeCell.Row OrElse _currentColumn <> activeCell.Column Then
                ApplyHighlight()
            End If
        Catch ex As Exception
            Debug.WriteLine("更新高亮时出错: " & ex.Message)
        End Try
    End Sub

    ' 移除高亮
    Private Sub RemoveHighlight()
        Try
            ' 获取当前活动工作表
            Dim activeSheet As Excel.Worksheet = _appEvents.ActiveSheet

            ' 删除条件格式
            activeSheet.Cells.FormatConditions.Delete()

            Debug.WriteLine("高亮已移除")
        Catch ex As Exception
            Debug.WriteLine("移除高亮时出错: " & ex.Message)
        End Try
    End Sub

    ' 应用条件格式
    Private Sub ApplyConditionalFormatting(activeCell As Excel.Range, activeSheet As Excel.Worksheet)
        Try
            ' 设置屏幕更新为False，提高性能
            _appEvents.ScreenUpdating = False

            ' 应用行高亮
            If _rowDisplay Then
                Dim entireRow As Excel.Range = activeSheet.Rows(_currentRow)
                Dim rowFormat As Excel.FormatCondition = entireRow.FormatConditions.Add(
                    Type:=XlFormatConditionType.xlExpression,
                    Formula1:="=ROW()=" & _currentRow)

                With rowFormat
                    .Interior.Color = _rowColor
                    .StopIfTrue = False
                End With
            End If

            ' 应用列高亮
            If _columnDisplay Then
                Dim entireColumn As Excel.Range = activeSheet.Columns(_currentColumn)
                Dim colFormat As Excel.FormatCondition = entireColumn.FormatConditions.Add(
                    Type:=XlFormatConditionType.xlExpression,
                    Formula1:="=COLUMN()=" & _currentColumn)

                With colFormat
                    .Interior.Color = _columnColor
                    .StopIfTrue = False
                End With
            End If

            ' 应用单元格高亮（会覆盖行列高亮）
            Dim cellFormat As Excel.FormatCondition = activeCell.FormatConditions.Add(
                Type:=XlFormatConditionType.xlExpression,
                Formula1:="=TRUE")

            With cellFormat
                .Interior.Color = _cellColor
                .StopIfTrue = False
            End With

            ' 恢复屏幕更新
            _appEvents.ScreenUpdating = True

            Debug.WriteLine("条件格式已应用")
        Catch ex As Exception
            _appEvents.ScreenUpdating = True
            Debug.WriteLine("应用条件格式时出错: " & ex.Message)
        End Try
    End Sub

    ' 添加一个清除所有高亮的公共方法
    Public Sub ClearAllHighlights()
        RemoveHighlight()
    End Sub
End Class