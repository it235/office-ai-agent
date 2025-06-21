Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop
Imports ShareRibbon
Imports Button = System.Windows.Forms.Button
Imports Font = System.Drawing.Font
Imports Point = System.Drawing.Point
Imports ScrollBars = System.Windows.Forms.ScrollBars
Imports TextBox = System.Windows.Forms.TextBox

Public Class EnhancedPreviewAndConfirm
    ' 用于保存工作表状态信息的类
    Private Class WorksheetState
        Public Name As String
        Public Cells As Dictionary(Of String, Object)
        Public UsedRangeAddress As String
        Public SheetExists As Boolean

        Public Sub New(name As String)
            Me.Name = name
            Me.Cells = New Dictionary(Of String, Object)
            Me.SheetExists = True
        End Sub
    End Class

    ' 用于表示单元格差异的类
    Private Class CellDifference
        Public Address As String
        Public SheetName As String
        Public OldValue As Object
        Public NewValue As Object
        Public ChangeType As String ' "添加", "修改", "删除"

        Public Sub New(sheetName As String, address As String, oldValue As Object, newValue As Object, changeType As String)
            Me.SheetName = sheetName
            Me.Address = address
            Me.OldValue = oldValue
            Me.NewValue = newValue
            Me.ChangeType = changeType
        End Sub
    End Class

    ' 用于表示工作表差异的类
    Private Class SheetDifference
        Public SheetName As String
        Public ChangeType As String ' "添加", "删除", "修改"

        Public Sub New(sheetName As String, changeType As String)
            Me.SheetName = sheetName
            Me.ChangeType = changeType
        End Sub
    End Class


    ' 使用异步方式处理，避免界面卡死
    Public Async Function PreviewAndConfirmVbaExecutionAsync(vbaCode As String) As Task(Of Boolean)
        Dim application As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
        Dim originalWorkbook As Workbook = application.ActiveWorkbook

        If originalWorkbook Is Nothing Then
            MessageBox.Show("没有打开的工作簿，无法预览变更。", "预览错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        ' 步骤1: 捕获当前工作簿状态
        Dim beforeState = Await Task.Run(Function() CaptureWorkbookState(originalWorkbook))

        ' 步骤2: 创建临时工作簿副本来执行代码
        Dim tempWorkbookPath As String = Nothing
        Dim tempWorkbook As Workbook = Nothing
        Dim tempFileName As String = IO.Path.GetTempFileName()

        Try
            ' 使用SaveCopyAs代替SaveAs，这样不会改变原始工作簿的路径
            tempWorkbookPath = IO.Path.ChangeExtension(tempFileName, ".xlsx")
            application.DisplayAlerts = False
            originalWorkbook.SaveCopyAs(tempWorkbookPath)
            application.DisplayAlerts = True

            ' 打开刚刚创建的副本
            tempWorkbook = application.Workbooks.Open(tempWorkbookPath)
            tempWorkbook.Activate() ' 确保操作在临时工作簿上执行

            ' 异步执行VBA
            Dim executionResult = Await Task.Run(Function() ExecuteCodeInTemporaryModule(tempWorkbook, vbaCode))
            If Not executionResult Then Return False

            ' 步骤4: 捕获执行后的状态
            Dim afterState = Await Task.Run(Function() CaptureWorkbookState(tempWorkbook))

            ' 步骤5: 比较状态
            Dim cellDifferences As New List(Of CellDifference)()
            Dim sheetDifferences As New List(Of SheetDifference)()
            CompareWorkbookStates(beforeState, afterState, cellDifferences, sheetDifferences)

            ' 步骤6: 显示优化后的预览弹窗
            Dim userConfirmed = ShowDifferencePreview(vbaCode, cellDifferences, sheetDifferences)
            Return userConfirmed

        Catch ex As Exception
            MessageBox.Show("预览代码执行时出错: " & ex.Message, "预览错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            application.DisplayAlerts = False
            ' 关闭临时工作簿
            If tempWorkbook IsNot Nothing Then
                Try
                    tempWorkbook.Close(SaveChanges:=False)
                    Marshal.ReleaseComObject(tempWorkbook)
                Catch
                End Try
            End If

            ' 重新激活原始工作簿
            If originalWorkbook IsNot Nothing Then
                Try
                    originalWorkbook.Activate()
                Catch
                    ' 忽略激活错误
                End Try
            End If
            application.DisplayAlerts = True

            ' 删除临时文件
            Try
                If tempWorkbookPath IsNot Nothing AndAlso IO.File.Exists(tempWorkbookPath) Then
                    IO.File.Delete(tempWorkbookPath)
                End If
                If IO.File.Exists(tempFileName) Then
                    IO.File.Delete(tempFileName)
                End If
            Catch
                ' 忽略删除临时文件的错误
            End Try
        End Try
    End Function

    ' 将原有同步方法改为调用异步方法，避免卡UI
    Public Function PreviewAndConfirmVbaExecution(vbaCode As String) As Boolean
        Return PreviewAndConfirmVbaExecutionAsync(vbaCode).GetAwaiter().GetResult()
    End Function
    ' 优化预览弹窗布局
    Private Function ShowDifferencePreview(code As String,
                                          cellDifferences As List(Of CellDifference),
                                          sheetDifferences As List(Of SheetDifference)) As Boolean
        Dim previewForm As New Form() With {
            .Text = "VBA代码执行预览",
            .Size = New Size(950, 650),
            .StartPosition = FormStartPosition.CenterScreen,
            .MinimizeBox = False,
            .MaximizeBox = True,
            .FormBorderStyle = FormBorderStyle.Sizable
        }

        ' 主TabControl容器
        Dim tabControl As New TabControl() With {
            .Dock = DockStyle.Fill
        }

        ' 代码区
        Dim codeTab As New TabPage("VBA代码")
        Dim codeTextBox As New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Both,
            .Text = code,
            .Font = New Font("Consolas", 10),
            .Dock = DockStyle.Fill,
            .WordWrap = False
        }
        codeTab.Controls.Add(codeTextBox)

        ' 工作表变更
        Dim sheetTab As New TabPage("工作表变更")
        Dim sheetListView As New ListView() With {
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .Dock = DockStyle.Fill
        }
        sheetListView.Columns.Add("工作表名称", 150)
        sheetListView.Columns.Add("变更类型", 100)

        For Each diff In sheetDifferences
            Dim item As New ListViewItem(diff.SheetName)
            item.SubItems.Add(diff.ChangeType)
            Select Case diff.ChangeType
                Case "添加"
                    item.BackColor = Color.LightGreen
                Case "删除"
                    item.BackColor = Color.LightPink
            End Select
            sheetListView.Items.Add(item)
        Next
        If sheetDifferences.Count = 0 Then
            sheetListView.Items.Add(New ListViewItem("无工作表变更"))
        End If
        sheetTab.Controls.Add(sheetListView)

        ' 单元格变更
        Dim cellTab As New TabPage("单元格变更")
        Dim cellListView As New ListView() With {
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .Dock = DockStyle.Fill
        }
        cellListView.Columns.Add("工作表", 80)
        cellListView.Columns.Add("单元格", 80)
        cellListView.Columns.Add("变更类型", 80)
        cellListView.Columns.Add("原值", 150)
        cellListView.Columns.Add("新值", 150)

        For Each diff In cellDifferences
            Dim item As New ListViewItem(diff.SheetName)
            item.SubItems.Add(diff.Address)
            item.SubItems.Add(diff.ChangeType)
            item.SubItems.Add(If(diff.OldValue Is Nothing, "(空)", diff.OldValue.ToString()))
            item.SubItems.Add(If(diff.NewValue Is Nothing, "(空)", diff.NewValue.ToString()))
            Select Case diff.ChangeType
                Case "添加"
                    item.BackColor = Color.LightGreen
                Case "删除"
                    item.BackColor = Color.LightPink
                Case "修改"
                    item.BackColor = Color.LightYellow
            End Select
            cellListView.Items.Add(item)
        Next
        If cellDifferences.Count = 0 Then
            cellListView.Items.Add(New ListViewItem("无单元格变更"))
        End If
        cellTab.Controls.Add(cellListView)

        ' 摘要
        Dim summaryTab As New TabPage("变更摘要")
        Dim summaryTextBox As New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Fill,
            .Font = New Font("微软雅黑", 10)
        }
        summaryTextBox.Text = GenerateSummary(sheetDifferences, cellDifferences)
        summaryTab.Controls.Add(summaryTextBox)

        tabControl.TabPages.Add(summaryTab)
        tabControl.TabPages.Add(cellTab)
        tabControl.TabPages.Add(sheetTab)
        tabControl.TabPages.Add(codeTab)

        Dim buttonPanel As New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 50
        }

        ' 给 buttonPanel 加一个 FlowLayoutPanel，简化按钮布局
        Dim flowLayout As New FlowLayoutPanel() With {
            .FlowDirection = FlowDirection.RightToLeft,
            .Dock = DockStyle.Fill
        }
        buttonPanel.Controls.Add(flowLayout)

        Dim acceptButton As New Button() With {
            .Text = "应用变更",
            .DialogResult = DialogResult.Yes,
            .AutoSize = True
        }
        Dim cancelButton As New Button() With {
            .Text = "取消",
            .DialogResult = DialogResult.No,
            .AutoSize = True
        }

        ' 流式布局下从右至左添加
        flowLayout.Controls.Add(cancelButton)
        flowLayout.Controls.Add(acceptButton)

        ' 依次将 panel、tabControl 放到 form
        previewForm.Controls.Add(buttonPanel)
        previewForm.Controls.Add(tabControl)


        ' 自动定位
        'acceptButton.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        'cancelButton.Anchor = AnchorStyles.Right Or AnchorStyles.Top
        'acceptButton.Location = New Point(previewForm.ClientSize.Width - 240, 10)
        'cancelButton.Location = New Point(previewForm.ClientSize.Width - 120, 10)

        'buttonPanel.Controls.Add(acceptButton)
        'buttonPanel.Controls.Add(cancelButton)
        'previewForm.Controls.Add(buttonPanel)
        'previewForm.Controls.Add(tabControl)

        ' 默认确认按钮
        previewForm.AcceptButton = cancelButton
        ' 显示对话框后，若用户点“应用变更”则返回 True
        Return (previewForm.ShowDialog() = DialogResult.Yes)
    End Function

    Private Function GenerateSummary(sheetDiffs As List(Of SheetDifference),
                                    cellDiffs As List(Of CellDifference)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("# 变更摘要")
        sb.AppendLine()

        If sheetDiffs.Count > 0 Then
            sb.AppendLine("## 工作表变更")
            For Each diff In sheetDiffs
                sb.AppendLine($"- {diff.SheetName}: {diff.ChangeType}")
            Next
            sb.AppendLine()
        End If

        If cellDiffs.Count > 0 Then
            Dim grouped = cellDiffs.GroupBy(Function(d) d.SheetName)
            sb.AppendLine("## 单元格变更")
            For Each group In grouped
                sb.AppendLine($"### 工作表: {group.Key}")
                Dim addCount = group.Count(Function(d) d.ChangeType = "添加")
                Dim modifyCount = group.Count(Function(d) d.ChangeType = "修改")
                Dim deleteCount = group.Count(Function(d) d.ChangeType = "删除")

                If addCount > 0 Then
                    sb.AppendLine($"- 添加: {addCount} 个单元格")
                End If
                If modifyCount > 0 Then
                    sb.AppendLine($"- 修改: {modifyCount} 个单元格")
                End If
                If deleteCount > 0 Then
                    sb.AppendLine($"- 删除: {deleteCount} 个单元格")
                End If
                sb.AppendLine()
            Next
        End If

        If sheetDiffs.Count = 0 AndAlso cellDiffs.Count = 0 Then
            sb.AppendLine("此代码执行后没有发现数据变更.")
        End If
        Return sb.ToString()
    End Function


    ' 查找模块中的第一个过程名
    Private Function FindFirstProcedureName(comp As VBComponent) As String
        Try
            Dim codeModule As CodeModule = comp.CodeModule
            Dim lineCount As Integer = codeModule.CountOfLines
            Dim line As Integer = 1

            While line <= lineCount
                Dim procName As String = codeModule.ProcOfLine(line, vbext_ProcKind.vbext_pk_Proc)
                If Not String.IsNullOrEmpty(procName) Then
                    Return procName
                End If
                line = codeModule.ProcStartLine(procName, vbext_ProcKind.vbext_pk_Proc) + codeModule.ProcCountLines(procName, vbext_ProcKind.vbext_pk_Proc)
            End While

            Return String.Empty
        Catch
            ' 如果出错，尝试使用正则表达式从代码中提取
            Dim code As String = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            Dim match As Match = Regex.Match(code, "^\s*(Sub|Function)\s+(\w+)", RegexOptions.Multiline Or RegexOptions.IgnoreCase)

            If match.Success AndAlso match.Groups.Count > 2 Then
                Return match.Groups(2).Value
            End If

            Return String.Empty
        End Try
    End Function


    ' 检查代码是否包含过程声明
    Private Function ContainsProcedureDeclaration(code As String) As Boolean
        ' 使用简单的正则表达式检查是否包含 Sub 或 Function 声明
        Return Regex.IsMatch(code, "^\s*(Sub|Function)\s+\w+", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
    End Function

    ' 执行前端传来的 VBA 代码片段
    Private Function ExecuteCodeInTemporaryModule(workbook As Workbook, vbaCode As String)
        ' 获取 VBA 项目
        Dim vbProj As VBProject = workbook.VBProject

        ' 添加空值检查
        If vbProj Is Nothing Then
            Return Nothing
        End If

        Dim vbComp As VBComponent = Nothing
        Dim tempModuleName As String = "TempPreviewMod" & DateTime.Now.Ticks.ToString().Substring(0, 8)

        Try
            ' 创建临时模块
            vbComp = vbProj.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            vbComp.Name = tempModuleName

            ' 检查代码是否已包含 Sub/Function 声明
            If ContainsProcedureDeclaration(vbaCode) Then
                ' 代码已包含过程声明，直接添加
                vbComp.CodeModule.AddFromString(vbaCode)

                ' 查找第一个过程名并执行
                Dim procName As String = FindFirstProcedureName(vbComp)
                If Not String.IsNullOrEmpty(procName) Then
                    workbook.Application.Run(tempModuleName & "." & procName)
                Else
                    'MessageBox.Show("无法在代码中找到可执行的过程")
                    GlobalStatusStrip.ShowWarning("无法在代码中找到可执行的过程")
                End If
            Else
                ' 代码不包含过程声明，将其包装在 Auto_Run 过程中
                Dim wrappedCode As String = "Sub Auto_Run()" & vbNewLine &
                                           vbaCode & vbNewLine &
                                           "End Sub"
                vbComp.CodeModule.AddFromString(wrappedCode)

                ' 执行 Auto_Run 过程
                workbook.Application.Run(tempModuleName & ".Auto_Run")
            End If

        Catch ex As Exception
            MessageBox.Show("执行 临时VBA 代码时出错: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 无论成功还是失败，都删除临时模块
            Try
                If vbProj IsNot Nothing AndAlso vbComp IsNot Nothing Then
                    vbProj.VBComponents.Remove(vbComp)
                End If
            Catch
                ' 忽略清理错误
            End Try
        End Try
        Return True
    End Function

    ' 捕获工作簿状态
    Private Function CaptureWorkbookState(workbook As Workbook) As Dictionary(Of String, WorksheetState)
        Dim state As New Dictionary(Of String, WorksheetState)

        For Each worksheet As Worksheet In workbook.Worksheets
            Dim sheetState As New WorksheetState(worksheet.Name)

            ' 获取使用范围
            Dim usedRange As Range = worksheet.UsedRange
            sheetState.UsedRangeAddress = usedRange.Address

            ' 捕获所有单元格的值
            For Each cell As Range In usedRange
                Dim address As String = cell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                sheetState.Cells(address) = cell.Value2
            Next

            state(worksheet.Name) = sheetState
        Next

        Return state
    End Function

    ' 比较工作簿状态
    Private Sub CompareWorkbookStates(
        beforeState As Dictionary(Of String, WorksheetState),
        afterState As Dictionary(Of String, WorksheetState),
        cellDifferences As List(Of CellDifference),
        sheetDifferences As List(Of SheetDifference))

        ' 检查工作表级别的更改（添加/删除工作表）
        For Each beforeSheet In beforeState.Values
            If Not afterState.ContainsKey(beforeSheet.Name) Then
                ' 工作表被删除
                sheetDifferences.Add(New SheetDifference(beforeSheet.Name, "删除"))
            End If
        Next

        For Each afterSheet In afterState.Values
            If Not beforeState.ContainsKey(afterSheet.Name) Then
                ' 工作表被添加
                sheetDifferences.Add(New SheetDifference(afterSheet.Name, "添加"))

                ' 添加所有新工作表的单元格作为"添加"
                For Each cell In afterSheet.Cells
                    cellDifferences.Add(New CellDifference(
                        afterSheet.Name, cell.Key, Nothing, cell.Value, "添加"))
                Next
            Else
                ' 工作表存在于两个状态中，比较单元格
                Dim beforeSheet = beforeState(afterSheet.Name)

                ' 检查单元格更改
                For Each afterCell In afterSheet.Cells
                    Dim address As String = afterCell.Key
                    Dim newValue As Object = afterCell.Value

                    If beforeSheet.Cells.ContainsKey(address) Then
                        Dim oldValue As Object = beforeSheet.Cells(address)

                        ' 比较值是否相等
                        If Not AreValuesEqual(oldValue, newValue) Then
                            cellDifferences.Add(New CellDifference(
                                afterSheet.Name, address, oldValue, newValue, "修改"))
                        End If
                    Else
                        ' 新添加的单元格
                        cellDifferences.Add(New CellDifference(
                            afterSheet.Name, address, Nothing, newValue, "添加"))
                    End If
                Next

                ' 检查删除的单元格
                For Each beforeCell In beforeSheet.Cells
                    Dim address As String = beforeCell.Key
                    If Not afterSheet.Cells.ContainsKey(address) Then
                        cellDifferences.Add(New CellDifference(
                            beforeSheet.Name, address, beforeCell.Value, Nothing, "删除"))
                    End If
                Next
            End If
        Next
    End Sub

    ' 比较两个值是否相等
    Private Function AreValuesEqual(value1 As Object, value2 As Object) As Boolean
        If value1 Is Nothing AndAlso value2 Is Nothing Then
            Return True
        ElseIf value1 Is Nothing OrElse value2 Is Nothing Then
            Return False
        Else
            Return value1.ToString() = value2.ToString()
        End If
    End Function

End Class