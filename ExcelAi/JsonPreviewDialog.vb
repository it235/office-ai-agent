' ExcelAi\JsonPreviewDialog.vb
' JSON命令预览对话框：在执行JSON命令前展示差异预览

Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon

''' <summary>
''' JSON命令预览对话框
''' 类似VBA预览功能，在执行前展示将要进行的更改
''' </summary>
Public Class JsonPreviewDialog
    Inherits Form

    Private WithEvents tabControl As TabControl
    Private WithEvents tabSummary As TabPage
    Private WithEvents tabCellChanges As TabPage
    Private WithEvents tabJsonCode As TabPage
    
    Private summaryTextBox As RichTextBox
    Private cellChangesListView As ListView
    Private jsonCodeTextBox As RichTextBox
    
    Private WithEvents btnExecute As Button
    Private WithEvents btnCancel As Button

    Private _previewResult As JsonPreviewResult

    Public Sub New()
        InitializeComponent()
    End Sub

    ''' <summary>
    ''' 显示预览并返回用户选择
    ''' </summary>
    Public Function ShowPreview(previewResult As JsonPreviewResult) As DialogResult
        _previewResult = previewResult
        PopulatePreviewData()
        Return Me.ShowDialog()
    End Function

    Private Sub InitializeComponent()
        Me.Text = "JSON 命令预览"
        Me.Size = New Size(700, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 创建 TabControl
        tabControl = New TabControl()
        tabControl.Dock = DockStyle.Top
        tabControl.Height = 380

        ' Tab1: 执行计划摘要
        tabSummary = New TabPage("执行计划")
        summaryTextBox = New RichTextBox()
        summaryTextBox.Dock = DockStyle.Fill
        summaryTextBox.ReadOnly = True
        summaryTextBox.Font = New Font("Microsoft YaHei", 10)
        summaryTextBox.BackColor = Color.White
        tabSummary.Controls.Add(summaryTextBox)
        tabControl.TabPages.Add(tabSummary)

        ' Tab2: 单元格变更列表
        tabCellChanges = New TabPage("单元格变更")
        cellChangesListView = New ListView()
        cellChangesListView.Dock = DockStyle.Fill
        cellChangesListView.View = View.Details
        cellChangesListView.FullRowSelect = True
        cellChangesListView.GridLines = True
        cellChangesListView.Columns.Add("地址", 80)
        cellChangesListView.Columns.Add("变更类型", 80)
        cellChangesListView.Columns.Add("原值", 200)
        cellChangesListView.Columns.Add("新值", 200)
        tabCellChanges.Controls.Add(cellChangesListView)
        tabControl.TabPages.Add(tabCellChanges)

        ' Tab3: JSON命令详情
        tabJsonCode = New TabPage("JSON 命令")
        jsonCodeTextBox = New RichTextBox()
        jsonCodeTextBox.Dock = DockStyle.Fill
        jsonCodeTextBox.ReadOnly = True
        jsonCodeTextBox.Font = New Font("Consolas", 10)
        jsonCodeTextBox.BackColor = Color.FromArgb(30, 30, 30)
        jsonCodeTextBox.ForeColor = Color.FromArgb(220, 220, 220)
        tabJsonCode.Controls.Add(jsonCodeTextBox)
        tabControl.TabPages.Add(tabJsonCode)

        Me.Controls.Add(tabControl)

        ' 按钮面板
        Dim buttonPanel As New Panel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.Height = 60
        buttonPanel.Padding = New Padding(10)
        buttonPanel.BackColor = Color.FromArgb(245, 245, 245)

        btnCancel = New Button()
        btnCancel.Text = "取消"
        btnCancel.Size = New Size(100, 35)
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.FlatStyle = FlatStyle.Flat
        btnCancel.FlatAppearance.BorderColor = Color.Gray

        btnExecute = New Button()
        btnExecute.Text = "确认执行"
        btnExecute.Size = New Size(100, 35)
        btnExecute.BackColor = Color.FromArgb(74, 111, 165)
        btnExecute.ForeColor = Color.White
        btnExecute.FlatStyle = FlatStyle.Flat
        btnExecute.FlatAppearance.BorderSize = 0
        btnExecute.DialogResult = DialogResult.OK

        ' 使用FlowLayoutPanel来自动布局按钮
        Dim flowPanel As New FlowLayoutPanel()
        flowPanel.Dock = DockStyle.Right
        flowPanel.FlowDirection = FlowDirection.RightToLeft
        flowPanel.AutoSize = True
        flowPanel.Padding = New Padding(5)
        flowPanel.Controls.Add(btnExecute)
        flowPanel.Controls.Add(btnCancel)

        buttonPanel.Controls.Add(flowPanel)
        Me.Controls.Add(buttonPanel)

        Me.AcceptButton = btnExecute
        Me.CancelButton = btnCancel
    End Sub

    ''' <summary>
    ''' 填充预览数据
    ''' </summary>
    Private Sub PopulatePreviewData()
        If _previewResult Is Nothing Then Return

        ' 填充摘要
        PopulateSummary()

        ' 填充单元格变更
        PopulateCellChanges()

        ' 填充JSON代码
        PopulateJsonCode()
    End Sub

    Private Sub PopulateSummary()
        summaryTextBox.Clear()
        
        ' 标题
        summaryTextBox.SelectionFont = New Font("Microsoft YaHei", 14, FontStyle.Bold)
        summaryTextBox.SelectionColor = Color.FromArgb(74, 111, 165)
        summaryTextBox.AppendText("执行计划预览" & vbCrLf & vbCrLf)

        ' 摘要
        If Not String.IsNullOrEmpty(_previewResult.Summary) Then
            summaryTextBox.SelectionFont = New Font("Microsoft YaHei", 10)
            summaryTextBox.SelectionColor = Color.Black
            summaryTextBox.AppendText(_previewResult.Summary & vbCrLf & vbCrLf)
        End If

        ' 执行步骤
        If _previewResult.ExecutionPlan IsNot Nothing AndAlso _previewResult.ExecutionPlan.Count > 0 Then
            summaryTextBox.SelectionFont = New Font("Microsoft YaHei", 11, FontStyle.Bold)
            summaryTextBox.SelectionColor = Color.FromArgb(74, 111, 165)
            summaryTextBox.AppendText("执行步骤：" & vbCrLf)

            For Each execStep In _previewResult.ExecutionPlan
                Dim icon = GetStepIcon(execStep.Icon)
                summaryTextBox.SelectionFont = New Font("Microsoft YaHei", 10)
                summaryTextBox.SelectionColor = Color.Black
                summaryTextBox.AppendText($"  {execStep.StepNumber}. {icon} {execStep.Description}")
                
                If Not String.IsNullOrEmpty(execStep.WillModify) Then
                    summaryTextBox.SelectionColor = Color.FromArgb(230, 81, 0)
                    summaryTextBox.AppendText($" → {execStep.WillModify}")
                End If
                summaryTextBox.AppendText(vbCrLf)
            Next
        End If

        ' 变更统计
        If _previewResult.CellChanges IsNot Nothing AndAlso _previewResult.CellChanges.Count > 0 Then
            summaryTextBox.AppendText(vbCrLf)
            summaryTextBox.SelectionFont = New Font("Microsoft YaHei", 11, FontStyle.Bold)
            summaryTextBox.SelectionColor = Color.FromArgb(74, 111, 165)
            summaryTextBox.AppendText("预计变更：" & vbCrLf)

            Dim addedCount = _previewResult.CellChanges.Where(Function(c) c.ChangeType = "Added").Count()
            Dim modifiedCount = _previewResult.CellChanges.Where(Function(c) c.ChangeType = "Modified").Count()
            Dim deletedCount = _previewResult.CellChanges.Where(Function(c) c.ChangeType = "Deleted").Count()

            summaryTextBox.SelectionFont = New Font("Microsoft YaHei", 10)
            If addedCount > 0 Then
                summaryTextBox.SelectionColor = Color.Green
                summaryTextBox.AppendText($"  + 新增: {addedCount} 个单元格" & vbCrLf)
            End If
            If modifiedCount > 0 Then
                summaryTextBox.SelectionColor = Color.Orange
                summaryTextBox.AppendText($"  ~ 修改: {modifiedCount} 个单元格" & vbCrLf)
            End If
            If deletedCount > 0 Then
                summaryTextBox.SelectionColor = Color.Red
                summaryTextBox.AppendText($"  - 删除: {deletedCount} 个单元格" & vbCrLf)
            End If
        Else
            summaryTextBox.AppendText(vbCrLf)
            summaryTextBox.SelectionFont = New Font("Microsoft YaHei", 10)
            summaryTextBox.SelectionColor = Color.Gray
            summaryTextBox.AppendText("（此命令不会产生单元格变更预览）" & vbCrLf)
        End If
    End Sub

    Private Sub PopulateCellChanges()
        cellChangesListView.Items.Clear()

        If _previewResult.CellChanges Is Nothing Then Return

        For Each change In _previewResult.CellChanges
            Dim item As New ListViewItem(change.Address)
            item.SubItems.Add(GetChangeTypeText(change.ChangeType))
            item.SubItems.Add(If(change.OldValue?.ToString(), ""))
            item.SubItems.Add(If(change.NewValue?.ToString(), ""))

            ' 根据变更类型设置颜色
            Select Case change.ChangeType
                Case "Added"
                    item.BackColor = Color.FromArgb(232, 245, 233) ' 淡绿色
                Case "Modified"
                    item.BackColor = Color.FromArgb(255, 243, 224) ' 淡橙色
                Case "Deleted"
                    item.BackColor = Color.FromArgb(255, 235, 238) ' 淡红色
            End Select

            cellChangesListView.Items.Add(item)
        Next
    End Sub

    Private Sub PopulateJsonCode()
        jsonCodeTextBox.Clear()

        If String.IsNullOrEmpty(_previewResult.OriginalJson) Then Return

        Try
            ' 格式化JSON
            Dim json = JObject.Parse(_previewResult.OriginalJson)
            Dim formattedJson = json.ToString(Formatting.Indented)
            jsonCodeTextBox.Text = formattedJson
        Catch
            jsonCodeTextBox.Text = _previewResult.OriginalJson
        End Try
    End Sub

    Private Function GetStepIcon(iconType As String) As String
        Select Case iconType?.ToLower()
            Case "search"
                Return "🔍"
            Case "data"
                Return "📊"
            Case "formula"
                Return "🧮"
            Case "chart"
                Return "📈"
            Case "format"
                Return "🎨"
            Case "clean"
                Return "🧹"
            Case Else
                Return "⚡"
        End Select
    End Function

    Private Function GetChangeTypeText(changeType As String) As String
        Select Case changeType
            Case "Added"
                Return "新增"
            Case "Modified"
                Return "修改"
            Case "Deleted"
                Return "删除"
            Case Else
                Return changeType
        End Select
    End Function

End Class
