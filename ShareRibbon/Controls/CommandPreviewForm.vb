Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq

''' <summary>
''' 命令预览表单 - 提供比MessageBox更丰富的预览体验
''' </summary>
Public Class CommandPreviewForm
    Inherits Form

    Private listView As ListView
    Private detailTextBox As TextBox
    Private confirmButton As Button
    Private cancelButton As Button
    Private splitContainer As SplitContainer

    Public Property IsConfirmed As Boolean = False

    ''' <summary>
    ''' 创建命令预览表单
    ''' </summary>
    ''' <param name="title">表单标题</param>
    ''' <param name="commands">命令列表(JArray或JObject)</param>
    Public Sub New(title As String, commands As JToken)
        Me.Text = title
        Me.Size = New Size(700, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(500, 350)

        InitializeComponents()
        LoadCommands(commands)
    End Sub

    Private Sub InitializeComponents()
        ' 创建分割容器
        splitContainer = New SplitContainer()
        splitContainer.Dock = DockStyle.Fill
        splitContainer.Orientation = Orientation.Horizontal
        splitContainer.SplitterDistance = 250

        ' 命令列表
        listView = New ListView()
        listView.Dock = DockStyle.Fill
        listView.View = View.Details
        listView.FullRowSelect = True
        listView.GridLines = True
        listView.Columns.Add("序号", 50)
        listView.Columns.Add("命令", 120)
        listView.Columns.Add("描述", 400)
        AddHandler listView.SelectedIndexChanged, AddressOf ListView_SelectedIndexChanged

        ' 详情文本框
        detailTextBox = New TextBox()
        detailTextBox.Dock = DockStyle.Fill
        detailTextBox.Multiline = True
        detailTextBox.ScrollBars = ScrollBars.Both
        detailTextBox.ReadOnly = True
        detailTextBox.Font = New Font("Consolas", 10)
        detailTextBox.BackColor = Color.FromArgb(250, 250, 250)

        splitContainer.Panel1.Controls.Add(listView)
        splitContainer.Panel2.Controls.Add(detailTextBox)

        ' 底部按钮面板
        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.FlowDirection = FlowDirection.RightToLeft
        buttonPanel.Height = 45
        buttonPanel.Padding = New Padding(10)

        cancelButton = New Button()
        cancelButton.Text = "取消"
        cancelButton.Size = New Size(80, 28)
        cancelButton.DialogResult = DialogResult.Cancel
        AddHandler cancelButton.Click, AddressOf CancelButton_Click

        confirmButton = New Button()
        confirmButton.Text = "确认执行"
        confirmButton.Size = New Size(90, 28)
        confirmButton.BackColor = Color.FromArgb(0, 120, 212)
        confirmButton.ForeColor = Color.White
        confirmButton.FlatStyle = FlatStyle.Flat
        AddHandler confirmButton.Click, AddressOf ConfirmButton_Click

        buttonPanel.Controls.Add(cancelButton)
        buttonPanel.Controls.Add(confirmButton)

        ' 顶部信息标签
        Dim infoLabel As New Label()
        infoLabel.Dock = DockStyle.Top
        infoLabel.Height = 30
        infoLabel.Text = "  选择命令查看详情，点击「确认执行」开始执行所有命令"
        infoLabel.TextAlign = ContentAlignment.MiddleLeft
        infoLabel.BackColor = Color.FromArgb(240, 240, 240)

        Me.Controls.Add(splitContainer)
        Me.Controls.Add(buttonPanel)
        Me.Controls.Add(infoLabel)

        Me.AcceptButton = confirmButton
        Me.CancelButton = cancelButton
    End Sub

    Private Sub LoadCommands(commands As JToken)
        listView.Items.Clear()

        If commands Is Nothing Then Return

        Dim cmdList As JArray
        If commands.Type = JTokenType.Array Then
            cmdList = CType(commands, JArray)
        ElseIf commands.Type = JTokenType.Object Then
            Dim cmdObj = CType(commands, JObject)
            If cmdObj("commands") IsNot Nothing Then
                cmdList = CType(cmdObj("commands"), JArray)
            Else
                ' 单命令
                cmdList = New JArray()
                cmdList.Add(cmdObj)
            End If
        Else
            Return
        End If

        Dim index = 1
        For Each cmd In cmdList
            If cmd.Type = JTokenType.Object Then
                Dim cmdObj = CType(cmd, JObject)
                Dim cmdName = cmdObj("command")?.ToString()
                Dim description = GetCommandDescription(cmdObj)

                Dim item As New ListViewItem(index.ToString())
                item.SubItems.Add(cmdName)
                item.SubItems.Add(description)
                item.Tag = cmdObj.ToString(Newtonsoft.Json.Formatting.Indented)

                listView.Items.Add(item)
                index += 1
            End If
        Next

        ' 自动调整列宽
        listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)

        ' 选中第一项
        If listView.Items.Count > 0 Then
            listView.Items(0).Selected = True
        End If
    End Sub

    Private Function GetCommandDescription(cmdObj As JObject) As String
        Dim sb As New StringBuilder()
        Dim cmdName = cmdObj("command")?.ToString()
        Dim params = cmdObj("params")

        Select Case cmdName?.ToLower()
            Case "inserttext"
                Dim content = params?("content")?.ToString()
                If Not String.IsNullOrEmpty(content) Then
                    sb.Append($"插入文本: {content.Substring(0, Math.Min(60, content.Length))}")
                    If content.Length > 60 Then sb.Append("...")
                End If

            Case "formattext"
                Dim range = params?("range")?.ToString()
                sb.Append($"格式化{If(range = "all", "全文", "选中内容")}: ")
                If params?("bold")?.Value(Of Boolean)() = True Then sb.Append("加粗 ")
                If params?("italic")?.Value(Of Boolean)() = True Then sb.Append("斜体 ")
                Dim fontSize = params?("fontSize")?.Value(Of Integer)()
                If fontSize > 0 Then sb.Append($"字号{fontSize} ")

            Case "replacetext"
                Dim find = params?("find")?.ToString()
                Dim replace = params?("replace")?.ToString()
                sb.Append($"替换 ""{find}"" 为 ""{replace}""")

            Case "inserttable"
                Dim rows = params?("rows")?.Value(Of Integer)()
                Dim cols = params?("cols")?.Value(Of Integer)()
                sb.Append($"插入 {rows}行×{cols}列 表格")

            Case "applystyle"
                Dim styleName = params?("styleName")?.ToString()
                sb.Append($"应用样式: {styleName}")

            Case "generatetoc"
                Dim levels = params?("levels")?.Value(Of Integer)()
                sb.Append($"生成目录 (级别: {If(levels > 0, levels, 3)})")

            Case "beautifydocument"
                sb.Append("美化文档格式")

            Case Else
                sb.Append(cmdName)
        End Select

        Return sb.ToString()
    End Function

    Private Sub ListView_SelectedIndexChanged(sender As Object, e As EventArgs)
        If listView.SelectedItems.Count > 0 Then
            Dim selectedItem = listView.SelectedItems(0)
            detailTextBox.Text = CStr(selectedItem.Tag)
        Else
            detailTextBox.Text = ""
        End If
    End Sub

    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs)
        IsConfirmed = True
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        IsConfirmed = False
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    ''' <summary>
    ''' 显示命令预览对话框
    ''' </summary>
    ''' <returns>用户是否确认执行</returns>
    Public Shared Function ShowPreview(title As String, commands As JToken) As Boolean
        Using form As New CommandPreviewForm(title, commands)
            form.ShowDialog()
            Return form.IsConfirmed
        End Using
    End Function
End Class
