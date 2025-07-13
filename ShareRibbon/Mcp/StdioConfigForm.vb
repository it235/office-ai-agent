Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Public Class StdioConfigForm
    Inherits Form

    Private _commandTextBox As TextBox
    Private _argumentsTextBox As TextBox
    Private _workingDirTextBox As TextBox
    Private _envVariablesGrid As DataGridView
    Private _envVariablesTextBox As TextBox
    Private _switchViewButton As Button
    Private _okButton As Button
    Private _cancelButton As Button
    Private _isGridView As Boolean = True

    Public Property Options As StdioOptions

    Public Sub New(options As StdioOptions)
        Me.Options = options
        InitializeComponent()
        LoadOptions()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Stdio连接配置"
        Me.Size = New Size(500, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        ' 命令输入
        Dim commandLabel As New Label()
        commandLabel.Text = "命令路径:"
        commandLabel.Location = New Point(20, 20)
        commandLabel.Width = 100
        Me.Controls.Add(commandLabel)

        _commandTextBox = New TextBox()
        _commandTextBox.Location = New Point(130, 17)
        _commandTextBox.Width = 250
        Me.Controls.Add(_commandTextBox)

        ' 参数输入
        Dim argsLabel As New Label()
        argsLabel.Text = "命令参数:"
        argsLabel.Location = New Point(20, 50)
        argsLabel.Width = 100
        Me.Controls.Add(argsLabel)

        _argumentsTextBox = New TextBox()
        _argumentsTextBox.Location = New Point(130, 47)
        _argumentsTextBox.Width = 340
        Me.Controls.Add(_argumentsTextBox)

        ' 环境变量标签和切换按钮
        Dim envLabel As New Label()
        envLabel.Text = "环境变量:"
        envLabel.Location = New Point(20, 110)
        envLabel.Width = 100
        Me.Controls.Add(envLabel)

        ' 添加切换视图按钮
        _switchViewButton = New Button()
        _switchViewButton.Text = "切换到文本视图"
        _switchViewButton.Location = New Point(330, 107)
        _switchViewButton.Width = 140
        AddHandler _switchViewButton.Click, AddressOf SwitchViewButton_Click
        Me.Controls.Add(_switchViewButton)

        ' 环境变量表格视图
        _envVariablesGrid = New DataGridView()
        _envVariablesGrid.Location = New Point(20, 140)
        _envVariablesGrid.Size = New Size(450, 270)
        _envVariablesGrid.AllowUserToAddRows = True
        _envVariablesGrid.AllowUserToDeleteRows = True
        _envVariablesGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        _envVariablesGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        _envVariablesGrid.Columns.Add("Key", "变量名")
        _envVariablesGrid.Columns.Add("Value", "变量值")
        Me.Controls.Add(_envVariablesGrid)

        ' 环境变量文本视图
        _envVariablesTextBox = New TextBox()
        _envVariablesTextBox.Location = New Point(20, 140)
        _envVariablesTextBox.Size = New Size(450, 270)
        _envVariablesTextBox.Multiline = True
        _envVariablesTextBox.ScrollBars = ScrollBars.Both
        _envVariablesTextBox.Font = New Font("Consolas", 9)
        _envVariablesTextBox.Visible = False
        Me.Controls.Add(_envVariablesTextBox)

        ' 底部按钮
        _okButton = New Button()
        _okButton.Text = "确定"
        _okButton.Location = New Point(310, 420)
        _okButton.Width = 80
        AddHandler _okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(_okButton)

        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(400, 420)
        _cancelButton.Width = 80
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)
    End Sub

    ' 从StdioOptions加载到表单
    Private Sub LoadOptions()
        _commandTextBox.Text = Options.Command
        _argumentsTextBox.Text = Options.Arguments

        ' 加载环境变量到表格
        LoadEnvToGrid()

        ' 同时加载到文本框
        LoadEnvToTextBox()
    End Sub

    Private Sub LoadEnvToGrid()
        _envVariablesGrid.Rows.Clear()
        For Each kvp In Options.EnvironmentVariables
            Dim rowIndex = _envVariablesGrid.Rows.Add()
            _envVariablesGrid.Rows(rowIndex).Cells("Key").Value = kvp.Key
            _envVariablesGrid.Rows(rowIndex).Cells("Value").Value = kvp.Value
        Next
    End Sub

    Private Sub LoadEnvToTextBox()
        Dim sb As New StringBuilder()
        For Each kvp In Options.EnvironmentVariables
            sb.AppendLine($"{kvp.Key}={kvp.Value}")
        Next
        _envVariablesTextBox.Text = sb.ToString()
    End Sub

    ' 从表单收集环境变量
    Private Sub CollectEnvironmentVariables()
        Options.EnvironmentVariables.Clear()

        If _isGridView Then
            ' 从表格收集
            For Each row As DataGridViewRow In _envVariablesGrid.Rows
                If row.IsNewRow Then Continue For

                Dim key = TryCast(row.Cells("Key").Value, String)
                Dim value = TryCast(row.Cells("Value").Value, String)

                If Not String.IsNullOrEmpty(key) Then
                    Options.EnvironmentVariables(key) = If(value IsNot Nothing, value, "")
                End If
            Next
        Else
            ' 从文本框收集
            For Each line In _envVariablesTextBox.Text.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                Dim parts = line.Split(New Char() {"="c}, 2)
                If parts.Length = 2 Then
                    Dim key = parts(0).Trim()
                    Dim value = parts(1).Trim()
                    If Not String.IsNullOrEmpty(key) Then
                        Options.EnvironmentVariables(key) = value
                    End If
                End If
            Next
        End If
    End Sub

    ' 切换视图按钮处理
    Private Sub SwitchViewButton_Click(sender As Object, e As EventArgs)
        If _isGridView Then
            ' 从表格切换到文本视图前，先更新环境变量集合
            CollectEnvironmentVariables()
            ' 然后加载到文本框
            LoadEnvToTextBox()

            ' 显示文本视图
            _envVariablesGrid.Visible = False
            _envVariablesTextBox.Visible = True
            _switchViewButton.Text = "切换到表格视图"
        Else
            ' 从文本视图切换到表格视图前，先更新环境变量集合
            CollectEnvironmentVariables()
            ' 然后加载到表格
            LoadEnvToGrid()

            ' 显示表格视图
            _envVariablesTextBox.Visible = False
            _envVariablesGrid.Visible = True
            _switchViewButton.Text = "切换到文本视图"
        End If

        _isGridView = Not _isGridView
    End Sub

    Private Sub BrowseButton_Click(sender As Object, e As EventArgs)
        Using dialog As New OpenFileDialog()
            dialog.Filter = "执行文件 (*.exe)|*.exe|JavaScript文件 (*.js)|*.js|Python文件 (*.py)|*.py|所有文件 (*.*)|*.*"
            dialog.Title = "选择命令文件"

            If dialog.ShowDialog() = DialogResult.OK Then
                _commandTextBox.Text = dialog.FileName
            End If
        End Using
    End Sub

    Private Sub BrowseDirButton_Click(sender As Object, e As EventArgs)
        Using dialog As New FolderBrowserDialog()
            dialog.Description = "选择工作目录"

            If dialog.ShowDialog() = DialogResult.OK Then
                _workingDirTextBox.Text = dialog.SelectedPath
            End If
        End Using
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        ' 保存命令和参数
        Options.Command = _commandTextBox.Text
        Options.Arguments = _argumentsTextBox.Text

        ' 收集环境变量
        CollectEnvironmentVariables()

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class