Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Text

Public Class StdioConfigForm
    Inherits Form

    Private _commandTextBox As TextBox
    Private _argumentsTextBox As TextBox
    Private _workingDirTextBox As TextBox
    Private _envVarsGrid As DataGridView
    Private _envVarsTextBox As TextBox  ' 新增：环境变量文本域
    Private _switchViewButton As Button ' 新增：切换视图按钮
    Private _okButton As Button
    Private _cancelButton As Button
    Private _isTextMode As Boolean = True ' 默认使用文本模式

    Public Property Options As StdioOptions

    Public Sub New(options As StdioOptions)
        Me.Options = options
        InitializeComponent()
        LoadOptions()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Stdio 连接高级设置"
        Me.Size = New Size(500, 550)  ' 增加高度以适应新控件
        Me.StartPosition = FormStartPosition.CenterParent

        ' 命令
        Dim commandLabel = New Label()
        commandLabel.Text = "命令路径:"
        commandLabel.Location = New Point(10, 15)
        commandLabel.Width = 100
        Me.Controls.Add(commandLabel)

        _commandTextBox = New TextBox()
        _commandTextBox.Location = New Point(120, 12)
        _commandTextBox.Width = 350
        Me.Controls.Add(_commandTextBox)

        ' 参数
        Dim argsLabel = New Label()
        argsLabel.Text = "命令参数:"
        argsLabel.Location = New Point(10, 45)
        argsLabel.Width = 100
        Me.Controls.Add(argsLabel)

        _argumentsTextBox = New TextBox()
        _argumentsTextBox.Location = New Point(120, 42)
        _argumentsTextBox.Width = 350
        Me.Controls.Add(_argumentsTextBox)

        ' 工作目录
        Dim workdirLabel = New Label()
        workdirLabel.Text = "工作目录:"
        workdirLabel.Location = New Point(10, 75)
        workdirLabel.Width = 100
        Me.Controls.Add(workdirLabel)

        _workingDirTextBox = New TextBox()
        _workingDirTextBox.Location = New Point(120, 72)
        _workingDirTextBox.Width = 350
        Me.Controls.Add(_workingDirTextBox)

        ' 环境变量标签和切换按钮
        Dim envVarsLabel = New Label()
        envVarsLabel.Text = "环境变量:"
        envVarsLabel.Location = New Point(10, 105)
        envVarsLabel.Width = 100
        Me.Controls.Add(envVarsLabel)

        _switchViewButton = New Button()
        _switchViewButton.Text = "切换到表格视图"
        _switchViewButton.Location = New Point(120, 102)
        _switchViewButton.Width = 150
        AddHandler _switchViewButton.Click, AddressOf SwitchViewButton_Click
        Me.Controls.Add(_switchViewButton)

        ' 环境变量文本域
        _envVarsTextBox = New TextBox()
        _envVarsTextBox.Location = New Point(10, 130)
        _envVarsTextBox.Size = New Size(465, 280)
        _envVarsTextBox.Multiline = True
        _envVarsTextBox.ScrollBars = ScrollBars.Both
        _envVarsTextBox.Font = New Font("Consolas", 9.0F)
        Me.Controls.Add(_envVarsTextBox)

        ' 环境变量表格
        _envVarsGrid = New DataGridView()
        _envVarsGrid.Location = New Point(10, 130)
        _envVarsGrid.Size = New Size(465, 280)
        _envVarsGrid.AutoGenerateColumns = False
        _envVarsGrid.AllowUserToAddRows = True
        _envVarsGrid.AllowUserToDeleteRows = True
        _envVarsGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        _envVarsGrid.Visible = False  ' 默认隐藏表格视图

        ' 添加列
        Dim nameColumn = New DataGridViewTextBoxColumn()
        nameColumn.HeaderText = "变量名"
        nameColumn.Name = "Name"
        nameColumn.Width = 150
        _envVarsGrid.Columns.Add(nameColumn)

        Dim valueColumn = New DataGridViewTextBoxColumn()
        valueColumn.HeaderText = "变量值"
        valueColumn.Name = "Value"
        valueColumn.Width = 270
        _envVarsGrid.Columns.Add(valueColumn)

        Me.Controls.Add(_envVarsGrid)

        ' 使用提示
        Dim helpLabel = New Label()
        helpLabel.Text = "提示: 每行一个环境变量，格式为 变量名=变量值"
        helpLabel.Location = New Point(10, 415)
        helpLabel.AutoSize = True
        helpLabel.ForeColor = Color.DarkBlue
        Me.Controls.Add(helpLabel)

        ' 按钮
        _okButton = New Button()
        _okButton.Text = "确定"
        _okButton.Location = New Point(325, 470)
        _okButton.Width = 70
        AddHandler _okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(_okButton)

        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(405, 470)
        _cancelButton.Width = 70
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)

        ' 设置初始视图
        UpdateViewMode()
    End Sub

    Private Sub SwitchViewButton_Click(sender As Object, e As EventArgs)
        ' 切换视图前，保存当前视图的数据
        If _isTextMode Then
            ' 从文本模式切换到表格模式，需要解析文本
            TextToGrid()
        Else
            ' 从表格模式切换到文本模式，需要生成文本
            GridToText()
        End If

        ' 切换视图模式
        _isTextMode = Not _isTextMode
        UpdateViewMode()
    End Sub

    Private Sub UpdateViewMode()
        If _isTextMode Then
            _envVarsTextBox.Visible = True
            _envVarsGrid.Visible = False
            _switchViewButton.Text = "切换到表格视图"
        Else
            _envVarsTextBox.Visible = False
            _envVarsGrid.Visible = True
            _switchViewButton.Text = "切换到文本视图"
        End If
    End Sub

    Private Sub TextToGrid()
        ' 清空网格
        _envVarsGrid.Rows.Clear()

        ' 解析文本中的环境变量
        Dim lines = _envVarsTextBox.Text.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

        For Each line In lines
            ' 跳过注释和空行
            If line.Trim().StartsWith("#") OrElse String.IsNullOrWhiteSpace(line) Then Continue For

            ' 解析 key=value 格式
            Dim equalIndex = line.IndexOf("=")
            If equalIndex > 0 Then
                Dim key = line.Substring(0, equalIndex).Trim()
                Dim value = If(equalIndex < line.Length - 1, line.Substring(equalIndex + 1).Trim(), "")

                ' 添加到表格
                Dim rowIndex = _envVarsGrid.Rows.Add()
                _envVarsGrid.Rows(rowIndex).Cells("Name").Value = key
                _envVarsGrid.Rows(rowIndex).Cells("Value").Value = value
            End If
        Next
    End Sub

    Private Sub GridToText()
        Dim sb = New StringBuilder()

        ' 将表格中的数据转换为文本
        For Each row As DataGridViewRow In _envVarsGrid.Rows
            If row.IsNewRow Then Continue For

            Dim name = row.Cells("Name").Value?.ToString()
            Dim value = row.Cells("Value").Value?.ToString()

            If Not String.IsNullOrEmpty(name) Then
                sb.AppendLine($"{name}={value}")
            End If
        Next

        _envVarsTextBox.Text = sb.ToString()
    End Sub

    Private Sub LoadOptions()
        _commandTextBox.Text = Options.Command
        _argumentsTextBox.Text = Options.Arguments
        _workingDirTextBox.Text = Options.WorkingDirectory

        ' 加载环境变量到文本框
        Dim sb = New StringBuilder()
        For Each kvp In Options.EnvironmentVariables
            sb.AppendLine($"{kvp.Key}={kvp.Value}")
        Next
        _envVarsTextBox.Text = sb.ToString()

        ' 也加载到表格
        _envVarsGrid.Rows.Clear()
        For Each kvp In Options.EnvironmentVariables
            Dim index = _envVarsGrid.Rows.Add()
            Dim row = _envVarsGrid.Rows(index)
            row.Cells("Name").Value = kvp.Key
            row.Cells("Value").Value = kvp.Value
        Next
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        ' 保存命令和参数
        Options.Command = _commandTextBox.Text.Trim()
        Options.Arguments = _argumentsTextBox.Text.Trim()
        Options.WorkingDirectory = _workingDirTextBox.Text.Trim()

        ' 保存环境变量
        Options.EnvironmentVariables.Clear()

        ' 根据当前视图模式获取环境变量
        If _isTextMode Then
            ' 从文本解析
            Dim lines = _envVarsTextBox.Text.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            For Each line In lines
                If line.Trim().StartsWith("#") OrElse String.IsNullOrWhiteSpace(line) Then Continue For

                Dim equalIndex = line.IndexOf("=")
                If equalIndex > 0 Then
                    Dim key = line.Substring(0, equalIndex).Trim()
                    Dim value = If(equalIndex < line.Length - 1, line.Substring(equalIndex + 1).Trim(), "")
                    Options.EnvironmentVariables(key) = value
                End If
            Next
        Else
            ' 从表格获取
            For Each row As DataGridViewRow In _envVarsGrid.Rows
                If Not row.IsNewRow AndAlso
                   row.Cells("Name").Value IsNot Nothing AndAlso
                   Not String.IsNullOrEmpty(row.Cells("Name").Value.ToString()) Then

                    Dim name = row.Cells("Name").Value.ToString()
                    Dim value = If(row.Cells("Value").Value Is Nothing, "", row.Cells("Value").Value.ToString())

                    Options.EnvironmentVariables(name) = value
                End If
            Next
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class