Imports System.Drawing
Imports System.Windows.Forms

Public Class BatchDataGenerationForm
    Inherits Form

    Private _fieldListView As ListView
    Private _addButton As Button
    Private _removeButton As Button
    Private _generateButton As Button
    Private _cancelButton As Button

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "批量数据生成"
        Me.Size = New Size(600, 400)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' 初始化ListView
        _fieldListView = New ListView()
        _fieldListView.Dock = DockStyle.Top
        _fieldListView.Height = 250
        _fieldListView.View = View.Details
        _fieldListView.FullRowSelect = True
        _fieldListView.Columns.Add("字段名", 100)
        _fieldListView.Columns.Add("单元格列", 100)
        _fieldListView.Columns.Add("字段描述", 350)
        Me.Controls.Add(_fieldListView)

        ' 创建按钮面板
        Dim buttonPanel As New Panel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.Height = 50
        Me.Controls.Add(buttonPanel)

        ' 添加按钮
        _addButton = New Button()
        _addButton.Text = "添加字段"
        _addButton.Location = New Point(10, 10)
        _addButton.Width = 100
        AddHandler _addButton.Click, AddressOf AddButton_Click
        buttonPanel.Controls.Add(_addButton)

        ' 移除按钮
        _removeButton = New Button()
        _removeButton.Text = "移除字段"
        _removeButton.Location = New Point(120, 10)
        _removeButton.Width = 100
        AddHandler _removeButton.Click, AddressOf RemoveButton_Click
        buttonPanel.Controls.Add(_removeButton)

        ' 生成按钮
        _generateButton = New Button()
        _generateButton.Text = "生成数据"
        _generateButton.Location = New Point(380, 10)
        _generateButton.Width = 100
        AddHandler _generateButton.Click, AddressOf GenerateButton_Click
        buttonPanel.Controls.Add(_generateButton)

        ' 取消按钮
        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(490, 10)
        _cancelButton.Width = 100
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        buttonPanel.Controls.Add(_cancelButton)
    End Sub

    Private Sub AddButton_Click(sender As Object, e As EventArgs)
        ' 创建字段输入对话框
        Using inputForm As New FieldInputForm()
            If inputForm.ShowDialog() = DialogResult.OK Then
                Dim item As New ListViewItem(inputForm.FieldName)
                item.SubItems.Add(inputForm.CellColumn)
                item.SubItems.Add(inputForm.FieldDescription)
                _fieldListView.Items.Add(item)
            End If
        End Using
    End Sub

    Private Sub RemoveButton_Click(sender As Object, e As EventArgs)
        If _fieldListView.SelectedItems.Count > 0 Then
            _fieldListView.Items.Remove(_fieldListView.SelectedItems(0))
        End If
    End Sub

    Private Sub GenerateButton_Click(sender As Object, e As EventArgs)
        ' 实现数据生成逻辑
        ' 这里将根据Excel应用程序的上下文生成数据
        MessageBox.Show("数据生成功能将在Excel环境中实现")
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class

' 用于输入字段信息的子表单
Public Class FieldInputForm
    Inherits Form

    Private _fieldNameTextBox As TextBox
    Private _cellColumnTextBox As TextBox
    Private _fieldDescTextBox As TextBox
    Private _okButton As Button
    Private _cancelButton As Button

    Public Property FieldName As String
    Public Property CellColumn As String
    Public Property FieldDescription As String

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "添加字段"
        Me.Size = New Size(400, 220)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 字段名称
        Dim fieldNameLabel As New Label()
        fieldNameLabel.Text = "字段名称:"
        fieldNameLabel.Location = New Point(10, 15)
        fieldNameLabel.Width = 80
        Me.Controls.Add(fieldNameLabel)

        _fieldNameTextBox = New TextBox()
        _fieldNameTextBox.Location = New Point(100, 12)
        _fieldNameTextBox.Width = 280
        Me.Controls.Add(_fieldNameTextBox)

        ' 单元格列
        Dim cellColumnLabel As New Label()
        cellColumnLabel.Text = "单元格列:"
        cellColumnLabel.Location = New Point(10, 45)
        cellColumnLabel.Width = 80
        Me.Controls.Add(cellColumnLabel)

        _cellColumnTextBox = New TextBox()
        _cellColumnTextBox.Location = New Point(100, 42)
        _cellColumnTextBox.Width = 280
        Me.Controls.Add(_cellColumnTextBox)

        ' 字段描述
        Dim fieldDescLabel As New Label()
        fieldDescLabel.Text = "字段描述:"
        fieldDescLabel.Location = New Point(10, 75)
        fieldDescLabel.Width = 80
        Me.Controls.Add(fieldDescLabel)

        _fieldDescTextBox = New TextBox()
        _fieldDescTextBox.Location = New Point(100, 72)
        _fieldDescTextBox.Width = 280
        _fieldDescTextBox.Height = 60
        _fieldDescTextBox.Multiline = True
        Me.Controls.Add(_fieldDescTextBox)

        ' 确定按钮
        _okButton = New Button()
        _okButton.Text = "确定"
        _okButton.Location = New Point(210, 150)
        _okButton.Width = 80
        AddHandler _okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(_okButton)

        ' 取消按钮
        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(300, 150)
        _cancelButton.Width = 80
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        FieldName = _fieldNameTextBox.Text
        CellColumn = _cellColumnTextBox.Text
        FieldDescription = _fieldDescTextBox.Text
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class