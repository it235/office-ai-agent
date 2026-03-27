Imports System.Drawing
Imports System.Windows.Forms
Imports ShareRibbon

''' <summary>字段定义（用于批量数据生成）</summary>
Public Class FieldDefinition
    Public Property FieldName As String
    Public Property CellColumn As String
    Public Property FieldDescription As String
End Class

Public Class BatchDataGenerationForm
    Inherits Form

    Private _fieldListView As ListView
    Private _addButton As Button
    Private _removeButton As Button
    Private _generateButton As Button
    Private _cancelButton As Button
    Private _rowCountInput As NumericUpDown

    Public ReadOnly Property Fields As List(Of FieldDefinition)
        Get
            Dim list As New List(Of FieldDefinition)()
            For Each item As ListViewItem In _fieldListView.Items
                list.Add(New FieldDefinition With {
                    .FieldName = item.Text,
                    .CellColumn = item.SubItems(1).Text,
                    .FieldDescription = item.SubItems(2).Text
                })
            Next
            Return list
        End Get
    End Property

    Public ReadOnly Property RowCount As Integer
        Get
            Return CInt(_rowCountInput.Value)
        End Get
    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "批量数据生成"
        Me.Size = New Size(620, 450)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.FormBorderStyle = FormBorderStyle.FixedDialog

        ' 字段列表
        _fieldListView = New ListView()
        _fieldListView.Dock = DockStyle.Top
        _fieldListView.Height = 270
        _fieldListView.View = View.Details
        _fieldListView.FullRowSelect = True
        _fieldListView.GridLines = True
        _fieldListView.Columns.Add("字段名", 120)
        _fieldListView.Columns.Add("目标列（如A）", 110)
        _fieldListView.Columns.Add("字段描述", 360)
        Me.Controls.Add(_fieldListView)

        ' 行数设置面板
        Dim rowCountPanel As New Panel()
        rowCountPanel.Dock = DockStyle.Top
        rowCountPanel.Height = 40
        Me.Controls.Add(rowCountPanel)

        Dim rowCountLabel As New Label()
        rowCountLabel.Text = "生成行数："
        rowCountLabel.Location = New Point(10, 12)
        rowCountLabel.Width = 70
        rowCountPanel.Controls.Add(rowCountLabel)

        _rowCountInput = New NumericUpDown()
        _rowCountInput.Location = New Point(85, 10)
        _rowCountInput.Width = 80
        _rowCountInput.Minimum = 1
        _rowCountInput.Maximum = 500
        _rowCountInput.Value = 10
        rowCountPanel.Controls.Add(_rowCountInput)

        ' 添加一个提示标签
        Dim hintLabel As New Label()
        hintLabel.Text = "提示：目标列填写Excel列字母（如A、B、C），AI将按字段描述生成对应内容"
        hintLabel.Location = New Point(180, 12)
        hintLabel.Width = 420
        hintLabel.ForeColor = Drawing.Color.Gray
        rowCountPanel.Controls.Add(hintLabel)

        ' 操作按钮面板
        Dim buttonPanel As New Panel()
        buttonPanel.Dock = DockStyle.Bottom
        buttonPanel.Height = 50
        Me.Controls.Add(buttonPanel)

        _addButton = New Button()
        _addButton.Text = "添加字段"
        _addButton.Location = New Point(10, 12)
        _addButton.Width = 100
        AddHandler _addButton.Click, AddressOf AddButton_Click
        buttonPanel.Controls.Add(_addButton)

        _removeButton = New Button()
        _removeButton.Text = "移除字段"
        _removeButton.Location = New Point(120, 12)
        _removeButton.Width = 100
        AddHandler _removeButton.Click, AddressOf RemoveButton_Click
        buttonPanel.Controls.Add(_removeButton)

        _generateButton = New Button()
        _generateButton.Text = "生成数据"
        _generateButton.Location = New Point(390, 12)
        _generateButton.Width = 100
        AddHandler _generateButton.Click, AddressOf GenerateButton_Click
        buttonPanel.Controls.Add(_generateButton)

        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(500, 12)
        _cancelButton.Width = 100
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        buttonPanel.Controls.Add(_cancelButton)
    End Sub

    Private Sub AddButton_Click(sender As Object, e As EventArgs)
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
        If _fieldListView.Items.Count = 0 Then
            MessageBox.Show("请先添加至少一个字段。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
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
        Me.Size = New Size(400, 230)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        Dim fieldNameLabel As New Label()
        fieldNameLabel.Text = "字段名称:"
        fieldNameLabel.Location = New Point(10, 15)
        fieldNameLabel.Width = 80
        Me.Controls.Add(fieldNameLabel)

        _fieldNameTextBox = New TextBox()
        _fieldNameTextBox.Location = New Point(100, 12)
        _fieldNameTextBox.Width = 280
        Me.Controls.Add(_fieldNameTextBox)

        Dim cellColumnLabel As New Label()
        cellColumnLabel.Text = "目标列:"
        cellColumnLabel.Location = New Point(10, 45)
        cellColumnLabel.Width = 80
        Me.Controls.Add(cellColumnLabel)

        _cellColumnTextBox = New TextBox()
        _cellColumnTextBox.Location = New Point(100, 42)
        _cellColumnTextBox.Width = 80
        _cellColumnTextBox.MaxLength = 3
        Me.Controls.Add(_cellColumnTextBox)

        Dim colHint As New Label()
        colHint.Text = "（Excel列字母，如 A、B、AA）"
        colHint.Location = New Point(190, 45)
        colHint.Width = 200
        colHint.ForeColor = Drawing.Color.Gray
        Me.Controls.Add(colHint)

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

        _okButton = New Button()
        _okButton.Text = "确定"
        _okButton.Location = New Point(210, 155)
        _okButton.Width = 80
        AddHandler _okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(_okButton)

        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(300, 155)
        _cancelButton.Width = 80
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(_fieldNameTextBox.Text) Then
            MessageBox.Show("请输入字段名称。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim colText = _cellColumnTextBox.Text.Trim().ToUpper()
        If String.IsNullOrWhiteSpace(colText) Then
            MessageBox.Show("请输入目标列（如 A、B）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        ' 只允许 1~3 个大写字母（对应 Excel 最大列 XFD=16384）；
        ' 不校验会导致 ColumnLetterToIndex 返回 0，数据静默写入失败，用户不知道哪里错了
        If colText.Length > 3 OrElse colText.Any(Function(ch) ch < "A"c OrElse ch > "Z"c) Then
            MessageBox.Show("目标列只能是 1~3 个英文字母（如 A、B、AA、XFD）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        FieldName = _fieldNameTextBox.Text.Trim()
        CellColumn = colText
        FieldDescription = _fieldDescTextBox.Text.Trim()
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class
