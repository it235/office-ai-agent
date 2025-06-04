Imports System.Drawing
Imports System.Windows.Forms
Imports ShareRibbon

Public Class TextPreviewForm
    Inherits Form

    Private textBox As TextBox
    Private inputTextBox As TextBox
    Private confirmButton As Button
    Private descriptionLabel1 As Label
    Private descriptionLabel As Label
    Private useButton As Button

    Public Property IsConfirmed As Boolean
    Public Property InputText As String
    Public Property DEFAULT_INPUT_TEXT As String = "注意，注意，需要在这里输入你整理表格的诉求！"
    Public Property DEMO_Q As String = "请按照采购日期、商品、单价、重量整理下方数据"

    Public Sub New(text As String)
        Me.Text = "问题描述&内容预览"
        Me.Size = New Size(500, 500)
        Me.StartPosition = FormStartPosition.CenterScreen


        descriptionLabel1 = New Label()
        descriptionLabel1.Text = "问题示例：" & DEMO_Q
        descriptionLabel1.Dock = DockStyle.Fill
        descriptionLabel1.TextAlign = ContentAlignment.MiddleLeft

        useButton = New Button()
        useButton.Text = "使用示例"
        AddHandler useButton.Click, AddressOf UseButtonButton_Click

        textBox = New TextBox()
        textBox.Multiline = True
        textBox.ScrollBars = ScrollBars.Vertical
        textBox.Dock = DockStyle.Fill
        textBox.Text = text
        textBox.ReadOnly = True

        inputTextBox = New TextBox()
        inputTextBox.Multiline = True
        inputTextBox.ScrollBars = ScrollBars.Vertical
        inputTextBox.Dock = DockStyle.Top
        inputTextBox.Height = 100
        inputTextBox.Text = If(String.IsNullOrEmpty(InputText), DEFAULT_INPUT_TEXT, InputText)
        inputTextBox.ForeColor = If(String.IsNullOrEmpty(InputText), Color.Gray, Color.Black)

        AddHandler inputTextBox.Enter, AddressOf inputTextBox_Enter ' 添加 Enter 事件处理程序
        AddHandler inputTextBox.Leave, AddressOf inputTextBox_Leave ' 添加 Leave 事件处理程序

        confirmButton = New Button()
        confirmButton.Text = "执行数据整理"
        AddHandler confirmButton.Click, AddressOf ConfirmButton_Click

        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.FlowDirection = FlowDirection.LeftToRight
        buttonPanel.Dock = DockStyle.Fill
        buttonPanel.Controls.Add(confirmButton)
        buttonPanel.AutoSize = True
        buttonPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink
        buttonPanel.WrapContents = False

        descriptionLabel = New Label()
        descriptionLabel.Text = "以下是您选中的单元格，请确认"
        descriptionLabel.Dock = DockStyle.Fill
        descriptionLabel.TextAlign = ContentAlignment.MiddleCenter

        Dim mainPanel As New TableLayoutPanel()
        mainPanel.Dock = DockStyle.Fill
        mainPanel.RowCount = 4
        mainPanel.ColumnCount = 1
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 100.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainPanel.Controls.Add(descriptionLabel1, 0, 0)
        mainPanel.Controls.Add(useButton, 0, 1)

        mainPanel.Controls.Add(inputTextBox, 0, 2)
        mainPanel.Controls.Add(buttonPanel, 0, 3)
        mainPanel.Controls.Add(descriptionLabel, 0, 4)
        mainPanel.Controls.Add(textBox, 0, 5)

        Me.Controls.Add(mainPanel)

        '加入底部告警栏
        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    Private Sub inputTextBox_Enter(sender As Object, e As EventArgs)
        If inputTextBox.Text = DEFAULT_INPUT_TEXT Then
            inputTextBox.Text = ""
            inputTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub inputTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(inputTextBox.Text) Then
            inputTextBox.Text = DEFAULT_INPUT_TEXT
            inputTextBox.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(inputTextBox.Text) Or DEFAULT_INPUT_TEXT = inputTextBox.Text Then
            'MessageBox.Show("请输入问题内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            GlobalStatusStrip.ShowWarning("请输入问题内容！")
            Return
        End If

        IsConfirmed = True
        InputText = inputTextBox.Text
        Me.Close()
    End Sub

    Private Sub UseButtonButton_Click(sender As Object, e As EventArgs)
        inputTextBox.Text = DEMO_Q
        inputTextBox.ForeColor = Color.Black
    End Sub
End Class


