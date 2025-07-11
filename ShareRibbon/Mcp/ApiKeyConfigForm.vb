Imports System.Drawing
Imports System.Windows.Forms

Public Class ApiKeyConfigForm
    Inherits Form

    Private _apiKeyTextBox As TextBox
    Private _useHeaderRadio As RadioButton
    Private _useUrlRadio As RadioButton
    Private _okButton As Button
    Private _cancelButton As Button

    Public Property ApiKey As String
    Public Property AddToUrl As Boolean = False

    Public Sub New(apiKey As String)
        Me.ApiKey = apiKey
        InitializeComponent()
        _apiKeyTextBox.Text = apiKey
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "API 密钥设置"
        Me.Size = New Size(400, 200)
        Me.StartPosition = FormStartPosition.CenterParent

        ' API 密钥输入框
        Dim keyLabel = New Label()
        keyLabel.Text = "API 密钥:"
        keyLabel.Location = New Point(10, 15)
        keyLabel.Width = 70
        Me.Controls.Add(keyLabel)

        _apiKeyTextBox = New TextBox()
        _apiKeyTextBox.Location = New Point(85, 12)
        _apiKeyTextBox.Width = 290
        _apiKeyTextBox.PasswordChar = "*"c
        Me.Controls.Add(_apiKeyTextBox)

        ' 选项：如何传递密钥
        Dim optionsGroup = New GroupBox()
        optionsGroup.Text = "传递方式"
        optionsGroup.Location = New Point(10, 45)
        optionsGroup.Size = New Size(365, 80)
        Me.Controls.Add(optionsGroup)

        _useHeaderRadio = New RadioButton()
        _useHeaderRadio.Text = "通过 Authorization 头部传递 (推荐)"
        _useHeaderRadio.Location = New Point(10, 20)
        _useHeaderRadio.Width = 300
        _useHeaderRadio.Checked = True
        optionsGroup.Controls.Add(_useHeaderRadio)

        _useUrlRadio = New RadioButton()
        _useUrlRadio.Text = "通过 URL 参数传递 (格式: ?api_key=xxx)"
        _useUrlRadio.Location = New Point(10, 45)
        _useUrlRadio.Width = 300
        optionsGroup.Controls.Add(_useUrlRadio)

        ' 按钮
        _okButton = New Button()
        _okButton.Text = "确定"
        _okButton.Location = New Point(230, 135)
        _okButton.Width = 70
        AddHandler _okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(_okButton)

        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(310, 135)
        _cancelButton.Width = 70
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        ApiKey = _apiKeyTextBox.Text.Trim()
        AddToUrl = _useUrlRadio.Checked
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class