Imports System.Drawing
Imports System.Windows.Forms

Public Class TranslateGlobalSettingsForm
    Inherits Form

    Public Property Settings As TranslateSettings

    Private qpsNumeric As NumericUpDown
    Private enableSelectionCheck As CheckBox
    Private promptTextBox As TextBox
    Private okButton As Button
    Private cancelButton As Button

    Public Sub New()
        Me.Text = "全局翻译设置"
        Me.Size = New Size(480, 360)
        Me.StartPosition = FormStartPosition.CenterParent

        Dim lblQps As New Label() With {.Text = "每秒最大请求数:", .Location = New Point(12, 12), .AutoSize = True}
        Me.Controls.Add(lblQps)

        qpsNumeric = New NumericUpDown() With {
            .Location = New Point(140, 10),
            .Minimum = 1,
            .Maximum = 100,
            .Value = 5
        }
        Me.Controls.Add(qpsNumeric)

        enableSelectionCheck = New CheckBox() With {
            .Text = "启用划词翻译（选中文本自动翻译）",
            .Location = New Point(12, 44),
            .AutoSize = True
        }
        Me.Controls.Add(enableSelectionCheck)

        Dim lblPrompt As New Label() With {.Text = "翻译提示词（system prompt）:", .Location = New Point(12, 80), .AutoSize = True}
        Me.Controls.Add(lblPrompt)

        promptTextBox = New TextBox() With {
            .Location = New Point(12, 104),
            .Size = New Size(440, 160),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical
        }
        Me.Controls.Add(promptTextBox)

        okButton = New Button() With {.Text = "保存", .Location = New Point(300, 280), .Size = New Size(70, 28)}
        AddHandler okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(okButton)

        cancelButton = New Button() With {.Text = "取消", .Location = New Point(380, 280), .Size = New Size(70, 28)}
        AddHandler cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(cancelButton)
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        If Settings Is Nothing Then Settings = New TranslateSettings()
        qpsNumeric.Value = Settings.MaxRequestsPerSecond
        enableSelectionCheck.Checked = Settings.EnableSelectionTranslate
        promptTextBox.Text = Settings.PromptText
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        Settings.MaxRequestsPerSecond = CInt(qpsNumeric.Value)
        Settings.EnableSelectionTranslate = enableSelectionCheck.Checked
        Settings.PromptText = promptTextBox.Text
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class