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
    Public Property DEFAULT_INPUT_TEXT As String = "ע�⣬ע�⣬��Ҫ�����������������������"
    Public Property DEMO_Q As String = "�밴�ղɹ����ڡ���Ʒ�����ۡ����������·�����"

    Public Sub New(text As String)
        Me.Text = "��������&����Ԥ��"
        Me.Size = New Size(500, 500)
        Me.StartPosition = FormStartPosition.CenterScreen


        descriptionLabel1 = New Label()
        descriptionLabel1.Text = "����ʾ����" & DEMO_Q
        descriptionLabel1.Dock = DockStyle.Fill
        descriptionLabel1.TextAlign = ContentAlignment.MiddleLeft

        useButton = New Button()
        useButton.Text = "ʹ��ʾ��"
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

        AddHandler inputTextBox.Enter, AddressOf inputTextBox_Enter ' ��� Enter �¼��������
        AddHandler inputTextBox.Leave, AddressOf inputTextBox_Leave ' ��� Leave �¼��������

        confirmButton = New Button()
        confirmButton.Text = "ִ����������"
        AddHandler confirmButton.Click, AddressOf ConfirmButton_Click

        Dim buttonPanel As New FlowLayoutPanel()
        buttonPanel.FlowDirection = FlowDirection.LeftToRight
        buttonPanel.Dock = DockStyle.Fill
        buttonPanel.Controls.Add(confirmButton)
        buttonPanel.AutoSize = True
        buttonPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink
        buttonPanel.WrapContents = False

        descriptionLabel = New Label()
        descriptionLabel.Text = "��������ѡ�еĵ�Ԫ����ȷ��"
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

        '����ײ��澯��
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
            'MessageBox.Show("�������������ݡ�", "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            GlobalStatusStrip.ShowWarning("�������������ݣ�")
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


