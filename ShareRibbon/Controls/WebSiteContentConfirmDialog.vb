Imports System.Drawing
Imports System.Windows.Forms

Public Class WebSiteContentConfirmDialog
    Inherits Form

    Private _content As String
        Private _previewText As String
        Private _tag As String
        Private _path As String

        Public Sub New(content As String, tag As String, path As String)
        'InitializeComponent()
        _content = content
            _tag = tag
            _path = path
            _previewText = If(content.Length > 50, content.Substring(0, 50) & "...", content)
            InitializeUI()
        End Sub

        Private Sub InitializeUI()
            Text = "ȷ��ѡ��"
            StartPosition = FormStartPosition.CenterScreen
            Size = New Size(500, 300)
            MinimizeBox = False
            MaximizeBox = False
            FormBorderStyle = FormBorderStyle.FixedDialog

            ' ����Ԥ���ı���
            Dim previewBox As New TextBox With {
                .Multiline = True,
                .ReadOnly = True,
                .ScrollBars = ScrollBars.Vertical,
                .Dock = DockStyle.Top,
                .Height = 180,
                .Text = $"��ѡ��Ԫ��: <{_tag}>{Environment.NewLine}·��: {_path}{Environment.NewLine}Ԥ��: {_content}"
            }
            Controls.Add(previewBox)

            ' ������ť���
            Dim buttonPanel As New FlowLayoutPanel With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .Height = 40,
                .Padding = New Padding(5)
            }

            ' ����������ť
            Dim btnCancel As New Button With {
                .Text = "ȡ������",
                .DialogResult = DialogResult.Cancel,
                .Width = 100
            }

            Dim btnUseContent As New Button With {
                .Text = "ֱ��ʹ������",
                .DialogResult = DialogResult.Yes,
                .Width = 120
            }

            Dim btnAiChat As New Button With {
                .Text = "����AI����",
                .DialogResult = DialogResult.No,
                .Width = 100
            }

            ' ��Ӱ�ť�����
            buttonPanel.Controls.Add(btnCancel)
            buttonPanel.Controls.Add(btnUseContent)
            buttonPanel.Controls.Add(btnAiChat)
            Controls.Add(buttonPanel)

            AcceptButton = btnUseContent
            CancelButton = btnCancel
        End Sub

        Public ReadOnly Property Content As String
            Get
                Return _content
            End Get
        End Property
    End Class
