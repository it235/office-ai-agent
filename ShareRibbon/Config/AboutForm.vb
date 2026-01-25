' ShareRibbon\Config\AboutForm.vb
Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' 关于对话框 - 显示插件信息和开源地址
''' </summary>
Public Class AboutForm
    Inherits Form

    Private lblTitle As Label
    Private lblDescription As Label
    Private lblAuthor As Label
    Private lblDataPath As Label
    Private lblBili As LinkLabel
    Private lblGitee As LinkLabel
    Private lblGithub As LinkLabel
    Private btnClose As Button

    Public Sub New()
        InitializeComponents()
    End Sub

    Private Sub InitializeComponents()
        Me.Text = "关于 Office MOSS 助手"
        Me.Size = New Size(450, 420)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White

        ' 标题
        lblTitle = New Label()
        lblTitle.Text = "Office MOSS 助手"
        lblTitle.Font = New Font("微软雅黑", 16, FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(74, 111, 165)
        lblTitle.Location = New Point(20, 20)
        lblTitle.AutoSize = True
        Me.Controls.Add(lblTitle)

        ' 描述
        lblDescription = New Label()
        lblDescription.Text = "大家好，我是B站的君哥，账号【君哥聊编程】。" & vbCrLf & vbCrLf &
                             "该插件的灵感来自一位B站粉丝，他从事银行审计工作，" & vbCrLf &
                             "经常与表格打交道。很多时候表格中的数据无法通过" & vbCrLf &
                             "固定的公式来计算，但在人类理解上具有相同的意义，" & vbCrLf &
                             "所以 Office MOSS 诞生了。" & vbCrLf & vbCrLf &
                             "插件持续优化中，欢迎留言或评论，不断完善该插件。"
        lblDescription.Font = New Font("微软雅黑", 9)
        lblDescription.ForeColor = Color.FromArgb(80, 80, 80)
        lblDescription.Location = New Point(20, 55)
        lblDescription.Size = New Size(400, 130)
        Me.Controls.Add(lblDescription)

        ' 数据路径
        lblDataPath = New Label()
        lblDataPath.Text = "数据存放目录: 我的文档\" & ConfigSettings.OfficeAiAppDataFolder
        lblDataPath.Font = New Font("微软雅黑", 9)
        lblDataPath.ForeColor = Color.Gray
        lblDataPath.Location = New Point(20, 190)
        lblDataPath.AutoSize = True
        Me.Controls.Add(lblDataPath)

        ' 开源地址标题
        Dim lblOpenSource As New Label()
        lblOpenSource.Text = "开源地址:"
        lblOpenSource.Font = New Font("微软雅黑", 9, FontStyle.Bold)
        lblOpenSource.Location = New Point(20, 225)
        lblOpenSource.AutoSize = True
        Me.Controls.Add(lblOpenSource)

        ' B站链接
        lblBili = New LinkLabel()
        lblBili.Text = "bilibili: https://www.bilibili.com/video/BV17vNRz1ELn"
        lblBili.Font = New Font("微软雅黑", 9)
        lblBili.Location = New Point(20, 250)
        lblBili.AutoSize = True
        lblBili.LinkColor = Color.FromArgb(74, 111, 165)
        AddHandler lblBili.LinkClicked, AddressOf Bilibili_LinkClicked
        Me.Controls.Add(lblBili)

        ' Gitee链接
        lblGitee = New LinkLabel()
        lblGitee.Text = "Gitee: https://gitee.com/it235/office-ai-agent"
        lblGitee.Font = New Font("微软雅黑", 9)
        lblGitee.Location = New Point(20, 275)
        lblGitee.AutoSize = True
        lblGitee.LinkColor = Color.FromArgb(74, 111, 165)
        AddHandler lblGitee.LinkClicked, AddressOf Gitee_LinkClicked
        Me.Controls.Add(lblGitee)

        ' Github链接
        lblGithub = New LinkLabel()
        lblGithub.Text = "Github: https://github.com/it235/office-ai-agent"
        lblGithub.Font = New Font("微软雅黑", 9)
        lblGithub.Location = New Point(20, 300)
        lblGithub.AutoSize = True
        lblGithub.LinkColor = Color.FromArgb(74, 111, 165)
        AddHandler lblGithub.LinkClicked, AddressOf Github_LinkClicked
        Me.Controls.Add(lblGithub)

        ' 关闭按钮
        btnClose = New Button()
        btnClose.Text = "关闭"
        btnClose.Size = New Size(80, 30)
        btnClose.Location = New Point(350, 330)
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.BackColor = Color.FromArgb(74, 111, 165)
        btnClose.ForeColor = Color.White
        btnClose.Font = New Font("微软雅黑", 9)
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        Me.Controls.Add(btnClose)
        Me.AcceptButton = btnClose
    End Sub

    Private Sub Bilibili_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Try
            System.Diagnostics.Process.Start("https://www.bilibili.com/video/BV17vNRz1ELn")
        Catch ex As Exception
            MessageBox.Show("无法打开链接: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Gitee_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Try
            System.Diagnostics.Process.Start("https://gitee.com/it235/office-ai-agent")
        Catch ex As Exception
            MessageBox.Show("无法打开链接: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Github_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Try
            System.Diagnostics.Process.Start("https://github.com/it235/office-ai-agent")
        Catch ex As Exception
            MessageBox.Show("无法打开链接: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class
