' ShareRibbon\Config\SkillsConfigForm.vb
' Skills配置窗口：展示Claude规范的Skills

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports ShareRibbon.Services

''' <summary>
''' Skills配置窗口：展示Claude规范的Skills（只读列表）
''' </summary>
Public Class SkillsConfigForm
    Inherits Form

    Private listBox As ListBox
    Private _skills As New List(Of SkillFileDefinition)()

    ' 详情区域控件
    Private lblName As Label
    Private lblDescription As Label
    Private lblLicense As Label
    Private lblCompatibility As Label
    Private lblAllowedTools As Label
    Private lblAuthor As Label
    Private lblVersion As Label
    Private txtContent As TextBox

    Public Sub New()
        Me.Text = "Skills配置"
        Me.Size = New Size(850, 550)
        Me.MinimumSize = New Size(700, 450)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Microsoft YaHei UI", 9)
        AddHandler Me.FormClosing, AddressOf OnFormClosing
        AddHandler Me.Shown, AddressOf OnFormShown
        InitializeUI()
    End Sub

    Private Sub OnFormShown(sender As Object, e As EventArgs)
        LoadSkills()
    End Sub

    Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
        If Me.Controls.Contains(GlobalStatusStrip.StatusStrip) Then
            Me.Controls.Remove(GlobalStatusStrip.StatusStrip)
        End If
    End Sub

    Private Sub InitializeUI()
        ' 顶部说明
        Dim lblInfo As New Label() With {
            .Text = "Skills目录：Documents\OfficeAiAppData\Skills，将符合Claude规范的Skills目录拷贝到此即可",
            .Location = New Point(15, 10),
            .Size = New Size(800, 20),
            .ForeColor = Color.Gray,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }
        Me.Controls.Add(lblInfo)

        ' 可拖拽分隔的左右布局
        Dim split As New SplitContainer() With {
            .Location = New Point(15, 35),
            .Size = New Size(805, 430),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .SplitterDistance = 280,
            .Panel1MinSize = 150,
            .Panel2MinSize = 300,
            .FixedPanel = FixedPanel.None
        }
        split.Panel1.SuspendLayout()
        split.Panel2.SuspendLayout()

        ' 左侧：Skills列表
        Dim lblList As New Label() With {
            .Text = "已安装的Skills：",
            .Location = New Point(0, 0),
            .Size = New Size(270, 20),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
        }
        split.Panel1.Controls.Add(lblList)
        listBox = New ListBox() With {
            .Location = New Point(0, 22),
            .Size = New Size(276, 403),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .HorizontalScrollbar = True,
            .HorizontalExtent = 500
        }
        AddHandler listBox.SelectedIndexChanged, AddressOf ListSelectionChanged
        split.Panel1.Controls.Add(listBox)

        ' 右侧：详情区域
        Dim xRight As Integer = 10
        Dim p2Y As Integer = 0

        ' 名称
        lblName = New Label() With {
            .Text = "名称：",
            .Location = New Point(xRight, p2Y),
            .Size = New Size(70, 20),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left,
            .Font = New Font(Me.Font, FontStyle.Bold)
        }
        split.Panel2.Controls.Add(lblName)
        Dim txtName As New Label() With {
            .Name = "txtName",
            .Location = New Point(xRight + 70, p2Y),
            .Size = New Size(430, 20),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .ForeColor = Color.FromArgb(70, 130, 180)
        }
        split.Panel2.Controls.Add(txtName)
        p2Y += 28

        ' 描述
        lblDescription = New Label() With {
            .Text = "描述：",
            .Location = New Point(xRight, p2Y),
            .Size = New Size(70, 20),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
        }
        split.Panel2.Controls.Add(lblDescription)
        Dim txtDescription As New Label() With {
            .Name = "txtDescription",
            .Location = New Point(xRight + 70, p2Y),
            .Size = New Size(430, 40),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .ForeColor = Color.DarkGray
        }
        split.Panel2.Controls.Add(txtDescription)
        p2Y += 48

        ' 元数据行
        Dim metadataPanel As New Panel() With {
            .Location = New Point(xRight + 70, p2Y),
            .Size = New Size(430, 80),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .BackColor = Color.FromArgb(245, 245, 245)
        }
        split.Panel2.Controls.Add(metadataPanel)

        lblLicense = New Label() With {
            .Text = "许可证：",
            .Location = New Point(5, 5),
            .Size = New Size(60, 18),
            .ForeColor = Color.Gray
        }
        metadataPanel.Controls.Add(lblLicense)
        Dim txtLicense As New Label() With {
            .Name = "txtLicense",
            .Location = New Point(70, 5),
            .Size = New Size(355, 18),
            .ForeColor = Color.FromArgb(100, 100, 100)
        }
        metadataPanel.Controls.Add(txtLicense)

        lblCompatibility = New Label() With {
            .Text = "环境：",
            .Location = New Point(5, 28),
            .Size = New Size(60, 18),
            .ForeColor = Color.Gray
        }
        metadataPanel.Controls.Add(lblCompatibility)
        Dim txtCompatibility As New Label() With {
            .Name = "txtCompatibility",
            .Location = New Point(70, 28),
            .Size = New Size(355, 18),
            .ForeColor = Color.FromArgb(100, 100, 100)
        }
        metadataPanel.Controls.Add(txtCompatibility)

        lblAllowedTools = New Label() With {
            .Text = "工具：",
            .Location = New Point(5, 51),
            .Size = New Size(60, 18),
            .ForeColor = Color.Gray
        }
        metadataPanel.Controls.Add(lblAllowedTools)
        Dim txtAllowedTools As New Label() With {
            .Name = "txtAllowedTools",
            .Location = New Point(70, 51),
            .Size = New Size(355, 18),
            .ForeColor = Color.FromArgb(100, 100, 100)
        }
        metadataPanel.Controls.Add(txtAllowedTools)

        lblAuthor = New Label() With {
            .Text = "作者：",
            .Location = New Point(5, 74),
            .Size = New Size(60, 18),
            .ForeColor = Color.Gray
        }
        metadataPanel.Controls.Add(lblAuthor)
        Dim txtAuthor As New Label() With {
            .Name = "txtAuthor",
            .Location = New Point(70, 74),
            .Size = New Size(200, 18),
            .ForeColor = Color.FromArgb(100, 100, 100)
        }
        metadataPanel.Controls.Add(txtAuthor)

        lblVersion = New Label() With {
            .Text = "版本：",
            .Location = New Point(280, 74),
            .Size = New Size(50, 18),
            .ForeColor = Color.Gray
        }
        metadataPanel.Controls.Add(lblVersion)
        Dim txtVersion As New Label() With {
            .Name = "txtVersion",
            .Location = New Point(330, 74),
            .Size = New Size(95, 18),
            .ForeColor = Color.FromArgb(100, 100, 100)
        }
        metadataPanel.Controls.Add(txtVersion)

        p2Y += 88

        ' 内容区域
        Dim lblContent As New Label() With {
            .Text = "Skill内容：",
            .Location = New Point(xRight, p2Y),
            .Size = New Size(70, 20),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
        }
        split.Panel2.Controls.Add(lblContent)
        txtContent = New TextBox() With {
            .Location = New Point(xRight + 70, p2Y),
            .Size = New Size(430, 210),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .ReadOnly = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .BackColor = Color.White,
            .Font = New Font("Consolas", 8.5)
        }
        split.Panel2.Controls.Add(txtContent)

        split.Panel1.ResumeLayout(False)
        split.Panel2.ResumeLayout(False)
        Me.Controls.Add(split)

        ' 底部按钮
        Dim y = 475
        Dim btnOpenDir As New Button() With {
            .Text = "打开Skills目录",
            .Location = New Point(15, y),
            .Size = New Size(120, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left,
            .BackColor = Color.FromArgb(70, 130, 180),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnOpenDir.Click, AddressOf BtnOpenDirClick
        Me.Controls.Add(btnOpenDir)

        Dim btnRefresh As New Button() With {
            .Text = "刷新列表",
            .Location = New Point(145, y),
            .Size = New Size(100, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        }
        AddHandler btnRefresh.Click, AddressOf BtnRefreshClick
        Me.Controls.Add(btnRefresh)

        Dim btnClose As New Button() With {
            .Text = "关闭",
            .Location = New Point(740, y),
            .Size = New Size(80, 28),
            .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        }
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    Private Sub LoadSkills()
        Try
            SkillsDirectoryService.EnsureDirectoryExists()
            _skills = SkillsDirectoryService.GetAllSkills(forceRefresh:=True)

            listBox.Items.Clear()
            For Each skill In _skills
                listBox.Items.Add(New ListItem With {.Skill = skill, .DisplayText = skill.Name})
            Next

            If listBox.Items.Count = 0 Then
                listBox.Items.Add("(暂无Skills，请打开Skills目录添加)")
            End If
        Catch ex As Exception
            listBox.Items.Clear()
            listBox.Items.Add("(加载失败: " & ex.Message & ")")
            GlobalStatusStrip.ShowWarning("加载Skills失败")
        End Try
    End Sub

    Private Sub ListSelectionChanged(sender As Object, e As EventArgs)
        Dim item = TryCast(listBox.SelectedItem, ListItem)
        If item Is Nothing Then
            ClearDetail()
            Return
        End If

        Dim skill = item.Skill
        If skill Is Nothing Then
            ClearDetail()
            Return
        End If

        ' 更新详情
        Dim txtName = Me.Controls.Find("txtName", True).FirstOrDefault()
        If txtName IsNot Nothing Then txtName.Text = skill.Name

        Dim txtDescription = Me.Controls.Find("txtDescription", True).FirstOrDefault()
        If txtDescription IsNot Nothing Then txtDescription.Text = If(String.IsNullOrWhiteSpace(skill.Description), "(无描述)", skill.Description)

        Dim txtLicense = Me.Controls.Find("txtLicense", True).FirstOrDefault()
        If txtLicense IsNot Nothing Then txtLicense.Text = If(String.IsNullOrWhiteSpace(skill.License), "-", skill.License)

        Dim txtCompatibility = Me.Controls.Find("txtCompatibility", True).FirstOrDefault()
        If txtCompatibility IsNot Nothing Then txtCompatibility.Text = If(String.IsNullOrWhiteSpace(skill.Compatibility), "-", skill.Compatibility)

        Dim txtAllowedTools = Me.Controls.Find("txtAllowedTools", True).FirstOrDefault()
        If txtAllowedTools IsNot Nothing Then txtAllowedTools.Text = If(String.IsNullOrWhiteSpace(skill.AllowedToolsText), "-", skill.AllowedToolsText)

        Dim txtAuthor = Me.Controls.Find("txtAuthor", True).FirstOrDefault()
        If txtAuthor IsNot Nothing Then txtAuthor.Text = If(String.IsNullOrWhiteSpace(skill.Author), "-", skill.Author)

        Dim txtVersion = Me.Controls.Find("txtVersion", True).FirstOrDefault()
        If txtVersion IsNot Nothing Then txtVersion.Text = If(String.IsNullOrWhiteSpace(skill.Version), "-", skill.Version)

        txtContent.Text = If(String.IsNullOrWhiteSpace(skill.Content), "(无内容)", skill.Content)
    End Sub

    Private Sub ClearDetail()
        Dim txtName = Me.Controls.Find("txtName", True).FirstOrDefault()
        If txtName IsNot Nothing Then txtName.Text = ""

        Dim txtDescription = Me.Controls.Find("txtDescription", True).FirstOrDefault()
        If txtDescription IsNot Nothing Then txtDescription.Text = ""

        Dim txtLicense = Me.Controls.Find("txtLicense", True).FirstOrDefault()
        If txtLicense IsNot Nothing Then txtLicense.Text = ""

        Dim txtCompatibility = Me.Controls.Find("txtCompatibility", True).FirstOrDefault()
        If txtCompatibility IsNot Nothing Then txtCompatibility.Text = ""

        Dim txtAllowedTools = Me.Controls.Find("txtAllowedTools", True).FirstOrDefault()
        If txtAllowedTools IsNot Nothing Then txtAllowedTools.Text = ""

        Dim txtAuthor = Me.Controls.Find("txtAuthor", True).FirstOrDefault()
        If txtAuthor IsNot Nothing Then txtAuthor.Text = ""

        Dim txtVersion = Me.Controls.Find("txtVersion", True).FirstOrDefault()
        If txtVersion IsNot Nothing Then txtVersion.Text = ""

        txtContent.Text = ""
    End Sub

    Private Sub BtnOpenDirClick(sender As Object, e As EventArgs)
        Try
            SkillsDirectoryService.OpenSkillsDirectory()
            GlobalStatusStrip.ShowInfo("已打开Skills目录")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("打开目录失败: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnRefreshClick(sender As Object, e As EventArgs)
        LoadSkills()
        GlobalStatusStrip.ShowInfo("已刷新Skills列表")
    End Sub

    Private Class ListItem
        Public Property Skill As SkillFileDefinition
        Public Property DisplayText As String
        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class
End Class
