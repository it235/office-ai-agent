Imports System.Drawing
Imports System.Windows.Forms

Public Class MCPPresetForm
    Inherits Form

    Private _presetListView As ListView
    Private _okButton As Button
    Private _cancelButton As Button

    Public Property SelectedUrl As String
    Public Property SelectedApiKey As String

    Public Sub New()
        InitializeComponent()
        LoadPresets()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "MCP服务器预设"
        Me.Size = New Size(600, 400)
        Me.StartPosition = FormStartPosition.CenterParent

        _presetListView = New ListView()
        _presetListView.Location = New Point(10, 10)
        _presetListView.Size = New Size(570, 300)
        _presetListView.View = View.Details
        _presetListView.FullRowSelect = True
        _presetListView.Columns.Add("名称", 150)
        _presetListView.Columns.Add("URL", 300)
        _presetListView.Columns.Add("类型", 100)
        Me.Controls.Add(_presetListView)

        _okButton = New Button()
        _okButton.Text = "确定"
        _okButton.Location = New Point(420, 330)
        _okButton.Width = 80
        AddHandler _okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(_okButton)

        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(510, 330)
        _cancelButton.Width = 80
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)
    End Sub

    Private Sub LoadPresets()
        ' 预设的 MCP 服务器配置，修复 JS 路径的格式
        Dim presets = {
        New With {.Name = "本地 HTTP 服务器", .Url = "http://localhost:3000", .Type = "HTTP", .ApiKey = ""},
        New With {.Name = "本地 Node.js MCP 服务器", .Url = "stdio://node?args=\""F: \\ai\\node\\first-mcp-server\\build\\index.js\""", .Type = "Stdio", .ApiKey = ""},
        New With {.Name = "GitHub MCP 服务器", .Url = "stdio://npx?args=@modelcontextprotocol/server-github", .Type = "Stdio", .ApiKey = ""},
        New With {.Name = "GitLab MCP 服务器", .Url = "stdio://npx?args=@modelcontextprotocol/server-gitlab", .Type = "Stdio", .ApiKey = ""},
        New With {.Name = "MySQL MCP 服务器", .Url = "stdio://npx?args=@modelcontextprotocol/server-mysql", .Type = "Stdio", .ApiKey = ""},
        New With {.Name = "Python MCP 服务器", .Url = "stdio://python?args=\""-m mcp_server\""", .Type = "Stdio", .ApiKey = ""}
    }

        For Each preset In presets
            Dim item = New ListViewItem(preset.Name)
            item.SubItems.Add(preset.Url)
            item.SubItems.Add(preset.Type)
            item.Tag = preset
            _presetListView.Items.Add(item)
        Next
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        If _presetListView.SelectedItems.Count > 0 Then
            Dim selected = _presetListView.SelectedItems(0).Tag
            SelectedUrl = selected.Url
            SelectedApiKey = selected.ApiKey
            Me.DialogResult = DialogResult.OK
        End If
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class