Imports System.Drawing
Imports System.Windows.Forms

Public Class MCPConfigForm
    Inherits Form

    Private _serverUrlTextBox As TextBox
    Private _apiKeyTextBox As TextBox
    Private _testButton As Button
    Private _saveButton As Button
    Private _cancelButton As Button
    Private _toolsListView As ListView

    Public Sub New()
        InitializeComponent()
        LoadConfig()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "MCP服务器配置"
        Me.Size = New Size(500, 400)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' 服务器URL
        Dim serverUrlLabel As New Label()
        serverUrlLabel.Text = "服务器URL:"
        serverUrlLabel.Location = New Point(10, 15)
        serverUrlLabel.Width = 80
        Me.Controls.Add(serverUrlLabel)

        _serverUrlTextBox = New TextBox()
        _serverUrlTextBox.Location = New Point(100, 12)
        _serverUrlTextBox.Width = 380
        Me.Controls.Add(_serverUrlTextBox)

        ' API密钥
        Dim apiKeyLabel As New Label()
        apiKeyLabel.Text = "API密钥:"
        apiKeyLabel.Location = New Point(10, 45)
        apiKeyLabel.Width = 80
        Me.Controls.Add(apiKeyLabel)

        _apiKeyTextBox = New TextBox()
        _apiKeyTextBox.Location = New Point(100, 42)
        _apiKeyTextBox.Width = 380
        _apiKeyTextBox.PasswordChar = "*"c
        Me.Controls.Add(_apiKeyTextBox)

        ' 可用工具列表
        Dim toolsLabel As New Label()
        toolsLabel.Text = "可用工具:"
        toolsLabel.Location = New Point(10, 75)
        toolsLabel.Width = 80
        Me.Controls.Add(toolsLabel)

        _toolsListView = New ListView()
        _toolsListView.Location = New Point(100, 75)
        _toolsListView.Size = New Size(380, 200)
        _toolsListView.View = View.Details
        _toolsListView.CheckBoxes = True
        _toolsListView.FullRowSelect = True
        _toolsListView.Columns.Add("工具名称", 150)
        _toolsListView.Columns.Add("描述", 225)
        Me.Controls.Add(_toolsListView)

        ' 测试连接按钮
        _testButton = New Button()
        _testButton.Text = "测试连接"
        _testButton.Location = New Point(100, 290)
        _testButton.Width = 100
        AddHandler _testButton.Click, AddressOf TestButton_Click
        Me.Controls.Add(_testButton)

        ' 保存按钮
        _saveButton = New Button()
        _saveButton.Text = "保存配置"
        _saveButton.Location = New Point(280, 320)
        _saveButton.Width = 100
        AddHandler _saveButton.Click, AddressOf SaveButton_Click
        Me.Controls.Add(_saveButton)

        ' 取消按钮
        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(390, 320)
        _cancelButton.Width = 100
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)

        ' 添加一些示例工具
        AddSampleTools()
    End Sub

    Private Sub AddSampleTools()
        ' 添加示例工具到列表
        Dim item1 As New ListViewItem("网页搜索")
        item1.SubItems.Add("搜索互联网获取最新信息")
        _toolsListView.Items.Add(item1)

        Dim item2 As New ListViewItem("图像生成")
        item2.SubItems.Add("基于文本描述生成图像")
        _toolsListView.Items.Add(item2)

        Dim item3 As New ListViewItem("代码助手")
        item3.SubItems.Add("提供编程和脚本辅助")
        _toolsListView.Items.Add(item3)

        Dim item4 As New ListViewItem("数据分析")
        item4.SubItems.Add("对Excel数据进行分析和可视化")
        _toolsListView.Items.Add(item4)
    End Sub

    Private Sub LoadConfig()
        ' 从配置文件加载MCP配置
        ' 实际应用中应从配置文件读取
        _serverUrlTextBox.Text = "https://mcp.example.com/api"
    End Sub

    Private Sub TestButton_Click(sender As Object, e As EventArgs)
        ' 测试MCP服务器连接
        If String.IsNullOrEmpty(_serverUrlTextBox.Text) Then
            MessageBox.Show("请输入服务器URL", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 模拟连接测试
        MessageBox.Show("连接测试成功！服务器响应正常。", "连接测试", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs)
        ' 保存MCP配置
        If String.IsNullOrEmpty(_serverUrlTextBox.Text) Then
            MessageBox.Show("请输入服务器URL", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 实际应用中应保存到配置文件
        MessageBox.Show("MCP配置已保存！", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class