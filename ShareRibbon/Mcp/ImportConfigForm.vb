Imports System.Drawing
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ImportConfigForm
    Inherits Form

    Private _jsonTextBox As TextBox
    Private _okButton As Button
    Private _cancelButton As Button

    Public Property ConfigJson As String
        Get
            Return _jsonTextBox.Text
        End Get
        Set(value As String)
            _jsonTextBox.Text = value
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "导入MCP配置"
        Me.Size = New Size(600, 450)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.FormBorderStyle = FormBorderStyle.FixedDialog

        ' 添加说明标签
        Dim instructionLabel As New Label()
        instructionLabel.Text = "请在下方粘贴MCP服务器配置的JSON文本:"
        instructionLabel.Location = New Point(20, 20)
        instructionLabel.AutoSize = True
        Me.Controls.Add(instructionLabel)

        ' 添加文本框用于粘贴JSON
        _jsonTextBox = New TextBox()
        _jsonTextBox.Multiline = True
        _jsonTextBox.ScrollBars = ScrollBars.Both
        _jsonTextBox.Location = New Point(20, 50)
        _jsonTextBox.Size = New Size(550, 300)
        _jsonTextBox.Font = New Font("Consolas", 9)
        _jsonTextBox.AcceptsReturn = True
        _jsonTextBox.AcceptsTab = True
        Me.Controls.Add(_jsonTextBox)

        ' 示例文本按钮
        Dim exampleButton As New Button()
        exampleButton.Text = "填充示例"
        exampleButton.Location = New Point(20, 360)
        exampleButton.Size = New Size(100, 30)
        AddHandler exampleButton.Click, AddressOf ExampleButton_Click
        Me.Controls.Add(exampleButton)

        ' 验证按钮
        Dim validateButton As New Button()
        validateButton.Text = "验证JSON"
        validateButton.Location = New Point(130, 360)
        validateButton.Size = New Size(100, 30)
        AddHandler validateButton.Click, AddressOf ValidateButton_Click
        Me.Controls.Add(validateButton)

        ' 确定按钮
        _okButton = New Button()
        _okButton.Text = "导入"
        _okButton.DialogResult = DialogResult.OK
        _okButton.Location = New Point(370, 360)
        _okButton.Size = New Size(90, 30)
        AddHandler _okButton.Click, AddressOf OkButton_Click
        Me.Controls.Add(_okButton)

        ' 取消按钮
        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.DialogResult = DialogResult.Cancel
        _cancelButton.Location = New Point(480, 360)
        _cancelButton.Size = New Size(90, 30)
        Me.Controls.Add(_cancelButton)

        Me.AcceptButton = _okButton
        Me.CancelButton = _cancelButton
    End Sub

    Private Sub ExampleButton_Click(sender As Object, e As EventArgs)
        ' 提供一个示例JSON格式
        Dim exampleJson As String = "
{
  ""mcpServers"": {
    ""amap-maps"": {
      ""command"": ""npx"",
      ""args"": [
        ""-y"",
        ""@amap/amap-maps-mcp-server""
      ],
      ""env"": {
        ""AMAP_MAPS_API_KEY"": ""请替换为您的API密钥""
      }
    }
  }
}".Trim()

        _jsonTextBox.Text = exampleJson
    End Sub

    Private Sub ValidateButton_Click(sender As Object, e As EventArgs)
        Try
            ' 尝试解析JSON以验证格式
            Dim json = _jsonTextBox.Text.Trim()
            If String.IsNullOrEmpty(json) Then
                MessageBox.Show("请先输入JSON配置", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim config = JObject.Parse(json)

            ' 检查是否包含mcpServers节点
            If config("mcpServers") Is Nothing Then
                MessageBox.Show("JSON格式有效，但缺少必要的'mcpServers'节点", "验证提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            MessageBox.Show("JSON格式有效！", "验证成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"JSON格式无效: {ex.Message}", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        Try
            ' 验证JSON格式
            Dim json = _jsonTextBox.Text.Trim()
            If String.IsNullOrEmpty(json) Then
                MessageBox.Show("请先输入JSON配置", "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.DialogResult = DialogResult.None
                Return
            End If

            ' 尝试解析，确保格式正确
            Dim config = JObject.Parse(json)

            ' 检查是否包含mcpServers节点
            If config("mcpServers") Is Nothing Then
                MessageBox.Show("JSON格式有效，但缺少必要的'mcpServers'节点", "导入提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Me.DialogResult = DialogResult.None
                Return
            End If

            ' 设置结果并关闭窗体
            ConfigJson = json
            Me.DialogResult = DialogResult.OK
        Catch ex As Exception
            MessageBox.Show($"JSON格式无效: {ex.Message}", "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.DialogResult = DialogResult.None
        End Try
    End Sub
End Class