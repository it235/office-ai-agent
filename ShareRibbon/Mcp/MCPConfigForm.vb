Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class MCPConfigForm
    Inherits Form

    Private _serverUrlTextBox As TextBox
    Private _apiKeyTextBox As TextBox
    Private _testButton As Button
    Private _saveButton As Button
    Private _cancelButton As Button
    Private _statusLabel As Label
    Private _tabControl As TabControl

    ' 各功能的ListView
    Private _toolsListView As ListView
    Private _resourcesListView As ListView
    Private _promptsListView As ListView

    ' 测试区域
    Private _testToolCombo As ComboBox
    Private _testParametersTextBox As TextBox
    Private _testResultTextBox As TextBox
    Private _executeTestButton As Button

    Private _mcpClient As StreamJsonRpcMCPClient
    Private _currentTools As List(Of MCPToolInfo)

    ' 新增连接管理相关成员
    Private _connectionsListView As ListView
    Private _addConnectionButton As Button
    Private _removeConnectionButton As Button
    Private _connectionNameTextBox As TextBox
    ' 修改字段类型
    Private _currentConnections As List(Of MCPConnectionConfig)

    Private _currentConnectionName As String = String.Empty

    ' 在 CreateConnectionConfigArea 方法中添加成员变量引用
    Private _advancedButton As Button

    Public Sub New()
        _mcpClient = New StreamJsonRpcMCPClient()
        _currentTools = New List(Of MCPToolInfo)()
        _currentConnections = MCPConnectionManager.LoadConnections()
        InitializeComponent()
        LoadConfig()

        ' 添加窗体加载事件处理
        AddHandler Me.Load, AddressOf MCPConfigForm_Load
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "通用MCP客户端配置"
        Me.Size = New Size(900, 600)  ' 增加窗体宽度
        Me.StartPosition = FormStartPosition.CenterScreen

        ' 状态标签
        _statusLabel = New Label()
        _statusLabel.Text = "就绪 - 请配置MCP服务器连接"
        _statusLabel.Location = New Point(210, 10)  ' 向右移动
        _statusLabel.Width = 680
        _statusLabel.ForeColor = Color.Blue
        Me.Controls.Add(_statusLabel)

        ' 添加左侧连接列表
        Dim connectionListLabel = New Label()
        connectionListLabel.Text = "已保存的连接:"
        connectionListLabel.Location = New Point(10, 10)
        connectionListLabel.Width = 180
        Me.Controls.Add(connectionListLabel)

        _connectionsListView = New ListView()
        _connectionsListView.Location = New Point(10, 30)
        _connectionsListView.Size = New Size(180, 450)
        _connectionsListView.View = View.Details
        _connectionsListView.FullRowSelect = True
        _connectionsListView.Columns.Add("连接名称", 120)
        _connectionsListView.Columns.Add("状态", 50)
        AddHandler _connectionsListView.SelectedIndexChanged, AddressOf ConnectionsListView_SelectedIndexChanged
        Me.Controls.Add(_connectionsListView)

        ' 添加/移除连接按钮
        _addConnectionButton = New Button()
        _addConnectionButton.Text = "添加"
        _addConnectionButton.Location = New Point(10, 490)
        _addConnectionButton.Width = 80
        AddHandler _addConnectionButton.Click, AddressOf AddConnectionButton_Click
        Me.Controls.Add(_addConnectionButton)

        _removeConnectionButton = New Button()
        _removeConnectionButton.Text = "移除"
        _removeConnectionButton.Location = New Point(100, 490)
        _removeConnectionButton.Width = 80
        AddHandler _removeConnectionButton.Click, AddressOf RemoveConnectionButton_Click
        Me.Controls.Add(_removeConnectionButton)

        ' 选项卡控件
        _tabControl = New TabControl()
        _tabControl.Location = New Point(210, 130)
        _tabControl.Size = New Size(670, 370)
        Me.Controls.Add(_tabControl)

        ' 连接配置区域
        CreateConnectionConfigArea()

        ' 创建各个选项卡
        CreateToolsTab()
        CreateResourcesTab()
        CreatePromptsTab()
        CreateTestTab()

        ' 底部按钮
        CreateBottomButtons()

        ' 加载保存的连接到列表
        LoadConnectionsList()
        SetupConnectionsContextMenu()
    End Sub

    Private _connectionDescriptionTextBox As TextBox

    Private Sub CreateConnectionConfigArea()
        ' 连接名称输入框
        Dim nameLabel As New Label()
        nameLabel.Text = "连接名称:"
        nameLabel.Location = New Point(210, 40)
        nameLabel.Width = 80
        Me.Controls.Add(nameLabel)

        _connectionNameTextBox = New TextBox()
        _connectionNameTextBox.Location = New Point(300, 37)
        _connectionNameTextBox.Width = 180  ' 缩短宽度从300改为180
        Me.Controls.Add(_connectionNameTextBox)

        ' 连接类型选择
        Dim typeLabel As New Label()
        typeLabel.Text = "连接类型:"
        typeLabel.Location = New Point(210, 70)
        typeLabel.Width = 80
        Me.Controls.Add(typeLabel)

        Dim typeCombo = New ComboBox()
        typeCombo.Location = New Point(300, 67)
        typeCombo.Width = 150
        typeCombo.DropDownStyle = ComboBoxStyle.DropDownList
        typeCombo.Items.AddRange({"HTTP/SSE", "Stdio (本地进程)"})
        typeCombo.SelectedIndex = 0
        AddHandler typeCombo.SelectedIndexChanged, AddressOf TypeCombo_SelectedIndexChanged
        Me.Controls.Add(typeCombo)

        ' 服务器URL/命令
        Dim serverUrlLabel As New Label()
        serverUrlLabel.Text = "服务器URL:"
        serverUrlLabel.Location = New Point(210, 100)
        serverUrlLabel.Width = 80
        Me.Controls.Add(serverUrlLabel)

        _serverUrlTextBox = New TextBox()
        _serverUrlTextBox.Location = New Point(300, 97)
        _serverUrlTextBox.Width = 400
        _serverUrlTextBox.Text = "http://localhost:3000"
        Me.Controls.Add(_serverUrlTextBox)

        ' 删除预设按钮代码块

        ' 高级设置按钮 - 调整位置，填补预设按钮的空缺
        _advancedButton = New Button()
        _advancedButton.Text = "高级设置"
        _advancedButton.Location = New Point(710, 67)  ' 移到预设按钮的位置
        _advancedButton.Width = 170  ' 增加宽度
        _advancedButton.Enabled = False ' 默认禁用（HTTP模式）
        AddHandler _advancedButton.Click, AddressOf AdvancedButton_Click
        Me.Controls.Add(_advancedButton)

        ' 测试连接按钮
        _testButton = New Button()
        _testButton.Text = "连接并探索"
        _testButton.Location = New Point(710, 97)
        _testButton.Width = 170
        AddHandler _testButton.Click, AddressOf TestConnectionAsync
        Me.Controls.Add(_testButton)
    End Sub

    ' 底部按钮移动位置
    Private Sub CreateBottomButtons()
        ' 添加导入配置按钮
        Dim importConfigButton As New Button()
        importConfigButton.Text = "导入配置"
        importConfigButton.Location = New Point(570, 520)
        importConfigButton.Width = 100
        AddHandler importConfigButton.Click, AddressOf ImportConfigButton_Click
        Me.Controls.Add(importConfigButton)

        _saveButton = New Button()
        _saveButton.Text = "保存配置"
        _saveButton.Location = New Point(680, 520)
        _saveButton.Width = 100
        AddHandler _saveButton.Click, AddressOf SaveButton_Click
        Me.Controls.Add(_saveButton)

        _cancelButton = New Button()
        _cancelButton.Text = "取消"
        _cancelButton.Location = New Point(790, 520)
        _cancelButton.Width = 100
        AddHandler _cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(_cancelButton)
    End Sub
    Private Sub ImportConfigButton_Click(sender As Object, e As EventArgs)
        ' 显示导入配置对话框
        Using importForm As New ImportConfigForm()
            If importForm.ShowDialog() = DialogResult.OK Then
                Try
                    ' 获取用户输入的JSON配置
                    Dim jsonConfig = importForm.ConfigJson

                    ' 导入并合并配置
                    Dim importedCount = MCPConnectionManager.ImportAndMergeConfig(jsonConfig)

                    If importedCount > 0 Then
                        ' 重新加载连接列表
                        _currentConnections = MCPConnectionManager.LoadConnections()
                        LoadConnectionsList()

                        MessageBox.Show($"成功导入并合并了 {importedCount} 个连接配置", "导入成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("没有找到有效的MCP服务器配置", "导入提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Catch ex As Exception
                    MessageBox.Show($"导入配置失败: {ex.Message}", "导入错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    ' 修改窗体加载事件，确保选择生效
    Private Sub MCPConfigForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' 如果列表有项目但没有选中项，强制触发选择事件
        If _connectionsListView.Items.Count > 0 AndAlso _connectionsListView.SelectedItems.Count > 0 Then
            ' 手动调用选择事件处理程序
            ConnectionsListView_SelectedIndexChanged(_connectionsListView, EventArgs.Empty)
        End If
    End Sub
    ' 添加新方法：加载连接列表
    Private Sub LoadConnectionsList()
        _connectionsListView.Items.Clear()

        For Each connection In _currentConnections
            Dim item = New ListViewItem(connection.Name)
            item.SubItems.Add(If(connection.Enabled, "启用", "禁用"))
            item.Tag = connection
            _connectionsListView.Items.Add(item)
        Next

        ' 如果列表有项目但没有选中项，自动选择第一个
        If _connectionsListView.Items.Count > 0 AndAlso _connectionsListView.SelectedItems.Count = 0 Then
            _connectionsListView.Items(0).Selected = True
            _connectionsListView.Select() ' 确保ListView获得焦点
        End If
    End Sub
    ' 修改 PresetButton_Click 方法

    Private Sub ConnectionsListView_SelectedIndexChanged(sender As Object, e As EventArgs)
        If _connectionsListView.SelectedItems.Count > 0 Then
            Dim selectedConnection = CType(_connectionsListView.SelectedItems(0).Tag, MCPConnectionConfig)

            ' 填充连接信息到表单
            _connectionNameTextBox.Text = selectedConnection.Name

            ' 设置连接类型（在设置URL之前）
            Dim connectionType = selectedConnection.ConnectionType
            Dim typeCombo As ComboBox = Nothing

            For Each ctrl As Control In Me.Controls
                If TypeOf ctrl Is ComboBox AndAlso ctrl.Location.Y = 67 Then
                    typeCombo = CType(ctrl, ComboBox)
                    If connectionType.Equals("Stdio", StringComparison.OrdinalIgnoreCase) Then
                        typeCombo.SelectedIndex = 1  ' Stdio 模式
                        _advancedButton.Enabled = True ' 启用高级设置按钮
                    Else
                        typeCombo.SelectedIndex = 0  ' HTTP 模式
                        _advancedButton.Enabled = False ' 禁用高级设置按钮
                    End If
                    Exit For
                End If
            Next

            ' 设置URL
            _serverUrlTextBox.Text = selectedConnection.Url

            ' 保存当前连接名称
            _currentConnectionName = selectedConnection.Name

            ' 加载连接中保存的工具信息
            LoadToolsFromConnection(selectedConnection)

            ' 立即应用更改
            Application.DoEvents()
        End If
    End Sub
    ' 从保存的连接中加载工具信息
    Private Sub LoadToolsFromConnection(connection As MCPConnectionConfig)
        ' 清除当前工具列表
        _currentTools.Clear()
        _toolsListView.Items.Clear()
        _testToolCombo.Items.Clear()

        ' 检查连接中是否有保存的工具
        If connection.Tools IsNot Nothing AndAlso connection.Tools.Count > 0 Then
            ' 将保存的工具信息转换为MCPToolInfo对象并添加到列表
            For Each toolJson In connection.Tools
                Try
                    Dim tool As New MCPToolInfo()

                    ' 从function格式提取信息
                    If toolJson("type")?.ToString() = "function" AndAlso toolJson("function") IsNot Nothing Then
                        Dim functionObj = toolJson("function")

                        ' 提取工具名称和描述
                        tool.Name = functionObj("name")?.ToString()
                        tool.Description = functionObj("description")?.ToString()

                        ' 提取参数架构
                        If functionObj("parameters") IsNot Nothing Then
                            tool.InputSchema = functionObj("parameters")
                        End If

                        ' 添加到工具列表
                        _currentTools.Add(tool)

                        ' 添加到ListView
                        Dim item = New ListViewItem(tool.Name)
                        item.SubItems.Add(If(tool.Description Is Nothing, "", tool.Description))
                        item.SubItems.Add(If(tool.InputSchema IsNot Nothing, "是", "否"))
                        item.Tag = tool
                        _toolsListView.Items.Add(item)

                        ' 添加到测试工具下拉列表
                        _testToolCombo.Items.Add(tool.Name)
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"加载工具时出错: {ex.Message}")
                End Try
            Next

            ' 如果有工具，选择第一个
            If _testToolCombo.Items.Count > 0 Then
                _testToolCombo.SelectedIndex = 0
            End If

            UpdateStatus($"已从保存的连接加载 {_currentTools.Count} 个工具", Color.Green)
        Else
            UpdateStatus("此连接没有保存的工具信息，需要重新连接服务器进行探索", Color.Blue)
        End If
    End Sub

    ' 添加新方法：添加连接按钮处理
    Private Sub AddConnectionButton_Click(sender As Object, e As EventArgs)
        ' 清空表单以创建新连接
        _connectionNameTextBox.Text = "新连接_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
        _serverUrlTextBox.Text = "http://localhost:3000"

        ' 设置为HTTP模式
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is ComboBox AndAlso ctrl.Location.Y = 67 Then
                Dim combo = CType(ctrl, ComboBox)
                combo.SelectedIndex = 0  ' HTTP 模式
                Exit For
            End If
        Next

        ' 清除当前连接名称
        _currentConnectionName = String.Empty
    End Sub

    ' 添加新方法：移除连接按钮处理
    Private Sub RemoveConnectionButton_Click(sender As Object, e As EventArgs)
        If _connectionsListView.SelectedItems.Count > 0 Then
            Dim selectedConnection = CType(_connectionsListView.SelectedItems(0).Tag, MCPConnectionConfig)

            If MessageBox.Show($"确定要移除连接 '{selectedConnection.Name}' 吗?", "确认移除",
                              MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                ' 移除连接
                _currentConnections = MCPConnectionManager.RemoveConnection(_currentConnections, selectedConnection.Name)

                ' 重新加载列表
                LoadConnectionsList()

                ' 清空表单
                _connectionNameTextBox.Text = String.Empty
                _serverUrlTextBox.Text = "http://localhost:3000"

                ' 清除当前连接名称
                _currentConnectionName = String.Empty
            End If
        Else
            MessageBox.Show("请先选择要移除的连接", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' 添加高级设置对话框处理
    Private Sub AdvancedButton_Click(sender As Object, e As EventArgs)
        ' 如果按钮被禁用，直接返回
        If Not _advancedButton.Enabled Then Return

        Dim isStdio = _serverUrlTextBox.Text.StartsWith("stdio://") OrElse
         (_serverUrlTextBox.Text.Contains("高级设置") AndAlso Not _serverUrlTextBox.Text.StartsWith("http"))

        If isStdio Then
            ' 创建新的StdioOptions对象
            Dim options As New StdioOptions()

            ' 如果有选中的连接，尝试从连接配置加载设置
            If Not String.IsNullOrEmpty(_currentConnectionName) AndAlso
           _currentConnections.Exists(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase)) Then

                Dim connection = _currentConnections.Find(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase))

                If _serverUrlTextBox.Text.StartsWith("stdio://") Then
                    ' 从URL解析基本选项
                    options = StdioOptions.Parse(_serverUrlTextBox.Text)

                    ' 命令和参数从URL解析，但环境变量从连接配置加载（更可靠）
                    options.EnvironmentVariables.Clear()

                    ' 复制环境变量
                    For Each kvp In connection.EnvironmentVariables
                        options.EnvironmentVariables.Add(kvp.Key, kvp.Value)
                    Next
                Else
                    ' 创建新选项，使用基本默认值
                    options.Command = "node"
                    options.Arguments = "-r ts-node/register src/server.ts"

                    ' 加载已保存的环境变量
                    For Each kvp In connection.EnvironmentVariables
                        options.EnvironmentVariables.Add(kvp.Key, kvp.Value)
                    Next
                End If
            Else
                ' 新连接，使用基本设置或从URL解析
                If _serverUrlTextBox.Text.StartsWith("stdio://") Then
                    options = StdioOptions.Parse(_serverUrlTextBox.Text)
                Else
                    options.Command = "node"
                    options.Arguments = "-r ts-node/register src/server.ts"
                End If
            End If

            ' 显示Stdio配置对话框
            Using stdioForm As New StdioConfigForm(options)
                If stdioForm.ShowDialog() = DialogResult.OK Then
                    ' 更新文本框显示
                    _serverUrlTextBox.Text = stdioForm.Options.ToUrl()

                    ' 如果有选中的连接，立即更新其环境变量
                    If Not String.IsNullOrEmpty(_currentConnectionName) AndAlso
                   _currentConnections.Exists(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase)) Then

                        Dim connection = _currentConnections.Find(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase))
                        connection.EnvironmentVariables.Clear()

                        ' 复制环境变量
                        For Each kvp In stdioForm.Options.EnvironmentVariables
                            connection.EnvironmentVariables.Add(kvp.Key, kvp.Value)
                        Next

                        ' 立即保存到文件
                        MCPConnectionManager.SaveConnections(_currentConnections)
                    End If
                End If
            End Using
        Else
            ' HTTP 模式下的设置
            MessageBox.Show("HTTP/SSE 模式不需要额外配置。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ImportOfficialConfig_Click(sender As Object, e As EventArgs)
        ' 直接导入官方格式，无需转换
        Using openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "JSON文件|*.json"
            openFileDialog.Title = "选择官方MCP配置文件"

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Try
                    Dim json = File.ReadAllText(openFileDialog.FileName)
                    Dim config = JObject.Parse(json)

                    If config("mcpServers") IsNot Nothing Then
                        Dim imported = 0
                        Dim serversObj = config("mcpServers").ToObject(Of JObject)()

                        For Each server In serversObj.Properties()
                            Dim serverId = server.Name
                            Dim serverConfig = server.Value.ToObject(Of MCPConnectionConfig)()

                            ' 如果名称为空，使用ID作为名称
                            If String.IsNullOrEmpty(serverConfig.Name) Then
                                serverConfig.Name = serverId
                            End If

                            ' 检查是否已存在同名连接
                            Dim existingIndex = _currentConnections.FindIndex(Function(c) c.Name.Equals(serverConfig.Name, StringComparison.OrdinalIgnoreCase))
                            If existingIndex >= 0 Then
                                _currentConnections(existingIndex) = serverConfig
                            Else
                                _currentConnections.Add(serverConfig)
                            End If

                            imported += 1
                        Next

                        ' 保存并刷新列表
                        MCPConnectionManager.SaveConnections(_currentConnections)
                        LoadConnectionsList()

                        MessageBox.Show($"成功导入 {imported} 个连接", "导入成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("无效的MCP配置文件格式", "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Catch ex As Exception
                    MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    ' 更新类型选择处理
    Private Sub TypeCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim combo = CType(sender, ComboBox)
        Select Case combo.SelectedIndex
            Case 0 ' HTTP/SSE
                _serverUrlTextBox.ReadOnly = False
                _serverUrlTextBox.BackColor = SystemColors.Window
                _serverUrlTextBox.Text = "http://localhost:3000"
                _statusLabel.Text = "HTTP/SSE 模式: 输入服务器URL即可连接"
                _advancedButton.Enabled = False ' 禁用高级设置按钮
            Case 1 ' Stdio
                ' 使服务器URL文本框只读，突出显示需要配置
                _serverUrlTextBox.ReadOnly = True
                _serverUrlTextBox.BackColor = Color.LightYellow
                _serverUrlTextBox.Text = "【请点击右侧""高级设置""按钮配置本地进程】"
                _statusLabel.Text = "Stdio 模式: 需要通过高级设置配置本地进程"
                _advancedButton.Enabled = True ' 启用高级设置按钮
        End Select
    End Sub

    ' 更新LoadConfig方法
    Private Sub LoadConfig()
        _serverUrlTextBox.Text = "http://localhost:3000"
    End Sub

    Private Sub CreateToolsTab()
        Dim toolsTab = New TabPage("工具(Tools)")

        ' 创建分割容器 - 不使用 DockStyle.Fill
        Dim splitter = New SplitContainer()
        splitter.Location = New Point(0, 0)
        splitter.Size = New Size(660, 340)  ' 手动设置尺寸，留出边距
        splitter.Orientation = Orientation.Vertical
        splitter.Panel1MinSize = 100   ' 减小最小尺寸
        splitter.Panel2MinSize = 100
        splitter.SplitterWidth = 5
        splitter.FixedPanel = FixedPanel.None
        splitter.SplitterDistance = 396  ' 直接设置为60% (660 * 0.6)

        ' 左侧工具列表
        _toolsListView = New ListView()
        _toolsListView.Dock = DockStyle.Fill
        _toolsListView.View = View.Details
        _toolsListView.FullRowSelect = True
        _toolsListView.HideSelection = False
        _toolsListView.Columns.Add("工具名称", 150)
        _toolsListView.Columns.Add("描述", 200)
        _toolsListView.Columns.Add("有输入架构", 80)

        ' 工具详情富文本框
        Dim detailsBox = New RichTextBox()
        detailsBox.Dock = DockStyle.Fill
        detailsBox.ReadOnly = True
        detailsBox.BackColor = SystemColors.Window
        detailsBox.Font = New Font("Consolas", 9)

        ' 添加工具选择事件处理程序
        AddHandler _toolsListView.SelectedIndexChanged, Sub(sender, e)
                                                            If _toolsListView.SelectedItems.Count > 0 Then
                                                                Dim selectedTool = CType(_toolsListView.SelectedItems(0).Tag, MCPToolInfo)
                                                                ShowToolDetails(selectedTool, detailsBox)
                                                            End If
                                                        End Sub

        AddHandler _toolsListView.DoubleClick, AddressOf ToolsListView_DoubleClick

        ' 添加控件到分割容器
        splitter.Panel1.Controls.Add(_toolsListView)
        splitter.Panel2.Controls.Add(detailsBox)

        toolsTab.Controls.Add(splitter)
        _tabControl.TabPages.Add(toolsTab)
    End Sub

    ' 新增方法：显示工具详情
    Private Sub ShowToolDetails(tool As MCPToolInfo, detailsBox As RichTextBox)
        Try
            ' 创建详细描述文本
            Dim sb As New StringBuilder()

            sb.AppendLine("【工具名称】")
            sb.AppendLine(tool.Name)
            sb.AppendLine()

            sb.AppendLine("【工具描述】")
            sb.AppendLine(If(String.IsNullOrEmpty(tool.Description), "(无描述)", tool.Description))
            sb.AppendLine()

            ' 如果有输入架构，显示详细参数信息
            If tool.InputSchema IsNot Nothing Then
                sb.AppendLine("【参数架构】")

                Dim schema = tool.InputSchema
                Dim schemaObj As JObject = Nothing

                ' 尝试将schema转换为JObject以便访问其属性
                If TypeOf schema Is JObject Then
                    schemaObj = DirectCast(schema, JObject)
                ElseIf TypeOf schema Is String Then
                    Try
                        schemaObj = JObject.Parse(schema.ToString())
                    Catch
                        ' 解析失败时直接显示原始字符串
                        sb.AppendLine(schema.ToString())
                    End Try
                End If

                ' 显示架构详情
                If schemaObj IsNot Nothing Then
                    ' 提取类型
                    sb.AppendLine($"类型: {If(schemaObj("type") IsNot Nothing, schemaObj("type").ToString(), "Object")}")

                    ' 如果有properties属性，说明是对象类型
                    If schemaObj("properties") IsNot Nothing Then
                        sb.AppendLine("")
                        sb.AppendLine("参数列表:")

                        Dim props = schemaObj("properties")
                        ' 正确的方式是先检查元素是否存在
                        Dim required As New List(Of String)()
                        If schemaObj("required") IsNot Nothing Then
                            Try
                                required = schemaObj("required").ToObject(Of List(Of String))
                            Catch ex As Exception
                                ' 可能的转换错误，使用空列表
                                Debug.WriteLine($"无法转换required数组: {ex.Message}")
                            End Try
                        End If

                        ' 将 StringBuilder 的内容设置到 RichTextBox
                        detailsBox.Clear()
                        detailsBox.AppendText(sb.ToString())

                        ' 遍历所有属性
                        For Each prop In props.Children(Of JProperty)()
                            Dim propName = prop.Name
                            Dim propObj = prop.Value

                            ' 参数名称和是否必需
                            Dim paramText As String = $"· {propName}"
                            detailsBox.AppendText(paramText)

                            If required.Contains(propName) Then
                                ' 添加"(必需)"，并设置为红色
                                detailsBox.SelectionStart = detailsBox.TextLength
                                detailsBox.SelectionColor = Color.Red
                                detailsBox.AppendText(" (必需)")
                                detailsBox.SelectionColor = detailsBox.ForeColor ' 恢复默认颜色
                            End If

                            detailsBox.AppendText(Environment.NewLine)

                            ' 参数类型
                            Dim propType = If(propObj("type") IsNot Nothing, propObj("type").ToString(), "any")

                            ' 参数描述
                            Dim propDesc = propObj("description")?.ToString()
                            If Not String.IsNullOrEmpty(propDesc) Then
                                detailsBox.AppendText($"  描述: {propDesc}{Environment.NewLine}")
                            End If

                            ' 枚举值
                            If propObj("enum") IsNot Nothing Then
                                detailsBox.AppendText($"  可选值: ")
                                Dim enumVals = propObj("enum").ToObject(Of List(Of Object))()
                                detailsBox.AppendText(String.Join(", ", enumVals) & Environment.NewLine)
                            End If

                            ' 默认值
                            If propObj("default") IsNot Nothing Then
                                detailsBox.AppendText($"  默认值: {propObj("default")}{Environment.NewLine}")
                            End If

                            detailsBox.AppendText(Environment.NewLine)
                        Next

                        ' 如果有示例，显示示例
                        If schemaObj("examples") IsNot Nothing Then
                            detailsBox.AppendText(Environment.NewLine & "示例:" & Environment.NewLine)
                            For Each example In schemaObj("examples")
                                detailsBox.AppendText(example.ToString(Newtonsoft.Json.Formatting.Indented) & Environment.NewLine)
                            Next
                        End If

                        ' 已经直接操作了 RichTextBox，所以不需要 return
                        Return
                    End If

                    ' 如果有示例，显示示例
                    If schemaObj("examples") IsNot Nothing Then
                        sb.AppendLine("")
                        sb.AppendLine("示例:")
                        For Each example In schemaObj("examples")
                            sb.AppendLine(example.ToString(Newtonsoft.Json.Formatting.Indented))
                        Next
                    End If
                Else
                    ' 如果无法解析为JObject，直接显示序列化的JSON
                    sb.AppendLine(JsonConvert.SerializeObject(tool.InputSchema, Newtonsoft.Json.Formatting.Indented))
                End If
            Else
                sb.AppendLine("【参数架构】")
                sb.AppendLine("(无参数架构)")
            End If

            ' 如果前面没有直接操作 RichTextBox，则在这里设置文本
            detailsBox.Text = sb.ToString()
            detailsBox.SelectionStart = 0
            detailsBox.ScrollToCaret()
        Catch ex As Exception
            detailsBox.Text = $"无法加载工具详情: {ex.Message}"
        End Try
    End Sub

    Private Sub CreateResourcesTab()
        Dim resourcesTab = New TabPage("资源(Resources)")

        ' 创建分割容器 - 不使用 DockStyle.Fill
        Dim splitter = New SplitContainer()
        splitter.Location = New Point(0, 0)
        splitter.Size = New Size(660, 340)  ' 手动设置尺寸
        splitter.Orientation = Orientation.Vertical
        splitter.Panel1MinSize = 100
        splitter.Panel2MinSize = 100
        splitter.SplitterWidth = 5
        splitter.FixedPanel = FixedPanel.None
        splitter.SplitterDistance = 396  ' 60%

        ' 左侧资源列表
        _resourcesListView = New ListView()
        _resourcesListView.Dock = DockStyle.Fill
        _resourcesListView.View = View.Details
        _resourcesListView.FullRowSelect = True
        _resourcesListView.Columns.Add("URI", 180)
        _resourcesListView.Columns.Add("名称", 120)
        _resourcesListView.Columns.Add("MIME类型", 100)

        ' 右侧详情面板
        Dim detailsBox = New RichTextBox()
        detailsBox.Dock = DockStyle.Fill
        detailsBox.ReadOnly = True
        detailsBox.BackColor = SystemColors.Window
        detailsBox.Font = New Font("Consolas", 9)

        ' 添加选择事件处理程序
        AddHandler _resourcesListView.SelectedIndexChanged, Sub(sender, e)
                                                                If _resourcesListView.SelectedItems.Count > 0 Then
                                                                    Dim resource = CType(_resourcesListView.SelectedItems(0).Tag, MCPResourceInfo)
                                                                    detailsBox.Text = $"URI: {resource.Uri}{Environment.NewLine}{Environment.NewLine}" &
                                                                        $"名称: {resource.Name}{Environment.NewLine}{Environment.NewLine}" &
                                                                        $"描述: {resource.Description}{Environment.NewLine}{Environment.NewLine}" &
                                                                        $"MIME类型: {resource.MimeType}"
                                                                End If
                                                            End Sub

        AddHandler _resourcesListView.DoubleClick, AddressOf ResourcesListView_DoubleClick

        ' 添加控件到分割容器
        splitter.Panel1.Controls.Add(_resourcesListView)
        splitter.Panel2.Controls.Add(detailsBox)

        resourcesTab.Controls.Add(splitter)
        _tabControl.TabPages.Add(resourcesTab)
    End Sub

    Private Sub CreatePromptsTab()
        Dim promptsTab = New TabPage("提示(Prompts)")

        ' 创建分割容器 - 不使用 DockStyle.Fill
        Dim splitter = New SplitContainer()
        splitter.Location = New Point(0, 0)
        splitter.Size = New Size(660, 340)  ' 手动设置尺寸
        splitter.Orientation = Orientation.Vertical
        splitter.Panel1MinSize = 100
        splitter.Panel2MinSize = 100
        splitter.SplitterWidth = 5
        splitter.FixedPanel = FixedPanel.None
        splitter.SplitterDistance = 396  ' 60%

        ' 左侧提示列表
        _promptsListView = New ListView()
        _promptsListView.Dock = DockStyle.Fill
        _promptsListView.View = View.Details
        _promptsListView.FullRowSelect = True
        _promptsListView.Columns.Add("提示名称", 180)
        _promptsListView.Columns.Add("描述", 200)
        _promptsListView.Columns.Add("参数数量", 80)

        ' 右侧详情面板
        Dim detailsBox = New RichTextBox()
        detailsBox.Dock = DockStyle.Fill
        detailsBox.ReadOnly = True
        detailsBox.BackColor = SystemColors.Window
        detailsBox.Font = New Font("Consolas", 9)

        ' 添加选择事件处理程序
        AddHandler _promptsListView.SelectedIndexChanged, Sub(sender, e)
                                                              If _promptsListView.SelectedItems.Count > 0 Then
                                                                  Dim prompt = CType(_promptsListView.SelectedItems(0).Tag, MCPPromptInfo)
                                                                  Dim sb = New StringBuilder()
                                                                  sb.AppendLine($"名称: {prompt.Name}")
                                                                  sb.AppendLine($"描述: {prompt.Description}")
                                                                  sb.AppendLine()

                                                                  If prompt.Arguments IsNot Nothing AndAlso prompt.Arguments.Count > 0 Then
                                                                      sb.AppendLine("参数列表:")
                                                                      For Each arg In prompt.Arguments
                                                                          sb.AppendLine($"· {arg.Name}")
                                                                          sb.AppendLine($"  描述: {arg.Description}")
                                                                          sb.AppendLine($"  必需: {arg.Required}")
                                                                          sb.AppendLine()
                                                                      Next
                                                                  Else
                                                                      sb.AppendLine("无参数")
                                                                  End If

                                                                  detailsBox.Text = sb.ToString()
                                                              End If
                                                          End Sub

        ' 添加控件到分割容器
        splitter.Panel1.Controls.Add(_promptsListView)
        splitter.Panel2.Controls.Add(detailsBox)

        promptsTab.Controls.Add(splitter)
        _tabControl.TabPages.Add(promptsTab)
    End Sub

    Private Sub CreateTestTab()
        Dim testTab = New TabPage("测试工具")

        ' 工具选择
        Dim toolLabel = New Label()
        toolLabel.Text = "选择工具:"
        toolLabel.Location = New Point(10, 10)
        toolLabel.Width = 80
        testTab.Controls.Add(toolLabel)

        _testToolCombo = New ComboBox()
        _testToolCombo.Location = New Point(100, 7)
        _testToolCombo.Width = 450  ' 增加宽度从300改为450，以便显示更长的工具名称
        _testToolCombo.DropDownStyle = ComboBoxStyle.DropDownList
        AddHandler _testToolCombo.SelectedIndexChanged, AddressOf TestToolCombo_SelectedIndexChanged
        testTab.Controls.Add(_testToolCombo)

        _executeTestButton = New Button()
        _executeTestButton.Text = "执行"
        _executeTestButton.Location = New Point(560, 5)  ' 调整位置
        _executeTestButton.Width = 90  ' 略微增加按钮宽度
        AddHandler _executeTestButton.Click, AddressOf ExecuteTestAsync
        testTab.Controls.Add(_executeTestButton)

        ' 参数输入
        Dim paramsLabel = New Label()
        paramsLabel.Text = "参数(JSON):"
        paramsLabel.Location = New Point(10, 40)
        paramsLabel.Width = 100
        testTab.Controls.Add(paramsLabel)

        _testParametersTextBox = New TextBox()
        _testParametersTextBox.Location = New Point(10, 60)
        _testParametersTextBox.Size = New Size(650, 100)  ' 增加宽度
        _testParametersTextBox.Multiline = True
        _testParametersTextBox.ScrollBars = ScrollBars.Both
        _testParametersTextBox.Text = "{""location"": ""北京""}"
        _testParametersTextBox.Font = New Font("Consolas", 9)  ' 使用等宽字体，方便编辑JSON
        testTab.Controls.Add(_testParametersTextBox)

        ' 结果显示
        Dim resultLabel = New Label()
        resultLabel.Text = "执行结果:"
        resultLabel.Location = New Point(10, 170)
        resultLabel.Width = 100
        testTab.Controls.Add(resultLabel)

        _testResultTextBox = New TextBox()
        _testResultTextBox.Location = New Point(10, 190)
        _testResultTextBox.Size = New Size(650, 170)  ' 增加宽度
        _testResultTextBox.Multiline = True
        _testResultTextBox.ScrollBars = ScrollBars.Both
        _testResultTextBox.ReadOnly = True
        _testResultTextBox.Font = New Font("Consolas", 9)  ' 使用等宽字体，方便查看JSON
        testTab.Controls.Add(_testResultTextBox)

        _tabControl.TabPages.Add(testTab)
    End Sub

    ' 添加工具选择变更事件处理
    Private Sub TestToolCombo_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            If _testToolCombo.SelectedItem Is Nothing Then Return

            Dim selectedToolName = _testToolCombo.SelectedItem.ToString()

            ' 查找选中的工具信息
            Dim selectedTool As MCPToolInfo = Nothing
            For Each tool In _currentTools
                If tool.Name = selectedToolName Then
                    selectedTool = tool
                    Exit For
                End If
            Next

            If selectedTool IsNot Nothing Then
                ' 生成参数模板
                Dim paramTemplate = GenerateParameterTemplate(selectedTool)
                _testParametersTextBox.Text = paramTemplate
            End If
        Catch ex As Exception
            Debug.WriteLine($"生成参数模板时出错: {ex.Message}")
        End Try
    End Sub


    ' 修改参数模板生成方法，只包含必需参数
    Private Function GenerateParameterTemplate(tool As MCPToolInfo) As String
        Try
            If tool.InputSchema Is Nothing Then
                Return "{}"
            End If

            Dim schema = tool.InputSchema
            Dim schemaObj As JObject = Nothing

            ' 解析JSON架构
            If TypeOf schema Is JObject Then
                schemaObj = DirectCast(schema, JObject)
            ElseIf TypeOf schema Is String Then
                Try
                    schemaObj = JObject.Parse(schema.ToString())
                Catch
                    Return "{}"
                End Try
            Else
                Return "{}"
            End If

            ' 创建参数模板对象
            Dim paramObj As New JObject()

            ' 获取必需参数列表
            Dim requiredProps As New List(Of String)()
            If schemaObj("required") IsNot Nothing Then
                Try
                    requiredProps = schemaObj("required").ToObject(Of List(Of String))()
                    'Debug.WriteLine($"必需参数: {String.Join(", ", requiredProps)}")
                Catch ex As Exception
                    Debug.WriteLine($"解析required数组失败: {ex.Message}")
                End Try
            End If

            ' 如果存在properties属性
            If schemaObj("properties") IsNot Nothing Then
                Dim props = schemaObj("properties")

                ' 遍历所有属性
                For Each prop In props.Children(Of JProperty)()
                    Dim propName = prop.Name
                    Dim propObj = prop.Value

                    ' 只添加必需参数到模板中
                    If requiredProps.Contains(propName) Then
                        ' 添加属性到参数模板
                        Dim defaultValue As JToken = Nothing

                        ' 检查是否有默认值或示例
                        If propObj("default") IsNot Nothing Then
                            defaultValue = propObj("default")
                        ElseIf propObj("examples") IsNot Nothing AndAlso propObj("examples").Count > 0 Then
                            defaultValue = propObj("examples")(0)
                        Else
                            ' 根据类型生成默认值
                            Dim propType = If(propObj("type") IsNot Nothing, propObj("type").ToString(), "string")
                            Select Case propType.ToLower()
                                Case "string"
                                    defaultValue = "示例值"
                                    ' 尝试从描述中提取合适的示例
                                    Dim desc = propObj("description")?.ToString()
                                    If Not String.IsNullOrEmpty(desc) Then
                                        If desc.Contains("例如") OrElse desc.Contains("示例") Then
                                            Dim exStart = Math.Max(desc.IndexOf("例如"), desc.IndexOf("示例"))
                                            If exStart > 0 Then
                                                Dim exEnd = desc.IndexOf("。", exStart)
                                                If exEnd > exStart Then
                                                    defaultValue = desc.Substring(exStart, exEnd - exStart)
                                                End If
                                            End If
                                        End If

                                        ' 特殊参数示例
                                        If propName.ToLower().Contains("location") Then
                                            defaultValue = "北京"
                                        ElseIf propName.ToLower().Contains("address") Then
                                            defaultValue = "北京市海淀区上地十街10号"
                                        ElseIf propName.ToLower().Contains("query") Then
                                            defaultValue = "天安门"
                                        ElseIf propName.ToLower().Contains("origin") Then
                                            defaultValue = "北京站"
                                        ElseIf propName.ToLower().Contains("destination") Then
                                            defaultValue = "北京西站"
                                        End If
                                    End If
                                Case "number", "integer"
                                    defaultValue = 0
                                Case "boolean"
                                    defaultValue = False
                                Case "array"
                                    defaultValue = New JArray()
                                Case "object"
                                    defaultValue = New JObject()
                                Case Else
                                    defaultValue = Nothing
                            End Select
                        End If

                        ' 添加到参数对象
                        If defaultValue IsNot Nothing Then
                            paramObj(propName) = defaultValue
                        End If
                    End If
                Next
            End If

            ' 如果没有必需参数，添加一个提示
            If paramObj.Count = 0 AndAlso Not requiredProps.Any() Then
                ' 检查是否有任何可选参数
                If schemaObj("properties") IsNot Nothing AndAlso schemaObj("properties").Count() > 0 Then
                    ' 添加注释提示用户这个工具有可选参数但没有必需参数
                    paramObj("_注释") = "此工具没有必需参数，但有可选参数。请根据需要添加。"
                End If
            End If

            ' 格式化为可读的JSON
            Return JsonConvert.SerializeObject(paramObj, Formatting.Indented)
        Catch ex As Exception
            Debug.WriteLine($"生成参数模板失败: {ex.Message}")
            Return "{}"
        End Try
    End Function
    ' 将 MCP 工具转换为大模型函数调用格式
    Private Function ConvertToolToLLMFunction(tool As MCPToolInfo) As JObject
        Dim functionObj = New JObject()

        ' 设置工具类型为 function
        functionObj("type") = "function"

        ' 创建 function 对象
        Dim functionData = New JObject()
        functionData("name") = tool.Name
        functionData("description") = If(String.IsNullOrEmpty(tool.Description), $"MCP tool: {tool.Name}", tool.Description)

        ' 创建参数架构
        Dim parameters = New JObject()
        parameters("type") = "object"

        ' 创建属性对象
        Dim properties = New JObject()

        ' 创建 required 数组
        Dim required = New JArray()

        ' 从 InputSchema 解析参数
        If tool.InputSchema IsNot Nothing Then
            Dim schema As JObject = Nothing

            ' 尝试将 InputSchema 转换为 JObject
            If TypeOf tool.InputSchema Is JObject Then
                schema = DirectCast(tool.InputSchema, JObject)
            ElseIf TypeOf tool.InputSchema Is String Then
                Try
                    schema = JObject.Parse(tool.InputSchema.ToString())
                Catch
                    ' 解析失败，创建默认架构
                    properties("input") = New JObject()
                    properties("input")("type") = "string"
                    properties("input")("description") = "Input for the tool"
                    required.Add("input")
                End Try
            End If

            If schema IsNot Nothing Then
                ' 检查是否有 properties
                If schema("properties") IsNot Nothing Then
                    ' 复制属性
                    properties = schema("properties").DeepClone().ToObject(Of JObject)()

                    ' 获取必需参数
                    If schema("required") IsNot Nothing Then
                        required = schema("required").DeepClone().ToObject(Of JArray)()
                    End If
                Else
                    ' 没有 properties，尝试创建默认参数
                    properties("input") = New JObject()
                    properties("input")("type") = "string"
                    properties("input")("description") = "Input for the tool"
                    required.Add("input")
                End If
            End If
        Else
            ' 没有 InputSchema，创建默认参数
            properties("input") = New JObject()
            properties("input")("type") = "string"
            properties("input")("description") = "Input for the tool"
            required.Add("input")
        End If

        ' 设置参数架构
        parameters("properties") = properties
        parameters("required") = required

        ' 设置函数参数
        functionData("parameters") = parameters

        ' 设置函数对象
        functionObj("function") = functionData

        Return functionObj
    End Function
    Private Sub UpdateToolsList()
        _toolsListView.Items.Clear()
        _testToolCombo.Items.Clear()

        For Each tool In _currentTools
            Dim item = New ListViewItem(tool.Name)
            item.SubItems.Add(If(tool.Description Is Nothing, "", tool.Description))
            item.SubItems.Add(If(tool.InputSchema IsNot Nothing, "是", "否"))
            item.Tag = tool
            _toolsListView.Items.Add(item)

            _testToolCombo.Items.Add(tool.Name)
        Next

        If _testToolCombo.Items.Count > 0 Then
            _testToolCombo.SelectedIndex = 0
        End If

    End Sub

    Private Sub UpdateResourcesList(resources As List(Of MCPResourceInfo))
        _resourcesListView.Items.Clear()

        For Each resource In resources
            Dim item = New ListViewItem(resource.Uri)
            item.SubItems.Add(If(resource.Name Is Nothing, "", resource.Name))
            item.SubItems.Add(If(resource.Description Is Nothing, "", resource.Description))
            item.SubItems.Add(If(resource.MimeType Is Nothing, "", resource.MimeType))
            item.Tag = resource
            _resourcesListView.Items.Add(item)
        Next
    End Sub

    Private Sub UpdatePromptsList(prompts As List(Of MCPPromptInfo))
        _promptsListView.Items.Clear()

        For Each prompt In prompts
            Dim item = New ListViewItem(prompt.Name)
            item.SubItems.Add(If(prompt.Description Is Nothing, "", prompt.Description))
            item.SubItems.Add($"{If(prompt.Arguments Is Nothing, 0, prompt.Arguments.Count)}")
            item.Tag = prompt
            _promptsListView.Items.Add(item)
        Next
    End Sub
    ' 修改 LoadServerCapabilitiesAsync 方法
    Private Async Function LoadServerCapabilitiesAsync() As Task
        Try
            UpdateStatus("正在加载服务器功能...", Color.Blue)

            ' 初始化工具集合，避免空引用
            If _currentTools Is Nothing Then
                _currentTools = New List(Of MCPToolInfo)()
            End If

            ' 加载工具
            If _mcpClient.ServerCapabilities?.Tools Then
                Try
                    _currentTools = Await _mcpClient.ListToolsAsync()
                    UpdateToolsList()
                Catch toolEx As Exception
                    Debug.WriteLine($"加载工具失败: {toolEx.Message}")
                    ' 失败时初始化为空列表，避免后续引用错误
                    _currentTools = New List(Of MCPToolInfo)()
                End Try
            End If

            ' 加载资源
            If _mcpClient.ServerCapabilities?.Resources Then
                Try
                    Dim resources = Await _mcpClient.ListResourcesAsync()
                    UpdateResourcesList(resources)
                Catch resEx As Exception
                    Debug.WriteLine($"加载资源失败: {resEx.Message}")
                End Try
            End If

            ' 加载提示
            If _mcpClient.ServerCapabilities?.Prompts Then
                Try
                    Dim prompts = Await _mcpClient.ListPromptsAsync()
                    UpdatePromptsList(prompts)
                Catch promptEx As Exception
                    Debug.WriteLine($"加载提示失败: {promptEx.Message}")
                End Try
            End If

            UpdateStatus($"服务器功能加载完成！工具: {_currentTools.Count}个", Color.Green)

        Catch ex As Exception
            UpdateStatus($"加载服务器功能失败: {ex.Message}", Color.Red)
        End Try
    End Function


    Private Async Sub TestConnectionAsync(sender As Object, e As EventArgs)
        ' 验证连接名称
        If String.IsNullOrEmpty(_connectionNameTextBox.Text) Then
            UpdateStatus("请输入连接名称", Color.Red)
            Return
        End If

        If String.IsNullOrEmpty(_serverUrlTextBox.Text) OrElse
       (_serverUrlTextBox.Text.Contains("请点击") AndAlso _serverUrlTextBox.Text.Contains("高级设置")) Then
            UpdateStatus("请先配置连接信息", Color.Red)
            Return
        End If

        Try
            _testButton.Enabled = False
            UpdateStatus("正在连接MCP服务器...", Color.Blue)

            ' 不使用 API 密钥
            Dim apiKey As String = Nothing

            ' 使用新的配置方法
            Await _mcpClient.ConfigureAsync(_serverUrlTextBox.Text, apiKey)

            ' 初始化连接
            Dim result = Await _mcpClient.InitializeAsync()

            If Not result.Success Then
                UpdateStatus($"连接失败: {result.ErrorMessage}", Color.Red)
                Return
            End If

            ' 更新状态信息，显示详细的服务器信息
            Dim transportText = If(_mcpClient.TransportType = MCPTransportType.Stdio, "Stdio", "HTTP/SSE")
            Dim serverInfo = If(result.ServerInfo IsNot Nothing,
           $"{result.ServerInfo.Name} v{result.ServerInfo.Version}",
           "未知服务器")
            Dim protocolInfo = If(Not String.IsNullOrEmpty(result.ProtocolVersion),
            $"协议版本: {result.ProtocolVersion}",
            "")

            UpdateStatus($"连接成功！传输类型: {transportText}, 服务器: {serverInfo} {protocolInfo}", Color.Green)

            ' 加载服务器功能
            Await LoadServerCapabilitiesAsync()

            ' 保存或更新连接配置
            Dim connectionName = _connectionNameTextBox.Text.Trim()
            'Dim connectionDescription = _connectionDescriptionTextBox.Text.Trim()
            Dim connectionType = If(_mcpClient.TransportType = MCPTransportType.Stdio, "Stdio", "HTTP")

            ' 创建或更新连接配置
            Dim connection As MCPConnectionConfig



            If Not String.IsNullOrEmpty(_currentConnectionName) AndAlso
       _currentConnections.Exists(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase)) Then
                ' 更新现有连接
                connection = _currentConnections.Find(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase))
                connection.Name = connectionName
                'connection.Description = connectionDescription
                connection.Url = _serverUrlTextBox.Text
                ' 移除 connection.LastConnected = DateTime.Now

                ' 如果名称已变更，需要从列表中移除旧的
                If Not connectionName.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase) Then
                    _currentConnections.RemoveAll(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase))
                    _currentConnections.Add(connection)
                End If
            Else
                ' 创建新连接
                connection = New MCPConnectionConfig(connectionName, _serverUrlTextBox.Text, connectionType) With {
            .Description = String.Empty}

                ' 检查是否已存在同名连接
                Dim existingIndex = _currentConnections.FindIndex(Function(c) c.Name.Equals(connectionName, StringComparison.OrdinalIgnoreCase))
                If existingIndex >= 0 Then
                    _currentConnections(existingIndex) = connection
                Else
                    _currentConnections.Add(connection)
                End If
            End If

            ' 清除之前的工具列表
            connection.Tools.Clear()

            ' 转换工具列表为大模型函数调用格式
            If _currentTools IsNot Nothing AndAlso _currentTools.Count > 0 Then
                For Each tool In _currentTools
                    Dim llmFunction = ConvertToolToLLMFunction(tool)
                    connection.Tools.Add(llmFunction)
                Next
                Debug.WriteLine($"已转换 {connection.Tools.Count} 个工具为大模型函数格式")
            Else
                Debug.WriteLine("没有可用的工具信息")
            End If

            ' 保存连接配置 - 使用新的保存方法
            MCPConnectionManager.SaveConnections(_currentConnections)

            ' 更新当前连接名称
            _currentConnectionName = connectionName

            ' 重新加载连接列表
            LoadConnectionsList()

            _connectionsListView.SelectedItems.Clear() ' 先清除所有选中项

            ' 选中当前连接
            For i As Integer = 0 To _connectionsListView.Items.Count - 1
                Dim item = _connectionsListView.Items(i)
                Dim itemConnection = CType(item.Tag, MCPConnectionConfig)

                If itemConnection.Name.Equals(connectionName, StringComparison.OrdinalIgnoreCase) Then
                    item.Selected = True
                    _connectionsListView.Focus() ' 确保ListView获得焦点
                    _connectionsListView.EnsureVisible(i)
                    Exit For
                End If
            Next

        Catch ex As Exception
            UpdateStatus($"连接失败: {ex.Message}", Color.Red)
        Finally
            _testButton.Enabled = True
        End Try
    End Sub

    ' 更新 ExecuteTestAsync 方法，添加更多错误处理
    Private Async Sub ExecuteTestAsync(sender As Object, e As EventArgs)
        If _testToolCombo.SelectedItem Is Nothing Then
            MessageBox.Show("请选择要测试的工具", "提示")
            Return
        End If

        Try
            _executeTestButton.Enabled = False
            _testResultTextBox.Text = "正在执行..."

            Dim toolName = _testToolCombo.SelectedItem.ToString()
            Dim parameters As Object = Nothing

            ' 解析参数
            If Not String.IsNullOrWhiteSpace(_testParametersTextBox.Text) Then
                Try
                    parameters = Newtonsoft.Json.JsonConvert.DeserializeObject(_testParametersTextBox.Text)
                    ' 调试输出参数
                    Debug.WriteLine($"测试工具参数: {_testParametersTextBox.Text}")
                Catch jsonEx As Exception
                    _testResultTextBox.Text = $"参数JSON格式错误: {jsonEx.Message}"
                    _executeTestButton.Enabled = True
                    Return
                End Try
            End If

            ' 调用工具
            Debug.WriteLine($"开始调用工具: {toolName}")
            Dim result = Await _mcpClient.CallToolAsync(toolName, parameters)
            Debug.WriteLine($"工具调用完成: {toolName}, 是否错误: {result.IsError}")

            ' 显示结果
            If result.IsError Then
                _testResultTextBox.Text = $"执行失败: {result.ErrorMessage}"
            Else
                Dim resultText As New StringBuilder()

                ' 显示原始JSON以便调试
                resultText.AppendLine("【原始JSON响应】")
                Dim rawJson = Newtonsoft.Json.JsonConvert.SerializeObject(result, Newtonsoft.Json.Formatting.Indented)
                resultText.AppendLine(rawJson)
                resultText.AppendLine()

                resultText.AppendLine("【解析后内容】")

                ' 处理各种内容类型
                If result.Content IsNot Nothing AndAlso result.Content.Count > 0 Then
                    For Each content In result.Content
                        resultText.AppendLine($"内容类型: {content.Type}")

                        If content.Type = "text" AndAlso Not String.IsNullOrEmpty(content.Text) Then
                            resultText.AppendLine(content.Text)
                        ElseIf content.Type = "image" Then
                            resultText.AppendLine($"[图像内容] MIME类型: {content.MimeType}")
                        ElseIf Not String.IsNullOrEmpty(content.Data) Then
                            resultText.AppendLine($"[数据] {content.Data}")
                        Else
                            resultText.AppendLine("[无可显示内容]")
                        End If

                        resultText.AppendLine()
                    Next
                Else
                    resultText.AppendLine("响应中没有内容数据")
                End If

                _testResultTextBox.Text = resultText.ToString()
            End If

        Catch ex As Exception
            ' 显示更详细的错误信息
            Debug.WriteLine($"执行测试时发生异常: {ex.ToString()}")
            _testResultTextBox.Text = $"执行失败: {ex.Message}{Environment.NewLine}{Environment.NewLine}详细错误:{Environment.NewLine}{ex.ToString()}"
        Finally
            _executeTestButton.Enabled = True
        End Try
    End Sub

    ' 改进 ToolsListView_DoubleClick 方法
    Private Async Sub ToolsListView_DoubleClick(sender As Object, e As EventArgs)
        If _toolsListView.SelectedItems.Count = 0 Then Return

        Dim selectedTool = CType(_toolsListView.SelectedItems(0).Tag, MCPToolInfo)

        ' 创建一个详细信息窗口，代替简单的消息框
        Dim detailForm As New Form()
        detailForm.Text = $"工具详情: {selectedTool.Name}"
        detailForm.Size = New Size(700, 500)
        detailForm.StartPosition = FormStartPosition.CenterParent

        Dim detailBox As New RichTextBox()
        detailBox.Dock = DockStyle.Fill
        detailBox.ReadOnly = True
        detailBox.Font = New Font("Consolas", 9)
        detailForm.Controls.Add(detailBox)

        ' 填充详细信息
        ShowToolDetails(selectedTool, detailBox)

        ' 在测试标签页添加测试按钮
        Dim testButton As New Button()
        testButton.Text = "使用此工具测试"
        testButton.Dock = DockStyle.Bottom
        testButton.Height = 30

        AddHandler testButton.Click, Sub(s, ev)
                                         ' 切换到测试标签页并选择此工具
                                         _tabControl.SelectedIndex = 3 ' 假设测试标签页是第4个
                                         For i As Integer = 0 To _testToolCombo.Items.Count - 1
                                             If _testToolCombo.Items(i).ToString() = selectedTool.Name Then
                                                 _testToolCombo.SelectedIndex = i
                                                 Exit For
                                             End If
                                         Next
                                         detailForm.Close()
                                     End Sub

        detailForm.Controls.Add(testButton)

        ' 显示窗口
        detailForm.ShowDialog()
    End Sub

    Private Async Sub ResourcesListView_DoubleClick(sender As Object, e As EventArgs)
        If _resourcesListView.SelectedItems.Count = 0 Then Return

        Dim selectedResource = CType(_resourcesListView.SelectedItems(0).Tag, MCPResourceInfo)

        Try
            ' 读取资源内容
            Dim content = Await _mcpClient.ReadResourceAsync(selectedResource.Uri)

            Dim details = $"资源URI: {selectedResource.Uri}" & vbCrLf &
                         $"名称: {selectedResource.Name}" & vbCrLf &
                         $"描述: {selectedResource.Description}" & vbCrLf &
                         $"MIME类型: {selectedResource.MimeType}" & vbCrLf & vbCrLf

            For Each contentItem In content.Contents
                If Not String.IsNullOrEmpty(contentItem.Text) Then
                    details += $"内容: {contentItem.Text}" & vbCrLf
                End If
            Next

            MessageBox.Show(details, "资源详情", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"读取资源失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub UpdateStatus(message As String, color As Color)
        _statusLabel.Text = message
        _statusLabel.ForeColor = color
        Application.DoEvents()
    End Sub

    ' 修改 SaveButton_Click 方法，增加连接保存功能
    Private Sub SaveButton_Click(sender As Object, e As EventArgs)
        ' 验证连接名称
        If String.IsNullOrEmpty(_connectionNameTextBox.Text) Then
            MessageBox.Show("请输入连接名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 获取连接类型
        Dim connectionType As String = "HTTP"
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is ComboBox AndAlso ctrl.Location.Y = 67 Then
                Dim combo = CType(ctrl, ComboBox)
                connectionType = If(combo.SelectedIndex = 1, "Stdio", "HTTP")
                Exit For
            End If
        Next

        ' 保存连接配置
        Dim connectionName = _connectionNameTextBox.Text.Trim()
        'Dim connectionDescription = _connectionDescriptionTextBox.Text.Trim()  ' 获取描述信息

        ' 创建或更新连接配置
        Dim connection As MCPConnectionConfig

        If Not String.IsNullOrEmpty(_currentConnectionName) AndAlso
   _currentConnections.Exists(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase)) Then
            ' 更新现有连接
            connection = _currentConnections.Find(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase))
            connection.Name = connectionName
            connection.Description = ""  ' 保存描述
            connection.Url = _serverUrlTextBox.Text
            connection.ConnectionType = connectionType

            ' 如果是Stdio连接，保存环境变量
            If connectionType = "Stdio" AndAlso _serverUrlTextBox.Text.StartsWith("stdio://") Then
                Dim options = StdioOptions.Parse(_serverUrlTextBox.Text)
                connection.EnvironmentVariables.Clear()
                For Each kvp In options.EnvironmentVariables
                    connection.EnvironmentVariables.Add(kvp.Key, kvp.Value)
                Next
            End If

            ' 如果名称已变更，需要从列表中移除旧的
            If Not connectionName.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase) Then
                _currentConnections.RemoveAll(Function(c) c.Name.Equals(_currentConnectionName, StringComparison.OrdinalIgnoreCase))
                _currentConnections.Add(connection)
            End If
        Else
            ' 创建新连接
            connection = New MCPConnectionConfig(connectionName, _serverUrlTextBox.Text, connectionType) With {
        .Enabled = True,
        .Description = ""
    }

            ' 如果是Stdio连接，保存环境变量
            If connectionType = "Stdio" AndAlso _serverUrlTextBox.Text.StartsWith("stdio://") Then
                Dim options = StdioOptions.Parse(_serverUrlTextBox.Text)
                For Each kvp In options.EnvironmentVariables
                    connection.EnvironmentVariables.Add(kvp.Key, kvp.Value)
                Next
            End If

            ' 检查是否已存在同名连接
            Dim existingIndex = _currentConnections.FindIndex(Function(c) c.Name.Equals(connectionName, StringComparison.OrdinalIgnoreCase))
            If existingIndex >= 0 Then
                _currentConnections(existingIndex) = connection
            Else
                _currentConnections.Add(connection)
            End If
        End If

        ' 保存连接配置
        MCPConnectionManager.SaveConnections(_currentConnections)

        ' 更新当前连接名称
        _currentConnectionName = connectionName

        ' 重新加载连接列表
        LoadConnectionsList()

        'MessageBox.Show("配置已保存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    ' 为连接列表添加右键菜单
    Private Sub SetupConnectionsContextMenu()
        Dim menu = New ContextMenuStrip()

        Dim enableItem = New ToolStripMenuItem("启用")
        AddHandler enableItem.Click, Sub(sender, e)
                                         If _connectionsListView.SelectedItems.Count > 0 Then
                                             Dim connection = CType(_connectionsListView.SelectedItems(0).Tag, MCPConnectionConfig)
                                             connection.Enabled = True
                                             MCPConnectionManager.SaveConnections(_currentConnections)
                                             LoadConnectionsList()
                                         End If
                                     End Sub

        Dim disableItem = New ToolStripMenuItem("禁用")
        AddHandler disableItem.Click, Sub(sender, e)
                                          If _connectionsListView.SelectedItems.Count > 0 Then
                                              Dim connection = CType(_connectionsListView.SelectedItems(0).Tag, MCPConnectionConfig)
                                              connection.Enabled = False
                                              MCPConnectionManager.SaveConnections(_currentConnections)
                                              LoadConnectionsList()
                                          End If
                                      End Sub

        menu.Items.Add(enableItem)
        menu.Items.Add(disableItem)

        _connectionsListView.ContextMenuStrip = menu
    End Sub
    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            _mcpClient?.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class