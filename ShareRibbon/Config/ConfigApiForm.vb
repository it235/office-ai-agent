Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Security.Policy
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon.ConfigManager

''' <summary>
''' API配置窗体 - 双Tab布局 (云端模型/本地模型)
''' </summary>
Public Class ConfigApiForm
    Inherits Form

    ' 主控件
    Private mainTabControl As TabControl
    Private cloudTab As TabPage
    Private localTab As TabPage

    ' 云端模型Tab控件
    Private cloudProviderListBox As ListBox
    Private cloudPlatformLabel As Label
    Private cloudPlatformTextBox As TextBox
    Private cloudUrlLabel As Label
    Private cloudUrlTextBox As TextBox
    Private cloudApiKeyTextBox As TextBox
    Private cloudGetApiKeyButton As Button
    Private cloudModelCheckedListBox As CheckedListBox
    Private cloudRefreshModelsButton As Button
    Private cloudTranslateCheckBox As CheckBox
    Private cloudSaveButton As Button
    Private cloudDeleteButton As Button

    ' 本地模型Tab控件
    Private localProviderListBox As ListBox
    Private localPlatformTextBox As TextBox
    Private localUrlTextBox As TextBox
    Private localApiKeyTextBox As TextBox
    Private localDefaultKeyLabel As Label
    Private localModelCheckedListBox As CheckedListBox
    Private localRefreshModelsButton As Button
    Private localTranslateCheckBox As CheckBox
    Private localSaveButton As Button
    Private localDeleteButton As Button
    Private localAddButton As Button

    ' 当前选中的配置
    Private currentCloudConfig As ConfigItem
    Private currentLocalConfig As ConfigItem

    Public Sub New()
        InitializeForm()
        InitializeCloudTab()
        InitializeLocalTab()
        LoadDataToUI()
    End Sub

    ''' <summary>
    ''' 初始化窗体
    ''' </summary>
    Private Sub InitializeForm()
        Me.Text = "配置大模型API"
        Me.Size = New Size(700, 550)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 创建TabControl
        mainTabControl = New TabControl()
        mainTabControl.Dock = DockStyle.Fill
        Me.Controls.Add(mainTabControl)

        ' 创建云端模型Tab
        cloudTab = New TabPage()
        cloudTab.Text = "云端模型"
        cloudTab.Padding = New Padding(10)
        mainTabControl.TabPages.Add(cloudTab)

        ' 创建本地模型Tab
        localTab = New TabPage()
        localTab.Text = "本地模型"
        localTab.Padding = New Padding(10)
        mainTabControl.TabPages.Add(localTab)
    End Sub

    ''' <summary>
    ''' 初始化云端模型Tab
    ''' </summary>
    Private Sub InitializeCloudTab()
        ' 左侧：服务商列表
        Dim providerLabel As New Label()
        providerLabel.Text = "服务商列表："
        providerLabel.Location = New Point(10, 10)
        providerLabel.AutoSize = True
        cloudTab.Controls.Add(providerLabel)

        cloudProviderListBox = New ListBox()
        cloudProviderListBox.Location = New Point(10, 30)
        cloudProviderListBox.Size = New Size(180, 380)
        AddHandler cloudProviderListBox.SelectedIndexChanged, AddressOf CloudProviderListBox_SelectedIndexChanged
        cloudTab.Controls.Add(cloudProviderListBox)

        ' 添加新服务按钮
        Dim cloudAddButton As New Button()
        cloudAddButton.Text = "添加新服务"
        cloudAddButton.Location = New Point(10, 415)
        cloudAddButton.Size = New Size(180, 30)
        AddHandler cloudAddButton.Click, AddressOf CloudAddButton_Click
        cloudTab.Controls.Add(cloudAddButton)

        ' 右侧：配置面板
        Dim rightX As Integer = 210

        ' 平台名称 (Label for preset, TextBox for custom)
        cloudPlatformLabel = New Label()
        cloudPlatformLabel.Location = New Point(rightX, 10)
        cloudPlatformLabel.Size = New Size(440, 25)
        cloudPlatformLabel.Font = New Font(Me.Font.FontFamily, 11, FontStyle.Bold)
        cloudTab.Controls.Add(cloudPlatformLabel)

        cloudPlatformTextBox = New TextBox()
        cloudPlatformTextBox.Location = New Point(rightX, 10)
        cloudPlatformTextBox.Size = New Size(440, 25)
        cloudPlatformTextBox.Font = New Font(Me.Font.FontFamily, 11, FontStyle.Bold)
        cloudPlatformTextBox.Visible = False
        cloudTab.Controls.Add(cloudPlatformTextBox)

        ' API URL
        Dim urlTitleLabel As New Label()
        urlTitleLabel.Text = "API端点："
        urlTitleLabel.Location = New Point(rightX, 45)
        urlTitleLabel.AutoSize = True
        cloudTab.Controls.Add(urlTitleLabel)

        cloudUrlLabel = New Label()
        cloudUrlLabel.Location = New Point(rightX, 65)
        cloudUrlLabel.Size = New Size(440, 20)
        cloudUrlLabel.ForeColor = Color.DarkBlue
        cloudTab.Controls.Add(cloudUrlLabel)

        cloudUrlTextBox = New TextBox()
        cloudUrlTextBox.Location = New Point(rightX, 65)
        cloudUrlTextBox.Size = New Size(440, 20)
        cloudUrlTextBox.Visible = False
        cloudTab.Controls.Add(cloudUrlTextBox)

        ' API Key
        Dim apiKeyLabel As New Label()
        apiKeyLabel.Text = "API Key："
        apiKeyLabel.Location = New Point(rightX, 95)
        apiKeyLabel.AutoSize = True
        cloudTab.Controls.Add(apiKeyLabel)

        cloudApiKeyTextBox = New TextBox()
        cloudApiKeyTextBox.Location = New Point(rightX, 115)
        cloudApiKeyTextBox.Size = New Size(340, 25)
        cloudApiKeyTextBox.PasswordChar = "*"c
        AddHandler cloudApiKeyTextBox.Enter, AddressOf CloudApiKeyTextBox_Enter
        AddHandler cloudApiKeyTextBox.Leave, AddressOf CloudApiKeyTextBox_Leave
        cloudTab.Controls.Add(cloudApiKeyTextBox)

        ' 获取ApiKey按钮
        cloudGetApiKeyButton = New Button()
        cloudGetApiKeyButton.Text = "获取Key"
        cloudGetApiKeyButton.Location = New Point(rightX + 350, 113)
        cloudGetApiKeyButton.Size = New Size(90, 27)
        AddHandler cloudGetApiKeyButton.Click, AddressOf CloudGetApiKeyButton_Click
        cloudTab.Controls.Add(cloudGetApiKeyButton)

        ' 模型列表标题
        Dim modelLabel As New Label()
        modelLabel.Text = "模型列表："
        modelLabel.Location = New Point(rightX, 150)
        modelLabel.AutoSize = True
        cloudTab.Controls.Add(modelLabel)

        ' 刷新模型按钮
        cloudRefreshModelsButton = New Button()
        cloudRefreshModelsButton.Text = "刷新列表"
        cloudRefreshModelsButton.Location = New Point(rightX + 350, 145)
        cloudRefreshModelsButton.Size = New Size(90, 25)
        AddHandler cloudRefreshModelsButton.Click, AddressOf CloudRefreshModelsButton_Click
        cloudTab.Controls.Add(cloudRefreshModelsButton)

        ' 模型CheckedListBox
        cloudModelCheckedListBox = New CheckedListBox()
        cloudModelCheckedListBox.Location = New Point(rightX, 175)
        cloudModelCheckedListBox.Size = New Size(440, 200)
        cloudModelCheckedListBox.CheckOnClick = True
        AddHandler cloudModelCheckedListBox.ItemCheck, AddressOf CloudModelCheckedListBox_ItemCheck
        cloudTab.Controls.Add(cloudModelCheckedListBox)

        ' 用于翻译复选框
        cloudTranslateCheckBox = New CheckBox()
        cloudTranslateCheckBox.Text = "用于翻译"
        cloudTranslateCheckBox.Location = New Point(rightX, 385)
        cloudTranslateCheckBox.AutoSize = True
        cloudTab.Controls.Add(cloudTranslateCheckBox)

        ' 翻译提示
        Dim cloudTranslateTip As New Label()
        cloudTranslateTip.Text = "勾选后，翻译功能将使用该模型"
        cloudTranslateTip.Location = New Point(rightX + 85, 387)
        cloudTranslateTip.ForeColor = Color.Gray
        cloudTranslateTip.Font = New Font(Me.Font.FontFamily, 8)
        cloudTranslateTip.AutoSize = True
        cloudTab.Controls.Add(cloudTranslateTip)

        ' 验证并保存按钮
        cloudSaveButton = New Button()
        cloudSaveButton.Text = "验证并保存"
        cloudSaveButton.Location = New Point(rightX + 200, 410)
        cloudSaveButton.Size = New Size(110, 35)
        AddHandler cloudSaveButton.Click, AddressOf CloudSaveButton_Click
        cloudTab.Controls.Add(cloudSaveButton)

        ' 删除按钮
        cloudDeleteButton = New Button()
        cloudDeleteButton.Text = "删除"
        cloudDeleteButton.Location = New Point(rightX + 330, 410)
        cloudDeleteButton.Size = New Size(110, 35)
        AddHandler cloudDeleteButton.Click, AddressOf CloudDeleteButton_Click
        cloudTab.Controls.Add(cloudDeleteButton)
    End Sub

    ''' <summary>
    ''' 初始化本地模型Tab
    ''' </summary>
    Private Sub InitializeLocalTab()
        ' 左侧：服务商列表
        Dim providerLabel As New Label()
        providerLabel.Text = "本地服务列表："
        providerLabel.Location = New Point(10, 10)
        providerLabel.AutoSize = True
        localTab.Controls.Add(providerLabel)

        localProviderListBox = New ListBox()
        localProviderListBox.Location = New Point(10, 30)
        localProviderListBox.Size = New Size(180, 380)
        AddHandler localProviderListBox.SelectedIndexChanged, AddressOf LocalProviderListBox_SelectedIndexChanged
        localTab.Controls.Add(localProviderListBox)

        ' 添加新服务按钮
        localAddButton = New Button()
        localAddButton.Text = "添加新服务"
        localAddButton.Location = New Point(10, 415)
        localAddButton.Size = New Size(180, 30)
        AddHandler localAddButton.Click, AddressOf LocalAddButton_Click
        localTab.Controls.Add(localAddButton)

        ' 右侧：配置面板
        Dim rightX As Integer = 210

        ' 服务名称
        Dim platformLabel As New Label()
        platformLabel.Text = "服务名称："
        platformLabel.Location = New Point(rightX, 10)
        platformLabel.AutoSize = True
        localTab.Controls.Add(platformLabel)

        localPlatformTextBox = New TextBox()
        localPlatformTextBox.Location = New Point(rightX, 30)
        localPlatformTextBox.Size = New Size(440, 25)
        localTab.Controls.Add(localPlatformTextBox)

        ' API URL
        Dim urlLabel As New Label()
        urlLabel.Text = "API端点 (可编辑)："
        urlLabel.Location = New Point(rightX, 65)
        urlLabel.AutoSize = True
        localTab.Controls.Add(urlLabel)

        localUrlTextBox = New TextBox()
        localUrlTextBox.Location = New Point(rightX, 85)
        localUrlTextBox.Size = New Size(440, 25)
        localTab.Controls.Add(localUrlTextBox)

        ' API Key
        Dim apiKeyLabel As New Label()
        apiKeyLabel.Text = "API Key (大多数本地服务可留空)："
        apiKeyLabel.Location = New Point(rightX, 120)
        apiKeyLabel.AutoSize = True
        localTab.Controls.Add(apiKeyLabel)

        localApiKeyTextBox = New TextBox()
        localApiKeyTextBox.Location = New Point(rightX, 140)
        localApiKeyTextBox.Size = New Size(440, 25)
        localTab.Controls.Add(localApiKeyTextBox)

        ' 默认Key提示
        localDefaultKeyLabel = New Label()
        localDefaultKeyLabel.Location = New Point(rightX, 168)
        localDefaultKeyLabel.Size = New Size(440, 20)
        localDefaultKeyLabel.ForeColor = Color.Gray
        localDefaultKeyLabel.Font = New Font(Me.Font.FontFamily, 8)
        localTab.Controls.Add(localDefaultKeyLabel)

        ' 模型列表标题
        Dim modelLabel As New Label()
        modelLabel.Text = "模型列表："
        modelLabel.Location = New Point(rightX, 195)
        modelLabel.AutoSize = True
        localTab.Controls.Add(modelLabel)

        ' 刷新模型按钮
        localRefreshModelsButton = New Button()
        localRefreshModelsButton.Text = "刷新列表"
        localRefreshModelsButton.Location = New Point(rightX + 350, 190)
        localRefreshModelsButton.Size = New Size(90, 25)
        AddHandler localRefreshModelsButton.Click, AddressOf LocalRefreshModelsButton_Click
        localTab.Controls.Add(localRefreshModelsButton)

        ' 模型CheckedListBox
        localModelCheckedListBox = New CheckedListBox()
        localModelCheckedListBox.Location = New Point(rightX, 220)
        localModelCheckedListBox.Size = New Size(440, 150)
        localModelCheckedListBox.CheckOnClick = True
        AddHandler localModelCheckedListBox.ItemCheck, AddressOf LocalModelCheckedListBox_ItemCheck
        localTab.Controls.Add(localModelCheckedListBox)

        ' 用于翻译复选框
        localTranslateCheckBox = New CheckBox()
        localTranslateCheckBox.Text = "用于翻译"
        localTranslateCheckBox.Location = New Point(rightX, 380)
        localTranslateCheckBox.AutoSize = True
        localTab.Controls.Add(localTranslateCheckBox)

        ' 翻译提示
        Dim localTranslateTip As New Label()
        localTranslateTip.Text = "勾选后，翻译功能将使用该模型"
        localTranslateTip.Location = New Point(rightX + 85, 382)
        localTranslateTip.ForeColor = Color.Gray
        localTranslateTip.Font = New Font(Me.Font.FontFamily, 8)
        localTranslateTip.AutoSize = True
        localTab.Controls.Add(localTranslateTip)

        ' 保存按钮
        localSaveButton = New Button()
        localSaveButton.Text = "验证并保存"
        localSaveButton.Location = New Point(rightX + 200, 410)
        localSaveButton.Size = New Size(110, 35)
        AddHandler localSaveButton.Click, AddressOf LocalSaveButton_Click
        localTab.Controls.Add(localSaveButton)

        ' 删除按钮
        localDeleteButton = New Button()
        localDeleteButton.Text = "删除"
        localDeleteButton.Location = New Point(rightX + 330, 410)
        localDeleteButton.Size = New Size(110, 35)
        AddHandler localDeleteButton.Click, AddressOf LocalDeleteButton_Click
        localTab.Controls.Add(localDeleteButton)
    End Sub

    ''' <summary>
    ''' 加载数据到UI
    ''' </summary>
    Private Sub LoadDataToUI()
        ' 加载云端服务商
        cloudProviderListBox.Items.Clear()
        For Each config In ConfigData.Where(Function(c) c.providerType = ProviderType.Cloud)
            cloudProviderListBox.Items.Add(config)
        Next
        If cloudProviderListBox.Items.Count > 0 Then
            ' 选中当前使用的配置
            Dim selectedIndex = 0
            For i = 0 To cloudProviderListBox.Items.Count - 1
                Dim item = CType(cloudProviderListBox.Items(i), ConfigItem)
                If item.selected Then
                    selectedIndex = i
                    Exit For
                End If
            Next
            cloudProviderListBox.SelectedIndex = selectedIndex
        End If

        ' 加载本地服务商
        localProviderListBox.Items.Clear()
        For Each config In ConfigData.Where(Function(c) c.providerType = ProviderType.Local)
            localProviderListBox.Items.Add(config)
        Next
        If localProviderListBox.Items.Count > 0 Then
            Dim selectedIndex = 0
            For i = 0 To localProviderListBox.Items.Count - 1
                Dim item = CType(localProviderListBox.Items(i), ConfigItem)
                If item.selected Then
                    selectedIndex = i
                    Exit For
                End If
            Next
            localProviderListBox.SelectedIndex = selectedIndex
        End If
    End Sub

#Region "云端模型事件处理"

    Private Sub CloudProviderListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cloudProviderListBox.SelectedItem Is Nothing Then Return

        currentCloudConfig = CType(cloudProviderListBox.SelectedItem, ConfigItem)
        
        ' 根据是否为预置配置切换显示模式
        Dim isPreset = currentCloudConfig.isPreset
        
        ' 平台名称：预置用Label，自定义用TextBox
        cloudPlatformLabel.Visible = isPreset
        cloudPlatformTextBox.Visible = Not isPreset
        If isPreset Then
            cloudPlatformLabel.Text = currentCloudConfig.pltform
        Else
            cloudPlatformTextBox.Text = currentCloudConfig.pltform
        End If
        
        ' API URL：预置用Label，自定义用TextBox
        cloudUrlLabel.Visible = isPreset
        cloudUrlTextBox.Visible = Not isPreset
        If isPreset Then
            cloudUrlLabel.Text = currentCloudConfig.url
        Else
            cloudUrlTextBox.Text = currentCloudConfig.url
        End If
        
        cloudApiKeyTextBox.Text = If(String.IsNullOrEmpty(currentCloudConfig.key), "", currentCloudConfig.key)
        cloudTranslateCheckBox.Checked = currentCloudConfig.translateSelected

        ' 加载模型列表
        cloudModelCheckedListBox.Items.Clear()
        For Each model In currentCloudConfig.model
            Dim displayText = If(String.IsNullOrEmpty(model.displayName), model.modelName, model.displayName)
            cloudModelCheckedListBox.Items.Add(model, model.selected)
        Next

        ' 控制删除按钮可见性（预置配置不可删除）
        cloudDeleteButton.Enabled = Not isPreset
    End Sub

    Private Sub CloudApiKeyTextBox_Enter(sender As Object, e As EventArgs)
        cloudApiKeyTextBox.PasswordChar = Nothing
    End Sub

    Private Sub CloudApiKeyTextBox_Leave(sender As Object, e As EventArgs)
        cloudApiKeyTextBox.PasswordChar = "*"c
    End Sub

    Private Sub CloudGetApiKeyButton_Click(sender As Object, e As EventArgs)
        If currentCloudConfig Is Nothing OrElse String.IsNullOrEmpty(currentCloudConfig.registerUrl) Then
            MessageBox.Show("该服务商暂无注册链接", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Try
            Process.Start(New ProcessStartInfo(currentCloudConfig.registerUrl) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show($"无法打开浏览器: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Async Sub CloudRefreshModelsButton_Click(sender As Object, e As EventArgs)
        If currentCloudConfig Is Nothing Then Return

        Dim apiKey = cloudApiKeyTextBox.Text
        If String.IsNullOrEmpty(apiKey) Then
            MessageBox.Show("请先输入API Key", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 对于自定义配置，从TextBox获取URL
        Dim apiUrl = If(currentCloudConfig.isPreset, currentCloudConfig.url, cloudUrlTextBox.Text)
        If String.IsNullOrEmpty(apiUrl) Then
            MessageBox.Show("请先输入API端点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        cloudRefreshModelsButton.Enabled = False
        cloudRefreshModelsButton.Text = "刷新中..."
        Cursor = Cursors.WaitCursor

        ' 将用户输入同步到配置对象，防止刷新过程中丢失
        If Not currentCloudConfig.isPreset Then
            currentCloudConfig.pltform = cloudPlatformTextBox.Text
            currentCloudConfig.url = cloudUrlTextBox.Text
        End If
        currentCloudConfig.key = apiKey

        Try
            Dim models = Await ModelApiClient.GetModelsAsync(apiUrl, apiKey)
            If models.Count > 0 Then
                ' 保留已有模型的选中状态，添加新模型
                For Each modelName In models
                    Dim existing = currentCloudConfig.model.FirstOrDefault(Function(m) m.modelName = modelName)
                    If existing Is Nothing Then
                        currentCloudConfig.model.Add(New ConfigItemModel() With {
                            .modelName = modelName,
                            .displayName = modelName
                        })
                    End If
                Next

                ' 仅刷新模型列表，保持用户输入不变
                RefreshCloudModelList()
                MessageBox.Show($"已获取 {models.Count} 个模型", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("未获取到模型列表，请检查API Key是否正确", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"刷新模型列表失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            cloudRefreshModelsButton.Enabled = True
            cloudRefreshModelsButton.Text = "刷新列表"
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub CloudModelCheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        ' 单选逻辑：只允许选中一个模型
        If e.NewValue = CheckState.Checked Then
            For i = 0 To cloudModelCheckedListBox.Items.Count - 1
                If i <> e.Index Then
                    cloudModelCheckedListBox.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Async Sub CloudSaveButton_Click(sender As Object, e As EventArgs)
        If currentCloudConfig Is Nothing Then Return

        ' 对于自定义配置，从TextBox获取平台名和URL
        Dim platformName As String
        Dim apiUrl As String
        If currentCloudConfig.isPreset Then
            platformName = currentCloudConfig.pltform
            apiUrl = currentCloudConfig.url
        Else
            platformName = cloudPlatformTextBox.Text
            apiUrl = cloudUrlTextBox.Text
            
            If String.IsNullOrEmpty(platformName) Then
                MessageBox.Show("请输入服务名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            
            If String.IsNullOrEmpty(apiUrl) OrElse Not (apiUrl.StartsWith("http://") OrElse apiUrl.StartsWith("https://")) Then
                MessageBox.Show("请输入有效的API端点 (以http://或https://开头)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        End If

        Dim apiKey = cloudApiKeyTextBox.Text
        If String.IsNullOrEmpty(apiKey) Then
            MessageBox.Show("请输入API Key", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 获取选中的模型
        Dim selectedModelName As String = ""
        For i = 0 To cloudModelCheckedListBox.Items.Count - 1
            If cloudModelCheckedListBox.GetItemChecked(i) Then
                Dim model = CType(cloudModelCheckedListBox.Items(i), ConfigItemModel)
                selectedModelName = model.modelName
                Exit For
            End If
        Next

        If String.IsNullOrEmpty(selectedModelName) Then
            MessageBox.Show("请选择一个模型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        cloudSaveButton.Enabled = False
        cloudSaveButton.Text = "验证中..."
        Cursor = Cursors.WaitCursor

        ' 将用户输入同步到配置对象，防止异步验证期间UI事件覆盖输入
        If Not currentCloudConfig.isPreset Then
            currentCloudConfig.pltform = platformName
            currentCloudConfig.url = apiUrl
        End If
        currentCloudConfig.key = apiKey

        Try
            ' 验证API
            Dim validationResult = Await ValidateApiAsync(apiUrl, apiKey, selectedModelName)
            If validationResult Then
                ' 更新配置
                currentCloudConfig.pltform = platformName
                currentCloudConfig.url = apiUrl
                currentCloudConfig.key = apiKey
                currentCloudConfig.validated = True
                currentCloudConfig.translateSelected = cloudTranslateCheckBox.Checked

                ' 更新模型选中状态
                For Each model In currentCloudConfig.model
                    model.selected = (model.modelName = selectedModelName)
                Next

                ' 更新全局选中状态
                For Each config In ConfigData
                    config.selected = (config Is currentCloudConfig)
                    If config IsNot currentCloudConfig Then
                        config.translateSelected = If(cloudTranslateCheckBox.Checked, False, config.translateSelected)
                    End If
                Next

                ' 更新全局配置
                ConfigSettings.ApiUrl = currentCloudConfig.url
                ConfigSettings.ApiKey = apiKey
                ConfigSettings.platform = currentCloudConfig.pltform
                ConfigSettings.ModelName = selectedModelName

                Dim selectedModel = currentCloudConfig.model.FirstOrDefault(Function(m) m.modelName = selectedModelName)
                If selectedModel IsNot Nothing Then
                    ConfigSettings.mcpable = selectedModel.mcpable
                End If

                ' 保存配置
                SaveConfig()

                MessageBox.Show("配置已保存", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("API验证失败，请检查API Key和模型名称是否正确", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"验证失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            cloudSaveButton.Enabled = True
            cloudSaveButton.Text = "验证并保存"
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub CloudDeleteButton_Click(sender As Object, e As EventArgs)
        If currentCloudConfig Is Nothing Then Return
        If currentCloudConfig.isPreset Then
            MessageBox.Show("预置配置不可删除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If MessageBox.Show($"确定要删除 {currentCloudConfig.pltform} 吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ConfigData.Remove(currentCloudConfig)
            SaveConfig()
            LoadDataToUI()
        End If
    End Sub

    Private Sub CloudAddButton_Click(sender As Object, e As EventArgs)
        ' 创建新的云端服务配置
        Dim newConfig As New ConfigItem() With {
            .pltform = "新云端服务",
            .url = "https://api.example.com/v1/chat/completions",
            .providerType = ProviderType.Cloud,
            .isPreset = False,
            .key = "",
            .registerUrl = "",
            .translateSelected = True,
            .model = New List(Of ConfigItemModel)()
        }

        ConfigData.Add(newConfig)
        cloudProviderListBox.Items.Add(newConfig)
        cloudProviderListBox.SelectedItem = newConfig
    End Sub

#End Region

#Region "本地模型事件处理"

    Private Sub LocalProviderListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        If localProviderListBox.SelectedItem Is Nothing Then Return

        currentLocalConfig = CType(localProviderListBox.SelectedItem, ConfigItem)
        
        ' 更新配置面板
        localPlatformTextBox.Text = currentLocalConfig.pltform
        localUrlTextBox.Text = currentLocalConfig.url
        localApiKeyTextBox.Text = If(String.IsNullOrEmpty(currentLocalConfig.key), "", currentLocalConfig.key)
        localDefaultKeyLabel.Text = If(String.IsNullOrEmpty(currentLocalConfig.defaultApiKey), "", $"提示: 默认APIKey为 '{currentLocalConfig.defaultApiKey}'，大多数情况可留空")
        localTranslateCheckBox.Checked = currentLocalConfig.translateSelected

        ' 加载模型列表
        localModelCheckedListBox.Items.Clear()
        For Each model In currentLocalConfig.model
            Dim displayText = If(String.IsNullOrEmpty(model.displayName), model.modelName, model.displayName)
            localModelCheckedListBox.Items.Add(model, model.selected)
        Next

        ' 控制删除按钮可见性（预置配置可删除但会提示）
        localDeleteButton.Enabled = True
        localPlatformTextBox.ReadOnly = currentLocalConfig.isPreset
    End Sub

    Private Sub LocalAddButton_Click(sender As Object, e As EventArgs)
        ' 创建新的本地服务配置
        Dim newConfig As New ConfigItem() With {
            .pltform = "新本地服务",
            .url = "http://localhost:8000/v1/chat/completions",
            .providerType = ProviderType.Local,
            .isPreset = False,
            .key = "",
            .defaultApiKey = "",
            .translateSelected = True,
            .model = New List(Of ConfigItemModel)()
        }

        ConfigData.Add(newConfig)
        localProviderListBox.Items.Add(newConfig)
        localProviderListBox.SelectedItem = newConfig
    End Sub

    Private Async Sub LocalRefreshModelsButton_Click(sender As Object, e As EventArgs)
        If currentLocalConfig Is Nothing Then Return

        Dim apiUrl = localUrlTextBox.Text
        Dim apiKey = If(String.IsNullOrEmpty(localApiKeyTextBox.Text), currentLocalConfig.defaultApiKey, localApiKeyTextBox.Text)

        If String.IsNullOrEmpty(apiUrl) Then
            MessageBox.Show("请先输入API端点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        localRefreshModelsButton.Enabled = False
        localRefreshModelsButton.Text = "刷新中..."
        Cursor = Cursors.WaitCursor

        ' 将用户输入同步到配置对象，防止刷新过程中丢失
        currentLocalConfig.pltform = localPlatformTextBox.Text
        currentLocalConfig.url = apiUrl
        If Not String.IsNullOrEmpty(localApiKeyTextBox.Text) Then
            currentLocalConfig.key = localApiKeyTextBox.Text
        End If

        Try
            Dim models = Await ModelApiClient.GetModelsAsync(apiUrl, apiKey)
            If models.Count > 0 Then
                ' 清空并重新加载模型列表
                currentLocalConfig.model.Clear()
                For Each modelName In models
                    currentLocalConfig.model.Add(New ConfigItemModel() With {
                        .modelName = modelName,
                        .displayName = modelName,
                        .selected = (currentLocalConfig.model.Count = 0)
                    })
                Next

                ' 仅刷新模型列表，保持用户输入不变
                RefreshLocalModelList()
                MessageBox.Show($"已获取 {models.Count} 个模型", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("未获取到模型列表，请确保本地服务已启动", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"刷新模型列表失败: {ex.Message}" & vbCrLf & "请确保本地服务已启动", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            localRefreshModelsButton.Enabled = True
            localRefreshModelsButton.Text = "刷新列表"
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LocalModelCheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        ' 单选逻辑
        If e.NewValue = CheckState.Checked Then
            For i = 0 To localModelCheckedListBox.Items.Count - 1
                If i <> e.Index Then
                    localModelCheckedListBox.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Async Sub LocalSaveButton_Click(sender As Object, e As EventArgs)
        If currentLocalConfig Is Nothing Then Return

        Dim platformName = localPlatformTextBox.Text
        Dim apiUrl = localUrlTextBox.Text
        Dim apiKey = If(String.IsNullOrEmpty(localApiKeyTextBox.Text), currentLocalConfig.defaultApiKey, localApiKeyTextBox.Text)

        If String.IsNullOrEmpty(platformName) Then
            MessageBox.Show("请输入服务名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If String.IsNullOrEmpty(apiUrl) OrElse Not (apiUrl.StartsWith("http://") OrElse apiUrl.StartsWith("https://")) Then
            MessageBox.Show("请输入有效的API端点 (以http://或https://开头)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 获取选中的模型
        Dim selectedModelName As String = ""
        For i = 0 To localModelCheckedListBox.Items.Count - 1
            If localModelCheckedListBox.GetItemChecked(i) Then
                Dim model = CType(localModelCheckedListBox.Items(i), ConfigItemModel)
                selectedModelName = model.modelName
                Exit For
            End If
        Next

        localSaveButton.Enabled = False
        localSaveButton.Text = "验证中..."
        Cursor = Cursors.WaitCursor

        ' 将用户输入同步到配置对象，防止异步验证期间UI事件覆盖输入
        currentLocalConfig.pltform = platformName
        currentLocalConfig.url = apiUrl
        currentLocalConfig.key = apiKey

        Try
            ' 本地模型验证 - 尝试连接
            Dim validationResult = True
            If Not String.IsNullOrEmpty(selectedModelName) Then
                validationResult = Await ValidateApiAsync(apiUrl, apiKey, selectedModelName)
            End If

            If validationResult OrElse String.IsNullOrEmpty(selectedModelName) Then
                ' 更新配置
                currentLocalConfig.pltform = platformName
                currentLocalConfig.url = apiUrl
                currentLocalConfig.key = apiKey
                currentLocalConfig.validated = validationResult
                currentLocalConfig.translateSelected = localTranslateCheckBox.Checked

                ' 更新模型选中状态
                For Each model In currentLocalConfig.model
                    model.selected = (model.modelName = selectedModelName)
                Next

                ' 更新全局选中状态
                For Each config In ConfigData
                    config.selected = (config Is currentLocalConfig)
                    If config IsNot currentLocalConfig Then
                        config.translateSelected = If(localTranslateCheckBox.Checked, False, config.translateSelected)
                    End If
                Next

                ' 更新全局配置
                If Not String.IsNullOrEmpty(selectedModelName) Then
                    ConfigSettings.ApiUrl = apiUrl
                    ConfigSettings.ApiKey = apiKey
                    ConfigSettings.platform = platformName
                    ConfigSettings.ModelName = selectedModelName
                End If

                ' 保存配置
                SaveConfig()

                ' 刷新ListBox显示
                Dim selectedIndex = localProviderListBox.SelectedIndex
                localProviderListBox.Items(selectedIndex) = currentLocalConfig

                MessageBox.Show("配置已保存", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                MessageBox.Show("无法连接到本地服务，请确保服务已启动", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            localSaveButton.Enabled = True
            localSaveButton.Text = "验证并保存"
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LocalDeleteButton_Click(sender As Object, e As EventArgs)
        If currentLocalConfig Is Nothing Then Return

        Dim message = If(currentLocalConfig.isPreset, $"{currentLocalConfig.pltform} 是预置配置，删除后重启将恢复，确定要删除吗？", $"确定要删除 {currentLocalConfig.pltform} 吗？")

        If MessageBox.Show(message, "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ConfigData.Remove(currentLocalConfig)
            SaveConfig()
            LoadDataToUI()
        End If
    End Sub

#End Region

#Region "通用方法"

    ''' <summary>
    ''' 仅刷新云端模型CheckedListBox，不影响其他输入控件
    ''' </summary>
    Private Sub RefreshCloudModelList()
        cloudModelCheckedListBox.Items.Clear()
        If currentCloudConfig Is Nothing Then Return
        For Each model In currentCloudConfig.model
            cloudModelCheckedListBox.Items.Add(model, model.selected)
        Next
    End Sub

    ''' <summary>
    ''' 仅刷新本地模型CheckedListBox，不影响其他输入控件
    ''' </summary>
    Private Sub RefreshLocalModelList()
        localModelCheckedListBox.Items.Clear()
        If currentLocalConfig Is Nothing Then Return
        For Each model In currentLocalConfig.model
            localModelCheckedListBox.Items.Add(model, model.selected)
        Next
    End Sub

    ''' <summary>
    ''' 验证API连接
    ''' </summary>
    Private Async Function ValidateApiAsync(apiUrl As String, apiKey As String, modelName As String) As Task(Of Boolean)
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(60)

                Dim requestBody As String
                Dim request As HttpRequestMessage

                ' 检查是否是Anthropic
                If apiUrl.Contains("anthropic.com") Then
                    ' Anthropic格式
                    requestBody = $"{{""model"": ""{modelName}"", ""max_tokens"": 100, ""messages"": [{{""role"": ""user"", ""content"": ""hi""}}]}}"
                    request = New HttpRequestMessage(HttpMethod.Post, apiUrl)
                    request.Headers.Add("x-api-key", apiKey)
                    request.Headers.Add("anthropic-version", "2023-06-01")
                Else
                    ' OpenAI兼容格式
                    requestBody = $"{{""model"": ""{modelName}"", ""stream"": false, ""messages"": [{{""role"": ""user"", ""content"": ""hi""}}]}}"
                    request = New HttpRequestMessage(HttpMethod.Post, apiUrl)
                    request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)
                End If

                request.Content = New StringContent(requestBody, Encoding.UTF8, "application/json")

                Using response = Await client.SendAsync(request)
                    Return response.IsSuccessStatusCode
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine($"API验证异常: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

End Class
