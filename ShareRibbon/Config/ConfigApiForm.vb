Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Security.Policy
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ShareRibbon.ConfigManager
Imports Services.SkillsService
Imports Services.SkillsDirectoryService
Imports Services.SkillsFileDefinition
Imports Markdig
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

''' <summary>
''' API配置窗体 - 四Tab布局 (云端模型/本地模型/场景与Skills/记忆管理)
''' </summary>
Public Class ConfigApiForm
    Inherits Form

    ' 主控件
    Private mainTabControl As TabControl
    Private cloudTab As TabPage
    Private localTab As TabPage
    Private skillsTab As TabPage
    Private memoryTab As TabPage

    ' 云端模型Tab控件
    Private cloudProviderListBox As ListBox
    Private cloudPlatformLabel As Label
    Private cloudPlatformTextBox As TextBox
    Private cloudUrlLabel As Label
    Private cloudUrlTextBox As TextBox
    Private cloudApiKeyTextBox As TextBox
    Private cloudGetApiKeyButton As Button
    Private cloudChatModelCheckedListBox As CheckedListBox
    Private cloudEmbeddingModelCheckedListBox As CheckedListBox
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
    Private localChatModelCheckedListBox As CheckedListBox
    Private localEmbeddingModelCheckedListBox As CheckedListBox
    Private localRefreshModelsButton As Button
    Private localTranslateCheckBox As CheckBox
    Private localSaveButton As Button
    Private localDeleteButton As Button
    Private localAddButton As Button

    ' Skills Tab控件
    Private skillsListBox As ListBox
    Private _skills As New List(Of SkillFileDefinition)()
    Private skillsWebView2 As WebView2
    Private _webView2Initialized As Boolean = False
    Private scriptsPanel As Panel
    Private referencesPanel As Panel
    Private assetsPanel As Panel

    ' 新增：右侧内部分割容器的字段，便于在事件中访问并延迟调整
    Private skillsRightSplit As SplitContainer

    ' 当前选中的配置
    Private currentCloudConfig As ConfigItem
    Private currentLocalConfig As ConfigItem
    Private _applicationInfo As ApplicationInfo

    Public Sub New(appInfo As ApplicationInfo)
        _applicationInfo = appInfo
        InitializeForm()
        InitializeCloudTab()
        InitializeLocalTab()
        InitializeSkillsTab()
        InitializeMemoryTab()
        LoadDataToUI()
    End Sub

    ''' <summary>
    ''' 初始化窗体
    ''' </summary>
    Private Sub InitializeForm()
        Me.Text = "配置大模型API"
        Me.Size = New Size(1050, 800)
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

        ' 创建Skills Tab
        skillsTab = New TabPage()
        skillsTab.Text = "Skills"
        skillsTab.Padding = New Padding(10)
        mainTabControl.TabPages.Add(skillsTab)

        ' 创建记忆管理Tab
        memoryTab = New TabPage()
        memoryTab.Text = "记忆管理"
        memoryTab.Padding = New Padding(10)
        mainTabControl.TabPages.Add(memoryTab)
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
        cloudProviderListBox.Size = New Size(180, 480)
        AddHandler cloudProviderListBox.SelectedIndexChanged, AddressOf CloudProviderListBox_SelectedIndexChanged
        cloudTab.Controls.Add(cloudProviderListBox)

        ' 添加新服务按钮
        Dim cloudAddButton As New Button()
        cloudAddButton.Text = "添加新服务"
        cloudAddButton.Location = New Point(10, 515)
        cloudAddButton.Size = New Size(180, 30)
        AddHandler cloudAddButton.Click, AddressOf CloudAddButton_Click
        cloudTab.Controls.Add(cloudAddButton)

        ' 右侧：配置面板
        Dim rightX As Integer = 210

        ' 平台名称 (Label for preset, TextBox for custom)
        cloudPlatformLabel = New Label()
        cloudPlatformLabel.Location = New Point(rightX, 10)
        cloudPlatformLabel.Size = New Size(590, 25)
        cloudPlatformLabel.Font = New Font(Me.Font.FontFamily, 11, FontStyle.Bold)
        cloudTab.Controls.Add(cloudPlatformLabel)

        cloudPlatformTextBox = New TextBox()
        cloudPlatformTextBox.Location = New Point(rightX, 10)
        cloudPlatformTextBox.Size = New Size(590, 25)
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
        cloudUrlLabel.Size = New Size(590, 20)
        cloudUrlLabel.ForeColor = Color.DarkBlue
        cloudTab.Controls.Add(cloudUrlLabel)

        cloudUrlTextBox = New TextBox()
        cloudUrlTextBox.Location = New Point(rightX, 65)
        cloudUrlTextBox.Size = New Size(590, 20)
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
        cloudApiKeyTextBox.Size = New Size(490, 25)
        cloudApiKeyTextBox.PasswordChar = "*"c
        AddHandler cloudApiKeyTextBox.Enter, AddressOf CloudApiKeyTextBox_Enter
        AddHandler cloudApiKeyTextBox.Leave, AddressOf CloudApiKeyTextBox_Leave
        cloudTab.Controls.Add(cloudApiKeyTextBox)

        ' 获取ApiKey按钮
        cloudGetApiKeyButton = New Button()
        cloudGetApiKeyButton.Text = "获取Key"
        cloudGetApiKeyButton.Location = New Point(rightX + 500, 113)
        cloudGetApiKeyButton.Size = New Size(90, 27)
        AddHandler cloudGetApiKeyButton.Click, AddressOf CloudGetApiKeyButton_Click
        cloudTab.Controls.Add(cloudGetApiKeyButton)

        ' 对话模型列表标题
        Dim chatModelLabel As New Label()
        chatModelLabel.Text = "对话模型："
        chatModelLabel.Location = New Point(rightX, 150)
        chatModelLabel.AutoSize = True
        cloudTab.Controls.Add(chatModelLabel)

        ' 对话模型CheckedListBox
        cloudChatModelCheckedListBox = New CheckedListBox()
        cloudChatModelCheckedListBox.Location = New Point(rightX, 175)
        cloudChatModelCheckedListBox.Size = New Size(285, 180)
        cloudChatModelCheckedListBox.CheckOnClick = True
        AddHandler cloudChatModelCheckedListBox.ItemCheck, AddressOf CloudChatModelCheckedListBox_ItemCheck
        cloudTab.Controls.Add(cloudChatModelCheckedListBox)

        ' 向量模型列表标题
        Dim embeddingModelLabel As New Label()
        embeddingModelLabel.Text = "向量模型："
        embeddingModelLabel.Location = New Point(rightX + 305, 150)
        embeddingModelLabel.AutoSize = True
        cloudTab.Controls.Add(embeddingModelLabel)

        ' 向量模型CheckedListBox
        cloudEmbeddingModelCheckedListBox = New CheckedListBox()
        cloudEmbeddingModelCheckedListBox.Location = New Point(rightX + 305, 175)
        cloudEmbeddingModelCheckedListBox.Size = New Size(285, 180)
        cloudEmbeddingModelCheckedListBox.CheckOnClick = True
        AddHandler cloudEmbeddingModelCheckedListBox.ItemCheck, AddressOf CloudEmbeddingModelCheckedListBox_ItemCheck
        cloudTab.Controls.Add(cloudEmbeddingModelCheckedListBox)

        ' 刷新模型按钮
        cloudRefreshModelsButton = New Button()
        cloudRefreshModelsButton.Text = "刷新列表"
        cloudRefreshModelsButton.Location = New Point(rightX + 450, 145)
        cloudRefreshModelsButton.Size = New Size(140, 25)
        AddHandler cloudRefreshModelsButton.Click, AddressOf CloudRefreshModelsButton_Click
        cloudTab.Controls.Add(cloudRefreshModelsButton)


        ' 用于翻译复选框
        cloudTranslateCheckBox = New CheckBox()
        cloudTranslateCheckBox.Text = "用于翻译"
        cloudTranslateCheckBox.Location = New Point(rightX, 365)
        cloudTranslateCheckBox.AutoSize = True
        cloudTab.Controls.Add(cloudTranslateCheckBox)

        ' 翻译提示
        Dim cloudTranslateTip As New Label()
        cloudTranslateTip.Text = "勾选后，翻译功能将使用该模型"
        cloudTranslateTip.Location = New Point(rightX + 85, 367)
        cloudTranslateTip.ForeColor = Color.Gray
        cloudTranslateTip.Font = New Font(Me.Font.FontFamily, 8)
        cloudTranslateTip.AutoSize = True
        cloudTab.Controls.Add(cloudTranslateTip)

        ' 验证并保存按钮
        cloudSaveButton = New Button()
        cloudSaveButton.Text = "验证并保存"
        cloudSaveButton.Location = New Point(rightX + 320, 410)
        cloudSaveButton.Size = New Size(130, 35)
        AddHandler cloudSaveButton.Click, AddressOf CloudSaveButton_Click
        cloudTab.Controls.Add(cloudSaveButton)

        ' 删除按钮
        cloudDeleteButton = New Button()
        cloudDeleteButton.Text = "删除"
        cloudDeleteButton.Location = New Point(rightX + 460, 410)
        cloudDeleteButton.Size = New Size(130, 35)
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
        localProviderListBox.Size = New Size(180, 480)
        AddHandler localProviderListBox.SelectedIndexChanged, AddressOf LocalProviderListBox_SelectedIndexChanged
        localTab.Controls.Add(localProviderListBox)

        ' 添加新服务按钮
        localAddButton = New Button()
        localAddButton.Text = "添加新服务"
        localAddButton.Location = New Point(10, 515)
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
        localPlatformTextBox.Size = New Size(590, 25)
        localTab.Controls.Add(localPlatformTextBox)

        ' API URL
        Dim urlLabel As New Label()
        urlLabel.Text = "API端点 (可编辑)："
        urlLabel.Location = New Point(rightX, 65)
        urlLabel.AutoSize = True
        localTab.Controls.Add(urlLabel)

        localUrlTextBox = New TextBox()
        localUrlTextBox.Location = New Point(rightX, 85)
        localUrlTextBox.Size = New Size(590, 25)
        localTab.Controls.Add(localUrlTextBox)

        ' API Key
        Dim apiKeyLabel As New Label()
        apiKeyLabel.Text = "API Key (大多数本地服务可留空)："
        apiKeyLabel.Location = New Point(rightX, 120)
        apiKeyLabel.AutoSize = True
        localTab.Controls.Add(apiKeyLabel)

        localApiKeyTextBox = New TextBox()
        localApiKeyTextBox.Location = New Point(rightX, 140)
        localApiKeyTextBox.Size = New Size(590, 25)
        localTab.Controls.Add(localApiKeyTextBox)

        ' 默认Key提示
        localDefaultKeyLabel = New Label()
        localDefaultKeyLabel.Location = New Point(rightX, 168)
        localDefaultKeyLabel.Size = New Size(590, 20)
        localDefaultKeyLabel.ForeColor = Color.Gray
        localDefaultKeyLabel.Font = New Font(Me.Font.FontFamily, 8)
        localTab.Controls.Add(localDefaultKeyLabel)

        ' 对话模型列表标题
        Dim chatModelLabel As New Label()
        chatModelLabel.Text = "对话模型："
        chatModelLabel.Location = New Point(rightX, 195)
        chatModelLabel.AutoSize = True
        localTab.Controls.Add(chatModelLabel)

        ' 对话模型CheckedListBox
        localChatModelCheckedListBox = New CheckedListBox()
        localChatModelCheckedListBox.Location = New Point(rightX, 220)
        localChatModelCheckedListBox.Size = New Size(285, 130)
        localChatModelCheckedListBox.CheckOnClick = True
        AddHandler localChatModelCheckedListBox.ItemCheck, AddressOf LocalChatModelCheckedListBox_ItemCheck
        localTab.Controls.Add(localChatModelCheckedListBox)

        ' 向量模型列表标题
        Dim embeddingModelLabel As New Label()
        embeddingModelLabel.Text = "向量模型："
        embeddingModelLabel.Location = New Point(rightX + 305, 195)
        embeddingModelLabel.AutoSize = True
        localTab.Controls.Add(embeddingModelLabel)

        ' 向量模型CheckedListBox
        localEmbeddingModelCheckedListBox = New CheckedListBox()
        localEmbeddingModelCheckedListBox.Location = New Point(rightX + 305, 220)
        localEmbeddingModelCheckedListBox.Size = New Size(285, 130)
        localEmbeddingModelCheckedListBox.CheckOnClick = True
        AddHandler localEmbeddingModelCheckedListBox.ItemCheck, AddressOf LocalEmbeddingModelCheckedListBox_ItemCheck
        localTab.Controls.Add(localEmbeddingModelCheckedListBox)

        ' 刷新模型按钮
        localRefreshModelsButton = New Button()
        localRefreshModelsButton.Text = "刷新列表"
        localRefreshModelsButton.Location = New Point(rightX + 450, 190)
        localRefreshModelsButton.Size = New Size(140, 25)
        AddHandler localRefreshModelsButton.Click, AddressOf LocalRefreshModelsButton_Click
        localTab.Controls.Add(localRefreshModelsButton)

        ' 用于翻译复选框
        localTranslateCheckBox = New CheckBox()
        localTranslateCheckBox.Text = "用于翻译"
        localTranslateCheckBox.Location = New Point(rightX, 360)
        localTranslateCheckBox.AutoSize = True
        localTab.Controls.Add(localTranslateCheckBox)

        ' 翻译提示
        Dim localTranslateTip As New Label()
        localTranslateTip.Text = "勾选后，翻译功能将使用该模型"
        localTranslateTip.Location = New Point(rightX + 85, 362)
        localTranslateTip.ForeColor = Color.Gray
        localTranslateTip.Font = New Font(Me.Font.FontFamily, 8)
        localTranslateTip.AutoSize = True
        localTab.Controls.Add(localTranslateTip)

        ' 保存按钮
        localSaveButton = New Button()
        localSaveButton.Text = "验证并保存"
        localSaveButton.Location = New Point(rightX + 320, 410)
        localSaveButton.Size = New Size(130, 35)
        AddHandler localSaveButton.Click, AddressOf LocalSaveButton_Click
        localTab.Controls.Add(localSaveButton)

        ' 删除按钮
        localDeleteButton = New Button()
        localDeleteButton.Text = "删除"
        localDeleteButton.Location = New Point(rightX + 460, 410)
        localDeleteButton.Size = New Size(130, 35)
        AddHandler localDeleteButton.Click, AddressOf LocalDeleteButton_Click
        localTab.Controls.Add(localDeleteButton)
    End Sub
    Private Sub AdjustRightSplitAfterHandle(rightSplit As SplitContainer)
        ' Only run when the control has a meaningful height
        If rightSplit Is Nothing OrElse rightSplit.Height <= 1 Then Return

        ' 提高期望的上半部分高度，使内容区域更高
        Dim desired As Integer = 420
        Dim minAllowed As Integer = 300
        Dim panel2Min As Integer = 120

        ' clamp panel2Min so it doesn't exceed the container height minus 1
        Dim finalPanel2Min = Math.Min(panel2Min, Math.Max(0, rightSplit.Height - 1))
        ' clamp panel1Min so there's at least 1px spare
        Dim finalPanel1Min = Math.Min(minAllowed, Math.Max(0, rightSplit.Height - finalPanel2Min - 1))

        ' Compute a valid splitter position within the final bounds
        Dim lower = finalPanel1Min
        Dim upper = Math.Max(0, rightSplit.Height - finalPanel2Min)
        Dim finalSplitter = Math.Min(Math.Max(desired, lower), upper)

        rightSplit.SuspendLayout()
        Try
            ' Ensure SplitterDistance is first set to a value valid for the final min sizes
            rightSplit.SplitterDistance = finalSplitter
            rightSplit.Panel1MinSize = finalPanel1Min
            rightSplit.Panel2MinSize = finalPanel2Min
        Finally
            rightSplit.ResumeLayout()
        End Try
    End Sub

    Private Sub SkillsTab_SelectedIndexChanged(sender As Object, e As EventArgs)
        If mainTabControl.SelectedTab Is skillsTab Then
            ' 一次性处理：移除 handler 避免重复触发
            RemoveHandler mainTabControl.SelectedIndexChanged, AddressOf SkillsTab_SelectedIndexChanged

            ' 延迟到消息循环，确保所有布局完成
            Me.BeginInvoke(Sub()
                               Try
                                   AdjustRightSplitAfterHandle(skillsRightSplit)
                               Catch ex As Exception
                                   ' 忽略单次异常
                               End Try
                           End Sub)
        End If
    End Sub

    ''' <summary>
    ''' 初始化场景与Skills Tab
    ''' </summary>
    Private Sub InitializeSkillsTab()
        ' 顶部说明
        Dim lblInfo As New Label() With {
            .Text = "Skills目录：Documents\OfficeAiAppData\Skills，将符合Claude规范的Skills目录拷贝到此即可",
            .Location = New Point(12, 12),
            .Size = New Size(1080, 24),
            .ForeColor = Color.Gray,
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        skillsTab.Controls.Add(lblInfo)

        ' 主分隔：左侧列表，右侧详情
        Dim mainSplit As New SplitContainer() With {
            .Location = New Point(12, 42),
            .Size = New Size(1080, 570),
            .SplitterDistance = 250,
            .Panel1MinSize = 180,
            .Panel2MinSize = 380,
            .FixedPanel = FixedPanel.None
        }
        mainSplit.Panel1.SuspendLayout()
        mainSplit.Panel2.SuspendLayout()

        ' 左侧：Skills列表
        Dim lblList As New Label() With {
            .Text = "已安装的Skills：",
            .Location = New Point(0, 0),
            .Size = New Size(245, 24),
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        mainSplit.Panel1.Controls.Add(lblList)
        skillsListBox = New ListBox() With {
            .Location = New Point(0, 28),
            .Size = New Size(245, 540),
            .HorizontalScrollbar = True,
            .HorizontalExtent = 600,
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        AddHandler skillsListBox.SelectedIndexChanged, AddressOf SkillsListBox_SelectedIndexChanged
        mainSplit.Panel1.Controls.Add(skillsListBox)

        ' 右侧：详情区域 - 再分成上下两部分，上面是内容区（高度更大）
        skillsRightSplit = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal
        }
        ' 先添加到父 Panel（确保 Parent 存在），再延后设置与布局相关的属性
        mainSplit.Panel2.Controls.Add(skillsRightSplit)

        ' 延迟调整：在用户切换到 Skills 标签页后再做一次调整（保证布局已完成）
        AddHandler mainTabControl.SelectedIndexChanged, AddressOf SkillsTab_SelectedIndexChanged

        skillsRightSplit.Panel1.SuspendLayout()
        skillsRightSplit.Panel2.SuspendLayout()

        ' 上半部分：内容预览（WebView2）
        Dim topContentPanel As New Panel() With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(8)
        }
        skillsRightSplit.Panel1.Controls.Add(topContentPanel)

        Dim lblContent As New Label() With {
            .Text = "内容：",
            .Dock = DockStyle.Top,
            .Height = 28,
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        topContentPanel.Controls.Add(lblContent)

        skillsWebView2 = New WebView2() With {
            .Dock = DockStyle.Fill
        }
        topContentPanel.Controls.Add(skillsWebView2)

        ' 下半部分：元数据和目录文件列表
        Dim bottomDetailPanel As New Panel() With {
            .Dock = DockStyle.Fill
        }
        skillsRightSplit.Panel2.Controls.Add(bottomDetailPanel)
        Dim detailY As Integer = 0

        ' 名称
        Dim lblName As New Label() With {
            .Text = "名称：",
            .Location = New Point(12, detailY),
            .Size = New Size(85, 24),
            .Font = New Font(Me.Font.FontFamily, 9.5, FontStyle.Bold)
        }
        bottomDetailPanel.Controls.Add(lblName)
        Dim txtName As New Label() With {
            .Name = "SkillsTxtName",
            .Location = New Point(102, detailY),
            .Size = New Size(690, 24),
            .ForeColor = Color.FromArgb(70, 130, 180),
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        bottomDetailPanel.Controls.Add(txtName)
        detailY += 34

        ' 描述
        Dim lblDescription As New Label() With {
            .Text = "描述：",
            .Location = New Point(12, detailY),
            .Size = New Size(85, 24),
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        bottomDetailPanel.Controls.Add(lblDescription)
        Dim txtDescription As New Label() With {
            .Name = "SkillsTxtDescription",
            .Location = New Point(102, detailY),
            .Size = New Size(690, 42),
            .ForeColor = Color.DarkGray,
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        bottomDetailPanel.Controls.Add(txtDescription)
        detailY += 52

        ' 元数据行
        Dim metadataPanel As New Panel() With {
            .Location = New Point(102, detailY),
            .Size = New Size(690, 100),
            .BackColor = Color.FromArgb(245, 245, 245)
        }
        bottomDetailPanel.Controls.Add(metadataPanel)

        Dim lblLicense As New Label() With {
            .Text = "许可证：",
            .Location = New Point(8, 8),
            .Size = New Size(72, 22),
            .ForeColor = Color.Gray,
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        metadataPanel.Controls.Add(lblLicense)
        Dim txtLicense As New Label() With {
            .Name = "SkillsTxtLicense",
            .Location = New Point(84, 8),
            .Size = New Size(595, 22),
            .ForeColor = Color.FromArgb(100, 100, 100),
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        metadataPanel.Controls.Add(txtLicense)

        Dim lblAuthor As New Label() With {
            .Text = "作者：",
            .Location = New Point(8, 34),
            .Size = New Size(72, 22),
            .ForeColor = Color.Gray,
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        metadataPanel.Controls.Add(lblAuthor)
        Dim txtAuthor As New Label() With {
            .Name = "SkillsTxtAuthor",
            .Location = New Point(84, 34),
            .Size = New Size(280, 22),
            .ForeColor = Color.FromArgb(100, 100, 100),
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        metadataPanel.Controls.Add(txtAuthor)

        Dim lblVersion As New Label() With {
            .Text = "版本：",
            .Location = New Point(380, 34),
            .Size = New Size(60, 22),
            .ForeColor = Color.Gray,
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        metadataPanel.Controls.Add(lblVersion)
        Dim txtVersion As New Label() With {
            .Name = "SkillsTxtVersion",
            .Location = New Point(445, 34),
            .Size = New Size(235, 22),
            .ForeColor = Color.FromArgb(100, 100, 100),
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        metadataPanel.Controls.Add(txtVersion)
        detailY += 108

        ' 目录文件列表容器
        Dim filesContainerPanel As New Panel() With {
            .Location = New Point(12, detailY),
            .Size = New Size(790, 96)
        }
        bottomDetailPanel.Controls.Add(filesContainerPanel)

        ' scripts目录面板
        scriptsPanel = New Panel() With {
            .Location = New Point(0, 0),
            .Size = New Size(255, 96),
            .Visible = False
        }
        filesContainerPanel.Controls.Add(scriptsPanel)
        Dim lblScripts As New Label() With {
            .Text = "scripts：",
            .Location = New Point(0, 0),
            .Size = New Size(255, 22),
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        scriptsPanel.Controls.Add(lblScripts)
        Dim scriptsList As New ListBox() With {
            .Name = "ScriptsListBox",
            .Location = New Point(0, 26),
            .Size = New Size(255, 70),
            .HorizontalScrollbar = True,
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        scriptsPanel.Controls.Add(scriptsList)

        ' references目录面板
        referencesPanel = New Panel() With {
            .Location = New Point(0, 0),
            .Size = New Size(255, 96),
            .Visible = False
        }
        filesContainerPanel.Controls.Add(referencesPanel)
        Dim lblReferences As New Label() With {
            .Text = "references：",
            .Location = New Point(0, 0),
            .Size = New Size(255, 22),
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        referencesPanel.Controls.Add(lblReferences)
        Dim referencesList As New ListBox() With {
            .Name = "ReferencesListBox",
            .Location = New Point(0, 26),
            .Size = New Size(255, 70),
            .HorizontalScrollbar = True,
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        referencesPanel.Controls.Add(referencesList)

        ' assets目录面板
        assetsPanel = New Panel() With {
            .Location = New Point(0, 0),
            .Size = New Size(255, 96),
            .Visible = False
        }
        filesContainerPanel.Controls.Add(assetsPanel)
        Dim lblAssets As New Label() With {
            .Text = "assets：",
            .Location = New Point(0, 0),
            .Size = New Size(255, 22),
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        assetsPanel.Controls.Add(lblAssets)
        Dim assetsList As New ListBox() With {
            .Name = "AssetsListBox",
            .Location = New Point(0, 26),
            .Size = New Size(255, 70),
            .HorizontalScrollbar = True,
            .Font = New Font(Me.Font.FontFamily, 9)
        }
        assetsPanel.Controls.Add(assetsList)

        mainSplit.Panel1.ResumeLayout(False)
        mainSplit.Panel2.ResumeLayout(False)
        skillsRightSplit.Panel1.ResumeLayout(False)
        skillsRightSplit.Panel2.ResumeLayout(False)
        skillsTab.Controls.Add(mainSplit)

        ' 底部按钮
        Dim btnOpenDir As New Button() With {
            .Text = "打开Skills目录",
            .Location = New Point(12, 620),
            .Size = New Size(145, 34),
            .BackColor = Color.FromArgb(70, 130, 180),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        AddHandler btnOpenDir.Click, AddressOf SkillsOpenDir_Click
        skillsTab.Controls.Add(btnOpenDir)

        Dim btnRefresh As New Button() With {
            .Text = "刷新列表",
            .Location = New Point(167, 620),
            .Size = New Size(120, 34),
            .Font = New Font(Me.Font.FontFamily, 9.5)
        }
        AddHandler btnRefresh.Click, AddressOf SkillsRefresh_Click
        skillsTab.Controls.Add(btnRefresh)

        ' 初始化 WebView2
        InitializeSkillsWebView2()
    End Sub


    ''' <summary>
    ''' 初始化 WebView2
    ''' </summary>
    Private Async Sub InitializeSkillsWebView2()
        Try
            Dim userDataFolder As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                ConfigSettings.OfficeAiAppDataFolder,
                "SkillsWebView2Data")

            If Not Directory.Exists(userDataFolder) Then
                Directory.CreateDirectory(userDataFolder)
            End If

            Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
            Await skillsWebView2.EnsureCoreWebView2Async(env)

            If skillsWebView2.CoreWebView2 IsNot Nothing Then
                skillsWebView2.CoreWebView2.Settings.IsScriptEnabled = True
                skillsWebView2.CoreWebView2.Settings.AreDevToolsEnabled = False
                _webView2Initialized = True
            End If
        Catch ex As Exception
            Debug.WriteLine($"WebView2初始化失败: {ex.Message}")
        End Try
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

        ' 加载Skills
        LoadSkills()
    End Sub

    ''' <summary>
    ''' 加载Skills
    ''' </summary>
    Private Sub LoadSkills()
        Try
            SkillsDirectoryService.EnsureDirectoryExists()
            _skills = SkillsDirectoryService.GetAllSkills(forceRefresh:=True)

            skillsListBox.Items.Clear()
            For Each skill In _skills
                Dim folderName = If(Directory.Exists(skill.FilePath), Path.GetFileName(skill.FilePath), Path.GetFileNameWithoutExtension(skill.FilePath))
                skillsListBox.Items.Add(New SkillsListItem With {.Skill = skill, .DisplayText = folderName})
            Next

            If skillsListBox.Items.Count = 0 Then
                skillsListBox.Items.Add("(暂无Skills，请打开Skills目录添加)")
            End If
        Catch ex As Exception
            skillsListBox.Items.Clear()
            skillsListBox.Items.Add("(加载失败: " & ex.Message & ")")
        End Try
    End Sub



    ''' <summary>
    ''' Skills列表选中事件
    ''' </summary>
    Private Sub SkillsListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim item = TryCast(skillsListBox.SelectedItem, SkillsListItem)
        If item Is Nothing Then
            ClearSkillsDetail()
            Return
        End If

        Dim skill = item.Skill
        If skill Is Nothing Then
            ClearSkillsDetail()
            Return
        End If

        ' 更新详情
        Dim txtName = Me.Controls.Find("SkillsTxtName", True).FirstOrDefault()
        If txtName IsNot Nothing Then txtName.Text = skill.Name

        Dim txtDescription = Me.Controls.Find("SkillsTxtDescription", True).FirstOrDefault()
        If txtDescription IsNot Nothing Then txtDescription.Text = If(String.IsNullOrWhiteSpace(skill.Description), "(无描述)", skill.Description)

        Dim txtLicense = Me.Controls.Find("SkillsTxtLicense", True).FirstOrDefault()
        If txtLicense IsNot Nothing Then txtLicense.Text = If(String.IsNullOrWhiteSpace(skill.License), "-", skill.License)

        Dim txtAuthor = Me.Controls.Find("SkillsTxtAuthor", True).FirstOrDefault()
        If txtAuthor IsNot Nothing Then txtAuthor.Text = If(String.IsNullOrWhiteSpace(skill.Author), "-", skill.Author)

        Dim txtVersion = Me.Controls.Find("SkillsTxtVersion", True).FirstOrDefault()
        If txtVersion IsNot Nothing Then txtVersion.Text = If(String.IsNullOrWhiteSpace(skill.Version), "-", skill.Version)

        ' 加载并显示 scripts、references、assets 目录文件列表
        LoadAndShowSkillDirectoryFiles(skill)

        ' 使用 WebView2 和 Markdig 显示 Markdown 预览
        ShowMarkdownInWebView2(skill.Content)
    End Sub


    ''' <summary>
    ''' 加载并显示Skill目录下的文件列表（有就显示，没有就隐藏）
    ''' </summary>
    Private Sub LoadAndShowSkillDirectoryFiles(skill As SkillFileDefinition)
        scriptsPanel.Visible = False
        referencesPanel.Visible = False
        assetsPanel.Visible = False

        If String.IsNullOrWhiteSpace(skill.FilePath) OrElse Not Directory.Exists(skill.FilePath) Then
            Return
        End If

        Dim currentX As Integer = 0
        Dim panelWidth As Integer = 255
        Dim gap As Integer = 10

        ' 加载 scripts 目录
        Dim scriptsDir = Path.Combine(skill.FilePath, "scripts")
        If Directory.Exists(scriptsDir) Then
            Dim scriptFiles = Directory.GetFiles(scriptsDir)
            If scriptFiles.Length > 0 Then
                Dim scriptsList = TryCast(scriptsPanel.Controls.Find("ScriptsListBox", True).FirstOrDefault(), ListBox)
                If scriptsList IsNot Nothing Then
                    scriptsList.Items.Clear()
                    For Each file In scriptFiles
                        scriptsList.Items.Add(Path.GetFileName(file))
                    Next
                End If
                scriptsPanel.Location = New Point(currentX, 0)
                scriptsPanel.Visible = True
                currentX += panelWidth + gap
            End If
        End If

        ' 加载 references 目录
        Dim referencesDir = Path.Combine(skill.FilePath, "references")
        If Directory.Exists(referencesDir) Then
            Dim refFiles = Directory.GetFiles(referencesDir)
            If refFiles.Length > 0 Then
                Dim referencesList = TryCast(referencesPanel.Controls.Find("ReferencesListBox", True).FirstOrDefault(), ListBox)
                If referencesList IsNot Nothing Then
                    referencesList.Items.Clear()
                    For Each file In refFiles
                        referencesList.Items.Add(Path.GetFileName(file))
                    Next
                End If
                referencesPanel.Location = New Point(currentX, 0)
                referencesPanel.Visible = True
                currentX += panelWidth + gap
            End If
        End If

        ' 加载 assets 目录
        Dim assetsDir = Path.Combine(skill.FilePath, "assets")
        If Directory.Exists(assetsDir) Then
            Dim assetFiles = Directory.GetFiles(assetsDir)
            If assetFiles.Length > 0 Then
                Dim assetsList = TryCast(assetsPanel.Controls.Find("AssetsListBox", True).FirstOrDefault(), ListBox)
                If assetsList IsNot Nothing Then
                    assetsList.Items.Clear()
                    For Each file In assetFiles
                        assetsList.Items.Add(Path.GetFileName(file))
                    Next
                End If
                assetsPanel.Location = New Point(currentX, 0)
                assetsPanel.Visible = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' 将Markdown转换为HTML并在WebView2中显示（使用Markdig）
    ''' </summary>
    Private Sub ShowMarkdownInWebView2(markdown As String)
        If String.IsNullOrWhiteSpace(markdown) Then
            markdown = "(无内容)"
        End If

        Try
            ' 使用 Markdig 转换
            Dim htmlContent = Markdig.Markdown.ToHtml(markdown)

            ' 生成完整的HTML文档
            Dim fullHtml = $"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1'>
    <style>
        body {{ 
            font-family: 'Microsoft YaHei', 'Segoe UI', Arial, sans-serif; 
            padding: 20px; 
            line-height: 1.7; 
            color: #333;
            background-color: #fff;
        }}
        h1, h2, h3, h4, h5, h6 {{ 
            color: #1a1a1a; 
            margin-top: 1.5em;
            margin-bottom: 0.5em;
            font-weight: 600;
        }}
        h1 {{ border-bottom: 2px solid #eaecef; padding-bottom: 0.3em; font-size: 1.8em; }}
        h2 {{ border-bottom: 1px solid #eaecef; padding-bottom: 0.3em; font-size: 1.5em; }}
        h3 {{ font-size: 1.25em; }}
        p {{ margin: 1em 0; }}
        code {{ 
            background: #f6f8fa; 
            padding: 0.2em 0.4em; 
            border-radius: 3px; 
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 0.9em;
            color: #e83e8c;
        }}
        pre {{ 
            background: #f6f8fa; 
            padding: 16px; 
            border-radius: 6px; 
            overflow-x: auto;
            margin: 1em 0;
        }}
        pre code {{ 
            background: none; 
            padding: 0;
            color: #333;
        }}
        blockquote {{ 
            border-left: 4px solid #dfe2e5; 
            padding-left: 16px; 
            margin: 1em 0; 
            color: #6a737d;
            background: #f8f9fa;
            padding: 10px 16px;
            border-radius: 0 6px 6px 0;
        }}
        ul, ol {{ 
            padding-left: 2em; 
            margin: 1em 0;
        }}
        li {{ margin: 0.3em 0; }}
        table {{ 
            border-collapse: collapse; 
            width: 100%; 
            margin: 1em 0;
        }}
        th, td {{ 
            border: 1px solid #dfe2e5; 
            padding: 8px 12px; 
            text-align: left;
        }}
        th {{ background: #f6f8fa; font-weight: 600; }}
        a {{ color: #0366d6; text-decoration: none; }}
        a:hover {{ text-decoration: underline; }}
        img {{ max-width: 100%; height: auto; }}
        hr {{ 
            border: none; 
            border-top: 1px solid #eaecef; 
            margin: 2em 0;
        }}
    </style>
</head>
<body>
{htmlContent}
</body>
</html>"

            ' 显示在 WebView2 中
            If _webView2Initialized AndAlso skillsWebView2.CoreWebView2 IsNot Nothing Then
                skillsWebView2.NavigateToString(fullHtml)
            Else
                ' 如果 WebView2 还没初始化，先设置 DocumentText（临时方案）
                skillsWebView2.NavigateToString(fullHtml)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Markdown渲染失败: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 清空Skills详情
    ''' </summary>
    Private Sub ClearSkillsDetail()
        Dim txtName = Me.Controls.Find("SkillsTxtName", True).FirstOrDefault()
        If txtName IsNot Nothing Then txtName.Text = ""

        Dim txtDescription = Me.Controls.Find("SkillsTxtDescription", True).FirstOrDefault()
        If txtDescription IsNot Nothing Then txtDescription.Text = ""

        Dim txtLicense = Me.Controls.Find("SkillsTxtLicense", True).FirstOrDefault()
        If txtLicense IsNot Nothing Then txtLicense.Text = ""

        Dim txtAuthor = Me.Controls.Find("SkillsTxtAuthor", True).FirstOrDefault()
        If txtAuthor IsNot Nothing Then txtAuthor.Text = ""

        Dim txtVersion = Me.Controls.Find("SkillsTxtVersion", True).FirstOrDefault()
        If txtVersion IsNot Nothing Then txtVersion.Text = ""

        scriptsPanel.Visible = False
        referencesPanel.Visible = False
        assetsPanel.Visible = False

        If _webView2Initialized AndAlso skillsWebView2.CoreWebView2 IsNot Nothing Then
            skillsWebView2.NavigateToString("")
        End If
    End Sub

    ''' <summary>
    ''' 打开Skills目录 - 如果选中了Skill则打开该Skill的目录
    ''' </summary>
    Private Sub SkillsOpenDir_Click(sender As Object, e As EventArgs)
        Try
            Dim item = TryCast(skillsListBox.SelectedItem, SkillsListItem)
            If item IsNot Nothing AndAlso item.Skill IsNot Nothing Then
                SkillsDirectoryService.OpenSkillDirectory(item.Skill)
            Else
                SkillsDirectoryService.OpenSkillsDirectory()
            End If
        Catch ex As Exception
        End Try
    End Sub


    ''' <summary>
    ''' 刷新Skills列表
    ''' </summary>
    Private Sub SkillsRefresh_Click(sender As Object, e As EventArgs)
        LoadSkills()
    End Sub

    ''' <summary>
    ''' Skills列表项
    ''' </summary>
    Private Class SkillsListItem
        Public Property Skill As SkillFileDefinition
        Public Property DisplayText As String
        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class

#Region "云端模型事件处理"

    Private Sub CloudProviderListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        If cloudProviderListBox.SelectedItem Is Nothing Then Return

        currentCloudConfig = CType(cloudProviderListBox.SelectedItem, ConfigItem)

        Dim isPreset = currentCloudConfig.isPreset

        cloudPlatformLabel.Visible = isPreset
        cloudPlatformTextBox.Visible = Not isPreset
        If isPreset Then
            cloudPlatformLabel.Text = currentCloudConfig.pltform
        Else
            cloudPlatformTextBox.Text = currentCloudConfig.pltform
        End If

        cloudUrlLabel.Visible = isPreset
        cloudUrlTextBox.Visible = Not isPreset
        If isPreset Then
            cloudUrlLabel.Text = currentCloudConfig.url
        Else
            cloudUrlTextBox.Text = currentCloudConfig.url
        End If

        cloudApiKeyTextBox.Text = If(String.IsNullOrEmpty(currentCloudConfig.key), "", currentCloudConfig.key)
        cloudTranslateCheckBox.Checked = currentCloudConfig.translateSelected

        RefreshCloudModelLists()

        cloudDeleteButton.Enabled = Not isPreset
    End Sub

    Private Sub RefreshCloudModelLists()
        cloudChatModelCheckedListBox.Items.Clear()
        cloudEmbeddingModelCheckedListBox.Items.Clear()
        If currentCloudConfig Is Nothing Then Return

        For Each model In currentCloudConfig.model
            If model.modelType = ModelType.Chat Then
                cloudChatModelCheckedListBox.Items.Add(model, model.selected)
            ElseIf model.modelType = ModelType.Embedding Then
                cloudEmbeddingModelCheckedListBox.Items.Add(model, model.selected)
            End If
        Next
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

        Dim apiUrl = If(currentCloudConfig.isPreset, currentCloudConfig.url, cloudUrlTextBox.Text)
        If String.IsNullOrEmpty(apiUrl) Then
            MessageBox.Show("请先输入API端点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        cloudRefreshModelsButton.Enabled = False
        cloudRefreshModelsButton.Text = "刷新中..."
        Cursor = Cursors.WaitCursor

        If Not currentCloudConfig.isPreset Then
            currentCloudConfig.pltform = cloudPlatformTextBox.Text
            currentCloudConfig.url = cloudUrlTextBox.Text
        End If
        currentCloudConfig.key = apiKey

        Try
            Dim models = Await ModelApiClient.GetModelsAsync(apiUrl, apiKey)
            If models.Count > 0 Then
                For Each modelName In models
                    Dim existing = currentCloudConfig.model.FirstOrDefault(Function(m) m.modelName = modelName)
                    If existing Is Nothing Then
                        currentCloudConfig.model.Add(New ConfigItemModel() With {
                            .modelName = modelName,
                            .displayName = modelName,
                            .modelType = ModelType.Chat
                        })
                    End If
                Next

                RefreshCloudModelLists()
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

    Private Sub CloudChatModelCheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        If e.NewValue = CheckState.Checked Then
            For i = 0 To cloudChatModelCheckedListBox.Items.Count - 1
                If i <> e.Index Then
                    cloudChatModelCheckedListBox.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Sub CloudEmbeddingModelCheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        If e.NewValue = CheckState.Checked Then
            For i = 0 To cloudEmbeddingModelCheckedListBox.Items.Count - 1
                If i <> e.Index Then
                    cloudEmbeddingModelCheckedListBox.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Async Sub CloudSaveButton_Click(sender As Object, e As EventArgs)
        If currentCloudConfig Is Nothing Then Return

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

        Dim selectedChatModelName As String = ""
        For i = 0 To cloudChatModelCheckedListBox.Items.Count - 1
            If cloudChatModelCheckedListBox.GetItemChecked(i) Then
                Dim model = CType(cloudChatModelCheckedListBox.Items(i), ConfigItemModel)
                selectedChatModelName = model.modelName
                Exit For
            End If
        Next

        If String.IsNullOrEmpty(selectedChatModelName) Then
            MessageBox.Show("请选择一个对话模型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim selectedEmbeddingModelName As String = ""
        For i = 0 To cloudEmbeddingModelCheckedListBox.Items.Count - 1
            If cloudEmbeddingModelCheckedListBox.GetItemChecked(i) Then
                Dim model = CType(cloudEmbeddingModelCheckedListBox.Items(i), ConfigItemModel)
                selectedEmbeddingModelName = model.modelName
                Exit For
            End If
        Next

        cloudSaveButton.Enabled = False
        cloudSaveButton.Text = "验证中..."
        Cursor = Cursors.WaitCursor

        If Not currentCloudConfig.isPreset Then
            currentCloudConfig.pltform = platformName
            currentCloudConfig.url = apiUrl
        End If
        currentCloudConfig.key = apiKey

        Try
            Dim validationResult = Await ValidateApiAsync(apiUrl, apiKey, selectedChatModelName)
            If validationResult Then
                currentCloudConfig.pltform = platformName
                currentCloudConfig.url = apiUrl
                currentCloudConfig.key = apiKey
                currentCloudConfig.validated = True
                currentCloudConfig.translateSelected = cloudTranslateCheckBox.Checked

                For Each model In currentCloudConfig.model
                    If model.modelType = ModelType.Chat Then
                        model.selected = (model.modelName = selectedChatModelName)
                    ElseIf model.modelType = ModelType.Embedding Then
                        model.selected = (model.modelName = selectedEmbeddingModelName)
                    End If
                Next

                For Each config In ConfigData
                    config.selected = (config Is currentCloudConfig)
                    If config IsNot currentCloudConfig Then
                        config.translateSelected = If(cloudTranslateCheckBox.Checked, False, config.translateSelected)
                    End If
                Next

                ConfigSettings.ApiUrl = currentCloudConfig.url
                ConfigSettings.ApiKey = apiKey
                ConfigSettings.platform = currentCloudConfig.pltform
                ConfigSettings.ModelName = selectedChatModelName

                Dim selectedChatModel = currentCloudConfig.model.FirstOrDefault(Function(m) m.modelName = selectedChatModelName)
                If selectedChatModel IsNot Nothing Then
                    ConfigSettings.mcpable = selectedChatModel.mcpable
                End If

                If Not String.IsNullOrEmpty(selectedEmbeddingModelName) Then
                    ConfigSettings.EmbeddingModel = selectedEmbeddingModelName
                End If

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

        localPlatformTextBox.Text = currentLocalConfig.pltform
        localUrlTextBox.Text = currentLocalConfig.url
        localApiKeyTextBox.Text = If(String.IsNullOrEmpty(currentLocalConfig.key), "", currentLocalConfig.key)
        localDefaultKeyLabel.Text = If(String.IsNullOrEmpty(currentLocalConfig.defaultApiKey), "", $"提示: 默认APIKey为 '{currentLocalConfig.defaultApiKey}'，大多数情况可留空")
        localTranslateCheckBox.Checked = currentLocalConfig.translateSelected

        RefreshLocalModelLists()

        localDeleteButton.Enabled = True
        localPlatformTextBox.ReadOnly = currentLocalConfig.isPreset
    End Sub

    Private Sub RefreshLocalModelLists()
        localChatModelCheckedListBox.Items.Clear()
        localEmbeddingModelCheckedListBox.Items.Clear()
        If currentLocalConfig Is Nothing Then Return

        For Each model In currentLocalConfig.model
            If model.modelType = ModelType.Chat Then
                localChatModelCheckedListBox.Items.Add(model, model.selected)
            ElseIf model.modelType = ModelType.Embedding Then
                localEmbeddingModelCheckedListBox.Items.Add(model, model.selected)
            End If
        Next
    End Sub

    Private Async Sub LocalRefreshModelsButton_Click(sender As Object, e As EventArgs)
        If currentLocalConfig Is Nothing Then Return

        Dim apiKey = localApiKeyTextBox.Text
        If String.IsNullOrEmpty(apiKey) AndAlso Not String.IsNullOrEmpty(currentLocalConfig.defaultApiKey) Then
            apiKey = currentLocalConfig.defaultApiKey
        End If

        Dim apiUrl = localUrlTextBox.Text
        If String.IsNullOrEmpty(apiUrl) Then
            MessageBox.Show("请先输入API端点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        localRefreshModelsButton.Enabled = False
        localRefreshModelsButton.Text = "刷新中..."
        Cursor = Cursors.WaitCursor

        currentLocalConfig.pltform = localPlatformTextBox.Text
        currentLocalConfig.url = apiUrl
        currentLocalConfig.key = apiKey

        Try
            Dim models = Await ModelApiClient.GetModelsAsync(apiUrl, apiKey)
            If models.Count > 0 Then
                For Each modelName In models
                    Dim existing = currentLocalConfig.model.FirstOrDefault(Function(m) m.modelName = modelName)
                    If existing Is Nothing Then
                        currentLocalConfig.model.Add(New ConfigItemModel() With {
                            .modelName = modelName,
                            .displayName = modelName,
                            .modelType = ModelType.Chat
                        })
                    End If
                Next

                RefreshLocalModelLists()
                MessageBox.Show($"已获取 {models.Count} 个模型", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("未获取到模型列表，请检查API端点是否正确", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"刷新模型列表失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            localRefreshModelsButton.Enabled = True
            localRefreshModelsButton.Text = "刷新列表"
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LocalChatModelCheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        If e.NewValue = CheckState.Checked Then
            For i = 0 To localChatModelCheckedListBox.Items.Count - 1
                If i <> e.Index Then
                    localChatModelCheckedListBox.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Sub LocalEmbeddingModelCheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs)
        If e.NewValue = CheckState.Checked Then
            For i = 0 To localEmbeddingModelCheckedListBox.Items.Count - 1
                If i <> e.Index Then
                    localEmbeddingModelCheckedListBox.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Async Sub LocalSaveButton_Click(sender As Object, e As EventArgs)
        If currentLocalConfig Is Nothing Then Return

        Dim platformName = localPlatformTextBox.Text
        Dim apiUrl = localUrlTextBox.Text
        If String.IsNullOrEmpty(platformName) Then
            MessageBox.Show("请输入服务名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If String.IsNullOrEmpty(apiUrl) OrElse Not (apiUrl.StartsWith("http://") OrElse apiUrl.StartsWith("https://")) Then
            MessageBox.Show("请输入有效的API端点 (以http://或https://开头)", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim apiKey = localApiKeyTextBox.Text
        If String.IsNullOrEmpty(apiKey) AndAlso Not String.IsNullOrEmpty(currentLocalConfig.defaultApiKey) Then
            apiKey = currentLocalConfig.defaultApiKey
        End If

        Dim selectedChatModelName As String = ""
        For i = 0 To localChatModelCheckedListBox.Items.Count - 1
            If localChatModelCheckedListBox.GetItemChecked(i) Then
                Dim model = CType(localChatModelCheckedListBox.Items(i), ConfigItemModel)
                selectedChatModelName = model.modelName
                Exit For
            End If
        Next

        If String.IsNullOrEmpty(selectedChatModelName) Then
            MessageBox.Show("请选择一个对话模型", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim selectedEmbeddingModelName As String = ""
        For i = 0 To localEmbeddingModelCheckedListBox.Items.Count - 1
            If localEmbeddingModelCheckedListBox.GetItemChecked(i) Then
                Dim model = CType(localEmbeddingModelCheckedListBox.Items(i), ConfigItemModel)
                selectedEmbeddingModelName = model.modelName
                Exit For
            End If
        Next

        localSaveButton.Enabled = False
        localSaveButton.Text = "验证中..."
        Cursor = Cursors.WaitCursor

        currentLocalConfig.pltform = platformName
        currentLocalConfig.url = apiUrl
        currentLocalConfig.key = apiKey

        Try
            Dim validationResult = Await ValidateApiAsync(apiUrl, apiKey, selectedChatModelName)
            If validationResult Then
                currentLocalConfig.validated = True
                currentLocalConfig.translateSelected = localTranslateCheckBox.Checked

                For Each model In currentLocalConfig.model
                    If model.modelType = ModelType.Chat Then
                        model.selected = (model.modelName = selectedChatModelName)
                    ElseIf model.modelType = ModelType.Embedding Then
                        model.selected = (model.modelName = selectedEmbeddingModelName)
                    End If
                Next

                For Each config In ConfigData
                    config.selected = (config Is currentLocalConfig)
                    If config IsNot currentLocalConfig Then
                        config.translateSelected = If(localTranslateCheckBox.Checked, False, config.translateSelected)
                    End If
                Next

                ConfigSettings.ApiUrl = currentLocalConfig.url
                ConfigSettings.ApiKey = apiKey
                ConfigSettings.platform = currentLocalConfig.pltform
                ConfigSettings.ModelName = selectedChatModelName

                Dim selectedChatModel = currentLocalConfig.model.FirstOrDefault(Function(m) m.modelName = selectedChatModelName)
                If selectedChatModel IsNot Nothing Then
                    ConfigSettings.mcpable = selectedChatModel.mcpable
                End If

                If Not String.IsNullOrEmpty(selectedEmbeddingModelName) Then
                    ConfigSettings.EmbeddingModel = selectedEmbeddingModelName
                End If

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
            localSaveButton.Enabled = True
            localSaveButton.Text = "验证并保存"
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub LocalDeleteButton_Click(sender As Object, e As EventArgs)
        If currentLocalConfig Is Nothing Then Return
        If currentLocalConfig.isPreset Then
            MessageBox.Show("预置配置不可删除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If MessageBox.Show($"确定要删除 {currentLocalConfig.pltform} 吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ConfigData.Remove(currentLocalConfig)
            SaveConfig()
            LoadDataToUI()
        End If
    End Sub

    Private Sub LocalAddButton_Click(sender As Object, e As EventArgs)
        Dim newConfig As New ConfigItem() With {
            .pltform = "新本地服务",
            .url = "http://localhost:11434/v1/chat/completions",
            .providerType = ProviderType.Local,
            .isPreset = False,
            .key = "",
            .defaultApiKey = "",
            .registerUrl = "",
            .translateSelected = True,
            .model = New List(Of ConfigItemModel)()
        }

        ConfigData.Add(newConfig)
        localProviderListBox.Items.Add(newConfig)
        localProviderListBox.SelectedItem = newConfig
    End Sub

#End Region

#Region "辅助方法"

    Private Async Function ValidateApiAsync(apiUrl As String, apiKey As String, modelName As String) As Task(Of Boolean)
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Using client As New HttpClient()
                client.Timeout = TimeSpan.FromSeconds(30)
                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", apiKey)

                ' 构造一个简单的聊天请求来验证API
                Dim requestBody = New JObject() From {
                    {"model", modelName},
                    {"messages", New JArray() From {
                        New JObject() From {
                            {"role", "user"},
                            {"content", "Hi"}
                        }
                    }},
                    {"max_tokens", 5}
                }

                Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(apiUrl, content)

                Return response.IsSuccessStatusCode
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub SaveConfig()
        Try
            ConfigManager.SaveConfig()
        Catch ex As Exception
            MessageBox.Show($"保存配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

    ' 记忆管理Tab控件
    Private chkUseContextBuilder As CheckBox
    Private chkEnableUserProfile As CheckBox
    Private numRagTopN As NumericUpDown
    Private numAtomicMaxLen As NumericUpDown
    Private numSessionSummaryLimit As NumericUpDown
    Private chkEnableAgenticSearch As CheckBox
    Private listMemory As ListBox
    Private txtMemoryContent As TextBox
    Private txtUserProfile As TextBox
    Private _memoryRecords As New List(Of AtomicMemoryRecord)()
    Private _memorySplitContainer As SplitContainer
    Private _memoryListSplitContainer As SplitContainer

    ''' <summary>
    ''' 初始化记忆管理Tab - 左右布局
    ''' </summary>
    Private Sub InitializeMemoryTab()
    ' 可拖拽分隔的左右布局
    _memorySplitContainer = New SplitContainer() With {
            .Location = New Point(10, 10),
            .Size = New Size(850, 490),
            .Panel1MinSize = 200,
            .Panel2MinSize = 300,
            .FixedPanel = FixedPanel.None
        }
    _memorySplitContainer.Panel1.SuspendLayout()
    _memorySplitContainer.Panel2.SuspendLayout()

    ' 左侧：记忆配置
    Dim lblConfigTitle As New Label() With {
            .Text = "记忆配置",
            .Location = New Point(10, 10),
            .Size = New Size(260, 20),
            .Font = New Font(Me.Font.FontFamily, 10, FontStyle.Bold)
        }
    _memorySplitContainer.Panel1.Controls.Add(lblConfigTitle)

    Dim y As Integer = 40
    chkUseContextBuilder = New CheckBox() With {
            .Text = "使用 ContextBuilder（分层组装 Memory/Skills）",
            .Location = New Point(10, y),
            .Size = New Size(260, 24),
            .Checked = MemoryConfig.UseContextBuilder
        }
    _memorySplitContainer.Panel1.Controls.Add(chkUseContextBuilder)
    y += 32

    chkEnableUserProfile = New CheckBox() With {
            .Text = "启用用户画像",
            .Location = New Point(10, y),
            .Size = New Size(200, 24),
            .Checked = MemoryConfig.EnableUserProfile
        }
    AddHandler chkEnableUserProfile.CheckedChanged, Sub(s, e)
                                                        txtUserProfile.Enabled = chkEnableUserProfile.Checked
                                                    End Sub
    _memorySplitContainer.Panel1.Controls.Add(chkEnableUserProfile)
    y += 28

    Dim lblRag As New Label() With {.Text = "RAG 检索条数 (1-20)：", .Location = New Point(10, y + 2), .Size = New Size(160, 20)}
    _memorySplitContainer.Panel1.Controls.Add(lblRag)
    numRagTopN = New NumericUpDown() With {
            .Location = New Point(175, y),
            .Size = New Size(80, 24),
            .Minimum = 1,
            .Maximum = 20,
            .Value = MemoryConfig.RagTopN
        }
    _memorySplitContainer.Panel1.Controls.Add(numRagTopN)
    y += 32

    Dim lblAtomic As New Label() With {.Text = "记忆片段最大长度 (50-500)：", .Location = New Point(10, y + 2), .Size = New Size(160, 20)}
    _memorySplitContainer.Panel1.Controls.Add(lblAtomic)
    numAtomicMaxLen = New NumericUpDown() With {
            .Location = New Point(175, y),
            .Size = New Size(80, 24),
            .Minimum = 50,
            .Maximum = 500,
            .Value = MemoryConfig.AtomicContentMaxLength
        }
    _memorySplitContainer.Panel1.Controls.Add(numAtomicMaxLen)
    y += 32

    Dim lblSummary As New Label() With {.Text = "近期会话摘要条数 (1-15)：", .Location = New Point(10, y + 2), .Size = New Size(160, 20)}
    _memorySplitContainer.Panel1.Controls.Add(lblSummary)
    numSessionSummaryLimit = New NumericUpDown() With {
            .Location = New Point(175, y),
            .Size = New Size(80, 24),
            .Minimum = 1,
            .Maximum = 15,
            .Value = MemoryConfig.SessionSummaryLimit
        }
    _memorySplitContainer.Panel1.Controls.Add(numSessionSummaryLimit)
    y += 32

    chkEnableAgenticSearch = New CheckBox() With {
            .Text = "启用 MCP 记忆搜索（Agentic Search）",
            .Location = New Point(10, y),
            .Size = New Size(260, 24),
            .Checked = MemoryConfig.EnableAgenticSearch
        }
    _memorySplitContainer.Panel1.Controls.Add(chkEnableAgenticSearch)
    y += 40

    ' 用户画像编辑区
    Dim lblProfile As New Label() With {.Text = "用户画像内容：", .Location = New Point(10, y), .Size = New Size(200, 20)}
    _memorySplitContainer.Panel1.Controls.Add(lblProfile)
    y += 22
    txtUserProfile = New TextBox() With {
            .Location = New Point(10, y),
            .Size = New Size(250, 120),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Enabled = MemoryConfig.EnableUserProfile
        }
    _memorySplitContainer.Panel1.Controls.Add(txtUserProfile)
    y += 130

    ' 保存配置按钮
    Dim btnSaveConfig As New Button() With {
            .Text = "保存配置",
            .Location = New Point(10, y),
            .Size = New Size(120, 30),
            .BackColor = Color.FromArgb(70, 130, 180),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
    AddHandler btnSaveConfig.Click, AddressOf SaveMemoryConfigClick
    _memorySplitContainer.Panel1.Controls.Add(btnSaveConfig)

    ' 右侧：记忆片段和用户画像（Tab形式）
    Dim tabControl As New TabControl() With {
            .Dock = DockStyle.Fill
        }

    ' Tab 1: 记忆片段（可拖拽分隔条）
    Dim tabMemory As New TabPage("记忆片段")
    _memoryListSplitContainer = New SplitContainer() With {
            .Location = New Point(0, 0),
            .Size = New Size(640, 440),          ' 合理非零大小
            .SplitterDistance = 260,             ' 在 Panel1MinSize 与 Width - Panel2MinSize 之间
            .Panel1MinSize = 100,
            .Panel2MinSize = 150,
            .Dock = DockStyle.Fill
        }

    Dim lblList As New Label() With {.Text = "记忆片段列表（最近 100 条）", .Location = New Point(5, 5), .Size = New Size(200, 20)}
    _memoryListSplitContainer.Panel1.Controls.Add(lblList)
    listMemory = New ListBox() With {
            .Location = New Point(5, 28),
            .Size = New Size(250, 420),
            .DisplayMember = "DisplayText",
            .HorizontalScrollbar = True,
            .HorizontalExtent = 3000
        }
    AddHandler listMemory.SelectedIndexChanged, AddressOf MemorySelectionChanged
    _memoryListSplitContainer.Panel1.Controls.Add(listMemory)

    txtMemoryContent = New TextBox() With {
            .Location = New Point(5, 5),
            .Size = New Size(300, 380),
            .Multiline = True,
            .ScrollBars = ScrollBars.Both,
            .ReadOnly = True
        }
    _memoryListSplitContainer.Panel2.Controls.Add(txtMemoryContent)
    Dim btnRefreshMemory As New Button() With {.Text = "刷新", .Location = New Point(5, 390), .Size = New Size(70, 28)}
    AddHandler btnRefreshMemory.Click, AddressOf LoadMemories
    _memoryListSplitContainer.Panel2.Controls.Add(btnRefreshMemory)
    Dim btnDeleteMemory As New Button() With {.Text = "删除选中", .Location = New Point(85, 390), .Size = New Size(80, 28)}
    AddHandler btnDeleteMemory.Click, AddressOf BtnDeleteMemoryClick
    _memoryListSplitContainer.Panel2.Controls.Add(btnDeleteMemory)
    Dim btnCopyMemory As New Button() With {.Text = "复制选中", .Location = New Point(175, 390), .Size = New Size(70, 28)}
    AddHandler btnCopyMemory.Click, Sub(s, ev)
                                        If listMemory.SelectedItem IsNot Nothing Then
                                            Try
                                                Clipboard.SetText(listMemory.SelectedItem.ToString())
                                            Catch ex As Exception
                                            End Try
                                        End If
                                    End Sub
    _memoryListSplitContainer.Panel2.Controls.Add(btnCopyMemory)
    tabMemory.Controls.Add(_memoryListSplitContainer)
    tabControl.TabPages.Add(tabMemory)

    _memorySplitContainer.Panel2.Controls.Add(tabControl)

    _memorySplitContainer.Panel1.ResumeLayout(False)
    _memorySplitContainer.Panel2.ResumeLayout(False)
    memoryTab.Controls.Add(_memorySplitContainer)

    ' 加载数据
    LoadMemoryConfig()
    LoadMemories()
End Sub

''' <summary>
''' 加载记忆配置
''' </summary>
Private Sub LoadMemoryConfig()
    chkUseContextBuilder.Checked = MemoryConfig.UseContextBuilder
    chkEnableUserProfile.Checked = MemoryConfig.EnableUserProfile
    numRagTopN.Value = MemoryConfig.RagTopN
    numAtomicMaxLen.Value = MemoryConfig.AtomicContentMaxLength
    numSessionSummaryLimit.Value = MemoryConfig.SessionSummaryLimit
    chkEnableAgenticSearch.Checked = MemoryConfig.EnableAgenticSearch
    txtUserProfile.Enabled = MemoryConfig.EnableUserProfile
    Try
        txtUserProfile.Text = MemoryRepository.GetUserProfile()
    Catch
        txtUserProfile.Text = ""
    End Try
End Sub

''' <summary>
''' 保存记忆配置
''' </summary>
Private Sub SaveMemoryConfigClick(sender As Object, e As EventArgs)
    Try
        MemoryConfig.UseContextBuilder = chkUseContextBuilder.Checked
        MemoryConfig.EnableUserProfile = chkEnableUserProfile.Checked
        MemoryConfig.RagTopN = CInt(numRagTopN.Value)
        MemoryConfig.AtomicContentMaxLength = CInt(numAtomicMaxLen.Value)
        MemoryConfig.SessionSummaryLimit = CInt(numSessionSummaryLimit.Value)
        MemoryConfig.EnableAgenticSearch = chkEnableAgenticSearch.Checked
        If chkEnableUserProfile.Checked Then
            MemoryRepository.UpdateUserProfile(txtUserProfile.Text)
        End If
        MessageBox.Show("记忆配置已保存", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
    Catch ex As Exception
        MessageBox.Show("保存失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
End Sub

''' <summary>
''' 加载记忆片段
''' </summary>
Private Sub LoadMemories()
    Try
        OfficeAiDatabase.EnsureInitialized()
        _memoryRecords = MemoryRepository.ListAtomicMemories(100, 0)
        listMemory.DataSource = Nothing
        listMemory.Items.Clear()
        For Each r In _memoryRecords
            Dim preview = If(r.Content?.Length > 40, r.Content.Substring(0, 40) & "...", r.Content)
            listMemory.Items.Add(New MemoryItem With {.Record = r, .DisplayText = $"[{r.CreateTime}] {preview}"})
        Next
    Catch ex As Exception
        listMemory.Items.Clear()
        listMemory.Items.Add("(加载失败: " & ex.Message & ")")
    End Try
End Sub

''' <summary>
''' 记忆片段选中事件
''' </summary>
Private Sub MemorySelectionChanged(sender As Object, e As EventArgs)
    Dim item = TryCast(listMemory.SelectedItem, MemoryItem)
    If item Is Nothing Then
        txtMemoryContent.Text = ""
        Return
    End If
    txtMemoryContent.Text = item.Record.Content
End Sub

''' <summary>
''' 删除记忆片段
''' </summary>
Private Sub BtnDeleteMemoryClick(sender As Object, e As EventArgs)
    Dim item = TryCast(listMemory.SelectedItem, MemoryItem)
    If item Is Nothing Then
        MessageBox.Show("请先选择一条记录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return
    End If
    If MessageBox.Show("确定删除此条记忆片段？", "确认", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Return
    Try
        MemoryRepository.DeleteAtomicMemory(item.Record.Id)
        LoadMemories()
    Catch ex As Exception
        MessageBox.Show("删除失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
End Sub

''' <summary>
''' 记忆片段列表项
''' </summary>
Private Class MemoryItem
    Public Property Record As AtomicMemoryRecord
    Public Property DisplayText As String
    Public Overrides Function ToString() As String
        Return DisplayText
    End Function
End Class

''' <summary>
''' 获取完整异常信息
''' </summary>
Private Shared Function GetFullExceptionMessage(ex As Exception) As String
    Dim sb As New StringBuilder()
    Dim current As Exception = ex
    Dim depth As Integer = 0
    While current IsNot Nothing AndAlso depth < 5
        If depth > 0 Then sb.Append(" <- ")
        sb.Append(current.GetType().Name).Append(": ").Append(current.Message)
        current = current.InnerException
        depth += 1
    End While
    Return sb.ToString()
End Function

End Class