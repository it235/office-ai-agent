Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports Newtonsoft.Json
Imports AiHelper.ConfigManager

' 大模型提示词配置
Public Class ConfigPromptForm
    Inherits Form
    Public Shared Property ConfigPromptData As List(Of PromptConfigItem)

    ' 默认配置文件在当前用户，我的文档下
    Private Shared configFileName As String = "office_ai_prompt_config.json"
    Private Shared configFilePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        ConfigSettings.OfficeAiAppDataFolder, configFileName)

    Private descriptionLabel1 As Label

    Private currentPromptComboBox As ComboBox
    Private editConfigButton As Button
    Private promptContentBox As TextBox
    Private confirmButton As Button

    Private addConfigButton As Button
    Private newPromptComboBox As TextBox
    Private newPromptContent As TextBox
    Private saveConfigButton As Button

    Public Property propmtName As String
    Public Property propmtContent As String
    Public Property VBA_Q As String = "你是一名优秀的Excel资深专家，擅长写VBA代码。如果你有输出代码，请你仔细思考和检查，没有数据或格式不对的行可以跳过用处理，同时我不需要任何通篇大论的废话，请你智能快速一点。"

    Public Property EXCEL_TAB_Q As String = "你是一名优秀的Excel资深专家，擅长通过VBA代码根据数据生成各种图表，例如：饼图、折线图、柱状图，请根据我的问题给出我需要的VBA代码。同时我不需要任何通篇大论的废话，请你智能快速一点。"

    Public Sub LoadConfig()
        ' 初始化配置数据
        ConfigPromptData = New List(Of PromptConfigItem)()

        Dim vbap = New PromptConfigItem() With {
                .name = "VBA专家身份",
                .content = VBA_Q,
                .selected = True
            }

        Dim excelTabP = New PromptConfigItem() With {
                .name = "Excel表格专家身份",
                .content = EXCEL_TAB_Q,
                .selected = False
            }
        ' 添加默认配置
        If Not File.Exists(configFilePath) Then
            ConfigPromptData.Add(vbap)
            ConfigPromptData.Add(excelTabP)
        Else
            ' 加载自定义配置
            Dim json As String = File.ReadAllText(configFilePath)
            ConfigPromptData = JsonConvert.DeserializeObject(Of List(Of PromptConfigItem))(json)
        End If

        ' 初始化配置，将数据初始化到 ConfigSettings，方便全局调用
        For Each item In ConfigPromptData
            If item.selected Then
                ConfigSettings.propmtName = item.name
                ConfigSettings.propmtContent = item.content
            End If
        Next
    End Sub


    Public Sub New()
        LoadConfig()

        ' 初始化表单
        Me.Text = "配置聊天大模型提示词"
        Me.Size = New Size(480, 550)
        Me.StartPosition = FormStartPosition.CenterScreen ' 设置表单居中显示

        descriptionLabel1 = New Label()
        descriptionLabel1.Text = "提示词相当于给AI设定对应的身份，这样才能找到该领域的问题，回答起来更专业，例如：你是一名Excel VBA专家，接下来的问题都和Excel以及VBA相关"
        descriptionLabel1.Dock = DockStyle.Top
        descriptionLabel1.Height = 40
        descriptionLabel1.Margin = New Padding(10, 10, 10, 10)
        descriptionLabel1.TextAlign = ContentAlignment.MiddleLeft
        Me.Controls.Add(descriptionLabel1)

        ' 初始化模型选择 ComboBox
        currentPromptComboBox = New ComboBox()
        currentPromptComboBox.DisplayMember = "name"
        currentPromptComboBox.ValueMember = "value"
        currentPromptComboBox.Location = New Point(10, 50)
        currentPromptComboBox.Size = New Size(260, 30)
        AddHandler currentPromptComboBox.SelectedIndexChanged, AddressOf propmtCombBox_SelectedIndexChanged
        Me.Controls.Add(currentPromptComboBox)

        ' 初始化编辑配置按钮
        editConfigButton = New Button()
        editConfigButton.Text = "修改"
        editConfigButton.Font = New Font(editConfigButton.Font.FontFamily, 8) ' 设置字体大小
        editConfigButton.Location = New Point(280, 50)
        editConfigButton.Size = New Size(40, currentPromptComboBox.Height + 2)
        AddHandler editConfigButton.Click, AddressOf EditConfigButton_Click

        Me.Controls.Add(editConfigButton)


        ' 用来接收之前选择的提示词名称和提示词内容
        Dim propmtNameForDB As String
        Dim propmtContentForDB As String

        For Each config In ConfigPromptData
            If config.selected Then
                propmtNameForDB = config.name
                propmtContentForDB = config.content
            End If
        Next

        ' 提示词内容预览框
        promptContentBox = New TextBox()
        promptContentBox.Multiline = True
        promptContentBox.ScrollBars = ScrollBars.Vertical
        promptContentBox.Text = propmtContentForDB
        promptContentBox.ForeColor = Color.Gray
        promptContentBox.Location = New Point(10, 80)
        promptContentBox.Size = New Size(360, 120)
        promptContentBox.ReadOnly = True
        Me.Controls.Add(promptContentBox)


        ' 初始化确认按钮
        confirmButton = New Button()
        confirmButton.Text = "使用该提示词"
        confirmButton.Location = New Point(50, 210)
        confirmButton.Size = New Size(100, 30)
        AddHandler confirmButton.Click, AddressOf ConfirmButton_Click
        Me.Controls.Add(confirmButton)

        ' 初始化添加配置按钮
        addConfigButton = New Button()
        addConfigButton.Text = "添加新提示词"
        addConfigButton.Location = New Point(170, 210)
        addConfigButton.Size = New Size(100, 30)
        AddHandler addConfigButton.Click, AddressOf AddConfigButton_Click
        Me.Controls.Add(addConfigButton)

        ' 初始化新配置控件
        newPromptComboBox = New TextBox()
        newPromptComboBox.Text = NEW_NAME_C
        newPromptComboBox.ForeColor = Color.Gray
        newPromptComboBox.Location = New Point(10, 250)
        newPromptComboBox.Size = New Size(260, 30)
        newPromptComboBox.Visible = False
        AddHandler newPromptComboBox.Enter, AddressOf NewModelPlatformTextBox_Enter
        AddHandler newPromptComboBox.Leave, AddressOf NewModelPlatformTextBox_Leave
        Me.Controls.Add(newPromptComboBox)

        newPromptContent = New TextBox()
        newPromptContent.Multiline = True
        newPromptContent.ScrollBars = ScrollBars.Vertical
        'newPromptContent.Text = If(String.IsNullOrEmpty(propmtContentForDB), "输入提示词内容", propmtContentForDB)
        newPromptContent.ForeColor = If(String.IsNullOrEmpty(propmtContentForDB), Color.Gray, Color.Black)
        newPromptContent.Location = New Point(10, 290)
        newPromptContent.Size = New Size(360, 120)
        newPromptContent.Visible = False
        AddHandler newPromptContent.Enter, AddressOf ApiKeyTextBox_Enter ' 添加 Enter 事件处理程序
        AddHandler newPromptContent.Leave, AddressOf ApiKeyTextBox_Leave ' 添加 Leave 事件处理程序
        Me.Controls.Add(newPromptContent)

        saveConfigButton = New Button()
        saveConfigButton.Text = "保存"
        saveConfigButton.Location = New Point(100, 420)
        saveConfigButton.Size = New Size(100, 30)
        saveConfigButton.Visible = False
        AddHandler saveConfigButton.Click, AddressOf SaveConfigButton_Click
        Me.Controls.Add(saveConfigButton)

        ' 加载配置到复选框
        For Each configItem In ConfigPromptData
            currentPromptComboBox.Items.Add(configItem)
        Next

        ' 设置之前选择的模型
        If Not String.IsNullOrEmpty(propmtNameForDB) Then
            For i As Integer = 0 To currentPromptComboBox.Items.Count - 1
                If CType(currentPromptComboBox.Items(i), PromptConfigItem).name = propmtNameForDB Then
                    currentPromptComboBox.SelectedIndex = i
                    Exit For
                End If
            Next
        Else
            If currentPromptComboBox.Items.Count > 0 Then
                currentPromptComboBox.SelectedIndex = 0
            End If
        End If

        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub


    Private Sub EditConfigButton_Click(sender As Object, e As EventArgs)
        ' 获取选中的模型提示词
        Dim selectedPlatform As PromptConfigItem = CType(currentPromptComboBox.SelectedItem, PromptConfigItem)

        ' 将选中的数据带入到新配置控件中
        newPromptComboBox.Text = selectedPlatform.name
        newPromptComboBox.ForeColor = Color.Black

        newPromptContent.Text = selectedPlatform.content
        newPromptContent.ForeColor = Color.Black

        ' 显示新配置控件
        Me.Size = New Size(480, 550)
        newPromptComboBox.Visible = True
        newPromptContent.Visible = True
        saveConfigButton.Visible = True
    End Sub


    ' 切换提示词后的确认按钮
    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs)

        ' 获取选中的提示词名称和提升词内容
        Dim selectedPlatform As PromptConfigItem = CType(currentPromptComboBox.SelectedItem, PromptConfigItem)
        Dim name As String = selectedPlatform.name
        Dim content As String = selectedPlatform.content

        ' 重置选择后的selected属性和key
        For Each config In ConfigPromptData
            config.selected = False
            If selectedPlatform.name = config.name Then
                config.selected = True
                config.name = name
                config.content = content
            End If
        Next

        ' 保存到文件
        SaveConfig()

        ' 刷新内存中的api配置
        ConfigSettings.propmtName = name
        ConfigSettings.propmtContent = content

        ' 关闭对话框
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub


    Public Shared Sub SaveConfig()
        Dim json As String = JsonConvert.SerializeObject(ConfigPromptData, Formatting.Indented)
        ' 如果configFilePath的目录不存在就创建
        Dim dir = Path.GetDirectoryName(configFilePath)
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
        '如果文件不存在就创建
        If Not File.Exists(configFilePath) Then
            File.Create(configFilePath).Dispose()
        End If
        File.WriteAllText(configFilePath, json)
    End Sub


    Private Sub propmtCombBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' 根据选中的提示词名称显示不同的内容
        Dim selectedModel As PromptConfigItem = CType(currentPromptComboBox.SelectedItem, PromptConfigItem)
        promptContentBox.Clear()
        promptContentBox.Text = selectedModel.content
        'promptContentBox.ForeColor = Color.Black
    End Sub

    Private Sub AddConfigButton_Click(sender As Object, e As EventArgs)
        ' 显示新配置控件
        Me.Size = New Size(480, 550)
        newPromptComboBox.Visible = True
        newPromptContent.Visible = True
        saveConfigButton.Visible = True
    End Sub


    Private Sub SaveConfigButton_Click(sender As Object, e As EventArgs)
        ' 获取新配置
        Dim name As String = newPromptComboBox.Text
        Dim content As String = newPromptContent.Text

        If String.IsNullOrWhiteSpace(name) Or name = NEW_NAME_C Then
            GlobalStatusStrip.ShowWarning("请输入提示词名称！")
            Return
        End If

        If String.IsNullOrWhiteSpace(content) Then
            GlobalStatusStrip.ShowWarning("请输入提示词内容！")
            Return
        End If

        ' 检查是否存在相同的 propmtName
        Dim existingItem As PromptConfigItem = ConfigPromptData.FirstOrDefault(Function(item) item.name = name)
        If existingItem IsNot Nothing Then
            ' 更新已有的 propmtName 数据
            existingItem.name = name
            existingItem.content = content
            existingItem.selected = True
        Else
            ' 用户本地新增模型到 ComboBox
            Dim newItem As New PromptConfigItem() With {
                .name = name,
                .content = content,
                .selected = True
            }
            ConfigPromptData.Add(newItem)
            currentPromptComboBox.Items.Add(newItem)
            currentPromptComboBox.SelectedItem = newItem
        End If

        promptContentBox.Text = content

        newPromptComboBox.Clear()
        newPromptContent.Clear()

        ' 保存到文件
        SaveConfig()

        ConfigSettings.propmtContent = content
        ConfigSettings.propmtName = name


        Me.Size = New Size(480, 550)
        newPromptComboBox.Visible = False
        newPromptContent.Visible = False
        saveConfigButton.Visible = False
    End Sub

    Private Property NEW_NAME_C As String = "取个响亮的名称，例如：Excel函数专家"

    Private Sub NewModelPlatformTextBox_Enter(sender As Object, e As EventArgs)
        If newPromptComboBox.Text = NEW_NAME_C Then
            newPromptComboBox.Text = ""
            newPromptComboBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub NewModelPlatformTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(newPromptComboBox.Text) Then
            newPromptComboBox.Text = NEW_NAME_C
            newPromptComboBox.ForeColor = Color.Gray
        End If
    End Sub

    Private Property NEW_CONTENT_C As String = "输入大模型提示词内容，为其设定一个身份，例如：你是一个非常厉害的Excel大师，擅长各种VBA代码"
    Private Sub ApiKeyTextBox_Enter(sender As Object, e As EventArgs)
        If newPromptContent.Text = NEW_CONTENT_C Then
            newPromptContent.Text = ""
            newPromptContent.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ApiKeyTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(promptContentBox.Text) Then
            newPromptContent.Text = NEW_CONTENT_C
            newPromptContent.ForeColor = Color.Gray
        End If
    End Sub


    ' 提示词配置（每次仅可使用1个）
    Public Class PromptConfigItem
        Public Property name As String
        Public Property content As String
        Public Property selected As Boolean
        Public Overrides Function ToString() As String
            Return content
        End Function
    End Class
End Class

