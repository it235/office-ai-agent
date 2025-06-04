Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports Newtonsoft.Json
Imports ShareRibbon.ConfigManager

Public Class ConfigApiForm
    Inherits Form

    Private modelComboBox As ComboBox
    ' 编辑按钮
    Private editConfigButton As Button
    Private apiKeyTextBox As TextBox
    Private modelNameComboBox As ComboBox
    Private confirmButton As Button
    Private addConfigButton As Button
    Private newModelPlatformTextBox As TextBox
    Private newApiUrlTextBox As TextBox
    Private newModelNameTextBoxes As List(Of TextBox)
    Private addModelNameButton As Button
    Private saveConfigButton As Button


    Public Property platform As String
    Public Property apiUrl As String
    Public Property apiKey As String
    Public Property modelName As String


    Public Sub New()
        ' 初始化表单
        Me.Text = "配置大模型API"
        Me.Size = New Size(350, 350)
        Me.StartPosition = FormStartPosition.CenterScreen ' 设置表单居中显示

        ' 初始化模型选择 ComboBox
        modelComboBox = New ComboBox()
        modelComboBox.DisplayMember = "pltform"
        modelComboBox.ValueMember = "url"
        modelComboBox.Location = New Point(10, 10)
        modelComboBox.Size = New Size(260, 30)
        AddHandler modelComboBox.SelectedIndexChanged, AddressOf ModelComboBox_SelectedIndexChanged
        Me.Controls.Add(modelComboBox)

        ' 初始化编辑配置按钮
        editConfigButton = New Button()
        editConfigButton.Text = "修改"
        editConfigButton.Font = New Font(editConfigButton.Font.FontFamily, 8) ' 设置字体大小
        editConfigButton.Location = New Point(280, 10)
        editConfigButton.Size = New Size(40, modelComboBox.Height + 2)
        AddHandler editConfigButton.Click, AddressOf EditConfigButton_Click
        Me.Controls.Add(editConfigButton)

        ' 初始化模型名称选择 ComboBox
        modelNameComboBox = New ComboBox()
        modelNameComboBox.Location = New Point(10, 50)
        modelNameComboBox.Size = New Size(260, 30)
        modelNameComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        modelNameComboBox.AutoCompleteSource = AutoCompleteSource.ListItems
        Me.Controls.Add(modelNameComboBox)

        ' 用来接收之前选择的模型和 API Key
        Dim platformForDB As String
        Dim apiUrlForDB As String
        Dim apiKeyForDB As String
        Dim modelNameForDB As String

        For Each config In ConfigData
            If config.selected Then
                platformForDB = config.pltform
                apiKeyForDB = config.key
                apiUrlForDB = config.url
                For Each item_m In config.model
                    If item_m.selected Then
                        modelNameForDB = item_m.modelName
                    End If
                Next
            End If
        Next

        ' 初始化 API Key 输入框
        apiKeyTextBox = New TextBox()
        apiKeyTextBox.Text = If(String.IsNullOrEmpty(apiKeyForDB), "输入 API Key", apiKeyForDB)
        apiKeyTextBox.ForeColor = If(String.IsNullOrEmpty(apiKeyForDB), Color.Gray, Color.Black)
        apiKeyTextBox.Location = New Point(10, 90)
        apiKeyTextBox.Size = New Size(260, 30)
        AddHandler apiKeyTextBox.Enter, AddressOf ApiKeyTextBox_Enter ' 添加 Enter 事件处理程序
        AddHandler apiKeyTextBox.Leave, AddressOf ApiKeyTextBox_Leave ' 添加 Leave 事件处理程序
        Me.Controls.Add(apiKeyTextBox)

        ' 初始化确认按钮
        confirmButton = New Button()
        confirmButton.Text = "确认"
        confirmButton.Location = New Point(100, 130)
        confirmButton.Size = New Size(100, 30)
        AddHandler confirmButton.Click, AddressOf ConfirmButton_Click
        Me.Controls.Add(confirmButton)

        ' 初始化添加配置按钮
        addConfigButton = New Button()
        addConfigButton.Text = "添加模型配置"
        addConfigButton.Location = New Point(100, 170)
        addConfigButton.Size = New Size(100, 30)
        AddHandler addConfigButton.Click, AddressOf AddConfigButton_Click
        Me.Controls.Add(addConfigButton)

        ' 初始化新配置控件
        newModelPlatformTextBox = New TextBox()
        newModelPlatformTextBox.Text = "模型平台"
        newModelPlatformTextBox.ForeColor = Color.Gray
        newModelPlatformTextBox.Location = New Point(10, 210)
        newModelPlatformTextBox.Size = New Size(260, 30)
        newModelPlatformTextBox.Visible = False
        AddHandler newModelPlatformTextBox.Enter, AddressOf NewModelPlatformTextBox_Enter
        AddHandler newModelPlatformTextBox.Leave, AddressOf NewModelPlatformTextBox_Leave
        Me.Controls.Add(newModelPlatformTextBox)

        newApiUrlTextBox = New TextBox()
        newApiUrlTextBox.Text = "API URL"
        newApiUrlTextBox.ForeColor = Color.Gray
        newApiUrlTextBox.Location = New Point(10, 250)
        newApiUrlTextBox.Size = New Size(260, 30)
        newApiUrlTextBox.Visible = False
        AddHandler newApiUrlTextBox.Enter, AddressOf NewApiUrlTextBox_Enter
        AddHandler newApiUrlTextBox.Leave, AddressOf NewApiUrlTextBox_Leave
        Me.Controls.Add(newApiUrlTextBox)

        newModelNameTextBoxes = New List(Of TextBox)()
        AddNewModelNameTextBox(False)

        addModelNameButton = New Button()
        addModelNameButton.Text = "+"
        addModelNameButton.Location = New Point(280, 290)
        addModelNameButton.Size = New Size(20, 20)
        addModelNameButton.Visible = False
        AddHandler addModelNameButton.Click, AddressOf AddModelNameButton_Click
        Me.Controls.Add(addModelNameButton)

        saveConfigButton = New Button()
        saveConfigButton.Text = "保存"
        saveConfigButton.Location = New Point(100, 420)
        saveConfigButton.Size = New Size(100, 30)
        saveConfigButton.Visible = False
        AddHandler saveConfigButton.Click, AddressOf SaveConfigButton_Click
        Me.Controls.Add(saveConfigButton)

        ' 加载配置到复选框
        For Each configItem In ConfigData
            modelComboBox.Items.Add(configItem)
        Next

        ' 设置之前选择的模型
        If Not String.IsNullOrEmpty(platformForDB) Then
            For i As Integer = 0 To modelComboBox.Items.Count - 1
                If CType(modelComboBox.Items(i), ConfigManager.ConfigItem).pltform = platformForDB Then
                    modelComboBox.SelectedIndex = i
                    Exit For
                End If
            Next
        Else
            If modelComboBox.Items.Count > 0 Then
                modelComboBox.SelectedIndex = 0
            End If
        End If

        ' 设置之前选择的模型名称
        If Not String.IsNullOrEmpty(modelNameForDB) Then
            For i As Integer = 0 To modelNameComboBox.Items.Count - 1
                If modelNameComboBox.Items(i).ToString() = modelNameForDB Then
                    modelNameComboBox.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        ' 设置之前的 API Key
        If Not String.IsNullOrEmpty(apiKeyForDB) Then
            apiKeyTextBox.Text = apiKeyForDB
            apiKeyTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ApiKeyTextBox_Enter(sender As Object, e As EventArgs)
        If apiKeyTextBox.Text = "输入 API Key" Then
            apiKeyTextBox.Text = ""
            apiKeyTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ApiKeyTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(apiKeyTextBox.Text) Then
            apiKeyTextBox.Text = "输入 API Key"
            apiKeyTextBox.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub EditConfigButton_Click(sender As Object, e As EventArgs)
        ' 获取选中的模型和 API Key
        Dim selectedPlatform As ConfigManager.ConfigItem = CType(modelComboBox.SelectedItem, ConfigManager.ConfigItem)
        Dim selectedModelName As String = If(modelNameComboBox.SelectedItem IsNot Nothing, modelNameComboBox.SelectedItem.ToString(), modelNameComboBox.Text)

        ' 将选中的数据带入到新配置控件中
        newModelPlatformTextBox.Text = selectedPlatform.pltform
        newModelPlatformTextBox.ForeColor = Color.Black
        newApiUrlTextBox.Text = selectedPlatform.url
        newApiUrlTextBox.ForeColor = Color.Black

        ' 清空并重新添加 newModelNameTextBoxes
        For Each textBox In newModelNameTextBoxes
            Me.Controls.Remove(textBox)
        Next
        newModelNameTextBoxes.Clear()

        For Each model In selectedPlatform.model
            AddNewModelNameTextBox(True)
            Dim newModelNameTextBox = newModelNameTextBoxes.Last()
            newModelNameTextBox.Text = model.modelName
            newModelNameTextBox.ForeColor = Color.Black
            If model.modelName = selectedModelName Then
                newModelNameTextBox.BackColor = Color.LightBlue ' 标记选中的模型名称
            End If
        Next

        ' 显示新配置控件
        Me.Size = New Size(350, 500)
        newModelPlatformTextBox.Visible = True
        newApiUrlTextBox.Visible = True
        For Each textBox In newModelNameTextBoxes
            textBox.Visible = True
        Next
        addModelNameButton.Visible = True
        saveConfigButton.Visible = True
    End Sub


    ' 切换大模型后的确认按钮
    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs)

        ' 获取选中的模型和 API Key
        Dim selectedPlatform As ConfigManager.ConfigItem = CType(modelComboBox.SelectedItem, ConfigManager.ConfigItem)
        Dim apiUrl As String = selectedPlatform.url
        Dim selectedModelName As String = If(modelNameComboBox.SelectedItem IsNot Nothing, modelNameComboBox.SelectedItem.ToString(), modelNameComboBox.Text)
        Dim inputApiKey As String = apiKeyTextBox.Text


        ' 更新选中的模型的 API Key
        'selectedItem.key = inputApiKey

        ' 重置选择后的selected属性和key
        For Each config In ConfigData
            config.selected = False
            If selectedPlatform.pltform = config.pltform Then
                config.selected = True
                config.key = inputApiKey
                For Each item_m In config.model
                    item_m.selected = False
                    If item_m.modelName = selectedModelName Then
                        item_m.selected = True
                    End If
                Next
            End If

        Next

        ' 保存到文件
        SaveConfig()

        ' 刷新内存中的api配置
        ConfigSettings.ApiUrl = apiUrl
        ConfigSettings.ApiKey = inputApiKey
        ConfigSettings.platform = selectedPlatform.pltform
        ConfigSettings.ModelName = selectedModelName

        ' 关闭对话框
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub ModelComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' 根据选中的模型更新模型名称选择 ComboBox
        modelNameComboBox.Items.Clear()
        Dim selectedModel As ConfigManager.ConfigItem = CType(modelComboBox.SelectedItem, ConfigManager.ConfigItem)
        For Each ModelNameT In selectedModel.model
            modelNameComboBox.Items.Add(ModelNameT)
        Next
        If modelNameComboBox.Items.Count > 0 Then
            modelNameComboBox.SelectedIndex = 0
        End If

        ' 更新 API Key
        apiKeyTextBox.Text = selectedModel.key
        apiKeyTextBox.ForeColor = If(String.IsNullOrEmpty(selectedModel.key), Color.Gray, Color.Black)
    End Sub

    Private Sub AddConfigButton_Click(sender As Object, e As EventArgs)
        ' 显示新配置控件
        Me.Size = New Size(350, 500)
        newModelPlatformTextBox.Visible = True
        newApiUrlTextBox.Visible = True
        For Each textBox In newModelNameTextBoxes
            textBox.Visible = True
        Next
        addModelNameButton.Visible = True
        saveConfigButton.Visible = True

    End Sub

    Private Sub AddModelNameButton_Click(sender As Object, e As EventArgs)
        AddNewModelNameTextBox(True)
    End Sub

    Private Sub AddNewModelNameTextBox(display As Boolean)
        Dim newModelNameTextBox As New TextBox()
        newModelNameTextBox.Text = "具体模型"
        newModelNameTextBox.ForeColor = Color.Gray
        newModelNameTextBox.Location = New Point(10, 290 + newModelNameTextBoxes.Count * 40)
        newModelNameTextBox.Size = New Size(260, 30)
        newModelNameTextBox.Visible = display
        AddHandler newModelNameTextBox.Enter, AddressOf NewModelNameTextBox_Enter
        AddHandler newModelNameTextBox.Leave, AddressOf NewModelNameTextBox_Leave
        Me.Controls.Add(newModelNameTextBox)
        newModelNameTextBoxes.Add(newModelNameTextBox)

        ' 只有第二行及之后的行才添加减号按钮
        If newModelNameTextBoxes.Count > 1 Then
            Dim removeButton As New Button()
            removeButton.Text = "-"
            removeButton.Location = New Point(280, 290 + (newModelNameTextBoxes.Count - 1) * 40)
            removeButton.Size = New Size(20, 20)
            removeButton.Visible = display
            AddHandler removeButton.Click, Sub(sender As Object, e As EventArgs)
                                               Me.Controls.Remove(newModelNameTextBox)
                                               Me.Controls.Remove(removeButton)
                                               newModelNameTextBoxes.Remove(newModelNameTextBox)
                                               Me.Refresh()
                                           End Sub
            Me.Controls.Add(removeButton)
        End If
        Me.Refresh()
    End Sub


    Private Sub SaveConfigButton_Click(sender As Object, e As EventArgs)
        ' 获取新配置
        Dim newModelPlatform As String = newModelPlatformTextBox.Text
        Dim newApiUrl As String = newApiUrlTextBox.Text
        Dim newModels As New List(Of ConfigItemModel)()
        For Each textBox In newModelNameTextBoxes
            If textBox.Text <> "具体模型" AndAlso Not String.IsNullOrWhiteSpace(textBox.Text) Then
                newModels.Add(New ConfigItemModel() With {.modelName = textBox.Text, .selected = True})

            End If
        Next

        ' 如果newApiUrl不是以http://或https://开头，则报错异常提示
        If Not newApiUrl.StartsWith("http://") And Not newApiUrl.StartsWith("https://") Then
            MessageBox.Show("API URL 必须以 http:// 或 https:// 开头", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If



        ' 检查是否存在相同的 platform
        Dim existingItem As ConfigManager.ConfigItem = ConfigData.FirstOrDefault(Function(item) item.pltform = newModelPlatform)
        If existingItem IsNot Nothing Then
            ' 更新已有的 platform 数据
            existingItem.url = newApiUrl
            existingItem.model = newModels
            existingItem.selected = True
        Else
            ' 用户本地新增模型到 ComboBox
            Dim newItem As New ConfigManager.ConfigItem() With {
            .pltform = newModelPlatform,
            .url = newApiUrl,
            .model = newModels,
            .selected = True
        }
            ConfigData.Add(newItem)
            modelComboBox.Items.Add(newItem)
            modelComboBox.SelectedItem = newItem
        End If

        ' 保存到文件
        SaveConfig()

        'modelComboBox.Items.Add(newItem)
        'modelComboBox.SelectedItem = newItem

        modelNameComboBox.Items.Clear()
        For Each model In newModels
            modelNameComboBox.Items.Add(model)
        Next
        If modelNameComboBox.Items.Count > 0 Then
            modelNameComboBox.SelectedIndex = 0
        End If


        newModelPlatformTextBox.Text = "模型平台"
        newModelPlatformTextBox.ForeColor = Color.Gray
        newApiUrlTextBox.Text = "API URL"
        newApiUrlTextBox.ForeColor = Color.Gray
        For Each textBox In newModelNameTextBoxes
            textBox.Text = "具体模型"
            textBox.ForeColor = Color.Gray
        Next

        Me.Size = New Size(350, 300)
        newModelPlatformTextBox.Visible = False
        newApiUrlTextBox.Visible = False
        For Each textBox In newModelNameTextBoxes
            textBox.Visible = False
        Next
        addModelNameButton.Visible = False
        saveConfigButton.Visible = False
    End Sub



    Private Sub NewModelPlatformTextBox_Enter(sender As Object, e As EventArgs)
        If newModelPlatformTextBox.Text = "模型平台" Then
            newModelPlatformTextBox.Text = ""
            newModelPlatformTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub NewModelPlatformTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(newModelPlatformTextBox.Text) Then
            newModelPlatformTextBox.Text = "模型平台"
            newModelPlatformTextBox.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub NewApiUrlTextBox_Enter(sender As Object, e As EventArgs)
        If newApiUrlTextBox.Text = "API URL" Then
            newApiUrlTextBox.Text = ""
            newApiUrlTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub NewApiUrlTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(newApiUrlTextBox.Text) Then
            newApiUrlTextBox.Text = "API URL"
            newApiUrlTextBox.ForeColor = Color.Gray
        End If
    End Sub
    Private Sub NewModelNameTextBox_Enter(sender As Object, e As EventArgs)
        If CType(sender, TextBox).Text = "具体模型" Then
            CType(sender, TextBox).Text = ""
            CType(sender, TextBox).ForeColor = Color.Black
        End If
    End Sub

    Private Sub NewModelNameTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(CType(sender, TextBox).Text) Then
            CType(sender, TextBox).Text = "具体模型"
            CType(sender, TextBox).ForeColor = Color.Gray
        End If
    End Sub
End Class

