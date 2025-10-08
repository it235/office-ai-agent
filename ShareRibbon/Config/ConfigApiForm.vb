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
    Private getApiKeyButton As Button


    Public Property platform As String
    Public Property apiUrl As String
    Public Property apiKey As String
    Public Property modelName As String


    Public Sub New()
        ' 初始化表单
        Me.Text = "配置大模型API"
        Me.Size = New Size(450, 350)
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
        editConfigButton.Size = New Size(80, modelComboBox.Height + 2)
        AddHandler editConfigButton.Click, AddressOf EditConfigButton_Click
        Me.Controls.Add(editConfigButton)

        ' 初始化获取ApiKey按钮
        getApiKeyButton = New Button()
        getApiKeyButton.Text = "获取ApiKey"
        getApiKeyButton.Font = New Font(getApiKeyButton.Font.FontFamily, 8) ' 设置字体大小
        getApiKeyButton.Location = New Point(280, 90) ' 位置
        getApiKeyButton.Size = New Size(80, modelComboBox.Height + 2) ' 按钮大小
        'getApiKeyButton.ForeColor = Color.Blue ' 使用蓝色字体以表示这是一个链接
        AddHandler getApiKeyButton.Click, AddressOf GetApiKeyButton_Click
        Me.Controls.Add(getApiKeyButton)

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
        Me.Size = New Size(450, 500)
        newModelPlatformTextBox.Visible = True
        newApiUrlTextBox.Visible = True
        For Each textBox In newModelNameTextBoxes
            textBox.Visible = True
        Next
        addModelNameButton.Visible = True
        saveConfigButton.Visible = True
    End Sub

    ' 处理获取ApiKey按钮点击事件
    Private Sub GetApiKeyButton_Click(sender As Object, e As EventArgs)
        ' 指定URL
        Dim urll As String = "https://cloud.siliconflow.cn/i/PGhr3knx"
        Try
            ' 尝试使用Edge浏览器打开URL
            Process.Start("microsoft-edge:" & urll)
        Catch ex As Exception
            ' 如果无法使用Edge，则使用默认浏览器
            Try
                Process.Start(urll)
            Catch ex2 As Exception
                MessageBox.Show("无法打开浏览器。请手动访问: " & urll, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Try
    End Sub

    ' 切换大模型后的确认按钮
    Private Async Sub ConfirmButton_Click(sender As Object, e As EventArgs)
        ' 获取选中的模型和API Key
        Dim selectedPlatform As ConfigManager.ConfigItem = CType(modelComboBox.SelectedItem, ConfigManager.ConfigItem)
        Dim apiUrl As String = selectedPlatform.url
        Dim selectedModelName As String = If(modelNameComboBox.SelectedItem IsNot Nothing, modelNameComboBox.SelectedItem.ToString(), modelNameComboBox.Text)
        Dim inputApiKey As String = apiKeyTextBox.Text

        ' 检查API Key是否有效
        If inputApiKey = "输入 API Key" OrElse String.IsNullOrWhiteSpace(inputApiKey) Then
            MessageBox.Show("请输入有效的API Key", "验证失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 获取当前选中的模型对象
        Dim selectedModel As ConfigManager.ConfigItemModel = Nothing
        For Each model In selectedPlatform.model
            If model.modelName = selectedModelName Then
                selectedModel = model
                Exit For
            End If
        Next

        ' 判断是否需要验证：
        ' 1. 如果之前已验证过且API Key未变更，则无需再次验证
        ' 2. 如果之前未验证过或API Key已变更，则需要验证
        Dim needValidation As Boolean = True


        ' 检查两层验证状态
        If selectedPlatform.validated AndAlso selectedPlatform.key = inputApiKey AndAlso
           selectedModel IsNot Nothing AndAlso selectedModel.mcpValidated Then
            needValidation = False
        End If

        ' 如果不需要验证，直接保存并退出
        If Not needValidation Then
            ' 重置选择后的selected属性
            For Each config In ConfigData
                config.selected = False
                If selectedPlatform.pltform = config.pltform Then
                    config.selected = True
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
            ConfigSettings.mcpable = selectedModel.mcpable

            ' 关闭对话框
            Me.DialogResult = DialogResult.OK
            Me.Close()
            Return
        End If

        ' 需要验证，显示加载提示
        Cursor = Cursors.WaitCursor
        confirmButton.Enabled = False
        confirmButton.Text = "验证中..."

        GlobalStatusStripAll.ShowWarning("推理模型比普通模型会更加慢一些，请耐心等待")
        Try
            ' 首先使用简单的请求体进行快速验证
            Dim simpleRequestBody As String = $"{{""model"": ""{selectedModelName}"", ""stream"": true ,""messages"": [{{""role"": ""user"", ""content"": ""hi""}}]}}"
            Dim response As String = Await SendHttpRequestForValidation(apiUrl, inputApiKey, simpleRequestBody)

            ' 检查响应是否有效
            Dim validationSuccess As Boolean = Not String.IsNullOrEmpty(response)

            If validationSuccess Then

                Dim mcpSupported As Boolean = False


                ' 重置选择后的selected属性和key，设置validated为true
                For Each config In ConfigData
                    config.selected = False
                    If selectedPlatform.pltform = config.pltform Then
                        config.selected = True
                        config.key = inputApiKey
                        config.validated = True ' 标记为已验证
                        For Each item_m In config.model
                            item_m.selected = False
                            If item_m.modelName = selectedModelName Then
                                item_m.mcpable = mcpSupported
                                item_m.mcpValidated = False
                                item_m.selected = True
                            End If
                        Next
                    End If
                Next

                ' 保存到文件
                SaveConfig()

                If mcpSupported Then
                    Debug.WriteLine($"检测到 {selectedModelName} 模型支持MCP工具功能！", "MCP功能支持", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                ' 刷新内存中的api配置
                ConfigSettings.ApiUrl = apiUrl
                ConfigSettings.ApiKey = inputApiKey
                ConfigSettings.platform = selectedPlatform.pltform
                ConfigSettings.ModelName = selectedModelName
                ConfigSettings.mcpable = mcpSupported


                ' 检测MCP功能
                ' 验证成功后，异步检查function tools支持
                ' 注意：这里我们不等待结果，让它在后台运行
                Task.Run(Async Function()
                             Try

                                 Dim mcpSupportedTemp As Boolean = Await CheckFunctionToolsSupport(apiUrl, inputApiKey, selectedModelName)

                                 ' 更新配置中的MCP支持状态
                                 For Each config In ConfigData
                                     If config.pltform = selectedPlatform.pltform Then
                                         For Each item_m In config.model
                                             If item_m.modelName = selectedModelName Then
                                                 item_m.mcpable = mcpSupportedTemp
                                                 item_m.mcpValidated = True
                                                 ConfigSettings.mcpable = mcpSupportedTemp
                                                 Exit For
                                             End If
                                         Next
                                         Exit For
                                     End If
                                 Next

                                 ' 保存更新后的配置
                                 SaveConfig()

                                 If mcpSupportedTemp Then
                                     Debug.WriteLine($"检测到 {selectedModelName} 模型支持MCP工具功能！")
                                 End If
                             Catch ex As Exception
                                 Debug.WriteLine($"后台检查MCP支持时出错: {ex.Message}")
                             End Try
                         End Function)

                ' 关闭对话框
                Me.DialogResult = DialogResult.OK
                Me.Close()

            Else
                ' 验证失败，提示用户修改
                MessageBox.Show("API验证失败。请检查API URL、模型名称和API Key是否正确。", "验证失败",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)

                ' 标记为未验证
                selectedPlatform.validated = False
                If selectedModel IsNot Nothing Then
                    selectedModel.mcpValidated = False
                End If
            End If
        Catch ex As Exception
            ' 处理异常
            MessageBox.Show($"验证过程中出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' 标记为未验证
            selectedPlatform.validated = False
            If selectedModel IsNot Nothing Then
                selectedModel.mcpValidated = False
            End If
        Finally
            ' 恢复按钮状态
            confirmButton.Enabled = True
            confirmButton.Text = "确认"
            Cursor = Cursors.Default
        End Try
    End Sub

    ' 首先，添加一个异步方法来检查function tools支持
    Private Async Function CheckFunctionToolsSupport(apiUrl As String, apiKey As String, modelName As String) As Task(Of Boolean)
        Try
            ' 构建一个带tools定义的请求体
            Dim functionToolRequestBody As String = "{" &
            $"""model"": ""{modelName}"", ""stream"": true," &
            $"""messages"": [{{""role"": ""user"", ""content"": ""请计算5+7的结果，并通过工具函数返回""}}]," &
            $"""tools"": [" &
                "{" &
                    """type"": ""function""," &
                    """function"": {" &
                        """name"": ""calculator""," &
                        """description"": ""计算数学表达式的结果""," &
                        """parameters"": {" &
                            """type"": ""object""," &
                            """properties"": {" &
                                """result"": {" &
                                    """type"": ""number""," &
                                    """description"": ""计算结果""" &
                                "}" &
                            "}," &
                            """required"": [""result""]" &
                        "}" &
                    "}" &
                "}" &
            "]," &
            """tool_choice"": ""auto""" &
        "}"
            Dim toolResponse As String = Await SendHttpRequestForValidation(apiUrl, apiKey, functionToolRequestBody, True)

            Return Not String.IsNullOrEmpty(toolResponse)

        Catch ex As Exception
            Debug.WriteLine($"检查function tools支持时出错: {ex.Message}")
            Return False
        End Try
    End Function

    ' 用于验证的API请求方法
    Private Async Function SendHttpRequestForValidation(apiUrl As String, apiKey As String, requestBody As String, Optional checkFunctionTool As Boolean = False) As Task(Of String)
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Using client As New Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(60)
                Dim request As New Net.Http.HttpRequestMessage(Net.Http.HttpMethod.Post, apiUrl)
                request.Headers.Authorization = New Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey)
                request.Content = New Net.Http.StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")

                Using response As Net.Http.HttpResponseMessage = Await client.SendAsync(request, Net.Http.HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()
                    Debug.WriteLine($"[HTTP] 校验API响应状态码: {response.StatusCode}")
                    If response.StatusCode <> Net.HttpStatusCode.OK Then
                        Return String.Empty
                    End If

                    If Not checkFunctionTool Then
                        Return "OK"
                    End If

                    Using responseStream As IO.Stream = Await response.Content.ReadAsStreamAsync()
                        Using reader As New IO.StreamReader(responseStream, System.Text.Encoding.UTF8)
                            Dim buffer(40960) As Char
                            Dim readCount As Integer
                            Dim chunkBuilder As New StringBuilder()
                            Do
                                readCount = Await reader.ReadAsync(buffer, 0, buffer.Length)
                                If readCount = 0 Then Exit Do
                                Dim chunkT As String = New String(buffer, 0, readCount)
                                chunkT = chunkT.Replace("data:", "")
                                chunkBuilder.Append(chunkT)
                                If chunkBuilder.ToString().TrimEnd({ControlChars.Cr, ControlChars.Lf, " "c}).EndsWith("}") Then
                                    Dim chunk As String = chunkBuilder.ToString()
                                    If chunk.Trim() = "" Then
                                        Continue Do
                                    End If

                                    ' 按行分割处理
                                    Dim lines = chunk.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                    For Each line In lines
                                        If line = "[DONE]" OrElse line.Trim() = "" Then Continue For
                                        If Not line.TrimStart().StartsWith("{") Then Continue For
                                        Try
                                            Dim jsonObj = Newtonsoft.Json.Linq.JObject.Parse(line)
                                            Dim delta = jsonObj("choices")?(0)?("delta")
                                            If delta IsNot Nothing Then
                                                ' 推理模型：reasoning_content，普通模型：content
                                                If Not String.IsNullOrEmpty(delta("reasoning_content")?.ToString()) OrElse
                                               Not String.IsNullOrEmpty(delta("content")?.ToString()) Then
                                                    Return line ' API验证成功
                                                End If
                                                If checkFunctionTool Then
                                                    ' function tool相关字段
                                                    If delta("tool_calls") IsNot Nothing OrElse
                                                   delta("function_call") IsNot Nothing OrElse
                                                   delta("tools") IsNot Nothing OrElse
                                                   (jsonObj("capabilities") IsNot Nothing AndAlso jsonObj("capabilities")("tools") IsNot Nothing) Then
                                                        Return line ' function tool支持
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Exception
                                            ' 忽略解析错误
                                        End Try
                                    Next
                                    chunkBuilder.Clear()
                                End If

                            Loop
                        End Using
                    End Using
                End Using
            End Using
            Return String.Empty
        Catch ex As Exception
            Debug.WriteLine($"API验证请求失败: {ex.Message}")
            Return String.Empty
        End Try
    End Function

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
        Me.Size = New Size(450, 500)
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

        Me.Size = New Size(450, 300)
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

