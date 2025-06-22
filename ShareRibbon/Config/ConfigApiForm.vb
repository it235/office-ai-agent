Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Security.Policy
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports ShareRibbon.ConfigManager
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading.Tasks
Imports System.Diagnostics
Public Class ConfigApiForm
    Inherits Form

    Private modelComboBox As ComboBox
    ' �༭��ť
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
        ' ��ʼ����
        Me.Text = "���ô�ģ��API"
        Me.Size = New Size(450, 350)
        Me.StartPosition = FormStartPosition.CenterScreen ' ���ñ�������ʾ

        ' ��ʼ��ģ��ѡ�� ComboBox
        modelComboBox = New ComboBox()
        modelComboBox.DisplayMember = "pltform"
        modelComboBox.ValueMember = "url"
        modelComboBox.Location = New Point(10, 10)
        modelComboBox.Size = New Size(260, 30)
        AddHandler modelComboBox.SelectedIndexChanged, AddressOf ModelComboBox_SelectedIndexChanged
        Me.Controls.Add(modelComboBox)

        ' ��ʼ���༭���ð�ť
        editConfigButton = New Button()
        editConfigButton.Text = "�޸�"
        editConfigButton.Font = New Font(editConfigButton.Font.FontFamily, 8) ' ���������С
        editConfigButton.Location = New Point(280, 10)
        editConfigButton.Size = New Size(80, modelComboBox.Height + 2)
        AddHandler editConfigButton.Click, AddressOf EditConfigButton_Click
        Me.Controls.Add(editConfigButton)

        ' ��ʼ����ȡApiKey��ť
        getApiKeyButton = New Button()
        getApiKeyButton.Text = "��ȡApiKey"
        getApiKeyButton.Font = New Font(getApiKeyButton.Font.FontFamily, 8) ' ���������С
        getApiKeyButton.Location = New Point(280, 90) ' λ��
        getApiKeyButton.Size = New Size(80, modelComboBox.Height + 2) ' ��ť��С
        'getApiKeyButton.ForeColor = Color.Blue ' ʹ����ɫ�����Ա�ʾ����һ������
        AddHandler getApiKeyButton.Click, AddressOf GetApiKeyButton_Click
        Me.Controls.Add(getApiKeyButton)

        ' ��ʼ��ģ������ѡ�� ComboBox
        modelNameComboBox = New ComboBox()
        modelNameComboBox.Location = New Point(10, 50)
        modelNameComboBox.Size = New Size(260, 30)
        modelNameComboBox.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        modelNameComboBox.AutoCompleteSource = AutoCompleteSource.ListItems
        Me.Controls.Add(modelNameComboBox)

        ' ��������֮ǰѡ���ģ�ͺ� API Key
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

        ' ��ʼ�� API Key �����
        apiKeyTextBox = New TextBox()
        apiKeyTextBox.Text = If(String.IsNullOrEmpty(apiKeyForDB), "���� API Key", apiKeyForDB)
        apiKeyTextBox.ForeColor = If(String.IsNullOrEmpty(apiKeyForDB), Color.Gray, Color.Black)
        apiKeyTextBox.Location = New Point(10, 90)
        apiKeyTextBox.Size = New Size(260, 30)
        AddHandler apiKeyTextBox.Enter, AddressOf ApiKeyTextBox_Enter ' ��� Enter �¼��������
        AddHandler apiKeyTextBox.Leave, AddressOf ApiKeyTextBox_Leave ' ��� Leave �¼��������
        Me.Controls.Add(apiKeyTextBox)

        ' ��ʼ��ȷ�ϰ�ť
        confirmButton = New Button()
        confirmButton.Text = "ȷ��"
        confirmButton.Location = New Point(100, 130)
        confirmButton.Size = New Size(100, 30)
        AddHandler confirmButton.Click, AddressOf ConfirmButton_Click
        Me.Controls.Add(confirmButton)

        ' ��ʼ��������ð�ť
        addConfigButton = New Button()
        addConfigButton.Text = "���ģ������"
        addConfigButton.Location = New Point(100, 170)
        addConfigButton.Size = New Size(100, 30)
        AddHandler addConfigButton.Click, AddressOf AddConfigButton_Click
        Me.Controls.Add(addConfigButton)

        ' ��ʼ�������ÿؼ�
        newModelPlatformTextBox = New TextBox()
        newModelPlatformTextBox.Text = "ģ��ƽ̨"
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
        saveConfigButton.Text = "����"
        saveConfigButton.Location = New Point(100, 420)
        saveConfigButton.Size = New Size(100, 30)
        saveConfigButton.Visible = False
        AddHandler saveConfigButton.Click, AddressOf SaveConfigButton_Click
        Me.Controls.Add(saveConfigButton)

        ' �������õ���ѡ��
        For Each configItem In ConfigData
            modelComboBox.Items.Add(configItem)
        Next

        ' ����֮ǰѡ���ģ��
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

        ' ����֮ǰѡ���ģ������
        If Not String.IsNullOrEmpty(modelNameForDB) Then
            For i As Integer = 0 To modelNameComboBox.Items.Count - 1
                If modelNameComboBox.Items(i).ToString() = modelNameForDB Then
                    modelNameComboBox.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        ' ����֮ǰ�� API Key
        If Not String.IsNullOrEmpty(apiKeyForDB) Then
            apiKeyTextBox.Text = apiKeyForDB
            apiKeyTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ApiKeyTextBox_Enter(sender As Object, e As EventArgs)
        If apiKeyTextBox.Text = "���� API Key" Then
            apiKeyTextBox.Text = ""
            apiKeyTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub ApiKeyTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(apiKeyTextBox.Text) Then
            apiKeyTextBox.Text = "���� API Key"
            apiKeyTextBox.ForeColor = Color.Gray
        End If
    End Sub

    Private Sub EditConfigButton_Click(sender As Object, e As EventArgs)
        ' ��ȡѡ�е�ģ�ͺ� API Key
        Dim selectedPlatform As ConfigManager.ConfigItem = CType(modelComboBox.SelectedItem, ConfigManager.ConfigItem)
        Dim selectedModelName As String = If(modelNameComboBox.SelectedItem IsNot Nothing, modelNameComboBox.SelectedItem.ToString(), modelNameComboBox.Text)

        ' ��ѡ�е����ݴ��뵽�����ÿؼ���
        newModelPlatformTextBox.Text = selectedPlatform.pltform
        newModelPlatformTextBox.ForeColor = Color.Black
        newApiUrlTextBox.Text = selectedPlatform.url
        newApiUrlTextBox.ForeColor = Color.Black

        ' ��ղ�������� newModelNameTextBoxes
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
                newModelNameTextBox.BackColor = Color.LightBlue ' ���ѡ�е�ģ������
            End If
        Next

        ' ��ʾ�����ÿؼ�
        Me.Size = New Size(450, 500)
        newModelPlatformTextBox.Visible = True
        newApiUrlTextBox.Visible = True
        For Each textBox In newModelNameTextBoxes
            textBox.Visible = True
        Next
        addModelNameButton.Visible = True
        saveConfigButton.Visible = True
    End Sub

    ' �����ȡApiKey��ť����¼�
    Private Sub GetApiKeyButton_Click(sender As Object, e As EventArgs)
        ' ָ��URL
        Dim urll As String = "https://cloud.siliconflow.cn/i/PGhr3knx"
        Try
            ' ����ʹ��Edge�������URL
            Process.Start("microsoft-edge:" & urll)
        Catch ex As Exception
            ' ����޷�ʹ��Edge����ʹ��Ĭ�������
            Try
                Process.Start(urll)
            Catch ex2 As Exception
                MessageBox.Show("�޷�������������ֶ�����: " & urll, "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Try
    End Sub

    ' �л���ģ�ͺ��ȷ�ϰ�ť
    Private Async Sub ConfirmButton_Click(sender As Object, e As EventArgs)
        ' ��ȡѡ�е�ģ�ͺ�API Key
        Dim selectedPlatform As ConfigManager.ConfigItem = CType(modelComboBox.SelectedItem, ConfigManager.ConfigItem)
        Dim apiUrl As String = selectedPlatform.url
        Dim selectedModelName As String = If(modelNameComboBox.SelectedItem IsNot Nothing, modelNameComboBox.SelectedItem.ToString(), modelNameComboBox.Text)
        Dim inputApiKey As String = apiKeyTextBox.Text

        ' ���API Key�Ƿ���Ч
        If inputApiKey = "���� API Key" OrElse String.IsNullOrWhiteSpace(inputApiKey) Then
            MessageBox.Show("��������Ч��API Key", "��֤ʧ��", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' �ж��Ƿ���Ҫ��֤��
        ' 1. ���֮ǰ����֤����API Keyδ������������ٴ���֤
        ' 2. ���֮ǰδ��֤����API Key�ѱ��������Ҫ��֤
        Dim needValidation As Boolean = True

        ' ����Ƿ�����֤����API Keyδ���
        If selectedPlatform.validated AndAlso selectedPlatform.key = inputApiKey Then
            needValidation = False
        End If

        ' �������Ҫ��֤��ֱ�ӱ��沢�˳�
        If Not needValidation Then
            ' ����ѡ����selected����
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

            ' ���浽�ļ�
            SaveConfig()

            ' ˢ���ڴ��е�api����
            ConfigSettings.ApiUrl = apiUrl
            ConfigSettings.ApiKey = inputApiKey
            ConfigSettings.platform = selectedPlatform.pltform
            ConfigSettings.ModelName = selectedModelName

            ' �رնԻ���
            Me.DialogResult = DialogResult.OK
            Me.Close()
            Return
        End If

        ' ��Ҫ��֤����ʾ������ʾ
        Cursor = Cursors.WaitCursor
        confirmButton.Enabled = False
        confirmButton.Text = "��֤��..."

        Try
            ' ����һ���򵥵�������
            Dim requestBody As String = $"{{""model"": ""{selectedModelName}"", ""messages"": [{{""role"": ""user"", ""content"": ""hi""}}]}}"

            ' ����API��֤
            Dim response As String = Await SendHttpRequestForValidation(apiUrl, inputApiKey, requestBody)

            ' �����Ӧ�Ƿ���Ч
            Dim validationSuccess As Boolean = Not String.IsNullOrEmpty(response) AndAlso
                                         (response.Contains("content") OrElse response.Contains("message"))

            If validationSuccess Then
                ' ��֤�ɹ����������ò�����

                ' ����ѡ����selected���Ժ�key������validatedΪtrue
                For Each config In ConfigData
                    config.selected = False
                    If selectedPlatform.pltform = config.pltform Then
                        config.selected = True
                        config.key = inputApiKey
                        config.validated = True ' ���Ϊ����֤
                        For Each item_m In config.model
                            item_m.selected = False
                            If item_m.modelName = selectedModelName Then
                                item_m.selected = True
                            End If
                        Next
                    End If
                Next

                ' ���浽�ļ�
                SaveConfig()

                ' ˢ���ڴ��е�api����
                ConfigSettings.ApiUrl = apiUrl
                ConfigSettings.ApiKey = inputApiKey
                ConfigSettings.platform = selectedPlatform.pltform
                ConfigSettings.ModelName = selectedModelName

                ' �رնԻ���
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Else
                ' ��֤ʧ�ܣ���ʾ�û��޸�
                MessageBox.Show("API��֤ʧ�ܡ�����API URL��ģ�����ƺ�API Key�Ƿ���ȷ��", "��֤ʧ��",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning)

                ' ���Ϊδ��֤
                selectedPlatform.validated = False
            End If
        Catch ex As Exception
            ' �����쳣
            MessageBox.Show($"��֤�����г���: {ex.Message}", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' ���Ϊδ��֤
            selectedPlatform.validated = False
        Finally
            ' �ָ���ť״̬
            confirmButton.Enabled = True
            confirmButton.Text = "ȷ��"
            Cursor = Cursors.Default
        End Try
    End Sub

    ' ������֤��API���󷽷�
    Private Async Function SendHttpRequestForValidation(apiUrl As String, apiKey As String, requestBody As String) As Task(Of String)
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            Using client As New Net.Http.HttpClient()
                client.Timeout = TimeSpan.FromSeconds(15) ' �϶̵ĳ�ʱʱ�䣬ֻ������֤
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim content As New Net.Http.StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
                Dim response As Net.Http.HttpResponseMessage = Await client.PostAsync(apiUrl, content)

                ' ������������ش���״̬�룬������׳��쳣
                response.EnsureSuccessStatusCode()

                ' ��ȡ��������Ӧ����
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            ' �����ﴦ���쳣������ʾ��Ϣ����Ϊ���ǻ��ڵ��÷�������ʾ
            Debug.WriteLine($"API��֤����ʧ��: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Sub ModelComboBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' ����ѡ�е�ģ�͸���ģ������ѡ�� ComboBox
        modelNameComboBox.Items.Clear()
        Dim selectedModel As ConfigManager.ConfigItem = CType(modelComboBox.SelectedItem, ConfigManager.ConfigItem)
        For Each ModelNameT In selectedModel.model
            modelNameComboBox.Items.Add(ModelNameT)
        Next
        If modelNameComboBox.Items.Count > 0 Then
            modelNameComboBox.SelectedIndex = 0
        End If

        ' ���� API Key
        apiKeyTextBox.Text = selectedModel.key
        apiKeyTextBox.ForeColor = If(String.IsNullOrEmpty(selectedModel.key), Color.Gray, Color.Black)
    End Sub

    Private Sub AddConfigButton_Click(sender As Object, e As EventArgs)
        ' ��ʾ�����ÿؼ�
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
        newModelNameTextBox.Text = "����ģ��"
        newModelNameTextBox.ForeColor = Color.Gray
        newModelNameTextBox.Location = New Point(10, 290 + newModelNameTextBoxes.Count * 40)
        newModelNameTextBox.Size = New Size(260, 30)
        newModelNameTextBox.Visible = display
        AddHandler newModelNameTextBox.Enter, AddressOf NewModelNameTextBox_Enter
        AddHandler newModelNameTextBox.Leave, AddressOf NewModelNameTextBox_Leave
        Me.Controls.Add(newModelNameTextBox)
        newModelNameTextBoxes.Add(newModelNameTextBox)

        ' ֻ�еڶ��м�֮����в���Ӽ��Ű�ť
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
        ' ��ȡ������
        Dim newModelPlatform As String = newModelPlatformTextBox.Text
        Dim newApiUrl As String = newApiUrlTextBox.Text
        Dim newModels As New List(Of ConfigItemModel)()
        For Each textBox In newModelNameTextBoxes
            If textBox.Text <> "����ģ��" AndAlso Not String.IsNullOrWhiteSpace(textBox.Text) Then
                newModels.Add(New ConfigItemModel() With {.modelName = textBox.Text, .selected = True})

            End If
        Next

        ' ���newApiUrl������http://��https://��ͷ���򱨴��쳣��ʾ
        If Not newApiUrl.StartsWith("http://") And Not newApiUrl.StartsWith("https://") Then
            MessageBox.Show("API URL ������ http:// �� https:// ��ͷ", "����", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If



        ' ����Ƿ������ͬ�� platform
        Dim existingItem As ConfigManager.ConfigItem = ConfigData.FirstOrDefault(Function(item) item.pltform = newModelPlatform)
        If existingItem IsNot Nothing Then
            ' �������е� platform ����
            existingItem.url = newApiUrl
            existingItem.model = newModels
            existingItem.selected = True
        Else
            ' �û���������ģ�͵� ComboBox
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

        ' ���浽�ļ�
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


        newModelPlatformTextBox.Text = "ģ��ƽ̨"
        newModelPlatformTextBox.ForeColor = Color.Gray
        newApiUrlTextBox.Text = "API URL"
        newApiUrlTextBox.ForeColor = Color.Gray
        For Each textBox In newModelNameTextBoxes
            textBox.Text = "����ģ��"
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
        If newModelPlatformTextBox.Text = "ģ��ƽ̨" Then
            newModelPlatformTextBox.Text = ""
            newModelPlatformTextBox.ForeColor = Color.Black
        End If
    End Sub

    Private Sub NewModelPlatformTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(newModelPlatformTextBox.Text) Then
            newModelPlatformTextBox.Text = "ģ��ƽ̨"
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
        If CType(sender, TextBox).Text = "����ģ��" Then
            CType(sender, TextBox).Text = ""
            CType(sender, TextBox).ForeColor = Color.Black
        End If
    End Sub

    Private Sub NewModelNameTextBox_Leave(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(CType(sender, TextBox).Text) Then
            CType(sender, TextBox).Text = "����ģ��"
            CType(sender, TextBox).ForeColor = Color.Gray
        End If
    End Sub
End Class

