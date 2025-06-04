Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports Newtonsoft.Json
Imports AiHelper.ConfigManager

' ��ģ����ʾ������
Public Class ConfigPromptForm
    Inherits Form
    Public Shared Property ConfigPromptData As List(Of PromptConfigItem)

    ' Ĭ�������ļ��ڵ�ǰ�û����ҵ��ĵ���
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
    Public Property VBA_Q As String = "����һ�������Excel����ר�ң��ó�дVBA���롣�������������룬������ϸ˼���ͼ�飬û�����ݻ��ʽ���Ե��п��������ô���ͬʱ�Ҳ���Ҫ�κ�ͨƪ���۵ķϻ����������ܿ���һ�㡣"

    Public Property EXCEL_TAB_Q As String = "����һ�������Excel����ר�ң��ó�ͨ��VBA��������������ɸ���ͼ�����磺��ͼ������ͼ����״ͼ��������ҵ������������Ҫ��VBA���롣ͬʱ�Ҳ���Ҫ�κ�ͨƪ���۵ķϻ����������ܿ���һ�㡣"

    Public Sub LoadConfig()
        ' ��ʼ����������
        ConfigPromptData = New List(Of PromptConfigItem)()

        Dim vbap = New PromptConfigItem() With {
                .name = "VBAר�����",
                .content = VBA_Q,
                .selected = True
            }

        Dim excelTabP = New PromptConfigItem() With {
                .name = "Excel���ר�����",
                .content = EXCEL_TAB_Q,
                .selected = False
            }
        ' ���Ĭ������
        If Not File.Exists(configFilePath) Then
            ConfigPromptData.Add(vbap)
            ConfigPromptData.Add(excelTabP)
        Else
            ' �����Զ�������
            Dim json As String = File.ReadAllText(configFilePath)
            ConfigPromptData = JsonConvert.DeserializeObject(Of List(Of PromptConfigItem))(json)
        End If

        ' ��ʼ�����ã������ݳ�ʼ���� ConfigSettings������ȫ�ֵ���
        For Each item In ConfigPromptData
            If item.selected Then
                ConfigSettings.propmtName = item.name
                ConfigSettings.propmtContent = item.content
            End If
        Next
    End Sub


    Public Sub New()
        LoadConfig()

        ' ��ʼ����
        Me.Text = "���������ģ����ʾ��"
        Me.Size = New Size(480, 550)
        Me.StartPosition = FormStartPosition.CenterScreen ' ���ñ�������ʾ

        descriptionLabel1 = New Label()
        descriptionLabel1.Text = "��ʾ���൱�ڸ�AI�趨��Ӧ����ݣ����������ҵ�����������⣬�ش�������רҵ�����磺����һ��Excel VBAר�ң������������ⶼ��Excel�Լ�VBA���"
        descriptionLabel1.Dock = DockStyle.Top
        descriptionLabel1.Height = 40
        descriptionLabel1.Margin = New Padding(10, 10, 10, 10)
        descriptionLabel1.TextAlign = ContentAlignment.MiddleLeft
        Me.Controls.Add(descriptionLabel1)

        ' ��ʼ��ģ��ѡ�� ComboBox
        currentPromptComboBox = New ComboBox()
        currentPromptComboBox.DisplayMember = "name"
        currentPromptComboBox.ValueMember = "value"
        currentPromptComboBox.Location = New Point(10, 50)
        currentPromptComboBox.Size = New Size(260, 30)
        AddHandler currentPromptComboBox.SelectedIndexChanged, AddressOf propmtCombBox_SelectedIndexChanged
        Me.Controls.Add(currentPromptComboBox)

        ' ��ʼ���༭���ð�ť
        editConfigButton = New Button()
        editConfigButton.Text = "�޸�"
        editConfigButton.Font = New Font(editConfigButton.Font.FontFamily, 8) ' ���������С
        editConfigButton.Location = New Point(280, 50)
        editConfigButton.Size = New Size(40, currentPromptComboBox.Height + 2)
        AddHandler editConfigButton.Click, AddressOf EditConfigButton_Click

        Me.Controls.Add(editConfigButton)


        ' ��������֮ǰѡ�����ʾ�����ƺ���ʾ������
        Dim propmtNameForDB As String
        Dim propmtContentForDB As String

        For Each config In ConfigPromptData
            If config.selected Then
                propmtNameForDB = config.name
                propmtContentForDB = config.content
            End If
        Next

        ' ��ʾ������Ԥ����
        promptContentBox = New TextBox()
        promptContentBox.Multiline = True
        promptContentBox.ScrollBars = ScrollBars.Vertical
        promptContentBox.Text = propmtContentForDB
        promptContentBox.ForeColor = Color.Gray
        promptContentBox.Location = New Point(10, 80)
        promptContentBox.Size = New Size(360, 120)
        promptContentBox.ReadOnly = True
        Me.Controls.Add(promptContentBox)


        ' ��ʼ��ȷ�ϰ�ť
        confirmButton = New Button()
        confirmButton.Text = "ʹ�ø���ʾ��"
        confirmButton.Location = New Point(50, 210)
        confirmButton.Size = New Size(100, 30)
        AddHandler confirmButton.Click, AddressOf ConfirmButton_Click
        Me.Controls.Add(confirmButton)

        ' ��ʼ��������ð�ť
        addConfigButton = New Button()
        addConfigButton.Text = "�������ʾ��"
        addConfigButton.Location = New Point(170, 210)
        addConfigButton.Size = New Size(100, 30)
        AddHandler addConfigButton.Click, AddressOf AddConfigButton_Click
        Me.Controls.Add(addConfigButton)

        ' ��ʼ�������ÿؼ�
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
        'newPromptContent.Text = If(String.IsNullOrEmpty(propmtContentForDB), "������ʾ������", propmtContentForDB)
        newPromptContent.ForeColor = If(String.IsNullOrEmpty(propmtContentForDB), Color.Gray, Color.Black)
        newPromptContent.Location = New Point(10, 290)
        newPromptContent.Size = New Size(360, 120)
        newPromptContent.Visible = False
        AddHandler newPromptContent.Enter, AddressOf ApiKeyTextBox_Enter ' ��� Enter �¼��������
        AddHandler newPromptContent.Leave, AddressOf ApiKeyTextBox_Leave ' ��� Leave �¼��������
        Me.Controls.Add(newPromptContent)

        saveConfigButton = New Button()
        saveConfigButton.Text = "����"
        saveConfigButton.Location = New Point(100, 420)
        saveConfigButton.Size = New Size(100, 30)
        saveConfigButton.Visible = False
        AddHandler saveConfigButton.Click, AddressOf SaveConfigButton_Click
        Me.Controls.Add(saveConfigButton)

        ' �������õ���ѡ��
        For Each configItem In ConfigPromptData
            currentPromptComboBox.Items.Add(configItem)
        Next

        ' ����֮ǰѡ���ģ��
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
        ' ��ȡѡ�е�ģ����ʾ��
        Dim selectedPlatform As PromptConfigItem = CType(currentPromptComboBox.SelectedItem, PromptConfigItem)

        ' ��ѡ�е����ݴ��뵽�����ÿؼ���
        newPromptComboBox.Text = selectedPlatform.name
        newPromptComboBox.ForeColor = Color.Black

        newPromptContent.Text = selectedPlatform.content
        newPromptContent.ForeColor = Color.Black

        ' ��ʾ�����ÿؼ�
        Me.Size = New Size(480, 550)
        newPromptComboBox.Visible = True
        newPromptContent.Visible = True
        saveConfigButton.Visible = True
    End Sub


    ' �л���ʾ�ʺ��ȷ�ϰ�ť
    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs)

        ' ��ȡѡ�е���ʾ�����ƺ�����������
        Dim selectedPlatform As PromptConfigItem = CType(currentPromptComboBox.SelectedItem, PromptConfigItem)
        Dim name As String = selectedPlatform.name
        Dim content As String = selectedPlatform.content

        ' ����ѡ����selected���Ժ�key
        For Each config In ConfigPromptData
            config.selected = False
            If selectedPlatform.name = config.name Then
                config.selected = True
                config.name = name
                config.content = content
            End If
        Next

        ' ���浽�ļ�
        SaveConfig()

        ' ˢ���ڴ��е�api����
        ConfigSettings.propmtName = name
        ConfigSettings.propmtContent = content

        ' �رնԻ���
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub


    Public Shared Sub SaveConfig()
        Dim json As String = JsonConvert.SerializeObject(ConfigPromptData, Formatting.Indented)
        ' ���configFilePath��Ŀ¼�����ھʹ���
        Dim dir = Path.GetDirectoryName(configFilePath)
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
        '����ļ������ھʹ���
        If Not File.Exists(configFilePath) Then
            File.Create(configFilePath).Dispose()
        End If
        File.WriteAllText(configFilePath, json)
    End Sub


    Private Sub propmtCombBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' ����ѡ�е���ʾ��������ʾ��ͬ������
        Dim selectedModel As PromptConfigItem = CType(currentPromptComboBox.SelectedItem, PromptConfigItem)
        promptContentBox.Clear()
        promptContentBox.Text = selectedModel.content
        'promptContentBox.ForeColor = Color.Black
    End Sub

    Private Sub AddConfigButton_Click(sender As Object, e As EventArgs)
        ' ��ʾ�����ÿؼ�
        Me.Size = New Size(480, 550)
        newPromptComboBox.Visible = True
        newPromptContent.Visible = True
        saveConfigButton.Visible = True
    End Sub


    Private Sub SaveConfigButton_Click(sender As Object, e As EventArgs)
        ' ��ȡ������
        Dim name As String = newPromptComboBox.Text
        Dim content As String = newPromptContent.Text

        If String.IsNullOrWhiteSpace(name) Or name = NEW_NAME_C Then
            GlobalStatusStrip.ShowWarning("��������ʾ�����ƣ�")
            Return
        End If

        If String.IsNullOrWhiteSpace(content) Then
            GlobalStatusStrip.ShowWarning("��������ʾ�����ݣ�")
            Return
        End If

        ' ����Ƿ������ͬ�� propmtName
        Dim existingItem As PromptConfigItem = ConfigPromptData.FirstOrDefault(Function(item) item.name = name)
        If existingItem IsNot Nothing Then
            ' �������е� propmtName ����
            existingItem.name = name
            existingItem.content = content
            existingItem.selected = True
        Else
            ' �û���������ģ�͵� ComboBox
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

        ' ���浽�ļ�
        SaveConfig()

        ConfigSettings.propmtContent = content
        ConfigSettings.propmtName = name


        Me.Size = New Size(480, 550)
        newPromptComboBox.Visible = False
        newPromptContent.Visible = False
        saveConfigButton.Visible = False
    End Sub

    Private Property NEW_NAME_C As String = "ȡ�����������ƣ����磺Excel����ר��"

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

    Private Property NEW_CONTENT_C As String = "�����ģ����ʾ�����ݣ�Ϊ���趨һ����ݣ����磺����һ���ǳ�������Excel��ʦ���ó�����VBA����"
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


    ' ��ʾ�����ã�ÿ�ν���ʹ��1����
    Public Class PromptConfigItem
        Public Property name As String
        Public Property content As String
        Public Property selected As Boolean
        Public Overrides Function ToString() As String
            Return content
        End Function
    End Class
End Class

