Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' 自动补全设置保存后触发的事件参数
''' </summary>
Public Class AutocompleteSettingsSavedEventArgs
    Inherits EventArgs

    Public Property EnableAutocomplete As Boolean
    Public Property AutocompleteDelayMs As Integer
    Public Property AutocompleteShortcut As String

    Public Sub New(enableAutocomplete As Boolean, delayMs As Integer, shortcut As String)
        Me.EnableAutocomplete = enableAutocomplete
        Me.AutocompleteDelayMs = delayMs
        Me.AutocompleteShortcut = shortcut
    End Sub
End Class

''' <summary>
''' 自动补全设置对话框
''' </summary>
Public Class AutocompleteSettingsForm
    Inherits Form

    ''' <summary>
    ''' 设置保存后触发的事件，各宿主应用可订阅此事件来同步状态
    ''' </summary>
    Public Shared Event SettingsSaved As EventHandler(Of AutocompleteSettingsSavedEventArgs)

    Private chkEnable As CheckBox
    Private lblShortcut As Label
    Private cmbShortcut As ComboBox
    Private lblDelay As Label
    Private numDelay As NumericUpDown
    Private lblDelayUnit As Label
    Private btnSave As Button
    Private btnCancel As Button
    Private lblDescription As Label

    ' 快捷键选项
    Private ReadOnly ShortcutOptions As String() = {
        "Ctrl+Enter",
        "Alt+/",
        "右箭头 →",
        "Ctrl+."
    }

    Public Sub New()
        InitializeComponent()
        LoadCurrentSettings()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "自动补全设置"
        Me.Size = New Size(400, 280)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        Dim yPos As Integer = 20

        ' 启用自动补全复选框
        chkEnable = New CheckBox()
        chkEnable.Text = "启用 AI 自动补全"
        chkEnable.Location = New Point(20, yPos)
        chkEnable.Size = New Size(340, 24)
        chkEnable.Font = New Font(chkEnable.Font.FontFamily, 10)
        AddHandler chkEnable.CheckedChanged, AddressOf ChkEnable_CheckedChanged
        Me.Controls.Add(chkEnable)

        yPos += 35

        ' 描述标签
        lblDescription = New Label()
        lblDescription.Text = "输入文本后，AI会自动显示灰色补全建议。按下快捷键即可接受补全。"
        lblDescription.Location = New Point(20, yPos)
        lblDescription.Size = New Size(340, 36)
        lblDescription.ForeColor = Color.Gray
        Me.Controls.Add(lblDescription)

        yPos += 45

        ' 快捷键标签
        lblShortcut = New Label()
        lblShortcut.Text = "接受补全快捷键:"
        lblShortcut.Location = New Point(20, yPos + 3)
        lblShortcut.Size = New Size(120, 20)
        Me.Controls.Add(lblShortcut)

        ' 快捷键下拉框
        cmbShortcut = New ComboBox()
        cmbShortcut.Location = New Point(150, yPos)
        cmbShortcut.Size = New Size(200, 24)
        cmbShortcut.DropDownStyle = ComboBoxStyle.DropDownList
        cmbShortcut.Items.AddRange(ShortcutOptions)
        Me.Controls.Add(cmbShortcut)

        yPos += 40

        ' 延迟标签
        lblDelay = New Label()
        lblDelay.Text = "触发延迟:"
        lblDelay.Location = New Point(20, yPos + 3)
        lblDelay.Size = New Size(120, 20)
        Me.Controls.Add(lblDelay)

        ' 延迟数字输入框
        numDelay = New NumericUpDown()
        numDelay.Location = New Point(150, yPos)
        numDelay.Size = New Size(100, 24)
        numDelay.Minimum = 100
        numDelay.Maximum = 5000
        numDelay.Increment = 100
        numDelay.Value = 800
        Me.Controls.Add(numDelay)

        ' 延迟单位标签
        lblDelayUnit = New Label()
        lblDelayUnit.Text = "毫秒 (输入后等待多久触发补全)"
        lblDelayUnit.Location = New Point(260, yPos + 3)
        lblDelayUnit.Size = New Size(200, 20)
        lblDelayUnit.ForeColor = Color.Gray
        Me.Controls.Add(lblDelayUnit)

        yPos += 50

        ' 保存按钮
        btnSave = New Button()
        btnSave.Text = "保存"
        btnSave.Location = New Point(180, yPos)
        btnSave.Size = New Size(80, 30)
        btnSave.DialogResult = DialogResult.OK
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        Me.Controls.Add(btnSave)

        ' 取消按钮
        btnCancel = New Button()
        btnCancel.Text = "取消"
        btnCancel.Location = New Point(270, yPos)
        btnCancel.Size = New Size(80, 30)
        btnCancel.DialogResult = DialogResult.Cancel
        Me.Controls.Add(btnCancel)

        Me.AcceptButton = btnSave
        Me.CancelButton = btnCancel

        ' 初始状态
        UpdateControlsEnabled()
    End Sub

    ''' <summary>
    ''' 加载当前设置
    ''' </summary>
    Private Sub LoadCurrentSettings()
        chkEnable.Checked = ChatSettings.EnableAutocomplete
        numDelay.Value = Math.Max(numDelay.Minimum, Math.Min(numDelay.Maximum, ChatSettings.AutocompleteDelayMs))

        ' 设置快捷键下拉框选项
        Dim shortcutIndex = Array.IndexOf(ShortcutOptions, ChatSettings.AutocompleteShortcut)
        If shortcutIndex >= 0 Then
            cmbShortcut.SelectedIndex = shortcutIndex
        Else
            ' 默认选择 Ctrl+Enter
            cmbShortcut.SelectedIndex = 0
        End If

        UpdateControlsEnabled()
    End Sub

    ''' <summary>
    ''' 复选框状态变化时更新控件启用状态
    ''' </summary>
    Private Sub ChkEnable_CheckedChanged(sender As Object, e As EventArgs)
        UpdateControlsEnabled()
    End Sub

    ''' <summary>
    ''' 根据启用状态更新控件
    ''' </summary>
    Private Sub UpdateControlsEnabled()
        Dim enabled = chkEnable.Checked
        lblShortcut.Enabled = enabled
        cmbShortcut.Enabled = enabled
        lblDelay.Enabled = enabled
        numDelay.Enabled = enabled
        lblDelayUnit.Enabled = enabled
    End Sub

    ''' <summary>
    ''' 保存设置
    ''' </summary>
    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        Try
            ' 更新静态属性
            ChatSettings.EnableAutocomplete = chkEnable.Checked
            ChatSettings.AutocompleteDelayMs = CInt(numDelay.Value)
            ChatSettings.AutocompleteShortcut = If(cmbShortcut.SelectedItem IsNot Nothing,
                                                   cmbShortcut.SelectedItem.ToString(),
                                                   "Ctrl+.")

            ' 保存到文件（调用现有的保存方法）
            SaveAutocompleteSettings()

            ' 触发设置保存事件，通知各宿主应用同步状态
            RaiseEvent SettingsSaved(Me, New AutocompleteSettingsSavedEventArgs(
                ChatSettings.EnableAutocomplete,
                ChatSettings.AutocompleteDelayMs,
                ChatSettings.AutocompleteShortcut))

            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("保存设置失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 保存自动补全设置到文件
    ''' </summary>
    Private Sub SaveAutocompleteSettings()
        Try
            Dim settingsPath = GetSettingsFilePath()
            Dim settingsObj As Newtonsoft.Json.Linq.JObject

            ' 读取现有设置
            If IO.File.Exists(settingsPath) Then
                Dim jsonContent = IO.File.ReadAllText(settingsPath)
                settingsObj = Newtonsoft.Json.Linq.JObject.Parse(jsonContent)
            Else
                settingsObj = New Newtonsoft.Json.Linq.JObject()
            End If

            ' 更新自动补全设置
            settingsObj("enableAutocomplete") = ChatSettings.EnableAutocomplete
            settingsObj("autocompleteDelayMs") = ChatSettings.AutocompleteDelayMs
            settingsObj("autocompleteShortcut") = ChatSettings.AutocompleteShortcut

            ' 确保目录存在
            IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(settingsPath))

            ' 保存设置
            IO.File.WriteAllText(settingsPath, settingsObj.ToString(Newtonsoft.Json.Formatting.Indented))

        Catch ex As Exception
            Debug.WriteLine($"保存自动补全设置失败: {ex.Message}")
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 获取设置文件路径
    ''' </summary>
    Private Function GetSettingsFilePath() As String
        Dim fileName As String = "office_ai_chat_settings.json"
        Return IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            ConfigSettings.OfficeAiAppDataFolder,
            fileName)
    End Function

End Class
