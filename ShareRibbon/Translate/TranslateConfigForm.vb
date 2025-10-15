Imports System.Drawing
Imports System.Windows.Forms
Imports Newtonsoft.Json

Public Class TranslateConfigForm
    Inherits Form

    Private enableCheckBox As CheckBox
    Private sourceLangCombo As ComboBox
    Private targetLangCombo As ComboBox
    Private platformsPanel As FlowLayoutPanel
    Private configureButton As Button
    Private saveButton As Button
    Private cancelButton As Button

    Private settings As TranslateSettings

    Private supportedLanguages As String() = New String() {
        "auto", "en", "zh", "ja", "ko", "fr", "de", "es", "ru", "pt", "it", "vi", "th", "id", "ar"
    }

    Public Sub New()
        Me.Text = "翻译配置"
        Me.Size = New Size(520, 500)
        Me.StartPosition = FormStartPosition.CenterParent

        settings = TranslateSettings.Load()

        enableCheckBox = New CheckBox() With {
            .Text = "启用翻译功能",
            .Checked = settings.Enabled,
            .Location = New Point(12, 12),
            .AutoSize = True
        }
        Me.Controls.Add(enableCheckBox)

        Dim lblFrom As New Label() With {
            .Text = "原始语言:",
            .Location = New Point(12, 44),
            .AutoSize = True
        }
        Me.Controls.Add(lblFrom)

        sourceLangCombo = New ComboBox() With {
            .Location = New Point(80, 40),
            .Size = New Size(120, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        sourceLangCombo.Items.AddRange(supportedLanguages)
        sourceLangCombo.SelectedItem = If(String.IsNullOrEmpty(settings.SourceLanguage), "auto", settings.SourceLanguage)
        Me.Controls.Add(sourceLangCombo)

        Dim lblTo As New Label() With {
            .Text = "目标语言:",
            .Location = New Point(220, 44),
            .AutoSize = True
        }
        Me.Controls.Add(lblTo)

        targetLangCombo = New ComboBox() With {
            .Location = New Point(288, 40),
            .Size = New Size(120, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        targetLangCombo.Items.AddRange(supportedLanguages)
        targetLangCombo.SelectedItem = If(String.IsNullOrEmpty(settings.TargetLanguage), "zh", settings.TargetLanguage)
        Me.Controls.Add(targetLangCombo)

        platformsPanel = New FlowLayoutPanel() With {
            .Location = New Point(12, 80),
            .Size = New Size(480, 320),
            .AutoScroll = True,
            .FlowDirection = FlowDirection.TopDown,
            .WrapContents = False
        }
        Me.Controls.Add(platformsPanel)

        configureButton = New Button() With {
            .Text = "配置全局翻译设置",
            .Location = New Point(12, 410),
            .Size = New Size(160, 28)
        }
        AddHandler configureButton.Click, AddressOf ConfigureButton_Click
        Me.Controls.Add(configureButton)

        saveButton = New Button() With {
            .Text = "保存",
            .Location = New Point(320, 410),
            .Size = New Size(80, 28)
        }
        AddHandler saveButton.Click, AddressOf SaveButton_Click
        Me.Controls.Add(saveButton)

        cancelButton = New Button() With {
            .Text = "取消",
            .Location = New Point(410, 410),
            .Size = New Size(80, 28)
        }
        AddHandler cancelButton.Click, AddressOf CancelButton_Click
        Me.Controls.Add(cancelButton)

        LoadPlatforms()
    End Sub

    Private Sub LoadPlatforms()
        platformsPanel.Controls.Clear()

        ' 列出已经验证的 validated = true 的平台
        For Each cfg In ConfigManager.ConfigData
            If cfg.validated Then
                Dim p As New Panel() With {
                    .Width = platformsPanel.ClientSize.Width - 25,
                    .Height = 56,
                    .BorderStyle = BorderStyle.None
                }

                Dim lbl As New Label() With {
                    .Text = cfg.pltform,
                    .Location = New Point(6, 6),
                    .AutoSize = True
                }
                p.Controls.Add(lbl)

                ' 模型下拉
                Dim modelCombo As New ComboBox() With {
                    .Location = New Point(6, 26),
                    .Size = New Size(320, 24),
                    .DropDownStyle = ComboBoxStyle.DropDownList,
                    .Tag = cfg ' 保存引用
                }
                For Each m In cfg.model
                    modelCombo.Items.Add(m.modelName)
                Next
                ' 选择已标注的 selected 模型
                For i As Integer = 0 To cfg.model.Count - 1
                    If cfg.model(i).translateSelected Then
                        modelCombo.SelectedIndex = i
                        Exit For
                    End If
                Next
                If modelCombo.Items.Count > 0 AndAlso modelCombo.SelectedIndex < 0 Then
                    modelCombo.SelectedIndex = 0
                End If
                p.Controls.Add(modelCombo)

                ' 单选项：设为翻译平台（全局仅允许一个）
                Dim radio As New RadioButton() With {
                    .Text = "设为翻译平台",
                    .Location = New Point(340, 28),
                    .AutoSize = True,
                    .Tag = cfg
                }
                radio.Checked = cfg.translateSelected
                AddHandler radio.CheckedChanged, AddressOf TranslateRadio_CheckedChanged
                p.Controls.Add(radio)

                platformsPanel.Controls.Add(p)
            End If
        Next
    End Sub

    Private Sub TranslateRadio_CheckedChanged(sender As Object, e As EventArgs)
        Dim rb As RadioButton = CType(sender, RadioButton)
        If rb.Checked Then
            ' 取消其他所有 radio / cfg.translateSelected = False
            For Each cfg In ConfigManager.ConfigData
                cfg.translateSelected = False
            Next
            Dim cfgSel As ConfigManager.ConfigItem = CType(rb.Tag, ConfigManager.ConfigItem)
            cfgSel.translateSelected = True
            ' 同步界面上的 radio
            For Each ctrl In platformsPanel.Controls
                For Each c In CType(ctrl, Panel).Controls
                    Dim r = TryCast(c, RadioButton)
                    If r IsNot Nothing And r IsNot rb Then
                        r.Checked = False
                    End If
                Next
            Next
        End If
    End Sub

    Private Sub ConfigureButton_Click(sender As Object, e As EventArgs)
        Dim dlg As New TranslateGlobalSettingsForm()
        dlg.Settings = settings
        If dlg.ShowDialog() = DialogResult.OK Then
            settings = dlg.Settings
            settings.Save()
        End If
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs)
        settings.Enabled = enableCheckBox.Checked
        settings.SourceLanguage = CStr(sourceLangCombo.SelectedItem)
        settings.TargetLanguage = CStr(targetLangCombo.SelectedItem)
        settings.Save()

        ' 只允许一个平台 translateSelected = True
        Dim selectedPlatform As ConfigManager.ConfigItem = Nothing
        For Each ctrl In platformsPanel.Controls
            For Each c In CType(ctrl, Panel).Controls
                Dim r = TryCast(c, RadioButton)
                If r IsNot Nothing Then
                    Dim cfg = CType(r.Tag, ConfigManager.ConfigItem)
                    If r.Checked Then
                        cfg.translateSelected = True
                        selectedPlatform = cfg
                    Else
                        cfg.translateSelected = False
                    End If
                End If
            Next
        Next

        ' 只对选中的平台设置模型的 translateSelected
        For Each ctrl In platformsPanel.Controls
            Dim panel = CType(ctrl, Panel)
            Dim cfg As ConfigManager.ConfigItem = Nothing
            Dim modelCombo As ComboBox = Nothing

            For Each c In panel.Controls
                If TypeOf c Is ComboBox Then
                    modelCombo = CType(c, ComboBox)
                    cfg = CType(modelCombo.Tag, ConfigManager.ConfigItem)
                End If
            Next

            If cfg IsNot Nothing AndAlso modelCombo IsNot Nothing Then
                ' 只处理选中的平台
                If selectedPlatform Is cfg Then
                    For Each m In cfg.model
                        m.translateSelected = False
                    Next
                    If modelCombo.SelectedIndex >= 0 Then
                        Dim selectedModelName = CStr(modelCombo.SelectedItem)
                        Dim selectedModel = cfg.model.FirstOrDefault(Function(m) m.modelName = selectedModelName)
                        If selectedModel IsNot Nothing Then
                            selectedModel.translateSelected = True
                        End If
                    End If
                Else
                    ' 其他平台所有模型都为False
                    For Each m In cfg.model
                        m.translateSelected = False
                    Next
                End If
            End If
        Next

        ConfigManager.SaveConfig()
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class