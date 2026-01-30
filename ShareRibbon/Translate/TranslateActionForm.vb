Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' 翻译操作对话框 - 整合所有常用翻译设置
''' </summary>
Public Class TranslateActionForm
    Inherits Form

    ' 翻译范围
    Private grpScope As GroupBox
    Private rbAll As RadioButton
    Private rbSelection As RadioButton

    ' 翻译平台和模型
    Private grpModel As GroupBox
    Private cbPlatform As ComboBox
    Private cbModel As ComboBox

    ' 语言和领域
    Private grpLanguage As GroupBox
    Private cbTargetLang As ComboBox
    Private cbDomain As ComboBox
    Private btnEditDomain As Button

    ' 输出方式
    Private grpOutput As GroupBox
    Private rbSidePanel As RadioButton
    Private rbReplace As RadioButton
    Private rbImmersive As RadioButton
    Private rbNewDoc As RadioButton

    ' 沉浸式翻译样式
    Private pnlImmersiveStyle As Panel
    Private chkPreserveFormat As CheckBox
    Private btnColor As Button
    Private chkItalic As CheckBox
    Private selectedColor As Color = Color.FromArgb(102, 102, 102)

    ' 按钮
    Private btnTranslate As Button
    Private btnCancel As Button

    ' 进度
    Private progressBar As ProgressBar
    Private lblProgress As Label

    Private _hasSelection As Boolean
    Private _appType As String

    ''' <summary>翻译范围：True=全部，False=选区</summary>
    Public Property TranslateAll As Boolean = True

    ''' <summary>输出模式</summary>
    Public Property OutputMode As TranslateOutputMode = TranslateOutputMode.Immersive

    ''' <summary>选择的领域</summary>
    Public Property SelectedDomain As String = "通用"

    ''' <summary>目标语言</summary>
    Public Property TargetLanguage As String = "zh"

    ''' <summary>源语言（固定为auto）</summary>
    Public Property SourceLanguage As String = "auto"

    Private targetLanguages As String() = {"en", "zh", "ja", "ko", "fr", "de", "es", "ru", "pt", "it"}
    Private languageNames As Dictionary(Of String, String) = New Dictionary(Of String, String) From {
        {"en", "英语"}, {"zh", "中文"}, {"ja", "日语"}, {"ko", "韩语"},
        {"fr", "法语"}, {"de", "德语"}, {"es", "西班牙语"}, {"ru", "俄语"},
        {"pt", "葡萄牙语"}, {"it", "意大利语"}
    }

    Public Sub New(hasSelection As Boolean, appType As String)
        _hasSelection = hasSelection
        _appType = appType

        Me.Text = "一键翻译"
        Me.Size = New Size(500, 560)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        InitializeControls()
        LoadSettings()
    End Sub

    Private Sub InitializeControls()
        Dim yPos = 10

        ' ========== 翻译平台和模型 ==========
        grpModel = New GroupBox() With {
            .Text = "翻译平台",
            .Location = New Point(10, yPos),
            .Size = New Size(465, 70)
        }
        Me.Controls.Add(grpModel)

        Dim lblPlatform As New Label() With {.Text = "平台:", .Location = New Point(15, 25), .AutoSize = True}
        grpModel.Controls.Add(lblPlatform)

        cbPlatform = New ComboBox() With {
            .Location = New Point(55, 22),
            .Size = New Size(150, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        AddHandler cbPlatform.SelectedIndexChanged, AddressOf PlatformChanged
        grpModel.Controls.Add(cbPlatform)

        Dim lblModel As New Label() With {.Text = "模型:", .Location = New Point(220, 25), .AutoSize = True}
        grpModel.Controls.Add(lblModel)

        cbModel = New ComboBox() With {
            .Location = New Point(260, 22),
            .Size = New Size(190, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        grpModel.Controls.Add(cbModel)

        yPos += 80

        ' ========== 语言和领域 ==========
        grpLanguage = New GroupBox() With {
            .Text = "翻译设置",
            .Location = New Point(10, yPos),
            .Size = New Size(465, 75)
        }
        Me.Controls.Add(grpLanguage)

        Dim lblTarget As New Label() With {.Text = "翻译为:", .Location = New Point(15, 28), .AutoSize = True}
        grpLanguage.Controls.Add(lblTarget)

        cbTargetLang = New ComboBox() With {
            .Location = New Point(70, 25),
            .Size = New Size(100, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        For Each lang In targetLanguages
            cbTargetLang.Items.Add(languageNames(lang))
        Next
        cbTargetLang.SelectedIndex = 1 ' 默认中文
        grpLanguage.Controls.Add(cbTargetLang)

        Dim lblDomain As New Label() With {.Text = "领域:", .Location = New Point(185, 28), .AutoSize = True}
        grpLanguage.Controls.Add(lblDomain)

        cbDomain = New ComboBox() With {
            .Location = New Point(225, 25),
            .Size = New Size(150, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        TranslateDomainManager.Load()
        For Each template In TranslateDomainManager.Templates
            cbDomain.Items.Add(template.Name)
        Next
        If cbDomain.Items.Count > 0 Then cbDomain.SelectedIndex = 0
        grpLanguage.Controls.Add(cbDomain)

        btnEditDomain = New Button() With {
            .Text = "编辑",
            .Location = New Point(385, 23),
            .Size = New Size(65, 26)
        }
        AddHandler btnEditDomain.Click, AddressOf EditDomain_Click
        grpLanguage.Controls.Add(btnEditDomain)

        yPos += 85

        ' ========== 翻译范围 ==========
        grpScope = New GroupBox() With {
            .Text = "翻译范围",
            .Location = New Point(10, yPos),
            .Size = New Size(465, 55)
        }
        Me.Controls.Add(grpScope)

        rbAll = New RadioButton() With {
            .Text = If(_appType = "Word", "整个文档", "所有幻灯片"),
            .Location = New Point(15, 22),
            .AutoSize = True,
            .Checked = True
        }
        grpScope.Controls.Add(rbAll)

        rbSelection = New RadioButton() With {
            .Text = If(_hasSelection, "仅选中内容", "仅选中内容（未选中）"),
            .Location = New Point(150, 22),
            .AutoSize = True,
            .Enabled = _hasSelection
        }
        grpScope.Controls.Add(rbSelection)

        yPos += 65

        ' ========== 输出方式 ==========
        grpOutput = New GroupBox() With {
            .Text = "输出方式",
            .Location = New Point(10, yPos),
            .Size = New Size(465, 80)
        }
        Me.Controls.Add(grpOutput)

        If _appType = "Excel" Then
            ' Excel只有3个选项：替换原文、右侧单元格、下方单元格
            rbReplace = New RadioButton() With {
                .Text = "替换原文",
                .Location = New Point(15, 22),
                .AutoSize = True,
                .Checked = True
            }
            AddHandler rbReplace.CheckedChanged, AddressOf OutputModeChanged
            grpOutput.Controls.Add(rbReplace)

            rbImmersive = New RadioButton() With {
                .Text = "译文放在右侧单元格",
                .Location = New Point(15, 48),
                .AutoSize = True
            }
            AddHandler rbImmersive.CheckedChanged, AddressOf OutputModeChanged
            grpOutput.Controls.Add(rbImmersive)

            rbNewDoc = New RadioButton() With {
                .Text = "译文放在下方单元格",
                .Location = New Point(200, 22),
                .AutoSize = True
            }
            AddHandler rbNewDoc.CheckedChanged, AddressOf OutputModeChanged
            grpOutput.Controls.Add(rbNewDoc)

            ' Excel不需要侧栏选项
            rbSidePanel = New RadioButton() With {.Visible = False}
        Else
            ' Word/PowerPoint的原有选项
            ' 沉浸式翻译作为默认选项
            rbImmersive = New RadioButton() With {
                .Text = If(_appType = "Word", "沉浸式翻译（译文放在每段后面）", "沉浸式翻译（译文放在每页后面）"),
                .Location = New Point(15, 22),
                .AutoSize = True,
                .Checked = (_appType = "Word") ' Word默认选沉浸式
            }
            AddHandler rbImmersive.CheckedChanged, AddressOf OutputModeChanged
            grpOutput.Controls.Add(rbImmersive)

            rbSidePanel = New RadioButton() With {
                .Text = "仅显示在侧栏（不修改原文）",
                .Location = New Point(15, 48),
                .AutoSize = True,
                .Checked = (_appType = "PowerPoint") ' PowerPoint默认选侧栏
            }
            AddHandler rbSidePanel.CheckedChanged, AddressOf OutputModeChanged
            grpOutput.Controls.Add(rbSidePanel)

            rbReplace = New RadioButton() With {
                .Text = "替换原文",
                .Location = New Point(280, 22),
                .AutoSize = True
            }
            AddHandler rbReplace.CheckedChanged, AddressOf OutputModeChanged
            grpOutput.Controls.Add(rbReplace)

            rbNewDoc = New RadioButton() With {
                .Text = If(_appType = "Word", "创建新文档", "创建新演示文稿"),
                .Location = New Point(280, 48),
                .AutoSize = True
            }
            AddHandler rbNewDoc.CheckedChanged, AddressOf OutputModeChanged
            grpOutput.Controls.Add(rbNewDoc)
        End If

        yPos += 90

        ' ========== 沉浸式翻译样式（Excel不显示） ==========
        pnlImmersiveStyle = New Panel() With {
            .Location = New Point(10, yPos),
            .Size = New Size(465, 70),
            .Visible = (_appType <> "Excel") ' Excel不显示沉浸式样式面板
        }
        Me.Controls.Add(pnlImmersiveStyle)

        ' 保持格式选项
        chkPreserveFormat = New CheckBox() With {
            .Text = "保持原文格式（译文继承原文样式）",
            .Location = New Point(5, 8),
            .AutoSize = True,
            .Checked = True
        }
        AddHandler chkPreserveFormat.CheckedChanged, AddressOf PreserveFormatChanged
        pnlImmersiveStyle.Controls.Add(chkPreserveFormat)

        Dim lblStyle As New Label() With {.Text = "自定义译文样式:", .Location = New Point(5, 38), .AutoSize = True}
        pnlImmersiveStyle.Controls.Add(lblStyle)

        btnColor = New Button() With {
            .Location = New Point(120, 33),
            .Size = New Size(70, 26),
            .BackColor = selectedColor,
            .FlatStyle = FlatStyle.Flat,
            .Text = "颜色",
            .Enabled = False
        }
        AddHandler btnColor.Click, AddressOf ColorButton_Click
        pnlImmersiveStyle.Controls.Add(btnColor)

        chkItalic = New CheckBox() With {
            .Text = "斜体",
            .Location = New Point(205, 37),
            .AutoSize = True,
            .Checked = True,
            .Enabled = False
        }
        pnlImmersiveStyle.Controls.Add(chkItalic)

        yPos += 75

        ' ========== 进度条 ==========
        progressBar = New ProgressBar() With {
            .Location = New Point(10, yPos),
            .Size = New Size(465, 20),
            .Visible = False
        }
        Me.Controls.Add(progressBar)

        lblProgress = New Label() With {
            .Location = New Point(10, yPos + 22),
            .Size = New Size(465, 20),
            .Text = "",
            .Visible = False
        }
        Me.Controls.Add(lblProgress)

        yPos += 50

        ' ========== 按钮 ==========
        btnTranslate = New Button() With {
            .Text = "开始翻译",
            .Location = New Point(280, yPos),
            .Size = New Size(90, 35),
            .DialogResult = DialogResult.OK
        }
        AddHandler btnTranslate.Click, AddressOf TranslateButton_Click
        Me.Controls.Add(btnTranslate)
        Me.AcceptButton = btnTranslate

        btnCancel = New Button() With {
            .Text = "取消",
            .Location = New Point(380, yPos),
            .Size = New Size(90, 35),
            .DialogResult = DialogResult.Cancel
        }
        Me.Controls.Add(btnCancel)
        Me.CancelButton = btnCancel
    End Sub

    Private Sub LoadSettings()
        ' 加载平台列表
        cbPlatform.Items.Clear()
        Dim hasValidated = False
        Dim selectedPlatformIdx = -1
        Dim idx = 0

        For Each cfg In ConfigManager.ConfigData
            If cfg.validated Then
                cbPlatform.Items.Add(cfg.pltform)
                If cfg.translateSelected Then
                    selectedPlatformIdx = idx
                End If
                hasValidated = True
                idx += 1
            End If
        Next

        If Not hasValidated Then
            cbPlatform.Items.Add("（请先配置API）")
            cbPlatform.SelectedIndex = 0
            cbPlatform.Enabled = False
            cbModel.Enabled = False
            btnTranslate.Enabled = False
        Else
            cbPlatform.SelectedIndex = If(selectedPlatformIdx >= 0, selectedPlatformIdx, 0)
        End If

        ' 加载设置
        Dim settings = TranslateSettings.Load()

        ' 目标语言
        For i = 0 To targetLanguages.Length - 1
            If targetLanguages(i) = settings.TargetLanguage Then
                cbTargetLang.SelectedIndex = i
                Exit For
            End If
        Next

        ' 领域
        For i = 0 To cbDomain.Items.Count - 1
            If cbDomain.Items(i).ToString() = settings.CurrentDomain Then
                cbDomain.SelectedIndex = i
                Exit For
            End If
        Next

        ' 输出模式
        Select Case settings.OutputMode
            Case TranslateOutputMode.Replace : rbReplace.Checked = True
            Case TranslateOutputMode.SidePanel : rbSidePanel.Checked = True
            Case TranslateOutputMode.NewDocument : rbNewDoc.Checked = True
            Case Else : rbImmersive.Checked = True
        End Select

        ' 沉浸式样式
        chkItalic.Checked = settings.ImmersiveTranslationItalic
        chkPreserveFormat.Checked = settings.PreserveFormatting
        Try
            Dim colorHex = settings.ImmersiveTranslationColor.TrimStart("#"c)
            Dim r = Convert.ToInt32(colorHex.Substring(0, 2), 16)
            Dim g = Convert.ToInt32(colorHex.Substring(2, 2), 16)
            Dim b = Convert.ToInt32(colorHex.Substring(4, 2), 16)
            selectedColor = Color.FromArgb(r, g, b)
            btnColor.BackColor = selectedColor
        Catch
        End Try

        ' 更新样式控件状态
        UpdateStyleControlsState()
    End Sub

    Private Sub PlatformChanged(sender As Object, e As EventArgs)
        cbModel.Items.Clear()

        Dim platformName = cbPlatform.SelectedItem?.ToString()
        If String.IsNullOrEmpty(platformName) Then Return

        Dim cfg = ConfigManager.ConfigData.FirstOrDefault(Function(c) c.pltform = platformName)
        If cfg Is Nothing OrElse cfg.model Is Nothing Then Return

        Dim selectedModelIdx = -1
        For i = 0 To cfg.model.Count - 1
            cbModel.Items.Add(cfg.model(i).modelName)
            If cfg.model(i).translateSelected Then
                selectedModelIdx = i
            End If
        Next

        If cbModel.Items.Count > 0 Then
            cbModel.SelectedIndex = If(selectedModelIdx >= 0, selectedModelIdx, 0)
        End If
    End Sub

    Private Sub OutputModeChanged(sender As Object, e As EventArgs)
        ' 只有选择沉浸式翻译且非Excel时才显示样式设置
        pnlImmersiveStyle.Visible = rbImmersive.Checked AndAlso _appType <> "Excel"
    End Sub

    Private Sub PreserveFormatChanged(sender As Object, e As EventArgs)
        UpdateStyleControlsState()
    End Sub

    Private Sub UpdateStyleControlsState()
        ' 当保持原文格式时，禁用自定义样式选项
        Dim enableCustomStyle = Not chkPreserveFormat.Checked
        btnColor.Enabled = enableCustomStyle
        chkItalic.Enabled = enableCustomStyle
    End Sub

    Private Sub ColorButton_Click(sender As Object, e As EventArgs)
        Using dlg As New ColorDialog()
            dlg.Color = selectedColor
            dlg.FullOpen = True
            If dlg.ShowDialog() = DialogResult.OK Then
                selectedColor = dlg.Color
                btnColor.BackColor = selectedColor
            End If
        End Using
    End Sub

    Private Sub EditDomain_Click(sender As Object, e As EventArgs)
        Dim dlg As New TranslateGlobalSettingsForm()
        dlg.Settings = TranslateSettings.Load()
        If dlg.ShowDialog() = DialogResult.OK Then
            dlg.Settings.Save()
            ' 刷新领域列表
            cbDomain.Items.Clear()
            TranslateDomainManager.Load()
            For Each template In TranslateDomainManager.Templates
                cbDomain.Items.Add(template.Name)
            Next
            If cbDomain.Items.Count > 0 Then cbDomain.SelectedIndex = 0
        End If
    End Sub

    Private Sub TranslateButton_Click(sender As Object, e As EventArgs)
        ' 保存平台和模型选择
        Dim platformName = cbPlatform.SelectedItem?.ToString()
        If Not String.IsNullOrEmpty(platformName) Then
            For Each cfg In ConfigManager.ConfigData
                cfg.translateSelected = (cfg.pltform = platformName)
                If cfg.translateSelected AndAlso cfg.model IsNot Nothing Then
                    Dim selectedModelName = cbModel.SelectedItem?.ToString()
                    For Each m In cfg.model
                        m.translateSelected = (m.modelName = selectedModelName)
                    Next
                End If
            Next
            ConfigManager.SaveConfig()
        End If

        ' 设置返回值
        TranslateAll = rbAll.Checked
        SourceLanguage = "auto"
        TargetLanguage = targetLanguages(cbTargetLang.SelectedIndex)
        SelectedDomain = cbDomain.SelectedItem?.ToString()

        If rbReplace.Checked Then
            OutputMode = TranslateOutputMode.Replace
        ElseIf rbSidePanel.Checked Then
            OutputMode = TranslateOutputMode.SidePanel
        ElseIf rbNewDoc.Checked Then
            OutputMode = TranslateOutputMode.NewDocument
        Else
            OutputMode = TranslateOutputMode.Immersive
        End If

        ' 保存设置
        Dim settings = TranslateSettings.Load()
        settings.TargetLanguage = TargetLanguage
        settings.CurrentDomain = SelectedDomain
        settings.OutputMode = OutputMode
        settings.ImmersiveTranslationItalic = chkItalic.Checked
        settings.ImmersiveTranslationColor = $"#{selectedColor.R:X2}{selectedColor.G:X2}{selectedColor.B:X2}"
        settings.PreserveFormatting = chkPreserveFormat.Checked
        settings.Save()
    End Sub

    Public Sub ShowProgress(current As Integer, total As Integer, message As String)
        If Me.InvokeRequired Then
            Me.Invoke(Sub() ShowProgress(current, total, message))
            Return
        End If

        progressBar.Visible = True
        lblProgress.Visible = True
        progressBar.Maximum = total
        progressBar.Value = Math.Min(current, total)
        lblProgress.Text = message
        btnTranslate.Enabled = False
    End Sub

    Public Sub HideProgress()
        If Me.InvokeRequired Then
            Me.Invoke(Sub() HideProgress())
            Return
        End If

        progressBar.Visible = False
        lblProgress.Visible = False
        btnTranslate.Enabled = True
    End Sub
End Class
