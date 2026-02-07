Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' 翻译领域模板管理对话框 - 简化版，专注于领域管理
''' </summary>
Public Class TranslateGlobalSettingsForm
    Inherits Form

    Public Property Settings As TranslateSettings

    ' 领域管理
    Private grpDomain As GroupBox
    Private lstDomains As ListBox
    Private txtDomainDesc As TextBox
    Private btnAddDomain As Button
    Private btnEditDomain As Button
    Private btnDeleteDomain As Button

    ' 高级设置
    Private grpAdvanced As GroupBox
    Private numBatchSize As NumericUpDown
    Private chkShowProgress As CheckBox

    Private btnOk As Button
    Private btnCancel As Button

    Public Sub New()
        Me.Text = "翻译领域管理"
        Me.Size = New Size(520, 480)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        InitializeControls()
    End Sub

    Private Sub InitializeControls()
        Dim yPos = 10

        ' ========== 领域模板管理 ==========
        grpDomain = New GroupBox() With {
            .Text = "翻译领域模板",
            .Location = New Point(10, yPos),
            .Size = New Size(485, 320)
        }
        Me.Controls.Add(grpDomain)

        lstDomains = New ListBox() With {
            .Location = New Point(15, 25),
            .Size = New Size(200, 200)
        }
        AddHandler lstDomains.SelectedIndexChanged, AddressOf DomainSelected
        grpDomain.Controls.Add(lstDomains)

        Dim lblDesc As New Label() With {
            .Text = "领域提示词预览:",
            .Location = New Point(230, 25),
            .AutoSize = True
        }
        grpDomain.Controls.Add(lblDesc)

        txtDomainDesc = New TextBox() With {
            .Location = New Point(230, 45),
            .Size = New Size(240, 180),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .ReadOnly = True,
            .BackColor = Color.White
        }
        grpDomain.Controls.Add(txtDomainDesc)

        btnAddDomain = New Button() With {
            .Text = "添加领域",
            .Location = New Point(15, 235),
            .Size = New Size(90, 28)
        }
        AddHandler btnAddDomain.Click, AddressOf AddDomain_Click
        grpDomain.Controls.Add(btnAddDomain)

        btnEditDomain = New Button() With {
            .Text = "编辑",
            .Location = New Point(115, 235),
            .Size = New Size(50, 28)
        }
        AddHandler btnEditDomain.Click, AddressOf EditDomain_Click
        grpDomain.Controls.Add(btnEditDomain)

        btnDeleteDomain = New Button() With {
            .Text = "删除",
            .Location = New Point(170, 235),
            .Size = New Size(50, 28)
        }
        AddHandler btnDeleteDomain.Click, AddressOf DeleteDomain_Click
        grpDomain.Controls.Add(btnDeleteDomain)

        Dim lblTip As New Label() With {
            .Text = "提示: 内置领域不可删除，您可以添加自定义领域模板",
            .Location = New Point(15, 275),
            .Size = New Size(455, 35),
            .ForeColor = Color.Gray
        }
        grpDomain.Controls.Add(lblTip)

        yPos += 330

        ' ========== 高级设置 ==========
        grpAdvanced = New GroupBox() With {
            .Text = "高级设置",
            .Location = New Point(10, yPos),
            .Size = New Size(485, 60)
        }
        Me.Controls.Add(grpAdvanced)

        Dim lblBatch As New Label() With {.Text = "每批翻译段落数:", .Location = New Point(15, 25), .AutoSize = True}
        grpAdvanced.Controls.Add(lblBatch)

        numBatchSize = New NumericUpDown() With {
            .Location = New Point(120, 22),
            .Size = New Size(60, 24),
            .Minimum = 0,
            .Maximum = 20,
            .Value = 5
        }
        grpAdvanced.Controls.Add(numBatchSize)

        chkShowProgress = New CheckBox() With {
            .Text = "显示翻译进度",
            .Location = New Point(200, 24),
            .AutoSize = True,
            .Checked = True
        }
        grpAdvanced.Controls.Add(chkShowProgress)

        yPos += 70

        ' ========== 按钮 ==========
        btnOk = New Button() With {
            .Text = "确定",
            .Location = New Point(320, yPos),
            .Size = New Size(80, 32),
            .DialogResult = DialogResult.OK
        }
        AddHandler btnOk.Click, AddressOf OkButton_Click
        Me.Controls.Add(btnOk)
        Me.AcceptButton = btnOk

        btnCancel = New Button() With {
            .Text = "取消",
            .Location = New Point(410, yPos),
            .Size = New Size(80, 32),
            .DialogResult = DialogResult.Cancel
        }
        Me.Controls.Add(btnCancel)
        Me.CancelButton = btnCancel
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        If Settings Is Nothing Then Settings = New TranslateSettings()

        numBatchSize.Value = Settings.BatchSize
        chkShowProgress.Checked = Settings.ShowProgress

        LoadDomainTemplates()
    End Sub

    Private Sub LoadDomainTemplates()
        lstDomains.Items.Clear()
        TranslateDomainManager.Load()
        For Each template In TranslateDomainManager.Templates
            Dim displayText = template.Name
            If template.IsBuiltIn Then displayText &= " [内置]"
            lstDomains.Items.Add(displayText)
        Next

        If lstDomains.Items.Count > 0 Then
            lstDomains.SelectedIndex = 0
        End If
    End Sub

    Private Sub DomainSelected(sender As Object, e As EventArgs)
        If lstDomains.SelectedIndex < 0 OrElse lstDomains.SelectedIndex >= TranslateDomainManager.Templates.Count Then
            txtDomainDesc.Text = ""
            Return
        End If

        Dim template = TranslateDomainManager.Templates(lstDomains.SelectedIndex)
        txtDomainDesc.Text = template.SystemPrompt
    End Sub

    Private Sub AddDomain_Click(sender As Object, e As EventArgs)
        Dim dlg As New DomainTemplateEditForm()
        If dlg.ShowDialog() = DialogResult.OK Then
            TranslateDomainManager.AddTemplate(dlg.Template)
            LoadDomainTemplates()
        End If
    End Sub

    Private Sub EditDomain_Click(sender As Object, e As EventArgs)
        If lstDomains.SelectedIndex < 0 Then
            MessageBox.Show("请先选择要编辑的领域模板。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim template = TranslateDomainManager.Templates(lstDomains.SelectedIndex)
        If template.IsBuiltIn Then
            MessageBox.Show("内置模板不可编辑，您可以添加新的自定义模板。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim dlg As New DomainTemplateEditForm()
        dlg.Template = template
        If dlg.ShowDialog() = DialogResult.OK Then
            TranslateDomainManager.Save()
            LoadDomainTemplates()
        End If
    End Sub

    Private Sub DeleteDomain_Click(sender As Object, e As EventArgs)
        If lstDomains.SelectedIndex < 0 Then
            MessageBox.Show("请先选择要删除的领域模板。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim template = TranslateDomainManager.Templates(lstDomains.SelectedIndex)
        If template.IsBuiltIn Then
            MessageBox.Show("内置模板不可删除。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        If MessageBox.Show($"确定要删除领域模板 '{template.Name}' 吗？", "确认删除",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            TranslateDomainManager.RemoveTemplate(template.Name)
            LoadDomainTemplates()
        End If
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        Settings.BatchSize = CInt(numBatchSize.Value)
        Settings.ShowProgress = chkShowProgress.Checked
    End Sub
End Class

''' <summary>
''' 领域模板编辑对话框
''' </summary>
Public Class DomainTemplateEditForm
    Inherits Form

    Public Property Template As TranslateDomainTemplate

    Private txtName As TextBox
    Private txtDescription As TextBox
    Private txtSystemPrompt As TextBox
    Private btnOk As Button
    Private btnCancel As Button

    Public Sub New()
        Me.Text = "编辑领域模板"
        Me.Size = New Size(500, 420)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        InitializeControls()
    End Sub

    Private Sub InitializeControls()
        Dim yPos = 15

        Dim lblName As New Label() With {.Text = "领域名称:", .Location = New Point(12, yPos + 3), .AutoSize = True}
        Me.Controls.Add(lblName)

        txtName = New TextBox() With {
            .Location = New Point(90, yPos),
            .Size = New Size(380, 24)
        }
        Me.Controls.Add(txtName)
        yPos += 35

        Dim lblDesc As New Label() With {.Text = "领域描述:", .Location = New Point(12, yPos + 3), .AutoSize = True}
        Me.Controls.Add(lblDesc)

        txtDescription = New TextBox() With {
            .Location = New Point(90, yPos),
            .Size = New Size(380, 24)
        }
        Me.Controls.Add(txtDescription)
        yPos += 35

        Dim lblPrompt As New Label() With {.Text = "系统提示词:", .Location = New Point(12, yPos), .AutoSize = True}
        Me.Controls.Add(lblPrompt)
        yPos += 22

        txtSystemPrompt = New TextBox() With {
            .Location = New Point(12, yPos),
            .Size = New Size(458, 220),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical
        }
        Me.Controls.Add(txtSystemPrompt)
        yPos += 230

        btnOk = New Button() With {
            .Text = "保存",
            .Location = New Point(300, yPos),
            .Size = New Size(80, 32),
            .DialogResult = DialogResult.OK
        }
        AddHandler btnOk.Click, AddressOf OkButton_Click
        Me.Controls.Add(btnOk)
        Me.AcceptButton = btnOk

        btnCancel = New Button() With {
            .Text = "取消",
            .Location = New Point(390, yPos),
            .Size = New Size(80, 32),
            .DialogResult = DialogResult.Cancel
        }
        Me.Controls.Add(btnCancel)
        Me.CancelButton = btnCancel
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        If Template IsNot Nothing Then
            txtName.Text = Template.Name
            txtDescription.Text = Template.Description
            txtSystemPrompt.Text = Template.SystemPrompt
            txtName.ReadOnly = Template.IsBuiltIn
        Else
            Template = New TranslateDomainTemplate()
        End If
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtName.Text) Then
            MessageBox.Show("请输入领域名称。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Me.DialogResult = DialogResult.None
            Return
        End If

        Template.Name = txtName.Text.Trim()
        Template.Description = txtDescription.Text.Trim()
        Template.SystemPrompt = txtSystemPrompt.Text
    End Sub
End Class
