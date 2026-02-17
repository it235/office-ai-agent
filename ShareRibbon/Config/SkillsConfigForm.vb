' ShareRibbon\Config\SkillsConfigForm.vb
' Skills / 场景提示词配置（prompt_template 表）

Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' 场景与 Skills 配置窗口：管理 prompt_template 表中的系统提示词与 Skills
''' </summary>
Public Class SkillsConfigForm
    Inherits Form

    Private scenarioCombo As ComboBox
    Private listBox As ListBox
    Private txtName As TextBox
    Private txtContent As TextBox
    Private txtSupportedApps As TextBox
    Private chkIsSkill As CheckBox
    Private lblSupportedApps As Label
    Private _records As New List(Of PromptTemplateRecord)()
    Private _currentScenario As String = "excel"

    Public Sub New()
        Me.Text = "场景与 Skills 配置"
        Me.Size = New Size(700, 500)
        Me.MinimumSize = New Size(500, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Font = New Font("Microsoft YaHei UI", 9)
        AddHandler Me.FormClosing, AddressOf OnFormClosing
        AddHandler Me.Shown, AddressOf OnFormShown
        InitializeUI()
    End Sub

    Private Sub OnFormShown(sender As Object, e As EventArgs)
        LoadScenario(_currentScenario)
    End Sub

    Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
        If Me.Controls.Contains(GlobalStatusStrip.StatusStrip) Then
            Me.Controls.Remove(GlobalStatusStrip.StatusStrip)
        End If
    End Sub

    Private Sub InitializeUI()
        Dim y As Integer = 15

        ' 场景选择
        Dim lblScenario As New Label() With {.Text = "场景：", .Location = New Point(15, y + 3), .Size = New Size(45, 20), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        Me.Controls.Add(lblScenario)
        scenarioCombo = New ComboBox() With {
            .Location = New Point(65, y),
            .Size = New Size(120, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left
        }
        scenarioCombo.Items.AddRange({"excel", "word", "ppt", "common"})
        scenarioCombo.SelectedIndex = 0
        AddHandler scenarioCombo.SelectedIndexChanged, AddressOf ScenarioChanged
        Me.Controls.Add(scenarioCombo)

        ' 可拖拽分隔的左右布局
        Dim split As New SplitContainer() With {
            .Location = New Point(15, 52),
            .Size = New Size(665, 305),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .SplitterDistance = 200,
            .Panel1MinSize = 80,
            .Panel2MinSize = 200,
            .FixedPanel = FixedPanel.None
        }
        split.Panel1.SuspendLayout()
        split.Panel2.SuspendLayout()

        ' 左侧：列表（可横向拖拽调整宽度）
        Dim lblList As New Label() With {.Text = "列表（系统提示词 is_skill=0，Skills is_skill=1）", .Location = New Point(0, 0), .Size = New Size(180, 20), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        split.Panel1.Controls.Add(lblList)
        listBox = New ListBox() With {
            .Location = New Point(0, 22),
            .Size = New Size(196, 278),
            .DisplayMember = "DisplayText",
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right,
            .HorizontalScrollbar = True,
            .HorizontalExtent = 3000
        }
        AddHandler listBox.SelectedIndexChanged, AddressOf ListSelectionChanged
        split.Panel1.Controls.Add(listBox)

        ' 右侧：编辑区
        Dim xRight As Integer = 10
        Dim p2Y As Integer = 0
        Dim lblName As New Label() With {.Text = "名称：", .Location = New Point(xRight, p2Y + 2), .Size = New Size(60, 20), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        split.Panel2.Controls.Add(lblName)
        txtName = New TextBox() With {.Location = New Point(xRight + 60, p2Y), .Size = New Size(400, 24), .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right}
        split.Panel2.Controls.Add(txtName)
        p2Y += 35
        Dim lblContent As New Label() With {.Text = "内容：", .Location = New Point(xRight, p2Y), .Size = New Size(60, 20), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        split.Panel2.Controls.Add(lblContent)
        txtContent = New TextBox() With {
            .Location = New Point(xRight + 60, p2Y),
            .Size = New Size(400, 130),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        }
        split.Panel2.Controls.Add(txtContent)
        p2Y += 140
        chkIsSkill = New CheckBox() With {.Text = "是 Skill（is_skill=1）", .Location = New Point(xRight + 60, p2Y), .Size = New Size(180, 24), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        AddHandler chkIsSkill.CheckedChanged, AddressOf IsSkillChanged
        split.Panel2.Controls.Add(chkIsSkill)
        p2Y += 30
        lblSupportedApps = New Label() With {.Text = "supported_apps：", .Location = New Point(xRight, p2Y), .Size = New Size(100, 20), .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        split.Panel2.Controls.Add(lblSupportedApps)
        txtSupportedApps = New TextBox() With {.Location = New Point(xRight + 105, p2Y - 2), .Size = New Size(350, 24), .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right}
        txtSupportedApps.Visible = False
        split.Panel2.Controls.Add(txtSupportedApps)

        split.Panel1.ResumeLayout(False)
        split.Panel2.ResumeLayout(False)
        Me.Controls.Add(split)

        ' 按钮
        y = 365
        Dim btnAdd As New Button() With {.Text = "新增", .Location = New Point(15, y), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnAdd.Click, AddressOf BtnAddClick
        Me.Controls.Add(btnAdd)
        Dim btnSave As New Button() With {.Text = "保存", .Location = New Point(90, y), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnSave.Click, AddressOf BtnSaveClick
        Me.Controls.Add(btnSave)
        Dim btnDelete As New Button() With {.Text = "删除", .Location = New Point(165, y), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnDelete.Click, AddressOf BtnDeleteClick
        Me.Controls.Add(btnDelete)
        Dim btnImport As New Button() With {.Text = "导入", .Location = New Point(245, y), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnImport.Click, AddressOf BtnImportClick
        Me.Controls.Add(btnImport)
        Dim btnCopy As New Button() With {.Text = "复制选中", .Location = New Point(325, y), .Size = New Size(70, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Left}
        AddHandler btnCopy.Click, Sub(s, ev)
                                     If listBox.SelectedItem IsNot Nothing Then
                                         Try
                                             Clipboard.SetText(listBox.SelectedItem.ToString())
                                             GlobalStatusStrip.ShowInfo("已复制")
                                         Catch ex As Exception
                                             GlobalStatusStrip.ShowWarning("复制失败: " & ex.Message)
                                         End Try
                                     Else
                                         GlobalStatusStrip.ShowWarning("请先选择一项")
                                     End If
                                 End Sub
        Me.Controls.Add(btnCopy)

        Dim btnClose As New Button() With {.Text = "关闭", .Location = New Point(600, y), .Size = New Size(80, 28), .Anchor = AnchorStyles.Bottom Or AnchorStyles.Right}
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    Private Sub ScenarioChanged(sender As Object, e As EventArgs)
        _currentScenario = If(scenarioCombo.SelectedItem Is Nothing, "excel", scenarioCombo.SelectedItem.ToString())
        LoadScenario(_currentScenario)
    End Sub

    Private Sub LoadScenario(scenario As String)
        Try
            OfficeAiDatabase.EnsureInitialized()
            _records = PromptTemplateRepository.ListByScenario(scenario)
            listBox.DataSource = Nothing
            listBox.Items.Clear()
            For Each r In _records
                Dim suffix = If(r.IsSkill = 1, " [Skill]", " [系统]")
                listBox.Items.Add(New ListItem With {.Record = r, .DisplayText = (If(r.TemplateName, "(未命名)") & suffix)})
            Next
        Catch ex As Exception
            listBox.Items.Clear()
            listBox.Items.Add("(加载失败: " & ex.Message & ")")
            GlobalStatusStrip.ShowWarning("加载失败")
        End Try
    End Sub

    Private Sub ListSelectionChanged(sender As Object, e As EventArgs)
        Dim item = TryCast(listBox.SelectedItem, ListItem)
        If item Is Nothing Then
            txtName.Text = ""
            txtContent.Text = ""
            chkIsSkill.Checked = False
            txtSupportedApps.Text = ""
            Return
        End If
        Dim r = item.Record
        txtName.Text = r.TemplateName
        txtContent.Text = r.Content
        chkIsSkill.Checked = (r.IsSkill = 1)
        txtSupportedApps.Visible = chkIsSkill.Checked
        lblSupportedApps.Visible = chkIsSkill.Checked
        If chkIsSkill.Checked AndAlso Not String.IsNullOrEmpty(r.ExtraJson) Then
            Try
                Dim jo = JObject.Parse(r.ExtraJson)
                Dim arr = If(jo("supported_apps"), jo("supportedApps"))
                If arr IsNot Nothing AndAlso TypeOf arr Is JArray Then
                    txtSupportedApps.Text = String.Join(", ", arr.Select(Function(t) t.ToString()))
                Else
                    txtSupportedApps.Text = ""
                End If
            Catch
                txtSupportedApps.Text = r.ExtraJson
            End Try
        Else
            txtSupportedApps.Text = ""
        End If
    End Sub

    Private Sub IsSkillChanged(sender As Object, e As EventArgs)
        txtSupportedApps.Visible = chkIsSkill.Checked
        lblSupportedApps.Visible = chkIsSkill.Checked
    End Sub

    Private Sub BtnAddClick(sender As Object, e As EventArgs)
        txtName.Text = ""
        txtContent.Text = ""
        chkIsSkill.Checked = False
        txtSupportedApps.Text = ""
        txtSupportedApps.Visible = False
        lblSupportedApps.Visible = False
    End Sub

    Private Sub BtnSaveClick(sender As Object, e As EventArgs)
        Try
            OfficeAiDatabase.EnsureInitialized()
            Dim extra As String = ""
            If chkIsSkill.Checked AndAlso Not String.IsNullOrEmpty(txtSupportedApps.Text.Trim()) Then
                Dim apps = txtSupportedApps.Text.Split({","c}, StringSplitOptions.RemoveEmptyEntries)
                Dim arr As New JArray()
                For Each a In apps
                    arr.Add(a.Trim())
                Next
                extra = New JObject() From {{"supported_apps", arr}}.ToString()
            End If

            Dim item = TryCast(listBox.SelectedItem, ListItem)
            If item IsNot Nothing Then
                Dim r = item.Record
                r.TemplateName = txtName.Text.Trim()
                r.Content = txtContent.Text
                r.IsSkill = If(chkIsSkill.Checked, 1, 0)
                r.ExtraJson = extra
                r.Scenario = _currentScenario
                PromptTemplateRepository.Update(r)
                GlobalStatusStrip.ShowInfo("已保存")
            Else
                Dim r As New PromptTemplateRecord With {
                    .TemplateName = txtName.Text.Trim(),
                    .Content = txtContent.Text,
                    .IsSkill = If(chkIsSkill.Checked, 1, 0),
                    .ExtraJson = extra,
                    .Scenario = _currentScenario,
                    .Sort = _records.Count
                }
                PromptTemplateRepository.Insert(r)
                GlobalStatusStrip.ShowInfo("已新增")
            End If
            LoadScenario(_currentScenario)
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("保存失败: " & ex.Message)
        End Try
    End Sub

    Private Sub BtnImportClick(sender As Object, e As EventArgs)
        Using dlg As New OpenFileDialog()
            dlg.Filter = "JSON 或 Markdown (*.json;*.md)|*.json;*.md|JSON (*.json)|*.json|Markdown (*.md)|*.md|All (*.*)|*.*"
            dlg.FilterIndex = 1
            If dlg.ShowDialog() <> DialogResult.OK Then Return
            Try
                Dim content = File.ReadAllText(dlg.FileName)
                Dim ext = Path.GetExtension(dlg.FileName).ToLowerInvariant()
                Dim name = Path.GetFileNameWithoutExtension(dlg.FileName)
                Dim record As PromptTemplateRecord = Nothing

                If ext = ".json" Then
                    Dim jo = JObject.Parse(content)
                    Dim pt = jo("promptTemplate")
                    Dim ct = jo("content")
                    Dim pm = jo("prompt")
                    Dim promptTemplate = If(pt IsNot Nothing, pt.ToString(), If(ct IsNot Nothing, ct.ToString(), If(pm IsNot Nothing, pm.ToString(), "")))
                    If String.IsNullOrWhiteSpace(promptTemplate) Then
                        GlobalStatusStrip.ShowWarning("JSON 中需包含 promptTemplate、content 或 prompt 字段")
                        Return
                    End If
                    Dim sn = jo("skillName")
                    Dim nm = jo("name")
                    Dim skillName = If(sn IsNot Nothing, sn.ToString(), If(nm IsNot Nothing, nm.ToString(), name))
                    Dim supportedApps = If(jo("supported_apps"), jo("supportedApps"))
                    Dim extraJo As New JObject()
                    If supportedApps IsNot Nothing AndAlso TypeOf supportedApps Is JArray Then
                        extraJo("supported_apps") = supportedApps
                    End If
                    Dim params = If(jo("parameters"), jo("params"))
                    If params IsNot Nothing Then extraJo("parameters") = params
                    Dim extra = If(extraJo.Count > 0, extraJo.ToString(), "")
                    record = New PromptTemplateRecord With {
                        .TemplateName = skillName,
                        .Content = promptTemplate,
                        .IsSkill = 1,
                        .ExtraJson = extra,
                        .Scenario = _currentScenario,
                        .Sort = _records.Count
                    }
                Else
                    record = New PromptTemplateRecord With {
                        .TemplateName = name,
                        .Content = content,
                        .IsSkill = 1,
                        .ExtraJson = "",
                        .Scenario = _currentScenario,
                        .Sort = _records.Count
                    }
                End If
                PromptTemplateRepository.Insert(record)
                GlobalStatusStrip.ShowInfo("已导入 Skill: " & record.TemplateName)
                LoadScenario(_currentScenario)
            Catch ex As Exception
                GlobalStatusStrip.ShowWarning("导入失败: " & ex.Message)
            End Try
        End Using
    End Sub

    Private Sub BtnDeleteClick(sender As Object, e As EventArgs)
        Dim item = TryCast(listBox.SelectedItem, ListItem)
        If item Is Nothing Then
            GlobalStatusStrip.ShowWarning("请先选择一项")
            Return
        End If
        If MessageBox.Show("确定删除此项？", "确认", MessageBoxButtons.YesNo) <> DialogResult.Yes Then Return
        Try
            PromptTemplateRepository.Delete(item.Record.Id)
            GlobalStatusStrip.ShowInfo("已删除")
            LoadScenario(_currentScenario)
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning("删除失败: " & ex.Message)
        End Try
    End Sub

    Private Class ListItem
        Public Property Record As PromptTemplateRecord
        Public Property DisplayText As String
        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class
End Class
