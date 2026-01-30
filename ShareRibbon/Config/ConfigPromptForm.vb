Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Imports Newtonsoft.Json

' 大模型提示词配置 - 现代化UI
Public Class ConfigPromptForm
    Inherits Form
    Private ReadOnly _applicationInfo As ApplicationInfo

    Public Shared Property ConfigPromptData As List(Of PromptConfigItem)

    ' UI控件
    Private tabControl As TabControl
    Private tabBasic As TabPage
    Private tabAdvanced As TabPage
    Private tabQuickQuestions As TabPage

    ' 基础配置控件
    Private promptListBox As ListBox
    Private promptNameTextBox As TextBox
    Private promptContentTextBox As TextBox
    Private btnAdd As Button
    Private btnDelete As Button
    Private btnUse As Button
    Private btnSave As Button

    ' 高级配置控件
    Private jsonSchemaTextBox As TextBox
    Private btnSaveSchema As Button
    Private btnResetSchema As Button

    ' 快捷问题控件
    Private quickQuestionsListBox As ListBox
    Private quickQuestionTextBox As TextBox
    Private btnAddQuestion As Button
    Private btnDeleteQuestion As Button
    Private btnSaveQuestions As Button
    Private btnResetQuestions As Button

    ' 快捷问题数据
    Private Shared _quickQuestions As List(Of String)

    ' 默认快捷问题（与前端predefinedPrompts保持一致）
    Private Shared ReadOnly DEFAULT_QUICK_QUESTIONS As String() = {
        "帮我把A列加B列的值写入C列",
        "帮我把Sheet1和Sheet2的表格按名字合并",
        "帮我把Sheet1的数据，按照中文名称拆分成多个xlsx文件",
        "给我将我选中的Word内容格式调整一下",
        "给我生成一个3页的周报PPT文件",
        "什么？没有你想要的，点击此处维护吧"
    }

    Private Const MAX_QUICK_QUESTIONS As Integer = 6

    ' 属性
    Public Property propmtName As String
    Public Property propmtContent As String

    ' 默认提示词
    Private ReadOnly DEFAULT_PROMPTS As New Dictionary(Of String, String) From {
        {"Excel", "你是一名Excel专家，擅长数据分析、公式计算和VBA编程。如果用户需求明确，返回JSON命令执行操作；如果需求不明确，请先询问澄清。"},
        {"Word", "你是一名Word文档专家，擅长文档编辑、格式排版和内容生成。如果用户需求明确，返回JSON命令执行操作；如果需求不明确，请先询问澄清。"},
        {"PowerPoint", "你是一名PowerPoint演示专家，擅长幻灯片设计、动画效果和内容创作。如果用户需求明确，返回JSON命令执行操作；如果需求不明确，请先询问澄清。"}
    }

    Public Sub New(applicationInfo As ApplicationInfo)
        _applicationInfo = applicationInfo
        LoadConfig()
        InitializeUI()
    End Sub

    Private Sub InitializeUI()
        ' 窗体设置
        Me.Text = $"提示词配置 - {_applicationInfo.Type}"
        Me.Size = New Size(600, 520)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' 创建TabControl
        tabControl = New TabControl() With {
            .Location = New Point(10, 10),
            .Size = New Size(565, 420),
            .Font = New Font("Microsoft YaHei UI", 9)
        }

        ' 基础配置页
        tabBasic = New TabPage("聊天提示词")
        InitializeBasicTab()
        tabControl.TabPages.Add(tabBasic)

        ' 高级配置页
        tabAdvanced = New TabPage("JSON格式约束")
        InitializeAdvancedTab()
        tabControl.TabPages.Add(tabAdvanced)

        ' 快捷问题配置页
        tabQuickQuestions = New TabPage("快捷问题")
        InitializeQuickQuestionsTab()
        tabControl.TabPages.Add(tabQuickQuestions)

        Me.Controls.Add(tabControl)

        ' 底部关闭按钮
        Dim btnClose As New Button() With {
            .Text = "关闭",
            .Location = New Point(490, 440),
            .Size = New Size(80, 30)
        }
        AddHandler btnClose.Click, Sub(s, e) Me.Close()
        Me.Controls.Add(btnClose)

        Me.Controls.Add(GlobalStatusStrip.StatusStrip)
    End Sub

    Private Sub InitializeBasicTab()
        ' 说明标签
        Dim lblDesc As New Label() With {
            .Text = "提示词为AI设定身份角色，让回答更专业。选择一个提示词后点击「使用」生效。",
            .Location = New Point(10, 10),
            .Size = New Size(530, 20),
            .ForeColor = Color.Gray
        }
        tabBasic.Controls.Add(lblDesc)

        ' 左侧：提示词列表
        Dim lblList As New Label() With {
            .Text = "已保存的提示词：",
            .Location = New Point(10, 35),
            .AutoSize = True
        }
        tabBasic.Controls.Add(lblList)

        promptListBox = New ListBox() With {
            .Location = New Point(10, 55),
            .Size = New Size(180, 200),
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        AddHandler promptListBox.SelectedIndexChanged, AddressOf PromptListBox_SelectedIndexChanged
        tabBasic.Controls.Add(promptListBox)

        ' 列表操作按钮
        btnUse = New Button() With {
            .Text = "使用选中",
            .Location = New Point(10, 260),
            .Size = New Size(85, 28),
            .BackColor = Color.FromArgb(70, 130, 180),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnUse.Click, AddressOf BtnUse_Click
        tabBasic.Controls.Add(btnUse)

        btnDelete = New Button() With {
            .Text = "删除",
            .Location = New Point(105, 260),
            .Size = New Size(85, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDelete.Click, AddressOf BtnDelete_Click
        tabBasic.Controls.Add(btnDelete)

        ' 右侧：编辑区域
        Dim lblName As New Label() With {
            .Text = "提示词名称：",
            .Location = New Point(210, 35),
            .AutoSize = True
        }
        tabBasic.Controls.Add(lblName)

        promptNameTextBox = New TextBox() With {
            .Location = New Point(210, 55),
            .Size = New Size(330, 25)
        }
        tabBasic.Controls.Add(promptNameTextBox)

        Dim lblContent As New Label() With {
            .Text = "提示词内容：",
            .Location = New Point(210, 85),
            .AutoSize = True
        }
        tabBasic.Controls.Add(lblContent)

        promptContentTextBox = New TextBox() With {
            .Location = New Point(210, 105),
            .Size = New Size(330, 150),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        tabBasic.Controls.Add(promptContentTextBox)

        ' 编辑操作按钮
        btnAdd = New Button() With {
            .Text = "新增/保存",
            .Location = New Point(210, 260),
            .Size = New Size(100, 28),
            .BackColor = Color.FromArgb(60, 179, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAdd.Click, AddressOf BtnAdd_Click
        tabBasic.Controls.Add(btnAdd)

        Dim btnClear As New Button() With {
            .Text = "清空输入",
            .Location = New Point(320, 260),
            .Size = New Size(80, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnClear.Click, Sub(s, e)
                                       promptNameTextBox.Clear()
                                       promptContentTextBox.Clear()
                                       promptListBox.ClearSelected()
                                   End Sub
        tabBasic.Controls.Add(btnClear)

        ' 当前使用的提示词显示
        Dim lblCurrent As New Label() With {
            .Text = "当前使用：",
            .Location = New Point(10, 300),
            .AutoSize = True,
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Bold)
        }
        tabBasic.Controls.Add(lblCurrent)

        Dim lblCurrentValue As New Label() With {
            .Name = "lblCurrentValue",
            .Text = If(String.IsNullOrEmpty(ConfigSettings.propmtName), "(未设置)", ConfigSettings.propmtName),
            .Location = New Point(80, 300),
            .Size = New Size(460, 20),
            .ForeColor = Color.FromArgb(70, 130, 180)
        }
        tabBasic.Controls.Add(lblCurrentValue)

        ' 加载数据到列表
        RefreshPromptList()
    End Sub

    Private Sub InitializeAdvancedTab()
        ' 说明标签
        Dim lblDesc As New Label() With {
            .Text = $"JSON格式约束用于规范AI返回的命令格式，确保可正确解析执行。当前应用：{_applicationInfo.Type}",
            .Location = New Point(10, 10),
            .Size = New Size(530, 20),
            .ForeColor = Color.Gray
        }
        tabAdvanced.Controls.Add(lblDesc)

        Dim lblWarning As New Label() With {
            .Text = "⚠ 修改此内容可能导致命令执行失败，请谨慎操作！",
            .Location = New Point(10, 32),
            .Size = New Size(530, 20),
            .ForeColor = Color.OrangeRed,
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Bold)
        }
        tabAdvanced.Controls.Add(lblWarning)

        ' JSON Schema 编辑框
        jsonSchemaTextBox = New TextBox() With {
            .Location = New Point(10, 55),
            .Size = New Size(530, 270),
            .Multiline = True,
            .ScrollBars = ScrollBars.Both,
            .Font = New Font("Consolas", 9),
            .WordWrap = False
        }
        tabAdvanced.Controls.Add(jsonSchemaTextBox)

        ' 加载当前的 JSON Schema
        LoadJsonSchema()

        ' 操作按钮
        btnSaveSchema = New Button() With {
            .Text = "保存修改",
            .Location = New Point(10, 335),
            .Size = New Size(100, 30),
            .BackColor = Color.FromArgb(60, 179, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSaveSchema.Click, AddressOf BtnSaveSchema_Click
        tabAdvanced.Controls.Add(btnSaveSchema)

        btnResetSchema = New Button() With {
            .Text = "恢复默认",
            .Location = New Point(120, 335),
            .Size = New Size(100, 30),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnResetSchema.Click, AddressOf BtnResetSchema_Click
        tabAdvanced.Controls.Add(btnResetSchema)
    End Sub

    Private Sub LoadJsonSchema()
        Try
            Dim schema = PromptManager.Instance.GetJsonSchemaConstraint(_applicationInfo.Type.ToString())
            jsonSchemaTextBox.Text = If(String.IsNullOrEmpty(schema), "(无配置)", schema)
        Catch ex As Exception
            jsonSchemaTextBox.Text = $"(加载失败: {ex.Message})"
        End Try
    End Sub

    Private Sub BtnSaveSchema_Click(sender As Object, e As EventArgs)
        Try
            ' 保存到 PromptManager
            PromptManager.Instance.UpdateJsonSchemaConstraint(_applicationInfo.Type.ToString(), jsonSchemaTextBox.Text)
            PromptManager.Instance.SavePromptConfiguration()
            GlobalStatusStrip.ShowInfo("JSON格式约束已保存！")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning($"保存失败: {ex.Message}")
        End Try
    End Sub

    Private Sub BtnResetSchema_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("确定要恢复默认的JSON格式约束吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                PromptManager.Instance.ResetJsonSchemaConstraint(_applicationInfo.Type.ToString())
                PromptManager.Instance.SavePromptConfiguration()
                LoadJsonSchema()
                GlobalStatusStrip.ShowInfo("已恢复默认配置！")
            Catch ex As Exception
                GlobalStatusStrip.ShowWarning($"恢复失败: {ex.Message}")
            End Try
        End If
    End Sub

    ' ============ 快捷问题Tab初始化 ============
    Private Sub InitializeQuickQuestionsTab()
        ' 加载快捷问题数据
        LoadQuickQuestions()

        ' 说明标签
        Dim lblDesc As New Label() With {
            .Text = "快捷问题会在输入框中按 # 键时显示，方便快速选择常用问题。最多可维护6条。",
            .Location = New Point(10, 10),
            .Size = New Size(530, 20),
            .ForeColor = Color.Gray
        }
        tabQuickQuestions.Controls.Add(lblDesc)

        ' 左侧：快捷问题列表
        Dim lblList As New Label() With {
            .Text = "已维护的快捷问题：",
            .Location = New Point(10, 35),
            .AutoSize = True
        }
        tabQuickQuestions.Controls.Add(lblList)

        quickQuestionsListBox = New ListBox() With {
            .Location = New Point(10, 55),
            .Size = New Size(530, 150),
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        AddHandler quickQuestionsListBox.SelectedIndexChanged, AddressOf QuickQuestionsListBox_SelectedIndexChanged
        tabQuickQuestions.Controls.Add(quickQuestionsListBox)

        ' 编辑区域
        Dim lblEdit As New Label() With {
            .Text = "编辑问题内容：",
            .Location = New Point(10, 215),
            .AutoSize = True
        }
        tabQuickQuestions.Controls.Add(lblEdit)

        quickQuestionTextBox = New TextBox() With {
            .Location = New Point(10, 235),
            .Size = New Size(530, 25),
            .Font = New Font("Microsoft YaHei UI", 9)
        }
        tabQuickQuestions.Controls.Add(quickQuestionTextBox)

        ' 操作按钮行
        btnAddQuestion = New Button() With {
            .Text = "新增/更新",
            .Location = New Point(10, 270),
            .Size = New Size(90, 28),
            .BackColor = Color.FromArgb(60, 179, 113),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnAddQuestion.Click, AddressOf BtnAddQuestion_Click
        tabQuickQuestions.Controls.Add(btnAddQuestion)

        btnDeleteQuestion = New Button() With {
            .Text = "删除选中",
            .Location = New Point(110, 270),
            .Size = New Size(90, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDeleteQuestion.Click, AddressOf BtnDeleteQuestion_Click
        tabQuickQuestions.Controls.Add(btnDeleteQuestion)

        btnSaveQuestions = New Button() With {
            .Text = "保存配置",
            .Location = New Point(350, 270),
            .Size = New Size(90, 28),
            .BackColor = Color.FromArgb(70, 130, 180),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnSaveQuestions.Click, AddressOf BtnSaveQuestions_Click
        tabQuickQuestions.Controls.Add(btnSaveQuestions)

        btnResetQuestions = New Button() With {
            .Text = "恢复默认",
            .Location = New Point(450, 270),
            .Size = New Size(90, 28),
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnResetQuestions.Click, AddressOf BtnResetQuestions_Click
        tabQuickQuestions.Controls.Add(btnResetQuestions)

        ' 提示信息
        Dim lblTip As New Label() With {
            .Text = "💡 提示：保存后，在聊天输入框中按 # 键即可看到最新的快捷问题列表。",
            .Location = New Point(10, 310),
            .Size = New Size(530, 20),
            .ForeColor = Color.FromArgb(70, 130, 180),
            .Font = New Font("Microsoft YaHei UI", 9, FontStyle.Italic)
        }
        tabQuickQuestions.Controls.Add(lblTip)

        ' 刷新列表
        RefreshQuickQuestionsList()
    End Sub

    Private Sub QuickQuestionsListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        If quickQuestionsListBox.SelectedItem IsNot Nothing Then
            quickQuestionTextBox.Text = quickQuestionsListBox.SelectedItem.ToString()
        End If
    End Sub

    Private Sub BtnAddQuestion_Click(sender As Object, e As EventArgs)
        Dim question = quickQuestionTextBox.Text.Trim()
        If String.IsNullOrEmpty(question) Then
            GlobalStatusStrip.ShowWarning("请输入快捷问题内容！")
            Return
        End If

        If quickQuestionsListBox.SelectedIndex >= 0 Then
            ' 更新选中项
            _quickQuestions(quickQuestionsListBox.SelectedIndex) = question
            GlobalStatusStrip.ShowInfo("已更新快捷问题！")
        Else
            ' 新增
            If _quickQuestions.Count >= MAX_QUICK_QUESTIONS Then
                GlobalStatusStrip.ShowWarning($"最多只能维护{MAX_QUICK_QUESTIONS}条快捷问题！")
                Return
            End If
            _quickQuestions.Add(question)
            GlobalStatusStrip.ShowInfo("已添加快捷问题！")
        End If

        RefreshQuickQuestionsList()
        quickQuestionTextBox.Clear()
        quickQuestionsListBox.ClearSelected()
    End Sub

    Private Sub BtnDeleteQuestion_Click(sender As Object, e As EventArgs)
        If quickQuestionsListBox.SelectedIndex < 0 Then
            GlobalStatusStrip.ShowWarning("请先选择要删除的快捷问题！")
            Return
        End If

        Dim selectedIndex = quickQuestionsListBox.SelectedIndex
        If MessageBox.Show($"确定要删除「{_quickQuestions(selectedIndex)}」吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            _quickQuestions.RemoveAt(selectedIndex)
            RefreshQuickQuestionsList()
            quickQuestionTextBox.Clear()
            GlobalStatusStrip.ShowInfo("已删除！")
        End If
    End Sub

    Private Sub BtnSaveQuestions_Click(sender As Object, e As EventArgs)
        Try
            SaveQuickQuestions()
            GlobalStatusStrip.ShowInfo("快捷问题配置已保存！重新打开聊天面板后生效。")
        Catch ex As Exception
            GlobalStatusStrip.ShowWarning($"保存失败: {ex.Message}")
        End Try
    End Sub

    Private Sub BtnResetQuestions_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("确定要恢复默认的快捷问题吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            _quickQuestions = DEFAULT_QUICK_QUESTIONS.ToList()
            RefreshQuickQuestionsList()
            SaveQuickQuestions()
            GlobalStatusStrip.ShowInfo("已恢复默认快捷问题！")
        End If
    End Sub

    Private Sub RefreshQuickQuestionsList()
        quickQuestionsListBox.Items.Clear()
        For Each q In _quickQuestions
            quickQuestionsListBox.Items.Add(q)
        Next
    End Sub

    ' ============ 快捷问题数据持久化 ============
    Private Sub LoadQuickQuestions()
        _quickQuestions = New List(Of String)()
        Dim filePath = GetQuickQuestionsFilePath()

        If File.Exists(filePath) Then
            Try
                Dim json = File.ReadAllText(filePath)
                _quickQuestions = JsonConvert.DeserializeObject(Of List(Of String))(json)
            Catch ex As Exception
                Debug.WriteLine($"加载快捷问题失败: {ex.Message}")
            End Try
        End If

        ' 如果为空，使用默认值
        If _quickQuestions Is Nothing OrElse _quickQuestions.Count = 0 Then
            _quickQuestions = DEFAULT_QUICK_QUESTIONS.ToList()
        End If
    End Sub

    Private Sub SaveQuickQuestions()
        Dim filePath = GetQuickQuestionsFilePath()
        Dim dir = Path.GetDirectoryName(filePath)
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If

        Dim json = JsonConvert.SerializeObject(_quickQuestions, Formatting.Indented)
        File.WriteAllText(filePath, json)
    End Sub

    Private Function GetQuickQuestionsFilePath() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "OfficeAiAppData",
            "quick_questions_config.json")
    End Function

    ''' <summary>
    ''' 获取当前快捷问题列表（供HTML页面调用）
    ''' </summary>
    Public Shared Function GetQuickQuestionsList() As List(Of String)
        If _quickQuestions IsNot Nothing AndAlso _quickQuestions.Count > 0 Then
            Return _quickQuestions
        End If

        ' 尝试从文件加载
        Dim filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "OfficeAiAppData",
            "quick_questions_config.json")

        If File.Exists(filePath) Then
            Try
                Dim json = File.ReadAllText(filePath)
                Dim questions = JsonConvert.DeserializeObject(Of List(Of String))(json)
                If questions IsNot Nothing AndAlso questions.Count > 0 Then
                    Return questions
                End If
            Catch ex As Exception
                Debug.WriteLine($"读取快捷问题失败: {ex.Message}")
            End Try
        End If

        ' 返回默认值
        Return DEFAULT_QUICK_QUESTIONS.ToList()
    End Function

    Private Sub RefreshPromptList()
        promptListBox.Items.Clear()
        For Each item In ConfigPromptData
            promptListBox.Items.Add(item)
        Next

        ' 选中当前使用的
        For i As Integer = 0 To promptListBox.Items.Count - 1
            If CType(promptListBox.Items(i), PromptConfigItem).selected Then
                promptListBox.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub PromptListBox_SelectedIndexChanged(sender As Object, e As EventArgs)
        If promptListBox.SelectedItem IsNot Nothing Then
            Dim item = CType(promptListBox.SelectedItem, PromptConfigItem)
            promptNameTextBox.Text = item.name
            promptContentTextBox.Text = item.content
        End If
    End Sub

    Private Sub BtnUse_Click(sender As Object, e As EventArgs)
        If promptListBox.SelectedItem Is Nothing Then
            GlobalStatusStrip.ShowWarning("请先选择一个提示词！")
            Return
        End If

        Dim selectedItem = CType(promptListBox.SelectedItem, PromptConfigItem)

        ' 更新选中状态
        For Each item In ConfigPromptData
            item.selected = (item.name = selectedItem.name)
        Next

        ' 保存并更新全局配置
        SaveConfig()
        ConfigSettings.propmtName = selectedItem.name
        ConfigSettings.propmtContent = selectedItem.content

        ' 更新显示
        Dim lblCurrentValue = tabBasic.Controls.Find("lblCurrentValue", False).FirstOrDefault()
        If lblCurrentValue IsNot Nothing Then
            lblCurrentValue.Text = selectedItem.name
        End If

        GlobalStatusStrip.ShowInfo($"已启用提示词：{selectedItem.name}")
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs)
        Dim name = promptNameTextBox.Text.Trim()
        Dim content = promptContentTextBox.Text.Trim()

        If String.IsNullOrEmpty(name) Then
            GlobalStatusStrip.ShowWarning("请输入提示词名称！")
            Return
        End If

        If String.IsNullOrEmpty(content) Then
            GlobalStatusStrip.ShowWarning("请输入提示词内容！")
            Return
        End If

        ' 检查是否存在
        Dim existingItem = ConfigPromptData.FirstOrDefault(Function(item) item.name = name)
        If existingItem IsNot Nothing Then
            ' 更新
            existingItem.content = content
            GlobalStatusStrip.ShowInfo($"已更新提示词：{name}")
        Else
            ' 新增
            ConfigPromptData.Add(New PromptConfigItem() With {
                .name = name,
                .content = content,
                .selected = False
            })
            GlobalStatusStrip.ShowInfo($"已添加提示词：{name}")
        End If

        SaveConfig()
        RefreshPromptList()
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As EventArgs)
        If promptListBox.SelectedItem Is Nothing Then
            GlobalStatusStrip.ShowWarning("请先选择要删除的提示词！")
            Return
        End If

        Dim selectedItem = CType(promptListBox.SelectedItem, PromptConfigItem)

        If selectedItem.selected Then
            GlobalStatusStrip.ShowWarning("不能删除当前正在使用的提示词！")
            Return
        End If

        If MessageBox.Show($"确定要删除「{selectedItem.name}」吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ConfigPromptData.Remove(selectedItem)
            SaveConfig()
            RefreshPromptList()
            promptNameTextBox.Clear()
            promptContentTextBox.Clear()
            GlobalStatusStrip.ShowInfo("已删除！")
        End If
    End Sub

    Public Sub LoadConfig()
        ConfigPromptData = New List(Of PromptConfigItem)()

        If File.Exists(configFilePath) Then
            Try
                Dim json As String = File.ReadAllText(configFilePath)
                ConfigPromptData = JsonConvert.DeserializeObject(Of List(Of PromptConfigItem))(json)
            Catch ex As Exception
                Debug.WriteLine($"加载提示词配置失败: {ex.Message}")
            End Try
        End If

        ' 如果为空，添加默认配置
        If ConfigPromptData Is Nothing OrElse ConfigPromptData.Count = 0 Then
            ConfigPromptData = New List(Of PromptConfigItem)()
            Dim defaultPrompt = GetDefaultPrompt()
            ConfigPromptData.Add(defaultPrompt)
            SaveConfig()
        End If

        ' 初始化全局配置
        For Each item In ConfigPromptData
            If item.selected Then
                ConfigSettings.propmtName = item.name
                ConfigSettings.propmtContent = item.content
                Exit For
            End If
        Next
    End Sub

    Private Function GetDefaultPrompt() As PromptConfigItem
        Dim appType = _applicationInfo.Type.ToString()
        Dim content = If(DEFAULT_PROMPTS.ContainsKey(appType), DEFAULT_PROMPTS(appType), "你是一名Office办公专家。")

        Return New PromptConfigItem() With {
            .name = $"{appType}助手",
            .content = content,
            .selected = True
        }
    End Function

    Public Sub SaveConfig()
        Try
            Dim dir = Path.GetDirectoryName(configFilePath)
            If Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim json As String = JsonConvert.SerializeObject(ConfigPromptData, Formatting.Indented)
            File.WriteAllText(configFilePath, json)
        Catch ex As Exception
            Debug.WriteLine($"保存提示词配置失败: {ex.Message}")
        End Try
    End Sub

    Private ReadOnly Property configFilePath As String
        Get
            Return _applicationInfo.GetPromptConfigFilePath()
        End Get
    End Property

    ' 提示词配置项
    Public Class PromptConfigItem
        Public Property name As String
        Public Property content As String
        Public Property selected As Boolean
        Public Overrides Function ToString() As String
            Return name
        End Function
    End Class
End Class
