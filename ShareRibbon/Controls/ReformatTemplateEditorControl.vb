' ShareRibbon\Controls\ReformatTemplateEditorControl.vb
' 排版模板编辑器控件 - 用于CustomTaskPane

Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms
Imports Newtonsoft.Json

''' <summary>
''' 排版模板编辑器控件 - 使用TabControl组织各设置区域
''' </summary>
Public Class ReformatTemplateEditorControl
    Inherits UserControl

    Private _template As ReformatTemplate
    Private _isNewTemplate As Boolean
    Private _previewCallback As PreviewStyleCallback
    Private _placeholderPreviewCallback As TemplatePlaceholderPreviewCallback

    ' 跟踪上一次选中的索引（用于切换前保存）
    Private _lastLayoutIndex As Integer = -1
    Private _lastStyleIndex As Integer = -1

    ' 防止递归更新的标志
    Private _isUpdatingList As Boolean = False

    ' 顶部工具栏
    Private pnlToolbar As Panel
    Private lblTitle As Label
    Private btnSave As Button
    Private btnCancel As Button

    ' 基本信息区（顶部固定）
    Private pnlBasicInfo As Panel
    Private txtName As TextBox
    Private cboCategory As ComboBox
    Private txtDescription As TextBox

    ' Tab控件
    Private tabControl As TabControl
    Private tabLayout As TabPage
    Private tabBodyStyles As TabPage
    Private tabPageSettings As TabPage
    Private tabAiGuidance As TabPage

    ' 版式骨架区
    Private lstLayoutElements As ListBox
    Private txtElementName As TextBox
    Private cboElementType As ComboBox
    Private cboElementFontCN As ComboBox
    Private cboElementFontSize As ComboBox
    Private chkElementBold As CheckBox
    Private cboElementAlignment As ComboBox
    Private pnlElementColor As Panel
    Private lblElementColor As Label
    Private btnElementColor As Button
    Private txtElementPlaceholder As TextBox
    Private btnAddElement As Button
    Private btnRemoveElement As Button
    
    ' 占位符预览区
    Private pnlPlaceholderPreview As Panel
    Private lblPlaceholderTitle As Label
    Private txtPlaceholderContent As TextBox
    Private btnUpdatePlaceholder As Button

    ' 正文样式区
    Private lstBodyStyles As ListBox
    Private txtStyleName As TextBox
    Private cboStyleFontCN As ComboBox
    Private cboStyleFontSize As ComboBox
    Private chkStyleBold As CheckBox
    Private cboStyleAlignment As ComboBox
    Private numStyleFirstIndent As NumericUpDown
    Private numStyleLineSpacing As NumericUpDown
    Private pnlStyleColor As Panel
    Private lblStyleColor As Label
    Private btnStyleColor As Button
    Private btnAddStyle As Button
    Private btnRemoveStyle As Button

    ' 页面设置区
    Private numMarginTop As NumericUpDown
    Private numMarginBottom As NumericUpDown
    Private numMarginLeft As NumericUpDown
    Private numMarginRight As NumericUpDown

    ' AI说明区
    Private txtAiGuidance As TextBox

    ' 事件
    Public Event TemplateSaved As EventHandler(Of ReformatTemplate)
    Public Event EditorClosed As EventHandler

    Public Sub New(Optional template As ReformatTemplate = Nothing, Optional previewCallback As PreviewStyleCallback = Nothing, Optional placeholderPreviewCallback As TemplatePlaceholderPreviewCallback = Nothing)
        _isNewTemplate = template Is Nothing
        If template IsNot Nothing Then
            Dim json = JsonConvert.SerializeObject(template)
            _template = JsonConvert.DeserializeObject(Of ReformatTemplate)(json)
        Else
            _template = New ReformatTemplate()
        End If
        _previewCallback = previewCallback
        _placeholderPreviewCallback = placeholderPreviewCallback

        InitializeControl()
        LoadTemplateData()
    End Sub

    Public Sub SetPreviewCallback(callback As PreviewStyleCallback)
        _previewCallback = callback
    End Sub
    
    Public Sub SetPlaceholderPreviewCallback(callback As TemplatePlaceholderPreviewCallback)
        _placeholderPreviewCallback = callback
    End Sub

    Private Sub InitializeControl()
        Me.BackColor = Color.FromArgb(248, 249, 250)
        Me.Padding = New Padding(0)

        CreateToolbar()
        CreateBasicInfoPanel()
        CreateTabControl()
    End Sub

    Private Sub CreateToolbar()
        pnlToolbar = New Panel With {
            .Dock = DockStyle.Top,
            .Height = 45,
            .BackColor = Color.FromArgb(102, 126, 234),
            .Padding = New Padding(10, 6, 10, 6)
        }

        lblTitle = New Label With {
            .Text = If(_isNewTemplate, "新建模板", "编辑模板"),
            .ForeColor = Color.White,
            .Font = New Font("Microsoft YaHei UI", 10, FontStyle.Bold),
            .AutoSize = True,
            .Location = New Point(10, 12)
        }
        pnlToolbar.Controls.Add(lblTitle)

        btnSave = New Button With {
            .Text = "保存",
            .Size = New Size(55, 28),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(76, 175, 80),
            .ForeColor = Color.White,
            .Cursor = Cursors.Hand
        }
        btnSave.FlatAppearance.BorderSize = 0
        AddHandler btnSave.Click, AddressOf BtnSave_Click
        pnlToolbar.Controls.Add(btnSave)

        btnCancel = New Button With {
            .Text = "取消",
            .Size = New Size(55, 28),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(180, 180, 180),
            .ForeColor = Color.FromArgb(60, 60, 60),
            .Cursor = Cursors.Hand
        }
        btnCancel.FlatAppearance.BorderSize = 0
        AddHandler btnCancel.Click, AddressOf BtnCancel_Click
        pnlToolbar.Controls.Add(btnCancel)

        AddHandler pnlToolbar.Resize, Sub(s, e)
                                          btnSave.Location = New Point(pnlToolbar.Width - 65, 8)
                                          btnCancel.Location = New Point(pnlToolbar.Width - 125, 8)
                                      End Sub

        Me.Controls.Add(pnlToolbar)
    End Sub

    Private Sub CreateBasicInfoPanel()
        pnlBasicInfo = New Panel With {
            .Dock = DockStyle.Top,
            .Height = 95,
            .Padding = New Padding(10, 8, 10, 5),
            .BackColor = Color.White
        }

        ' 第一行：名称和分类
        Dim lblName As New Label With {.Text = "名称:", .Location = New Point(10, 12), .AutoSize = True}
        pnlBasicInfo.Controls.Add(lblName)

        txtName = New TextBox With {.Location = New Point(50, 9), .Width = 130}
        pnlBasicInfo.Controls.Add(txtName)

        Dim lblCategory As New Label With {.Text = "分类:", .Location = New Point(190, 12), .AutoSize = True}
        pnlBasicInfo.Controls.Add(lblCategory)

        cboCategory = New ComboBox With {
            .Location = New Point(230, 9),
            .Width = 90,
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cboCategory.Items.AddRange({"通用", "行政", "学术", "商务"})
        pnlBasicInfo.Controls.Add(cboCategory)

        ' 第二行：描述
        Dim lblDesc As New Label With {.Text = "描述:", .Location = New Point(10, 42), .AutoSize = True}
        pnlBasicInfo.Controls.Add(lblDesc)

        txtDescription = New TextBox With {
            .Location = New Point(50, 39),
            .Width = 270,
            .Height = 45,
            .Multiline = True
        }
        pnlBasicInfo.Controls.Add(txtDescription)

        Me.Controls.Add(pnlBasicInfo)
        pnlBasicInfo.BringToFront()
    End Sub

    Private Sub CreateTabControl()
        tabControl = New TabControl With {
            .Dock = DockStyle.Fill,
            .Font = New Font("Microsoft YaHei UI", 9)
        }

        ' 版式骨架Tab
        tabLayout = New TabPage("版式骨架")
        CreateLayoutTab()
        tabControl.TabPages.Add(tabLayout)

        ' 正文样式Tab
        tabBodyStyles = New TabPage("正文样式")
        CreateBodyStylesTab()
        tabControl.TabPages.Add(tabBodyStyles)

        ' 页面设置Tab
        tabPageSettings = New TabPage("页面设置")
        CreatePageSettingsTab()
        tabControl.TabPages.Add(tabPageSettings)

        ' 占位符预览Tab
        Dim tabPlaceholderPreview As New TabPage("占位符预览")
        CreatePlaceholderPreviewTab(tabPlaceholderPreview)
        tabControl.TabPages.Add(tabPlaceholderPreview)
        
        ' AI说明Tab
        tabAiGuidance = New TabPage("AI说明")
        CreateAiGuidanceTab()
        tabControl.TabPages.Add(tabAiGuidance)

        Me.Controls.Add(tabControl)
        tabControl.BringToFront()
    End Sub

    Private Sub CreateLayoutTab()
        tabLayout.Padding = New Padding(8)

        ' 左侧：列表
        Dim lblList As New Label With {.Text = "元素列表:", .Location = New Point(8, 8), .AutoSize = True}
        tabLayout.Controls.Add(lblList)

        lstLayoutElements = New ListBox With {
            .Location = New Point(8, 28),
            .Size = New Size(110, 180)
        }
        AddHandler lstLayoutElements.SelectedIndexChanged, AddressOf LstLayoutElements_SelectedIndexChanged
        tabLayout.Controls.Add(lstLayoutElements)

        btnAddElement = New Button With {.Text = "添加", .Location = New Point(8, 212), .Size = New Size(50, 25)}
        AddHandler btnAddElement.Click, AddressOf BtnAddElement_Click
        tabLayout.Controls.Add(btnAddElement)

        btnRemoveElement = New Button With {.Text = "删除", .Location = New Point(62, 212), .Size = New Size(50, 25)}
        AddHandler btnRemoveElement.Click, AddressOf BtnRemoveElement_Click
        tabLayout.Controls.Add(btnRemoveElement)

        ' 右侧：编辑区
        Dim pnlEdit As New Panel With {
            .Location = New Point(128, 8),
            .Size = New Size(200, 230),
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.White
        }

        Dim y = 10
        pnlEdit.Controls.Add(New Label With {.Text = "名称:", .Location = New Point(8, y), .AutoSize = True})
        txtElementName = New TextBox With {.Location = New Point(60, y - 3), .Width = 125}
        pnlEdit.Controls.Add(txtElementName)
        y += 30

        pnlEdit.Controls.Add(New Label With {.Text = "元素:", .Location = New Point(8, y), .AutoSize = True})
        cboElementType = New ComboBox With {.Location = New Point(60, y - 3), .Width = 125, .DropDownStyle = ComboBoxStyle.DropDownList}
        ' 使用中文友好的显示，内部值保持英文
        cboElementType.Items.AddRange({"文本", "红线", "分隔线"})
        pnlEdit.Controls.Add(cboElementType)
        y += 30

        pnlEdit.Controls.Add(New Label With {.Text = "字体:", .Location = New Point(8, y), .AutoSize = True})
        cboElementFontCN = New ComboBox With {.Location = New Point(60, y - 3), .Width = 125, .DropDownStyle = ComboBoxStyle.DropDown}
        cboElementFontCN.Items.AddRange(GetCommonChineseFonts())
        cboElementFontCN.Text = "宋体"
        AddHandler cboElementFontCN.TextChanged, AddressOf OnLayoutStyleChanged
        AddHandler cboElementFontCN.SelectedIndexChanged, AddressOf OnLayoutStyleChanged
        pnlEdit.Controls.Add(cboElementFontCN)
        y += 30

        pnlEdit.Controls.Add(New Label With {.Text = "字号:", .Location = New Point(8, y), .AutoSize = True})
        cboElementFontSize = New ComboBox With {.Location = New Point(60, y - 3), .Width = 60, .DropDownStyle = ComboBoxStyle.DropDown}
        cboElementFontSize.Items.AddRange(GetCommonFontSizes())
        cboElementFontSize.Text = "12"
        AddHandler cboElementFontSize.TextChanged, AddressOf OnLayoutStyleChanged
        AddHandler cboElementFontSize.SelectedIndexChanged, AddressOf OnLayoutStyleChanged
        pnlEdit.Controls.Add(cboElementFontSize)
        chkElementBold = New CheckBox With {.Text = "粗体", .Location = New Point(125, y - 3), .AutoSize = True}
        AddHandler chkElementBold.CheckedChanged, AddressOf OnLayoutStyleChanged
        pnlEdit.Controls.Add(chkElementBold)
        y += 30

        pnlEdit.Controls.Add(New Label With {.Text = "对齐:", .Location = New Point(8, y), .AutoSize = True})
        cboElementAlignment = New ComboBox With {.Location = New Point(60, y - 3), .Width = 90, .DropDownStyle = ComboBoxStyle.DropDownList}
        cboElementAlignment.Items.AddRange({"左对齐", "居中", "右对齐", "两端对齐"})
        AddHandler cboElementAlignment.SelectedIndexChanged, AddressOf OnLayoutStyleChanged
        pnlEdit.Controls.Add(cboElementAlignment)
        y += 30
        
        pnlEdit.Controls.Add(New Label With {.Text = "颜色:", .Location = New Point(8, y), .AutoSize = True})
        btnElementColor = New Button With {
            .Location = New Point(60, y - 3),
            .Size = New Size(30, 24),
            .BackColor = ColorTranslator.FromHtml("#000000")
        }
        AddHandler btnElementColor.Click, AddressOf OnElementColorClick
        pnlEdit.Controls.Add(btnElementColor)
        lblElementColor = New Label With {
            .Location = New Point(95, y),
            .Size = New Size(80, 24),
            .Text = "#000000",
            .ForeColor = Color.Black
        }
        AddHandler lblElementColor.DoubleClick, AddressOf OnElementColorClick
        pnlEdit.Controls.Add(lblElementColor)
        y += 30

        pnlEdit.Controls.Add(New Label With {.Text = "占位符:", .Location = New Point(8, y), .AutoSize = True})
        txtElementPlaceholder = New TextBox With {
            .Location = New Point(60, y - 3),
            .Size = New Size(125, 23),
            .Text = "{{content}}"
        }
        AddHandler txtElementPlaceholder.TextChanged, AddressOf OnElementPlaceholderChanged
        pnlEdit.Controls.Add(txtElementPlaceholder)

        tabLayout.Controls.Add(pnlEdit)
    End Sub

    Private Sub CreateBodyStylesTab()
        tabBodyStyles.Padding = New Padding(8)

        ' 左侧：列表
        Dim lblList As New Label With {.Text = "样式列表:", .Location = New Point(8, 8), .AutoSize = True}
        tabBodyStyles.Controls.Add(lblList)

        lstBodyStyles = New ListBox With {
            .Location = New Point(8, 28),
            .Size = New Size(110, 180)
        }
        AddHandler lstBodyStyles.SelectedIndexChanged, AddressOf LstBodyStyles_SelectedIndexChanged
        tabBodyStyles.Controls.Add(lstBodyStyles)

        btnAddStyle = New Button With {.Text = "添加", .Location = New Point(8, 212), .Size = New Size(50, 25)}
        AddHandler btnAddStyle.Click, AddressOf BtnAddStyle_Click
        tabBodyStyles.Controls.Add(btnAddStyle)

        btnRemoveStyle = New Button With {.Text = "删除", .Location = New Point(62, 212), .Size = New Size(50, 25)}
        AddHandler btnRemoveStyle.Click, AddressOf BtnRemoveStyle_Click
        tabBodyStyles.Controls.Add(btnRemoveStyle)

        ' 右侧：编辑区
        Dim pnlEdit As New Panel With {
            .Location = New Point(128, 8),
            .Size = New Size(200, 230),
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.White
        }

        Dim y = 10
        pnlEdit.Controls.Add(New Label With {.Text = "名称:", .Location = New Point(8, y), .AutoSize = True})
        txtStyleName = New TextBox With {.Location = New Point(60, y - 3), .Width = 125}
        pnlEdit.Controls.Add(txtStyleName)
        y += 28

        pnlEdit.Controls.Add(New Label With {.Text = "字体:", .Location = New Point(8, y), .AutoSize = True})
        cboStyleFontCN = New ComboBox With {.Location = New Point(60, y - 3), .Width = 125, .DropDownStyle = ComboBoxStyle.DropDown}
        cboStyleFontCN.Items.AddRange(GetCommonChineseFonts())
        cboStyleFontCN.Text = "宋体"
        AddHandler cboStyleFontCN.TextChanged, AddressOf OnBodyStyleChanged
        AddHandler cboStyleFontCN.SelectedIndexChanged, AddressOf OnBodyStyleChanged
        pnlEdit.Controls.Add(cboStyleFontCN)
        y += 28

        pnlEdit.Controls.Add(New Label With {.Text = "字号:", .Location = New Point(8, y), .AutoSize = True})
        cboStyleFontSize = New ComboBox With {.Location = New Point(60, y - 3), .Width = 55, .DropDownStyle = ComboBoxStyle.DropDown}
        cboStyleFontSize.Items.AddRange(GetCommonFontSizes())
        cboStyleFontSize.Text = "12"
        AddHandler cboStyleFontSize.TextChanged, AddressOf OnBodyStyleChanged
        AddHandler cboStyleFontSize.SelectedIndexChanged, AddressOf OnBodyStyleChanged
        pnlEdit.Controls.Add(cboStyleFontSize)
        chkStyleBold = New CheckBox With {.Text = "粗体", .Location = New Point(120, y - 3), .AutoSize = True}
        AddHandler chkStyleBold.CheckedChanged, AddressOf OnBodyStyleChanged
        pnlEdit.Controls.Add(chkStyleBold)
        y += 28

        pnlEdit.Controls.Add(New Label With {.Text = "对齐:", .Location = New Point(8, y), .AutoSize = True})
        cboStyleAlignment = New ComboBox With {.Location = New Point(60, y - 3), .Width = 80, .DropDownStyle = ComboBoxStyle.DropDownList}
        cboStyleAlignment.Items.AddRange({"左对齐", "居中", "右对齐", "两端对齐"})
        AddHandler cboStyleAlignment.SelectedIndexChanged, AddressOf OnBodyStyleChanged
        pnlEdit.Controls.Add(cboStyleAlignment)
        y += 28

        pnlEdit.Controls.Add(New Label With {.Text = "缩进:", .Location = New Point(8, y), .AutoSize = True})
        numStyleFirstIndent = New NumericUpDown With {.Location = New Point(60, y - 3), .Width = 55, .Minimum = 0, .Maximum = 10, .Value = 2, .DecimalPlaces = 1}
        AddHandler numStyleFirstIndent.ValueChanged, AddressOf OnBodyStyleChanged
        pnlEdit.Controls.Add(numStyleFirstIndent)
        pnlEdit.Controls.Add(New Label With {.Text = "字符", .Location = New Point(120, y), .AutoSize = True})
        y += 28

        pnlEdit.Controls.Add(New Label With {.Text = "行距:", .Location = New Point(8, y), .AutoSize = True})
        numStyleLineSpacing = New NumericUpDown With {.Location = New Point(60, y - 3), .Width = 55, .Minimum = 1, .Maximum = 3, .Value = 1.5D, .DecimalPlaces = 1, .Increment = 0.1D}
        AddHandler numStyleLineSpacing.ValueChanged, AddressOf OnBodyStyleChanged
        pnlEdit.Controls.Add(numStyleLineSpacing)
        pnlEdit.Controls.Add(New Label With {.Text = "倍", .Location = New Point(120, y), .AutoSize = True})
        y += 28
        
        pnlEdit.Controls.Add(New Label With {.Text = "颜色:", .Location = New Point(8, y), .AutoSize = True})
        btnStyleColor = New Button With {
            .Location = New Point(60, y - 3),
            .Size = New Size(30, 24),
            .BackColor = ColorTranslator.FromHtml("#000000")
        }
        AddHandler btnStyleColor.Click, AddressOf OnStyleColorClick
        pnlEdit.Controls.Add(btnStyleColor)
        lblStyleColor = New Label With {
            .Location = New Point(95, y),
            .Size = New Size(80, 24),
            .Text = "#000000",
            .ForeColor = Color.Black
        }
        AddHandler lblStyleColor.DoubleClick, AddressOf OnStyleColorClick
        pnlEdit.Controls.Add(lblStyleColor)

        tabBodyStyles.Controls.Add(pnlEdit)
    End Sub

    Private Sub CreatePageSettingsTab()
        tabPageSettings.Padding = New Padding(15)

        Dim grp As New GroupBox With {
            .Text = "页边距 (cm)",
            .Location = New Point(15, 15),
            .Size = New Size(300, 130),
            .Font = New Font("Microsoft YaHei UI", 9)
        }

        grp.Controls.Add(New Label With {.Text = "上:", .Location = New Point(20, 30), .AutoSize = True})
        numMarginTop = New NumericUpDown With {.Location = New Point(50, 27), .Width = 70, .Minimum = 0, .Maximum = 10, .Value = 2.54D, .DecimalPlaces = 2, .Increment = 0.1D}
        grp.Controls.Add(numMarginTop)

        grp.Controls.Add(New Label With {.Text = "下:", .Location = New Point(140, 30), .AutoSize = True})
        numMarginBottom = New NumericUpDown With {.Location = New Point(170, 27), .Width = 70, .Minimum = 0, .Maximum = 10, .Value = 2.54D, .DecimalPlaces = 2, .Increment = 0.1D}
        grp.Controls.Add(numMarginBottom)

        grp.Controls.Add(New Label With {.Text = "左:", .Location = New Point(20, 70), .AutoSize = True})
        numMarginLeft = New NumericUpDown With {.Location = New Point(50, 67), .Width = 70, .Minimum = 0, .Maximum = 10, .Value = 3.18D, .DecimalPlaces = 2, .Increment = 0.1D}
        grp.Controls.Add(numMarginLeft)

        grp.Controls.Add(New Label With {.Text = "右:", .Location = New Point(140, 70), .AutoSize = True})
        numMarginRight = New NumericUpDown With {.Location = New Point(170, 67), .Width = 70, .Minimum = 0, .Maximum = 10, .Value = 3.18D, .DecimalPlaces = 2, .Increment = 0.1D}
        grp.Controls.Add(numMarginRight)

        tabPageSettings.Controls.Add(grp)
    End Sub

    Private Sub CreatePlaceholderPreviewTab(tab As TabPage)
        tab.Padding = New Padding(10)
        
        ' 标题
        lblPlaceholderTitle = New Label With {
            .Text = "实时占位符预览",
            .Location = New Point(10, 10),
            .AutoSize = True,
            .Font = New Font("Microsoft YaHei UI", 10, FontStyle.Bold)
        }
        tab.Controls.Add(lblPlaceholderTitle)
        
        ' 预览内容区域
        txtPlaceholderContent = New TextBox With {
            .Location = New Point(10, 40),
            .Size = New Size(310, 180),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical,
            .Text = "在此区域编辑占位符内容，实时预览效果..."
        }
        AddHandler txtPlaceholderContent.TextChanged, AddressOf OnPlaceholderContentChanged
        tab.Controls.Add(txtPlaceholderContent)
        
        ' 更新按钮
        btnUpdatePlaceholder = New Button With {
            .Text = "更新预览",
            .Location = New Point(10, 230),
            .Size = New Size(80, 25)
        }
        AddHandler btnUpdatePlaceholder.Click, AddressOf OnUpdatePlaceholderClick
        tab.Controls.Add(btnUpdatePlaceholder)
    End Sub
    
    Private Sub CreateAiGuidanceTab()
        tabAiGuidance.Padding = New Padding(10)

        Dim lbl As New Label With {
            .Text = "AI补充说明（给AI的额外上下文信息）:",
            .Location = New Point(10, 10),
            .AutoSize = True
        }
        tabAiGuidance.Controls.Add(lbl)

        txtAiGuidance = New TextBox With {
            .Location = New Point(10, 35),
            .Size = New Size(310, 200),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical
        }
        tabAiGuidance.Controls.Add(txtAiGuidance)
    End Sub

    Private Sub LoadTemplateData()
        If _template Is Nothing Then Return

        txtName.Text = _template.Name
        txtDescription.Text = _template.Description
        If cboCategory.Items.Contains(_template.Category) Then
            cboCategory.SelectedItem = _template.Category
        Else
            cboCategory.SelectedIndex = 0
        End If

        lstLayoutElements.Items.Clear()
        If _template.Layout?.Elements IsNot Nothing Then
            For Each el In _template.Layout.Elements
                lstLayoutElements.Items.Add(el)
            Next
        End If

        lstBodyStyles.Items.Clear()
        If _template.BodyStyles IsNot Nothing Then
            For Each style In _template.BodyStyles
                lstBodyStyles.Items.Add(style)
            Next
        End If

        If _template.PageSettings?.Margins IsNot Nothing Then
            numMarginTop.Value = CDec(Math.Max(0, Math.Min(10, _template.PageSettings.Margins.Top)))
            numMarginBottom.Value = CDec(Math.Max(0, Math.Min(10, _template.PageSettings.Margins.Bottom)))
            numMarginLeft.Value = CDec(Math.Max(0, Math.Min(10, _template.PageSettings.Margins.Left)))
            numMarginRight.Value = CDec(Math.Max(0, Math.Min(10, _template.PageSettings.Margins.Right)))
        End If

        txtAiGuidance.Text = _template.AiGuidance
    End Sub

#Region "事件处理"

    Private Sub LstLayoutElements_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' 防止递归
        If _isUpdatingList Then Return

        ' 先保存上一个选中项的数据
        If _lastLayoutIndex >= 0 AndAlso _lastLayoutIndex < lstLayoutElements.Items.Count Then
            SaveLayoutElementAt(_lastLayoutIndex)
        End If

        ' 更新索引并加载新数据
        _lastLayoutIndex = lstLayoutElements.SelectedIndex

        If lstLayoutElements.SelectedItem Is Nothing Then Return
        Dim el = CType(lstLayoutElements.SelectedItem, LayoutElement)

        If txtElementName IsNot Nothing Then txtElementName.Text = If(el.Name, "")
        If cboElementType IsNot Nothing AndAlso el.ElementType IsNot Nothing Then
            cboElementType.SelectedIndex = GetElementTypeIndex(el.ElementType)
        End If

        If el.Font IsNot Nothing Then
            If cboElementFontCN IsNot Nothing Then cboElementFontCN.Text = If(el.Font.FontNameCN, "宋体")
            If cboElementFontSize IsNot Nothing Then cboElementFontSize.Text = el.Font.FontSize.ToString()
            If chkElementBold IsNot Nothing Then chkElementBold.Checked = el.Font.Bold
        End If

        If el.Paragraph IsNot Nothing AndAlso cboElementAlignment IsNot Nothing AndAlso el.Paragraph.Alignment IsNot Nothing Then
            cboElementAlignment.SelectedIndex = GetAlignmentIndex(el.Paragraph.Alignment)
        End If
        
        ' 处理颜色设置
        If el.Color IsNot Nothing AndAlso lblElementColor IsNot Nothing Then
            Try
                Dim color = ColorTranslator.FromHtml(el.Color.FontColor)
                lblElementColor.Text = el.Color.FontColor
                If btnElementColor IsNot Nothing Then
                    btnElementColor.BackColor = color
                End If
            Catch ex As Exception
                Debug.WriteLine($"颜色解析错误: {ex.Message}")
                lblElementColor.Text = "#000000"
                If btnElementColor IsNot Nothing Then
                    btnElementColor.BackColor = Color.Black
                End If
            End Try
        Else
            lblElementColor.Text = "#000000"
            If btnElementColor IsNot Nothing Then
                btnElementColor.BackColor = Color.Black
            End If
        End If
                
        ' 处理占位符内容
        If txtElementPlaceholder IsNot Nothing Then
            txtElementPlaceholder.Text = If(el.PlaceholderContent, "{{content}}")
        End If
                
        ' 更新占位符预览内容
        If txtPlaceholderContent IsNot Nothing Then
            txtPlaceholderContent.Text = If(el.PlaceholderContent, "{{content}}")
        End If
    End Sub

    Private Sub LstBodyStyles_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' 防止递归
        If _isUpdatingList Then Return

        ' 先保存上一个选中项的数据
        If _lastStyleIndex >= 0 AndAlso _lastStyleIndex < lstBodyStyles.Items.Count Then
            SaveBodyStyleAt(_lastStyleIndex)
        End If

        ' 更新索引并加载新数据
        _lastStyleIndex = lstBodyStyles.SelectedIndex

        If lstBodyStyles.SelectedItem Is Nothing Then Return
        Dim style = CType(lstBodyStyles.SelectedItem, StyleRule)

        If txtStyleName IsNot Nothing Then txtStyleName.Text = If(style.RuleName, "")

        If style.Font IsNot Nothing Then
            If cboStyleFontCN IsNot Nothing Then cboStyleFontCN.Text = If(style.Font.FontNameCN, "宋体")
            If cboStyleFontSize IsNot Nothing Then cboStyleFontSize.Text = style.Font.FontSize.ToString()
            If chkStyleBold IsNot Nothing Then chkStyleBold.Checked = style.Font.Bold
        End If

        If style.Paragraph IsNot Nothing Then
            If cboStyleAlignment IsNot Nothing AndAlso style.Paragraph.Alignment IsNot Nothing Then
                cboStyleAlignment.SelectedIndex = GetAlignmentIndex(style.Paragraph.Alignment)
            End If
            If numStyleFirstIndent IsNot Nothing Then numStyleFirstIndent.Value = CDec(Math.Max(0, Math.Min(10, style.Paragraph.FirstLineIndent)))
            If numStyleLineSpacing IsNot Nothing Then numStyleLineSpacing.Value = CDec(Math.Max(1, Math.Min(3, style.Paragraph.LineSpacing)))
        End If
        
        ' 处理颜色设置
        If style.Color IsNot Nothing AndAlso lblStyleColor IsNot Nothing Then
            Try
                Dim color = ColorTranslator.FromHtml(style.Color.FontColor)
                lblStyleColor.Text = style.Color.FontColor
                If btnStyleColor IsNot Nothing Then
                    btnStyleColor.BackColor = color
                End If
            Catch ex As Exception
                Debug.WriteLine($"颜色解析错误: {ex.Message}")
                lblStyleColor.Text = "#000000"
                If btnStyleColor IsNot Nothing Then
                    btnStyleColor.BackColor = Color.Black
                End If
            End Try
        Else
            lblStyleColor.Text = "#000000"
            If btnStyleColor IsNot Nothing Then
                btnStyleColor.BackColor = Color.Black
            End If
        End If
    End Sub

    Private Sub BtnAddElement_Click(sender As Object, e As EventArgs)
        Dim newEl As New LayoutElement With {
            .Name = "新元素" & (lstLayoutElements.Items.Count + 1),
            .ElementType = "text",
            .Font = New FontConfig With {.FontNameCN = "宋体", .FontSize = 12},
            .Paragraph = New ParagraphConfig With {.Alignment = "left"},
            .Color = New ColorConfig With {.FontColor = "#000000"}
        }
        If _template.Layout Is Nothing Then _template.Layout = New LayoutConfig()
        If _template.Layout.Elements Is Nothing Then _template.Layout.Elements = New List(Of LayoutElement)()
        _template.Layout.Elements.Add(newEl)
        lstLayoutElements.Items.Add(newEl)
        lstLayoutElements.SelectedItem = newEl
    End Sub

    Private Sub BtnRemoveElement_Click(sender As Object, e As EventArgs)
        If lstLayoutElements.SelectedItem Is Nothing Then Return
        Dim el = CType(lstLayoutElements.SelectedItem, LayoutElement)
        _template.Layout?.Elements?.Remove(el)
        lstLayoutElements.Items.Remove(el)
    End Sub

    Private Sub BtnAddStyle_Click(sender As Object, e As EventArgs)
        Dim newStyle As New StyleRule With {
            .RuleName = "新样式" & (lstBodyStyles.Items.Count + 1),
            .Font = New FontConfig With {.FontNameCN = "宋体", .FontSize = 12},
            .Paragraph = New ParagraphConfig With {.Alignment = "justify", .FirstLineIndent = 2, .LineSpacing = 1.5},
            .Color = New ColorConfig With {.FontColor = "#000000"}
        }
        If _template.BodyStyles Is Nothing Then _template.BodyStyles = New List(Of StyleRule)()
        _template.BodyStyles.Add(newStyle)
        lstBodyStyles.Items.Add(newStyle)
        lstBodyStyles.SelectedItem = newStyle
    End Sub

    Private Sub BtnRemoveStyle_Click(sender As Object, e As EventArgs)
        If lstBodyStyles.SelectedItem Is Nothing Then Return
        Dim style = CType(lstBodyStyles.SelectedItem, StyleRule)
        _template.BodyStyles?.Remove(style)
        lstBodyStyles.Items.Remove(style)
    End Sub

    Private Sub OnLayoutStyleChanged(sender As Object, e As EventArgs)
        If _previewCallback Is Nothing Then Return
        Try
            _previewCallback.Invoke(
                If(cboElementFontCN?.Text, "宋体"),
                ParseFontSize(cboElementFontSize?.Text),
                If(chkElementBold IsNot Nothing, chkElementBold.Checked, False),
                GetAlignmentValue(cboElementAlignment?.SelectedIndex), 0, 1.5)
        Catch ex As Exception
            Debug.WriteLine($"OnLayoutStyleChanged error: {ex.Message}")
        End Try
        SaveCurrentLayoutElement()
    End Sub
    
    Private Sub OnElementColorClick(sender As Object, e As EventArgs)
        Dim colorDlg As New ColorDialog With {
            .FullOpen = True,
            .AnyColor = True,
            .AllowFullOpen = True
        }
            
        If cboElementFontCN IsNot Nothing Then
            Try
                colorDlg.Color = btnElementColor.BackColor
            Catch
                colorDlg.Color = Color.Black
            End Try
        End If
            
        If colorDlg.ShowDialog() = DialogResult.OK Then
            Dim selectedColor = colorDlg.Color
            Dim colorHex = ColorTranslator.ToHtml(selectedColor)
                
            If btnElementColor IsNot Nothing Then
                btnElementColor.BackColor = selectedColor
            End If
            If lblElementColor IsNot Nothing Then
                lblElementColor.Text = colorHex
            End If
                
            SaveCurrentLayoutElement()
        End If
    End Sub
        
    Private Sub OnElementPlaceholderChanged(sender As Object, e As EventArgs)
        SaveCurrentLayoutElement()
        ' 触发实时预览
        TriggerPlaceholderPreview()
    End Sub
        
    Private Sub OnPlaceholderContentChanged(sender As Object, e As EventArgs)
        ' 实时预览内容变化
        TriggerPlaceholderPreview()
    End Sub
        
    Private Sub OnUpdatePlaceholderClick(sender As Object, e As EventArgs)
        TriggerPlaceholderPreview()
    End Sub

    Private Sub OnBodyStyleChanged(sender As Object, e As EventArgs)
        If _previewCallback Is Nothing Then Return
        Try
            _previewCallback.Invoke(
                If(cboStyleFontCN?.Text, "宋体"),
                ParseFontSize(cboStyleFontSize?.Text),
                If(chkStyleBold IsNot Nothing, chkStyleBold.Checked, False),
                GetAlignmentValue(cboStyleAlignment?.SelectedIndex),
                If(numStyleFirstIndent IsNot Nothing, CDbl(numStyleFirstIndent.Value), 0),
                If(numStyleLineSpacing IsNot Nothing, CDbl(numStyleLineSpacing.Value), 1.5))
        Catch ex As Exception
            Debug.WriteLine($"OnBodyStyleChanged error: {ex.Message}")
        End Try
        SaveCurrentBodyStyle()
    End Sub
    
    Private Sub OnStyleColorClick(sender As Object, e As EventArgs)
        Dim colorDlg As New ColorDialog With {
            .FullOpen = True,
            .AnyColor = True,
            .AllowFullOpen = True
        }
            
        If btnStyleColor IsNot Nothing Then
            Try
                colorDlg.Color = btnStyleColor.BackColor
            Catch
                colorDlg.Color = Color.Black
            End Try
        End If
            
        If colorDlg.ShowDialog() = DialogResult.OK Then
            Dim selectedColor = colorDlg.Color
            Dim colorHex = ColorTranslator.ToHtml(selectedColor)
                
            If btnStyleColor IsNot Nothing Then
                btnStyleColor.BackColor = selectedColor
            End If
            If lblStyleColor IsNot Nothing Then
                lblStyleColor.Text = colorHex
            End If
                
            SaveCurrentBodyStyle()
        End If
    End Sub

    ''' <summary>
    ''' 按索引保存版式元素数据
    ''' </summary>
    Private Sub SaveLayoutElementAt(index As Integer)
        If lstLayoutElements Is Nothing OrElse index < 0 OrElse index >= lstLayoutElements.Items.Count Then Return
        Try
            Dim el = CType(lstLayoutElements.Items(index), LayoutElement)
            If txtElementName IsNot Nothing Then el.Name = txtElementName.Text
            If cboElementType IsNot Nothing Then el.ElementType = GetElementTypeValue(cboElementType.SelectedIndex)
            If el.Font Is Nothing Then el.Font = New FontConfig()
            If cboElementFontCN IsNot Nothing Then el.Font.FontNameCN = cboElementFontCN.Text
            If cboElementFontSize IsNot Nothing Then el.Font.FontSize = ParseFontSize(cboElementFontSize.Text)
            If chkElementBold IsNot Nothing Then el.Font.Bold = chkElementBold.Checked
            If el.Paragraph Is Nothing Then el.Paragraph = New ParagraphConfig()
            If cboElementAlignment IsNot Nothing Then el.Paragraph.Alignment = GetAlignmentValue(cboElementAlignment.SelectedIndex)
                        
            ' 处理颜色设置
            If el.Color Is Nothing Then el.Color = New ColorConfig()
            If lblElementColor IsNot Nothing Then
                el.Color.FontColor = lblElementColor.Text
            Else
                el.Color.FontColor = "#000000"
            End If
            
            ' 处理占位符内容
            If txtElementPlaceholder IsNot Nothing Then
                el.PlaceholderContent = txtElementPlaceholder.Text
            Else
                el.PlaceholderContent = "{{content}}"
            End If
                        
            ' 更新列表显示（设置标志防止触发事件递归）
            _isUpdatingList = True
            lstLayoutElements.Items(index) = el
            _isUpdatingList = False
        Catch ex As Exception
            _isUpdatingList = False
            Debug.WriteLine($"SaveLayoutElementAt error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 按索引保存正文样式数据
    ''' </summary>
    Private Sub SaveBodyStyleAt(index As Integer)
        If lstBodyStyles Is Nothing OrElse index < 0 OrElse index >= lstBodyStyles.Items.Count Then Return
        Try
            Dim style = CType(lstBodyStyles.Items(index), StyleRule)
            If txtStyleName IsNot Nothing Then style.RuleName = txtStyleName.Text
            If style.Font Is Nothing Then style.Font = New FontConfig()
            If cboStyleFontCN IsNot Nothing Then style.Font.FontNameCN = cboStyleFontCN.Text
            If cboStyleFontSize IsNot Nothing Then style.Font.FontSize = ParseFontSize(cboStyleFontSize.Text)
            If chkStyleBold IsNot Nothing Then style.Font.Bold = chkStyleBold.Checked
            If style.Paragraph Is Nothing Then style.Paragraph = New ParagraphConfig()
            If cboStyleAlignment IsNot Nothing Then style.Paragraph.Alignment = GetAlignmentValue(cboStyleAlignment.SelectedIndex)
            If numStyleFirstIndent IsNot Nothing Then style.Paragraph.FirstLineIndent = CDbl(numStyleFirstIndent.Value)
            If numStyleLineSpacing IsNot Nothing Then style.Paragraph.LineSpacing = CDbl(numStyleLineSpacing.Value)
                        
            ' 处理颜色设置
            If style.Color Is Nothing Then style.Color = New ColorConfig()
            If lblStyleColor IsNot Nothing Then
                style.Color.FontColor = lblStyleColor.Text
            Else
                style.Color.FontColor = "#000000"
            End If
                        
            ' 更新列表显示（设置标志防止触发事件递归）
            _isUpdatingList = True
            lstBodyStyles.Items(index) = style
            _isUpdatingList = False
        Catch ex As Exception
            _isUpdatingList = False
            Debug.WriteLine($"SaveBodyStyleAt error: {ex.Message}")
        End Try
    End Sub

    Private Sub SaveCurrentLayoutElement()
        If lstLayoutElements Is Nothing OrElse lstLayoutElements.SelectedItem Is Nothing Then Return
        Try
            Dim el = CType(lstLayoutElements.SelectedItem, LayoutElement)
            If txtElementName IsNot Nothing Then el.Name = txtElementName.Text
            If cboElementType IsNot Nothing Then el.ElementType = GetElementTypeValue(cboElementType.SelectedIndex)
            If el.Font Is Nothing Then el.Font = New FontConfig()
            If cboElementFontCN IsNot Nothing Then el.Font.FontNameCN = cboElementFontCN.Text
            If cboElementFontSize IsNot Nothing Then el.Font.FontSize = ParseFontSize(cboElementFontSize.Text)
            If chkElementBold IsNot Nothing Then el.Font.Bold = chkElementBold.Checked
            If el.Paragraph Is Nothing Then el.Paragraph = New ParagraphConfig()
            If cboElementAlignment IsNot Nothing Then el.Paragraph.Alignment = GetAlignmentValue(cboElementAlignment.SelectedIndex)
                        
            ' 处理颜色设置
            If el.Color Is Nothing Then el.Color = New ColorConfig()
            If lblElementColor IsNot Nothing Then
                el.Color.FontColor = lblElementColor.Text
            Else
                el.Color.FontColor = "#000000"
            End If
            
            ' 处理占位符内容
            If txtElementPlaceholder IsNot Nothing Then
                el.PlaceholderContent = txtElementPlaceholder.Text
            Else
                el.PlaceholderContent = "{{content}}"
            End If
                        
            Dim idx = lstLayoutElements.SelectedIndex
            If idx >= 0 AndAlso idx < lstLayoutElements.Items.Count Then
                _isUpdatingList = True
                lstLayoutElements.Items(idx) = el
                _isUpdatingList = False
            End If
        Catch ex As Exception
            _isUpdatingList = False
            Debug.WriteLine($"SaveCurrentLayoutElement error: {ex.Message}")
        End Try
    End Sub

    Private Sub SaveCurrentBodyStyle()
        If lstBodyStyles Is Nothing OrElse lstBodyStyles.SelectedItem Is Nothing Then Return
        Try
            Dim style = CType(lstBodyStyles.SelectedItem, StyleRule)
            If txtStyleName IsNot Nothing Then style.RuleName = txtStyleName.Text
            If style.Font Is Nothing Then style.Font = New FontConfig()
            If cboStyleFontCN IsNot Nothing Then style.Font.FontNameCN = cboStyleFontCN.Text
            If cboStyleFontSize IsNot Nothing Then style.Font.FontSize = ParseFontSize(cboStyleFontSize.Text)
            If chkStyleBold IsNot Nothing Then style.Font.Bold = chkStyleBold.Checked
            If style.Paragraph Is Nothing Then style.Paragraph = New ParagraphConfig()
            If cboStyleAlignment IsNot Nothing Then style.Paragraph.Alignment = GetAlignmentValue(cboStyleAlignment.SelectedIndex)
            If numStyleFirstIndent IsNot Nothing Then style.Paragraph.FirstLineIndent = CDbl(numStyleFirstIndent.Value)
            If numStyleLineSpacing IsNot Nothing Then style.Paragraph.LineSpacing = CDbl(numStyleLineSpacing.Value)
                        
            ' 处理颜色设置
            If style.Color Is Nothing Then style.Color = New ColorConfig()
            If lblStyleColor IsNot Nothing Then
                style.Color.FontColor = lblStyleColor.Text
            Else
                style.Color.FontColor = "#000000"
            End If
                        
            Dim idx = lstBodyStyles.SelectedIndex
            If idx >= 0 AndAlso idx < lstBodyStyles.Items.Count Then
                _isUpdatingList = True
                lstBodyStyles.Items(idx) = style
                _isUpdatingList = False
            End If
        Catch ex As Exception
            _isUpdatingList = False
            Debug.WriteLine($"SaveCurrentBodyStyle error: {ex.Message}")
        End Try
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtName.Text) Then
            MessageBox.Show("请输入模板名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' 先保存当前正在编辑的元素和样式
        SaveCurrentLayoutElement()
        SaveCurrentBodyStyle()

        _template.Name = txtName.Text.Trim()
        _template.Description = txtDescription.Text
        _template.Category = cboCategory.SelectedItem?.ToString()
        _template.AiGuidance = txtAiGuidance.Text
        If _template.PageSettings Is Nothing Then _template.PageSettings = New PageConfig()
        If _template.PageSettings.Margins Is Nothing Then _template.PageSettings.Margins = New MarginsConfig()
        _template.PageSettings.Margins.Top = CDbl(numMarginTop.Value)
        _template.PageSettings.Margins.Bottom = CDbl(numMarginBottom.Value)
        _template.PageSettings.Margins.Left = CDbl(numMarginLeft.Value)
        _template.PageSettings.Margins.Right = CDbl(numMarginRight.Value)
        If _isNewTemplate Then
            ReformatTemplateManager.Instance.AddTemplate(_template)
        Else
            ReformatTemplateManager.Instance.UpdateTemplate(_template)
        End If
        RaiseEvent TemplateSaved(Me, _template)
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs)
        RaiseEvent EditorClosed(Me, EventArgs.Empty)
    End Sub

#End Region

#Region "辅助方法"

    ''' <summary>
    ''' 获取常用中文字体列表
    ''' </summary>
    Private Shared Function GetCommonChineseFonts() As String()
        Return New String() {
            "宋体", "黑体", "楷体", "仿宋", "微软雅黑",
            "华文宋体", "华文黑体", "华文楷体", "华文仿宋", "华文中宋",
            "方正小标宋简体", "方正仿宋简体", "方正楷体简体", "方正黑体简体",
            "Times New Roman", "Arial"
        }
    End Function

    ''' <summary>
    ''' 获取常用字号列表（与Office一致）
    ''' </summary>
    Private Shared Function GetCommonFontSizes() As String()
        Return New String() {
            "初号", "小初", "一号", "小一", "二号", "小二",
            "三号", "小三", "四号", "小四", "五号", "小五",
            "六号", "小六", "七号", "八号",
            "8", "9", "10", "10.5", "11", "12", "14", "16",
            "18", "20", "22", "24", "26", "28", "36", "48", "72"
        }
    End Function

    ''' <summary>
    ''' 解析字号（支持中文字号名称和数字）
    ''' </summary>
    Private Shared Function ParseFontSize(sizeText As String) As Double
        If String.IsNullOrWhiteSpace(sizeText) Then Return 12.0
        Select Case sizeText.Trim()
            Case "初号" : Return 42
            Case "小初" : Return 36
            Case "一号" : Return 26
            Case "小一" : Return 24
            Case "二号" : Return 22
            Case "小二" : Return 18
            Case "三号" : Return 16
            Case "小三" : Return 15
            Case "四号" : Return 14
            Case "小四" : Return 12
            Case "五号" : Return 10.5
            Case "小五" : Return 9
            Case "六号" : Return 7.5
            Case "小六" : Return 6.5
            Case "七号" : Return 5.5
            Case "八号" : Return 5
            Case Else
                Dim result As Double
                If Double.TryParse(sizeText, result) Then Return result
                Return 12.0
        End Select
    End Function

    ''' <summary>
    ''' 获取元素类型的中文索引
    ''' </summary>
    Private Shared Function GetElementTypeIndex(elementType As String) As Integer
        Select Case elementType
            Case "text" : Return 0
            Case "redLine" : Return 1
            Case "separator" : Return 2
            Case Else : Return 0
        End Select
    End Function

    ''' <summary>
    ''' 获取元素类型的英文值
    ''' </summary>
    Private Shared Function GetElementTypeValue(index As Integer) As String
        Select Case index
            Case 0 : Return "text"
            Case 1 : Return "redLine"
            Case 2 : Return "separator"
            Case Else : Return "text"
        End Select
    End Function

    ''' <summary>
    ''' 获取对齐方式的中文索引
    ''' </summary>
    Private Shared Function GetAlignmentIndex(alignment As String) As Integer
        Select Case alignment
            Case "left" : Return 0
            Case "center" : Return 1
            Case "right" : Return 2
            Case "justify" : Return 3
            Case Else : Return 0
        End Select
    End Function

    ''' <summary>
    ''' 获取对齐方式的英文值
    ''' </summary>
    Private Shared Function GetAlignmentValue(index As Integer?) As String
        If index Is Nothing Then Return "left"
        Select Case index.Value
            Case 0 : Return "left"
            Case 1 : Return "center"
            Case 2 : Return "right"
            Case 3 : Return "justify"
            Case Else : Return "left"
        End Select
    End Function
    
    ''' <summary>
    ''' 触发占位符预览
    ''' </summary>
    Private Sub TriggerPlaceholderPreview()
        If _placeholderPreviewCallback Is Nothing OrElse lstLayoutElements.SelectedItem Is Nothing Then Return
        
        Try
            Dim el = CType(lstLayoutElements.SelectedItem, LayoutElement)
            Dim placeholderId = If(el.Name, "UnknownElement")
            Dim content = If(txtPlaceholderContent?.Text, "{{content}}")
            Dim fontConfig = el.Font
            Dim paragraphConfig = el.Paragraph
            Dim colorConfig = el.Color
            
            If fontConfig Is Nothing Then fontConfig = New FontConfig()
            If paragraphConfig Is Nothing Then paragraphConfig = New ParagraphConfig()
            If colorConfig Is Nothing Then colorConfig = New ColorConfig()
            
            _placeholderPreviewCallback.Invoke(placeholderId, content, fontConfig, paragraphConfig, colorConfig)
            
        Catch ex As Exception
            Debug.WriteLine($"占位符预览失败: {ex.Message}")
        End Try
    End Sub

#End Region

End Class
